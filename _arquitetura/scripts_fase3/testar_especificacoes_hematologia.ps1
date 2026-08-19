# testar_especificacoes_hematologia.ps1 - prova estrutural e de nao regressao
#
# Compara a tabela fato antes/depois da instalacao formal. Os 1.575 vereditos
# Westgard e os 1.125 Sigma existentes devem permanecer identicos. A prova roda
# os motores em memoria e fecha os arquivos sem salvar.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).

param(
    [Parameter(Mandatory = $true)][string]$Referencia,
    [Parameter(Mandatory = $true)][string]$Candidato,
    [Parameter(Mandatory = $true)][string]$OutCsv
)

$ErrorActionPreference = 'Stop'
$Referencia = (Resolve-Path -LiteralPath $Referencia).ProviderPath
$Candidato = (Resolve-Path -LiteralPath $Candidato).ProviderPath
foreach ($p in @($Referencia, $Candidato)) {
    if ($p -notmatch '(?i)(_EM_DESENVOLVIMENTO|build_hardening)') {
        throw "Recusado: teste somente em clone _EM_DESENVOLVIMENTO ou build_hardening: $p"
    }
}

function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $ultimo = $_
            if ($tentativa -eq 2) {
                try { Start-Process excel.exe -WindowStyle Hidden -ErrorAction SilentlyContinue | Out-Null; Start-Sleep 5 } catch { }
            }
            Start-Sleep -Seconds ($tentativa * 2)
        }
    }
    throw "Excel COM nao subiu: $($ultimo.Exception.Message)"
}

function Mapa-Cabecalhos($ws) {
    $m = @{}
    $ult = $ws.Cells.Item(1, $ws.Columns.Count).End(-4159).Column
    for ($c = 1; $c -le $ult; $c++) {
        $nome = [string]$ws.Cells.Item(1, $c).Value2
        if ($nome) { $m[$nome] = $c }
    }
    return $m
}

function Foto-BI([string]$Caminho, [bool]$ComEspecificacoes) {
    $xl = Novo-Excel
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $xl.AutomationSecurity = 1
    $wb = $xl.Workbooks.Open($Caminho, 0, $true)
    try {
        try { $wb.EnableAutoRecover = $false } catch { }
        foreach ($rotina in @('AtualizarCalc','AtualizarPainelEng','AtualizarEstatisticaAba')) {
            $xl.Run("$($wb.Name)!$rotina") | Out-Null
        }
        if ($ComEspecificacoes) { $xl.Run("$($wb.Name)!AtualizarEngEspec") | Out-Null }
        $xl.Run("$($wb.Name)!AtualizarBIData") | Out-Null

        $bi = $wb.Worksheets.Item('BI_Data')
        $cab = Mapa-Cabecalhos $bi
        foreach ($obrigatorio in @('ID_Result_Global','Analito','Veredito','Sigma')) {
            if (-not $cab.ContainsKey($obrigatorio)) { throw "BI_Data sem coluna $obrigatorio em $Caminho" }
        }
        $ult = $bi.Cells.Item($bi.Rows.Count, $cab['ID_Result_Global']).End(-4162).Row
        $dados = @{}
        $analitosDB = @{}
        for ($r = 2; $r -le $ult; $r++) {
            $chave = [string]$bi.Cells.Item($r, $cab['ID_Result_Global']).Value2
            if (-not $chave) { continue }
            $analito = [string]$bi.Cells.Item($r, $cab['Analito']).Value2
            $analitosDB[$analito] = $true
            $dados[$chave] = [PSCustomObject]@{
                Veredito = [string]$bi.Cells.Item($r, $cab['Veredito']).Value2
                Sigma = $bi.Cells.Item($r, $cab['Sigma']).Value2
            }
        }

        $estrutura = $null
        if ($ComEspecificacoes) {
            foreach ($nome in @('Cfg_Especificacoes','DB_Especificacoes','Eng_Especificacoes')) {
                try { $null = $wb.Worksheets.Item($nome) } catch { throw "Aba ausente no candidato: $nome" }
            }
            foreach ($nome in @('engEspAnalito','engCVtp','engBIAStp','engETp','engEspAno','engEspSituacao')) {
                try { $null = $wb.Names.Item($nome) } catch { throw "Nome definido ausente: $nome" }
            }

            $cfg = $wb.Worksheets.Item('Cfg_Especificacoes')
            $db = $wb.Worksheets.Item('DB_Especificacoes')
            $eng = $wb.Worksheets.Item('Eng_Especificacoes')
            $an = $wb.Worksheets.Item('Analitos')

            $fontesCfg = 0
            for ($r = 7; $r -le 40; $r++) { if ([string]$cfg.Cells.Item($r, 1).Value2) { $fontesCfg++ } }
            $linhasDB = 0
            for ($r = 4; $r -le [Math]::Max(4, $db.Cells.Item($db.Rows.Count, 1).End(-4162).Row); $r++) {
                if ([string]$db.Cells.Item($r, 1).Value2) { $linhasDB++ }
            }

            $fallbackDB = 0
            $fallbackDivergente = 0
            for ($r = 4; $r -le 43; $r++) {
                $nome = [string]$eng.Cells.Item($r, 1).Value2
                if (-not $nome -or -not $analitosDB.ContainsKey($nome)) { continue }
                $situacao = [string]$eng.Cells.Item($r, 9).Value2
                $f = $an.Range('A4:A43').Find($nome)
                $etLegado = if ($null -ne $f) { $an.Cells.Item($f.Row, 18).Value2 } else { $null }
                if ($situacao -eq 'FALLBACK LEGADO') {
                    $fallbackDB++
                    $etEng = $eng.Cells.Item($r, 6).Value2
                    if (-not (Is-IgualNumero $etLegado $etEng)) { $fallbackDivergente++ }
                }
            }

            $refsEng = 0
            foreach ($ws in @($wb.Worksheets)) {
                $cel = $null
                try { $cel = $ws.Cells.Find('engETp', [System.Reflection.Missing]::Value, -4123, 2) } catch { }
                if ($null -ne $cel) { $refsEng++ }
            }

            $nomesVBA = @{}
            foreach ($comp in @($wb.VBProject.VBComponents)) { $nomesVBA[$comp.Name] = $true }
            $estrutura = [PSCustomObject]@{
                FontePadraoVazia = ([string]$cfg.Range('B2').Value2 -eq '')
                RigorPadraoVazio = ([string]$cfg.Range('B3').Value2 -eq '')
                FontesCfg = $fontesCfg
                LinhasDB = $linhasDB
                FallbackDB = $fallbackDB
                FallbackDivergente = $fallbackDivergente
                ReferenciasEngETp = $refsEng
                TemModulo = $nomesVBA.ContainsKey('mEspecificacoes')
                TemFormulario = $nomesVBA.ContainsKey('frmEspecificacoes')
            }
        }

        return [PSCustomObject]@{ Dados = $dados; Estrutura = $estrutura }
    }
    finally {
        try { $wb.Close($false) } catch { }
        try { $xl.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    }
}

function Is-IgualNumero($A, $B) {
    if (($null -eq $A -or "$A" -eq '') -and ($null -eq $B -or "$B" -eq '')) { return $true }
    try { $da = [double]$A } catch { return $false }
    try { $db = [double]$B } catch { return $false }
    return ([Math]::Abs($da - $db) -le 0.000000000001)
}

$antes = Foto-BI $Referencia $false
$depois = Foto-BI $Candidato $true

$difVeredito = 0
$difSigma = 0
$nVerAntes = 0
$nVerDepois = 0
$nSigmaAntes = 0
$nSigmaDepois = 0
$detalhes = New-Object System.Collections.ArrayList

foreach ($chave in $antes.Dados.Keys) {
    $a = $antes.Dados[$chave]
    if ($a.Veredito) { $nVerAntes++ }
    if ($null -ne $a.Sigma -and "$($a.Sigma)" -ne '') { $nSigmaAntes++ }
    if (-not $depois.Dados.ContainsKey($chave)) {
        $difVeredito++
        if ($detalhes.Count -lt 50) { [void]$detalhes.Add("CHAVE_AUSENTE;$chave") }
        continue
    }
    $d = $depois.Dados[$chave]
    if ($a.Veredito -ne $d.Veredito) {
        $difVeredito++
        if ($detalhes.Count -lt 50) { [void]$detalhes.Add("VEREDITO;$chave;$($a.Veredito);$($d.Veredito)") }
    }
    if (-not (Is-IgualNumero $a.Sigma $d.Sigma)) {
        $difSigma++
        if ($detalhes.Count -lt 50) { [void]$detalhes.Add("SIGMA;$chave;$($a.Sigma);$($d.Sigma)") }
    }
}
foreach ($d in $depois.Dados.Values) {
    if ($d.Veredito) { $nVerDepois++ }
    if ($null -ne $d.Sigma -and "$($d.Sigma)" -ne '') { $nSigmaDepois++ }
}

$e = $depois.Estrutura
$okEstrutura = $e.FontePadraoVazia -and $e.RigorPadraoVazio -and $e.FontesCfg -eq 0 -and `
                $e.LinhasDB -eq 0 -and $e.FallbackDB -eq 15 -and $e.FallbackDivergente -eq 0 -and `
                $e.ReferenciasEngETp -eq 0 -and $e.TemModulo -and $e.TemFormulario
$okVeredito = ($nVerAntes -eq 1575 -and $nVerDepois -eq 1575 -and $difVeredito -eq 0)
$okSigma = ($nSigmaAntes -eq 1125 -and $nSigmaDepois -eq 1125 -and $difSigma -eq 0)

$linhas = @(
    'Metrica;Esperado;Referencia;Candidato;Diferencas;Status',
    "Westgard_Veredito;1575;$nVerAntes;$nVerDepois;$difVeredito;$(if($okVeredito){'OK'}else{'FALHA'})",
    "Sigma;1125;$nSigmaAntes;$nSigmaDepois;$difSigma;$(if($okSigma){'OK'}else{'FALHA'})",
    "Fallback_QR_DB;15;15;$($e.FallbackDB);$($e.FallbackDivergente);$(if($e.FallbackDB -eq 15 -and $e.FallbackDivergente -eq 0){'OK'}else{'FALHA'})",
    "DB_formal_vazio;0;0;$($e.LinhasDB);0;$(if($e.LinhasDB -eq 0){'OK'}else{'FALHA'})",
    "Consumidores_engETp;0;0;$($e.ReferenciasEngETp);0;$(if($e.ReferenciasEngETp -eq 0){'OK'}else{'FALHA'})",
    "Estrutura_formal;OK;NA;$(if($okEstrutura){'OK'}else{'FALHA'});0;$(if($okEstrutura){'OK'}else{'FALHA'})"
)
if ($detalhes.Count -gt 0) { $linhas += ''; $linhas += 'Tipo;Chave;Antes;Depois'; $linhas += $detalhes }
$dir = Split-Path -Parent $OutCsv
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
[System.IO.File]::WriteAllLines($OutCsv, $linhas, (New-Object System.Text.UTF8Encoding($true)))

"Westgard: $nVerAntes -> $nVerDepois, diferencas $difVeredito"
"Sigma: $nSigmaAntes -> $nSigmaDepois, diferencas $difSigma"
"Fallback Q/R nos analitos com dados: $($e.FallbackDB), divergencias $($e.FallbackDivergente)"
"Relatorio: $OutCsv"

if (-not $okVeredito -or -not $okSigma -or -not $okEstrutura) {
    throw 'Teste de especificacoes da Hematologia reprovado.'
}
