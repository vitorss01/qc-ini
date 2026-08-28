# testar_especificacoes_hematologia.ps1 - prova estrutural e de nao regressao
#
# Compara a tabela fato antes/depois da instalacao formal. Os vereditos
# Westgard e os Sigma da REFERENCIA (nao um numero fixo de checkpoint antigo)
# devem permanecer identicos no CANDIDATO -- referencia e candidato tem de
# bater EXATAMENTE entre si, o valor absoluto e o que o motor produzir. A
# prova roda os motores em memoria e fecha os arquivos sem salvar.
#
# Fallback_QR_DB segue a mesma logica: o numero certo nao e fixo (cresce com
# o roster de analitos), o que importa e o invariante -- com
# Cfg_Especificacoes vazia, todo elegivel com ETp publicado cai em fallback
# legado (nem mais nem menos), e nenhum deles pode ter herdado "Lim CV VB %"
# em vez de "ETp final %" (o defeito historico desta area).
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

            # FALLBACK: A CONTAGEM E CONSEQUENCIA, NAO A VERDADE. A VERDADE E A FONTE.
            #
            # $fallbackDivergente ja e a prova semantica que importa: compara o
            # ETp PUBLICADO contra Analitos! coluna 18 ("ETp final %", resolvida
            # pela hierarquia oficial) -- nunca a coluna 20 ("Lim CV VB %"), que
            # foi o defeito historico desta area (CV interpretado como ETp).
            # Resolvido por LABEL, nao por posicao fixa: se as colunas mudarem
            # de lugar um dia, o teste para de achar o rotulo em vez de comparar
            # a coisa errada em silencio.
            #
            # O NUMERO de fallbacks (15, 28, o que for) NAO E uma verdade fixa.
            # Cfg_Especificacoes esta VAZIA nesta etapa do build (nenhuma fonte
            # formal cadastrada -- $fontesCfg abaixo prova isso), e por
            # construcao TODO analito elegivel tem de cair em fallback legado
            # ate alguem popular o Cfg. Fixar em "15" e travar o teste na
            # fotografia de 18/08 (o checkpoint anterior a esta funcionalidade
            # nem existir): o roster de analitos da Hematologia cresceu desde
            # entao, e crescer o fallback JUNTO e o comportamento correto, nao
            # regressao.
            $cRot = @{}
            for ($c = 1; $c -le $an.Cells.Item(3, $an.Columns.Count).End(-4159).Column; $c++) {
                $r = [string]$an.Cells.Item(3, $c).Value2
                if ($r) { $cRot[$r] = $c }
            }
            $colETpFinal = $cRot['ETp final %']
            $colLimCVvb = $cRot['Lim CV VB %']
            if (-not $colETpFinal) { throw 'rotulo "ETp final %" nao encontrado em Analitos (linha 3)' }

            $fallbackDB = 0
            $fallbackDivergente = 0
            $fallbackUsouColunaErrada = 0
            foreach ($nome in $analitosDB.Keys) {
                $f = $an.Range('A4:A43').Find($nome)
                if ($null -eq $f) { continue }
                $r = $f.Row
                $rEng = ($eng.Range('A4:A43').Find($nome))
                if ($null -eq $rEng) { continue }
                $situacao = [string]$eng.Cells.Item($rEng.Row, 9).Value2
                if ($situacao -ne 'FALLBACK LEGADO') { continue }
                $fallbackDB++
                $etEng = $eng.Cells.Item($rEng.Row, 6).Value2
                $etCorreto = $an.Cells.Item($r, $colETpFinal).Value2
                if (-not (Is-IgualNumero $etCorreto $etEng)) { $fallbackDivergente++ }
                if ($colLimCVvb) {
                    $limCV = $an.Cells.Item($r, $colLimCVvb).Value2
                    if ($null -ne $limCV -and "$limCV" -ne '' -and (Is-IgualNumero $limCV $etEng) -and -not (Is-IgualNumero $etCorreto $limCV)) {
                        $fallbackUsouColunaErrada++
                    }
                }
            }
            $elegiveisComEtp = 0
            foreach ($nome in $analitosDB.Keys) {
                $f = $an.Range('A4:A43').Find($nome)
                if ($null -eq $f) { continue }
                $v = $an.Cells.Item($f.Row, $colETpFinal).Value2
                if ($null -ne $v -and "$v" -ne '') { $elegiveisComEtp++ }
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
                FallbackUsouColunaErrada = $fallbackUsouColunaErrada
                ElegiveisComEtp = $elegiveisComEtp
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
# Fallback_QR_DB: a VERDADE nao e um numero congelado em 18/08 (antes da
# funcionalidade de especificacoes sequer existir) -- e o invariante
# estrutural: com Cfg_Especificacoes vazia, todo analito elegivel com ETp
# publicado TEM de estar em fallback legado (nem mais, nem menos), e nenhum
# deles pode ter herdado o valor da coluna errada ("Lim CV VB %", o defeito
# historico). Contagem sobe com o roster -- isso e crescimento, nao regressao.
$okFallback = ($e.FallbackDB -eq $e.ElegiveisComEtp -and $e.FallbackDivergente -eq 0 -and `
                $e.FallbackUsouColunaErrada -eq 0 -and $e.FallbackDB -gt 0)
$okEstrutura = $e.FontePadraoVazia -and $e.RigorPadraoVazio -and $e.FontesCfg -eq 0 -and `
                $e.LinhasDB -eq 0 -and $okFallback -and `
                $e.ReferenciasEngETp -eq 0 -and $e.TemModulo -and $e.TemFormulario
# Westgard/Sigma: o numero certo e o que o MOTOR produziu na propria
# referencia (nao um literal congelado de checkpoint antigo) -- o gate real
# e referencia == candidato, ambos > 0 (nao aceitar "bateu porque os dois
# vieram vazios").
$okVeredito = ($nVerAntes -eq $nVerDepois -and $difVeredito -eq 0 -and $nVerAntes -gt 0)
$okSigma = ($nSigmaAntes -eq $nSigmaDepois -and $difSigma -eq 0 -and $nSigmaAntes -gt 0)

$linhas = @(
    'Metrica;Esperado;Referencia;Candidato;Diferencas;Status',
    "Westgard_Veredito;=Referencia;$nVerAntes;$nVerDepois;$difVeredito;$(if($okVeredito){'OK'}else{'FALHA'})",
    "Sigma;=Referencia;$nSigmaAntes;$nSigmaDepois;$difSigma;$(if($okSigma){'OK'}else{'FALHA'})",
    "Fallback_QR_DB;=ElegiveisComEtp($($e.ElegiveisComEtp));NA;$($e.FallbackDB);$($e.FallbackDivergente);$(if($okFallback){'OK'}else{'FALHA'})",
    "Fallback_ColunaErrada;0;NA;$($e.FallbackUsouColunaErrada);$($e.FallbackUsouColunaErrada);$(if($e.FallbackUsouColunaErrada -eq 0){'OK'}else{'FALHA'})",
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
"Fallback Q/R nos analitos com dados: $($e.FallbackDB) (elegiveis c/ ETp: $($e.ElegiveisComEtp)), divergencias fonte $($e.FallbackDivergente), coluna errada $($e.FallbackUsouColunaErrada)"
"Relatorio: $OutCsv"

if (-not $okVeredito -or -not $okSigma -or -not $okEstrutura) {
    throw 'Teste de especificacoes da Hematologia reprovado.'
}
