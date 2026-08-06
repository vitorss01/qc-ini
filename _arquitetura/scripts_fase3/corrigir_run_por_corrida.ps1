# corrigir_run_por_corrida.ps1 - o RUN volta a identificar a CORRIDA
#
# O QUE ESTAVA ERRADO
#
# ADR-004 e ADR-011: o RUN identifica a CORRIDA -- data + lote -- e os niveis da
# mesma corrida COMPARTILHAM o RUN. A chave logica e RUN + Analito + Nivel.
#
# AtribuirRUNs montava a chave da corrida com o CODIGO do lote, e o codigo
# EMBUTE O NIVEL: CodigoLote("8974", 1) = "QC-897401" e CodigoLote("8974", 2) =
# "QC-897402". Extraindo 6 caracteres saem "897401" e "897402" -- chaves
# diferentes para a MESMA corrida. Resultado: cada nivel ganhava o seu proprio
# RUN, e 18 corridas viravam 36.
#
# No grafico isso aparece como os niveis plotados em faixas separadas do eixo X
# (1..18 e 19..36) em vez de sobrepostos na mesma corrida.
#
# A conta certa e o NUCLEO do lote, que nao tem o nivel: NucleoLote("QC-897401")
# = "8974" para os dois niveis.
#
# DUAS METADES QUE PRECISAM CONCORDAR
#
# AtribuirRUNs monta a chave DUAS vezes -- uma lendo o banco (corridas que ja
# existem) e outra para os registros novos. Se as duas nao usarem a MESMA
# funcao, nada casa e todo registro novo ganha RUN inedito. Foi o que aconteceu
# quando so a primeira foi corrigida.
#
# MIGRACAO DO QUE JA ESTA GRAVADO
#
# Os registros que ja receberam RUN separado por nivel sao remapeados para o
# menor RUN da corrida. Nao e reuso de RUN (o contador nunca volta): e correcao
# de uma atribuicao errada. Os numeros que ficam sem uso permanecem sem uso.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\corrigir_run_por_corrida.ps1 -Workbook ..\..\QC_Bioquimica.xlsm [-Simular]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [switch]$Simular
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $u = $_
            if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }
            Start-Sleep -Seconds ($t * 2)
        }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { try { $wb.Close($false) } catch { }; $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvar = $false
try {
    $estruturaEstava = $wb.ProtectStructure
    if ($estruturaEstava) { $wb.Unprotect($SENHA) }

    # ------------------------------------------------------------- codigo ---
    $mod = $null
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq 'mImportar') { $mod = $c; break } }
    if ($mod -eq $null) { throw 'mImportar ausente' }
    $cm = $mod.CodeModule
    $linhas = $cm.Lines(1, $cm.CountOfLines) -split "`r?`n"

    $trocou = 0
    for ($i = 0; $i -lt $linhas.Count; $i++) {
        if ($linhas[$i] -match '^\s*''') { continue }
        if ($linhas[$i] -notmatch 'regs\(i, 4\)') { continue }
        $novo = [regex]::Replace($linhas[$i], 'Mid\$?\(\s*(CStr\(regs\(i, 4\)\))\s*,\s*4\s*,\s*6\s*\)', 'NucleoLote($1)')
        if ($novo -ne $linhas[$i]) { $linhas[$i] = $novo; $trocou++; "codigo linha $($i + 1): $($novo.Trim())" }
    }
    if ($trocou -eq 0) {
        "codigo: nenhuma troca necessaria (as duas metades ja usam NucleoLote)"
    }
    elseif (-not $Simular) {
        $cm.DeleteLines(1, $cm.CountOfLines)
        $cm.AddFromString(($linhas -join "`r`n"))
    }

    # ------------------------------------------------------------- dados ----
    $db = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq 'DB_Resultados') { $db = $w; break } }
    if ($db.ProtectContents) { try { $db.Unprotect($SENHA) } catch { } }
    $ult = $db.Cells($db.Rows.Count, 1).End(-4162).Row
    if ($ult -lt 4) { throw 'banco vazio' }

    # le tudo de uma vez
    $vals = $db.Range($db.Cells(4, 1), $db.Cells($ult, 7)).Value2
    $mapa = @{}     # chave da corrida -> menor RUN
    for ($r = 1; $r -le ($ult - 3); $r++) {
        $run = $vals[$r, 1]; $dt = $vals[$r, 2]; $lote = "$($vals[$r,4])".Trim()
        if ("$run" -eq '' -or "$dt" -eq '' -or $lote.Length -lt 6) { continue }
        $nucleo = $lote.Substring(3, $lote.Length - 5)
        $k = "{0}|{1}" -f [int][double]$dt, $nucleo
        $ri = [int][double]$run
        if (-not $mapa.ContainsKey($k) -or $ri -lt $mapa[$k]) { $mapa[$k] = $ri }
    }

    $mudar = @()
    for ($r = 1; $r -le ($ult - 3); $r++) {
        $run = $vals[$r, 1]; $dt = $vals[$r, 2]; $lote = "$($vals[$r,4])".Trim()
        if ("$run" -eq '' -or "$dt" -eq '' -or $lote.Length -lt 6) { continue }
        $nucleo = $lote.Substring(3, $lote.Length - 5)
        $k = "{0}|{1}" -f [int][double]$dt, $nucleo
        $alvo = $mapa[$k]
        if ([int][double]$run -ne $alvo) { $mudar += @{ Linha = $r + 3; De = [int][double]$run; Para = $alvo } }
    }

    "corridas distintas (data + nucleo do lote): $($mapa.Count)"
    "registros a remapear: $($mudar.Count)"
    if ($mudar.Count -gt 0) {
        $amostra = $mudar | Select-Object -First 3
        foreach ($m in $amostra) { "   linha $($m.Linha): RUN $($m.De) -> $($m.Para)" }
    }

    if ($Simular) { "SIMULACAO -- nada foi gravado."; $wb.Close($false); $xl.Quit(); exit 0 }

    foreach ($m in $mudar) { $db.Cells($m.Linha, 1).Value2 = [double]$m.Para }

    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    # -------------------------------------------------------- conferencia ---
    $erros = @()
    $vals2 = $db.Range($db.Cells(4, 1), $db.Cells($ult, 7)).Value2
    $porData = @{}
    $chaves = @{}
    for ($r = 1; $r -le ($ult - 3); $r++) {
        $run = "$($vals2[$r,1])".Trim(); $dt = $vals2[$r, 2]
        $niv = "$($vals2[$r,3])".Trim(); $an = "$($vals2[$r,5])".Trim(); $st = "$($vals2[$r,7])".Trim()
        if ($run -eq '' -or "$dt" -eq '') { continue }
        $kd = "{0}|{1}" -f [int][double]$dt, $niv
        $porData[$kd] = [int][double]$run
        if ($st -eq 'Ativo') {
            $kc = "$run|$an|$niv"
            if ($chaves.ContainsKey($kc)) { $erros += "chave duplicada apos o remapeamento: $kc" }
            else { $chaves[$kc] = $true }
        }
    }
    $datas = @($porData.Keys | ForEach-Object { $_.Split('|')[0] } | Select-Object -Unique)
    $divergentes = 0
    foreach ($d in $datas) {
        $r1 = if ($porData.ContainsKey("$d|1")) { $porData["$d|1"] } else { $null }
        $r2 = if ($porData.ContainsKey("$d|2")) { $porData["$d|2"] } else { $null }
        if ($r1 -ne $null -and $r2 -ne $null -and $r1 -ne $r2) { $divergentes++ }
    }
    if ($divergentes -gt 0) { $erros += "$divergentes data(s) ainda com RUN diferente entre os niveis" }

    $runsDistintos = (@($porData.Values) | Select-Object -Unique).Count
    if ($runsDistintos -ne $datas.Count) {
        $erros += "RUNs distintos ($runsDistintos) diferente de datas distintas ($($datas.Count))"
    }

    if ($erros.Count -gt 0) {
        $erros | Select-Object -First 8 | ForEach-Object { "  FALHA: $_" }
        throw "Correcao rejeitada: $($erros.Count) problema(s). Nada foi salvo."
    }
    "conferencia: $runsDistintos RUN(s) para $($datas.Count) data(s); niveis compartilhando RUN; nenhuma chave RUN+Analito+Nivel duplicada"

    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $salvar = $true
    $wb.Save()
    "Salvo: $Workbook"
}
finally {
    try { if ($salvar) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
}
