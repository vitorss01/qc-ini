# criar_especificacoes.ps1 - modulo de Especificacoes de Qualidade (ADR-022)
#
# Cria DB_Especificacoes (banco tabular) e Cfg_Especificacoes (fontes + padroes),
# e MIGRA o que ja existe nas colunas K..O da aba Analitos.
#
# POR QUE MIGRAR, E NAO CONVIVER
#
# A aba Analitos ja implementa a matematica certa (S = TEa/3, T = CVi*fi, R
# resolvido por cascata, fatores em U:W). O que esta errado e o MODELO:
#   - nao ha dimensao ANO, entao uma corrida de 2025 reaberta em 2030 seria
#     julgada pela meta de 2030;
#   - as specs moram no armazem POR LOTE (aInput = E4:P43, que inclui K..P),
#     mas TEa CLIA e constante regulatoria e CVi/CVg sao constantes biologicas
#     -- nao sao propriedades do lote;
#   - nao ha escolha de fonte, e o Fabricante nao participa.
#
# Manter as duas implementacoes criaria DUAS fontes de verdade para o ETp --
# exatamente a divergencia que o ADR-019 existe para impedir.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_especificacoes.ps1 -Workbook <x.xlsm> [-AnoBase 2026]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [int]$AnoBase = 2026
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

function Norm { param([string]$s)
    if (-not $s) { return '' }
    $n = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $n.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne 'NonSpacingMark') { [void]$sb.Append($ch) }
    }
    return $sb.ToString().ToUpperInvariant()
}
function Aba { param($Pasta, [string]$Nome)
    $alvo = Norm $Nome
    foreach ($ws in @($Pasta.Worksheets)) { if ((Norm $ws.Name) -eq $alvo) { return $ws } }
    return $null
}
function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch { $u = $_; if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }; Start-Sleep -Seconds ($t * 2) }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

$salvou = $false
$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $Workbook" }
$xl.Calculation = -4135

try {
    if ($wb.ProtectStructure) { $wb.Unprotect($SENHA) }
    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.ProtectContents) { try { $ws.Unprotect($SENHA) } catch { try { $ws.Unprotect() } catch { } } }
    }

    # ---- Cfg_Especificacoes: fontes e padroes ---------------------------
    $cfg = Aba $wb 'Cfg_Especificacoes'
    if ($cfg -ne $null) { $cfg.Visible = -1; $cfg.Delete() }
    $cfg = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $cfg.Name = 'Cfg_Especificacoes'

    $cfg.Range('A1').Value2 = 'ESPECIFICACOES DE QUALIDADE - CONFIGURACAO'
    $cfg.Range('A1').Font.Bold = $true
    $cfg.Range('A1').Font.Size = 13
    $cfg.Range('A2').Value2 = 'Fonte padrao'
    $cfg.Range('B2').Value2 = 'CLIA'
    $cfg.Range('A3').Value2 = 'Rigor padrao (VB)'
    $cfg.Range('B3').Value2 = 'MIN'
    $cfg.Range('A4').Value2 = 'Ano de referencia (0 = usar o ano do resultado)'
    $cfg.Range('B4').Value2 = [double]0
    $cfg.Range('A2:A4').Font.Bold = $true

    $cfg.Range('A6').Value2 = 'FONTE'
    $cfg.Range('B6').Value2 = 'MODELO'
    $cfg.Range('C6').Value2 = 'O que a fonte informa'
    $cfg.Range('A6:C6').Font.Bold = $true
    $cfg.Range('A6:C6').Interior.Color = 14277081

    # As tres fontes de partida. Ricos, EFLM, CAP, RCPA, Rilibak entram como
    # LINHA aqui, escolhendo um dos tres modelos -- sem tocar em codigo.
    $fontes = @(
        @('CLIA', 'ETP_DIRETO', 'ETp (%). CVtp = ETp/3'),
        @('Variacao Biologica', 'VB', 'CVi, CVg e rigor. Deriva CVtp, BIAStp e ETp'),
        @('Fabricante', 'CV_BIAS_DIRETO', 'CVtp e BIAStp (%). ETp = BIAStp + 1,65*CVtp')
    )
    for ($i = 0; $i -lt $fontes.Count; $i++) {
        $cfg.Cells.Item(7 + $i, 1).Value2 = [string]$fontes[$i][0]
        $cfg.Cells.Item(7 + $i, 2).Value2 = [string]$fontes[$i][1]
        $cfg.Cells.Item(7 + $i, 3).Value2 = [string]$fontes[$i][2]
    }
    $cfg.Columns.Item(1).ColumnWidth = 24
    $cfg.Columns.Item(2).ColumnWidth = 18
    $cfg.Columns.Item(3).ColumnWidth = 52
    "Cfg_Especificacoes: $($fontes.Count) fontes, padrao = CLIA / MIN"

    # ---- DB_Especificacoes ----------------------------------------------
    $db = Aba $wb 'DB_Especificacoes'
    if ($db -ne $null) { $db.Visible = -1; $db.Delete() }
    $db = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $db.Name = 'DB_Especificacoes'

    $db.Range('A1').Value2 = 'ESPECIFICACOES DE QUALIDADE - BANCO'
    $db.Range('A1').Font.Bold = $true
    $db.Range('A1').Font.Size = 13
    $db.Range('A2').Value2 = 'Uma linha por Ano + Fonte + Analito. A meta vigente e a de maior Ano que nao ultrapasse o ano do RESULTADO -- e por isso que uma corrida de 2025 continua julgada pela meta de 2025 em 2035.'
    $db.Range('A2').Font.Italic = $true

    $cab = @('ID', 'Ano', 'Fonte', 'Modelo', 'Analito', 'ETp %', 'CVi %', 'CVg %',
             'Rigor', 'CVtp %', 'BIAStp %', 'Ativo', 'Usuario', 'Cadastrado em')
    for ($k = 0; $k -lt $cab.Count; $k++) {
        $c = $db.Cells.Item(3, $k + 1)
        $c.Value2 = [string]$cab[$k]
        $c.Font.Bold = $true
    }
    $db.Range($db.Cells.Item(3, 1), $db.Cells.Item(3, $cab.Count)).Interior.Color = 14277081
    $db.Range($db.Cells.Item(3, 1), $db.Cells.Item(3, $cab.Count)).Borders.LineStyle = 1
    $db.Columns.Item(1).ColumnWidth = 13
    $db.Columns.Item(3).ColumnWidth = 20
    $db.Columns.Item(4).ColumnWidth = 17
    $db.Columns.Item(5).ColumnWidth = 26
    $db.Columns.Item(13).ColumnWidth = 14
    $db.Columns.Item(14).ColumnWidth = 17

    # ---- migracao do que ja existe na Analitos --------------------------
    $an = Aba $wb 'Analitos'
    if ($an -eq $null) { throw "aba Analitos nao encontrada" }
    $lin = 4
    $nCLIA = 0; $nVB = 0
    for ($r = 4; $r -le 43; $r++) {
        $nome = $an.Cells.Item($r, 1).Value2
        if ($nome -eq $null -or "$nome".Trim() -eq '') { continue }
        $tea = $an.Cells.Item($r, 11).Value2      # K
        $cvi = $an.Cells.Item($r, 13).Value2      # M
        $cvg = $an.Cells.Item($r, 14).Value2      # N
        $rig = $an.Cells.Item($r, 15).Value2      # O
        if ("$rig".Trim() -eq '') { $rig = 'MIN' }

        if ($tea -ne $null -and "$tea".Trim() -ne '') {
            $db.Cells.Item($lin, 1).Value2 = 'ESP-' + ('{0:D6}' -f ($lin - 3))
            $db.Cells.Item($lin, 2).Value2 = [double]$AnoBase
            $db.Cells.Item($lin, 3).Value2 = 'CLIA'
            $db.Cells.Item($lin, 4).Value2 = 'ETP_DIRETO'
            $db.Cells.Item($lin, 5).Value2 = [string]$nome
            $db.Cells.Item($lin, 6).Value2 = [double]$tea
            $db.Cells.Item($lin, 12).Value2 = 'Sim'
            $db.Cells.Item($lin, 13).Value2 = 'migracao'
            $db.Cells.Item($lin, 14).Value2 = [double](Get-Date).ToOADate()
            $lin++; $nCLIA++
        }
        if ($cvi -ne $null -and "$cvi".Trim() -ne '' -and $cvg -ne $null -and "$cvg".Trim() -ne '') {
            $db.Cells.Item($lin, 1).Value2 = 'ESP-' + ('{0:D6}' -f ($lin - 3))
            $db.Cells.Item($lin, 2).Value2 = [double]$AnoBase
            $db.Cells.Item($lin, 3).Value2 = 'Variacao Biologica'
            $db.Cells.Item($lin, 4).Value2 = 'VB'
            $db.Cells.Item($lin, 5).Value2 = [string]$nome
            $db.Cells.Item($lin, 7).Value2 = [double]$cvi
            $db.Cells.Item($lin, 8).Value2 = [double]$cvg
            $db.Cells.Item($lin, 9).Value2 = [string]"$rig".Trim()
            $db.Cells.Item($lin, 12).Value2 = 'Sim'
            $db.Cells.Item($lin, 13).Value2 = 'migracao'
            $db.Cells.Item($lin, 14).Value2 = [double](Get-Date).ToOADate()
            $lin++; $nVB++
        }
    }
    if ($lin -gt 4) {
        # Value2 grava o serial OLE; sem o formato a coluna mostraria 46000.
        $db.Range($db.Cells.Item(4, 14), $db.Cells.Item($lin - 1, 14)).NumberFormat = 'dd/mm/yyyy hh:mm'
    }
    "DB_Especificacoes: $nCLIA de CLIA + $nVB de Variacao Biologica migradas (ano $AnoBase)"

    if ($nCLIA -eq 0 -and $nVB -eq 0) {
        "  AVISO: nada a migrar -- a aba Analitos esta sem TEa/CVi/CVg (esperado apos a limpeza)"
    }

    $db.Visible = 2      # xlSheetVeryHidden: banco nao e tela
    $cfg.Visible = 2

    $xl.Calculation = -4105
    if ($wb.ProtectStructure -eq $false) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { $xl.Calculation = -4105 } catch { }
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
