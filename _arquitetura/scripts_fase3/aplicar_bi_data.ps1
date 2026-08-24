# aplicar_bi_data.ps1 - etapa de BUILD da camada BI (ADR-026)
#
# Cria a aba BI_Data, monta a tabela estruturada tblBI_Fato e RECONCILIA com a
# aba Calc antes de deixar passar.
#
# A reconciliacao nao e cosmetica: a camada BI recalcula Z e Westgard a partir
# do mesmo dado que o motor usa. Se as duas contas divergirem, uma das duas esta
# errada, e num sistema de CQ nao da para saber qual sem investigar. O script
# FALHA quando ha divergencia -- artefato com BI divergente nao e entregavel.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\aplicar_bi_data.ps1 -Workbook <alvo.xlsm>

param([Parameter(Mandatory = $true)][string]$Workbook)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

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

try {
    $estruturaEstava = $wb.ProtectStructure
    if ($estruturaEstava) { $wb.Unprotect($SENHA) }

    $temBI = $false
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq 'mBI') { $temBI = $true } }
    if (-not $temBI) { throw "mBI ausente -- aplicar_vba.ps1 precisa importa-lo antes desta etapa" }
    "mBI presente"

    $xl.Run('AtualizarBIData') | Out-Null
    "AtualizarBIData executada"

    $ws = $wb.Worksheets.Item('BI_Data')
    $ult = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    $nLin = $ult - 1
    "BI_Data: $nLin linha(s) de fato"
    if ($nLin -lt 1) { throw "BI_Data vazia" }

    # tabela estruturada, e nao faixa: o Power Query referencia pelo NOME
    $lo = $null
    foreach ($t in $ws.ListObjects) { if ($t.Name -eq 'tblBI_Fato') { $lo = $t } }
    if ($lo -eq $null) { throw "ListObject tblBI_Fato nao foi criado" }
    "tabela estruturada tblBI_Fato: $($lo.Range.Address()) ($($lo.ListRows.Count) linhas x $($lo.ListColumns.Count) colunas)"
    # 84 desde o ADR-040 (era 76 no ADR-035, 65 no ADR-033, 60 no ADR-026).
    # As oito novas fecham o plano de CQ: Sigma_Plano e Nivel_Governante, que
    # materializam o "pior nivel governa", a classificacao desse Sigma e as
    # cinco booleanas Usar_*, para o BI acender as regras sem procurar
    # substring dentro de uma cadeia de texto.
    if ($lo.ListColumns.Count -ne 84) { throw "esperadas 84 colunas, encontradas $($lo.ListColumns.Count)" }

    # chaves: ID_Result tem de ser UNICO -- e a granularidade declarada
    $dup = $xl.WorksheetFunction.SumProduct(
        $xl.WorksheetFunction.CountIf($lo.ListColumns.Item(1).DataBodyRange, $lo.ListColumns.Item(1).DataBodyRange)) - $nLin
    if ($dup -ne 0) { throw "ID_Result nao e unico: $dup repeticao(oes)" }
    "ID_Result unico nas $nLin linhas"

    # ---- reconciliacao com o motor --------------------------------------
    $xl.Run('AtualizarOperacao') | Out-Null
    $r = $xl.Run('ReconciliarComCalc')
    "ReconciliarComCalc: $r"
    $p = $r -split '\|'
    if ([int]$p[0] -lt 1) { throw "reconciliacao nao comparou nada -- prova vazia nao e prova" }
    if ($p[1] -ne '0') { throw "BI diverge do motor em $($p[1]) ponto(s): $($p[2])" }
    "reconciliacao: $($p[0]) pontos comparados, 0 divergencias"

    $ws.Visible = 0        # xlSheetHidden: e camada de dados, nao tela
    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
