# ajustes_usabilidade.ps1 - view que se atualiza sozinha + formularios mais altos
#
# ITEM 1 - a aba Resultados nao refletia o DB_Resultados
#
# Resultados NAO e uma view viva: e uma COPIA materializada. O VBA le o banco,
# monta a matriz e despeja nos blocos por nivel. Nao ha formula ligando as duas,
# entao apagar linha no banco nao muda nada na aba ate alguem chamar
# AtualizarViewResultados.
#
# A escolha de materializar e correta -- 1.000+ linhas com formula de busca
# travariam a planilha. O erro era a aba nao se atualizar ao ser aberta: uma
# tela que se apresenta como reflexo do banco e nao reflete ensina o usuario a
# desconfiar do sistema. Worksheet_Activate resolve sem custo: so paga quando a
# aba e aberta.
#
# ITEM 3 - formularios cortando palavras
#
# Altura do formulario e das legendas. Crescer so o formulario nao resolve
# legenda cortada; crescer so a legenda nao resolve controle cortado embaixo.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\ajustes_usabilidade.ps1 -Workbook <x.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [int]$CrescerForm = 34,
    [int]$CrescerLegenda = 6
)

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
    if ($wb.ProtectStructure) { $wb.Unprotect($SENHA) }
    $vbp = $wb.VBProject

    # ---- ITEM 1: Resultados se atualiza ao ser aberta -------------------
    $wsR = $wb.Worksheets.Item('Resultados')
    $compR = $null
    foreach ($c in $vbp.VBComponents) {
        if ($c.Type -eq 100 -and $c.Properties.Item('Name').Value -eq 'Resultados') { $compR = $c }
    }
    if ($compR -eq $null) { throw "modulo de documento da aba Resultados nao encontrado" }
    $cmR = $compR.CodeModule
    $txtR = ''
    if ($cmR.CountOfLines -gt 0) { $txtR = $cmR.Lines(1, $cmR.CountOfLines) }
    if ($txtR -match 'Worksheet_Activate') {
        "Resultados: Worksheet_Activate ja existia"
    }
    else {
        # NAO consultar ProcStartLine aqui: quando a rotina ainda nao existe o
        # COM lanca "Sub ou Function nao definida", e o retorno nem era usado.
        $codigo = @(
            "",
            "' A aba Resultados e uma COPIA do DB_Resultados, nao uma view viva:",
            "' nenhuma formula liga as duas. Sem isto, apagar ou corrigir linha no",
            "' banco nao aparecia aqui ate alguem clicar em Atualizar View -- e uma",
            "' tela que se diz reflexo do banco e nao reflete ensina o usuario a",
            "' desconfiar do sistema.",
            "Private Sub Worksheet_Activate()",
            "    On Error GoTo fim",
            "    Application.EnableEvents = False",
            "    AtualizarViewResultados",
            "fim:",
            "    Application.EnableEvents = True",
            "End Sub"
        ) -join "`r`n"
        $cmR.InsertLines(($cmR.CountOfLines + 1), $codigo)
        "Resultados: Worksheet_Activate instalado (atualiza a view ao abrir a aba)"
    }

    # ---- ITEM 3: formularios mais altos ---------------------------------
    $ajustados = @()
    foreach ($c in $vbp.VBComponents) {
        if ($c.Type -ne 3) { continue }        # 3 = vbext_ct_MSForm
        $des = $c.Designer
        if ($des -eq $null) { continue }
        $hAntes = $c.Properties.Item('Height').Value
        $c.Properties.Item('Height').Value = $hAntes + $CrescerForm

        # legenda cortada: cresce a altura do rotulo, nao a largura
        $nLbl = 0
        foreach ($ctl in $des.Controls) {
            $tipo = ''
            try { $tipo = $ctl.Name } catch { }
            $ehLabel = $false
            try { $ehLabel = ($ctl.Font -ne $null -and $ctl.WordWrap -ne $null) } catch { }
            if ($ehLabel) {
                try { $ctl.Height = $ctl.Height + $CrescerLegenda; $nLbl++ } catch { }
            }
        }
        $ajustados += "$($c.Name): altura $hAntes -> $($c.Properties.Item('Height').Value), $nLbl rotulo(s)"
    }
    if ($ajustados.Count -eq 0) { throw "nenhum UserForm encontrado" }
    "formularios ajustados ($($ajustados.Count)):"
    $ajustados | ForEach-Object { "  $_" }

    if ($wb.ProtectStructure -eq $false) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
