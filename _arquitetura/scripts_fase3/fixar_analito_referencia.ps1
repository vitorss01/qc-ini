# fixar_analito_referencia.ps1 - deixa o Painel num analito conhecido
#
# POR QUE ISTO EXISTE
#
# O Calc nao e uma area por analito: e a area de UM analito por vez, o que o
# spinner do Painel selecionou. Consequencia: o valor calculado de ~660 celulas
# do Calc, e o conteudo inteiro do Eng_Saida, dependem de qual analito estava
# selecionado quando o arquivo foi salvo.
#
# A linha de base de formulas guarda valores calculados. Enquanto a Glicose foi
# o analito nº 1, isso passou despercebido. Ao reordenar a aba Analitos
# (05/08/2026) a posicao 1 virou Lactato, que ainda nao tem resultado -- e a
# suite acusou 663 divergencias no Calc e o Eng_Saida vazio, sem que nada
# estivesse quebrado.
#
# Um controle que muda de resposta conforme o estado da interface nao serve para
# detectar regressao: ou ele e deterministico, ou vira ruido que se aprende a
# ignorar -- que e como um gate morre.
#
# Este script fixa o analito de referencia POR NOME antes da extracao. Nome, e
# nao posicao, justamente porque a posicao e o que muda.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\fixar_analito_referencia.ps1 -Workbook <x.xlsm> [-Analito Glicose]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Analito = 'Glicose'
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
    $pa = $wb.Worksheets.Item('Painel')
    $protPa = $pa.ProtectContents
    if ($protPa) { try { $pa.Unprotect($SENHA) } catch { } }

    $an = $wb.Worksheets.Item('Analitos')
    $pos = 0
    for ($i = 4; $i -le 43; $i++) {
        $v = $an.Cells.Item($i, 1).Value2
        if ($v -ne $null -and "$v".Trim() -eq $Analito) { $pos = $i - 3; break }
    }
    if ($pos -eq 0) { throw "analito de referencia '$Analito' nao esta cadastrado na aba Analitos" }

    $pa.Range('B3').Value2 = [double]$pos
    $xl.CalculateFullRebuild()
    $lido = "$($pa.Range('C3').Value2)"
    if ($lido -ne $Analito) { throw "Painel ficou em '$lido', esperado '$Analito'" }
    "analito de referencia: $Analito (posicao $pos)"

    if ($protPa) { $pa.Protect($SENHA, $true, $true, $true, $true) }
    if ($wb.ProtectStructure -eq $false) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    $salvou = $true
}
finally {
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
