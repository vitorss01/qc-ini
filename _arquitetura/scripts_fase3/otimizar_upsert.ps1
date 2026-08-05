# otimizar_upsert.ps1 - troca UpsertResultados pela versao que nao recalcula
#
# Substitui a rotina POR NOME dentro do modulo mDados, direto no projeto VBA do
# arquivo. Por nome, e nao por numero de linha: mDados muda com o tempo e
# amarrar em linha quebraria na proxima alteracao a montante.
#
# O ganho vem de suspender o recalculo durante a gravacao. O DB_Resultados tem
# ~45.000 formulas; com o calculo em automatico cada escrita de celula dispara
# recalculo das dependentes, e uma importacao de 1.200 resultados vira milhares
# de recalculos de dezenas de milhares de formulas.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\otimizar_upsert.ps1 -Workbook <x.xlsm> [-Patch <mDados_UPSERT.txt>] [-Modulo mDados]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Patch,
    [string]$Modulo = 'mDados'
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if (-not $Patch) {
    $Patch = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'src_producao\mDados_UPSERT.txt'
}
$Patch = (Resolve-Path -LiteralPath $Patch).ProviderPath
$SENHA = 'qcini2025'

# O .txt e ASCII por construcao (acento em fonte VBA ja corrompeu modulo neste
# projeto). Lido como UTF-8 mesmo assim, para o caso de um dia deixar de ser.
$linhas = [System.IO.File]::ReadAllLines($Patch, [System.Text.Encoding]::UTF8)
$corpo = @()
foreach ($l in $linhas) { if ($l -notmatch "^'@@FUNC") { $corpo += $l } }
$novoTexto = ($corpo -join "`r`n")
if ($novoTexto -notmatch 'Public Function UpsertResultados') { throw "patch nao contem UpsertResultados" }

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

    $comp = $null
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq $Modulo) { $comp = $c } }
    if ($comp -eq $null) { throw "modulo $Modulo nao encontrado" }
    $cm = $comp.CodeModule

    $ini = $cm.ProcStartLine('UpsertResultados', 0)
    $n = $cm.ProcCountLines('UpsertResultados', 0)
    if ($ini -le 0) { throw "UpsertResultados nao encontrado em $Modulo" }
    "antes : $Modulo com $($cm.CountOfLines) linhas; UpsertResultados em L$ini ($n linhas)"

    $cm.DeleteLines($ini, $n)
    $cm.InsertLines($ini, $novoTexto)
    "depois: $Modulo com $($cm.CountOfLines) linhas"

    # conferencia: a rotina existe uma vez so e traz as tres medidas
    $txt = $cm.Lines(1, $cm.CountOfLines)
    $qtd = ([regex]::Matches($txt, 'Function\s+UpsertResultados\s*\(')).Count
    if ($qtd -ne 1) { throw "UpsertResultados aparece $qtd vezes apos o patch" }
    foreach ($marca in @('xlCalculationManual', 'minTocada', 'calcAntes')) {
        if ($txt -notmatch [regex]::Escape($marca)) { throw "patch aplicado sem '$marca'" }
    }
    "conferencia: ok - rotina unica, calculo suspenso, escrita por faixa"

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
