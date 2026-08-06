# restaurar_bloco_eqc.ps1 - devolve o bloco de bias do EQC a aba Estatistica
#
# O QUE SE PERDEU, E COMO
#
# A aba Estatistica guardava em P:U (linhas 7..86) as medias de bias vindas de
# EQC_Dados, por nivel. O novo layout ocupa P e Q com "ETP VB %" e "Status ETP",
# entao o bloco tinha de mudar de endereco para AA:AF.
#
# montar_estatistica_v2.ps1 fazia isso com uma guarda de idempotencia ingenua:
# "se AA6 tem conteudo, ja mudei". A guarda olha o DESTINO para decidir sobre a
# ORIGEM. Entre execucoes sucessivas -- algumas salvando, outras recusadas pela
# conferencia -- houve uma passagem em que AA6 estava vazio E P:U JA CONTINHA AS
# COLUNAS NOVAS. A copia entao levou as colunas novas para AA:AF e apagou P:U,
# e o bloco do EQC deixou de existir: 480 formulas.
#
# Quem pegou foi a prova 4.1 (AUSENTE so na lista fechada), apontando
# Estatistica!T e Estatistica!U -- as duas colunas que ficaram sem destino. Sem
# essa prova o bloco teria sumido em silencio, e so faria falta no dia em que o
# laboratorio cadastrasse o primeiro ensaio de proficiencia.
#
# ORIGEM DA RESTAURACAO: uma copia do arquivo ANTERIOR a mudanca. As formulas
# sao trazidas como TEXTO e reescritas, nao copiadas por objeto, para que o
# destino seja exatamente AA:AF sem herdar referencias de P:U.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\restaurar_bloco_eqc.ps1 -Workbook ..\..\QC_Bioquimica.xlsm -Origem <copia.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [Parameter(Mandatory = $true)][string]$Origem
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$Origem = (Resolve-Path -LiteralPath $Origem).ProviderPath
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

function AbaEstat { param($W) foreach ($x in $W.Worksheets) { if ($x.Name -like 'Estat*') { return $x } } return $null }

$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1

# ---------- 1. le as formulas originais da copia ----------
$wbO = $xl.Workbooks.Open($Origem, 0, $true)
$estO = AbaEstat $wbO
if ($estO -eq $null) { $wbO.Close($false); $xl.Quit(); throw "aba Estatistica ausente na origem" }

$cab = @()
for ($c = 16; $c -le 21; $c++) { $cab += "$($estO.Cells(6, $c).Value2)" }
"cabecalho na origem (P6:U6): $($cab -join ' | ')"

$form = @{}
$nOrig = 0
for ($r = 7; $r -le 86; $r++) {
    for ($c = 16; $c -le 21; $c++) {
        $f = "$($estO.Cells($r, $c).Formula)"
        if ($f -like '=*') { $form["$r|$c"] = $f; $nOrig++ }
    }
}
"formulas lidas da origem: $nOrig"
$wbO.Close($false)

if ($nOrig -lt 400) { $xl.Quit(); throw "a origem tem so $nOrig formulas em P7:U86 -- nao e a versao anterior a mudanca" }
if (($cab -join '') -like '*Status ETP*') { $xl.Quit(); throw "a origem JA tem o layout novo em P:U -- escolha uma copia anterior" }

# ---------- 2. escreve em AA:AF do arquivo de trabalho ----------
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvar = $false
try {
    $estruturaEstava = $wb.ProtectStructure
    if ($estruturaEstava) { $wb.Unprotect($SENHA) }
    $est = AbaEstat $wb
    if ($est.ProtectContents) { try { $est.Unprotect($SENHA) } catch { } }

    $est.Range($est.Cells(6, 27), $est.Cells(86, 32)).ClearContents() | Out-Null
    for ($k = 0; $k -lt 6; $k++) { $est.Cells(6, 27 + $k).Value2 = [string]$cab[$k] }
    $est.Range($est.Cells(6, 27), $est.Cells(6, 32)).Font.Bold = $true

    $escritas = 0
    foreach ($k in $form.Keys) {
        $p = $k -split '\|'
        $r = [int]$p[0]
        $c = [int]$p[1]
        $est.Cells($r, $c + 11).Formula = $form[$k]      # P(16) -> AA(27): +11
        $escritas++
    }
    "formulas escritas em AA7:AF86: $escritas"

    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    # ---------- 3. confere ----------
    $erros = @()
    $comF = 0
    for ($r = 7; $r -le 86; $r++) {
        for ($c = 27; $c -le 32; $c++) { if ("$($est.Cells($r,$c).Formula)" -like '=*') { $comF++ } }
    }
    if ($comF -ne $nOrig) { $erros += "AA:AF ficou com $comF formulas, origem tinha $nOrig" }

    $cabNovo = @()
    for ($c = 27; $c -le 32; $c++) { $cabNovo += "$($est.Cells(6, $c).Value2)" }
    if (($cabNovo -join '|') -ne ($cab -join '|')) { $erros += "cabecalho AA6:AF6 nao confere: '$($cabNovo -join '|')'" }
    if (($cabNovo -join '') -like '*Status ETP*') { $erros += "AA:AF ainda contem as colunas NOVAS -- a restauracao nao pegou" }

    # as colunas do layout novo nao podem ter sido tocadas
    if ("$($est.Cells(6,16).Value2)" -notlike '*ETP VB*') { $erros += "P6 deveria ser 'ETP VB %', esta '$($est.Cells(6,16).Value2)'" }
    if ("$($est.Cells(6,17).Value2)" -notlike '*Status ETP*') { $erros += "Q6 deveria ser 'Status ETP', esta '$($est.Cells(6,17).Value2)'" }

    if ($erros.Count -gt 0) {
        $erros | ForEach-Object { "  FALHA: $_" }
        throw "Restauracao recusada: $($erros.Count) conferencia(s) falharam. Nada foi salvo."
    }
    "conferencia: $comF formulas em AA:AF, cabecalho identico a origem, layout novo intacto em P:Q"

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
