# forcar_option_explicit.ps1 - Option Explicit em todo componente com codigo
#
# Sem Option Explicit, uma variavel digitada errado nao e erro: o VBA cria uma
# nova, vazia. "alvoMedia" virando "alvoMedai" devolve zero e a estatistica sai
# errada SEM avisar. E a mesma familia dos outros defeitos deste projeto --
# falha que nao se anuncia.
#
# Aplica DENTRO do artefato, pelo CodeModule, e so onde ha codigo de verdade:
# modulo de documento vazio nao ganha nada e a mudanca seria ruido no diff.
#
# IDEMPOTENTE: componente que ja declara nao e tocado.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\forcar_option_explicit.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'

function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $ultimo = $_
            if ($tentativa -eq 2) {
                try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep -Seconds 5 } catch { }
            }
            Start-Sleep -Seconds ($tentativa * 2)
        }
    }
    throw "Excel COM nao subiu apos 6 tentativas: $($ultimo.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    throw "Arquivo aberto em SOMENTE LEITURA: $Workbook"
}

$aplicados = @()
$jaTinham = 0
$vazios = 0

foreach ($comp in $wb.VBProject.VBComponents) {
    $cm = $comp.CodeModule
    if ($cm.CountOfLines -lt 1) { $vazios++; continue }

    $codigo = $cm.Lines(1, $cm.CountOfLines)

    # "com codigo de verdade" = tem ao menos uma rotina. Modulo de documento so
    # com linhas em branco ou comentario nao conta.
    if ($codigo -notmatch '(?m)^\s*(Public |Private |Friend )?(Sub|Function|Property)\s+\w+') {
        $vazios++
        continue
    }
    if ($codigo -match '(?m)^\s*Option Explicit') { $jaTinham++; continue }

    # Vai na linha 1: Option so vale antes de qualquer declaracao ou rotina.
    $cm.InsertLines(1, 'Option Explicit')
    $aplicados += $comp.Name
}

"componentes com Option Explicit ja declarado : $jaTinham"
"componentes sem codigo (ignorados)           : $vazios"
"Option Explicit ACRESCENTADO em              : $(if ($aplicados.Count) { $aplicados -join ', ' } else { 'nenhum' })"

# Conferencia: relE o projeto e exige que nao sobre componente com rotina e sem
# a declaracao. Sem isto, um InsertLines que falhasse passaria como sucesso.
$faltando = @()
foreach ($comp in $wb.VBProject.VBComponents) {
    $cm = $comp.CodeModule
    if ($cm.CountOfLines -lt 1) { continue }
    $codigo = $cm.Lines(1, $cm.CountOfLines)
    if ($codigo -notmatch '(?m)^\s*(Public |Private |Friend )?(Sub|Function|Property)\s+\w+') { continue }
    if ($codigo -notmatch '(?m)^\s*Option Explicit') { $faltando += $comp.Name }
}
if ($faltando.Count -gt 0) {
    throw "Ainda sem Option Explicit apos a aplicacao: $($faltando -join ', ')"
}
"conferencia: ok - nenhum componente com rotina ficou sem Option Explicit"

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"Salvo: $Workbook"
