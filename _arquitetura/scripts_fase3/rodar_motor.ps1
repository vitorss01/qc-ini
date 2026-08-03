# rodar_motor.ps1 - executa rotinas do motor dentro do .xlsm de build
#
# Diferente dos demais scripts desta pasta, este ABRE COM MACROS HABILITADAS
# (AutomationSecurity = 1). Eventos ficam desligados para que nenhum
# Workbook_Open ou Worksheet_Change interfira no que estamos medindo.
#
# Usar somente em copia de build. Nunca apontar para a producao.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\rodar_motor.ps1 -Workbook <build.xlsm> -Rotinas AtualizarCalc[,Outra]
#   .\rodar_motor.ps1 -Workbook <build.xlsm> -Rotinas AtualizarCalc -Repeticoes 10

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [Parameter(Mandatory = $true)][string[]]$Rotinas,
    [int]$Repeticoes = 1,
    [string]$Analito
)

if ($Workbook -notmatch 'build_hardening') {
    throw "Recusado: este script so roda em copia de build (caminho deve conter build_hardening). Recebido: $Workbook"
}

$ErrorActionPreference = 'Stop'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1      # msoAutomationSecurityLow: permite macro

$wb = $xl.Workbooks.Open($Workbook)

# Somente leitura significa que OUTRA instancia do Excel ainda segura o arquivo.
# Sem esta guarda o script grava no vazio e reporta sucesso: DisplayAlerts=$false
# suprime o aviso do Excel, e o Save falha em silencio.
if ($wb.ReadOnly) {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    throw "Arquivo aberto em SOMENTE LEITURA (outra instancia do Excel o mantem travado): $Workbook"
}


if ($Analito) {
    $wb.Names('selAnalito').RefersToRange.Value2 = $Analito
    "selAnalito := $Analito"
}
"selAnalito atual: " + $wb.Names('selAnalito').RefersToRange.Value2

for ($r = 1; $r -le $Repeticoes; $r++) {
    foreach ($rot in $Rotinas) {
        $t0 = [Diagnostics.Stopwatch]::StartNew()
        $xl.Run("$($wb.Name)!$rot")
        $t0.Stop()
        "  [{0}/{1}] {2,-26} {3,7:N3} s" -f $r, $Repeticoes, $rot, $t0.Elapsed.TotalSeconds
    }
}

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"motor executado em: $Workbook"
