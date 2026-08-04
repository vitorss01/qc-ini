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

# Criar o Excel COM RESILIENCIA.
#
# O build sobe e derruba o Excel cerca de dez vezes. Sob esse ritmo o servidor
# COM as vezes recusa a proxima instancia com 0x80080005
# (CO_E_SERVER_EXEC_FAILURE) -- estado transitorio, nao defeito do script.
# Falhar na primeira tentativa jogava fora um build inteiro de varios minutos.
function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try {
            $app = New-Object -ComObject Excel.Application
            return $app
        }
        catch {
            $ultimo = $_
            # Depois de um periodo sem Excel rodando, a PRIMEIRA ativacao COM
            # costuma falhar com 0x80080005 mesmo com a maquina sadia. Lancar o
            # excel.exe uma vez levanta o servidor e as ativacoes seguintes
            # funcionam. Verificado nesta maquina: com um processo de pe, o
            # New-Object passa na hora.
            if ($tentativa -eq 2) {
                try {
                    Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null
                    Start-Sleep -Seconds 5
                }
                catch { }
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
$xl.AutomationSecurity = 1      # msoAutomationSecurityLow: permite macro

$wb = $xl.Workbooks.Open($Workbook)

# AutoRecuperacao DESLIGADA nesta copia de trabalho.
#
# O build encerra o Excel a forca varias vezes. Cada encerramento deixa um
# arquivo de recuperacao pendente; acumulados, o Excel passa a tentar exibir o
# painel "Recuperacao de Documento" ao iniciar e MORRE antes de responder --
# ate o excel.exe puro para de abrir, e a automacao falha com 0x80080005
# (CO_E_SERVER_EXEC_FAILURE), que parece defeito de COM e nao e.
#
# O artefato e reproduzivel por comando: nao ha o que recuperar aqui.
try { $wb.EnableAutoRecover = $false } catch { }


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
