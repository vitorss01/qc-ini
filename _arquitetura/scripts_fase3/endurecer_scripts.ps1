# endurecer_scripts.ps1 - aplica, de uma vez, os dois controles que faltavam em
# todos os scripts do pipeline que abrem o .xlsm.
#
# CONTEXTO. O pipeline reportava sucesso sobre etapas que nao aconteceram. Duas
# causas somadas:
#
#   1. Sem $ErrorActionPreference = 'Stop', erro de COM em PowerShell e
#      NAO-TERMINANTE: o Save falha, a mensagem vai para o stderr, o script
#      segue e imprime as proprias mensagens de sucesso.
#
#   2. $xl.Quit() NAO encerra o processo enquanto o PowerShell ainda segura
#      referencias COM ($wb, $ws, ...). O Excel fica vivo segurando o arquivo, e
#      a etapa SEGUINTE o abre em SOMENTE LEITURA -- silenciosamente, porque
#      DisplayAlerts = $false suprime o aviso. Dai o Save falha (causa 1) e o
#      trabalho evapora.
#
# O sintoma disso foi o motor corrigido nunca entrar no .xlsm, em duas maquinas,
# por semanas, com o build dizendo "Salvo:" a cada etapa.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\endurecer_scripts.ps1              (aplica)
#   .\endurecer_scripts.ps1 -Conferir    (so relata o que falta)

param([switch]$Conferir)

$ErrorActionPreference = 'Stop'
$enc = [System.Text.Encoding]::Default
$s = Split-Path -Parent $MyInvocation.MyCommand.Path

$alvos = @(
    'criar_eng_saida.ps1', 'criar_corridas.ps1', 'patch_forms_run.ps1',
    'redirecionar_calc.ps1', 'redirecionar_painel.ps1', 'redirecionar_estatistica.ps1',
    'rodar_motor.ps1', 'snapshot_formulas.ps1', 'varredura_adr019.ps1',
    'mapa_escritas.ps1', 'comparar_veredictos.ps1', 'snapshot_projeto.ps1'
)

$guarda = @'

# Somente leitura = outra instancia do Excel ainda segura o arquivo. Sem esta
# guarda o script grava no vazio e reporta sucesso.
if ($wb.ReadOnly) {
    $wb.Close($false); $xl.Quit()
    throw "Arquivo aberto em SOMENTE LEITURA (outra instancia do Excel o mantem travado): $Workbook"
}
'@

$relatorio = @()

foreach ($nome in $alvos) {
    $p = Join-Path $s $nome
    if (-not (Test-Path $p)) { $relatorio += "$nome : ausente"; continue }

    $txt = [System.IO.File]::ReadAllText($p, $enc)
    $mudou = $false
    $faltava = @()

    if ($txt -notmatch '\$ErrorActionPreference\s*=\s*[''"]Stop[''"]') {
        $faltava += 'ErrorActionPreference'
        if (-not $Conferir) {
            # logo apos o bloco param(...) ou, se nao houver, antes do primeiro
            # New-Object Excel
            if ($txt -match '(?ms)^(param\s*\([^\)]*\)\s*)') {
                $txt = $txt -replace '(?ms)^(param\s*\([^\)]*\)\s*)', "`$1`r`n`$ErrorActionPreference = 'Stop'`r`n"
            }
            else {
                $txt = "`$ErrorActionPreference = 'Stop'`r`n" + $txt
            }
            $mudou = $true
        }
    }

    if ($txt -notmatch '\$wb\.ReadOnly') {
        $faltava += 'guarda de somente-leitura'
        if (-not $Conferir -and $txt -match '\$wb\s*=\s*\$xl\.Workbooks\.Open\([^\)]*\)') {
            $txt = $txt -replace '(\$wb\s*=\s*\$xl\.Workbooks\.Open\([^\)]*\))', ("`$1`r`n" + $guarda)
            $mudou = $true
        }
    }

    if ($faltava.Count -eq 0) { $relatorio += "$nome : ok" }
    else { $relatorio += "$nome : faltava $($faltava -join ' + ')" }

    if ($mudou) { [System.IO.File]::WriteAllText($p, $txt, $enc) }
}

$relatorio | ForEach-Object { "  $_" }
if ($Conferir) { "(modo conferencia - nada foi alterado)" } else { "endurecimento aplicado" }
