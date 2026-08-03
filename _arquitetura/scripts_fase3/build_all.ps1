# build_all.ps1 - reconstroi o artefato do hardening a partir da producao
#
# Um comando so. Depois de um "git pull" noutra maquina, basta rodar isto para
# obter QC_Hematologia.xlsm com tudo aplicado: Eng_Saida, Corridas, o motor
# corrigido, os redirecionamentos e os formularios migrados.
#
# O .xlsm de build NAO e versionado -- e artefato, nao fonte. A fonte e
# snapshot_producao/<produto>/vba/ mais os patches de src_hardening1/, e este
# script e a receita que junta as duas coisas.
#
# A PRODUCAO NAO E TOCADA. Tudo acontece sobre uma copia em build_hardening1/.
#
# ATENCAO: exige "Confiar no acesso ao modelo de objeto do projeto do VBA"
# (Central de Confiabilidade > Configuracoes de Macro). Sem isso a aplicacao
# do VBA falha.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\build_all.ps1
#   .\build_all.ps1 -PularMotor      (nao executa o motor ao final)

param(
    [switch]$PularMotor
)

$ErrorActionPreference = 'Stop'

$s = Split-Path -Parent $MyInvocation.MyCommand.Path
$arq = Split-Path -Parent $s
$raiz = Split-Path -Parent $arq
$h = Join-Path $arq 'src_hardening1'
$bd = Join-Path $arq 'build_hardening1'
$prod = Join-Path $raiz 'QC_Hematologia.xlsm'
$alvo = Join-Path $bd 'QC_Hematologia.xlsm'

if (-not (Test-Path $prod)) { throw "Producao nao encontrada: $prod" }
New-Item -ItemType Directory -Force -Path $bd | Out-Null

# Excel orfao de execucao anterior mantem o arquivo travado e a copia falha em
# silencio -- o build seguiria sobre o artefato velho.
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }
Start-Sleep -Seconds 2

"== 0. copia limpa da producao"
Copy-Item $prod $alvo -Force

"== 1. gera os modulos VBA a partir dos patches"
& (Join-Path $s 'gerar_mEstatistica.ps1') `
    -Producao (Join-Path $arq 'snapshot_producao\Hematologia\vba\mEstatistica.bas') `
    -NovoSub (Join-Path $h 'AtualizarCalc.txt') `
    -NovoPainel (Join-Path $h 'AtualizarPainelEng.txt') `
    -NovoEstat (Join-Path $h 'AtualizarEstatisticaAba.txt') `
    -Saida (Join-Path $h 'mEstatistica.bas') | Select-Object -Last 2

& (Join-Path $s 'gerar_mDados.ps1') `
    -Producao (Join-Path $arq 'snapshot_producao\Hematologia\vba\mDados.bas') `
    -NovoRun (Join-Path $h 'mDados_RUN.txt') `
    -Saida (Join-Path $h 'mDados.bas') | Select-Object -Last 2

"== 2. Eng_Saida (camada de saida do motor)"
& (Join-Path $s 'criar_eng_saida.ps1') -Workbook $alvo | Select-Object -Last 1

"== 3. Corridas (identidade do RUN) + migracao"
& (Join-Path $s 'criar_corridas.ps1') -Workbook $alvo | Select-Object -Skip 1 -First 3

"== 4. aplica o VBA"
& (Join-Path $s 'aplicar_vba.ps1') -Workbook $alvo -Modulos @(
    (Join-Path $h 'mEstatistica.bas'),
    (Join-Path $h 'mDados.bas'),
    (Join-Path $h 'Planilha7.cls')
) | Select-Object -Last 1

"== 5. migra os formularios para a nova API do RUN"
& (Join-Path $s 'patch_forms_run.ps1') -Workbook $alvo | Select-Object -Last 2

"== 6. redireciona as abas de interface para Eng_Saida"
& (Join-Path $s 'redirecionar_calc.ps1') -Workbook $alvo -OutCsv (Join-Path $h 'marco2_celulas_redirecionadas.csv') | Select-Object -First 1
& (Join-Path $s 'redirecionar_painel.ps1') -Workbook $alvo -OutCsv (Join-Path $h 'marco3_celulas_redirecionadas.csv') | Select-Object -First 1
& (Join-Path $s 'redirecionar_estatistica.ps1') -Workbook $alvo -OutCsv (Join-Path $h 'marco4_celulas_redirecionadas.csv') | Select-Object -First 1

if (-not $PularMotor) {
    "== 7. executa o motor"
    & (Join-Path $s 'rodar_motor.ps1') -Workbook $alvo -Rotinas AtualizarCalc, AtualizarPainelEng, AtualizarEstatisticaAba | Select-Object -Last 4
}

""
"BUILD PRONTO: $alvo"
"Conferir com:"
"  .\snapshot_formulas.ps1 -Workbook '$alvo' -OutCsv `$env:TEMP\build.csv"
"  .\diff_formulas.ps1 -Referencia '$(Join-Path $arq 'snapshot_producao\Hematologia\formulas.csv')' -Candidato `$env:TEMP\build.csv"
"Aceite esperado: AUSENTE 0 - ALTERADA 4362 - VALOR 0 - EXTRA 9"
