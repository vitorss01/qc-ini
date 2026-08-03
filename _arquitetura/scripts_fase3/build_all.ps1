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

# A pasta de build fica FORA do OneDrive.
#
# Nao porque a sincronizacao tenha causado o problema -- essa hipotese foi
# testada e REFUTADA: fora do OneDrive o sintoma se repetiu identico. A causa
# real era o Excel orfao segurando o arquivo (ver Encerrar-Excel abaixo).
# Ainda assim, artefato binario nao pertence a pasta sincronizada: o Git ja
# versiona a fonte, e o .xlsm e produto.
#
# NAO usar %LOCALAPPDATA%: o Excel recusa abrir de AppData\Local e devolve
# "nao foi possivel encontrar" mesmo com o arquivo no lugar. %USERPROFILE%
# funciona, nao exige privilegio e fica fora da pasta sincronizada.
# O NOME PRECISA CONTER "build_hardening": rodar_motor.ps1 recusa executar
# fora de uma copia de build, justamente para o motor nunca rodar sobre a
# producao. A trava e boa -- o nome se adapta a ela, nao o contrario.
$bd = Join-Path $env:USERPROFILE 'QCINI_build_hardening1'
$prod = Join-Path $raiz 'QC_Hematologia.xlsm'
$alvo = Join-Path $bd 'QC_Hematologia.xlsm'

if (-not (Test-Path $prod)) { throw "Producao nao encontrada: $prod" }
New-Item -ItemType Directory -Force -Path $bd | Out-Null

# Encerra o Excel ENTRE as etapas.
#
# $xl.Quit() nao mata o processo enquanto o PowerShell ainda segura referencias
# COM. O Excel sobrevive segurando o .xlsm, e a etapa seguinte o abre em
# SOMENTE LEITURA sem avisar (DisplayAlerts = $false). O Save entao falha e a
# etapa inteira se perde reportando sucesso. Foi assim que o motor corrigido
# ficou fora do arquivo por semanas, em duas maquinas.

# NAO CANALIZAR AS ETAPAS PARA "Select-Object -First".
#
# "-First" INTERROMPE O PIPELINE A MONTANTE: assim que tem os N objetos que
# pediu, o PowerShell lanca StopUpstreamCommandsException e MATA o script
# produtor. As etapas morriam depois de imprimir suas primeiras linhas e ANTES
# do $wb.Save() -- gravando nada e parecendo bem-sucedidas.
#
# Isso explicava o padrao: criar_eng_saida usava "-Last 1" (que bufferiza tudo
# e nao interrompe) e persistia; criar_corridas usava "-Skip 1 -First 3" e a
# aba Corridas nunca aparecia no artefato.
#
# Aqui a saida e capturada em variavel e so depois recortada. O script sempre
# roda ate o fim.
function Mostrar {
    param([string[]]$Linhas, [int]$Primeiras = 0, [int]$Ultimas = 0, [int]$Pular = 0)
    if ($Linhas -eq $null) { return }
    $x = $Linhas
    if ($Pular -gt 0) { $x = $x | Select-Object -Skip $Pular }
    if ($Primeiras -gt 0) { $x = $x[0..([Math]::Min($Primeiras, $x.Count) - 1)] }
    elseif ($Ultimas -gt 0 -and $x.Count -gt $Ultimas) { $x = $x[($x.Count - $Ultimas)..($x.Count - 1)] }
    $x | ForEach-Object { $_ }
}

function Encerrar-Excel {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    $mortos = 0
    Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
        try { $_.Kill(); $mortos++ } catch { }
    }
    if ($mortos -gt 0) { Start-Sleep -Milliseconds 1200 }
}

# Excel orfao de execucao anterior mantem o arquivo travado e a copia falha em
# silencio -- o build seguiria sobre o artefato velho.
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }
Start-Sleep -Seconds 2

"== 0. copia limpa da producao"
Copy-Item $prod $alvo -Force

Encerrar-Excel
"== 1. gera os modulos VBA a partir dos patches"
Mostrar (& (Join-Path $s 'gerar_mEstatistica.ps1') `
    -Producao (Join-Path $arq 'snapshot_producao\Hematologia\vba\mEstatistica.bas') `
    -NovoSub (Join-Path $h 'AtualizarCalc.txt') `
    -NovoPainel (Join-Path $h 'AtualizarPainelEng.txt') `
    -NovoEstat (Join-Path $h 'AtualizarEstatisticaAba.txt') `
    -Saida (Join-Path $h 'mEstatistica.bas')) -Ultimas 2

Mostrar (& (Join-Path $s 'gerar_mDados.ps1') `
    -Producao (Join-Path $arq 'snapshot_producao\Hematologia\vba\mDados.bas') `
    -NovoRun (Join-Path $h 'mDados_RUN.txt') `
    -Saida (Join-Path $h 'mDados.bas')) -Ultimas 2

Mostrar (& (Join-Path $s 'gerar_mDados_audit.ps1') `
    -Entrada (Join-Path $h 'mDados.bas') `
    -Patch (Join-Path $h 'mDados_AUDIT.txt') `
    -Saida (Join-Path $h 'mDados.bas')) -Ultimas 2

Encerrar-Excel
"== 2. Eng_Saida (camada de saida do motor)"
Mostrar (& (Join-Path $s 'criar_eng_saida.ps1') -Workbook $alvo) -Ultimas 1

Encerrar-Excel
"== 3. Corridas (identidade do RUN) + migracao"
Mostrar (& (Join-Path $s 'criar_corridas.ps1') -Workbook $alvo) -Pular 1 -Primeiras 3

Encerrar-Excel
"== 3b. Audit_Log (trilha de auditoria encadeada por hash)"
Mostrar (& (Join-Path $s 'criar_audit_log.ps1') -Workbook $alvo) -Primeiras 1

Encerrar-Excel
"== 4. aplica o VBA"
& (Join-Path $s 'aplicar_vba.ps1') -Workbook $alvo -Modulos @(
    (Join-Path $h 'mEstatistica.bas'),
    (Join-Path $h 'mDados.bas'),
    (Join-Path $h 'mAuditoria.bas'),
    (Join-Path $h 'Planilha7.cls')
)

Encerrar-Excel
"== 5. migra os formularios para a nova API do RUN"
Mostrar (& (Join-Path $s 'patch_forms_run.ps1') -Workbook $alvo) -Ultimas 2

Encerrar-Excel
"== 6. redireciona as abas de interface para Eng_Saida"
Mostrar (& (Join-Path $s 'redirecionar_calc.ps1') -Workbook $alvo -OutCsv (Join-Path $h 'marco2_celulas_redirecionadas.csv')) -Primeiras 1
Encerrar-Excel
Mostrar (& (Join-Path $s 'redirecionar_painel.ps1') -Workbook $alvo -OutCsv (Join-Path $h 'marco3_celulas_redirecionadas.csv')) -Primeiras 1
Encerrar-Excel
Mostrar (& (Join-Path $s 'redirecionar_estatistica.ps1') -Workbook $alvo -OutCsv (Join-Path $h 'marco4_celulas_redirecionadas.csv')) -Primeiras 1

Encerrar-Excel
if (-not $PularMotor) {
    "== 7. executa o motor"
    Mostrar (& (Join-Path $s 'rodar_motor.ps1') -Workbook $alvo -Rotinas AtualizarCalc, AtualizarPainelEng, AtualizarEstatisticaAba) -Ultimas 4
}

""
"BUILD PRONTO: $alvo"
"Conferir com:"
"  .\snapshot_formulas.ps1 -Workbook '$alvo' -OutCsv `$env:TEMP\build.csv"
"  .\diff_formulas.ps1 -Referencia '$(Join-Path $arq 'snapshot_producao\Hematologia\formulas.csv')' -Candidato `$env:TEMP\build.csv"
"Aceite esperado: AUSENTE 0 - ALTERADA 4362 - VALOR 0 - EXTRA 9"
