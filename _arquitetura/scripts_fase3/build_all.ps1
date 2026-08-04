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

# Cada etapa roda em SEU PROPRIO PROCESSO do PowerShell.
#
# POR QUE. Invocadas com "&", as etapas compartilham o processo -- e o runspace
# mantem vivas as referencias COM (RCW) mesmo depois do $xl.Quit(). O Excel
# sobrevivia segurando o .xlsm, e a etapa seguinte o abria em SOMENTE LEITURA
# sem avisar. A solucao de antes era matar o Excel entre as etapas, e ela criou
# um problema pior: dez Kill() por build desestabilizam a ativacao DCOM, e o
# COM passa a recusar novas instancias com 0x80080005 -- a ponto de nem o
# excel.exe puro abrir.
#
# Com um processo por etapa, o encerramento do processo libera TODAS as
# referencias e o Excel sai sozinho, do jeito certo. Sem Kill, sem lock, sem
# DCOM instavel. Custa ~1s de inicializacao por etapa e paga com folga.
function Etapa {
    param(
        [Parameter(Mandatory = $true)][string]$Script,
        [string[]]$Argumentos = @(),
        [int]$Primeiras = 0,
        [int]$Ultimas = 0,
        [int]$Pular = 0
    )
    $caminho = Join-Path $s $Script
    $saida = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $caminho @Argumentos 2>&1
    $codigo = $LASTEXITCODE

    # Encerrar DEPOIS tambem. Um script que lanca excecao fecha a pasta mas
    # deixa o PROCESSO vivo -- o Quit nao encerra enquanto o runspace segura
    # referencias COM. A etapa seguinte encontrava o arquivo travado e
    # abortava com "somente leitura". Limpar na saida torna cada etapa
    # independente do estado deixado pela anterior.
    Encerrar-Excel

    $linhas = @($saida | ForEach-Object { $_.ToString() })
    $erros = @($linhas | Where-Object { $_ -match 'Exce(p|ç)|Error|throw|CategoryInfo' })

    if ($codigo -ne 0 -or $erros.Count -gt 0) {
        $linhas | ForEach-Object { "     $_" }
        throw "Etapa $Script falhou (codigo $codigo)"
    }
    Mostrar $linhas -Primeiras $Primeiras -Ultimas $Ultimas -Pular $Pular
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
Etapa 'criar_eng_saida.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 1

Encerrar-Excel
"== 3. Corridas (identidade do RUN) + migracao"
Etapa 'criar_corridas.ps1' -Argumentos @('-Workbook', $alvo) -Pular 1 -Primeiras 3

Encerrar-Excel
"== 3b. Audit_Log (trilha de auditoria encadeada por hash)"
Etapa 'criar_audit_log.ps1' -Argumentos @('-Workbook', $alvo) -Primeiras 1

Encerrar-Excel
"== 3c. tabelas de log dentro do DB_Resultados"
Etapa 'criar_logs_db.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3

Encerrar-Excel
"== 4. aplica o VBA"
# aplicar_vba recebe um ARRAY de modulos: continua no processo atual,
# seguido de Encerrar-Excel como rede de seguranca.
& (Join-Path $s 'aplicar_vba.ps1') -Workbook $alvo -Modulos @(
    (Join-Path $h 'mEstatistica.bas'),
    (Join-Path $h 'mDados.bas'),
    (Join-Path $h 'mAuditoria.bas'),
    (Join-Path $h 'mConfig.bas'),
    (Join-Path $h 'mLogDB.bas'),
    (Join-Path $h 'mRegistros.bas'),
    (Join-Path $h 'Planilha7.cls')
)

Encerrar-Excel
"== 4b. vigia da tabela de elegibilidade (item 2.5)"
Etapa 'instalar_cfg_watch.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3

Encerrar-Excel
"== 5. migra os formularios para a nova API do RUN"
Etapa 'patch_forms_run.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 2

Encerrar-Excel
"== 5b. formularios da Sprint NC"
Etapa 'criar_forms_nc.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3

Encerrar-Excel
Etapa 'criar_botoes_nc.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3

Encerrar-Excel
"== 6. redireciona as abas de interface para Eng_Saida"
Etapa 'redirecionar_calc.ps1' -Argumentos @('-Workbook', $alvo, '-OutCsv', (Join-Path $h 'marco2_celulas_redirecionadas.csv')) -Primeiras 1
Encerrar-Excel
Etapa 'redirecionar_painel.ps1' -Argumentos @('-Workbook', $alvo, '-OutCsv', (Join-Path $h 'marco3_celulas_redirecionadas.csv')) -Primeiras 1
Encerrar-Excel
Etapa 'redirecionar_estatistica.ps1' -Argumentos @('-Workbook', $alvo, '-OutCsv', (Join-Path $h 'marco4_celulas_redirecionadas.csv')) -Primeiras 1

Encerrar-Excel
if (-not $PularMotor) {
    "== 7. executa o motor"
        # Com powershell.exe -File, parametro de ARRAY vai como UM argumento
    # separado por virgulas. Elementos soltos viram argumentos posicionais e
    # a ligacao falha com 'nao e possivel processar a transformacao'.
    Etapa 'rodar_motor.ps1' -Argumentos @('-Workbook', $alvo, '-Rotinas', 'AtualizarCalc,AtualizarPainelEng,AtualizarEstatisticaAba') -Ultimas 4
}

""
"BUILD PRONTO: $alvo"
"Conferir com:"
"  .\snapshot_formulas.ps1 -Workbook '$alvo' -OutCsv `$env:TEMP\build.csv"
"  .\diff_formulas.ps1 -Referencia '$(Join-Path $arq 'snapshot_producao\Hematologia\formulas.csv')' -Candidato `$env:TEMP\build.csv"
"Aceite esperado: AUSENTE 0 - ALTERADA 4362 - VALOR 0 - EXTRA 9"
