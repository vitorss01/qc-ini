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
    [switch]$PularMotor,
    # Produto alvo. Remove a amarracao a Hematologia em nome de arquivo,
    # snapshot e pasta de build.
    [string]$Produto = 'Hematologia'
)

$ErrorActionPreference = 'Stop'

$s = Split-Path -Parent $MyInvocation.MyCommand.Path
$arq = Split-Path -Parent $s
$raiz = Split-Path -Parent $arq
$h = Join-Path $arq 'src_hardening1'

# Niveis de controle por produto.
#
# NAO e so um numero de script: define a GEOMETRIA da pasta de trabalho --
# quantos blocos de coluna o Calc tem, quantas linhas de nivel o Painel usa,
# quantas linhas a Estatistica ocupa e quantos graficos existem. Parametrizar os
# scripts e condicao necessaria para propagar, nao suficiente: Bioquimica e
# Imunologia ainda NAO tem motor (falta mEstatistica, mWestgardKnowledge,
# Cfg_Status e Eventos_Westgard) e suas planilhas tem geometria propria.
$NIVEIS_POR_PRODUTO = @{
    'Hematologia' = 3
    'Bioquimica'  = 2
    'Imunologia'  = 2
}
if (-not $NIVEIS_POR_PRODUTO.ContainsKey($Produto)) {
    throw "Produto desconhecido: $Produto. Conhecidos: $($NIVEIS_POR_PRODUTO.Keys -join ', ')"
}
$NLV = $NIVEIS_POR_PRODUTO[$Produto]
"produto: $Produto   niveis: $NLV"

$snapProduto = Join-Path $arq "snapshot_producao\$Produto"
if (-not (Test-Path $snapProduto)) {
    throw "Snapshot de producao ausente para $Produto : $snapProduto. Gere com snapshot_projeto.ps1 antes de construir."
}

# O motor tem UMA fonte, sempre a Hematologia.
#
# Bioquimica e Imunologia nunca receberam a Fase 3 e nao possuem mEstatistica
# nem mWestgardKnowledge -- nao ha o que "atualizar" nelas. O modulo da
# Hematologia e a base para todos os produtos, e o que muda entre eles e o NLV,
# reescrito por gerar_mEstatistica.ps1. Ter uma fonte so e o que impede os
# produtos de divergirem em regra de Westgard, como Calc e motor divergiram.
$snapMotor = Join-Path $arq 'snapshot_producao\Hematologia'

# Modulos gerados ficam em pasta por produto: a Bioquimica tem NLV=2 e nao pode
# sobrescrever o modulo da Hematologia.
$hp = Join-Path $h $Produto
New-Item -ItemType Directory -Force -Path $hp | Out-Null

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
$bd = Join-Path $env:USERPROFILE "QCINI_build_hardening1_$Produto"
$prod = Join-Path $raiz "QC_$Produto.xlsm"
$alvo = Join-Path $bd "QC_$Produto.xlsm"

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
    $erros = @($linhas | Where-Object { $_ -match 'Exce(p|�)|Error|throw|CategoryInfo' })

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
    -Producao (Join-Path $snapMotor 'vba\mEstatistica.bas') `
    -NovoSub (Join-Path $h 'AtualizarCalc.txt') `
    -NovoPainel (Join-Path $h 'AtualizarPainelEng.txt') `
    -NovoEstat (Join-Path $h 'AtualizarEstatisticaAba.txt') `
    -NLV $NLV `
    -Saida (Join-Path $hp 'mEstatistica.bas')) -Ultimas 3

Mostrar (& (Join-Path $s 'gerar_mDados.ps1') `
    -Producao (Join-Path $snapProduto 'vba\mDados.bas') `
    -NovoRun (Join-Path $h 'mDados_RUN.txt') `
    -Saida (Join-Path $hp 'mDados.bas')) -Ultimas 2

# Limite atingido passa a falhar ALTO, nao em silencio (inspecao 04/08/2026).
# mRegistros e copiado para a pasta do produto antes de ser corrigido: o
# original em src_hardening1 e compartilhado e nao pode ser mutado.
Copy-Item (Join-Path $h 'mRegistros.bas') (Join-Path $hp 'mRegistros.bas') -Force
Copy-Item (Join-Path $h 'mAuditoria.bas') (Join-Path $hp 'mAuditoria.bas') -Force
Mostrar (& (Join-Path $s 'corrigir_silenciamento.ps1') -Arquivo (Join-Path $hp 'mEstatistica.bas') -Alvo 'mEstatistica')
Mostrar (& (Join-Path $s 'corrigir_silenciamento.ps1') -Arquivo (Join-Path $hp 'mRegistros.bas') -Alvo 'mRegistros')
Mostrar (& (Join-Path $s 'buffer_dinamico.ps1') -Arquivo (Join-Path $hp 'mEstatistica.bas'))
Mostrar (& (Join-Path $s 'corrigir_rotulo_audit.ps1') -Arquivo (Join-Path $hp 'mAuditoria.bas') -Novo (Join-Path $h 'Rotulo.txt'))

Mostrar (& (Join-Path $s 'gerar_mDados_audit.ps1') `
    -Entrada (Join-Path $hp 'mDados.bas') `
    -Patch (Join-Path $h 'mDados_AUDIT.txt') `
    -Saida (Join-Path $hp 'mDados.bas')) -Ultimas 2

Encerrar-Excel
"== 2. Eng_Saida (camada de saida do motor)"
Etapa 'criar_eng_saida.ps1' -Argumentos @('-Workbook', $alvo, '-NLV', $NLV) -Ultimas 1

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
"== 3d. abas exigidas pelo motor (Cfg_Status, Eventos_Westgard)"
Etapa 'criar_abas_motor.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3

Encerrar-Excel
"== 4. aplica o VBA"
# aplicar_vba recebe um ARRAY de modulos: continua no processo atual,
# seguido de Encerrar-Excel como rede de seguranca.
& (Join-Path $s 'aplicar_vba.ps1') -Workbook $alvo -Modulos @(
    (Join-Path $hp 'mEstatistica.bas'),
    (Join-Path $hp 'mDados.bas'),
    (Join-Path $snapMotor 'vba\mWestgardKnowledge.bas'),
    # mUI do motor tambem: a versao pre-Fase-3 da Bioquimica define um
    # AtualizarEstatistica proprio (stub que so recalcula a aba), e nome publico
    # duplicado entre modulos IMPEDE O PROJETO DE COMPILAR. O sintoma e
    # "macro nao disponivel" no Application.Run -- a mesma familia do aS.
    # AtualizarEixos e HookCharts sao identicos nos dois e ja iteram sobre
    # ChartObjects.Count, entao a troca e segura e no-op na Hematologia.
    (Join-Path $snapMotor 'vba\mUI.bas'),
    (Join-Path $hp 'mAuditoria.bas'),
    (Join-Path $h 'mConfig.bas'),
    (Join-Path $h 'mLogDB.bas'),
    (Join-Path $hp 'mRegistros.bas'),
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
"== 5c. Option Explicit em todo componente com codigo"
Etapa 'forcar_option_explicit.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 4

Encerrar-Excel
"== 6. redireciona as abas de interface para Eng_Saida"
Etapa 'redirecionar_calc.ps1' -Argumentos @('-Workbook', $alvo, '-NLV', $NLV, '-OutCsv', (Join-Path $h 'marco2_celulas_redirecionadas.csv')) -Primeiras 1
Encerrar-Excel
Etapa 'redirecionar_painel.ps1' -Argumentos @('-Workbook', $alvo, '-NLV', $NLV, '-OutCsv', (Join-Path $h 'marco3_celulas_redirecionadas.csv')) -Primeiras 1
Encerrar-Excel
Etapa 'redirecionar_estatistica.ps1' -Argumentos @('-Workbook', $alvo, '-NLV', $NLV, '-OutCsv', (Join-Path $h 'marco4_celulas_redirecionadas.csv')) -Primeiras 1

Encerrar-Excel
"== 6b. Rep 1/2/3 viram uma coluna unica de nao conformidade"
Etapa 'consolidar_registros_nc.ps1' -Argumentos @('-Workbook', $alvo, '-NLV', $NLV) -Ultimas 6

Encerrar-Excel
if (-not $PularMotor) {
    "== 7. executa o motor"
        # Com powershell.exe -File, parametro de ARRAY vai como UM argumento
    # separado por virgulas. Elementos soltos viram argumentos posicionais e
    # a ligacao falha com 'nao e possivel processar a transformacao'.
    Etapa 'rodar_motor.ps1' -Argumentos @('-Workbook', $alvo, '-Rotinas', 'AtualizarCalc,AtualizarPainelEng,AtualizarEstatisticaAba') -Ultimas 4
}

Encerrar-Excel
"== 8. trava a estrutura (abas so pelo Modo Desenvolvedor)"
Etapa 'travar_estrutura.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3

""
"BUILD PRONTO: $alvo"
"Conferir com:"
"  .\snapshot_formulas.ps1 -Workbook '$alvo' -OutCsv `$env:TEMP\build.csv"
"  .\diff_formulas.ps1 -Referencia '$(Join-Path $snapProduto 'formulas.csv')' -Candidato `$env:TEMP\build.csv"
"Aceite esperado: AUSENTE 1080 - ALTERADA 4362 - VALOR 0 - EXTRA 9"
"  as 1.080 AUSENTE sao regRep2 e regRep3, removidas de proposito na etapa 6b"
