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
    # Modo de release/validacao: preserva integralmente os dados da fonte.
    # O comportamento historico (com fixture) continua sendo o padrao para a
    # suite de desenvolvimento existente.
    [switch]$SemFixtures,
    # Fonte e pasta de build opcionais permitem construir copias controladas
    # sem alterar nem depender do binario canonico da raiz.
    [string]$Fonte = '',
    [string]$PastaBuild = '',
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

# Analito de referencia do Painel, POR PRODUTO.
#
# Estava fixo em 'Glicose' na chamada da etapa 7d, o que derrubava o build
# INTEIRO da Hematologia -- que nao tem Glicose -- com "analito de referencia
# 'Glicose' nao esta cadastrado na aba Analitos". O passo esta certo e a
# amarracao e que nao estava: e um dado do produto, igual ao NLV, e pertence
# a este mapa e nao ao meio do script.
#
# O analito escolhido precisa TER RESULTADOS: a linha de base de formulas
# guarda valores calculados, e um analito sem dado produz Calc divergente e
# Eng_Saida vazio -- o ruido que a propria etapa 7d existe para eliminar.
# Verificado no banco: Glicose e WBC tem resultado nos respectivos produtos.
$REFERENCIA_POR_PRODUTO = @{
    'Hematologia' = 'WBC'
    'Bioquimica'  = 'Glicose'
    'Imunologia'  = ''
}
$ANALITO_REF = $REFERENCIA_POR_PRODUTO[$Produto]
"produto: $Produto   niveis: $NLV   analito de referencia: $(if ($ANALITO_REF) { $ANALITO_REF } else { '(nenhum)' })"

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
$bd = if ($PastaBuild) { $PastaBuild } else { Join-Path $env:USERPROFILE "QCINI_build_hardening1_$Produto" }
$prod = if ($Fonte) { $Fonte } else { Join-Path $raiz "QC_$Produto.xlsm" }
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
# O arquivo de producao e onde o usuario trabalha: se ele logar (LockApp
# protege todas as abas) e salvar, o arquivo fica gravado travado. O build nao
# pode depender de como a ultima sessao terminou -- aqui o artefato volta a um
# estado conhecido.
Etapa 'normalizar_protecao.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3

Encerrar-Excel
"== 0b. identidade persistente do contrato BI"
Etapa 'preparar_contrato_bi.ps1' -Argumentos @('-Workbook', $alvo, '-Produto', $Produto, '-Versao', '2.0.0') -Ultimas 3

# A FIXTURE PERTENCE AO BUILD, NAO A PRODUCAO.
#
# A suite sempre exercitou o sistema sobre os 1.000 resultados ficticios que
# moravam no arquivo de producao. Quando a producao foi limpa para o uso real,
# o artefato nasceu vazio e a suite travou lendo uma linha que nao existia.
# Dado de teste e responsabilidade de quem testa.
if ($SemFixtures) {
    "  fixture de teste: PULADA (-SemFixtures); dados da fonte preservados"
}
else {
    Etapa 'semear_dados_teste.ps1' -Argumentos @('-Workbook', $alvo, '-NLV', $NLV) -Ultimas 4
}
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
Copy-Item (Join-Path $h 'mImportar.bas') (Join-Path $hp 'mImportar.bas') -Force
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
    (Join-Path $hp 'mImportar.bas'),
    # mBanco (ADR-025). Sem ele o artefato nao compila: mDados chama
    # ExigirCapacidade e AtualizarFlagsBanco, que vivem aqui. Ficou de fora na
    # primeira entrega do ADR-025 e o resultado foi um artefato sem a capacidade
    # de 60 meses, enquanto a producao -- remendada a mao pelo instalador -- a
    # tinha. As provas passavam porque olhavam para o arquivo errado.
    (Join-Path $arq 'src_producao\mBanco.bas'),
    # mEQA (ADR-034): consolida EQA.CAP_Dados e EQA.Controllab_Dados na
    # EQA_Base. Vem ANTES do mCEQ, que le a base que ele monta.
    (Join-Path $arq 'src_producao\mEQA.bas'),
    # mCEQ (ADR-030/032/034): bias, SDI e status do EP. NAO estava na lista.
    # O artefato sairia sem ele e as 403 celulas da Estatistica e do Painel
    # que chamam BiasEQ/SDIeq/StatusSDIeq virariam #NOME?. E a mesma falha
    # que o mBanco pagou no ADR-025: producao remendada a mao passando nas
    # provas enquanto o build entregava outra coisa.
    (Join-Path $arq 'src_producao\mCEQ.bas'),
    # mQualidade (ADR-033): as escadas de Sigma e de orcamento de erro. Vem
    # ANTES do mBI, que as chama. Sem ele o artefato nao compila.
    (Join-Path $arq 'src_producao\mQualidade.bas'),
    # mPlanoQC (ADR-035): DPM, rendimento e o plano de CQ por Sigma.
    # Vem ANTES do mBI, que o chama.
    (Join-Path $arq 'src_producao\mPlanoQC.bas'),
    # mBI (ADR-026): camada de dados para o Power BI.
    (Join-Path $arq 'src_producao\mBI.bas'),
    (Join-Path $h 'Planilha7.cls'),
    # Salvamento controlado: recalcula motor, atualiza BI e bloqueia o Save se
    # a reconciliacao falhar.
    (Join-Path $h 'EstaPastaDeTrabalho.cls')
)

Encerrar-Excel
"== 4z. ADR-046: guarda de escrita em aba protegida (mSeguranca)"
# TEM de vir aqui: os modulos que acabaram de ser importados chamam
# LiberarEscrita/RestaurarProtecao, e mSeguranca e fonte de PRODUCAO -- o build
# nao a regera. Se a guarda so fosse instalada mais tarde (travar_estrutura.ps1
# roda depois do motor), a primeira execucao acharia "Sub ou Function nao
# definida" e o artefato sairia quebrado.
Etapa 'instalar_guarda_protecao.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 4

Encerrar-Excel
"== 4a. ADR-025: BA:BC viram VALOR e os nomes r* acompanham o dado"
# Depois do aplicar_vba (precisa do mBanco ja importado) e antes de qualquer
# etapa que leia o banco pelos nomes.
Etapa 'aplicar_adr025.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 6

Encerrar-Excel
"== 4a2. ADR-025: vigia de Status (BA:BC voltam a ser auto-corretivas)"
Mostrar (& python (Join-Path $s 'instalar_watch_status.py') $alvo) -Ultimas 3

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
"== 7a. ADR-045: mEstatPeriodo entra pelo build, e nao a mao"
# POR QUE ESTA ETAPA EXISTE.
#
# Ate aqui o mEstatPeriodo era o unico modulo de src_producao fora do build. O
# build_all parte da producao e SO adiciona modulos, entao a copia que ja vivia
# dentro do arquivo sobrevivia intacta e ninguem via falta -- nao havia #NOME?
# para denunciar. O resultado foi silencioso: a correcao do teto do LotesStore
# (ADR-045) ficou em src_producao por uma serie inteira de ADRs sem nunca
# chegar ao arquivo, e a producao seguiu varrendo "To 200", devolvendo alvo
# vazio do 5o lote em diante.
#
# E a mesma familia do mBanco (ADR-025) e do mCEQ (ADR-034): fonte a frente do
# artefato, com as provas olhando para o arquivo errado.
#
# DEPOIS do motor de proposito: o portao do instalador le n, media, DP e CV da
# Estatistica antes de importar e exige que o modulo novo os reproduza. Com a
# aba ainda vazia nao havia o que conferir.
#
# So a Bioquimica: e o unico produto cujas celulas chamam EstatPeriodo (323
# delas). Instalar na Hematologia poria um modulo que ninguem usa, e o portao
# nao teria linha para conferir.
if ($Produto -eq 'Bioquimica') {
    # ADR-049, ANTES do modulo: DB_Carimbo precisa existir para o Carimbo do
    # mEstatPeriodo poder le-lo, e para as formulas poderem cita-lo.
    Etapa 'instalar_carimbo_banco.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3

    Encerrar-Excel
    Etapa 'instalar_estat_periodo.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 5

    # ADR-049, DEPOIS do modulo: as 320 formulas passam a citar DB_Carimbo, e e
    # isso que poe a Estatistica no grafo de calculo do Excel. Sem este passo,
    # so um rebuild completo atualizaria -- e o fluxo real (mDados) chama
    # Application.Calculate, que nao faz rebuild.
    Encerrar-Excel
    Mostrar (& python (Join-Path $s 'ligar_carimbo_estatistica.py') $alvo) -Ultimas 3
} else {
    "  $Produto nao usa EstatPeriodo em celula: etapa nao se aplica"
}

Encerrar-Excel
"== 7b. importacao por aba (substitui o frmMassa)"
# O CSV DE ANALITOS E POR PRODUTO, E PRECISA SER PASSADO EXPLICITAMENTE.
#
# criar_aba_importar.ps1 tem 'analitos_bioquimica.csv' como valor padrao do
# parametro -Csv. Como a chamada nao passava nada, TODO produto recebia a aba
# Importar com os 31 analitos da Bioquimica -- inclusive a Hematologia, que tem
# 28 e outros nomes.
#
# O sintoma era mudo, e por isso perigoso: a aba montava, ficava bonita e
# passava em TODAS as provas estruturais (1.7 a 1.12), porque elas conferem a
# forma da aba contra ela mesma. So a prova de ponta a ponta pegava -- 3.16
# devolvia "ERRO|6", com as seis colunas rejeitadas por nao existirem no
# produto. Uma aba de importacao que nao importa nada.
$csvAnalitos = Join-Path $arq "src_producao\analitos_$($Produto.ToLowerInvariant()).csv"
if (-not (Test-Path $csvAnalitos)) {
    throw "CSV de analitos ausente para $Produto : $csvAnalitos. Sem ele a aba Importar sairia com os analitos de outro setor."
}
Etapa 'criar_aba_importar.ps1' -Argumentos @('-Workbook', $alvo, '-Csv', $csvAnalitos) -Ultimas 3

# O Calc e a area de UM analito por vez. Sem fixar qual, o valor de ~660
# celulas -- e o Eng_Saida inteiro -- depende de qual analito o spinner deixou
# selecionado, e a conferencia de formulas vira ruido. Fixar por NOME, nao por
# posicao: a posicao e justamente o que muda quando a lista e reordenada.
# O analito vem de $REFERENCIA_POR_PRODUTO. Produto sem referencia definida
# pula a etapa em vez de derrubar o build com o analito de outro setor.
if ($ANALITO_REF) {
    Etapa 'fixar_analito_referencia.ps1' -Argumentos @('-Workbook', $alvo, '-Analito', $ANALITO_REF) -Ultimas 1
    if (-not $PularMotor) {
        # Alterar o seletor invalida Eng_Saida. Reexecutar o motor garante que
        # Calc, Painel, Estatistica e a reconciliacao BI usem o mesmo analito.
        Etapa 'rodar_motor.ps1' -Argumentos @('-Workbook', $alvo, '-Rotinas', 'AtualizarCalc,AtualizarPainelEng,AtualizarEstatisticaAba') -Ultimas 4
    }
}
else {
    "  (sem analito de referencia definido para $Produto - etapa pulada)"
}

Encerrar-Excel
# A reconciliacao precisa ocorrer DEPOIS do redirecionamento e da execucao do
# motor. Antes disso, Calc ainda contem a implementacao antiga por formulas e
# nao e uma referencia valida para a saida de mEstatistica.
"== 7b2. ADR-026: camada de dados BI_Data + reconciliacao com o motor"
Etapa 'aplicar_bi_data.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 6

Encerrar-Excel
# Depois do motor: AtualizarPainelEng nao toca em titulo de eixo, mas a ordem
# aqui garante que o rotulo final e o ultimo estado gravado. Antes da trava,
# porque mexe em celula de aba protegida.
"== 7c. o identificador da corrida chama-se RUN (itens 6.2 e 6.3)"
Etapa 'padronizar_run.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 4

Encerrar-Excel
"== 7d. P0 de identidade/rotulos/data + padrao visual P3"
Etapa 'aplicar_release_visual.ps1' -Argumentos @('-Workbook', $alvo, '-Produto', $Produto, '-NLV', $NLV) -Ultimas 2

Encerrar-Excel

# Hematologia: estrutura formal de especificacoes sem corte prematuro.
#
# A referencia e tirada DEPOIS de todas as etapas comuns. Assim a prova mede
# apenas a introducao de Cfg/DB/Eng + mEspecificacoes. Analitos!Q:R continua
# operacional e nenhum consumidor e redirecionado para uma saida vazia.
$refEspec = $null
if ($Produto -eq 'Hematologia') {
    "== 7e. especificacoes formais da Hematologia com fallback Q:R"
    $refEspec = Join-Path $bd 'QC_Hematologia_ANTES_ESPEC_EM_DESENVOLVIMENTO.xlsm'
    Copy-Item $alvo $refEspec -Force

    Mostrar (& (Join-Path $s 'gerar_mEspecificacoes.ps1') `
        -Producao (Join-Path $arq 'src_producao\mEspecificacoes.bas') `
        -Produto $Produto `
        -Saida (Join-Path $hp 'mEspecificacoes.bas')) -Ultimas 3

    Etapa 'criar_estrutura_especificacoes.ps1' -Argumentos @('-Workbook', $alvo, '-Produto', $Produto) -Ultimas 4
    Encerrar-Excel
    & (Join-Path $s 'aplicar_vba.ps1') -Workbook $alvo -Modulos @((Join-Path $hp 'mEspecificacoes.bas'))
    Encerrar-Excel
    Etapa 'criar_form_especificacoes.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3
    Encerrar-Excel
    Etapa 'ligar_motor_especificacoes.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 2
    Encerrar-Excel
    Etapa 'rodar_motor.ps1' -Argumentos @('-Workbook', $alvo, '-Rotinas', 'AtualizarEngEspec,AtualizarBIData') -Ultimas 4
}

Encerrar-Excel
"== 7z. ADR-051: Eventos_Westgard e EQA_Base viram tabela nomeada"
# DEPOIS do motor: as tabelas sao dimensionadas pelo historico ja gravado.
# Antes dele, Eventos_Westgard estaria vazia e a tabela nasceria com uma linha
# em branco que so a proxima execucao corrigiria.
#
# ANTES da etapa 8: dali em diante a estrutura fica travada, e criar
# ListObject exige a pasta destravada.
Etapa 'criar_listobjects_dados.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 5

Encerrar-Excel
"== 8. trava a estrutura (abas so pelo Modo Desenvolvedor)"
Etapa 'travar_estrutura.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 3

Encerrar-Excel
"== 8b. estado seguro persistente para macros desabilitadas"
Etapa 'blindar_artefato.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 5

Encerrar-Excel
if ($Produto -eq 'Hematologia' -and -not $SemFixtures) {
    "== 8c. nao regressao das especificacoes (1.575 Westgard / 1.125 Sigma)"
    Etapa 'testar_especificacoes_hematologia.ps1' -Argumentos @(
        '-Referencia', $refEspec,
        '-Candidato', $alvo,
        '-OutCsv', (Join-Path $bd 'teste_especificacoes_hematologia.csv')
    ) -Ultimas 4
}
elseif ($Produto -eq 'Hematologia') {
    "  teste 1.575/1.125 pulado: -SemFixtures nao possui a massa controlada"
}

Encerrar-Excel
if ($Produto -eq 'Bioquimica') {
    "== 8d. campos de equipamento/serie/controle editaveis sob protecao"
    Etapa 'testar_release_visual_bio.ps1' -Argumentos @('-Workbook', $alvo) -Ultimas 2
}

""
"BUILD PRONTO: $alvo"
"Conferir com:"
"  .\snapshot_formulas.ps1 -Workbook '$alvo' -OutCsv `$env:TEMP\build.csv"
"  .\diff_formulas.ps1 -Referencia '$(Join-Path $snapProduto 'formulas.csv')' -Candidato `$env:TEMP\build.csv"
"Aceite esperado: AUSENTE 1080 - ALTERADA 4362 - VALOR 0 - EXTRA 9"
"  as 1.080 AUSENTE sao regRep2 e regRep3, removidas de proposito na etapa 6b"
