# verificar_tudo.ps1 - suite de verificacao do QC_INI, um comando so.
#
# POR QUE ESTE SCRIPT EXISTE. Em 03/08/2026 descobrimos que o motor corrigido
# nunca havia entrado no .xlsm -- em duas maquinas, por semanas -- enquanto o
# build reportava "Salvo:" a cada etapa. Havia lint de VBA, diff de formulas
# celula a celula e varredura de arquitetura, e nenhum controle verificava se a
# etapa anterior tinha ACONTECIDO. O pipeline reportava intencao, nao resultado.
#
# Esta suite verifica RESULTADO. Roda sem agente e sem depender de ninguem ler
# log de build: ao final escreve um relatorio com OK/FALHA por item do Quality
# Gate e devolve codigo de saida diferente de zero se algo falhou.
#
# EXIGE sessao de desktop ativa (automacao COM do Excel) e "Confiar no acesso ao
# modelo de objeto do projeto do VBA".
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\verificar_tudo.ps1                 (build + todas as verificacoes)
#   .\verificar_tudo.ps1 -PularBuild     (verifica o artefato que ja existe)

param(
    [switch]$PularBuild,
    # Produto alvo. Acompanha o -Produto do build_all.ps1: define o snapshot de
    # referencia, a pasta de build e o nome do arquivo.
    [string]$Produto = 'Hematologia'
)

$ErrorActionPreference = 'Stop'

$s = Split-Path -Parent $MyInvocation.MyCommand.Path
$arq = Split-Path -Parent $s
$h = Join-Path $arq 'src_hardening1'
# Modulos gerados por produto (mEstatistica com o NLV do setor, mDados).
$hp = Join-Path $h $Produto
$snap = Join-Path $arq "snapshot_producao\$Produto"
$bd = Join-Path $env:USERPROFILE "QCINI_build_hardening1_$Produto"
$alvo = Join-Path $bd "QC_$Produto.xlsm"
$relatorio = Join-Path $bd 'VERIFICACAO.md'

$itens = New-Object System.Collections.ArrayList
function Anotar {
    param([string]$Id, [string]$Nome, [bool]$Ok, [string]$Evidencia)
    [void]$itens.Add([pscustomobject]@{ Id = $Id; Nome = $Nome; Ok = $Ok; Evidencia = $Evidencia })
    $marca = if ($Ok) { 'OK   ' } else { 'FALHA' }
    "  [$marca] $Id  $Nome"
    if ($Evidencia) { "          $Evidencia" }
}

# Espera ATE nao sobrar processo, em vez de dormir um tempo fixo. Dormir "o
# suficiente" e chute: se o Excel demora mais, a etapa seguinte abre o arquivo
# travado em somente leitura.
function Encerrar-Excel {
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill() } catch { } }
    $limite = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $limite) {
        if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 400; return }
        Start-Sleep -Milliseconds 250
    }
    throw "Excel nao encerrou em 20s - o arquivo seguiria travado e a etapa seguinte gravaria no vazio"
}

# Normaliza codigo VBA para comparacao: tira cabecalho de atributos, espacos a
# direita e linhas vazias no fim. O que importa e o codigo, nao a formatacao que
# o Excel reescreve ao importar.
function Normalizar-VBA {
    param([string[]]$Linhas)
    $out = New-Object System.Collections.ArrayList
    foreach ($l in $Linhas) {
        if ($l -match '^\s*(VERSION|BEGIN|END|MultiUse|Attribute)\b') { continue }
        [void]$out.Add($l.TrimEnd())
    }
    while ($out.Count -gt 0 -and $out[$out.Count - 1] -eq '') { $out.RemoveAt($out.Count - 1) }
    while ($out.Count -gt 0 -and $out[0] -eq '') { $out.RemoveAt(0) }
    # Minusculas de proposito. O VBA e insensivel a maiusculas e o VBE UNIFICA a
    # grafia dos identificadores em todo o projeto ao importar: "ws.Rows.Count"
    # vira "ws.rows.Count" porque "rows" ja aparece assim noutro modulo. Comparar
    # respeitando caixa acusaria diferenca onde o codigo e o mesmo. (E a mesma
    # insensibilidade que transforma o identificador "aS" na palavra reservada As.)
    return (($out -join "`n").ToLowerInvariant())
}

function Hash-Texto {
    param([string]$Texto)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $b = [System.Text.Encoding]::UTF8.GetBytes($Texto)
    return ([BitConverter]::ToString($sha.ComputeHash($b))).Replace('-', '').ToLower()
}

""
"================================================================"
"  VERIFICACAO DO QC_INI - $Produto"
"  $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
"================================================================"

# ---------------------------------------------------------------- 0. build ----
if (-not $PularBuild) {
    ""
    "-- 0. BUILD ----------------------------------------------------"
    $saidaBuild = & (Join-Path $s 'build_all.ps1') -Produto $Produto 2>&1
    $okBuild = $?
    $saidaBuild | Where-Object { $_ -match '^==|^BUILD PRONTO' } | ForEach-Object { "     $_" }
    Anotar '0.1' 'build_all.ps1 concluiu' $okBuild ''
}
if (-not (Test-Path $alvo)) { throw "Artefato nao encontrado: $alvo (rode sem -PularBuild)" }

# O bloco 3 GRAVA no Audit_Log e adultera uma linha de proposito. Isso nao pode
# acontecer no artefato entregavel: a verificacao roda sobre uma copia, e o
# artefato sai da suite exatamente como o build o deixou.
Encerrar-Excel
$alvoOriginal = $alvo
$copia = Join-Path $bd "QC_${Produto}_verificacao.xlsm"
Copy-Item $alvo $copia -Force
$alvo = $copia

# ------------------------------------------------- 1. artefato x fonte --------
""
"-- 1. O ARTEFATO CONTEM O QUE A FONTE DIZ ----------------------"
"   (o controle que faltava: comparar o VBA DENTRO do .xlsm com o versionado)"

Encerrar-Excel
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
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
# 1 = msoAutomationSecurityLow. NAO usar 3 (msoAutomationSecurityForceDisable):
# ele DESABILITA as macros, e todo Application.Run falha com "as macros foram
# desabilitadas" -- que parece problema de confianca do Office e e so este valor.
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($alvo)

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

if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Artefato aberto em somente leitura" }

# A trava de estrutura (abas so pelo Modo Desenvolvedor) impede reexibir aba
# oculta, e varios testes precisam disso. Desprotege a estrutura da COPIA
# de verificacao durante os testes; nao afeta o entregavel.
$estruturaEstava = $wb.ProtectStructure
if ($estruturaEstava) { $wb.Unprotect('qcini2025') }

try {
    # mEstatistica e mDados sao gerados POR PRODUTO (a Bioquimica tem NLV=2) e
    # ficam em src_hardening1/<Produto>/. mAuditoria e compartilhado: nao
    # depende de geometria.
    $modulos = @(
        @{ Nome = 'mEstatistica'; Arquivo = (Join-Path $hp 'mEstatistica.bas') },
        @{ Nome = 'mDados'; Arquivo = (Join-Path $hp 'mDados.bas') },
        @{ Nome = 'mAuditoria'; Arquivo = (Join-Path $hp 'mAuditoria.bas') },
        @{ Nome = 'mImportar'; Arquivo = (Join-Path $hp 'mImportar.bas') }
    )

    # mEspecificacoes entra por OUTRO caminho: vem de src_producao e chega ao
    # artefato pela copia da producao (ADR-021), nao pelo aplicar_vba. Ficava
    # portanto FORA desta prova -- e era o unico modulo de calculo cuja deriva
    # entre a fonte versionada e o que roda dentro do arquivo ninguem media.
    #
    # So a Bioquimica tem o motor de especificacoes. Cobrar isso da Hematologia
    # reprovaria um produto por nao ter algo que ainda nao lhe foi entregue.
    $esperaEspec = ($Produto -eq 'Bioquimica')
    if ($esperaEspec) {
        $modulos += @{ Nome = 'mEspecificacoes'; Arquivo = (Join-Path $arq 'src_producao\mEspecificacoes.bas') }
        # mEstatPeriodo chega pelo mesmo caminho (copia da producao) e e a
        # camada que decide n, CV, Bias e ET de cada analito. Sem esta prova, a
        # fonte versionada e o que roda dentro do arquivo podem divergir sem que
        # nada acuse -- o defeito que o ADR-021 existe para impedir.
        $modulos += @{ Nome = 'mEstatPeriodo'; Arquivo = (Join-Path $arq 'src_producao\mEstatPeriodo.bas') }
    }
    foreach ($m in $modulos) {
        $comp = $null
        foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq $m.Nome) { $comp = $c; break } }
        if ($comp -eq $null) {
            Anotar '1.1' "modulo $($m.Nome) presente" $false 'componente ausente do projeto VBA'
            continue
        }
        $noArquivo = Normalizar-VBA ([System.IO.File]::ReadAllLines($m.Arquivo, [System.Text.Encoding]::Default))
        $noXlsm = Normalizar-VBA ($comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines) -split "`r?`n")
        $hA = Hash-Texto $noArquivo
        $hX = Hash-Texto $noXlsm
        Anotar '1.1' "$($m.Nome) identico a fonte" ($hA -eq $hX) "fonte $($hA.Substring(0,12)) / artefato $($hX.Substring(0,12))"
    }

    # formularios e botoes da Sprint NC
    $formsNC = @('frmResultadoNaoConforme', 'frmExcluirRegistroNC')
    foreach ($fn in $formsNC) {
        $achou = $false
        foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq $fn) { $achou = $true } }
        Anotar '1.4' "formulario $fn presente" $achou ''
    }
    foreach ($par in @(@('Resultados', 'btnNaoConforme'), @('Registros', 'btnExcluirNC'))) {
        $wsB = $wb.Worksheets.Item($par[0])
        $temBotao = $false
        $acao = ''
        foreach ($sh in $wsB.Shapes) { if ($sh.Name -eq $par[1]) { $temBotao = $true; $acao = $sh.OnAction } }
        Anotar '1.5' "botao $($par[1]) na aba $($par[0])" ($temBotao -and $acao -ne '') "OnAction = '$acao'"
    }

    # aS como identificador: a classe de erro que derrubava o projeto inteiro
    $comAS = 0
    foreach ($c in $wb.VBProject.VBComponents) {
        if ($c.CodeModule.CountOfLines -gt 0) {
            $txt = $c.CodeModule.Lines(1, $c.CodeModule.CountOfLines)
            $comAS += ([regex]::Matches($txt, '(?m)^(?!\s*'').*\baS\b\s+As\s')).Count
        }
    }
    Anotar '1.2' 'nenhum identificador aS/As no projeto' ($comAS -eq 0) "ocorrencias: $comAS"

    # abas obrigatorias
    $abas = @()
    foreach ($ws in $wb.Worksheets) { $abas += $ws.Name }
    foreach ($n in @('Eng_Saida', 'Corridas', 'Audit_Log', 'Importar')) {
        Anotar '1.3' "aba $n existe" ($abas -contains $n) ''
    }

    # --- aba Importar: a substituta do frmMassa -------------------------------
    # O usuario cola os dados na horizontal (uma corrida por linha) e clica em
    # Registrar. Aqui so a MONTAGEM e conferida; que ela grava (e que recusa
    # colagem invalida) e provado em 3.16, com o motor rodando.
    if ($abas -contains 'Importar') {
        $wsI = $wb.Worksheets.Item('Importar')

        # O cabecalho vem do CSV de definicao (ordem fixa do gestor), e a linha
        # 3 -- oculta -- diz para qual analito CADASTRADO cada coluna vai. Se o
        # de/para apontar para nome que nao existe na aba Analitos, o resultado
        # vira linha orfa que nenhum calculo encontra.
        # O CSV E O DO PRODUTO. Estava fixo em analitos_bioquimica.csv, e por
        # isso esta prova CONFIRMAVA o defeito em vez de pega-lo: a aba da
        # Hematologia era montada com os 31 analitos da Bioquimica e comparada
        # contra a mesma lista errada -- batia, e passava.
        $csvDef = Join-Path $arq "src_producao\analitos_$($Produto.ToLowerInvariant()).csv"
        if (-not (Test-Path $csvDef)) { throw "CSV de analitos ausente para $Produto : $csvDef" }
        $ld = [System.IO.File]::ReadAllLines($csvDef, [System.Text.Encoding]::UTF8)
        $sig = @(); $nom = @()
        for ($i = 1; $i -lt $ld.Count; $i++) {
            if ($ld[$i].Trim() -eq '') { continue }
            $p = $ld[$i].Split(','); $sig += $p[1]; $nom += $p[2]
        }
        $esperado = @('DATA', 'NIVEL', 'LOTE') + $sig
        $lido = @()
        for ($k = 0; $k -lt $esperado.Count; $k++) {
            $lido += "$($wsI.Cells.Item(4, $k + 2).Value2)"
        }
        $bate = ($lido -join '|') -eq ($esperado -join '|')
        Anotar '1.7' 'cabecalho da aba Importar segue a ordem definida' $bate `
            "$($sig.Count) siglas + DATA/NIVEL/LOTE"

        # de/para: toda coluna aponta para analito que existe
        $wsAn = $wb.Worksheets.Item('Analitos')
        $cad = @{}
        for ($i = 4; $i -le 43; $i++) {
            $v = $wsAn.Cells.Item($i, 1).Value2
            if ($v -ne $null -and "$v".Trim() -ne '') { $cad["$v".Trim()] = $true }
        }
        # COLUNA SEM DE/PARA E DEFEITO, NAO CASO PREVISTO.
        #
        # Aqui havia um "continue" para o de/para vazio, com a justificativa de
        # que o runtime recusaria. Recusa mesmo -- e esse era o problema: a aba
        # da Hematologia saiu com as 31 colunas da Bioquimica e NENHUMA mapeada,
        # e esta prova relatou "0 orfas" com ar de aprovacao. Uma coluna que nao
        # aponta para lugar nenhum nunca vai importar; que o runtime a recuse
        # depois nao torna a montagem correta.
        $orfas = @()
        $vazias = @()
        for ($k = 0; $k -lt $nom.Count; $k++) {
            $d = "$($wsI.Cells.Item(3, 5 + $k).Value2)".Trim()
            if ($d -eq '') { $vazias += $sig[$k]; continue }
            if (-not $cad.ContainsKey($d)) { $orfas += "$($sig[$k]) -> '$d'" }
        }
        $detalheMapa = "$($orfas.Count) orfa(s), $($vazias.Count) sem de/para"
        if ($orfas.Count) { $detalheMapa += " | orfas: $($orfas -join '; ')" }
        if ($vazias.Count) { $detalheMapa += " | vazias: $(($vazias | Select-Object -First 6) -join ', ')" }
        Anotar '1.11' 'toda coluna da aba Importar aponta para analito cadastrado' `
            (($orfas.Count -eq 0) -and ($vazias.Count -eq 0)) $detalheMapa
        Anotar '1.12' 'linha de/para fica oculta' $wsI.Rows.Item(3).Hidden ''

        # o botao que dispara a migracao
        $acaoReg = ''
        foreach ($sh in $wsI.Shapes) { if ($sh.Name -eq 'btnRegistrar') { $acaoReg = $sh.OnAction } }
        Anotar '1.8' 'botao Registrar na aba Importar' ($acaoReg -eq 'RegistrarImportacao') `
            "OnAction = '$acaoReg'"

        # o caminho de ida: o botao da aba Resultados nao pode mais abrir o
        # frmMassa, que foi substituido
        $acaoIda = ''
        $aindaMassa = $false
        foreach ($sh in $wb.Worksheets.Item('Resultados').Shapes) {
            if ($sh.OnAction -eq 'IrParaImportar') { $acaoIda = $sh.OnAction }
            if ($sh.OnAction -eq 'AbrirFormMassa') { $aindaMassa = $true }
        }
        Anotar '1.9' 'botao da aba Resultados leva para a aba Importar' `
            ($acaoIda -eq 'IrParaImportar' -and -not $aindaMassa) `
            "IrParaImportar=$($acaoIda -ne ''); AbrirFormMassa restante=$aindaMassa"

        # area de entrada destravada dentro de aba protegida: o usuario digita
        # so onde deve
        $entradaLivre = (-not $wsI.Cells.Item(5, 2).Locked)
        $cabTravado = $wsI.Cells.Item(4, 2).Locked
        Anotar '1.10' 'na aba Importar so a area de colagem e editavel' `
            ($entradaLivre -and $cabTravado -and $wsI.ProtectContents) `
            "entrada livre=$entradaLivre; cabecalho travado=$cabTravado; aba protegida=$($wsI.ProtectContents)"
    }

    # ------------------------------------------------ 2. motor executa --------
    ""
    "-- 2. O MOTOR EXECUTA ------------------------------------------"
    foreach ($rot in @('AtualizarCalc', 'AtualizarPainelEng', 'AtualizarEstatisticaAba', 'RegistrarEventosWestgard')) {
        $t0 = Get-Date
        $falhou = $false
        $msg = ''
        try { $xl.Run($rot) | Out-Null } catch { $falhou = $true; $msg = $_.Exception.Message }
        $dt = ((Get-Date) - $t0).TotalSeconds
        Anotar '2.1' "$rot" (-not $falhou) $(if ($falhou) { $msg } else { ('{0:N2}s' -f $dt) })
    }

    # Eng_Saida abastecida e com a chave logica ANALITO|RUN
    $eng = $wb.Worksheets.Item('Eng_Saida')
    $chave = [string]$eng.Cells.Item(3, 28).Value2
    Anotar '2.2' 'Eng_Saida publica chave ANALITO|RUN' ($chave -match '^.+\|\d+$') "AB3 = '$chave'"

    # ------------------------------------------- 3. auditoria de verdade ------
    ""
    "-- 3. A TRILHA DE AUDITORIA FUNCIONA ---------------------------"
    "   (nao basta gravar: adulterar uma linha tem que ser DETECTADO)"

    $au = $wb.Worksheets.Item('Audit_Log')
    $au.Visible = -1

    # 3.1 parecer tecnico curto e recusado
    $p3 = $xl.Run('ParecerValido', 'apenas tres palavras')
    $p6 = $xl.Run('ParecerValido', 'erro comprovado de pipetagem no nivel dois')
    Anotar '3.1' 'parecer curto recusado, parecer valido aceito' ((-not $p3) -and $p6) "3 palavras=$p3 / 7 palavras=$p6"

    # 3.2 grava tres eventos
    $antes = $xl.Run('UltimaLinhaAudit')
    for ($i = 1; $i -le 3; $i++) {
        $xl.Run('Auditar', 'DADO', 'TESTE_SUITE', 'suite', $i, (Get-Date), '', 'QC-52261101', 1, 'WBC',
            3.01, 3.5, 'Ativo', 'Excluido', 'Teste automatizado',
            'Registro de teste automatizado da suite de verificacao') | Out-Null
    }
    $depois = $xl.Run('UltimaLinhaAudit')
    Anotar '3.2' 'log grava e cresce' (($depois - $antes) -eq 3) "linhas: $antes -> $depois"

    # 3.3 cadeia integra
    $v1 = [string]$xl.Run('VerificarIntegridadeLog')
    Anotar '3.3' 'cadeia integra apos gravacao' ($v1 -like 'OK|*') "retorno: $v1"

    # 3.4 ADULTERACAO E DETECTADA -- o teste que da sentido ao resto
    $linhaAlvo = $depois - 1
    $original = $au.Cells.Item($linhaAlvo, 24).Value2       # ParecerTecnico (schema 2)
    $au.Unprotect('qcini2025')
    $au.Cells.Item($linhaAlvo, 24).Value2 = 'Parecer adulterado fora do sistema'
    $v2 = [string]$xl.Run('VerificarIntegridadeLog')
    $detectou = ($v2 -like "QUEBRADO|$linhaAlvo|*")
    Anotar '3.4' 'ADULTERACAO DETECTADA na linha certa' $detectou "retorno: $v2"

    # restaura e confirma que volta a fechar
    $au.Cells.Item($linhaAlvo, 24).Value2 = $original
    $v3 = [string]$xl.Run('VerificarIntegridadeLog')
    Anotar '3.5' 'cadeia volta a fechar apos restaurar' ($v3 -like 'OK|*') "retorno: $v3"

    # ---- 3.6/3.7 item 2.3: reenvio NAO ressuscita registro excluido ----
    # O UpsertResultados antigo forcava Status = Ativo em toda atualizacao:
    # reenviar a mesma chave trazia de volta uma linha excluida, sem rastro.
    # Era vetor de fraude. Aqui a suite exclui um registro real e reenvia o
    # mesmo valor, exigindo que o Status continue Excluido e que a tentativa
    # apareca no log como UPSERT_BLOQUEADO.
    $db = $wb.Worksheets.Item('DB_Resultados')
    $linhaTeste = 4
    $runT = [int]$db.Cells.Item($linhaTeste, 1).Value2
    $nivelT = [int]$db.Cells.Item($linhaTeste, 3).Value2
    $loteT = [string]$db.Cells.Item($linhaTeste, 4).Value2
    $analT = [string]$db.Cells.Item($linhaTeste, 5).Value2
    $valorT = $db.Cells.Item($linhaTeste, 6).Value2
    $dataT = $db.Cells.Item($linhaTeste, 2).Value2
    $statusOriginal = [string]$db.Cells.Item($linhaTeste, 7).Value2

    $db.Cells.Item($linhaTeste, 7).Value2 = 'Excluido'
    $antesLog = $xl.Run('UltimaLinhaAudit')

    # reenvio pela camada de dados, exatamente como um formulario faria
    $vba = $wb.VBProject.VBComponents.Add(1)
    $vba.Name = 'mTesteUpsert'
    $vba.CodeModule.AddFromString(@'
Public Function T_Reenvio(ByVal r As Long, ByVal dt As Variant, ByVal nv As Long, _
                          ByVal lt As String, ByVal an As String, ByVal vl As Variant) As String
    Dim regs(1 To 1, 1 To 7) As Variant
    regs(1, 1) = r: regs(1, 2) = dt: regs(1, 3) = nv
    regs(1, 4) = lt: regs(1, 5) = an: regs(1, 6) = vl: regs(1, 7) = "Ativo"
    T_Reenvio = UpsertResultados(regs)
End Function
'@)
    Start-Sleep -Milliseconds 300
    $ret = [string]$xl.Run('T_Reenvio', $runT, $dataT, $nivelT, $loteT, $analT, $valorT)
    $statusDepois = [string]$db.Cells.Item($linhaTeste, 7).Value2
    $depoisLog = $xl.Run('UltimaLinhaAudit')

    Anotar '3.6' 'reenvio NAO ressuscita registro excluido' ($statusDepois -eq 'Excluido') `
        "status antes=Excluido depois=$statusDepois | UpsertResultados devolveu '$ret' (novos|atualizados|bloqueados)"

    $acao = ''
    if ($depoisLog -gt $antesLog) { $acao = [string]$au.Cells.Item($depoisLog, 7).Value2 }
    Anotar '3.7' 'tentativa de reenvio fica registrada no log' ($acao -eq 'REENVIO_BLOQUEADO') `
        "acao registrada: '$acao'"

    # ---- 3.10/3.11 duas camadas de log com o MESMO identificador ----
    # O Event Store prova integridade; as tabelas do banco preservam a estrutura
    # por origem. Se as duas nao compartilhassem o ID, nao haveria como provar
    # que uma nao contradiz a outra.
    $logResAntes = $xl.Run('ContarLogDB', 'Resultados')
    $vbaEx = $wb.VBProject.VBComponents.Add(1)
    $vbaEx.Name = 'mTesteExcl'
    $vbaEx.CodeModule.AddFromString(@'
Public Function T_Excluir(ByVal r As Long, ByVal nv As Long, ByVal an As String) As Long
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(UCase$(Trim$(an))) = 1
    T_Excluir = ExcluirLogico(r, nv, d, _
        "Exclusao de teste automatizado da suite de verificacao", _
        "Excluido", "Teste automatizado", "Resultados")
End Function
'@)
    Start-Sleep -Milliseconds 300
    $exc = $xl.Run('T_Excluir', $runT, $nivelT, $analT)
    $logResDepois = $xl.Run('ContarLogDB', 'Resultados')
    Anotar '3.10' 'exclusao grava tambem em LOG_Resultados' (($logResDepois - $logResAntes) -eq $exc) `
        "linhas no log do banco: $logResAntes -> $logResDepois (excluidos: $exc)"

    $linhaLog = $xl.Run('UltimaLinhaLogDB', 'Resultados')
    $idNoBanco = [string]$db.Cells.Item($linhaLog, 59).Value2      # BG = ID_Auditoria
    $idNoEvento = [string]$au.Cells.Item($xl.Run('UltimaLinhaAudit'), 1).Value2
    Anotar '3.11' 'as duas camadas compartilham o ID_Auditoria' ($idNoBanco -eq $idNoEvento -and $idNoBanco -ne '') `
        "banco='$idNoBanco' evento='$idNoEvento'"

    $wb.VBProject.VBComponents.Remove($vbaEx)
    $db.Cells.Item($linhaTeste, 7).Value2 = $statusOriginal

    # ---- 3.12 a 3.15 Sprint NC: marcar nao conforme, de ponta a ponta ----
    # O que precisa ser verdade: o resultado SAI DOS CALCULOS, o valor ORIGINAL
    # permanece no banco, a ocorrencia aparece na aba Registros, e as tres
    # camadas registram com o MESMO identificador.
    $dbNC = $wb.Worksheets.Item('DB_Resultados')
    $regNC = $wb.Worksheets.Item('Registros')
    $regNC.Visible = -1
    $linNC = 10
    $runNC = [int]$dbNC.Cells.Item($linNC, 1).Value2
    $nivNC = [int]$dbNC.Cells.Item($linNC, 3).Value2
    $anaNC = [string]$dbNC.Cells.Item($linNC, 5).Value2
    $valNC = $dbNC.Cells.Item($linNC, 6).Value2
    $stNCantes = [string]$dbNC.Cells.Item($linNC, 7).Value2

    $regAntes = $xl.Run('PrimeiraLinhaLivreReg')
    $logAntes2 = $xl.Run('ContarLogDB', 'Resultados')

    # parecer curto: a operacao TEM de ser recusada
    $idVazio = [string]$xl.Run('MarcarNaoConforme', $runNC, $nivNC, $anaNC, 'Invalido', 'muito curto')
    Anotar '3.12' 'parecer curto impede marcar nao conforme' ($idVazio -eq '') `
        "retorno vazio, status permaneceu '$([string]$dbNC.Cells.Item($linNC,7).Value2)'"

    $idNC = [string]$xl.Run('MarcarNaoConforme', $runNC, $nivNC, $anaNC, 'Invalido',
        'Material hemolisado confirmado pela equipe tecnica')
    $stNCdepois = [string]$dbNC.Cells.Item($linNC, 7).Value2
    $valDepois = $dbNC.Cells.Item($linNC, 6).Value2

    Anotar '3.13' 'resultado sai dos calculos e o valor ORIGINAL fica' `
        (($stNCdepois -eq 'Invalido') -and ("$valDepois" -eq "$valNC")) `
        "status $stNCantes -> $stNCdepois | valor preservado: $valDepois"

    $regDepois = $xl.Run('PrimeiraLinhaLivreReg')
    $okReg = ($regDepois -gt $regAntes)
    $anaReg = ''
    if ($okReg) { $anaReg = [string]$regNC.Cells.Item($regAntes, 3).Value2 }
    Anotar '3.14' 'ocorrencia aparece na aba Registros' ($okReg -and $anaReg -eq $anaNC) `
        "linha ${regAntes}: analito '$anaReg', RUN $([string]$regNC.Cells.Item($regAntes,5).Value2)"

    $logDepois2 = $xl.Run('ContarLogDB', 'Resultados')
    $idBanco2 = [string]$dbNC.Cells.Item($xl.Run('UltimaLinhaLogDB', 'Resultados'), 59).Value2
    Anotar '3.15' 'as tres camadas registram com o mesmo ID' `
        (($logDepois2 -gt $logAntes2) -and ($idBanco2 -eq $idNC) -and ($idNC -ne '')) `
        "ID: '$idNC' | log do banco: $logAntes2 -> $logDepois2"

    # desfaz: devolve o status e limpa a vitrine
    $dbNC.Cells.Item($linNC, 7).Value2 = $stNCantes
    try { $regNC.Unprotect('qcini2025') } catch { }
    $regNC.Range($regNC.Cells.Item($regAntes, 2), $regNC.Cells.Item($regAntes, 13)).ClearContents()
    $regNC.Visible = 2

    # ---- 3.16 e 3.17 importacao por aba, de ponta a ponta ----
    # O pedido do gestor era simples: "esse botao faz com que os dados sumam
    # dali e ja migrem para o DB_Resultados". As duas metades disso sao
    # verificadas aqui -- migra o que esta valido, e NAO migra nada quando ha
    # qualquer erro (tudo-ou-nada).
    #
    # Chama ExecutarImportacao, que e o mesmo nucleo do botao Registrar; o botao
    # so acrescenta MsgBox, que travaria a execucao sem interface.
    $wsImp = $wb.Worksheets.Item('Importar')
    $wsImp.Visible = -1
    $lotesImp = $xl.Run('ListaLotes')
    $loteImp = [string]$lotesImp.Item(1)
    $analitosImp = $xl.Run('ListaAnalitos')

    # Escrita inline e com conversao explicita. Uma funcao auxiliar aqui esbarra
    # na conversao do COM (o setter de Value2 e propriedade parametrizada e a
    # ligacao tardia nao aceita o Int32 vindo de parametro nao tipado).
    $escreve = {
        param($Ws, $Linha, $Data, $Nivel, $Lote, $Valores)
        $Ws.Cells.Item([int]$Linha, 2).Value2 = [string]$Data
        $Ws.Cells.Item([int]$Linha, 3).Value2 = [double]$Nivel
        $Ws.Cells.Item([int]$Linha, 4).Value2 = [string]$Lote
        for ($k = 0; $k -lt $Valores.Count; $k++) {
            $Ws.Cells.Item([int]$Linha, 5 + $k).Value2 = [double]$Valores[$k]
        }
    }

    $ultDbAntes = $dbNC.Cells.Item($dbNC.Rows.Count, 1).End(-4162).Row
    $auditAntesImp = [int]$xl.Run('UltimaLinhaAudit')

    # (a) colagem valida: 2 niveis x 3 analitos = 6 resultados
    # Dados comecam na linha 5: a 3 e o de/para oculto e a 4 e o cabecalho.
    & $escreve $wsImp 5 '14/12/2029' 1 $loteImp @(90, 35, 2.6)
    & $escreve $wsImp 6 '14/12/2029' 2 $loteImp @(250, 80, 5.2)
    $retImp = [string]$xl.Run('ExecutarImportacao', $true)
    $ultDbOk = $dbNC.Cells.Item($dbNC.Rows.Count, 1).End(-4162).Row
    $abaLimpa = ("$($wsImp.Cells.Item(5,2).Value2)" -eq '' -and "$($wsImp.Cells.Item(6,2).Value2)" -eq '')
    Anotar '3.16' 'colagem valida migra para o DB_Resultados e some da aba' `
        ($retImp -like 'OK|*' -and ($ultDbOk - $ultDbAntes) -eq 6 -and $abaLimpa) `
        "retorno '$retImp' | banco $ultDbAntes -> $ultDbOk | aba limpa: $abaLimpa"

    # ---- 3.18 a importacao em massa deixa rastro ----
    #
    # A importacao e o caminho de entrada de MAIOR volume do sistema: uma
    # colagem grava dezenas de resultados de uma vez. ISO 15189 8.4 nao admite
    # entrada de dado sem registro de quem, quando e o que.
    #
    # A prova existe porque a lacuna existiu: na PRODUCAO, RegistrarLog e um
    # stub vazio, e a chamada em ExecutarImportacao nao gravava nada. Quem faz
    # dela um registro real e a camada de hardening, aplicada pelo build -- e
    # ate aqui ninguem media se ela de fato chegava ate a importacao. 3.16 e
    # 3.17 provavam que o dado migra; nenhum provava que o ato fica registrado.
    $auditDepoisImp = [int]$xl.Run('UltimaLinhaAudit')
    $acaoImp = ''
    if ($auditDepoisImp -gt $auditAntesImp) { $acaoImp = [string]$au.Cells.Item($auditDepoisImp, 7).Value2 }
    Anotar '3.18' 'importacao em massa deixa rastro na trilha de auditoria' `
        ($auditDepoisImp -gt $auditAntesImp) `
        "Audit_Log $auditAntesImp -> $auditDepoisImp, ultima acao '$acaoImp'"

    # (b) colagem com UM erro: nada pode ser gravado, nem a linha boa
    & $escreve $wsImp 5 '21/12/2029' 1 $loteImp @(91, 36, 2.7)
    & $escreve $wsImp 6 '21/12/2029' 2 'LOTE_INEXISTENTE' @(251, 81, 5.3)
    $retErr = [string]$xl.Run('ExecutarImportacao', $true)
    $ultDbErr = $dbNC.Cells.Item($dbNC.Rows.Count, 1).End(-4162).Row
    # Count de uma Collection do VBA e METODO, nao propriedade: sem os
    # parenteses a ligacao tardia devolve o objeto do metodo, nao o numero.
    $cErr = 5 + $analitosImp.Count() + 1
    $temMsg = ("$($wsImp.Cells.Item(5, $cErr).Value2)" -ne '')
    $preservou = ("$($wsImp.Cells.Item(5,2).Value2)" -ne '')
    Anotar '3.17' 'colagem com erro nao grava NADA (nem a linha valida)' `
        ($retErr -like 'ERRO|*' -and $ultDbErr -eq $ultDbOk -and $temMsg -and $preservou) `
        "retorno '$retErr' | banco intacto em $ultDbErr | erro apontado na aba: $temMsg"

    # Desfaz: limpa SO as colunas do banco (A:G) nas linhas de teste.
    #
    # Nao usar Rows.Delete aqui. O DB_Resultados nao e so A:G: tem o bloco de
    # formulas BA:BD ate a linha ~15000 e as duas tabelas de log em BG..BY e
    # CB..CT. Apagar a linha inteira desloca tudo isso para cima e destroi
    # formulas que nada tem a ver com o teste -- foi o que aconteceu em
    # 05/08/2026 e derrubou o item 4.1 com 18 formulas AUSENTES.
    try { $dbNC.Unprotect('qcini2025') } catch { }
    if ($ultDbOk -gt $ultDbAntes) {
        $dbNC.Range($dbNC.Cells.Item($ultDbAntes + 1, 1), $dbNC.Cells.Item($ultDbOk, 7)).ClearContents() | Out-Null
    }
    try { $wsImp.Unprotect('qcini2025') } catch { }
    $wsImp.Range($wsImp.Cells.Item(5, 2), $wsImp.Cells.Item(204, $cErr)).ClearContents()
    $wsImp.Protect('qcini2025', $true, $true, $true, $true)

    # ---- 3.8/3.9 item 2.5: alterar a elegibilidade fica registrado ----
    # Cfg_Status decide o que entra em media/DP/CV/Bias/Sigma/Westgard. Mudar uma
    # celula redefine RETROATIVAMENTE a estatistica de todo o historico. A
    # exigencia nao e impedir -- a tabela existe para ser extensivel (ADR-006) --
    # e sim que toda mudanca fique versionada e registrada.
    $cfg = $wb.Worksheets.Item('Cfg_Status')
    $cfg.Visible = -1
    $versaoAntes = $xl.Run('VersaoCfg')
    $antesLog2 = $xl.Run('UltimaLinhaAudit')

    try { $cfg.Unprotect('qcini2025') } catch { }
    $elegOriginal = [string]$cfg.Cells.Item(11, 3).Value2      # Treinamento: NAO
    $xl.EnableEvents = $true
    $cfg.Cells.Item(11, 3).Value2 = 'SIM'
    Start-Sleep -Milliseconds 400
    $xl.EnableEvents = $false

    $versaoDepois = $xl.Run('VersaoCfg')
    $depoisLog2 = $xl.Run('UltimaLinhaAudit')
    $acaoCfg = ''
    if ($depoisLog2 -gt $antesLog2) { $acaoCfg = [string]$au.Cells.Item($depoisLog2, 7).Value2 }

    Anotar '3.8' 'alterar elegibilidade sobe a versao da configuracao' ($versaoDepois -gt $versaoAntes) `
        "versao $versaoAntes -> $versaoDepois"
    Anotar '3.9' 'alteracao de elegibilidade vai para a auditoria' ($acaoCfg -like 'CFG_*') `
        "acao registrada: '$acaoCfg'"

    # restaura
    $xl.EnableEvents = $true
    $cfg.Cells.Item(11, 3).Value2 = $elegOriginal
    Start-Sleep -Milliseconds 400
    $xl.EnableEvents = $false
    $cfg.Visible = 2

    # ================= 7. MOTOR DE ESPECIFICACOES (ADR-022 / ADR-023) =========
    #
    # 575 linhas que decidem CONFORME / NAO CONFORME e que, ate aqui, nao tinham
    # UMA prova. E a camada mais nova e a de maior consequencia: Painel e
    # Estatistica passaram a ler dela, entao um erro aqui muda veredito de
    # qualidade sem mudar nada visivel no resto do sistema.
    #
    # Ausencia do modulo NAO e motivo para pular: e o achado. mEspecificacoes nao
    # tem instalador versionado e chegou ao arquivo pela VBE -- exatamente o que
    # o ADR-021 proibe. Se ele sumir num rebuild, estes itens reprovam.
    $ESP_ANALITO = 'ZZ_PROVA_ESPEC'
    $espOk = $true
    $espMsg = ''
    $wsEsp = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq 'DB_Especificacoes') { $wsEsp = $w; break } }
    $temModulo = $false
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq 'mEspecificacoes') { $temModulo = $true; break } }

    if (-not $esperaEspec) {
        # Produto que ainda nao recebeu o motor de especificacoes. Nao anotar
        # nada e o certo: um item que nao se aplica nao e uma prova que passou.
        "     (grupo 7 nao se aplica a $Produto - motor de especificacoes e da Bioquimica)"
    }
    elseif ($wsEsp -eq $null -or -not $temModulo) {
        $falta = @()
        if ($wsEsp -eq $null) { $falta += 'aba DB_Especificacoes' }
        if (-not $temModulo) { $falta += 'modulo mEspecificacoes' }
        foreach ($id in @('7.1', '7.2', '7.3', '7.4', '7.5')) {
            Anotar $id 'motor de especificacoes presente no artefato' $false "ausente: $($falta -join ' e ')"
        }
    }
    else {
        try { $wsEsp.Unprotect('qcini2025') } catch { }
        $espVisAntes = $wsEsp.Visible
        $wsEsp.Visible = -1
        $linEspAntes = [int]$xl.Run('UltimaLinhaEspec')

        # escreve uma linha do banco de especificacoes
        $gravaEsp = {
            param($Ws, $Lin, $Ano, $Fonte, $Modelo, $Analito, $Etp, $Cvi, $Cvg, $Rigor, $Cvtp, $Biastp, $Ativo)
            $Ws.Cells.Item([int]$Lin, 1).Value2 = [string]("ZZP-" + $Lin)
            $Ws.Cells.Item([int]$Lin, 2).Value2 = [double]$Ano
            $Ws.Cells.Item([int]$Lin, 3).Value2 = [string]$Fonte
            $Ws.Cells.Item([int]$Lin, 4).Value2 = [string]$Modelo
            $Ws.Cells.Item([int]$Lin, 5).Value2 = [string]$Analito
            if ($Etp -ne $null) { $Ws.Cells.Item([int]$Lin, 6).Value2 = [double]$Etp }
            if ($Cvi -ne $null) { $Ws.Cells.Item([int]$Lin, 7).Value2 = [double]$Cvi }
            if ($Cvg -ne $null) { $Ws.Cells.Item([int]$Lin, 8).Value2 = [double]$Cvg }
            $Ws.Cells.Item([int]$Lin, 9).Value2 = [string]$Rigor
            if ($Cvtp -ne $null) { $Ws.Cells.Item([int]$Lin, 10).Value2 = [double]$Cvtp }
            if ($Biastp -ne $null) { $Ws.Cells.Item([int]$Lin, 11).Value2 = [double]$Biastp }
            $Ws.Cells.Item([int]$Lin, 12).Value2 = [string]$Ativo
            $Ws.Cells.Item([int]$Lin, 13).Value2 = 'prova'
        }

        $l1 = $linEspAntes + 1
        $l2 = $linEspAntes + 2
        & $gravaEsp $wsEsp $l1 2020 'CLIA' 'ETP_DIRETO' $ESP_ANALITO 30 $null $null '' $null $null 'Sim'
        & $gravaEsp $wsEsp $l2 2026 'CLIA' 'ETP_DIRETO' $ESP_ANALITO 60 $null $null '' $null $null 'Sim'
        $xl.Run('InvalidarCacheEspec') | Out-Null

        # ---- 7.1 a meta e a do ANO DO RESULTADO, nunca a do ano corrente ----
        # Requisito central do ADR-022: reabrir uma corrida de 2023 em 2030 tem
        # de julga-la pela meta de 2023. A regra e de VIGENCIA -- maior Ano <=
        # ano do resultado -- e nao casamento exato, que abriria buraco em todo
        # ano sem cadastro.
        $r2023 = [string]$xl.Run('ResolverEspec', $ESP_ANALITO, 2023, 'CLIA')
        $r2026 = [string]$xl.Run('ResolverEspec', $ESP_ANALITO, 2026, 'CLIA')
        $r2019 = [string]$xl.Run('ResolverEspec', $ESP_ANALITO, 2019, 'CLIA')
        $p2023 = $r2023 -split '\|'
        $p2026 = $r2026 -split '\|'
        $vig2023 = ($p2023.Count -gt 5 -and $p2023[5] -eq '2020' -and [double]$p2023[3] -eq 30)
        $vig2026 = ($p2026.Count -gt 5 -and $p2026[5] -eq '2026' -and [double]$p2026[3] -eq 60)
        $antesDeTudo = ($r2019 -eq 'NAO_CADASTRADA')
        Anotar '7.1' 'meta resolvida pelo ano do resultado, com regra de vigencia' `
            ($vig2023 -and $vig2026 -and $antesDeTudo) `
            "2023->ano $($p2023[5]) ETp $($p2023[3]) | 2026->ano $($p2026[5]) ETp $($p2026[3]) | 2019->$r2019"

        # ---- 7.2 cada modelo deriva o que o ADR-022 manda ----
        $mEtp = [string]$xl.Run('MetasDaLinha', 'ETP_DIRETO', 30, '', '', '', '', '')
        $mVb = [string]$xl.Run('MetasDaLinha', 'VB', '', 4, 6, 'DES', '', '')
        $mCb = [string]$xl.Run('MetasDaLinha', 'CV_BIAS_DIRETO', '', '', '', '', 3, 2)
        $pE = $mEtp -split '\|'; $pV = $mVb -split '\|'; $pC = $mCb -split '\|'
        # ETP_DIRETO: CVtp = ETp/3, BIAStp indefinido
        $okE = ([Math]::Abs([double]$pE[0] - 10) -lt 0.0001 -and $pE[1] -eq '' -and [Math]::Abs([double]$pE[2] - 30) -lt 0.0001)
        # VB: CVtp = CVi*fi ; BIAStp = raiz(CVi^2+CVg^2)*fb ; ETp = BIAStp + 1,65*CVtp
        $cvVb = 4 * 0.5
        $biasVb = [Math]::Sqrt(16 + 36) * 0.25
        $etVb = $biasVb + 1.65 * $cvVb
        $okV = ([Math]::Abs([double]$pV[0] - $cvVb) -lt 0.0001 -and [Math]::Abs([double]$pV[1] - $biasVb) -lt 0.0001 -and [Math]::Abs([double]$pV[2] - $etVb) -lt 0.0001)
        # CV_BIAS_DIRETO: ETp = BIAStp + 1,65*CVtp
        $okC = ([Math]::Abs([double]$pC[2] - (2 + 1.65 * 3)) -lt 0.0001)
        Anotar '7.2' 'os tres modelos derivam CVtp/BIAStp/ETp conforme o ADR-022' `
            ($okE -and $okV -and $okC) `
            "ETP_DIRETO '$mEtp' | VB '$mVb' (esperado $cvVb/$([Math]::Round($biasVb,4))/$([Math]::Round($etVb,4))) | CV_BIAS '$mCb'"

        # ---- 7.3 quatro estados, e nenhuma ausencia vira aprovacao ----
        # O ponto do ADR-023: "sem meta" e "sem dado" tem donos opostos numa
        # auditoria, e fundir os dois esconde qual e. O que NAO pode acontecer,
        # em nenhum dos dois, e sair CONFORME.
        $semEspec = [string]$xl.Run('AvaliarConformidade', 'ZZ_INEXISTENTE_XYZ', 2026, 1, 1, 'CLIA')
        $semDados = [string]$xl.Run('AvaliarConformidade', $ESP_ANALITO, 2026, '', '', 'CLIA')
        # ETp 60 => CVtp 20. CV 5 e bias 1 passam; CV 40 estoura.
        $conf = [string]$xl.Run('AvaliarConformidade', $ESP_ANALITO, 2026, 5, 1, 'CLIA')
        $naoConf = [string]$xl.Run('AvaliarConformidade', $ESP_ANALITO, 2026, 40, 1, 'CLIA')
        $quatro = ($semEspec -eq 'SEM ESPECIFICACAO' -and $semDados -eq 'SEM DADOS' -and `
                   $conf -eq 'CONFORME' -and $naoConf -eq 'NAO CONFORME')
        $ausenciaNaoAprova = ($semEspec -ne 'CONFORME' -and $semDados -ne 'CONFORME')
        Anotar '7.3' 'conformidade tem quatro estados e ausencia nunca e aprovacao' `
            ($quatro -and $ausenciaNaoAprova) `
            "sem meta '$semEspec' | sem dado '$semDados' | dentro '$conf' | fora '$naoConf'"

        # ---- 7.4 especificacao desativada sai de vigencia, COM ACENTO ----
        # Regressao de defeito real: a comparacao era contra "NAO" cru, e o
        # gestor desativa digitando "Nao" com til -- que e como o proprio sistema
        # grava a Cfg_Status. UCase de "Nao" com til nao e "NAO", entao a linha
        # recem-desativada continuava valendo e seguia mudando veredito.
        # A prova usa DE PROPOSITO a grafia acentuada: com "NAO" cru ela passaria
        # mesmo com o defeito de volta.
        $naoAcentuado = [string][char]0x004E + [string][char]0x00E3 + 'o'   # "Nao" com til
        $wsEsp.Cells.Item($l2, 12).Value2 = $naoAcentuado
        $wsEsp.Cells.Item($l1, 12).Value2 = $naoAcentuado
        $xl.Run('InvalidarCacheEspec') | Out-Null
        $rDesativada = [string]$xl.Run('ResolverEspec', $ESP_ANALITO, 2026, 'CLIA')
        $sitDesativada = [string]$xl.Run('SituacaoEspec', $ESP_ANALITO, 2026, 'CLIA')
        $confDesativada = [string]$xl.Run('AvaliarConformidade', $ESP_ANALITO, 2026, 5, 1, 'CLIA')
        Anotar '7.4' 'especificacao desativada sai de vigencia (grafia acentuada)' `
            ($rDesativada -eq 'NAO_CADASTRADA' -and $confDesativada -eq 'SEM ESPECIFICACAO') `
            "desativada com '$naoAcentuado': resolve '$rDesativada', situacao '$sitDesativada', conformidade '$confDesativada'"

        # ---- 7.5 o protocolo do motor nao depende de localidade ----
        # Mesma familia do defeito que gravava "92,0028" e o Excel relia 920028
        # (item 2.2). Aqui o risco e CStr no protocolo interno: numa maquina
        # pt-BR a virgula viraria separador de campo do proprio retorno.
        $comDecimal = [string]$xl.Run('MetasDaLinha', 'CV_BIAS_DIRETO', '', '', '', '', 2.5, 1.25)
        $pD = $comDecimal -split '\|'
        $semVirgula = ($comDecimal -notmatch ',')
        $reversivel = ($pD.Count -eq 3 -and [Math]::Abs([double]$pD[2] - (1.25 + 1.65 * 2.5)) -lt 0.0001)
        Anotar '7.5' 'protocolo do motor de especificacoes e invariante de localidade' `
            ($semVirgula -and $reversivel) `
            "retorno '$comDecimal' (sem virgula: $semVirgula, reversivel: $reversivel)"

        # ---- limpeza: o banco de especificacoes volta ao que era ----
        $wsEsp.Range($wsEsp.Cells.Item($l1, 1), $wsEsp.Cells.Item($l2, 14)).ClearContents() | Out-Null
        $xl.Run('InvalidarCacheEspec') | Out-Null
        $linEspDepois = [int]$xl.Run('UltimaLinhaEspec')
        if ($linEspDepois -ne $linEspAntes) {
            "     AVISO: DB_Especificacoes ficou com $linEspDepois linhas (antes $linEspAntes)"
        }
        $wsEsp.Visible = $espVisAntes
    }

    $wb.VBProject.VBComponents.Remove($vba)
    $db.Cells.Item($linhaTeste, 7).Value2 = $statusOriginal

    $au.Visible = 2
    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect('qcini2025', $true, $false) }
    $wb.Save()
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    Encerrar-Excel
}

# ------------------------------------------------ 4. formulas e ADR-019 -------
""
"-- 4. FORMULAS E FRONTEIRA ARQUITETURAL ------------------------"
$csvBuild = Join-Path $env:TEMP 'verif_build.csv'
& (Join-Path $s 'snapshot_formulas.ps1') -Workbook $alvo -OutCsv $csvBuild | Out-Null
Encerrar-Excel

# Out-String -Stream: diff_formulas usa Format-Table, que emite OBJETOS DE
# FORMATACAO, nao texto. Sem converter, o log enche de
# "Microsoft.PowerShell.Commands.Internal.Format.FormatEntryData".
$csvDiv = Join-Path $env:TEMP 'verif_div.csv'
$diff = & (Join-Path $s 'diff_formulas.ps1') -Referencia (Join-Path $snap 'formulas.csv') -Candidato $csvBuild -OutCsv $csvDiv |
    Out-String -Stream
$diff | Where-Object { $_ -match '\S' } | Select-Object -First 12 | ForEach-Object { "     $_" }
$linhaValor = ($diff | Where-Object { $_ -match 'VALOR' }) -join ' '
$semValor = ($linhaValor -match 'VALOR\D*0\b')

# AUSENTE deixou de ser "tem de ser zero" e passou a ser LISTA FECHADA.
#
# A consolidacao da Sprint NC removeu de proposito regRep2 e regRep3: 1.080
# formulas, nas colunas Y,Z (nivel 1), AU,AV (nivel 2) e BQ,BR (nivel 3) do
# Calc. Exigir zero reprovaria uma mudanca aprovada; aceitar qualquer numero
# deixaria passar destruicao acidental -- que foi exatamente o incidente da
# Fase 3A. A checagem continua estrita: qualquer AUSENTE FORA dessas seis
# colunas reprova.
$COLS_NC = @('Y', 'Z', 'AU', 'AV', 'BQ', 'BR')
$ausentes = @()
if (Test-Path $csvDiv) {
    $ausentes = @(Import-Csv $csvDiv -Delimiter ';' | Where-Object { $_.Tipo -eq 'AUSENTE' })
}
$foraDaLista = @($ausentes | Where-Object {
        $aba = ($_.Chave -split '!')[0]
        $col = ($_.Chave -split '!')[1] -replace '\d', ''
        -not ($aba -eq 'Calc' -and $COLS_NC -contains $col)
    })
$okAusente = ($foraDaLista.Count -eq 0)
$detalhe = "{0} AUSENTE, {1} fora da lista fechada (regRep2/regRep3 do Calc)" -f $ausentes.Count, $foraDaLista.Count
if (-not $okAusente) {
    $detalhe += ' -> ' + (($foraDaLista | Select-Object -First 5 | ForEach-Object { $_.Chave }) -join ', ')
}
Anotar '4.1' 'AUSENTE so na lista fechada da Sprint NC' $okAusente $detalhe
Anotar '4.2' 'nenhuma formula virou VALOR' $semValor $linhaValor

Encerrar-Excel
# varredura_adr019 trabalha sobre o CSV de formulas ja extraido (-Formulas),
# nao sobre a pasta de trabalho: nao precisa reabrir o Excel.
$adr = & (Join-Path $s 'varredura_adr019.ps1') -Formulas $csvBuild | Out-String -Stream
$adr | Where-Object { $_ -match '\S' } | Select-Object -Last 6 | ForEach-Object { "     $_" }
$linhaBanco = ($adr | Where-Object { $_ -match 'sobre o banco|DB_Resultados' }) -join ' '
Anotar '4.3' 'nenhuma formula de interface calcula sobre o banco' ($adr -join ' ' -notmatch 'VIOLA') $linhaBanco
Encerrar-Excel


# ------------------------------------------------ 4b. desempenho --------------
""
"-- 4b. DESEMPENHO ---------------------------------------------"
"   (a gravacao em camada dupla nao pode tornar o sistema lento)"

Encerrar-Excel
$xlP = Novo-Excel
$xlP.Visible = $false; $xlP.DisplayAlerts = $false; $xlP.EnableEvents = $false
$xlP.AutomationSecurity = 1
$wbP = $xlP.Workbooks.Open($alvo)
    if ($wbP.ProtectStructure) { $wbP.Unprotect('qcini2025') }
try { $wbP.EnableAutoRecover = $false } catch { }
try {
    # motor completo
    $t0 = Get-Date
    $xlP.Run('AtualizarCalc') | Out-Null
    $xlP.Run('AtualizarPainelEng') | Out-Null
    $xlP.Run('AtualizarEstatisticaAba') | Out-Null
    $xlP.Run('RegistrarEventosWestgard') | Out-Null
    $tMotor = ((Get-Date) - $t0).TotalSeconds
    Anotar '4b.1' 'motor completo abaixo de 5s' ($tMotor -lt 5) ("{0:N2}s para as 4 rotinas sobre 1.575 registros" -f $tMotor)

    # 50 gravacoes nas DUAS camadas
    #
    # A PRIMEIRA GRAVACAO E DESCARTADA. Ela paga o custo de montar o cache de
    # sequencia (uma varredura do log inteiro) e nao representa o custo de
    # regime. Medir com ela dentro mistura duas coisas diferentes.
    $auP = $wbP.Worksheets.Item('Audit_Log')
    $auP.Visible = -1
    # '' e nao Empty: Empty e literal do VBA, nao do PowerShell -- vira palavra
    # solta e o analisador acusa "expressao ausente apos ','".
    $idAq = $xlP.Run('Auditar', 'DADO', 'AQUECIMENTO', 'suite', 0, (Get-Date), '', '', 0, '',
        '', '', '', '', 'Aquecimento do cache', '')
    $t0 = Get-Date
    for ($i = 1; $i -le 50; $i++) {
        $id = $xlP.Run('Auditar', 'DADO', 'TESTE_PERF', 'suite', $i, (Get-Date), '', 'QC-52261101', 1, 'WBC',
            3.0, 3.1, 'Ativo', 'Ativo', 'Medicao de desempenho', '')
        $xlP.Run('RegistrarLogDB', 'Resultados', $id, 'TESTE_PERF', $i, (Get-Date), 1, 'WBC',
            'QC-52261101', 3.1, 'Ativo', 'Ativo', 'Medicao de desempenho') | Out-Null
    }
    $tLog = ((Get-Date) - $t0).TotalSeconds
    $porEvento = $tLog / 50
    # LIMITE DE 250ms, e a folga e para RUIDO DE AMBIENTE, nao porque o codigo
    # piorou. Medicoes do mesmo codigo nesta maquina: 132ms com o Excel recem
    # iniciado, 187ms depois de horas de build. Um limite de 150ms cai no meio
    # dessa faixa e faz o teste falhar por acaso -- e teste instavel perde valor
    # rapido, porque as pessoas param de olhar o resultado.
    #
    # Cada evento inclui DUAS travessias Application.Run vindas do PowerShell,
    # que nao existem no uso real (o VBA chama direto). A medida e conservadora
    # de proposito: passando aqui, passa em producao.
    Anotar '4b.2' 'gravacao em camada dupla abaixo de 250ms por evento' ($porEvento -lt 0.25) `
        ("{0:N0}ms por evento ({1:N1}s para 50 eventos nas duas camadas)" -f ($porEvento * 1000), $tLog)

    # verificacao da cadeia com o log ja crescido
    $t0 = Get-Date
    $ver = [string]$xlP.Run('VerificarIntegridadeLog')
    $tVer = ((Get-Date) - $t0).TotalSeconds
    Anotar '4b.3' 'verificacao da cadeia integra e rapida' (($ver -like 'OK|*') -and $tVer -lt 10) `
        ("$ver em {0:N2}s" -f $tVer)
}
finally {
    try { $wbP.Close($false) } catch { }
    try { $xlP.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlP) | Out-Null
}
Encerrar-Excel

# ------------------- 2.2 o valor ORIGINAL fica recuperavel -------------------
# ISO 15189 8.4.2: apos emendar um resultado, o valor original tem de continuar
# recuperavel. Nao basta existir uma coluna "valor anterior" -- ela tem de
# conter o numero CERTO.
#
# Este controle existe porque em 04/08/2026 ela nao continha. Rotulo() devolvia
# String, e "92,0028" gravado numa celula Geral era relido pelo Excel como
# 920028 (virgula tratada como separador de milhar). O delta ficava correto,
# porque e gravado como numero -- e foi isso que escondeu o defeito.
# Pior: o hash da cadeia e calculado sobre o que se rele da celula, entao a
# verificacao de integridade APROVAVA o registro corrompido.
Encerrar-Excel
$xlV = Novo-Excel
$xlV.Visible = $false; $xlV.DisplayAlerts = $false; $xlV.EnableEvents = $false
$xlV.AutomationSecurity = 1
$wbV = $xlV.Workbooks.Open($alvo)
try { $wbV.EnableAutoRecover = $false } catch { }
try {
    $probe = @'
Option Explicit
Public Function ProvarValorOriginal() As String
    Dim d As Variant, regs() As Variant, valAnt As Variant, valNovo As Double
    Dim wsA As Worksheet, ult As Long, i As Long, lin As Long, gravado As Variant
    Dim e As Long, ds As String
    On Error GoTo f
    d = CarregarDB()
    valAnt = d(1, COL_RESULT)
    valNovo = CDbl(valAnt) + 7.77
    ReDim regs(1 To 1, 1 To 7)
    regs(1, COL_RUN) = CLng(d(1, COL_RUN)): regs(1, COL_DATA) = d(1, COL_DATA)
    regs(1, COL_NIVEL) = CLng(d(1, COL_NIVEL)): regs(1, COL_LOTE) = d(1, COL_LOTE)
    regs(1, COL_ANALITO) = CStr(d(1, COL_ANALITO)): regs(1, COL_RESULT) = valNovo
    regs(1, COL_STATUS) = ST_ATIVO
    UpsertResultados regs
    Set wsA = ThisWorkbook.Sheets(AUDIT)
    ult = UltimaLinhaAudit()
    For i = ult To AUDIT_R0 Step -1
        If InStr(1, UCase$(CStr(wsA.Cells(i, AU_ACAO).Value)), "ALTERAD") > 0 Then lin = i: Exit For
    Next i
    If lin = 0 Then ProvarValorOriginal = "SEM_EVENTO": Exit Function
    gravado = wsA.Cells(lin, AU_RESANT).Value
    If Not IsNumeric(gravado) Then
        ProvarValorOriginal = "NAO_NUMERICO|" & CStr(gravado)
    ElseIf Abs(CDbl(gravado) - CDbl(valAnt)) < 0.000001 Then
        ProvarValorOriginal = "OK|" & CStr(valAnt)
    Else
        ProvarValorOriginal = "DIVERGE|original=" & CStr(valAnt) & "|gravado=" & CStr(gravado)
    End If
    Exit Function
f:
    e = Err.Number: ds = Err.Description
    ProvarValorOriginal = "ERRO " & e & ": " & ds
End Function
'@
    $tmpP = Join-Path $env:TEMP 'mProva22.bas'
    Set-Content $tmpP $probe -Encoding Default
    foreach ($c in $wbV.VBProject.VBComponents) { if ($c.Name -eq 'mProva22') { $wbV.VBProject.VBComponents.Remove($c); break } }
    $wbV.VBProject.VBComponents.Import($tmpP) | Out-Null
    Remove-Item $tmpP -Force
    $r22 = [string]$xlV.Run("$($wbV.Name)!ProvarValorOriginal")
    Anotar '2.2' 'valor ORIGINAL recuperavel no Audit_Log' ($r22 -like 'OK|*') $r22
}
finally {
    try { $wbV.Close($false) } catch { }
    try { $xlV.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlV) | Out-Null
}
Encerrar-Excel

# --------------------------------- 1.6 nome publico duplicado ----------------
# Dois modulos com o mesmo nome publico impedem o projeto de COMPILAR, e o
# sintoma nao diz isso: Application.Run devolve "macro nao disponivel ou macros
# desabilitadas", que parece problema de confianca do Office.
#
# Aconteceu de verdade ao instalar o motor na Bioquimica: o mUI dela, anterior
# a Fase 3, definia um AtualizarEstatistica proprio -- um stub que so recalcula
# a aba -- e colidiu com o do mEstatistica. As 42 verificacoes existentes NAO
# pegaram, porque nenhuma chamava justamente esse procedimento.
Encerrar-Excel
$xlD = Novo-Excel
$xlD.Visible = $false; $xlD.DisplayAlerts = $false; $xlD.EnableEvents = $false
$xlD.AutomationSecurity = 1
$wbD = $xlD.Workbooks.Open($alvo)
try { $wbD.EnableAutoRecover = $false } catch { }
try {
    $publicos = @{}
    foreach ($comp in $wbD.VBProject.VBComponents) {
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -lt 1) { continue }
        foreach ($linha in ($cm.Lines(1, $cm.CountOfLines) -split "`r?`n")) {
            if ($linha -match '^\s*Public\s+(Sub|Function)\s+(\w+)') {
                $nome = $Matches[2]
                if (-not $publicos.ContainsKey($nome)) { $publicos[$nome] = @() }
                $publicos[$nome] += $comp.Name
            }
        }
    }
    $duplicados = @($publicos.Keys | Where-Object { ($publicos[$_] | Select-Object -Unique).Count -gt 1 })
    $detalheDup = if ($duplicados.Count) {
        ($duplicados | ForEach-Object { "$_ em $(($publicos[$_] | Select-Object -Unique) -join '+')" }) -join '; '
    }
    else { "$($publicos.Count) nomes publicos, nenhum repetido" }
    Anotar '1.6' 'nenhum nome publico duplicado entre modulos' ($duplicados.Count -eq 0) $detalheDup
}
finally {
    try { $wbD.Close($false) } catch { }
    try { $xlD.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlD) | Out-Null
}
Encerrar-Excel

# ------------------------------------------------- 4c. QA de hardening -------
# Itens 4.6 a 4.9 do Quality Gate. Sao os controles que pegam vazamento de
# estado -- a classe de defeito que nao aparece numa execucao unica e so se
# manifesta em producao, depois de o analista usar o sistema o dia inteiro.
""
"-- 4c. QA DE HARDENING (4.6 a 4.9) ----------------------------"

Encerrar-Excel
$xlQ = Novo-Excel
$xlQ.Visible = $false; $xlQ.DisplayAlerts = $false; $xlQ.EnableEvents = $false
$xlQ.AutomationSecurity = 1
$wbQ = $xlQ.Workbooks.Open($alvo)
try { $wbQ.EnableAutoRecover = $false } catch { }

try {
    # ---- 4.6 zero erro de formula ----
    # #N/A fica DE FORA da contagem, e nao por conveniencia: o Calc usa NA() de
    # proposito para abrir lacuna nas series do grafico -- ponto fora do filtro
    # ou sem dado. Contar #N/A como defeito reprovaria o comportamento correto.
    # Os outros seis erros nao tem uso legitimo aqui.
    $errosFormula = New-Object System.Collections.ArrayList
    foreach ($wsQ in $wbQ.Worksheets) {
        $urQ = $wsQ.UsedRange
        if ($urQ.Cells.Count -le 1) { continue }
        foreach ($tipoErro in @(-2146826281, -2146826273, -2146826265, -2146826259, -2146826252, -2146826288)) {
            # xlErrDiv0, xlErrValue, xlErrRef, xlErrName, xlErrNum, xlErrNull
            try {
                $achou = $urQ.Find($tipoErro, [Type]::Missing, -4163)   # xlValues
                if ($achou -ne $null) {
                    [void]$errosFormula.Add("$($wsQ.Name)!$($achou.Address($false,$false))")
                }
            }
            catch { }
        }
    }
    Anotar '4.6' 'nenhum erro de formula no artefato' ($errosFormula.Count -eq 0) `
    ("#N/A excluido de proposito (lacuna de serie do grafico); achados: " + $(if ($errosFormula.Count) { $errosFormula -join ', ' } else { 'nenhum' }))

    # ---- 4.7 estado global restaurado ----
    # Uma rotina que sai deixando ScreenUpdating desligado congela a tela do
    # analista; deixando Calculation manual, a planilha para de recalcular e o
    # painel mente sem dar erro.
    $xlQ.ScreenUpdating = $true
    $xlQ.EnableEvents = $true
    $xlQ.Calculation = -4105          # xlCalculationAutomatic
    $xlQ.Run("$($wbQ.Name)!AtualizarEstatistica") | Out-Null
    $estadoOk = ($xlQ.ScreenUpdating -eq $true) -and ($xlQ.Calculation -eq -4105)
    Anotar '4.7' 'estado global restaurado apos o motor' $estadoOk `
    ("ScreenUpdating=$($xlQ.ScreenUpdating) Calculation=$($xlQ.Calculation) (esperado True / -4105)")

    # ---- 4.8 idempotencia ----
    # Roda 1x, fotografa a saida, roda mais 10x e compara. Diferenca aqui
    # significa estado acumulando entre execucoes: colecao que nao zera,
    # contador que soma, linha que duplica.
    function Foto-Saida($wb) {
        $eng = $wb.Worksheets('Eng_Saida')
        $txt = New-Object System.Text.StringBuilder
        foreach ($rg in @('A3:AB182', 'A185:U187', 'A190:M309')) {
            $v = $eng.Range($rg).Value2
            if ($v -is [Array]) {
                for ($i = 1; $i -le $v.GetLength(0); $i++) {
                    for ($j = 1; $j -le $v.GetLength(1); $j++) { [void]$txt.Append([string]$v.GetValue($i, $j)).Append('|') }
                }
            }
        }
        return (Hash-Texto $txt.ToString())
    }
    $xlQ.Run("$($wbQ.Name)!AtualizarEstatistica") | Out-Null
    $h1 = Foto-Saida $wbQ
    for ($k = 1; $k -le 10; $k++) { $xlQ.Run("$($wbQ.Name)!AtualizarEstatistica") | Out-Null }
    $h10 = Foto-Saida $wbQ
    Anotar '4.8' 'idempotencia: 1x e 11x produzem saida identica' ($h1 -eq $h10) `
    ("1x $($h1.Substring(0,12)) / 11x $($h10.Substring(0,12))")

    # ---- 4.9 persistencia entre sessoes ----
    # Salva, FECHA o Excel de verdade e reabre noutro processo. Expoe dependencia
    # de variavel em memoria, cache Static e UserInterfaceOnly que nao persiste.
    $wbQ.Save()
}
finally {
    try { $wbQ.Close($true) } catch { }
    try { $xlQ.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlQ) | Out-Null
}
Encerrar-Excel

$xlQ2 = Novo-Excel
$xlQ2.Visible = $false; $xlQ2.DisplayAlerts = $false; $xlQ2.EnableEvents = $false
$xlQ2.AutomationSecurity = 1
$wbQ2 = $xlQ2.Workbooks.Open($alvo)
try { $wbQ2.EnableAutoRecover = $false } catch { }
try {
    $hAntes = Foto-Saida $wbQ2
    $xlQ2.Run("$($wbQ2.Name)!AtualizarEstatistica") | Out-Null
    $hDepois = Foto-Saida $wbQ2
    Anotar '4.9' 'persistencia: reabrir e recalcular nao muda a saida' ($hAntes -eq $hDepois) `
    ("reaberto $($hAntes.Substring(0,12)) / recalculado $($hDepois.Substring(0,12))")
}
finally {
    try { $wbQ2.Close($false) } catch { }
    try { $xlQ2.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlQ2) | Out-Null
}
Encerrar-Excel

# ------------------------------------------ 5. blindagem do entregavel --------
""
"-- 5. O ARQUIVO DISTRIBUIDO SAI TRAVADO ------------------------"
"   (cenario do auditor: abrir com macros DESABILITADAS)"

Encerrar-Excel
$dist = Join-Path $bd "QC_${Produto}_distribuicao.xlsm"
Copy-Item $alvoOriginal $dist -Force
& (Join-Path $s 'blindar_artefato.ps1') -Workbook $dist | ForEach-Object { "     $_" }
Encerrar-Excel

# Conferencia NO ARQUIVO, nao pelo Excel. Se a pergunta e "o que o auditor ve
# com macros desligadas", perguntar ao Excel com macros ligadas nao responde.
# O .xlsm e um zip: le-se o XML direto.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($dist)
try {
    function Ler-Entrada {
        param($Zip, [string]$Nome)
        $e = $Zip.Entries | Where-Object { $_.FullName -eq $Nome }
        if ($e -eq $null) { return '' }
        $sr = New-Object System.IO.StreamReader($e.Open())
        $t = $sr.ReadToEnd(); $sr.Close()
        return $t
    }

    $wbXml = Ler-Entrada $zip 'xl/workbook.xml'
    $totalAbas = ([regex]::Matches($wbXml, '<sheet ')).Count
    $veryHidden = ([regex]::Matches($wbXml, 'state="veryHidden"')).Count
    # 3 visiveis por desenho: Login, Audit_Log e Audit_Legenda. A trilha PRECISA
    # ser legivel com macros desabilitadas -- e assim que um auditor cauteloso
    # abre um arquivo desconhecido.
    Anotar '5.1' 'so Login e a trilha ficam visiveis no arquivo salvo' ($veryHidden -eq ($totalAbas - 3)) `
        "veryHidden $veryHidden de $totalAbas (visiveis: Login, Audit_Log, Audit_Legenda)"

    $wsXml = Ler-Entrada $zip 'xl/workbook.xml'
    $trilhaVisivel = ($wsXml -match 'name="Audit_Log"(?![^>]*veryHidden)')
    Anotar '5.5' 'trilha de auditoria legivel com macros desabilitadas' $trilhaVisivel `
        'Audit_Log nao esta oculta: e material de auditoria, feito para ser lido'

    Anotar '5.2' 'estrutura da pasta protegida' ($wbXml -match '<workbookProtection') ''

    # 5.6 de RUNTIME: o usuario nao pode criar aba. Testado no arquivo de
    # DISTRIBUICAO ($dist), que sai travado, numa instancia propria -- exige
    # que Worksheets.Add seja RECUSADO. E o comportamento que o gestor pediu;
    # o XML acima prova o estado, este prova o efeito.
    Encerrar-Excel
    $xlT = Novo-Excel; $xlT.Visible = $false; $xlT.DisplayAlerts = $false
    $xlT.EnableEvents = $false; $xlT.AutomationSecurity = 1
    $wbT = $xlT.Workbooks.Open($dist)
    $recusou = $false
    try { $wbT.Worksheets.Add() | Out-Null } catch { $recusou = $true }
    $wbT.Close($false); $xlT.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlT) | Out-Null
    Encerrar-Excel
    Anotar '5.6' 'usuario nao consegue criar aba (estrutura travada)' $recusou ``
        $(if ($recusou) { 'Worksheets.Add recusado pelo Excel' } else { 'FALHA: aba criada com estrutura travada' })

    $planilhas = $zip.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' }
    $comProt = 0
    foreach ($pl in $planilhas) {
        $t = Ler-Entrada $zip $pl.FullName
        if ($t -match '<sheetProtection') { $comProt++ }
    }
    Anotar '5.3' 'toda aba tem <sheetProtection> GRAVADA no arquivo' ($comProt -eq $planilhas.Count) `
        "$comProt de $($planilhas.Count) abas"

    # item 3.4: senha do projeto VBA. DPB vazio = sem senha.
    $vba = $zip.Entries | Where-Object { $_.FullName -eq 'xl/vbaProject.bin' }
    $temVba = ($vba -ne $null)
    $comSenha = $false
    if ($temVba) {
        $ms = New-Object System.IO.MemoryStream
        $vba.Open().CopyTo($ms)
        $bytes = $ms.ToArray(); $ms.Close()
        $txt = [System.Text.Encoding]::ASCII.GetString($bytes)
        # projeto sem senha grava DPB curto; com senha o blob passa de ~40 chars
        $m = [regex]::Match($txt, 'DPB="([0-9A-Fa-f]*)"')
        $comSenha = ($m.Success -and $m.Groups[1].Value.Length -gt 40)
    }
    Anotar '5.4' 'projeto VBA com senha (item 3.4)' $comSenha `
        $(if ($comSenha) { 'DPB preenchido' } else { 'PASSO MANUAL PENDENTE: VBE > Ferramentas > Propriedades do VBAProject > Protecao' })

    # ---------------- 6.2 / 6.3: o identificador chama-se RUN ----------------
    #
    # Le do ARQUIVO SALVO, nao do Excel: prova o que o usuario vai abrir.
    #
    # 6.2 nao e "a palavra corrida sumiu". A corrida e o EVENTO e o RUN e o
    # IDENTIFICADOR dele -- "Lancar corrida" e "DataCorrida" estao certos.
    # O que se verifica e que o identificador nao aparece mais como "Seq",
    # nome que nao existe desde a Fase 1.
    #
    # E verifica-se TAMBEM o inverso: que Audit_Legenda!B22 continua dizendo
    # "Seq", porque ali Seq e a sequencia de AUDITORIA (1a, 2a, 3a alteracao
    # do resultado), outro conceito. Um replace global de Seq->RUN passaria
    # em 6.2 e destruiria a explicacao da trilha para o auditor. Este item
    # falha nos DOIS sentidos, de proposito.
    $ssXml = Ler-Entrada $zip 'xl/sharedStrings.xml'
    $textos = @()
    foreach ($m in [regex]::Matches($ssXml, '<si>(.*?)</si>', 'Singleline')) {
        $textos += ($m.Groups[1].Value -replace '<[^>]+>', '')
    }
    # indices ainda REFERENCIADOS por alguma celula (o Excel deixa orfaos)
    $usados = @{}
    foreach ($pl in ($zip.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' })) {
        $t = Ler-Entrada $zip $pl.FullName
        foreach ($m in [regex]::Matches($t, '<c [^>]*t="s"[^>]*>\s*<v>(\d+)</v>')) {
            $usados[[int]$m.Groups[1].Value] = $true
        }
    }
    $vivos = @()
    foreach ($k in $usados.Keys) { if ($k -lt $textos.Count) { $vivos += $textos[$k] } }

    $seqIdent = @($vivos | Where-Object { $_ -eq 'Seq' -or $_ -match 'Seq\s*=\s*n' })
    $temRunDesc = @($vivos | Where-Object { $_ -match 'RUN = identificador da corrida' }).Count -gt 0
    $legendaIntacta = @($vivos | Where-Object { $_ -match 'ORIGINAL.*Seq = 1' }).Count -gt 0

    Anotar '6.2' 'identificador da corrida rotulado RUN (e Seq de auditoria preservado)' `
        (($seqIdent.Count -eq 0) -and $temRunDesc -and $legendaIntacta) `
        $(if ($seqIdent.Count -gt 0) { "ainda ha rotulo Seq como identificador: $($seqIdent -join ' / ')" }
          elseif (-not $temRunDesc) { 'FALHA: DB_Resultados!A2 nao descreve o RUN' }
          elseif (-not $legendaIntacta) { 'FALHA: Audit_Legenda perdeu o Seq de AUDITORIA (replace global?)' }
          else { 'nenhum Seq como identificador; Audit_Legenda!B22 preservada' })

    $graficos = @($zip.Entries | Where-Object { $_.FullName -match '^xl/charts/chart\d+\.xml$' })
    $eixosOk = 0; $eixosRuins = @()
    foreach ($g in $graficos) {
        $t = Ler-Entrada $zip $g.FullName
        # O eixo X identifica-se por axPos="b" (bottom), NAO pelo tipo do
        # elemento. Estes graficos sao de DISPERSAO: o RUN e numerico, entao o
        # eixo X e um <c:valAx>, nao um <c:catAx>. Procurar so em catAx nao
        # acha nada e reprova um arquivo correto -- foi o que aconteceu na
        # primeira versao deste item.
        $tit = ''
        foreach ($mc in [regex]::Matches($t, '<c:(catAx|valAx|dateAx)>.*?</c:\1>', 'Singleline')) {
            if ([regex]::Match($mc.Value, '<c:axPos val="b"').Success) {
                $mt = [regex]::Match($mc.Value, '<c:title>.*?</c:title>', 'Singleline')
                if ($mt.Success) { $tit = (($mt.Value -replace '<[^>]+>', '') -replace '\s+', ' ').Trim() }
                break
            }
        }
        if ($tit -eq 'RUN') { $eixosOk++ } else { $eixosRuins += "$($g.FullName): '$tit'" }
    }
    Anotar '6.3' 'eixo X de todo grafico rotulado RUN' `
        (($graficos.Count -gt 0) -and ($eixosRuins.Count -eq 0)) `
        $(if ($graficos.Count -eq 0) { 'FALHA: nenhum grafico no arquivo' }
          elseif ($eixosRuins.Count -gt 0) { "fora do padrao: $($eixosRuins -join ' / ')" }
          else { "$eixosOk de $($graficos.Count) graficos com eixo X = RUN" })
}
finally { $zip.Dispose() }

# ------------------------------------------------------- 6. relatorio ---------
$falhas = @($itens | Where-Object { -not $_.Ok })
$md = New-Object System.Collections.ArrayList
[void]$md.Add("# Verificacao do QC_INI - $Produto")
[void]$md.Add("")
[void]$md.Add("**Executada em:** $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')")
[void]$md.Add("**Artefato:** ``$alvo``")
[void]$md.Add("")
[void]$md.Add("| Item | Verificacao | Resultado | Evidencia |")
[void]$md.Add("|---|---|---|---|")
foreach ($i in $itens) {
    $r = if ($i.Ok) { 'OK' } else { '**FALHA**' }
    [void]$md.Add("| $($i.Id) | $($i.Nome) | $r | $($i.Evidencia) |")
}
[void]$md.Add("")
[void]$md.Add("**$($itens.Count - $falhas.Count) de $($itens.Count) verificacoes passaram.**")
if ($falhas.Count -gt 0) {
    [void]$md.Add("")
    [void]$md.Add("## Falhas")
    foreach ($f in $falhas) { [void]$md.Add("- **$($f.Id)** $($f.Nome) -- $($f.Evidencia)") }
}
[System.IO.File]::WriteAllLines($relatorio, $md.ToArray(), (New-Object System.Text.UTF8Encoding($false)))

""
"================================================================"
"  $($itens.Count - $falhas.Count) de $($itens.Count) verificacoes passaram"
if ($falhas.Count -gt 0) {
    "  FALHAS:"
    foreach ($f in $falhas) { "    - $($f.Id) $($f.Nome)" }
}
"  Relatorio: $relatorio"
"================================================================"

if ($falhas.Count -gt 0) { exit 1 }
exit 0
