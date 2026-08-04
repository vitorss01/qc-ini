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
    [switch]$PularBuild
)

$ErrorActionPreference = 'Stop'

$s = Split-Path -Parent $MyInvocation.MyCommand.Path
$arq = Split-Path -Parent $s
$h = Join-Path $arq 'src_hardening1'
$snap = Join-Path $arq 'snapshot_producao\Hematologia'
$bd = Join-Path $env:USERPROFILE 'QCINI_build_hardening1'
$alvo = Join-Path $bd 'QC_Hematologia.xlsm'
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
"  VERIFICACAO DO QC_INI - Hematologia"
"  $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
"================================================================"

# ---------------------------------------------------------------- 0. build ----
if (-not $PularBuild) {
    ""
    "-- 0. BUILD ----------------------------------------------------"
    $saidaBuild = & (Join-Path $s 'build_all.ps1') 2>&1
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
$copia = Join-Path $bd "QC_Hematologia_verificacao.xlsm"
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

try {
    $modulos = @(
        @{ Nome = 'mEstatistica'; Arquivo = (Join-Path $h 'mEstatistica.bas') },
        @{ Nome = 'mDados'; Arquivo = (Join-Path $h 'mDados.bas') },
        @{ Nome = 'mAuditoria'; Arquivo = (Join-Path $h 'mAuditoria.bas') }
    )
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
    foreach ($n in @('Eng_Saida', 'Corridas', 'Audit_Log')) {
        Anotar '1.3' "aba $n existe" ($abas -contains $n) ''
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

    $wb.VBProject.VBComponents.Remove($vba)
    $db.Cells.Item($linhaTeste, 7).Value2 = $statusOriginal

    $au.Visible = 2
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
$diff = & (Join-Path $s 'diff_formulas.ps1') -Referencia (Join-Path $snap 'formulas.csv') -Candidato $csvBuild |
    Out-String -Stream
$diff | Where-Object { $_ -match '\S' } | Select-Object -First 12 | ForEach-Object { "     $_" }
$linhaAusente = ($diff | Where-Object { $_ -match 'AUSENTE' }) -join ' '
$linhaValor = ($diff | Where-Object { $_ -match 'VALOR' }) -join ' '
$semAusente = ($linhaAusente -match 'AUSENTE\D*0\b')
$semValor = ($linhaValor -match 'VALOR\D*0\b')
Anotar '4.1' 'nenhuma formula AUSENTE' $semAusente $linhaAusente
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

# ------------------------------------------ 5. blindagem do entregavel --------
""
"-- 5. O ARQUIVO DISTRIBUIDO SAI TRAVADO ------------------------"
"   (cenario do auditor: abrir com macros DESABILITADAS)"

Encerrar-Excel
$dist = Join-Path $bd 'QC_Hematologia_distribuicao.xlsm'
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
}
finally { $zip.Dispose() }

# ------------------------------------------------------- 6. relatorio ---------
$falhas = @($itens | Where-Object { -not $_.Ok })
$md = New-Object System.Collections.ArrayList
[void]$md.Add("# Verificacao do QC_INI - Hematologia")
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
