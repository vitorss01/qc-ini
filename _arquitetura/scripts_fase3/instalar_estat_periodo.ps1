# instalar_estat_periodo.ps1 - poe mEstatPeriodo na producao e CONFERE
#
# CONFERE, NAO CONFIA: o modulo so e dado por instalado depois de RESPONDER com
# o mesmo numero que ja estava na tela. n, media, DP e CV do periodo anual de
# 2026 sao lidos ANTES do import e conferidos DEPOIS -- se o modulo novo
# divergir, e ele que esta errado, e a instalacao e recusada.
#
# Na migracao, a referencia eram as formulas antigas das colunas C..F. Hoje
# essas colunas ja chamam EstatPeriodo, entao compara-las com o modulo seria
# circular; a referencia passou a ser o estado anterior ao import.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\instalar_estat_periodo.ps1 -Workbook ..\..\QC_Bioquimica.xlsm

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Fonte = ''
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if ($Fonte -eq '') {
    $Fonte = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'src_producao\mEstatPeriodo.bas'
}
$Fonte = (Resolve-Path -LiteralPath $Fonte).ProviderPath

function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $u = $_
            if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }
            Start-Sleep -Seconds ($t * 2)
        }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { try { $wb.Close($false) } catch { }; $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvar = $false
try {
    # ---------------- linha de base: o ANTES, capturado antes do import -------
    #
    # Tem de ser lido AQUI, com o modulo que esta rodando hoje. Depois do
    # import nao existe mais "antes" para comparar -- e foi por isso que a
    # conferencia virou circular sem ninguem notar.
    $estB = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -like 'Estat*') { $estB = $w; break } }
    if ($estB -eq $null) { throw 'aba Estatistica ausente' }
    $cabB = 0
    for ($r = 1; $r -le 40; $r++) {
        if ("$($estB.Cells($r,1).Value2)".Trim() -eq 'Analito') { $cabB = $r; break }
    }
    if ($cabB -eq 0) { throw 'cabecalho Analito nao encontrado na Estatistica' }
    # A TABELA TERMINA NA PRIMEIRA LINHA VAZIA.
    #
    # Usar a ultima linha usada da coluna A varria para dentro do bloco de
    # margem critica (ADR-033), que tambem tem nome de analito na coluna A --
    # so que ali a coluna D e texto ("Margem critica"), e o cast estourava.
    # 87 pares onde ha 31 analitos x 2 niveis ja denunciava a invasao.
    $fimB = $cabB
    while ("$($estB.Cells($fimB + 1, 1).Value2)".Trim() -ne '') { $fimB++ }
    $base = @{}
    for ($r = ($cabB + 1); $r -le $fimB; $r++) {
        $aB = "$($estB.Cells($r,1).Value2)".Trim()
        if ($aB -eq '') { continue }
        $nvB = "$($estB.Cells($r,2).Value2)".Trim()
        $base["$aB|$nvB"] = @($estB.Cells($r,3).Value2, $estB.Cells($r,4).Value2,
                              $estB.Cells($r,5).Value2, $estB.Cells($r,6).Value2)
    }
    "linha de base: $($base.Count) pares analito|nivel lidos antes do import"

    $proj = $wb.VBProject
    $alvo = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'mEstatPeriodo') { $alvo = $c; break } }
    if ($alvo -ne $null) { $proj.VBComponents.Remove($alvo); "removido: mEstatPeriodo (versao anterior)" }

    $proj.VBComponents.Import($Fonte) | Out-Null
    $novo = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'mEstatPeriodo') { $novo = $c; break } }
    if ($novo -eq $null) { throw 'Import nao criou mEstatPeriodo' }
    "importado: mEstatPeriodo ($($novo.CodeModule.CountOfLines) linhas)"

    # ---------------- conferencia contra o estado anterior ----------------
    $est = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -like 'Estat*') { $est = $w; break } }
    if ($est -eq $null) { throw 'aba Estatistica ausente' }

    $ini = [double](Get-Date -Year 2026 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).ToOADate()
    $fim = [double](Get-Date -Year 2026 -Month 12 -Day 31 -Hour 0 -Minute 0 -Second 0).ToOADate()

    # A JANELA DE LINHAS VEM DO CONTEUDO, NAO DE UM NUMERO ESCRITO A MAO.
    #
    # Era "de 7 a 86". A reestruturacao do ADR-033 desceu o cabecalho para a
    # linha 7, entao a primeira leitura pegava o TEXTO do cabecalho e
    # [double]"n" estourava, derrubando a conferencia INTEIRA -- depois de o
    # modulo ja ter sido importado. Fail-closed salvou o arquivo, mas o portao
    # nunca chegou a medir nada.
    #
    # E a mesma deriva de coordenada que ja quebrou os testes tres vezes neste
    # projeto, e a cura e a mesma: achar o cabecalho pelo rotulo.
    $linCab = 0
    for ($r = 1; $r -le 40; $r++) {
        if ("$($est.Cells($r,1).Value2)".Trim() -eq 'Analito') { $linCab = $r; break }
    }
    if ($linCab -eq 0) { throw 'cabecalho Analito nao encontrado na Estatistica' }
    $linFim = $linCab
    while ("$($est.Cells($linFim + 1, 1).Value2)".Trim() -ne '') { $linFim++ }
    "  conferindo linhas $($linCab + 1) a $linFim (cabecalho na $linCab)"

    $erros = @()
    $conferidas = 0
    for ($r = ($linCab + 1); $r -le $linFim; $r++) {
        $a = "$($est.Cells($r,1).Value2)".Trim()
        if ($a -eq '') { continue }
        $niv = "$($est.Cells($r,2).Value2)".Trim()
        $nAnt = $est.Cells($r, 3).Value2
        # -as devolve $null quando nao ha numero ali: rotulo remanescente ou
        # linha sem medida nao sao caso de conferencia, e nao podem derrubar o
        # portao.
        $nNum = $nAnt -as [double]
        if ($nNum -eq $null -or $nNum -lt 2) { continue }

        # ASSINATURA DE 7 PARAMETROS, NAO 8.
        #
        # As exclusoes eram (inicio, fim) -- dois argumentos -- e passaram a ser
        # UM intervalo Nx2. O instalador ficou com a chamada antiga e o COM
        # devolvia "o numero de parametros nao coincide", DEPOIS do import.
        $nNovo = $xl.Run('EstatPeriodo', $a, $niv, 'N', $ini, $fim, '', '')
        $mNovo = $xl.Run('EstatPeriodo', $a, $niv, 'MEDIA', $ini, $fim, '', '')
        $dNovo = $xl.Run('EstatPeriodo', $a, $niv, 'DP', $ini, $fim, '', '')
        $cNovo = $xl.Run('EstatPeriodo', $a, $niv, 'CV', $ini, $fim, '', '')

        # A REFERENCIA E O ANTES, NAO A COLUNA.
        #
        # O cabecalho desta secao dizia "conferencia contra o caminho antigo", e
        # era verdade na migracao: as colunas C..F guardavam as formulas velhas.
        # Hoje C14 e =EstatPeriodo(...) -- comparar o modulo com uma celula que
        # chama o modulo nao prova nada, so parece provar.
        #
        # O que ainda da para provar honestamente e REGRESSAO: os valores lidos
        # ANTES do import, com o modulo que estava rodando, contra os de agora.
        $ref = $base["$a|$niv"]
        if ($ref -eq $null) { continue }
        $rN = $ref[0] -as [double]; $rM = $ref[1] -as [double]
        $rD = $ref[2] -as [double]; $rC = $ref[3] -as [double]
        if ($rN -eq $null -or $rM -eq $null -or $rD -eq $null -or $rC -eq $null) { continue }
        if (([double]$nNovo) -ne $rN) { $erros += "$a N$niv : n antes $($ref[0]), agora $nNovo" }
        if ([Math]::Abs([double]$mNovo - $rM) -gt 0.0001) { $erros += "$a N$niv : media antes $($ref[1]), agora $mNovo" }
        if ([Math]::Abs([double]$dNovo - $rD) -gt 0.0001) { $erros += "$a N$niv : DP antes $($ref[2]), agora $dNovo" }
        if ([Math]::Abs([double]$cNovo - $rC) -gt 0.0001) { $erros += "$a N$niv : CV antes $($ref[3]), agora $cNovo" }
        $conferidas++
    }

    # janela de exclusao tem de REDUZIR o n
    $exI = [double](Get-Date -Year 2026 -Month 1 -Day 1).ToOADate()
    $exF = [double](Get-Date -Year 2026 -Month 1 -Day 9).ToOADate()
    # LerExclusoes espera um bloco Nx2: inicio na coluna 1, fim na coluna 2.
    $exUm = New-Object 'object[,]' 1, 2
    $exUm[0, 0] = $exI; $exUm[0, 1] = $exF
    $exTudo = New-Object 'object[,]' 1, 2
    $exTudo[0, 0] = $ini; $exTudo[0, 1] = $fim

    $nCheio = $xl.Run('EstatPeriodo', 'Lactato', 1, 'N', $ini, $fim, '', '')
    $nCorte = $xl.Run('EstatPeriodo', 'Lactato', 1, 'N', $ini, $fim, $exUm, '')
    if ([double]$nCorte -ge [double]$nCheio) { $erros += "exclusao nao reduziu o n: cheio $nCheio, cortado $nCorte" }

    # exclusao cobrindo tudo tem de zerar
    $nZero = $xl.Run('EstatPeriodo', 'Lactato', 1, 'N', $ini, $fim, $exTudo, '')
    if ([double]$nZero -ne 0) { $erros += "exclusao total deveria zerar o n, deu $nZero" }

    # status: ausencia de limite NAO pode virar aprovacao
    $sSemLim = [string]$xl.Run('StatusCV', 2.5, '', '', '')
    $sSoCLIA = [string]$xl.Run('StatusCV', 2.5, 5.67, '', '')
    $sFora = [string]$xl.Run('StatusCV', 9.9, 5.67, '', '')
    $eSemLim = [string]$xl.Run('StatusETP', 10, '', '')
    $eAmbos = [string]$xl.Run('StatusETP', 30, 17, 12)
    if ($sSemLim -ne 'SEM LIMITE') { $erros += "StatusCV sem limite devolveu '$sSemLim'" }
    if ($sSoCLIA -ne 'OK (CLIA)') { $erros += "StatusCV so CLIA devolveu '$sSoCLIA'" }
    if ($sFora -notlike 'FORA*') { $erros += "StatusCV fora devolveu '$sFora'" }
    if ($eSemLim -ne 'SEM LIMITE') { $erros += "StatusETP sem limite devolveu '$eSemLim'" }
    if ($eAmbos -ne 'CRITICO: FORA AMBOS') { $erros += "StatusETP fora ambos devolveu '$eAmbos'" }

    # LimEspec: especificacao ausente tem de virar "", NUNCA 0.
    # Com 0 o StatusCV leria "limite zero" e reprovaria os 31 analitos em VB.
    #
    # O NOME DO ANALITO E MONTADO POR CODIGO DE CARACTERE, e nao escrito como
    # literal. Windows PowerShell 5.1 le o .ps1 como ANSI: um "Acido urico"
    # acentuado digitado aqui chega corrompido ao Excel, nao casa com nenhuma
    # linha do banco e LimEspec devolve vazio -- fazendo a conferencia reprovar
    # um modulo correto. Ja aconteceu nesta sessao.
    $A_ = [string][char]0x00C1      # A agudo maiusculo
    $u_ = [string][char]0x00FA      # u agudo
    $anTeste = $A_ + 'cido ' + $u_ + 'rico'
    $limCLIA = $xl.Run('LimEspec', $anTeste, 2026, 'CLIA', 'ETP')
    $limVB = $xl.Run('LimEspec', $anTeste, 2026, 'Variacao Biologica', 'CV')
    $limNada = $xl.Run('LimEspec', 'ZZ_NAO_EXISTE', 2026, 'CLIA', 'ETP')
    # O QUE SE EXIGE DEPENDE DE O MOTOR DE ESPECIFICACOES EXISTIR.
    #
    # LimEspec le EspecCVtp/EspecBIAStp/EspecETp, que vivem no mEspecificacoes.
    # Esse modulo HOJE nao existe em nenhum dos dois produtos -- so a
    # Hematologia tem etapa de build que o gera. Exigir "17" de um produto que
    # nao tem o banco ligado reprova um modulo correto, e foi o que aconteceu.
    #
    # A regra que este bloco existe para proteger nao e o numero 17: e o
    # ADR-023 -- LIMITE AUSENTE NAO E APROVACAO, e nunca vira zero. Zero faria
    # StatusCV ler "limite zero" e reprovar os 31 analitos. Essa exigencia vale
    # nos dois cenarios, e e ela que fica obrigatoria.
    $temEspec = $false
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'mEspecificacoes') { $temEspec = $true; break } }

    if ($temEspec) {
        if ("$limCLIA" -eq '' -or [double]$limCLIA -le 0) { $erros += "LimEspec CLIA/ETp de Acido urico devolveu '$limCLIA', esperado 17" }
    } else {
        "  mEspecificacoes ausente: exige-se degradacao para vazio, nao o limite"
        if ("$limCLIA" -ne '') { $erros += "sem mEspecificacoes, LimEspec devolveu '$limCLIA'; esperado vazio" }
        if ("$limCLIA" -eq '0') { $erros += "LimEspec devolveu ZERO: ausencia virando limite zero (ADR-023)" }
    }
    if ("$limVB" -ne '') { $erros += "LimEspec VB (nao cadastrada) devolveu '$limVB', esperado vazio" }
    if ("$limNada" -ne '') { $erros += "LimEspec de analito inexistente devolveu '$limNada', esperado vazio" }

    # e o encadeamento que importa: limite ausente nao pode reprovar NEM aprovar
    $sVazio = [string]$xl.Run('StatusCV', 2.39, $limCLIA, $limVB, '')
    $espSt = $(if ($temEspec) { 'OK (CLIA)' } else { 'SEM LIMITE' })
    if ($sVazio -ne $espSt) { $erros += "StatusCV com VB ausente devolveu '$sVazio', esperado '$espSt'" }

    # janela derivada
    $jIni = $xl.Run('JanelaInicio', 'TRIMESTRAL', 2026, 1, '', '')
    $jFim = $xl.Run('JanelaFim', 'TRIMESTRAL', 2026, 1, '', '')
    $espIni = (Get-Date -Year 2026 -Month 1 -Day 1).Date
    $espFim = (Get-Date -Year 2026 -Month 3 -Day 31).Date
    if ([datetime]$jIni -ne $espIni) { $erros += "JanelaInicio TRIMESTRAL/1 deu $jIni" }
    if ([datetime]$jFim -ne $espFim) { $erros += "JanelaFim TRIMESTRAL/1 deu $jFim" }

    # CONFERIR ZERO LINHA NAO E APROVAR.
    #
    # Sem este piso, uma mudanca de layout que fizesse o laco nao casar nenhuma
    # linha passaria como sucesso -- o portao diria OK tendo medido nada. E a
    # mesma cegueira que o ADR-046 encontrou na auditoria estatica.
    $minimo = 20
    if ($conferidas -lt $minimo) {
        # ${minimo} entre chaves: "$minimo:" faz o PowerShell ler o dois-pontos
        # como qualificador de drive e o arquivo nem chega a interpretar.
        $erros += "conferidas apenas $conferidas linha(s), minimo ${minimo} -- o portao nao mediu o suficiente para aprovar"
    }

    if ($erros.Count -gt 0) {
        $erros | Select-Object -First 12 | ForEach-Object { "  FALHA: $_" }
        throw "Instalacao recusada: $($erros.Count) conferencia(s) falharam. Nada foi salvo."
    }

    "conferencia: $conferidas linha(s) reproduzem o estado anterior ao import (n, media, DP, CV)"
    "conferencia: exclusao 01/01-09/01 levou n de $nCheio para $nCorte; exclusao total zerou"
    "conferencia: StatusCV/StatusETP nao aprovam sem limite; TRIMESTRAL/1 = 01/01 a 31/03"

    $salvar = $true
    $wb.Save()
    "Salvo: $Workbook"
}
finally {
    try { if ($salvar) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
}
