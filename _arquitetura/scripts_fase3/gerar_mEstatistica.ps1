# gerar_mEstatistica.ps1 - monta a versao do hardening de mEstatistica.bas
#
# Marco 2: substitui AtualizarCalc (linhas 299..512 da versao de producao) pela
# versao que publica em Eng_Saida, e acrescenta as constantes de layout EF0/NEF.
#
# A saida e determinista: mesma entrada, mesmo arquivo. Sem edicao manual dentro
# do editor do VBA -- o modulo e fonte versionada, o .xlsm e artefato.
#
# ATENCAO codificacao: modulos VBA sao ANSI (cp1252). Ler e escrever com
# [System.Text.Encoding]::Default preserva os acentos, inclusive em
# Sheets("Estatistica") com acento. Gravar como UTF-8 corromperia o modulo.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).

param(
    [Parameter(Mandatory = $true)][string]$Producao,     # mEstatistica.bas de producao
    [Parameter(Mandatory = $true)][string]$NovoSub,      # AtualizarCalc.txt
    [Parameter(Mandatory = $true)][string]$Saida,
    [string]$NovoPainel,                                 # AtualizarPainelEng.txt (Marco 3)
    [string]$NovoEstat,                                  # AtualizarEstatisticaAba.txt (Marco 4)
    # Niveis do produto alvo. O modulo de PRODUCAO usado como base e sempre o da
    # Hematologia (NLV=3) -- e o unico que existe. Para Bioquimica e Imunologia,
    # que tem 2 niveis, a constante e reescrita aqui.
    #
    # Isto NAO e um ajuste cosmetico: NLV governa quantos blocos de coluna o
    # motor escreve no Calc, quantas linhas de nivel publica no Painel e quantas
    # linhas ocupa na Estatistica. Se divergir da geometria da planilha, o motor
    # escreve num lugar e a interface le outro, sem erro visivel.
    [int]$NLV = 3
)

# Duas codificacoes de proposito, e trocar uma pela outra corrompe o modulo:
#   $enc      cp1252  - modulos VBA exportados/importados pelo Excel
#   $encFonte UTF-8   - os .txt de origem, escritos por editor moderno
# Ler um .txt UTF-8 como cp1252 transforma Sheets("Estatistica" com acento) em
# Sheets("EstatA-stica") e o procedimento falha em tempo de execucao.
$enc = [System.Text.Encoding]::Default
$encFonte = New-Object System.Text.UTF8Encoding($false)

$L = [System.IO.File]::ReadAllLines($Producao, $enc)

# --- fronteiras por MARCADOR, nao por indice -------------------------------
#
# Indices fixos (299, 512, 679, 750, 799) quebram a cada linha acrescentada no
# motor. A reescrita de AvaliarWestgard pelo ADR-041 deslocou tudo e o build
# parou com "Fronteira inicial inesperada na linha 299". O guarda-corpo estava
# certo; o metodo de localizar e que era fragil. Marcador nao desloca.
function Achar([object]$linhas, [string]$padrao, [int]$de = 0) {
    for ($i = $de; $i -lt $linhas.Count; $i++) {
        if ($linhas[$i] -like $padrao) { return $i }
    }
    return -1
}

function AcharEndSub([object]$linhas, [int]$de) {
    for ($i = $de; $i -lt $linhas.Count; $i++) {
        if ($linhas[$i].Trim() -eq 'End Sub') { return $i }
    }
    return -1
}

$iCte = Achar $L '*NFD*'
if ($iCte -lt 0) { throw 'Nao achei a linha de constantes (NFD).' }

$iCalcIni = Achar $L '*MOTOR: MONTAR Calc*'
if ($iCalcIni -lt 0) { throw 'Nao achei o marcador "MOTOR: MONTAR Calc".' }
$iCalcSub = Achar $L '*Sub AtualizarCalc*' $iCalcIni
if ($iCalcSub -lt 0) { throw 'Nao achei Public Sub AtualizarCalc.' }
$iCalcFim = AcharEndSub $L $iCalcSub
if ($iCalcFim -lt 0) { throw 'Nao achei o End Sub de AtualizarCalc.' }

$iPainelIni = Achar $L '*Sub AtualizarPainelEng*'
$iPainelFim = -1
if ($iPainelIni -ge 0) { $iPainelFim = AcharEndSub $L $iPainelIni }

$iEstatIni = Achar $L '*Sub AtualizarEstatisticaAba*'
$iEstatFim = -1
if ($iEstatIni -ge 0) { $iEstatFim = AcharEndSub $L $iEstatIni }

"fronteiras: constantes=$($iCte + 1) Calc=$($iCalcIni + 1)..$($iCalcFim + 1) " +
"Painel=$($iPainelIni + 1)..$($iPainelFim + 1) Estat=$($iEstatIni + 1)..$($iEstatFim + 1)"

$novo = [System.IO.File]::ReadAllLines($NovoSub, $encFonte)

$out = New-Object System.Collections.ArrayList

# constantes + as sete novas
for ($i = 0; $i -le $iCte; $i++) { [void]$out.Add($L[$i]) }
[void]$out.Add("Private Const EF0 As Long = 3          ' Eng_Saida: 1a coluna de bloco de nivel")
[void]$out.Add("Private Const NEF As Long = 7          ' Eng_Saida: campos por nivel")
[void]$out.Add("Private Const COL_FILTRO As Long = 24  ' Eng_Saida: filtro de data por corrida")
[void]$out.Add("Private Const COL_VALOR0 As Long = 25  ' Eng_Saida: 1a coluna de valor por nivel")
[void]$out.Add("Private Const COL_CHAVE As Long = 28   ' Eng_Saida: chave logica ANALITO|RUN")
[void]$out.Add("Private Const LINHA_STAT As Long = 185 ' Eng_Saida: 1a linha do bloco de estatistica")
[void]$out.Add("Private Const LINHA_EST As Long = 190  ' Eng_Saida: 1a linha da tabela de parametros")

# ate o inicio de AtualizarCalc, inalteradas
for ($i = $iCte + 1; $i -lt $iCalcIni; $i++) { [void]$out.Add($L[$i]) }

# AtualizarCalc substituida
foreach ($linha in $novo) { [void]$out.Add($linha) }
$prox = $iCalcFim + 1

# AtualizarPainelEng, se fornecida
if ($NovoPainel -and $iPainelIni -ge 0) {
    for ($i = $prox; $i -lt $iPainelIni; $i++) { [void]$out.Add($L[$i]) }
    foreach ($linha in [System.IO.File]::ReadAllLines($NovoPainel, $encFonte)) { [void]$out.Add($linha) }
    $prox = $iPainelFim + 1
}

# AtualizarEstatisticaAba, se fornecida
if ($NovoEstat -and $iEstatIni -ge 0) {
    for ($i = $prox; $i -lt $iEstatIni; $i++) { [void]$out.Add($L[$i]) }
    foreach ($linha in [System.IO.File]::ReadAllLines($NovoEstat, $encFonte)) { [void]$out.Add($linha) }
    $prox = $iEstatFim + 1
}

# resto inalterado
for ($i = $prox; $i -lt $L.Count; $i++) { [void]$out.Add($L[$i]) }

# --- niveis do produto ---
$linhaNLV = -1
for ($i = 0; $i -lt $out.Count; $i++) {
    if ($out[$i] -match '^\s*Public Const NLV\s+As\s+Long\s*=\s*(\d+)') {
        $linhaNLV = $i
        $nlvOrigem = [int]$Matches[1]
        break
    }
}
if ($linhaNLV -lt 0) { throw 'Constante Public Const NLV nao encontrada no modulo.' }
if ($nlvOrigem -ne $NLV) {
    $out[$linhaNLV] = "Public Const NLV As Long = $NLV          ' niveis deste setor"
    "NLV reescrito: $nlvOrigem -> $NLV (linha $($linhaNLV + 1))"
}
else {
    "NLV: $NLV (sem alteracao)"
}

# --- lote por semantica, nao por largura fixa ---
# Hematologia usa nucleo de seis digitos, mas Bioquimica possui lotes reais de
# quatro digitos. Mid$(...,4,6) incorporava o sufixo do nivel ao nucleo e fazia
# o motor publicar nRun=0. NucleoLote e a API comum de mDados para os produtos.
$nLotesNormalizados = 0
for ($i = 0; $i -lt $out.Count; $i++) {
    $antes = $out[$i]
    $depois = $antes.Replace('Mid$(CStr(mDB(i, COL_LOTE)), 4, 6)', 'NucleoLote(CStr(mDB(i, COL_LOTE)))')
    if ($depois -ne $antes) { $out[$i] = $depois; $nLotesNormalizados++ }
}
if ($nLotesNormalizados -lt 1) { throw 'Nenhuma comparacao de lote foi normalizada no motor.' }

# --- correcao de compilacao: aS/aM como identificadores ---
# VBA e insensivel a maiusculas, entao o identificador "aS" e o mesmo token que
# a palavra reservada "As" e o modulo NAO COMPILA. Como o VBA compila sob
# demanda, procedimento a procedimento, isso passava despercebido: AtualizarCalc
# rodava e AtualizarPainelEng, AtualizarEstatisticaAba e RegistrarEventosWestgard
# morriam em tempo de execucao com "nao foi possivel executar a macro".
# Foi tambem o gatilho da destruicao das formulas: quando a Fase 3A corrigiu o
# aS, esses tres procedimentos ganharam vida e sobrescreveram Painel e
# Estatistica, que ate entao sobreviviam por o motor nao conseguir compilar.
# -creplace (case-sensitive) de proposito: "As" reservado nao pode ser tocado.
$nRenomeadas = 0
for ($i = 0; $i -lt $out.Count; $i++) {
    $antes = $out[$i]
    $depois = $antes -creplace '\baS\b', 'alvoS' -creplace '\baM\b', 'alvoM'
    if ($depois -cne $antes) { $out[$i] = $depois; $nRenomeadas++ }
}

# --- lint: identificador com nome de palavra reservada ---
$reservadas = @('As', 'If', 'Then', 'Else', 'For', 'Next', 'To', 'Do', 'Loop',
    'And', 'Or', 'Not', 'Is', 'New', 'Set', 'Let', 'Sub', 'Function', 'End',
    'Dim', 'Type', 'Each', 'In', 'With', 'Call', 'Exit', 'Const', 'Byte')
$achados = New-Object System.Collections.ArrayList
for ($i = 0; $i -lt $out.Count; $i++) {
    $linha = $out[$i]
    if ($linha -match '^\s*''') { continue }
    foreach ($r in $reservadas) {
        # declaracao de variavel cujo nome so difere no caixa de uma reservada
        if ($linha -cmatch "\b(?!$r\b)(?i:$r)\b\s+As\s") {
            [void]$achados.Add("linha $($i + 1): $($linha.Trim())")
            break
        }
    }
}

[System.IO.File]::WriteAllLines($Saida, $out.ToArray(), $enc)

"producao : $($L.Count) linhas"
"gerado   : $($out.Count) linhas"
"aS/aM renomeados em $nRenomeadas linhas"
"lotes normalizados por NucleoLote em $nLotesNormalizados linhas"
if ($achados.Count -gt 0) {
    "LINT: identificador colidindo com palavra reservada -- o modulo nao vai compilar:"
    $achados | ForEach-Object { "  $_" }
    throw "Lint falhou: $($achados.Count) identificador(es) invalido(s)."
}
"lint    : ok (nenhum identificador colide com palavra reservada)"
"saida    : $Saida"
