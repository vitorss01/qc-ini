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
    [string]$NovoEstat                                   # AtualizarEstatisticaAba.txt (Marco 4)
)

# Duas codificacoes de proposito, e trocar uma pela outra corrompe o modulo:
#   $enc      cp1252  - modulos VBA exportados/importados pelo Excel
#   $encFonte UTF-8   - os .txt de origem, escritos por editor moderno
# Ler um .txt UTF-8 como cp1252 transforma Sheets("Estatistica" com acento) em
# Sheets("EstatA-stica") e o procedimento falha em tempo de execucao.
$enc = [System.Text.Encoding]::Default
$encFonte = New-Object System.Text.UTF8Encoding($false)

$L = [System.IO.File]::ReadAllLines($Producao, $enc)

# --- validacao das fronteiras: falha alto se o arquivo nao for o esperado ---
if ($L[298] -notlike "*MOTOR: MONTAR Calc*") {
    throw "Fronteira inicial inesperada na linha 299: $($L[298])"
}
if ($L[311] -notlike '*Sheets("Calc")*' -and $L[309] -notlike '*Sheets("Calc")*') {
    throw "Nao encontrei Set ws = Sheets(Calc) perto da linha 310"
}
if ($L[511].Trim() -ne 'End Sub') {
    throw "Fronteira final inesperada na linha 512: $($L[511])"
}
if ($L[15] -notlike '*NFD*') {
    throw "Fronteira de constantes inesperada na linha 16: $($L[15])"
}

$novo = [System.IO.File]::ReadAllLines($NovoSub, $encFonte)

$out = New-Object System.Collections.ArrayList

# 1..16 (constantes) + as duas novas
for ($i = 0; $i -le 15; $i++) { [void]$out.Add($L[$i]) }
[void]$out.Add("Private Const EF0 As Long = 3          ' Eng_Saida: 1a coluna de bloco de nivel")
[void]$out.Add("Private Const NEF As Long = 7          ' Eng_Saida: campos por nivel")
[void]$out.Add("Private Const COL_FILTRO As Long = 24  ' Eng_Saida: filtro de data por corrida")
[void]$out.Add("Private Const COL_VALOR0 As Long = 25  ' Eng_Saida: 1a coluna de valor por nivel")
[void]$out.Add("Private Const COL_CHAVE As Long = 28   ' Eng_Saida: chave logica ANALITO|RUN")
[void]$out.Add("Private Const LINHA_STAT As Long = 185 ' Eng_Saida: 1a linha do bloco de estatistica")
[void]$out.Add("Private Const LINHA_EST As Long = 190  ' Eng_Saida: 1a linha da tabela de parametros")

# 17..298 inalteradas
for ($i = 16; $i -le 297; $i++) { [void]$out.Add($L[$i]) }

# 299..512 substituidas (AtualizarCalc)
foreach ($linha in $novo) { [void]$out.Add($linha) }

# 513..678 inalteradas
for ($i = 512; $i -le 677; $i++) { [void]$out.Add($L[$i]) }

# 679..747 substituidas (AtualizarPainelEng), se fornecido
if ($NovoPainel) {
    if ($L[678] -notlike '*Sub AtualizarPainelEng*') {
        throw "Fronteira inesperada na linha 679: $($L[678])"
    }
    if ($L[746].Trim() -ne 'End Sub') {
        throw "Fronteira inesperada na linha 747: $($L[746])"
    }
    foreach ($linha in [System.IO.File]::ReadAllLines($NovoPainel, $encFonte)) { [void]$out.Add($linha) }
    $i0 = 747
}
else {
    for ($i = 678; $i -le 746; $i++) { [void]$out.Add($L[$i]) }
    $i0 = 747
}

# 748..749 inalteradas
for ($i = $i0; $i -le 748; $i++) { [void]$out.Add($L[$i]) }

# 750..799 substituidas (AtualizarEstatisticaAba), se fornecido
if ($NovoEstat) {
    if ($L[749] -notlike '*Sub AtualizarEstatisticaAba*') {
        throw "Fronteira inesperada na linha 750: $($L[749])"
    }
    if ($L[798].Trim() -ne 'End Sub') {
        throw "Fronteira inesperada na linha 799: $($L[798])"
    }
    foreach ($linha in [System.IO.File]::ReadAllLines($NovoEstat, $encFonte)) { [void]$out.Add($linha) }
    $i1 = 799
}
else {
    for ($i = 749; $i -le 798; $i++) { [void]$out.Add($L[$i]) }
    $i1 = 799
}

# 800..fim inalteradas
for ($i = $i1; $i -lt $L.Count; $i++) { [void]$out.Add($L[$i]) }

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
if ($achados.Count -gt 0) {
    "LINT: identificador colidindo com palavra reservada -- o modulo nao vai compilar:"
    $achados | ForEach-Object { "  $_" }
    throw "Lint falhou: $($achados.Count) identificador(es) invalido(s)."
}
"lint    : ok (nenhum identificador colide com palavra reservada)"
"saida    : $Saida"
