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
    [Parameter(Mandatory = $true)][string]$Producao,   # mEstatistica.bas de producao
    [Parameter(Mandatory = $true)][string]$NovoSub,    # AtualizarCalc.txt
    [Parameter(Mandatory = $true)][string]$Saida
)

$enc = [System.Text.Encoding]::Default
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

$novo = [System.IO.File]::ReadAllLines($NovoSub, $enc)

$out = New-Object System.Collections.ArrayList

# 1..16 (constantes) + as duas novas
for ($i = 0; $i -le 15; $i++) { [void]$out.Add($L[$i]) }
[void]$out.Add("Private Const EF0 As Long = 3          ' Eng_Saida: 1a coluna de bloco de nivel")
[void]$out.Add("Private Const NEF As Long = 7          ' Eng_Saida: campos por nivel")

# 17..298 inalteradas
for ($i = 16; $i -le 297; $i++) { [void]$out.Add($L[$i]) }

# 299..512 substituidas
foreach ($linha in $novo) { [void]$out.Add($linha) }

# 513..fim inalteradas
for ($i = 512; $i -lt $L.Count; $i++) { [void]$out.Add($L[$i]) }

[System.IO.File]::WriteAllLines($Saida, $out.ToArray(), $enc)

"producao : $($L.Count) linhas"
"gerado   : $($out.Count) linhas"
"saida    : $Saida"
