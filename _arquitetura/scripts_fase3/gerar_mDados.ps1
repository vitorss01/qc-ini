# gerar_mDados.ps1 - monta a versao do hardening de mDados.bas
#
# Substitui NovoRUN (linhas 49..66 da versao de producao) pela API de
# identidade da corrida: contador persistido, registro de corridas e
# ObterOuCriarRUN. Ver mDados_RUN.txt.
#
# Duas codificacoes, e trocar uma pela outra corrompe o modulo:
#   $enc      cp1252  - modulos VBA (mDados tem "Excluido" com acento)
#   $encFonte UTF-8   - o .txt de origem
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\gerar_mDados.ps1 -Producao <mDados.bas> -NovoRun <mDados_RUN.txt> -Saida <mDados.bas>

param(
    [Parameter(Mandatory = $true)][string]$Producao,
    [Parameter(Mandatory = $true)][string]$NovoRun,
    [Parameter(Mandatory = $true)][string]$Saida
)

$enc = [System.Text.Encoding]::Default
$encFonte = New-Object System.Text.UTF8Encoding($false)

$L = [System.IO.File]::ReadAllLines($Producao, $enc)

# fronteiras: falha alto se o arquivo nao for o esperado
if ($L[48] -notlike '*RUN = chave logica da corrida*') {
    throw "Fronteira inicial inesperada na linha 49: $($L[48])"
}
if ($L[49] -notlike '*Function NovoRUN*') {
    throw "Fronteira inesperada na linha 50: $($L[49])"
}
if ($L[65].Trim() -ne 'End Function') {
    throw "Fronteira final inesperada na linha 66: $($L[65])"
}

# O .txt vem em duas partes, separadas por marcadores. Em VBA, "Public Const"
# so vale na SECAO DE DECLARACOES, no topo do modulo: uma constante declarada
# depois da primeira Function nao compila -- o erro que aparece e "variavel nao
# definida" na primeira linha que usa a constante, longe da causa.
$patch = [System.IO.File]::ReadAllLines($NovoRun, $encFonte)
$iC = [array]::FindIndex($patch, [Predicate[string]] { $args[0].Trim() -eq "'@@CONSTANTES" })
$iF = [array]::FindIndex($patch, [Predicate[string]] { $args[0].Trim() -eq "'@@FUNCOES" })
if ($iC -lt 0 -or $iF -lt 0 -or $iF -le $iC) {
    throw "Marcadores '@@CONSTANTES / '@@FUNCOES ausentes ou fora de ordem em $NovoRun"
}
$constantes = $patch[($iC + 1)..($iF - 1)]
$funcoes = $patch[($iF + 1)..($patch.Count - 1)]

if ($L[16] -notlike '*ST_EXCLUIDO*') {
    throw "Fim do bloco de constantes inesperado na linha 17: $($L[16])"
}

$out = New-Object System.Collections.ArrayList

# 1..17: constantes de producao, e logo apos as novas (secao de declaracoes)
for ($i = 0; $i -le 16; $i++) { [void]$out.Add($L[$i]) }
[void]$out.Add('')
foreach ($linha in $constantes) { [void]$out.Add($linha) }

# 18..48 inalteradas
for ($i = 17; $i -le 47; $i++) { [void]$out.Add($L[$i]) }

# 49..66 (NovoRUN antigo) substituidas pelas novas funcoes
foreach ($linha in $funcoes) { [void]$out.Add($linha) }

# 67..fim inalteradas
for ($i = 66; $i -lt $L.Count; $i++) { [void]$out.Add($L[$i]) }

# lint: Const/Dim de modulo declarado depois da primeira Function ou Sub.
# VBA so aceita declaracao de modulo na secao de declaracoes; depois disso o
# erro aparece como "variavel nao definida" na linha que USA a constante --
# longe da causa, e facil de diagnosticar errado.
$primeiroProc = -1
for ($i = 0; $i -lt $out.Count; $i++) {
    if ($out[$i] -match '^\s*(Public |Private |Friend )?(Sub|Function)\s+\w+') { $primeiroProc = $i; break }
}
if ($primeiroProc -ge 0) {
    for ($i = $primeiroProc; $i -lt $out.Count; $i++) {
        if ($out[$i] -match '^\s*(Public|Private)\s+Const\s+') {
            throw "Const de modulo na linha $($i + 1), depois da primeira rotina (linha $($primeiroProc + 1)). Nao compila: $($out[$i].Trim())"
        }
    }
}

# lint: identificador colidindo com palavra reservada (a mesma classe do aS)
$reservadas = @('As', 'If', 'Then', 'Else', 'For', 'Next', 'To', 'Do', 'Loop',
    'And', 'Or', 'Not', 'Is', 'New', 'Set', 'Let', 'Sub', 'Function', 'End',
    'Dim', 'Type', 'Each', 'In', 'With', 'Call', 'Exit', 'Const', 'Byte',
    'Imp', 'Eqv', 'Xor', 'Mod', 'Like', 'Step', 'Case', 'Erase', 'Stop')
$achados = New-Object System.Collections.ArrayList
for ($i = 0; $i -lt $out.Count; $i++) {
    $linha = $out[$i]
    if ($linha -match '^\s*''') { continue }
    foreach ($r in $reservadas) {
        if ($linha -cmatch "\b(?!$r\b)(?i:$r)\b\s+As\s") {
            [void]$achados.Add("linha $($i + 1): $($linha.Trim())")
            break
        }
    }
}

[System.IO.File]::WriteAllLines($Saida, $out.ToArray(), $enc)

"producao : $($L.Count) linhas"
"gerado   : $($out.Count) linhas"
if ($achados.Count -gt 0) {
    "LINT: identificador colidindo com palavra reservada:"
    $achados | ForEach-Object { "  $_" }
    throw "Lint falhou: $($achados.Count) identificador(es) invalido(s)."
}
"lint     : ok"
"saida    : $Saida"
