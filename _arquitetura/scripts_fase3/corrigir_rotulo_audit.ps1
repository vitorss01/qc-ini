# corrigir_rotulo_audit.ps1 - o Audit_Log parava de guardar o valor original
#
# DEFEITO (achado em 04/08/2026 ao PROVAR o item 2.2, nao ao ler o codigo).
#
# Rotulo() devolvia String: um resultado 92,0028 virava CStr -> "92,0028" e era
# gravado numa celula de formato Geral. O Excel reinterpreta a virgula como
# separador de MILHAR e a celula passa a conter 920028 (Double), exibido
# "920.028". Valor anterior e novo ficavam multiplicados por 10.000.
#
# Escondeu-se por dois motivos: AU_DELTA e gravado como NUMERO e continuava
# certo (7,77), e o hash da cadeia e calculado sobre o que se RELE da celula --
# a verificacao de integridade APROVAVA o registro corrompido.
#
# O CODIGO NOVO VIVE EM src_hardening1/Rotulo.txt, nao neste script.
# Montar VBA multilinha com array + -join dentro do .ps1 se mostrou fragil:
# o bloco chegou a ser gravado numa unica linha, o que comentou a funcao
# inteira (tudo apos o primeiro apostrofo) e derrubou a compilacao -- com o
# sintoma "O documento nao foi salvo" no aplicar_vba, longe da causa.
# Codigo VBA em arquivo proprio, lido com a codificacao certa, e o padrao que
# o resto do pipeline ja usa (AtualizarCalc.txt, mDados_RUN.txt).
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\corrigir_rotulo_audit.ps1 -Arquivo <mAuditoria.bas> -Novo <Rotulo.txt>

param(
    [Parameter(Mandatory = $true)][string]$Arquivo,
    [Parameter(Mandatory = $true)][string]$Novo
)

$ErrorActionPreference = 'Stop'

# cp1252 para o modulo VBA; UTF-8 para o .txt escrito por editor moderno
$enc = [System.Text.Encoding]::Default
$encFonte = New-Object System.Text.UTF8Encoding($false)

$L = [System.IO.File]::ReadAllLines($Arquivo, $enc)

if (($L -join ' ') -like '*Rotulo NAO pode devolver String*') {
    "rotulo do audit: ja corrigido, nada a fazer"
    exit 0
}

# localiza a funcao antiga pelos limites reais, sem depender de texto exato
$ini = -1; $fim = -1
for ($i = 0; $i -lt $L.Count; $i++) {
    if ($L[$i] -match '^\s*Private Function Rotulo\(') { $ini = $i; break }
}
if ($ini -lt 0) { throw "Private Function Rotulo nao encontrada em $Arquivo" }
for ($i = $ini; $i -lt $L.Count; $i++) {
    if ($L[$i] -match '^\s*End Function') { $fim = $i; break }
}
if ($fim -lt 0) { throw "End Function de Rotulo nao encontrado em $Arquivo" }

$novoCodigo = [System.IO.File]::ReadAllLines($Novo, $encFonte)

$out = New-Object System.Collections.ArrayList
for ($i = 0; $i -lt $ini; $i++) { [void]$out.Add($L[$i]) }
foreach ($linha in $novoCodigo) { [void]$out.Add($linha) }
for ($i = $fim + 1; $i -lt $L.Count; $i++) { [void]$out.Add($L[$i]) }

# WriteAllLines usa a quebra de linha do ambiente e mantem uma linha por item:
# nao ha join manual, entao nao ha como colapsar o bloco.
[System.IO.File]::WriteAllLines($Arquivo, $out.ToArray(), $enc)

# conferencia: a funcao tem de existir, em linha propria, e devolver Variant
$V = [System.IO.File]::ReadAllLines($Arquivo, $enc)
$assinatura = @($V | Where-Object { $_ -match '^\s*Private Function Rotulo\(ByVal v As Variant\) As Variant\s*$' })
if ($assinatura.Count -ne 1) {
    throw "Conferencia falhou: esperava 1 assinatura de Rotulo em linha propria, encontrei $($assinatura.Count). O bloco pode ter sido colapsado."
}
"rotulo do audit: numero deixa de virar texto ($($L.Count) -> $($V.Count) linhas)"
