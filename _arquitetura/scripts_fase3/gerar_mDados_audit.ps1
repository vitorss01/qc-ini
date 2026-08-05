# gerar_mDados_audit.ps1 - aplica a camada de auditoria sobre mDados.bas
#
# Sprint HARDENING 3, itens 3.1 e 3.2 do Quality Gate.
#
# Substitui, POR NOME, os blocos UpsertResultados, ExcluirLogico e RegistrarLog
# pela versao auditada (mDados_AUDIT.txt). A substituicao e por nome e nao por
# numero de linha porque mDados.bas ja e um artefato gerado -- amarrar em linha
# aqui quebraria a cada mudanca a montante.
#
# Duas codificacoes, e trocar uma pela outra corrompe o modulo:
#   $enc      cp1252  - modulos VBA
#   $encFonte UTF-8   - o .txt de origem
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\gerar_mDados_audit.ps1 -Entrada <mDados.bas> -Patch <mDados_AUDIT.txt> -Saida <mDados.bas>

param(
    [Parameter(Mandatory = $true)][string]$Entrada,
    [Parameter(Mandatory = $true)][string]$Patch,
    [Parameter(Mandatory = $true)][string]$Saida
)

$ErrorActionPreference = 'Stop'
$enc = [System.Text.Encoding]::Default
$encFonte = New-Object System.Text.UTF8Encoding($false)

$L = [System.IO.File]::ReadAllLines($Entrada, $enc)
$P = [System.IO.File]::ReadAllLines($Patch, $encFonte)

# --- 1. quebra o patch em blocos nomeados -------------------------------------
$blocos = @{}
$nome = $null
$buf = New-Object System.Collections.ArrayList
foreach ($linha in $P) {
    if ($linha -match "^'@@FUNC\s+(\w+)\s*$") {
        if ($nome) { $blocos[$nome] = $buf.ToArray() }
        $nome = $Matches[1]
        $buf = New-Object System.Collections.ArrayList
    }
    elseif ($nome) { [void]$buf.Add($linha) }
}
if ($nome) { $blocos[$nome] = $buf.ToArray() }
if ($blocos.Count -eq 0) { throw "Nenhum bloco '@@FUNC encontrado em $Patch" }

# --- 2. localiza cada rotina no modulo ----------------------------------------
# Inclui as linhas de comentario imediatamente acima da assinatura: elas
# documentam a rotina e sao substituidas junto.
function Get-Bloco {
    param([string[]]$Linhas, [string]$Nome)
    $ini = -1; $fim = -1
    for ($i = 0; $i -lt $Linhas.Count; $i++) {
        if ($Linhas[$i] -match "^\s*(Public|Private|Friend)?\s*(Sub|Function)\s+$Nome\s*\(") { $ini = $i; break }
    }
    if ($ini -lt 0) { return $null }
    $tipo = if ($Linhas[$ini] -match '\bSub\b') { 'Sub' } else { 'Function' }
    for ($i = $ini + 1; $i -lt $Linhas.Count; $i++) {
        if ($Linhas[$i].Trim() -eq "End $tipo") { $fim = $i; break }
    }
    if ($fim -lt 0) { throw "Fim de $Nome nao encontrado (esperado 'End $tipo')" }
    while ($ini -gt 0 -and $Linhas[$ini - 1] -match "^\s*'") { $ini-- }
    return @{ Inicio = $ini; Fim = $fim }
}

$alvos = @()
foreach ($n in $blocos.Keys) {
    $b = Get-Bloco -Linhas $L -Nome $n
    if (-not $b) { throw "Rotina $n nao encontrada em $Entrada" }
    $alvos += [pscustomobject]@{ Nome = $n; Inicio = $b.Inicio; Fim = $b.Fim }
}
$alvos = $alvos | Sort-Object Inicio

# blocos nao podem se sobrepor
for ($i = 1; $i -lt $alvos.Count; $i++) {
    if ($alvos[$i].Inicio -le $alvos[$i - 1].Fim) {
        throw "Blocos sobrepostos: $($alvos[$i-1].Nome) e $($alvos[$i].Nome)"
    }
}

# --- 3. remonta --------------------------------------------------------------
$out = New-Object System.Collections.ArrayList
$cursor = 0
foreach ($a in $alvos) {
    for ($i = $cursor; $i -lt $a.Inicio; $i++) { [void]$out.Add($L[$i]) }
    foreach ($linha in $blocos[$a.Nome]) { [void]$out.Add($linha) }
    $cursor = $a.Fim + 1
}
for ($i = $cursor; $i -lt $L.Count; $i++) { [void]$out.Add($L[$i]) }

# --- 4. lint -----------------------------------------------------------------
# Const de modulo depois da primeira rotina nao compila em VBA.
$primeiroProc = -1
for ($i = 0; $i -lt $out.Count; $i++) {
    if ($out[$i] -match '^\s*(Public |Private |Friend )?(Sub|Function)\s+\w+') { $primeiroProc = $i; break }
}
if ($primeiroProc -ge 0) {
    for ($i = $primeiroProc; $i -lt $out.Count; $i++) {
        if ($out[$i] -match '^\s*(Public|Private)\s+Const\s+') {
            throw "Const de modulo na linha $($i + 1), depois da primeira rotina: $($out[$i].Trim())"
        }
    }
}

# Identificador colidindo com palavra reservada -- a mesma classe do aS.
$reservadas = @('As', 'If', 'Then', 'Else', 'For', 'Next', 'To', 'Do', 'Loop',
    'And', 'Or', 'Not', 'Is', 'New', 'Set', 'Let', 'Sub', 'Function', 'End',
    'Dim', 'Type', 'Each', 'In', 'With', 'Call', 'Exit', 'Const', 'Byte',
    'Imp', 'Eqv', 'Xor', 'Mod', 'Like', 'Step', 'Case', 'Erase', 'Stop')
$achados = New-Object System.Collections.ArrayList
for ($i = 0; $i -lt $out.Count; $i++) {
    if ($out[$i] -match '^\s*''') { continue }
    foreach ($r in $reservadas) {
        if ($out[$i] -cmatch "\b(?!$r\b)(?i:$r)\b\s+As\s") {
            [void]$achados.Add("linha $($i + 1): $($out[$i].Trim())")
            break
        }
    }
}

[System.IO.File]::WriteAllLines($Saida, $out.ToArray(), $enc)

"entrada  : $($L.Count) linhas"
"blocos   : $($alvos.Nome -join ', ')"
"gerado   : $($out.Count) linhas"
if ($achados.Count -gt 0) {
    "LINT: identificador colidindo com palavra reservada:"
    $achados | ForEach-Object { "  $_" }
    throw "Lint falhou: $($achados.Count) identificador(es) invalido(s)."
}
"lint     : ok"
"saida    : $Saida"
