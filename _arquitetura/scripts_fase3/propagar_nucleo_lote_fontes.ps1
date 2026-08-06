# propagar_nucleo_lote_fontes.ps1 - leva a correcao do nucleo do lote as FONTES
#
# POR QUE ESTE SCRIPT EXISTE
#
# corrigir_nucleo_lote.ps1 conserta o .xlsm de PRODUCAO. Isso resolve o hoje e
# nao resolve o amanha: aplicar_vba.ps1 reimporta mDados, mImportar e
# mEstatistica dos .bas versionados a cada build, e mEspecificacoes e conferido
# contra a fonte pela prova 1.1. Corrigir so o arquivo faria o proximo build
# REVERTER a correcao -- calado, como quase tudo neste projeto quando fonte e
# artefato discordam (ADR-021).
#
# O QUE TROCA
#
#   Mid$(<expr do codigo do lote>, 4, 6)   ->   NucleoLote(<expr>)
#   NucleoLote = Mid$(Trim$(codigo), 4, 6) ->   ... Len(Trim$(codigo)) - 5
#
# NAO TROCA a linha de atribuicao da propria NucleoLote pelo padrao geral: viraria
# NucleoLote = NucleoLote(...), recursao infinita. Ela tem tratamento proprio.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\propagar_nucleo_lote_fontes.ps1 [-Simular]

param([switch]$Simular)

$ErrorActionPreference = 'Stop'
$s = Split-Path -Parent $MyInvocation.MyCommand.Path
$arq = Split-Path -Parent $s

$alvos = @(
    'src_producao\mEspecificacoes.bas',
    'src_producao\mImportar.bas',
    'src_producao\mDados_UPSERT.txt',
    'src_hardening1\mImportar.bas',
    'src_hardening1\mDados.bas',
    'src_hardening1\mEstatistica.bas',
    'src_hardening1\Bioquimica\mImportar.bas',
    'src_hardening1\Bioquimica\mDados.bas',
    'src_hardening1\Bioquimica\mEstatistica.bas',
    'src_hardening1\Hematologia\mImportar.bas',
    'src_hardening1\Hematologia\mDados.bas',
    'src_hardening1\Hematologia\mEstatistica.bas',
    'snapshot_producao\Bioquimica\vba\mEntrada.bas',
    'snapshot_producao\Bioquimica\vba\mDados.bas',
    'snapshot_producao\Bioquimica\vba\mOperacao.bas',
    'snapshot_producao\Hematologia\vba\mEntrada.bas',
    'snapshot_producao\Hematologia\vba\mDados.bas',
    'snapshot_producao\Hematologia\vba\mOperacao.bas'
)

$enc = [System.Text.Encoding]::Default
$reMid = [regex]'Mid\$?\(\s*(.+?)\s*,\s*4\s*,\s*6\s*\)'
$totalArq = 0
$totalLin = 0

foreach ($rel in $alvos) {
    $p = Join-Path $arq $rel
    if (-not (Test-Path $p)) { continue }
    $linhas = [System.IO.File]::ReadAllLines($p, $enc)
    $mudou = 0
    for ($i = 0; $i -lt $linhas.Count; $i++) {
        $l = $linhas[$i]
        if ($l -match '^\s*''') { continue }
        if ($l -match 'Function\s+NucleoLote') { continue }

        # a definicao: troca o 6 pelo comprimento, sem virar chamada a si mesma
        if ($l -match '^\s*NucleoLote\s*=\s*Mid') {
            $novo = '    NucleoLote = Mid$(Trim$(codigo), 4, Len(Trim$(codigo)) - 5)'
            if ($novo -ne $l) { $linhas[$i] = $novo; $mudou++ }
            continue
        }

        if ($l -notmatch 'Mid\$?\(') { continue }
        if ($l -notmatch 'COL_LOTE|lote|Lote|codigo|regs\(i, 4\)') { continue }
        $novo = $reMid.Replace($l, 'NucleoLote($1)')
        if ($novo -ne $l) { $linhas[$i] = $novo; $mudou++ }
    }
    if ($mudou -gt 0) {
        "$rel : $mudou linha(s)"
        if (-not $Simular) { [System.IO.File]::WriteAllLines($p, $linhas, $enc) }
        $totalArq++
        $totalLin += $mudou
    }
}

""
"$totalLin linha(s) em $totalArq arquivo(s)"
if ($Simular) { "SIMULACAO -- nada foi gravado." }

# lint: nao pode sobrar nenhum Mid(...,4,6) sobre codigo de lote
$resto = @()
foreach ($rel in $alvos) {
    $p = Join-Path $arq $rel
    if (-not (Test-Path $p)) { continue }
    $i = 0
    foreach ($l in [System.IO.File]::ReadAllLines($p, $enc)) {
        $i++
        if ($l -match '^\s*''') { continue }
        if ($l -match 'Mid\$?\(\s*.+?\s*,\s*4\s*,\s*6\s*\)' -and $l -match 'COL_LOTE|lote|Lote|codigo|regs\(i, 4\)') {
            $resto += "$rel`:$i -> $($l.Trim())"
        }
    }
}
if ($resto.Count -gt 0 -and -not $Simular) {
    $resto | ForEach-Object { "  RESTOU: $_" }
    throw "Ainda ha $($resto.Count) extracao(oes) fixa de 6 sobre codigo de lote."
}
"lint: nenhuma extracao fixa de 6 restante"
