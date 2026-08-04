# buffer_dinamico.ps1 - fim do teto de 5.000 eventos de Westgard
#
# ACHADO DO TESTE DE ESTRESSE (04/08/2026). Com 10.000 linhas no banco, o
# buffer de eventos estourou: 5.000 registrados, 9.317 descartados. A guarda
# instalada de manha fez o que devia -- auditou e interrompeu em vez de
# descartar em silencio --, mas o teto em si era arbitrario.
#
# POR QUE UBound(mDB, 1) E O LIMITE CERTO, e nao um numero maior.
# Cada evento corresponde a uma combinacao (analito, nivel, corrida), e o
# codigo concatena as regras numa linha so ("13s+22s"). Logo o numero de
# eventos NUNCA excede o de linhas elegiveis do banco. O limite passa a ser
# correto por construcao, nao por folga -- trocar 5.000 por 20.000 apenas
# adiaria o mesmo erro.
#
# A LIMPEZA tambem estava engessada em "A4:H5003": historico maior que isso
# deixaria linhas orfas embaixo. Mas limpar ate ws.Rows.Count (1.048.576)
# custaria caro em TODA operacao do dia. Aqui limpa-se ate a ultima linha
# realmente usada, com piso na area antiga. Linhas 1 a 3 (titulo e cabecalho)
# nunca sao tocadas.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\buffer_dinamico.ps1 -Arquivo <mEstatistica.bas>

param(
    [Parameter(Mandatory = $true)][string]$Arquivo
)

$ErrorActionPreference = 'Stop'
$enc = [System.Text.Encoding]::Default
$texto = ([System.IO.File]::ReadAllLines($Arquivo, $enc) -join "`r`n")

if ($texto -like '*ultimaLinhaEv*') {
    "buffer dinamico: ja aplicado, nada a fazer"
    exit 0
}

# ---------- 1. limpeza proporcional ao historico existente ----------
$antesClear = '    ws.Range("A4:H5003").ClearContents'
$depoisClear = @(
    '    '' Limpa ate a ultima linha REALMENTE usada, com piso na area antiga.',
    '    '' Fixar em 5003 deixava orfas embaixo quando o historico crescia; usar',
    '    '' ws.Rows.Count limparia 1 milhao de linhas em toda operacao do dia.',
    '    '' Linhas 1 a 3 (titulo e cabecalho) nunca sao tocadas.',
    '    Dim ultimaLinhaEv As Long',
    '    ultimaLinhaEv = ws.Cells(ws.rows.Count, 1).End(xlUp).Row',
    '    If ultimaLinhaEv < 5003 Then ultimaLinhaEv = 5003',
    '    ws.Range(ws.Cells(4, 1), ws.Cells(ultimaLinhaEv, 8)).ClearContents'
) -join "`r`n"

if ($texto -notlike "*$antesClear*") { throw "Linha de limpeza A4:H5003 nao encontrada em $Arquivo" }
$texto = $texto.Replace($antesClear, $depoisClear)

# ---------- 2. buffer dimensionado pelo banco ----------
# Seguro: IsEmpty(mDB) ja foi testado antes deste ponto na rotina.
$antesRedim = '    ReDim ev(1 To 5000, 1 To 8)'
$depoisRedim = @(
    '    '' Dimensionado pelo BANCO, nao por numero fixo: eventos nunca excedem',
    '    '' as linhas elegiveis, porque cada linha de evento e uma combinacao',
    '    '' (analito, nivel, corrida) com as regras concatenadas. IsEmpty(mDB) ja',
    '    '' foi testado acima, entao o UBound e seguro.',
    '    ReDim ev(1 To UBound(mDB, 1), 1 To 8)'
) -join "`r`n"

if ($texto -notlike "*$antesRedim*") { throw "ReDim ev(1 To 5000) nao encontrado em $Arquivo" }
$texto = $texto.Replace($antesRedim, $depoisRedim)

# ---------- 3. a guarda de estouro fica, agora como invariante ----------
# Com o limite vindo do proprio banco, estourar passou a ser impossivel por
# construcao. A guarda permanece porque asserção que nunca dispara custa nada e
# denuncia o dia em que alguem mudar a regra de "um evento por combinacao".
$texto = $texto.Replace(
    'Buffer de " & UBound(ev, 1) & " eventos cheio; "',
    'Buffer dinamico de " & UBound(ev, 1) & " eventos cheio (INVARIANTE VIOLADA); "')

$texto = ($texto -replace "`r`n", "`n") -replace "`n", "`r`n"
[System.IO.File]::WriteAllText($Arquivo, $texto, $enc)
"buffer dinamico: limpeza proporcional + ReDim por UBound(mDB,1) aplicados"
