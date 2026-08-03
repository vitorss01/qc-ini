# varredura_adr019.ps1 - Marco 5 do Sprint HARDENING 1
#
# Verifica a regra do ADR-019 na versao refinada que o gestor aprovou:
#
#   PERMITIDO a planilha  selecionar, filtrar, ordenar e localizar pontos
#                         (COUNTIFS, SUMIFS, MAXIFS, MINIFS, INDEX/MATCH,
#                          AGGREGATE de posicao: 14 grande, 15 pequeno)
#   PROIBIDO a planilha   calcular PARAMETRO ESTATISTICO -- media, DP, CV,
#                         bias, erro total, sigma, z-score, regra de Westgard.
#                         Isso e do motor, publicado em Eng_Saida.
#
# A varredura nao procura so a string "DB_Resultados": os nomes definidos
# rAnalito, rValor, rRUN, rLote, rStatus, rData, rNivel, rFirst e rRunUnico
# apontam para o banco, e e por eles que as formulas o alcancam.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\varredura_adr019.ps1 -Formulas <snapshot.csv> [-OutCsv <achados.csv>]

param(
    [Parameter(Mandatory = $true)][string]$Formulas,
    [string]$OutCsv
)

$nomesBanco = @('rAnalito', 'rData', 'rFirst', 'rLote', 'rNivel', 'rRUN',
    'rRunUnico', 'rStatus', 'rValor')
$regexBanco = '(?i)(DB_Resultados|\b(' + ($nomesBanco -join '|') + ')\b)'

# funcoes que produzem parametro estatistico
$regexEstat = '(?i)\b(AVERAGEIFS|AVERAGE|STDEV\.?[SP]?|VAR\.?[SP]?|MEDIAN|' +
              'AGGREGATE\s*\(\s*(1|2|3|4|5|6|7|8)\s*,)'

# abas de interface: o que o usuario ve. Eng_Saida e camada do motor;
# DB_Resultados e o proprio banco; ambas ficam fora da regra.
$abasInterface = @('Painel', 'Calc', 'Liberação', 'Liberacao', 'Registros',
    'Início', 'Inicio', 'Configuração', 'Configuracao', 'Analitos')

$todas = Import-Csv $Formulas -Delimiter ';'

$achados = New-Object System.Collections.ArrayList
foreach ($f in $todas) {
    $aba = $f.Aba
    $ehInterface = ($abasInterface -contains $aba) -or ($aba -like 'Estat*')
    if (-not $ehInterface) { continue }
    if ($f.FormulaR1C1 -notmatch $regexBanco) { continue }

    $tipo = if ($f.FormulaR1C1 -match $regexEstat) { 'ESTATISTICA' } else { 'selecao' }
    [void]$achados.Add([pscustomobject]@{
            Aba      = $aba
            Endereco = $f.Endereco
            Tipo     = $tipo
            Formula  = $f.FormulaR1C1
        })
}

"formulas de interface que alcancam o banco: $($achados.Count)"
""
$achados | Group-Object Aba, Tipo | Sort-Object Name | ForEach-Object {
    "  {0,-26} n={1}" -f $_.Name, $_.Count
}
""
$violacoes = @($achados | Where-Object Tipo -eq 'ESTATISTICA')
if ($violacoes.Count -eq 0) {
    "ADR-019: OK - nenhuma formula de interface calcula parametro estatistico."
}
else {
    "ADR-019: $($violacoes.Count) formula(s) de interface calculando parametro estatistico:"
    $violacoes | Group-Object Aba | ForEach-Object { "  {0,-16} n={1}" -f $_.Name, $_.Count }
    $violacoes | Select-Object -First 5 | ForEach-Object {
        $t = $_.Formula; if ($t.Length -gt 130) { $t = $t.Substring(0, 130) + '...' }
        "    $($_.Aba)!$($_.Endereco)  $t"
    }
}

if ($OutCsv) {
    $achados | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    ""
    "detalhe em: $OutCsv"
}
