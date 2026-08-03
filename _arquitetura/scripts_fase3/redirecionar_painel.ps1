# redirecionar_painel.ps1 - Marco 3 do Sprint HARDENING 1
#
# O Painel deixa de calcular. Cada celula do bloco por nivel passa a ler
# Eng_Saida por INDEX(engPainel, nivel, coluna). A linha 185+t de Eng_Saida
# espelha a linha 7+t do Painel, coluna a coluna -- por isso o indice de coluna
# do INDEX e a propria coluna do Painel.
#
# 45 celulas eram formulas na producao (uma TERCEIRA implementacao de
# estatistica, alem do motor e do Calc) e passam a referencia: ALTERADA.
# 9 celulas (S,T,U das 3 linhas) nao tinham formula -- o motor escrevia valor
# direto nelas. Viram formula: EXTRA. As duas listas ficam declaradas.
#
# Rotulos de S6/T6/U6 e largura das colunas 19..21 eram refeitos pelo motor a
# cada execucao. Passam a ser escritos uma vez, aqui.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\redirecionar_painel.ps1 -Workbook <build.xlsm> [-OutCsv <lista.csv>]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$OutCsv
)

$NLV = 3
$LINHA_PAINEL = 7      # Painel: 1a linha de nivel
$colunas = @(2, 3, 4, 5, 6, 7, 8, 9, 10, 13, 14, 15, 16, 17, 18, 19, 20, 21)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3

$wb = $xl.Workbooks.Open($Workbook)
$xl.Calculation = -4135
$pnl = $wb.Worksheets('Painel')

$lista = New-Object System.Collections.ArrayList

for ($t = 0; $t -lt $NLV; $t++) {
    $linha = $LINHA_PAINEL + $t
    foreach ($c in $colunas) {
        $cel = $pnl.Cells($linha, $c)
        $tinha = $cel.HasFormula
        # O IF(...="","",...) nao e decorativo: INDEX sobre celula vazia devolve
        # 0, e o motor grava "" quando nao ha dado (n=0, sem violacao). Sem o
        # guarda, o Painel exibiria 0 onde a producao exibia vazio. O zero
        # legitimo das contagens sobrevive, porque em Excel 0="" e FALSO.
        $ref = "INDEX(engPainel,$($t + 1),$c)"
        $cel.FormulaR1C1 = "=IF($ref=" + '""' + "," + '""' + ",$ref)"
        [void]$lista.Add([pscustomobject]@{
                Nivel    = $t + 1
                Celula   = $cel.Address($false, $false)
                Coluna   = $c
                Tipo     = if ($tinha) { 'ALTERADA' } else { 'EXTRA' }
            })
    }
}

# rotulos fixos, antes refeitos pelo motor a cada execucao
$pnl.Cells(6, 19).Value2 = [char]0x00DA + 'lt. viola' + [char]0x00E7 + [char]0x00E3 + 'o'
$pnl.Cells(6, 20).Value2 = 'Classifica' + [char]0x00E7 + [char]0x00E3 + 'o'
$pnl.Cells(6, 21).Value2 = 'Hist' + [char]0x00F3 + 'rico'
$pnl.Range($pnl.Cells(6, 19), $pnl.Cells(6, 21)).Font.Bold = $true
$pnl.Columns(19).ColumnWidth = 18
$pnl.Columns(20).ColumnWidth = 18
$pnl.Columns(21).ColumnWidth = 22

$xl.Calculation = -4105
$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

"celulas redirecionadas: $($lista.Count)"
$lista | Group-Object Tipo | ForEach-Object { "  {0,-9} n={1}" -f $_.Name, $_.Count }
if ($OutCsv) {
    $lista | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    "lista fechada em: $OutCsv"
}
