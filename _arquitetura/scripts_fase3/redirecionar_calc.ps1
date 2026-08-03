# redirecionar_calc.ps1 - Marco 2 do Sprint HARDENING 1
#
# Reescreve os 6 campos de regra Westgard de cada nivel do Calc para que leiam
# Eng_Saida em vez de recalcular. Sao 6 campos x 3 niveis x 180 linhas = 3.240
# celulas. Todo o resto do Calc fica intacto.
#
# Por que casar por RUN e nao por posicao de linha: a coluna B do Calc monta a
# lista de corridas com rFirst=1 e loteAtivo, e o motor monta a dele com
# EhElegivel e o lote core. As duas listas coincidem hoje, mas nada garante que
# coincidam sempre. MATCH sobre engRUN e correto por construcao.
#
# Guarda preservada: a formula continua devolvendo 0 (flags) ou "" (veredicto)
# quando a celula de valor do nivel esta vazia, exatamente como a producao.
# Assim linhas sem dado nao mudam de aparencia.
#
# Mapa de colunas (CF0=6, NFD=22, NLV=3):
#   nivel 1: valor=F(6)   regras K..P   (11..16)
#   nivel 2: valor=AB(28) regras AG..AL (33..38)
#   nivel 3: valor=AX(50) regras BC..BH (55..60)
# Eng_Saida (EF0=3, NEF=7): nivel t ocupa 3+t*7 .. 3+t*7+6
#   +0 R13s  +1 R22s  +2 RR4s  +3 R41s  +4 R10x  +5 Alerta12s  +6 Veredicto
# O campo 11 do Calc (veredicto) le o +6; o Alerta12s fica disponivel e sem
# consumidor por enquanto, de proposito.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\redirecionar_calc.ps1 -Workbook <build.xlsm> [-OutCsv <lista_fechada.csv>]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$OutCsv
)

$NLV = 3; $KC0 = 3; $NK = 180; $CF0 = 6; $NFD = 22
$EF0 = 3; $NEF = 7

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3

$wb = $xl.Workbooks.Open($Workbook)
$xl.Calculation = -4135      # manual: so vale com pasta aberta
$calc = $wb.Worksheets('Calc')

$lista = New-Object System.Collections.ArrayList
$n = 0

for ($t = 0; $t -lt $NLV; $t++) {
    $colValor = $CF0 + $t * $NFD               # campo 1 do bloco: o valor
    $colRegra0 = $CF0 + $t * $NFD + 5          # campo 6: primeira regra
    $engBase = $EF0 + $t * $NEF

    for ($k = 0; $k -lt 6; $k++) {
        $colCalc = $colRegra0 + $k
        $colEng = $engBase + $k
        $ehVeredicto = ($k -eq 5)
        if ($ehVeredicto) { $colEng = $engBase + 6 }   # pula Alerta12s

        $vazio = if ($ehVeredicto) { '""' } else { '0' }
        $f = "=IF(RC$colValor=" + '""' + ",$vazio,IFERROR(INDEX(engDados,MATCH(RC2,engRUN,0),$colEng),$vazio))"

        $rng = $calc.Range($calc.Cells($KC0, $colCalc), $calc.Cells($KC0 + $NK - 1, $colCalc))
        $rng.FormulaR1C1 = $f
        $n += $NK

        [void]$lista.Add([pscustomobject]@{
                Nivel      = $t + 1
                Campo      = if ($ehVeredicto) { 'Veredicto' } else { @('R13s', 'R22s', 'RR4s', 'R41s', 'R10x')[$k] }
                ColunaCalc = $colCalc
                Intervalo  = $rng.Address()
                ColunaEng  = $colEng
                Formula    = $f
            })
    }
}

$xl.Calculation = -4105      # xlCalculationAutomatic
$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

"celulas redirecionadas: $n"
$lista | Format-Table Nivel, Campo, Intervalo, ColunaEng -AutoSize
if ($OutCsv) {
    $lista | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    "lista fechada em: $OutCsv"
}
