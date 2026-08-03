# comparar_veredictos.ps1 - mede o impacto real da correcao de Westgard
#
# Para cada analito, compara o veredicto REJEITADO/OK do Calc em duas versoes:
#   REFERENCIA = producao, onde as formulas do Calc calculam Westgard sozinhas
#   BUILD      = hardening, onde o Calc le Eng_Saida (motor)
#
# As duas divergem em 4 das 5 regras (41s 4 vs 3 pontos, 10x 10 vs 6, R4s
# amplitude>4 vs max>2 e min<-2, e 22s no mesmo nivel ausente da planilha).
# Este script diz em quantas corridas isso muda o que o analista ve.
#
# Colunas de veredicto no Calc: P (nivel 1), AL (nivel 2), BH (nivel 3),
# linhas 3..182.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\comparar_veredictos.ps1 -Referencia <producao_copia.xlsm> -Build <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Referencia,
    [Parameter(Mandatory = $true)][string]$Build,
    [string]$OutCsv
)

$KC0 = 3; $NK = 180
$colsVeredicto = @{ 1 = 16; 2 = 38; 3 = 60 }    # P, AL, BH

function Abrir($caminho) {
    $ErrorActionPreference = 'Stop'

$xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
    $xl.AutomationSecurity = 1      # precisa rodar o motor no build
    return @{ App = $xl; Wb = $xl.Workbooks.Open($caminho) }
}

function LerVeredictos($ctx, $analito, $rodarMotor) {
    $wb = $ctx.Wb
    $wb.Names('selAnalito').RefersToRange.Value2 = $analito
    # Out-Null obrigatorio: Application.Run devolve valor e Calculate emite saida.
    # Sem isso o retorno da funcao vira array e o hashtable se perde.
    if ($rodarMotor) { $ctx.App.Run("$($wb.Name)!AtualizarCalc") | Out-Null }
    $ctx.App.Calculate() | Out-Null
    $calc = $wb.Worksheets('Calc')
    $res = @{}
    foreach ($nivel in $colsVeredicto.Keys) {
        $c = $colsVeredicto[$nivel]
        $v = $calc.Range($calc.Cells($KC0, $c), $calc.Cells($KC0 + $NK - 1, $c)).Value2
        $lista = New-Object System.Collections.ArrayList
        for ($i = 1; $i -le $NK; $i++) { [void]$lista.Add([string]$v.GetValue($i, 1)) }
        $res[$nivel] = $lista
    }
    # coluna B = RUN, para identificar a corrida divergente
    $b = $calc.Range($calc.Cells($KC0, 2), $calc.Cells($KC0 + $NK - 1, 2)).Value2
    $runs = New-Object System.Collections.ArrayList
    for ($i = 1; $i -le $NK; $i++) { [void]$runs.Add([string]$b.GetValue($i, 1)) }
    $res['RUN'] = $runs
    return $res
}

$ref = Abrir $Referencia
$bld = Abrir $Build

# lista de analitos da aba Analitos (A4:A43)
$analitos = New-Object System.Collections.ArrayList
$aba = $ref.Wb.Worksheets('Analitos')
for ($r = 4; $r -le 43; $r++) {
    $a = [string]$aba.Cells($r, 1).Value2
    if ($a.Trim() -ne '') { [void]$analitos.Add($a.Trim()) }
}
"analitos: $($analitos.Count)"
""

$divs = New-Object System.Collections.ArrayList
foreach ($a in $analitos) {
    $vr = LerVeredictos $ref $a $false
    $vb = LerVeredictos $bld $a $true
    $n = 0
    foreach ($nivel in 1..3) {
        for ($i = 0; $i -lt $NK; $i++) {
            if ($vr[$nivel][$i] -ne $vb[$nivel][$i]) {
                $n++
                [void]$divs.Add([pscustomobject]@{
                        Analito = $a; Nivel = $nivel; RUN = $vr['RUN'][$i]
                        Producao = $vr[$nivel][$i]; Motor = $vb[$nivel][$i]
                    })
            }
        }
    }
    "{0,-10} divergencias={1}" -f $a, $n
}

""
"TOTAL de corridas com veredicto diferente: $($divs.Count)"
if ($divs.Count -gt 0) {
    $divs | Group-Object Producao, Motor | ForEach-Object { "  {0,-22} n={1}" -f $_.Name, $_.Count }
    if ($OutCsv) {
        $divs | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
        "detalhe em: $OutCsv"
    }
}

$ref.Wb.Close($false); $ref.App.Quit()
$bld.Wb.Close($false); $bld.App.Quit()
