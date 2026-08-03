# criar_corridas.ps1 - registro de corridas + contador persistido do RUN
#
# Cria a aba Corridas e migra o historico. A migracao e ADITIVA: nenhum RUN
# existente muda de valor. Constroi o registro a partir dos triplos distintos
# (RUN, Data, nucleo do lote) ja gravados em DB_Resultados e posiciona o
# contador em max(RUN) + 1.
#
# Consequencia deliberada: dias que de fato tiveram duas corridas continuam
# colapsados num RUN so. Isso NAO e corrigido retroativamente -- inventar
# separacao de historico seria pior que registrar a limitacao. O modelo novo
# vale da migracao em diante.
#
# Layout:
#   D1            contador (nome definido proxRUN)
#   linha 3       cabecalhos
#   linha 4+      RUN | Data | Hora | Turno | Lote | Equipamento | Usuario |
#                 CriadoEm | CriadoPor
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_corridas.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$R0 = 4

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3

$wb = $xl.Workbooks.Open($Workbook)

foreach ($ws in $wb.Worksheets) {
    if ($ws.Name -eq 'Corridas') { $ws.Visible = -1; $ws.Delete(); "Corridas anterior removida"; break }
}

$cor = $wb.Worksheets.Add()
$cor.Name = 'Corridas'

$cor.Range('A1').Value2 = 'Registro de corridas'
$cor.Range('C1').Value2 = 'Proximo RUN:'
$cor.Range('A1:C1').Font.Bold = $true

$cab = @('RUN', 'Data', 'Hora', 'Turno', 'Lote', 'Equipamento', 'Usuario', 'CriadoEm', 'CriadoPor')
for ($k = 0; $k -lt $cab.Count; $k++) { $cor.Cells(3, $k + 1).Value2 = $cab[$k] }
$cor.Range($cor.Cells(3, 1), $cor.Cells(3, $cab.Count)).Font.Bold = $true

# ---- migracao: triplos distintos (RUN, Data, nucleo do lote) do banco ----
$db = $wb.Worksheets('DB_Resultados')
$ult = $db.Cells($db.Rows.Count, 1).End(-4162).Row
"linhas no banco: $($ult - 3)"

$vistos = @{}
$ordem = New-Object System.Collections.ArrayList
if ($ult -ge $R0) {
    $dados = $db.Range($db.Cells($R0, 1), $db.Cells($ult, 7)).Value2
    $nL = $dados.GetLength(0)
    for ($i = 1; $i -le $nL; $i++) {
        $run = $dados.GetValue($i, 1)
        if ($run -eq $null -or -not ($run -is [double])) { continue }
        $r = [int]$run
        if ($vistos.ContainsKey($r)) { continue }
        $serial = $dados.GetValue($i, 2)
        $lote = [string]$dados.GetValue($i, 4)
        $core = if ($lote.Length -ge 9) { $lote.Substring(3, 6) } else { '' }
        $vistos[$r] = $true
        [void]$ordem.Add([pscustomobject]@{ RUN = $r; Serial = $serial; Lote = $core })
    }
}

# Escrita celula a celula, com .Value e valor ja tipado.
# Tentei os dois atalhos e nenhum vinga no PowerShell 5.1: .Value2 com escalar
# numerico escolhe a sobrecarga de String e estoura InvalidCast, e a atribuicao
# de um object[,] a um Range multi-celula nao vincula (nao lanca erro e nao
# grava nada). Sao 25 corridas x 4 campos -- o custo e irrelevante e o
# comportamento e previsivel.
$lista = @($ordem | Sort-Object RUN)
if ($lista.Count -gt 0) {
    $cor.Range($cor.Cells($R0, 5), $cor.Cells($R0 + $lista.Count - 1, 5)).NumberFormat = '@'
    for ($k = 0; $k -lt $lista.Count; $k++) {
        $lin = $R0 + $k
        $vRun = [int]$lista[$k].RUN
        $cor.Cells($lin, 1).Formula = [string]$vRun
        if ($lista[$k].Serial -ne $null) {
            # .Formula com STRING, nao .Value/.Value2 com numero ou DateTime.
            # Neste ambiente o binding COM de valor numerico e erratico: dentro
            # de um .ps1 lanca InvalidCast, no mesmo codigo em bloco inline
            # grava normalmente. String sempre funcionou. O Excel interpreta o
            # texto como numero e o formato de data cuida da exibicao.
            # InvariantCulture e obrigatorio: em pt-BR a virgula decimal
            # transformaria o serial noutro numero.
            $vData = ([double]$lista[$k].Serial).ToString([System.Globalization.CultureInfo]::InvariantCulture)
            $cor.Cells($lin, 2).Formula = $vData
        }
        $cor.Cells($lin, 5).Value = [string]$lista[$k].Lote
        $cor.Cells($lin, 9).Value = 'migracao'

        # Confere a gravacao da linha. Este ambiente ja engoliu escrita COM tres
        # vezes sem lancar erro; sem a leitura de volta, a migracao mente.
        if ([string]$cor.Cells($lin, 1).Value2 -eq '') { throw "RUN nao gravado na linha $lin" }
        if ($lista[$k].Serial -ne $null -and [string]$cor.Cells($lin, 2).Value2 -eq '') {
            throw "Data nao gravada na linha $lin (RUN $vRun)"
        }
    }
    $cor.Range($cor.Cells($R0, 2), $cor.Cells($R0 + $lista.Count - 1, 2)).NumberFormat = 'dd/mm/yyyy'
}

$maior = 0
if ($ordem.Count -gt 0) { $maior = [int]($ordem | Measure-Object RUN -Maximum).Maximum }
$prox = [int]($maior + 1)
# .Value com valor ja tipado, nao .Value2 com expressao: o binding COM do
# PowerShell 5.1 escolhe a sobrecarga de String e estoura InvalidCast.
$cor.Range('D1').Formula = [string]$prox

$wb.Names.Add('proxRUN', '=Corridas!$D$1') | Out-Null

$cor.Columns('A:I').AutoFit() | Out-Null
$cor.Visible = 0        # oculta, mas nao very hidden: e registro auditavel

# Conferencia de que a gravacao ocorreu de fato. Sem isto, uma falha de binding
# COM deixa a aba vazia e o script segue relatando sucesso -- foi o que
# aconteceu na primeira versao, e so apareceu no teste de comportamento.
$ultGravada = $cor.Cells($cor.Rows.Count, 1).End(-4162).Row
$gravadas = $ultGravada - $R0 + 1
if ($gravadas -ne $ordem.Count) {
    throw "Migracao inconsistente: esperadas $($ordem.Count) corridas, gravadas $gravadas."
}
if ([int]$cor.Range('D1').Value2 -ne $prox) {
    throw "Contador proxRUN nao gravado: esperado $prox, encontrado '$($cor.Range('D1').Value2)'."
}

"corridas migradas : $($ordem.Count)  (conferido: $gravadas linhas)"
"maior RUN         : $maior"
"proxRUN           : $prox"

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"Salvo: $Workbook"
