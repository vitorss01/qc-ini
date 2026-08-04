# snapshot_formulas.ps1 - inventario celula a celula das formulas de uma pasta .xlsm
#
# Produz o artefato de referencia exigido pelo criterio de aceite do Sprint HARDENING 1
# (QUALITY_GATE.md, item 2.1): endereco + formula em R1C1 + valor, por aba.
# Comparar dois CSVs desses da o diff exigido, sem depender de Python.
#
# ATENCAO - ambiente Windows PowerShell 5.1:
#   Nao usar acentos neste arquivo. O PS 5.1 le .ps1 como ANSI e caminhos/nomes
#   com acento chegam corrompidos ao Excel (SOLUCAO vira SOLUCÃƒO).
#   Resolver caminhos com curinga no chamador, ex.:
#     $root = (Get-Item "...\CLONE_PLANILHA_QC__*\qc-ini").FullName
#
# ATENCAO - Excel COM:
#   1) Excel nao abre dois arquivos de mesmo nome ao mesmo tempo. Para comparar
#      producao x build corrigida (ambos QC_Hematologia.xlsm), rodar em sequencia.
#   2) Abrir um arquivo pelo COM muda o diretorio de trabalho do processo.
#      Sempre usar caminhos absolutos.
#
# Uso:
#   .\snapshot_formulas.ps1 -Workbook <caminho.xlsm> -OutCsv <saida.csv>

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [Parameter(Mandatory = $true)][string]$OutCsv
)

$ErrorActionPreference = 'Stop'

# Criar o Excel COM RESILIENCIA.
#
# O build sobe e derruba o Excel cerca de dez vezes. Sob esse ritmo o servidor
# COM as vezes recusa a proxima instancia com 0x80080005
# (CO_E_SERVER_EXEC_FAILURE) -- estado transitorio, nao defeito do script.
# Falhar na primeira tentativa jogava fora um build inteiro de varios minutos.
function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try {
            $app = New-Object -ComObject Excel.Application
            return $app
        }
        catch {
            $ultimo = $_
            # Depois de um periodo sem Excel rodando, a PRIMEIRA ativacao COM
            # costuma falhar com 0x80080005 mesmo com a maquina sadia. Lancar o
            # excel.exe uma vez levanta o servidor e as ativacoes seguintes
            # funcionam. Verificado nesta maquina: com um processo de pe, o
            # New-Object passa na hora.
            if ($tentativa -eq 2) {
                try {
                    Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null
                    Start-Sleep -Seconds 5
                }
                catch { }
            }
            Start-Sleep -Seconds ($tentativa * 2)
        }
    }
    throw "Excel COM nao subiu apos 6 tentativas: $($ultimo.Exception.Message)"
}

$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3          # msoAutomationSecurityForceDisable: nao roda macro

$wb = $xl.Workbooks.Open($Workbook, 0, $true)

# AutoRecuperacao DESLIGADA nesta copia de trabalho.
#
# O build encerra o Excel a forca varias vezes. Cada encerramento deixa um
# arquivo de recuperacao pendente; acumulados, o Excel passa a tentar exibir o
# painel "Recuperacao de Documento" ao iniciar e MORRE antes de responder --
# ate o excel.exe puro para de abrir, e a automacao falha com 0x80080005
# (CO_E_SERVER_EXEC_FAILURE), que parece defeito de COM e nao e.
#
# O artefato e reproduzivel por comando: nao ha o que recuperar aqui.
try { $wb.EnableAutoRecover = $false } catch { }


   # UpdateLinks=0, ReadOnly=$true
$xl.Calculation = -4135                          # xlCalculationManual (so vale com wb aberto)

$sw = New-Object System.IO.StreamWriter($OutCsv, $false, [System.Text.UTF8Encoding]::new($true))
$sw.WriteLine('Aba;Endereco;FormulaR1C1;Valor')

$totalGeral = 0
foreach ($ws in $wb.Worksheets) {
    $ur = $ws.UsedRange
    if ($ur.Cells.Count -eq 0) { continue }

    # leitura em bloco: 2 chamadas COM por aba, nao uma por celula
    $fr = $ur.FormulaR1C1
    $vl = $ur.Value2
    $r0 = $ur.Row; $c0 = $ur.Column
    $nR = $ur.Rows.Count; $nC = $ur.Columns.Count
    $unico = ($nR -eq 1 -and $nC -eq 1)

    $n = 0
    for ($i = 1; $i -le $nR; $i++) {
        for ($j = 1; $j -le $nC; $j++) {
            if ($unico) { $f = $fr; $v = $vl }
            else { $f = $fr.GetValue($i, $j); $v = $vl.GetValue($i, $j) }
            if ($f -isnot [string]) { continue }
            if (-not $f.StartsWith('=')) { continue }
            $addr = $ws.Cells($r0 + $i - 1, $c0 + $j - 1).Address($false, $false)
            $fs = $f -replace '[\r\n;]', ' '
            $vs = ("$v") -replace '[\r\n;]', ' '
            $sw.WriteLine("$($ws.Name);$addr;$fs;$vs")
            $n++
        }
    }
    $totalGeral += $n
    "{0,-22} {1,8}" -f $ws.Name, $n | Write-Output
}
"{0,-22} {1,8}" -f 'TOTAL', $totalGeral | Write-Output

$sw.Close()
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
