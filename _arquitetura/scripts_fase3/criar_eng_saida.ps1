# criar_eng_saida.ps1 - Marco 1 do Sprint HARDENING 1
#
# Cria a aba Eng_Saida na pasta de trabalho indicada: a camada de saida do motor
# estatistico (ADR-019). O motor passa a escrever aqui; Painel, Estatistica e Calc
# passam a LER daqui por formula.
#
# Layout (espelha as linhas do Calc para permitir referencia linha a linha):
#   linha 1  metadados: analito, lote, carimbo de tempo, nRun
#   linha 2  cabecalhos
#   linha 3..182  uma linha por slot de RUN (KC0=3, NK=180, igual ao Calc)
#
#   col A  slot (1..180)
#   col B  RUN
#   por nivel t (0..NLV-1), bloco de 7 colunas a partir de EF0 = 3:
#     +0 R13s   +1 R22s   +2 RR4s   +3 R41s   +4 R10x   +5 Alerta12s   +6 Veredicto
#
# Sobre o Veredicto: guarda apenas REJEITADO/OK, e o alerta 12s fica na coluna
# propria. O motor internamente produz REJEITADO/ALERTA/OK, mas o Calc usa esse
# campo em RC38="OK" para separar pontos plotados como conformes e como rejeitados
# (colunas AE/AF). Publicar "ALERTA" ali faria pontos de alerta migrarem para a
# serie de rejeitados e mudaria o grafico. Manter dois campos preserva o
# comportamento atual e nao perde a informacao.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_eng_saida.ps1 -Workbook <copia_de_trabalho.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

# constantes espelhadas de mEstatistica.bas
$NLV = 3      # niveis
$KC0 = 3      # primeira linha de dados (igual ao Calc)
$NK = 180     # slots de RUN
$EF0 = 3      # primeira coluna de bloco de nivel
$NEF = 7      # campos por nivel

$campos = @('R13s', 'R22s', 'RR4s', 'R41s', 'R10x', 'Alerta12s', 'Veredicto')

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3

$wb = $xl.Workbooks.Open($Workbook)

# ---- remove versao anterior, se houver (idempotencia) ----
foreach ($ws in $wb.Worksheets) {
    if ($ws.Name -eq 'Eng_Saida') {
        $ws.Visible = -1
        $ws.Delete()
        "Eng_Saida anterior removida"
        break
    }
}

$eng = $wb.Worksheets.Add()
$eng.Name = 'Eng_Saida'

# ---- linha 1: metadados ----
$eng.Range('A1').Value2 = 'Eng_Saida'
$eng.Range('B1').Value2 = 'analito:'
$eng.Range('D1').Value2 = 'lote:'
$eng.Range('F1').Value2 = 'gerado em:'
$eng.Range('H1').Value2 = 'nRun:'
$eng.Range('A1:H1').Font.Bold = $true

# ---- linha 2: cabecalhos ----
$eng.Cells(2, 1).Value2 = 'Slot'
$eng.Cells(2, 2).Value2 = 'RUN'
for ($t = 0; $t -lt $NLV; $t++) {
    for ($k = 0; $k -lt $NEF; $k++) {
        $col = $EF0 + $t * $NEF + $k
        $eng.Cells(2, $col).Value2 = "N$($t + 1)_$($campos[$k])"
    }
}
$ultimaCol = $EF0 + $NLV * $NEF - 1
$eng.Range($eng.Cells(2, 1), $eng.Cells(2, $ultimaCol)).Font.Bold = $true
# sem AutoFilter de proposito: criaria o nome definido parasita
# Eng_Saida!_FilterDatabase, que poluiria o inventario de nomes e o diff.

# ---- coluna A: slots ----
for ($i = 1; $i -le $NK; $i++) {
    $eng.Cells($KC0 + $i - 1, 1).Value2 = $i
}

# ---- nomes definidos: o VBA e as formulas referenciam por nome, nao por endereco ----
$ultimaLinha = $KC0 + $NK - 1
$refDados = "=Eng_Saida!" + $eng.Range($eng.Cells($KC0, 1), $eng.Cells($ultimaLinha, $ultimaCol)).Address()
$wb.Names.Add('engDados', $refDados) | Out-Null
$wb.Names.Add('engAnalito', '=Eng_Saida!' + $eng.Range('C1').Address()) | Out-Null
$wb.Names.Add('engLote', '=Eng_Saida!' + $eng.Range('E1').Address()) | Out-Null
$wb.Names.Add('engCarimbo', '=Eng_Saida!' + $eng.Range('G1').Address()) | Out-Null
$wb.Names.Add('engNRun', '=Eng_Saida!' + $eng.Range('I1').Address()) | Out-Null

# ---- aparencia: e camada tecnica, nao vitrine ----
$eng.Columns('A').ColumnWidth = 6
$eng.Columns('B').ColumnWidth = 8
$eng.Visible = 2      # xlSheetVeryHidden: nao aparece nem no menu de reexibir

"Eng_Saida criada: linhas $KC0..$ultimaLinha, colunas 1..$ultimaCol ($($NLV * $NEF) campos, $NLV niveis)"
"CodeName: $($eng.CodeName)"
"Nomes definidos: engDados, engAnalito, engLote, engCarimbo, engNRun"

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"Salvo: $Workbook"
