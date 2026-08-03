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

# Bloco auxiliar a direita dos blocos de nivel (Marco 3): permite que
# AtualizarPainelEng leia tudo de Eng_Saida e nunca do Calc. Sem isso o motor
# leria colunas do Calc que hoje sao formulas apontando para ca -- e passaria a
# depender de o Excel ter recalculado antes, o que nao se pode garantir.
$COL_FILTRO = 24                    # X: 1 se a corrida passa no filtro de data
$COL_VALOR0 = 25                    # Y,Z,AA: valor por nivel

# AB: chave logica "ANALITO|RUN".
# O RUN sozinho NAO identifica a saida do motor: ele identifica a CORRIDA, e a
# mesma corrida cobre varios analitos. Casar so por RUN fazia o Calc encontrar a
# linha certa do analito ERRADO quando Eng_Saida estava de outro analito, e
# exibir o veredicto dele. Com a chave composta o MATCH simplesmente falha, e a
# formula devolve vazio em vez de um veredicto falso.
$COL_CHAVE = 28

# Bloco de estatistica por nivel (Marco 3), espelhando as colunas do Painel:
#   linha 184 cabecalho, linhas 185..187 = niveis 1..3
#   colunas B..U == colunas 2..21 do Painel (n, media, dp, cv, etp, bias, et,
#   sigma, veredicto, [K,L vazias], cnt1..5, total, ultViol, classif, historico)
$LINHA_STAT = 185

# Bloco da tabela de parametros da aba Estatistica (Marco 4):
#   linha 189 cabecalho, linhas 190..309 = 40 analitos x 3 niveis
#   colunas C..M == colunas 3..13 da aba Estatistica (bloco C7:M126)
$LINHA_EST = 190
$LINHAS_EST = 120

$camposStat = @{
    2 = 'n'; 3 = 'media'; 4 = 'dp'; 5 = 'cv'; 6 = 'etp'; 7 = 'bias'; 8 = 'et'
    9 = 'sigma'; 10 = 'veredicto'; 13 = 'cnt13s'; 14 = 'cnt22s'; 15 = 'cntR4s'
    16 = 'cnt41s'; 17 = 'cnt10x'; 18 = 'totalViol'; 19 = 'ultViolacao'
    20 = 'classificacao'; 21 = 'historico'
}

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
$eng.Cells(2, $COL_FILTRO).Value2 = 'Filtro'
for ($t = 0; $t -lt $NLV; $t++) {
    $eng.Cells(2, $COL_VALOR0 + $t).Value2 = "N$($t + 1)_Valor"
}

$eng.Cells(2, $COL_CHAVE).Value2 = 'Chave'
$ultimaCol = $COL_CHAVE
$eng.Range($eng.Cells(2, 1), $eng.Cells(2, $ultimaCol)).Font.Bold = $true

# ---- bloco de estatistica por nivel ----
$eng.Cells($LINHA_STAT - 1, 1).Value2 = 'Nivel'
foreach ($c in $camposStat.Keys) { $eng.Cells($LINHA_STAT - 1, $c).Value2 = $camposStat[$c] }
$eng.Range($eng.Cells($LINHA_STAT - 1, 1), $eng.Cells($LINHA_STAT - 1, 21)).Font.Bold = $true
# variavel intermediaria de proposito: a forma inline ".Value2 = $t + 1" faz o
# PowerShell 5.1 tentar converter Int32 em String e falhar no COM.
for ($t = 0; $t -lt $NLV; $t++) {
    $nivel = $t + 1
    $eng.Cells($LINHA_STAT + $t, 1).Value2 = $nivel
}

# ---- bloco da tabela de parametros (aba Estatistica) ----
$camposEst = @('n', 'media', 'dp', 'cv', 'etp', 'et_desej', 'et_otim', 'bias',
    'et', 'sigma', 'classificacao')
$eng.Cells($LINHA_EST - 1, 1).Value2 = 'Linha'
for ($k = 0; $k -lt $camposEst.Count; $k++) {
    $eng.Cells($LINHA_EST - 1, 3 + $k).Value2 = $camposEst[$k]
}
$eng.Range($eng.Cells($LINHA_EST - 1, 1), $eng.Cells($LINHA_EST - 1, 13)).Font.Bold = $true
for ($r = 0; $r -lt $LINHAS_EST; $r++) {
    $ordem = $r + 1
    $eng.Cells($LINHA_EST + $r, 1).Value2 = $ordem
}
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
# engRUN: chave de juncao. O Calc casa por RUN, nao por posicao de linha --
# a coluna B do Calc filtra por rFirst/loteAtivo e o motor filtra por
# EhElegivel/lote core, entao as duas listas podem divergir de ordem.
$refRun = "=Eng_Saida!" + $eng.Range($eng.Cells($KC0, 2), $eng.Cells($ultimaLinha, 2)).Address()
$wb.Names.Add('engRUN', $refRun) | Out-Null
# engChave: "ANALITO|RUN". E por ela que o Calc casa, nao por engRUN.
$refChave = "=Eng_Saida!" + $eng.Range($eng.Cells($KC0, $COL_CHAVE), $eng.Cells($ultimaLinha, $COL_CHAVE)).Address()
$wb.Names.Add('engChave', $refChave) | Out-Null
# engPainel: bloco de estatistica por nivel. O Painel le por INDEX(engPainel,
# nivel, coluna) usando a MESMA coluna que ele proprio ocupa -- a linha 185+t
# espelha a linha 7+t do Painel, coluna a coluna.
$refStat = "=Eng_Saida!" + $eng.Range($eng.Cells($LINHA_STAT, 1), $eng.Cells($LINHA_STAT + $NLV - 1, 21)).Address()
$wb.Names.Add('engPainel', $refStat) | Out-Null
# engEstat: comeca na coluna A pelo mesmo motivo de engPainel -- assim o indice
# de coluna do INDEX e a propria coluna da aba Estatistica.
$refEst = "=Eng_Saida!" + $eng.Range($eng.Cells($LINHA_EST, 1), $eng.Cells($LINHA_EST + $LINHAS_EST - 1, 13)).Address()
$wb.Names.Add('engEstat', $refEst) | Out-Null
$wb.Names.Add('engAnalito', '=Eng_Saida!' + $eng.Range('C1').Address()) | Out-Null
$wb.Names.Add('engLote', '=Eng_Saida!' + $eng.Range('E1').Address()) | Out-Null
$wb.Names.Add('engCarimbo', '=Eng_Saida!' + $eng.Range('G1').Address()) | Out-Null
$wb.Names.Add('engNRun', '=Eng_Saida!' + $eng.Range('I1').Address()) | Out-Null

# ---- aparencia: e camada tecnica, nao vitrine ----
$eng.Columns('A').ColumnWidth = 6
$eng.Columns('B').ColumnWidth = 8
$eng.Visible = 2      # xlSheetVeryHidden: nao aparece nem no menu de reexibir

"Eng_Saida criada:"
"  corridas   linhas $KC0..$ultimaLinha, colunas 1..$ultimaCol"
"  estatistica linhas $LINHA_STAT..$($LINHA_STAT + $NLV - 1), colunas 1..21"
"  parametros  linhas $LINHA_EST..$($LINHA_EST + $LINHAS_EST - 1), colunas 1..13"
"Nomes: engDados, engRUN, engChave, engPainel, engEstat, engAnalito, engLote, engCarimbo, engNRun"

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"Salvo: $Workbook"
