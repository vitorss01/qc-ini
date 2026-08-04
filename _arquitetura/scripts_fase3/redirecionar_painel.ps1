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
    [string]$OutCsv,
    # Niveis do produto: Hematologia 3, Bioquimica e Imunologia 2. Tem de
    # casar com a constante NLV de mEstatistica.bas.
    [int]$NLV = 3
)

$LINHA_PAINEL = 7      # Painel: 1a linha de nivel

# G(7), H(8) e I(9) ficam DE FORA do redirecionamento, por decisao do gestor:
#   G = bias, que vem de Estatistica!J e portanto do controle EXTERNO (EQC_Dados).
#       O bias do motor e outra grandeza -- desvio contra o alvo do lote --, e
#       para avaliacao de metodo a trueness se estima por comparacao
#       interlaboratorial, nao pela media atribuida do proprio controle interno.
#   H = Erro Total e I = Sigma derivam de G. Continuam formula, e o defeito real
#       delas (falta de Abs no bias) e corrigido no lugar, mais abaixo.
$colunas = @(2, 3, 4, 5, 6, 10, 13, 14, 15, 16, 17, 18, 19, 20, 21)

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
$xl.AutomationSecurity = 3

$wb = $xl.Workbooks.Open($Workbook)

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


# Somente leitura significa que OUTRA instancia do Excel ainda segura o arquivo.
# Sem esta guarda o script grava no vazio e reporta sucesso: DisplayAlerts=$false
# suprime o aviso do Excel, e o Save falha em silencio.
if ($wb.ReadOnly) {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    throw "Arquivo aberto em SOMENTE LEITURA (outra instancia do Excel o mantem travado): $Workbook"
}

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

# --- correcao no lugar: Erro Total e Sigma sem Abs no bias ---
# CLSI: TE = |bias| + 1,65*CV  e  Sigma = (ETp - |bias|)/CV.
# Sem o Abs, bias negativo (media abaixo do alvo) subestimava o Erro Total e
# superestimava o Sigma -- o metodo aparentava desempenho melhor do que tem.
# O bias (RC7) continua vindo de Estatistica!J, ou seja, do EQC.
for ($t = 0; $t -lt $NLV; $t++) {
    $linha = $LINHA_PAINEL + $t
    $pnl.Cells($linha, 8).FormulaR1C1 = '=IF(RC5="","",RC5*1.65+ABS(IF(ISNUMBER(RC7),RC7,0)))'
    $pnl.Cells($linha, 9).FormulaR1C1 = '=IF(OR(RC5="",RC5=0,RC6=""),"",(RC6-ABS(IF(ISNUMBER(RC7),RC7,0)))/RC5)'
    foreach ($c in @(8, 9)) {
        [void]$lista.Add([pscustomobject]@{
                Nivel  = $t + 1
                Celula = $pnl.Cells($linha, $c).Address($false, $false)
                Coluna = $c
                Tipo   = 'ALTERADA'
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
