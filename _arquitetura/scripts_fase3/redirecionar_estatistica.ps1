# redirecionar_estatistica.ps1 - Marco 4 do Sprint HARDENING 1
#
# A aba Estatistica deixa de calcular. O bloco C7:M126 -- 11 colunas x 120
# linhas = 1.320 celulas, exatamente a area que o motor sobrescrevia -- passa a
# ler Eng_Saida por INDEX(engEstat, linha, coluna).
#
# A linha 190+r de Eng_Saida espelha a linha 7+r da Estatistica, e engEstat
# comeca na coluna A, entao o indice de coluna do INDEX e a propria coluna da
# aba. Sem conversao, sem tabela de-para.
#
# Guarda IF(ref="","",ref) pelo mesmo motivo do Painel: INDEX sobre celula vazia
# devolve 0, e o motor grava "" para analito ausente ou n insuficiente. Sem o
# guarda, 120 linhas exibiriam 0 onde a producao exibia vazio.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\redirecionar_estatistica.ps1 -Workbook <build.xlsm> [-OutCsv <lista.csv>]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$OutCsv,
    # Niveis do produto: Hematologia 3, Bioquimica e Imunologia 2. Tem de
    # casar com a constante NLV de mEstatistica.bas.
    [int]$NLV = 3
)

# 1a linha do bloco da Estatistica.
#
# NAO PODE SER FIXA. O cabecalho de filtros cresceu em 06/08/2026 e a tabela
# desceu da linha 7 para a 14; com o numero cravado aqui, o build passou a
# escrever os redirecionamentos DENTRO do cabecalho -- 60 formulas em cima dos
# campos de filtro e a tabela real intocada. Localizar pela coluna A (que traz
# =Analitos!$A$n) amarra o script ao CONTEUDO e nao a uma coordenada.
$E0 = 0          # localizado apos abrir a pasta (ver abaixo)
$LINHAS = 40 * $NLV    # 40 analitos x niveis
$C0 = 3          # coluna C
$CN = 9          # coluna I -- para em I de proposito, ver abaixo

# J(10) a M(13) ficam DE FORA do redirecionamento, por decisao do gestor:
#   J = "Bias anual %" = AVERAGE(P:U), e P..U sao AVERAGEIFS sobre EQC_Dados.
#       Esse bias e medido contra o programa de comparacao interlaboratorial.
#       O bias do motor e desvio contra o alvo do lote -- outra grandeza. Para
#       avaliacao de metodo, trueness se estima por EQA, nao pela media
#       atribuida do proprio controle interno.
#   K = Erro Total, L = Sigma e M = Classificacao derivam de J. Continuam
#       formula; o defeito real de K e L (falta de Abs no bias) e corrigido no
#       lugar, mais abaixo.

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

$est = $null
foreach ($ws in $wb.Worksheets) { if ($ws.Name -like 'Estat*') { $est = $ws; break } }
# Agora que a aba existe, localiza a 1a linha do bloco PELO CONTEUDO.
for ($i = 3; $i -le 60; $i++) {
    $f = $est.Cells.Item($i, 1).Formula
    if ("$f" -like '=Analitos!*') { $E0 = $i; break }
}
if ($E0 -eq 0) { throw "bloco da Estatistica nao localizado (nenhuma celula da coluna A com =Analitos!)" }
"Estatistica: bloco comeca na linha $E0"

if ($est -eq $null) { throw "Aba Estatistica nao encontrada" }

$lista = New-Object System.Collections.ArrayList

for ($c = $C0; $c -le $CN; $c++) {
    $ref = "INDEX(engEstat,ROW()-$($E0 - 1),$c)"
    $f = "=IF($ref=" + '""' + "," + '""' + ",$ref)"
    $rng = $est.Range($est.Cells($E0, $c), $est.Cells($E0 + $LINHAS - 1, $c))
    $comFormula = 0
    foreach ($cel in $rng) { if ($cel.HasFormula) { $comFormula++ } }
    $rng.FormulaR1C1 = $f
    [void]$lista.Add([pscustomobject]@{
            Coluna       = $c
            Intervalo    = $rng.Address()
            EramFormula  = $comFormula
            EramConstante = $LINHAS - $comFormula
            Formula      = $f
        })
}

# --- correcao no lugar: Erro Total (K) e Sigma (L) sem Abs no bias (J) ---
# CLSI: TE = |bias| + 1,65*CV  e  Sigma = (ETp - |bias|)/CV.
# F=CV, G=ETp, J=bias(EQC). A classificacao M nao muda de formula: ela le L.
$est.Range($est.Cells($E0, 11), $est.Cells($E0 + $LINHAS - 1, 11)).FormulaR1C1 = `
    '=IF(OR(RC6="",RC10=""),"",RC6*1.65+ABS(RC10))'
$est.Range($est.Cells($E0, 12), $est.Cells($E0 + $LINHAS - 1, 12)).FormulaR1C1 = `
    '=IF(OR(RC6="",RC6=0,RC7="",RC10=""),"",(RC7-ABS(RC10))/RC6)'
[void]$lista.Add([pscustomobject]@{
        Coluna = 11; Intervalo = $est.Range($est.Cells($E0, 11), $est.Cells($E0 + $LINHAS - 1, 11)).Address()
        EramFormula = $LINHAS; EramConstante = 0; Formula = 'ABS no bias (Erro Total)'
    })
[void]$lista.Add([pscustomobject]@{
        Coluna = 12; Intervalo = $est.Range($est.Cells($E0, 12), $est.Cells($E0 + $LINHAS - 1, 12)).Address()
        EramFormula = $LINHAS; EramConstante = 0; Formula = 'ABS no bias (Sigma)'
    })

$xl.Calculation = -4105
$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

$totF = ($lista | Measure-Object EramFormula -Sum).Sum
$totC = ($lista | Measure-Object EramConstante -Sum).Sum
"celulas tocadas: $totF (redirecionadas C..I + corrigidas K,L)"
"  ALTERADA (eram formula)   n=$totF"
"  EXTRA    (eram constante) n=$totC"
$lista | Format-Table Coluna, Intervalo, EramFormula, EramConstante -AutoSize
if ($OutCsv) {
    $lista | Export-Csv -Path $OutCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    "lista fechada em: $OutCsv"
}
