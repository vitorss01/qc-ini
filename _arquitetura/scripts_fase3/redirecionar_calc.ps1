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
    [string]$OutCsv,
    # Niveis do produto: Hematologia 3, Bioquimica e Imunologia 2. Tem de
    # casar com a constante NLV de mEstatistica.bas.
    [int]$NLV = 3
)

$KC0 = 3; $NK = 180; $CF0 = 6; $NFD = 22
$EF0 = 3; $NEF = 7

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

        # Casa por ANALITO|RUN, nao por RUN. O RUN identifica a CORRIDA, que e
        # compartilhada entre analitos: casando so por ele, quando Eng_Saida
        # estava de outro analito o MATCH encontrava a linha e devolvia o
        # veredicto DAQUELE analito. Com a chave composta o MATCH falha e o
        # IFERROR devolve vazio -- erro visivel, nunca veredicto falso.
        $vazio = if ($ehVeredicto) { '""' } else { '0' }
        $chave = 'selAnalito&"|"&RC2'
        $f = "=IF(RC$colValor=" + '""' + ",$vazio,IFERROR(INDEX(engDados,MATCH($chave,engChave,0),$colEng),$vazio))"

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
