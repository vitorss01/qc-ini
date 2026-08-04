# consolidar_registros_nc.ps1 - Rep 1/2/3 viram UMA coluna de nao conformidade
#
# A Sprint NC ja gravava a ocorrencia na coluna F ("Rep 1") e o grafico ja
# plotava essa coluna. O que restava era estrutural e de rotulo: a aba anunciava
# tres colunas de repeticao que ninguem mais alimenta, e o marcador do grafico
# dizia "Repeticao" para um dado que agora significa nao conformidade. Num
# sistema de qualidade, rotulo que mente e defeito.
#
# O QUE MUDA
#   Registros  E  "Corrida (Seq)" -> "RUN"
#              F  "Rep 1"         -> "Resultado NC"
#              G  "Rep 2"         -> removida (rotulo e dados)
#              H  "Rep 3"         -> removida (rotulo e dados)
#              J  "Observacao"    -> "Parecer Tecnico"
#              M  "Status"        -> "Tipo de NC"
#   Calc       540 formulas mantidas  (X / AT / BP  = regRep1 = coluna F)
#            1.080 formulas removidas (Y,Z / AU,AV / BQ,BR = regRep2 e regRep3)
#   Graficos   3 x 2 series "Repeticao" removidas; a que sobra vira
#              "Nao conforme"
#   Nomes      regRep2 e regRep3 excluidos
#
# O QUE NAO MUDA
#   mRegistros.bas nao precisa de alteracao: ja escreve so em F (RG_RESULT).
#   A serie que sobra aponta para as MESMAS celulas de antes.
#
# BLOQUEIO DA DIGITACAO MANUAL (item do handoff)
#   As celulas de dados da Registros estavam com Locked = False -- por isso a
#   protecao aplicada por blindar_artefato.ps1 nao impedia digitacao. Aqui elas
#   passam a Locked = True. MarcarNaoConforme e ExcluirRegistroNC ja fazem
#   Unprotect/Protect com a senha, entao o fluxo pelo formulario continua.
#
# ORDEM NO PIPELINE: depois de redirecionar_calc.ps1 e ANTES de
# blindar_artefato.ps1 -- com a aba protegida, as escritas daqui falhariam.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\consolidar_registros_nc.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    # Niveis do produto: Hematologia 3, Bioquimica e Imunologia 2. Tem de
    # casar com a constante NLV de mEstatistica.bas.
    [int]$NLV = 3
)

$ErrorActionPreference = 'Stop'

# acentos por codigo: o arquivo e lido como ANSI e literais acentuados chegam
# corrompidos ao Excel
$TXT_PARECER = 'Parecer T' + [char]0x00E9 + 'cnico'
$TXT_NAOCONF = 'N' + [char]0x00E3 + 'o conforme'

$REG_R0 = 4; $REG_RN = 203
$CALC_R0 = 3; $CALC_RN = 182

# Colunas do Calc que espelham regRep2 e regRep3, DERIVADAS da geometria em vez
# de escritas na mao: o bloco do nivel t comeca em CF0 + t*NFD, e regRep1/2/3
# sao os campos 19, 20 e 21 do bloco. Com NLV=3 isso da Y,Z / AU,AV / BQ,BR --
# os mesmos valores de antes, agora corretos tambem para um produto de 2 niveis.
$CF0 = 6; $NFD = 22
$COLS_REMOVER = @()
for ($t = 0; $t -lt $NLV; $t++) {
    $COLS_REMOVER += ($CF0 + $t * $NFD + 19)     # regRep2
    $COLS_REMOVER += ($CF0 + $t * $NFD + 20)     # regRep3
}
# indice da serie de repeticao em cada grafico (as tres sao consecutivas)
$SERIE_BASE = 11

function Novo-Excel {
    $ultimo = $null
    for ($tentativa = 1; $tentativa -le 6; $tentativa++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $ultimo = $_
            if ($tentativa -eq 2) {
                try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep -Seconds 5 } catch { }
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
try { $wb.EnableAutoRecover = $false } catch { }

if ($wb.ReadOnly) {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    throw "Arquivo aberto em SOMENTE LEITURA (outra instancia do Excel o segura): $Workbook"
}

$reg = $wb.Worksheets('Registros')
$calc = $wb.Worksheets('Calc')
try { $reg.Unprotect('qcini2025') } catch { }
try { $calc.Unprotect('qcini2025') } catch { }

# ---------- 1. rotulos da aba Registros ----------
$reg.Cells(3, 5).Formula = 'RUN'
$reg.Cells(3, 6).Formula = 'Resultado NC'
$reg.Cells(3, 7).ClearContents()
$reg.Cells(3, 8).ClearContents()
$reg.Cells(3, 10).Formula = $TXT_PARECER
$reg.Cells(3, 13).Formula = 'Tipo de NC'
"rotulos: E=RUN  F=Resultado NC  G,H removidos  J=Parecer  M=Tipo de NC"

# ---------- 2. dados residuais de Rep 2 e Rep 3 ----------
$reg.Range($reg.Cells($REG_R0, 7), $reg.Cells($REG_RN, 8)).ClearContents()

# ---------- 3. formulas do Calc de regRep2 e regRep3 ----------
$limpas = 0
foreach ($c in $COLS_REMOVER) {
    $rng = $calc.Range($calc.Cells($CALC_R0, $c), $calc.Cells($CALC_RN, $c))
    $limpas += $rng.Count
    $rng.ClearContents()
}
"formulas do Calc removidas: $limpas"

# ---------- 4. series dos graficos ----------
# Remover de tras para frente: excluir a serie 12 primeiro renumeraria a 13.
$pnl = $wb.Worksheets('Painel')
$seriesRemovidas = 0
$graficos = 0
foreach ($co in $pnl.ChartObjects()) {
    $ch = $co.Chart
    $total = $ch.SeriesCollection().Count
    if ($total -ge ($SERIE_BASE + 2)) {
        $ch.SeriesCollection($SERIE_BASE + 2).Delete()
        $ch.SeriesCollection($SERIE_BASE + 1).Delete()
        $seriesRemovidas += 2
    }
    $ch.SeriesCollection($SERIE_BASE).Name = $TXT_NAOCONF
    $graficos++
}
"graficos ajustados: $graficos   series removidas: $seriesRemovidas"

# ---------- 5. nomes definidos orfaos ----------
$removidos = @()
foreach ($nome in @('regRep2', 'regRep3')) {
    foreach ($nm in $wb.Names) {
        if ($nm.Name -eq $nome) { $nm.Delete(); $removidos += $nome; break }
    }
}
"nomes removidos: $($removidos -join ', ')"

# ---------- 6. trava a digitacao manual ----------
# Sem isto a protecao da aba nao impede nada: as celulas estavam destravadas.
$reg.Range($reg.Cells($REG_R0, 2), $reg.Cells($REG_RN, 13)).Locked = $true
"celulas B$REG_R0`:M$REG_RN travadas"

# ---------- 7. conferencia ----------
$erros = @()
if ($reg.Cells(3, 6).Value2 -ne 'Resultado NC') { $erros += 'rotulo F nao gravado' }
if ("$($reg.Cells(3, 7).Value2)" -ne '') { $erros += 'rotulo G nao removido' }
if ($reg.Range('B4').Locked -ne $true) { $erros += 'celulas da Registros continuam destravadas' }
foreach ($c in $COLS_REMOVER) {
    if ($calc.Cells($CALC_R0, $c).HasFormula) { $erros += "coluna $c do Calc ainda tem formula" }
}
foreach ($co in $pnl.ChartObjects()) {
    if ($co.Chart.SeriesCollection().Count -ne 12) { $erros += "grafico $($co.Name) com $($co.Chart.SeriesCollection().Count) series (esperado 12)" }
}
if ($erros.Count -gt 0) {
    $erros | ForEach-Object { "  FALHA: $_" }
    throw "Consolidacao incompleta: $($erros.Count) verificacao(oes) falharam."
}
"conferencia: ok - 12 series por grafico, rotulos gravados, celulas travadas"

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"Salvo: $Workbook"
