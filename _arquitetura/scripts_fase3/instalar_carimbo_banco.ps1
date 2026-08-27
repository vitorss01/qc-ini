# instalar_carimbo_banco.ps1 - ADR-049: o carimbo do banco quem mantem e o Excel
#
# O DEFEITO
#
# mEstatPeriodo guarda os agregados em mAgg e decide reaproveitar pelo Carimbo,
# cuja unica impressao digital do DADO era a ULTIMA LINHA de DB_Resultados.
#
# Inclusao nova muda a ultima linha e invalida o cache. Edicao EM LINHA nao:
# marcar um resultado como nao conforme (mRegistros grava Status na propria
# linha) mantem a ultima linha, o carimbo fica identico e a Estatistica segue
# contando o resultado excluido. Medido, com o cache quente:
#
#   n inicial ................................ 25
#   apos mudar Status para Excluido .......... 25
#   apos Application.Calculate (F9) .......... 25
#   apos CalculateFullRebuild ................ 25
#
# E dai para media, DP, CV, Sigma, classificacao e plano de CQ.
#
# POR QUE UMA CELULA, E NAO UMA VARREDURA NA UDF
#
# Medido nesta base (6.808 resultados, 320 chamadas por recalculo):
#
#   cache quente, como esta hoje ............. 0,199 s
#   + agregados numericos dentro da UDF ...... 0,195 s  (dobra)
#   + agregados de texto dentro da UDF ....... 0,645 s  (triplica)
#
# Numa celula, o Excel recalcula os agregados UMA vez, e so quando o intervalo
# muda -- e a UDF le uma celula. O custo por chamada vira desprezivel.
#
# POR QUE NAO UM CONTADOR BUMPADO PELOS ESCRITORES
#
# Seis modulos escrevem em DB_Resultados. Um contador central exigiria que TODO
# caminho de escrita lembrasse de incrementa-lo, inclusive os que rodam com
# EnableEvents = False, e nao cobriria edicao manual na aba. O Excel cobre tudo
# isso de graca, porque e ele quem mantem o grafo de calculo.
#
# O QUE O CARIMBO DETECTA
#
#   inclusao / exclusao de linha ....... contagem e ultima linha
#   mudanca de Status .................. CONT.SE(status; "Ativo")
#   mudanca de resultado ............... SOMA(valor)
#   mudanca de nivel ................... SOMA(nivel)
#   mudanca de data .................... SOMA(data) e CONT.NUM(data)
#   mudanca de analito ................. SOMARPRODUTO(NUM.CARACT(analito))
#   mudanca de lote .................... SOMARPRODUTO(NUM.CARACT(lote))
#
# O que NAO detecta: uma troca de texto que preserve o comprimento total E
# todas as somas -- por exemplo dois analitos trocados entre si na mesma
# coluna. E o limite conhecido desta abordagem, e esta escrito aqui para nao
# virar surpresa.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso: .\instalar_carimbo_banco.ps1 -Workbook <arquivo.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$SENHA = 'qcini2025'
$ABA = 'DB_Resultados'
$NOME = 'DB_Carimbo'

# Colunas de DB_Resultados (as mesmas constantes EP_C_* de mEstatPeriodo)
$C_RUN = 1; $C_DATA = 2; $C_NIVEL = 3; $C_LOTE = 4; $C_ANALITO = 5
$C_VALOR = 6; $C_STATUS = 7
$R0 = 4
# Teto generoso e FIXO: o SOMARPRODUTO so recalcula quando o intervalo muda,
# entao a folga nao custa por chamada. Coluna inteira, nao -- SOMARPRODUTO
# sobre 1.048.576 linhas pesa de verdade.
$RN = 200004

function Novo-Excel {
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch { Start-Sleep -Seconds (1 + $t) }
    }
    throw 'nao consegui criar a instancia do Excel'
}

function Faixa([int]$col) {
    $L = [char](64 + $col)
    return "$ABA!`$$L`$$R0`:`$$L`$$RN"
}

$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvou = $false
try {
    $ws = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq $ABA) { $ws = $w; break } }
    if ($ws -eq $null) { throw "aba $ABA ausente" }

    # Celula do carimbo: primeira coluna livre a direita do bloco de dados.
    $col = $C_STATUS + 3
    $cel = $ws.Cells(1, $col)

    $estrutura = [bool]$wb.ProtectStructure
    if ($estrutura) { $wb.Unprotect($SENHA) }
    $protegida = [bool]$ws.ProtectContents
    if ($protegida) { $ws.Unprotect($SENHA) }

    # A formula usa nomes de funcao em INGLES porque .Formula fala ingles,
    # independentemente do idioma do Excel. .FormulaLocal falaria portugues e
    # quebraria numa instalacao em outro idioma.
    $f = '=COUNTIF(' + (Faixa $C_STATUS) + ',"Ativo")&"|"' +
         '&ROUND(SUM(' + (Faixa $C_VALOR) + '),6)&"|"' +
         '&SUM(' + (Faixa $C_NIVEL) + ')&"|"' +
         '&ROUND(SUM(' + (Faixa $C_DATA) + '),6)&"|"' +
         '&COUNT(' + (Faixa $C_DATA) + ')&"|"' +
         '&SUMPRODUCT(LEN(' + (Faixa $C_ANALITO) + '))&"|"' +
         '&SUMPRODUCT(LEN(' + (Faixa $C_LOTE) + '))&"|"' +
         '&COUNTA(' + (Faixa $C_RUN) + ')'

    $cel.Formula = $f
    $ws.Cells(2, $col).Value2 = 'CARIMBO DO BANCO (ADR-049) - nao apagar'
    $ws.Columns($col).Hidden = $true

    # Nome definido: mEstatPeriodo le pelo NOME, nunca por endereco. Mover a
    # coluna deixaria de quebrar o modulo em silencio.
    $existe = $false
    foreach ($n in $wb.Names) { if ($n.Name -eq $NOME) { $existe = $true; break } }
    if ($existe) { $wb.Names($NOME).Delete() }
    $L = [char](64 + $col)
    $wb.Names.Add($NOME, "=$ABA!`$$L`$1") | Out-Null

    $xl.CalculateFullRebuild()
    $v = [string]$cel.Value2
    if ($v -eq '' -or $v -like '*#*') { throw "carimbo nao avaliou: '$v'" }
    "carimbo instalado em $ABA coluna $L (oculta): $v"

    if ($protegida) {
        $ws.Protect($SENHA, $false, $true, $true, $true)
    }
    if ($estrutura) { $wb.Protect($SENHA, $true, $false) }

    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { $wb.Close($salvou) } catch { }
    try { $xl.Quit() } catch { }
}
