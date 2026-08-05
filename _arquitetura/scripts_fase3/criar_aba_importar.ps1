# criar_aba_importar.ps1 - substitui o frmMassa por uma aba de importacao
#
# Cria a aba "Importar" com cabecalho horizontal (Data | Nivel | Lote |
# <analitos do produto>), area de entrada liberada e um botao Registrar que
# chama mImportar.RegistrarImportacao. E troca o botao de importacao em massa da
# aba Resultados para navegar ate a aba nova.
#
# O cabecalho de analitos vem da aba Analitos do proprio produto -- Bioquimica
# tem 20, Hematologia 28. Parametrico por construcao.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_aba_importar.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'
# ATENCAO: PowerShell NAO distingue maiusculas em nome de variavel. Uma matriz
# chamada $cab e a constante $CAB seriam A MESMA variavel, e a matriz
# sobrescreveria o numero -- Cells.Item receberia um array como indice de linha
# e o COM devolveria "tipo nao correspondente". Dai o nome $cabecalho.
$CAB = 3
$R0 = 4
$RN = 203

# Escreve o rotulo de um Shape sem depender do tipo dele.
# Form control (Type 8) -> TextFrame.Characters(); AutoShape -> TextFrame2.
function Set-TextoShape {
    param($Shape, [string]$Texto)
    try { $Shape.TextFrame.Characters().Text = $Texto; return } catch { }
    $Shape.TextFrame2.TextRange.Text = $Texto
}

function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch { $u = $_; if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }; Start-Sleep -Seconds ($t * 2) }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

$salvou = $false
$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { $wb.Close($false); $xl.Quit(); throw "Somente leitura: $Workbook" }

# a trava de estrutura impede criar aba: abre a estrutura, cria, reprotege.
$estruturaEstava = $wb.ProtectStructure
if ($estruturaEstava) { $wb.Unprotect($SENHA) }

try {
    # ---- analitos do produto (aba Analitos, coluna A, linhas 4..43) ----
    $an = $wb.Worksheets.Item('Analitos')
    $analitos = @()
    for ($i = 4; $i -le 43; $i++) {
        $v = $an.Cells.Item($i, 1).Value2
        if ($v -ne $null -and "$v".Trim() -ne '') { $analitos += "$v" }
    }
    if ($analitos.Count -eq 0) { throw "nenhum analito na aba Analitos" }

    # ---- (re)cria a aba Importar ----
    foreach ($ws in @($wb.Worksheets)) { if ($ws.Name -eq 'Importar') { $ws.Visible = -1; $ws.Delete() } }
    $imp = $wb.Worksheets.Add()
    $imp.Name = 'Importar'

    $imp.Range('B1').Value2 = 'IMPORTACAO DE RESULTADOS'
    $imp.Range('B1').Font.Bold = $true
    $imp.Range('B1').Font.Size = 13
    $imp.Range('B2').Value2 = 'Cole os dados a partir da linha ' + $R0 + '. Uma corrida por linha. Ao clicar em Registrar, tudo e validado: se houver erro, NADA e gravado e as linhas com problema aparecem a direita. Sem erro, os dados migram para o banco e somem daqui.'
    $imp.Range('B2').Font.Italic = $true

    # cabecalho
    $cabecalho = @('Data', 'Nivel', 'Lote') + $analitos
    for ($k = 0; $k -lt $cabecalho.Count; $k++) {
        $c = $imp.Cells.Item($CAB, $k + 2)     # comeca em B
        $c.Value2 = $cabecalho[$k]
        $c.Font.Bold = $true
    }
    $rngCab = $imp.Range($imp.Cells.Item($CAB, 2), $imp.Cells.Item($CAB, $cabecalho.Count + 1))
    $rngCab.Interior.Color = 14277081
    $imp.Columns.Item(2).ColumnWidth = 12    # Data
    $imp.Columns.Item(3).ColumnWidth = 7     # Nivel
    $imp.Columns.Item(4).ColumnWidth = 12    # Lote

    # area de entrada como TEXTO, para 2,58 nao virar data nem numero regional torto
    $areaC0 = 2
    $areaC1 = $cabecalho.Count + 1
    $entrada = $imp.Range($imp.Cells.Item($R0, $areaC0), $imp.Cells.Item($RN, $areaC1))
    $entrada.Locked = $false                 # o usuario digita aqui
    $imp.Cells.Item($CAB - 1, $areaC1 + 2).Value2 = ''   # reserva a coluna de erro a direita

    # botao Registrar
    foreach ($sh in @($imp.Shapes)) { if ($sh.Name -eq 'btnRegistrar') { $sh.Delete() } }
    $b = $imp.Shapes.AddShape(5, 12, 4, 150, 30)   # rounded rect
    $b.Name = 'btnRegistrar'
    $b.TextFrame2.TextRange.Text = 'Registrar'
    $b.TextFrame2.TextRange.Font.Size = 11
    $b.TextFrame2.TextRange.Font.Bold = $true
    $b.Fill.ForeColor.RGB = 3506772           # verde
    $b.Line.Visible = $false
    $b.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    $b.OnAction = 'RegistrarImportacao'

    # protege a aba deixando so a area de entrada editavel
    $imp.Cells.Item($R0, $areaC0).Select() | Out-Null
    $imp.Protect($SENHA, $true, $true, $true, $true, $true, $true, $true, $true, $false, $false, $false, $false, $false)
    # AllowFormattingCells etc default false; a area Locked=$false continua editavel

    "aba Importar criada: $($analitos.Count) analitos, entrada em B$R0"

    # ---- troca o botao da aba Resultados (frmMassa -> IrParaImportar) ----
    # Os botoes antigos (btnOp0..btnOp3) sao FORM CONTROLS (Shape.Type = 8), e
    # form control NAO expoe TextFrame2.TextRange -- o texto vive em
    # TextFrame.Characters(). Os botoes que este projeto cria depois sao
    # AutoShape (Type = 1), onde vale o TextFrame2. Escrever pelo caminho errado
    # derruba o script com "A propriedade 'Text' ...". Dai o Set-TextoShape.
    $resu = $wb.Worksheets.Item('Resultados')
    $trocou = $false
    $jaEstava = $false
    foreach ($sh in @($resu.Shapes)) {
        $acao = ''
        try { $acao = $sh.OnAction } catch { }
        if ($acao -eq 'AbrirFormMassa') {
            $sh.OnAction = 'IrParaImportar'
            Set-TextoShape $sh 'Importar (colar)'
            $trocou = $true
        }
        elseif ($acao -eq 'IrParaImportar') {
            Set-TextoShape $sh 'Importar (colar)'
            $jaEstava = $true
        }
    }
    if ($trocou) { "botao da aba Resultados: AbrirFormMassa -> IrParaImportar" }
    elseif ($jaEstava) { "botao da aba Resultados: ja apontava para IrParaImportar" }
    else { throw "botao de importacao em massa nao encontrado na aba Resultados" }

    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    $salvou = $true
    "importacao por aba instalada"
}
finally {
    # SO salva se a etapa chegou ao fim. Fechar com $true em caso de erro grava
    # um artefato pela metade -- foi exatamente assim que a aba ficou criada com
    # o botao sem texto na execucao anterior.
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
