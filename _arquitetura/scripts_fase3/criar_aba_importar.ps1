# criar_aba_importar.ps1 - monta a aba Importar no artefato de build
#
# Substitui o frmMassa por uma ABA de entrada. O cabecalho NAO e gerado da aba
# Analitos: vem de src_producao/analitos_<produto>.csv, que fixa a ordem e as
# siglas definidas pelo gestor. Gerar dinamicamente faria a ordem das colunas
# mudar sozinha quando alguem cadastrasse um analito -- e a planilha que o
# laboratorio cola tem ordem fixa.
#
# LAYOUT
#   linha 3   OCULTA: nome do analito cadastrado para onde a coluna vai
#   linha 4   cabecalho visivel (siglas de bancada)
#   linha 5+  area de colagem, destravada
#
# A linha 3 e o de/para. O banco guarda o NOME cadastrado ("Glicose"), nao a
# sigla ("GLI"): e por esse nome que Calc, Painel, Estatistica e graficos acham
# o resultado. Manter isso como DADO -- e nao como tabela dentro do .bas --
# evita nome acentuado em fonte VBA e permite trocar o cabecalho sem recompilar.
#
# Este script NAO importa modulo VBA: quem faz isso e a etapa 4 do build.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_aba_importar.ps1 -Workbook <build.xlsm> [-Csv <analitos_x.csv>]

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Csv
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if (-not $Csv) {
    $Csv = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'src_producao\analitos_bioquimica.csv'
}
$Csv = (Resolve-Path -LiteralPath $Csv).ProviderPath

$SENHA = 'qcini2025'
$MAP = 3
$CAB = 4
$R0 = 5
$RN = 204

# --- CSV em UTF-8: nome de analito tem acento -----------------------------
$linhasCsv = [System.IO.File]::ReadAllLines($Csv, [System.Text.Encoding]::UTF8)
$def = @()
for ($i = 1; $i -lt $linhasCsv.Count; $i++) {
    $l = $linhasCsv[$i].Trim()
    if ($l -eq '') { continue }
    $p = $l.Split(',')
    if ($p.Count -lt 5) { throw "linha $($i+1) do CSV mal formada: $l" }
    $def += [pscustomobject]@{ Sigla = $p[1]; Nome = $p[2] }
}
if ($def.Count -eq 0) { throw "CSV sem definicoes" }

function Set-TextoShape {
    param($Shape, [string]$Texto)
    # Form control (Type 8) nao expoe TextFrame2.TextRange; AutoShape sim.
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

$estruturaEstava = $wb.ProtectStructure
if ($estruturaEstava) { $wb.Unprotect($SENHA) }

try {
    # ---- analitos cadastrados, para resolver o de/para -------------------
    $an = $wb.Worksheets.Item('Analitos')
    $cadastrados = @{}
    for ($i = 4; $i -le 43; $i++) {
        $v = $an.Cells.Item($i, 1).Value2
        if ($v -ne $null -and "$v".Trim() -ne '') { $cadastrados["$v".Trim()] = $true }
    }
    $semCadastro = @()
    foreach ($d in $def) { if (-not $cadastrados.ContainsKey($d.Nome)) { $semCadastro += $d.Sigla } }

    # ---- (re)cria a aba -------------------------------------------------
    foreach ($ws in @($wb.Worksheets)) { if ($ws.Name -eq 'Importar') { $ws.Visible = -1; $ws.Delete() } }
    $imp = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $imp.Name = 'Importar'

    $imp.Range('B1').Value2 = 'IMPORTACAO DE RESULTADOS'
    $imp.Range('B1').Font.Bold = $true
    $imp.Range('B1').Font.Size = 14
    $imp.Range('B2').Value2 = 'Cole os dados a partir da linha ' + $R0 + ' - uma corrida por linha, na ordem do cabecalho. Ao clicar em Registrar tudo e validado: havendo erro, NADA e gravado e as linhas com problema aparecem a direita. Sem erro, os dados migram para o DB_Resultados e somem daqui.'
    $imp.Range('B2').Font.Italic = $true

    $cabecalho = @('DATA', 'NIVEL', 'LOTE') + ($def | ForEach-Object { $_.Sigla })
    for ($k = 0; $k -lt $def.Count; $k++) {
        $nome = $def[$k].Nome
        if (-not $cadastrados.ContainsKey($nome)) { $nome = '' }
        $imp.Cells.Item($MAP, 5 + $k).Value2 = [string]$nome
    }
    $imp.Rows.Item($MAP).Hidden = $true

    for ($k = 0; $k -lt $cabecalho.Count; $k++) {
        $c = $imp.Cells.Item($CAB, $k + 2)
        $c.Value2 = [string]$cabecalho[$k]
        $c.Font.Bold = $true
        $c.HorizontalAlignment = -4108
    }
    $ultCol = $cabecalho.Count + 1
    $rngCab = $imp.Range($imp.Cells.Item($CAB, 2), $imp.Cells.Item($CAB, $ultCol))
    $rngCab.Interior.Color = 14277081
    $rngCab.Borders.LineStyle = 1

    # analito ainda nao cadastrado fica cinza: o usuario ve que ali nao entra dado
    for ($k = 0; $k -lt $def.Count; $k++) {
        if (-not $cadastrados.ContainsKey($def[$k].Nome)) {
            $imp.Cells.Item($CAB, 5 + $k).Interior.Color = 12632256
            $imp.Cells.Item($CAB, 5 + $k).Font.Color = 8421504
        }
    }

    $imp.Columns.Item(2).ColumnWidth = 12
    $imp.Columns.Item(3).ColumnWidth = 7
    $imp.Columns.Item(4).ColumnWidth = 12
    for ($k = 5; $k -le $ultCol; $k++) { $imp.Columns.Item($k).ColumnWidth = 8 }

    $entrada = $imp.Range($imp.Cells.Item($R0, 2), $imp.Cells.Item($RN, $ultCol))
    $entrada.Locked = $false
    $imp.Range($imp.Cells.Item($R0, 2), $imp.Cells.Item($RN, 2)).NumberFormat = '@'

    # ---- botao Registrar ------------------------------------------------
    foreach ($sh in @($imp.Shapes)) { $sh.Delete() }
    $b = $imp.Shapes.AddShape(5, 12, 6, 170, 34)
    $b.Name = 'btnRegistrar'
    $b.TextFrame2.TextRange.Text = 'Registrar'
    $b.TextFrame2.TextRange.Font.Size = 12
    $b.TextFrame2.TextRange.Font.Bold = $true
    $b.Fill.ForeColor.RGB = 3506772
    $b.Line.Visible = $false
    $b.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    $b.OnAction = 'RegistrarImportacao'

    $imp.Cells.Item($R0, 2).Select() | Out-Null
    $imp.Protect($SENHA, $true, $true, $true, $true)

    "aba Importar: $($def.Count) colunas de analito, entrada em B$R0"
    if ($semCadastro.Count -gt 0) { "  sem cadastro em Analitos ($($semCadastro.Count)): $($semCadastro -join ', ')" }

    # ---- botao da aba Resultados ----------------------------------------
    $resu = $wb.Worksheets.Item('Resultados')
    $trocou = $false
    foreach ($sh in @($resu.Shapes)) {
        $acao = ''
        try { $acao = $sh.OnAction } catch { }
        if ($acao -eq 'AbrirFormMassa' -or $acao -eq 'IrParaImportar') {
            $sh.OnAction = 'IrParaImportar'
            Set-TextoShape $sh 'Importar (colar)'
            $trocou = $true
        }
    }
    if (-not $trocou) { throw "botao de importacao em massa nao encontrado na aba Resultados" }
    "botao da aba Resultados -> IrParaImportar"

    # ---- frmMassa sai de cena -------------------------------------------
    # AbrirFormMassa e reapontado ANTES da remocao: mOperacao referencia o
    # formulario e nao compilaria sem ele.
    $vbp = $wb.VBProject
    foreach ($c in $vbp.VBComponents) {
        if ($c.Name -eq 'mOperacao') {
            $cm = $c.CodeModule
            $ini = $cm.ProcStartLine('AbrirFormMassa', 0)
            $n = $cm.ProcCountLines('AbrirFormMassa', 0)
            if ($ini -gt 0) {
                $cm.DeleteLines($ini, $n)
                $cm.InsertLines($ini, @(
                    "' O frmMassa foi substituido pela aba Importar (ADR-020). Este ponto de",
                    "' entrada continua existindo porque botoes antigos apontam para ele.",
                    "Public Sub AbrirFormMassa()",
                    "    IrParaImportar",
                    "End Sub"
                ) -join "`r`n")
            }
        }
    }
    foreach ($c in @($vbp.VBComponents)) {
        if ($c.Name -eq 'frmMassa') { $vbp.VBComponents.Remove($c); "frmMassa removido" }
    }

    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    $salvou = $true
    "importacao por aba instalada"
}
finally {
    # SO salva se a etapa chegou ao fim. Fechar com $true em caso de erro grava
    # um artefato pela metade.
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
