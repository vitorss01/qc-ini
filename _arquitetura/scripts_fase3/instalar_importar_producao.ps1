# instalar_importar_producao.ps1 - poe a aba Importar no arquivo de PRODUCAO
#
# Diferente de criar_aba_importar.ps1, que monta a aba no artefato de build a
# partir da aba Analitos. Aqui o cabecalho e FIXO, na ordem exata pedida pelo
# gestor, e o que varia e o mapeamento sigla -> analito cadastrado.
#
# O QUE ESTE SCRIPT FAZ
#   1. destrava as abas (o arquivo pode ter sido salvo logado, com tudo travado)
#   2. cria a aba Importar com:
#        linha 4 (visivel) : DATA | NIVEL | LOTE | as 31 siglas, na ordem pedida
#        linha 3 (OCULTA)  : nome do analito cadastrado para onde a coluna vai
#   3. cria o botao Registrar
#   4. troca o botao da aba Resultados para IrParaImportar
#   5. importa mImportar.bas, remove o frmMassa e reaponta AbrirFormMassa
#   6. SALVA o arquivo
#
# POR QUE A LINHA DE MAPEAMENTO
#   O banco guarda o NOME cadastrado ("Glicose"), nao a sigla ("GLI"): e por
#   esse nome que Calc, Painel, Estatistica e graficos acham o resultado.
#   Gravar a sigla quebraria todos em silencio. O mapeamento fica em celula --
#   dado, nao codigo -- para nao haver nome acentuado dentro do .bas.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\instalar_importar_producao.ps1 -Workbook <QC_Bioquimica.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Modulo
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if (-not $Modulo) {
    $Modulo = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'src_producao\mImportar.bas'
}
$Modulo = (Resolve-Path -LiteralPath $Modulo).ProviderPath

$SENHA = 'qcini2025'
$MAP = 3      # linha oculta de mapeamento
$CAB = 4      # linha do cabecalho visivel
$R0  = 5
$RN  = 204

# Cabecalho EXATO pedido pelo gestor. Ordem fixa, nao gerada.
$CABECALHO = @(
    'DATA', 'NIVEL', 'LOTE',
    'ALA', 'AUR', 'ALB', 'BD', 'BT', 'CAL', 'CFF', 'COL', 'HDL', 'CRE', 'FER',
    'FOS', 'GLI', 'HbA1c', 'MAG', 'PTN T', 'TRI', 'URE', 'AMI', 'CPK', 'FAL',
    'GGT', 'LDH', 'LPS', 'TGO', 'TGP', 'PCR', 'ECO2', 'NA', 'K', 'CL'
)

# Sigla -> analito cadastrado na aba Analitos.
#
# Chave em ASCII; o VALOR e procurado na aba Analitos por comparacao
# NORMALIZADA (sem acento, maiusculas, so alfanumerico). Assim o script nao
# precisa carregar "Acido urico" com acento e continua achando a linha certa.
$MAPA = @{
    'AUR'   = 'ACIDOURICO'
    'ALB'   = 'ALBUMINA'
    'BT'    = 'BILIRRUBINATOTAL'
    'CAL'   = 'CALCIO'
    'COL'   = 'COLESTEROLTOTAL'
    'HDL'   = 'HDLCOLESTEROL'
    'CRE'   = 'CREATININA'
    'FOS'   = 'FOSFORO'
    'GLI'   = 'GLICOSE'
    'PTN T' = 'PROTEINATOTAL'
    'TRI'   = 'TRIGLICERIDEOS'
    'URE'   = 'UREIA'
    'AMI'   = 'AMILASE'
    'FAL'   = 'FOSFATASEALCALINA'
    'GGT'   = 'GGT'
    'TGO'   = 'ASTTGO'
    'TGP'   = 'ALTTGP'
    'NA'    = 'SODIO'
    'K'     = 'POTASSIO'
    'CL'    = 'CLORO'
}

function Normalizar {
    param([string]$s)
    if (-not $s) { return '' }
    $n = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $n.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne 'NonSpacingMark') { [void]$sb.Append($ch) }
    }
    return ($sb.ToString().ToUpperInvariant() -replace '[^A-Z0-9]', '')
}

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

try {
    # ---- 1. destrava tudo -----------------------------------------------
    if ($wb.ProtectStructure) { $wb.Unprotect($SENHA); "estrutura destravada" }
    $destravadas = 0
    foreach ($ws in @($wb.Worksheets)) {
        if (-not $ws.ProtectContents) { continue }
        try { $ws.Unprotect($SENHA) } catch { try { $ws.Unprotect() } catch { } }
        if (-not $ws.ProtectContents) { $destravadas++ }
        else { throw "aba '$($ws.Name)' nao abriu com a senha do projeto" }
    }
    "abas destravadas: $destravadas"

    # ---- 2. le os analitos cadastrados ----------------------------------
    $an = $wb.Worksheets.Item('Analitos')
    $cadastrados = @{}          # normalizado -> nome exato
    for ($i = 4; $i -le 43; $i++) {
        $v = $an.Cells.Item($i, 1).Value2
        if ($v -ne $null -and "$v".Trim() -ne '') { $cadastrados[(Normalizar "$v")] = "$v" }
    }
    if ($cadastrados.Count -eq 0) { throw "nenhum analito na aba Analitos" }
    "analitos cadastrados: $($cadastrados.Count)"

    # resolve cada coluna do cabecalho para o nome cadastrado (ou vazio)
    $destino = @()
    $semCadastro = @()
    for ($k = 3; $k -lt $CABECALHO.Count; $k++) {
        $sigla = $CABECALHO[$k]
        $nome = ''
        if ($MAPA.ContainsKey($sigla)) {
            $chave = $MAPA[$sigla]
            if ($cadastrados.ContainsKey($chave)) { $nome = $cadastrados[$chave] }
            else { throw "mapeamento aponta para '$chave', que nao existe na aba Analitos (sigla $sigla)" }
        }
        $destino += $nome
        if ($nome -eq '') { $semCadastro += $sigla }
    }
    "colunas de analito : $($destino.Count)"
    "mapeadas           : $(($destino | Where-Object { $_ -ne '' }).Count)"
    if ($semCadastro.Count -gt 0) {
        "SEM CADASTRO ($($semCadastro.Count)): $($semCadastro -join ', ')"
    }

    # ---- 3. (re)cria a aba Importar -------------------------------------
    foreach ($ws in @($wb.Worksheets)) { if ($ws.Name -eq 'Importar') { $ws.Visible = -1; $ws.Delete() } }
    $imp = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $imp.Name = 'Importar'

    $imp.Range('B1').Value2 = 'IMPORTACAO DE RESULTADOS'
    $imp.Range('B1').Font.Bold = $true
    $imp.Range('B1').Font.Size = 14
    $imp.Range('B2').Value2 = 'Cole os dados a partir da linha ' + $R0 + ' - uma corrida por linha, na ordem do cabecalho. Ao clicar em Registrar tudo e validado: havendo erro, NADA e gravado e as linhas com problema aparecem a direita. Sem erro, os dados migram para o DB_Resultados e somem daqui.'
    $imp.Range('B2').Font.Italic = $true

    # linha OCULTA de mapeamento
    for ($k = 0; $k -lt $destino.Count; $k++) {
        $imp.Cells.Item($MAP, 5 + $k).Value2 = [string]$destino[$k]
    }
    $imp.Rows.Item($MAP).Hidden = $true

    # cabecalho visivel
    for ($k = 0; $k -lt $CABECALHO.Count; $k++) {
        $c = $imp.Cells.Item($CAB, $k + 2)      # comeca em B
        $c.Value2 = [string]$CABECALHO[$k]
        $c.Font.Bold = $true
        $c.HorizontalAlignment = -4108          # centro
    }
    $ultCol = $CABECALHO.Count + 1
    $rngCab = $imp.Range($imp.Cells.Item($CAB, 2), $imp.Cells.Item($CAB, $ultCol))
    $rngCab.Interior.Color = 14277081
    $rngCab.Borders.LineStyle = 1

    # coluna de analito nao cadastrado fica cinza: o usuario ve que ali nao entra dado
    for ($k = 0; $k -lt $destino.Count; $k++) {
        if ($destino[$k] -eq '') {
            $imp.Cells.Item($CAB, 5 + $k).Interior.Color = 12632256
            $imp.Cells.Item($CAB, 5 + $k).Font.Color = 8421504
        }
    }

    $imp.Columns.Item(2).ColumnWidth = 12    # DATA
    $imp.Columns.Item(3).ColumnWidth = 7     # NIVEL
    $imp.Columns.Item(4).ColumnWidth = 12    # LOTE
    for ($k = 5; $k -le $ultCol; $k++) { $imp.Columns.Item($k).ColumnWidth = 8 }

    # area de entrada destravada
    $entrada = $imp.Range($imp.Cells.Item($R0, 2), $imp.Cells.Item($RN, $ultCol))
    $entrada.Locked = $false
    $imp.Range($imp.Cells.Item($R0, 2), $imp.Cells.Item($RN, 2)).NumberFormat = '@'   # DATA como texto

    # ---- 4. botao Registrar ---------------------------------------------
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
    "botao Registrar criado"

    $imp.Cells.Item($R0, 2).Select() | Out-Null
    $imp.Protect($SENHA, $true, $true, $true, $true)

    # ---- 5. botao da aba Resultados -------------------------------------
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

    # ---- 6. VBA: importa mImportar, remove frmMassa ---------------------
    $vbp = $wb.VBProject
    foreach ($c in @($vbp.VBComponents)) {
        if ($c.Name -eq 'mImportar') { $vbp.VBComponents.Remove($c) }
    }
    $vbp.VBComponents.Import($Modulo) | Out-Null
    $comp = $null
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq 'mImportar') { $comp = $c } }
    if ($comp -eq $null) { throw "mImportar nao entrou no projeto" }
    "mImportar importado: $($comp.CodeModule.CountOfLines) linhas"

    # AbrirFormMassa passa a navegar. Precisa vir ANTES de remover o frmMassa:
    # o modulo mOperacao referencia o formulario e nao compilaria sem ele.
    foreach ($c in $vbp.VBComponents) {
        if ($c.Name -eq 'mOperacao') {
            $cm = $c.CodeModule
            $lin = $cm.ProcBodyLine('AbrirFormMassa', 0)
            $n = $cm.ProcCountLines('AbrirFormMassa', 0)
            $ini = $cm.ProcStartLine('AbrirFormMassa', 0)
            $cm.DeleteLines($ini, $n)
            $cm.InsertLines($ini, @(
                "' O frmMassa foi substituido pela aba Importar (ADR-020). Este ponto de",
                "' entrada continua existindo porque botoes antigos ainda apontam para ele.",
                "Public Sub AbrirFormMassa()",
                "    IrParaImportar",
                "End Sub"
            ) -join "`r`n")
            "AbrirFormMassa -> IrParaImportar"
        }
    }
    foreach ($c in @($vbp.VBComponents)) {
        if ($c.Name -eq 'frmMassa') { $vbp.VBComponents.Remove($c); "frmMassa removido" }
    }

    # ---- 7. salva -------------------------------------------------------
    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
