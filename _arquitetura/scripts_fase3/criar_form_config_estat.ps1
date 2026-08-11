# criar_form_config_estat.ps1 - painel de configuracao da analise estatistica
#
# CAMADA DE APRESENTACAO. O form escreve nas MESMAS celulas que a aba ja usa
# (B4, D4, F4, H4, B5, D5, B8:C10, E8:F9) e o motor mEstatPeriodo segue sendo o
# unico que calcula. A cadeia continua:
#
#   formulario -> celulas existentes -> formulas existentes -> motor -> resultado
#
# LAYOUT (o que o gestor pediu)
#
#   PERIODO
#     ( ) Mensal      Mes:  [ 01 - Janeiro  v ]
#     ( ) Trimestral  [ ]1o [ ]2o [ ]3o [ ]4o
#     ( ) Anual
#     Ano [ v ]   Lote [ v ]
#
#   EXCLUSOES        1..5, cada uma com De e Ate
#   PERIODO EFETIVO  previsao ANTES de aplicar
#   [APLICAR] [LIMPAR] [CANCELAR]
#
# O ANO FICA SEMPRE VISIVEL, e nao so no modo Anual: o motor pergunta "janeiro
# de QUAL ano?". Escondê-lo nos modos Mensal e Trimestral deixaria o usuario
# escolhendo um mes sem saber de que ano ele e.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
# O codigo do form vem de src_producao\frmConfigEstatistica.txt, lido como UTF-8.
#
# Uso:
#   .\criar_form_config_estat.ps1 -Workbook ..\..\QC_Bioquimica.xlsm

param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Fonte = ''
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if ($Fonte -eq '') {
    $Fonte = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'src_producao\frmConfigEstatistica.txt'
}
$Fonte = (Resolve-Path -LiteralPath $Fonte).ProviderPath
$SENHA = 'qcini2025'

# acentos por codigo
$a_ = [string][char]0x00E1; $e_ = [string][char]0x00E9; $i_ = [string][char]0x00ED
$o_ = [string][char]0x00F3; $C_ = [string][char]0x00C7; $A2 = [string][char]0x00C3
$O_ = [string][char]0x00C3
$capForm = 'Configura' + [string][char]0x00E7 + $A2 + 'o da an' + $a_ + 'lise'
$lblPeriodo = 'PER' + [string][char]0x00CD + 'ODO DE AN' + [string][char]0x00C1 + 'LISE'
$lblExcl = 'EXCLUS' + [string][char]0x00D5 + 'ES DE DADOS'
$lblMesCap = 'M' + $e_ + 's'
$q1 = '1' + [string][char]0x00BA + ' Tri'; $q2 = '2' + [string][char]0x00BA + ' Tri'
$q3 = '3' + [string][char]0x00BA + ' Tri'; $q4 = '4' + [string][char]0x00BA + ' Tri'
$ate = $a_ + 't' + $e_
$capAplicar = 'APLICAR'
$btnCfgTxt = 'CONFIGURAR AN' + [string][char]0x00C1 + 'LISE'

function Novo-Excel {
    $u = $null
    for ($t = 1; $t -le 6; $t++) {
        try { return (New-Object -ComObject Excel.Application) }
        catch {
            $u = $_
            if ($t -eq 2) { try { Start-Process excel.exe -WindowStyle Hidden -EA SilentlyContinue | Out-Null; Start-Sleep 5 } catch { } }
            Start-Sleep -Seconds ($t * 2)
        }
    }
    throw "Excel COM nao subiu: $($u.Exception.Message)"
}

# ---------------------------------------------------------------------------
# PASSO 1: remover a versao anterior, SALVAR E FECHAR.
#
# Remover e criar um UserForm na MESMA sessao do VBE falha com 0x800A004B
# (CTL_E_PATHFILEACCESSERROR): o VBE ainda segura o arquivo temporario do
# componente removido quando o Add tenta escrever o do novo. Fechar a pasta
# entre as duas operacoes libera o handle. Custa alguns segundos e evita um erro
# que parece permissao de disco e nao e.
$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Workbook)
if ($wb.ReadOnly) { try { $wb.Close($false) } catch { }; $xl.Quit(); throw "Somente leitura: $Workbook" }
$estruturaEstava = $wb.ProtectStructure
if ($estruturaEstava) { $wb.Unprotect($SENHA) }
$removeu = $false
foreach ($c in @($wb.VBProject.VBComponents)) {
    if ($c.Name -eq 'frmConfigEstatistica') { $wb.VBProject.VBComponents.Remove($c); $removeu = $true }
}
if ($removeu) {
    if ($estruturaEstava) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    "removido: frmConfigEstatistica anterior (pasta fechada para liberar o VBE)"
}
$wb.Close($removeu)
$xl.Quit()
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
Start-Sleep -Seconds 2

# ---------------------------------------------------------------------------
# PASSO 2: sessao nova, cria o formulario
$xl = Novo-Excel
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) { try { $wb.Close($false) } catch { }; $xl.Quit(); throw "Somente leitura: $Workbook" }

$salvar = $false
try {
    $estruturaEstava = $wb.ProtectStructure
    if ($estruturaEstava) { $wb.Unprotect($SENHA) }
    $proj = $wb.VBProject

    $f = $proj.VBComponents.Add(3)          # 3 = vbext_ct_MSForm
    $f.Properties.Item('Name').Value = 'frmConfigEstatistica'
    $f.Properties.Item('Caption').Value = $capForm
    $f.Properties.Item('Width').Value = 430
    $f.Properties.Item('Height').Value = 470
    $des = $f.Designer

    function Novo {
        param($D, [string]$Tipo, [string]$Nome, [int]$L, [int]$T, [int]$W, [int]$H, [string]$Cap = '')
        $ctl = $D.Controls.Add($Tipo, $Nome, $true)
        $ctl.Left = $L; $ctl.Top = $T; $ctl.Width = $W; $ctl.Height = $H
        if ($Cap -ne '') { try { $ctl.Caption = $Cap } catch { } }
        return $ctl
    }
    $LBL = 'Forms.Label.1'; $CBO = 'Forms.ComboBox.1'; $TXT = 'Forms.TextBox.1'
    $BTN = 'Forms.CommandButton.1'; $OPT = 'Forms.OptionButton.1'; $CHK = 'Forms.CheckBox.1'

    # ---------- periodo ----------
    Novo $des $LBL 'lblTitPer' 12 8 400 14 $lblPeriodo | Out-Null
    Novo $des $OPT 'optMensal' 20 28 80 16 'Mensal' | Out-Null
    Novo $des $LBL 'lblMes' 116 30 30 14 $lblMesCap | Out-Null
    Novo $des $CBO 'cboMes' 150 28 150 18 | Out-Null

    Novo $des $OPT 'optTrimestral' 20 52 90 16 'Trimestral' | Out-Null
    Novo $des $CHK 'chkQ1' 116 52 60 16 $q1 | Out-Null
    Novo $des $CHK 'chkQ2' 180 52 60 16 $q2 | Out-Null
    Novo $des $CHK 'chkQ3' 244 52 60 16 $q3 | Out-Null
    Novo $des $CHK 'chkQ4' 308 52 60 16 $q4 | Out-Null

    Novo $des $OPT 'optAnual' 20 76 80 16 'Anual' | Out-Null

    Novo $des $LBL 'lblAno' 116 78 26 14 'Ano' | Out-Null
    Novo $des $CBO 'cboAno' 146 76 70 18 | Out-Null
    Novo $des $LBL 'lblLote' 232 78 28 14 'Lote' | Out-Null
    Novo $des $CBO 'cboLote' 264 76 104 18 | Out-Null

    # ---------- exclusoes ----------
    Novo $des $LBL 'lblTitExc' 12 106 400 14 $lblExcl | Out-Null
    for ($k = 1; $k -le 5; $k++) {
        $y = 126 + ($k - 1) * 24
        Novo $des $LBL "lblEx$k" 20 ($y + 2) 62 14 ("Exclus" + $a_ + "o " + $k) | Out-Null
        Novo $des $TXT "txtEx${k}Ini" 88 $y 92 18 | Out-Null
        Novo $des $LBL "lblExAte$k" 186 ($y + 2) 20 14 $ate | Out-Null
        Novo $des $TXT "txtEx${k}Fim" 208 $y 92 18 | Out-Null
    }

    # ---------- previsao ----------
    Novo $des $LBL 'lblTitPrev' 12 254 400 14 ('PER' + [string][char]0x00CD + 'ODO EFETIVO') | Out-Null
    $prev = Novo $des $LBL 'lblPrev' 20 272 386 56 ''
    try { $prev.WordWrap = $true } catch { }

    # ---------- botoes ----------
    Novo $des $BTN 'btnAplicar' 20 340 120 30 $capAplicar | Out-Null
    Novo $des $BTN 'btnLimpar' 150 340 100 30 'LIMPAR' | Out-Null
    Novo $des $BTN 'btnCancelar' 286 340 120 30 'CANCELAR' | Out-Null

    # ---------- codigo ----------
    $encFonte = New-Object System.Text.UTF8Encoding($false)
    $codigo = [System.IO.File]::ReadAllText($Fonte, $encFonte)
    $f.CodeModule.AddFromString($codigo)
    "formulario montado: $($des.Controls.Count) controles, $($f.CodeModule.CountOfLines) linhas"

    # ---------- o abridor vive na FONTE do modulo, nao e injetado aqui ----------
    $mod = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'mEstatPeriodo') { $mod = $c; break } }
    if ($mod -eq $null) { throw 'mEstatPeriodo ausente' }
    $txtMod = $mod.CodeModule.Lines(1, $mod.CodeModule.CountOfLines)
    if ($txtMod -notmatch 'Sub AbrirConfigEstatistica') {
        throw "AbrirConfigEstatistica ausente de mEstatPeriodo. Rode antes: aplicar_vba.ps1 com src_producao\mEstatPeriodo.bas."
    }

    # ---------- botao na aba ----------
    $est = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -like 'Estat*') { $est = $w; break } }
    if ($est.ProtectContents) { try { $est.Unprotect($SENHA) } catch { } }
    foreach ($sh in @($est.Shapes)) { if ($sh.Name -eq 'btnConfigEstat') { $sh.Delete() } }
    $b = $est.Shapes.AddShape(5, $est.Range('J3').Left, $est.Range('J3').Top, 160, 26)
    $b.Name = 'btnConfigEstat'
    $b.TextFrame2.TextRange.Text = $btnCfgTxt
    $b.TextFrame2.TextRange.Font.Size = 9
    $b.TextFrame2.TextRange.Font.Bold = $true
    $b.OnAction = 'AbrirConfigEstatistica'

    # ---------- B11: a formula do periodo efetivo se perde em sessao manual ----
    # Ja se perdeu duas vezes. Restaurar aqui, junto do form, porque e ela que
    # deixa a exclusao ativa VISIVEL na aba -- sem isso o filtro age calado.
    if ("$($est.Range('B11').Formula)" -notlike '=*') {
        $est.Range('B11').Formula = '=PeriodoEfetivo(Estat_Ini_Efetiva,Estat_Fim_Efetiva,Estat_Exclusoes)'
        "B11: formula do periodo efetivo restaurada"
    }

    $xl.Calculation = -4105
    $wb.Application.CalculateFullRebuild()

    # ---------- conferencia ----------
    $erros = @()
    $nomes = @()
    foreach ($ctl in $des.Controls) { $nomes += $ctl.Name }
    $obrig = @('optMensal', 'optTrimestral', 'optAnual', 'cboMes', 'cboAno', 'cboLote',
        'chkQ1', 'chkQ2', 'chkQ3', 'chkQ4', 'lblPrev',
        'txtEx1Ini', 'txtEx1Fim', 'txtEx5Ini', 'txtEx5Fim',
        'btnAplicar', 'btnLimpar', 'btnCancelar')
    foreach ($o in $obrig) { if ($nomes -notcontains $o) { $erros += "controle ausente: $o" } }

    # NAO chamar AbrirConfigEstatistica: .Show e MODAL e bloqueia o COM para
    # sempre. Rodar uma funcao PURA do projeto prova que ele compila.
    try { $null = $xl.Run('JanelaInicio', 'TRIMESTRAL', 2026, 2, '', '') }
    catch { $erros += "o projeto VBA nao compila apos incluir o formulario: $($_.Exception.Message.Split([char]10)[0])" }

    $btnOk = $false
    foreach ($sh in $est.Shapes) { if ($sh.Name -eq 'btnConfigEstat' -and "$($sh.OnAction)" -like '*AbrirConfigEstatistica*') { $btnOk = $true } }
    if (-not $btnOk) { $erros += 'botao btnConfigEstat ausente ou sem OnAction' }

    if ("$($est.Range('B11').Value2)" -eq '') { $erros += 'B11 (periodo efetivo) nao calculou' }

    if ($erros.Count -gt 0) {
        $erros | ForEach-Object { "  FALHA: $_" }
        throw "Montagem recusada: $($erros.Count) problema(s). Nada foi salvo."
    }
    "conferencia: $($nomes.Count) controles, todos os obrigatorios, projeto compila, B11='$($est.Range('B11').Value2)'"

    if ($estruturaEstava -and -not $wb.ProtectStructure) { $wb.Protect($SENHA, $true, $false) }
    $salvar = $true
    $wb.Save()
    "Salvo: $Workbook"
}
finally {
    try { if ($salvar) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
}
