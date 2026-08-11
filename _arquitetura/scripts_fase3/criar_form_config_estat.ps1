# criar_form_config_estat.ps1 - UserForm de configuracao da analise estatistica
#
# CAMADA DE APRESENTACAO. O formulario le e escreve as MESMAS celulas que a aba
# ja usa (B4, D4, F4, H4, B5, D5, B8:C10, E8:F9) e o motor mEstatPeriodo segue
# sendo o unico que calcula. Nenhuma regra e reimplementada aqui.
#
# O QUE ELE RESOLVE
#
# Na aba, "Mes / Tri / Sem" pede um NUMERO: com Visao=TRIMESTRAL o usuario tem
# de saber de cabeca que 2 significa abril-junho. No formulario a lista muda com
# a visao e mostra "2 - 2o trimestre (abr-jun)". O numero que vai para a celula
# e o mesmo; o que muda e o que o usuario precisa saber.
#
# E a janela efetiva aparece ANTES de aplicar, calculada pelas proprias funcoes
# do motor (JanelaInicio/JanelaFim/PeriodoEfetivo) -- prever com uma copia da
# conta seria criar a segunda verdade que o ADR-019 existe para impedir.
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
$o_ = [string][char]0x00F3; $c_ = [string][char]0x00E7; $a2 = [string][char]0x00E3
$A2 = [string][char]0x00C3
$vis = 'Vis' + $a2 + 'o'
$mes = 'M' + $e_ + 's'
$per = 'Per' + $i_ + 'odo'
$exc = 'Exclus' + $o_ + 'es'
$cfgTit = 'CONFIGURA' + [string][char]0x00C7 + $A2 + 'O DA AN' + [string][char]0x00C1 + 'LISE'

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

    # remove versao anterior
    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -eq 'frmConfigEstatistica') { $proj.VBComponents.Remove($c); "removido: frmConfigEstatistica anterior" }
    }

    # 3 = vbext_ct_MSForm
    $f = $proj.VBComponents.Add(3)
    $f.Properties.Item('Name').Value = 'frmConfigEstatistica'
    $f.Properties.Item('Caption').Value = 'Configura' + [string][char]0x00E7 + $a2 + 'o da an' + $a_ + 'lise estat' + $i_ + 'stica'
    $f.Properties.Item('Width').Value = 460
    $f.Properties.Item('Height').Value = 430
    $des = $f.Designer

    function Novo {
        param($D, [string]$Tipo, [string]$Nome, [int]$L, [int]$T, [int]$W, [int]$H, [string]$Cap = '')
        $ctl = $D.Controls.Add($Tipo, $Nome, $true)
        $ctl.Left = $L; $ctl.Top = $T; $ctl.Width = $W; $ctl.Height = $H
        if ($Cap -ne '') { try { $ctl.Caption = $Cap } catch { } }
        return $ctl
    }

    $LBL = 'Forms.Label.1'; $CBO = 'Forms.ComboBox.1'; $TXT = 'Forms.TextBox.1'
    $BTN = 'Forms.CommandButton.1'; $FRM = 'Forms.Frame.1'

    # ---- bloco 1: periodo ----
    $fr1 = Novo $des $FRM 'frPeriodo' 10 6 434 118 ($per + ' analisado')
    Novo $des $LBL 'lblVisao' 22 30 60 16 $vis | Out-Null
    Novo $des $CBO 'cboVisao' 86 28 120 18 | Out-Null
    Novo $des $LBL 'lblAno' 222 30 30 16 'Ano' | Out-Null
    Novo $des $CBO 'cboAno' 254 28 70 18 | Out-Null
    Novo $des $LBL 'lblParte' 22 58 60 16 ($mes + '/Tri/Sem') | Out-Null
    Novo $des $CBO 'cboParte' 86 56 238 18 | Out-Null
    Novo $des $LBL 'lblLote' 336 30 30 16 'Lote' | Out-Null
    Novo $des $CBO 'cboLote' 336 48 96 18 | Out-Null
    Novo $des $LBL 'lblPers' 22 86 60 16 'Personalizado' | Out-Null
    Novo $des $TXT 'txtPersIni' 86 84 90 18 | Out-Null
    Novo $des $LBL 'lblAte' 182 86 20 16 ($a_ + 't' + $e_) | Out-Null
    Novo $des $TXT 'txtPersFim' 204 84 90 18 | Out-Null

    # ---- bloco 2: exclusoes ----
    $fr2 = Novo $des $FRM 'frExcl' 10 130 434 150 ($exc + ' - resultados destes intervalos ficam fora do calculo')
    for ($k = 1; $k -le 5; $k++) {
        $y = 22 + ($k - 1) * 24
        Novo $des $LBL "lblEx$k" 22 ($y + 132) 20 16 "$k" | Out-Null
        Novo $des $TXT "txtEx${k}Ini" 46 ($y + 130) 90 18 | Out-Null
        Novo $des $LBL "lblExAte$k" 142 ($y + 132) 20 16 ($a_ + 't' + $e_) | Out-Null
        Novo $des $TXT "txtEx${k}Fim" 164 ($y + 130) 90 18 | Out-Null
    }

    # ---- janela efetiva + botoes ----
    Novo $des $LBL 'lblJanela' 14 288 430 18 'Periodo que vai valer:' | Out-Null
    Novo $des $BTN 'btnAplicar' 14 316 130 28 'APLICAR AN' + [string][char]0x00C1 + 'LISE' | Out-Null
    Novo $des $BTN 'btnLimpar' 154 316 100 28 'LIMPAR' | Out-Null
    Novo $des $BTN 'btnFechar' 344 316 100 28 'FECHAR' | Out-Null

    # ---- codigo ----
    $enc = [System.Text.Encoding]::Default
    $encFonte = New-Object System.Text.UTF8Encoding($false)
    $codigo = [System.IO.File]::ReadAllText($Fonte, $encFonte)
    $f.CodeModule.AddFromString($codigo)
    "formulario montado: $($des.Controls.Count) controles, $($f.CodeModule.CountOfLines) linhas de codigo"

    # ---- abridor no modulo do motor ----
    $mod = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'mEstatPeriodo') { $mod = $c; break } }
    if ($mod -eq $null) { throw 'mEstatPeriodo ausente -- o form depende das funcoes dele' }
    # NAO ACRESCENTAR CODIGO AO MODULO AQUI.
    #
    # A primeira versao deste script fazia AddFromString de AbrirConfigEstatistica
    # direto no .xlsm. O modulo dentro do arquivo passou a ter uma rotina que a
    # fonte versionada nao tinha, e a prova 1.1 (modulo identico a fonte)
    # reprovou -- corretamente. E a mesma deriva fonte/artefato que o ADR-021
    # existe para impedir, cometida por mim.
    #
    # A rotina agora vive em src_producao\mEstatPeriodo.bas e chega pelo
    # instalador do modulo. Aqui so se CONFERE que ela esta presente.
    $txt = $mod.CodeModule.Lines(1, $mod.CodeModule.CountOfLines)
    if ($txt -notmatch 'Sub AbrirConfigEstatistica') {
        throw "AbrirConfigEstatistica ausente de mEstatPeriodo. Rode antes: instalar_estat_periodo.ps1 (a rotina vive na fonte versionada, nao e injetada aqui)."
    }
    "AbrirConfigEstatistica: presente no modulo (vindo da fonte versionada)"

    # ---- botao na aba ----
    $est = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -like 'Estat*') { $est = $w; break } }
    if ($est.ProtectContents) { try { $est.Unprotect($SENHA) } catch { } }
    foreach ($sh in @($est.Shapes)) { if ($sh.Name -eq 'btnConfigEstat') { $sh.Delete() } }
    $b = $est.Shapes.AddShape(5, $est.Range('J3').Left, $est.Range('J3').Top, 150, 28)
    $b.Name = 'btnConfigEstat'
    $b.TextFrame2.TextRange.Text = $cfgTit
    $b.TextFrame2.TextRange.Font.Size = 9
    $b.TextFrame2.TextRange.Font.Bold = $true
    $b.OnAction = 'AbrirConfigEstatistica'
    "botao 'btnConfigEstat' criado em J3 -> AbrirConfigEstatistica"

    # ---- conferencia ----
    $erros = @()
    $achou = $false
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'frmConfigEstatistica') { $achou = $true } }
    if (-not $achou) { $erros += 'frmConfigEstatistica nao entrou no projeto' }
    $obrig = @('cboVisao', 'cboAno', 'cboParte', 'cboLote', 'txtPersIni', 'txtPersFim',
        'txtEx1Ini', 'txtEx1Fim', 'txtEx5Ini', 'txtEx5Fim', 'lblJanela',
        'btnAplicar', 'btnLimpar', 'btnFechar')
    $nomes = @()
    foreach ($ctl in $des.Controls) { $nomes += $ctl.Name }
    foreach ($o in $obrig) { if ($nomes -notcontains $o) { $erros += "controle ausente: $o" } }
    # NAO CHAMAR AbrirConfigEstatistica AQUI.
    #
    # Ela faz frmConfigEstatistica.Show, que e um dialogo MODAL. Sem ninguem para
    # fecha-lo, o Run bloqueia para sempre e o script fica pendurado com o Excel
    # aberto -- foi o que aconteceu na primeira execucao, e so um kill resolveu.
    #
    # O que da para conferir sem abrir a janela e o que importa: o componente
    # existe, os controles estao la, o codigo compila (Run de uma funcao PURA do
    # mesmo projeto prova a compilacao) e o botao aponta para o nome certo.
    try {
        $null = $xl.Run('JanelaInicio', 'ANUAL', 2026, 1, '', '')
    }
    catch { $erros += "o projeto VBA nao compila apos incluir o formulario: $($_.Exception.Message.Split([char]10)[0])" }

    $temAbridor = $false
    $txtMod = $mod.CodeModule.Lines(1, $mod.CodeModule.CountOfLines)
    if ($txtMod -match 'Sub AbrirConfigEstatistica') { $temAbridor = $true }
    if (-not $temAbridor) { $erros += 'AbrirConfigEstatistica ausente de mEstatPeriodo' }

    $btnOk = $false
    foreach ($sh in $est.Shapes) { if ($sh.Name -eq 'btnConfigEstat' -and "$($sh.OnAction)" -like '*AbrirConfigEstatistica*') { $btnOk = $true } }
    if (-not $btnOk) { $erros += 'botao btnConfigEstat ausente ou sem OnAction' }
    if ($erros.Count -gt 0) {
        $erros | ForEach-Object { "  FALHA: $_" }
        throw "Montagem do formulario recusada: $($erros.Count) problema(s). Nada foi salvo."
    }
    "conferencia: form presente, $($nomes.Count) controles, todos os obrigatorios"

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
