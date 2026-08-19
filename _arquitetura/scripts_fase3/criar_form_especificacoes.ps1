# criar_form_especificacoes.ps1 - formulario UNICO de cadastro (ADR-022)
#
# UM formulario, nao tres.
#
# O gestor pediu tres telas (CLIA, Variacao Biologica, Fabricante). Tres telas
# significariam tres copias da mesma logica de gravar, validar e listar -- e
# tres lugares para divergir. Aqui a FONTE e um combo, e os campos de entrada
# mudam conforme o MODELO dela:
#
#   ETP_DIRETO      ETp %
#   VB              CVi %, CVg %, Rigor
#   CV_BIAS_DIRETO  CVtp %, BIAStp %
#
# Fonte nova cadastrada em Cfg_Especificacoes aparece no combo sozinha; se usar
# um dos tres modelos, o formulario ja sabe desenha-la.
#
# O formulario NAO calcula nada: ele chama mEspecificacoes.MetasDaLinha para a
# previa e GravarEspec para gravar. Duplicar a formula aqui criaria a segunda
# fonte de verdade que o ADR-019 existe para impedir.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_form_especificacoes.ps1 -Workbook <x.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
$SENHA = 'qcini2025'

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
    if ($wb.ProtectStructure) { $wb.Unprotect($SENHA) }
    $vbp = $wb.VBProject

    foreach ($c in @($vbp.VBComponents)) {
        if ($c.Name -eq 'frmEspecificacoes') { $vbp.VBComponents.Remove($c) }
    }
    $f = $vbp.VBComponents.Add(3)      # 3 = UserForm
    $f.Properties.Item('Name').Value = 'frmEspecificacoes'
    $f.Properties.Item('Caption').Value = 'Especificacoes de Qualidade'
    $f.Properties.Item('Width').Value = 560
    $f.Properties.Item('Height').Value = 470
    $des = $f.Designer
    if ($des -eq $null) { throw "Designer nulo" }

    function Ctl {
        param($D, [string]$Tipo, [string]$Nome, [double]$L, [double]$T, [double]$W, [double]$H, [string]$Txt = '')
        $c = $D.Controls.Add("Forms.$Tipo.1", $Nome, $true)
        $c.Left = $L; $c.Top = $T; $c.Width = $W; $c.Height = $H
        if ($Txt -ne '') {
            if ($Tipo -eq 'Label' -or $Tipo -eq 'CommandButton') { $c.Caption = $Txt } else { $c.Text = $Txt }
        }
        return $c
    }
    # ---- cabecalho
    $l = Ctl $des 'Label' 'lblFonte' 12 10 40 16 'Fonte'
    $cbo = Ctl $des 'ComboBox' 'cboFonte' 12 28 190 20
    $cbo.Style = 2      # fmStyleDropDownList: so escolhe o que existe
    $l = Ctl $des 'Label' 'lblAno' 216 10 40 16 'Ano'
    $t = Ctl $des 'TextBox' 'txtAno' 216 28 60 20
    $l = Ctl $des 'Label' 'lblModelo' 290 30 250 16 ''
    $l.ForeColor = 8421504
    # ---- lista
    $l = Ctl $des 'Label' 'lblLista' 12 56 400 14 'Analitos (o que ja esta cadastrado para este Ano e Fonte)'
    $lst = Ctl $des 'ListBox' 'lstAnalitos' 12 74 530 168
    $lst.ColumnCount = 5
    $lst.ColumnWidths = '150 pt;60 pt;70 pt;70 pt;70 pt'
    $lst.ColumnHeads = $false
    # ---- entradas
    $l = Ctl $des 'Label' 'lblEnt1' 12 252 90 16 ''
    $t = Ctl $des 'TextBox' 'txtEnt1' 12 270 80 20
    $l = Ctl $des 'Label' 'lblEnt2' 104 252 90 16 ''
    $t = Ctl $des 'TextBox' 'txtEnt2' 104 270 80 20
    $l = Ctl $des 'Label' 'lblEnt3' 196 252 90 16 ''
    $cboR = Ctl $des 'ComboBox' 'cboRigor' 196 270 80 20
    $cboR.Style = 2
    # ---- previa
    # Font.Bold pelo Designer devolve NullReference no PowerShell; o negrito
    # e aplicado no UserForm_Initialize, onde o objeto Font e o do MSForms.
    $l = Ctl $des 'Label' 'lblPrevia' 300 252 240 16 'Previa do calculo'
    $l = Ctl $des 'Label' 'lblResultado' 300 270 240 46 ''

    $l = Ctl $des 'Label' 'lblAviso' 12 300 280 30 ''
    $l.ForeColor = 192
    # ---- botoes
    $b = Ctl $des 'CommandButton' 'btnGravar' 12 342 130 30 'Gravar analito'
    $b = Ctl $des 'CommandButton' 'btnGravarLista' 150 342 150 30 'Gravar e proximo'
    $b = Ctl $des 'CommandButton' 'btnFechar' 412 342 130 30 'Fechar'

    $l = Ctl $des 'Label' 'lblRodape' 12 382 530 44 ''
    $l.ForeColor = 8421504

    # ---- codigo ----------------------------------------------------------
    $codigo = @'
Option Explicit
' Formulario UNICO de especificacoes. Os campos de entrada mudam conforme o
' MODELO da fonte escolhida.
'
' Este formulario NAO calcula meta: a previa vem de MetasDaLinha e a gravacao
' de GravarEspec, ambos em mEspecificacoes. Repetir a formula aqui criaria a
' segunda fonte de verdade que o ADR-019 existe para impedir.
Private carregando As Boolean

Private Sub UserForm_Initialize()
    Dim c As Collection, i As Long, rigor As String
    carregando = True
    Me.lblPrevia.Font.Bold = True
    Set c = ListaFontes()
    For i = 1 To c.Count
        Me.cboFonte.AddItem c(i)
    Next i
    If Me.cboFonte.ListCount > 0 Then Me.cboFonte.ListIndex = 0
    Me.cboRigor.AddItem "MIN"
    Me.cboRigor.AddItem "DES"
    Me.cboRigor.AddItem "OTI"
    rigor = RigorPadrao()
    If rigor <> "" Then Me.cboRigor.Value = rigor
    Me.txtAno.Text = CStr(Year(Date))
    Me.lblRodape.Caption = "A meta vigente para um resultado e a de maior Ano que nao ultrapasse o ano DELE. " & _
                           "Cadastrar 2027 nao altera como uma corrida de 2025 e julgada."
    carregando = False
    ConfigurarDisponibilidade
    AjustarCampos
    CarregarLista
End Sub

Private Sub ConfigurarDisponibilidade()
    Dim disponivel As Boolean
    disponivel = (Me.cboFonte.ListCount > 0)
    Me.cboFonte.Enabled = disponivel
    Me.lstAnalitos.Enabled = disponivel
    Me.btnGravar.Enabled = disponivel
    Me.btnGravarLista.Enabled = disponivel
    If Not disponivel Then
        Me.lblAviso.Caption = "Cadastre uma fonte e seu modelo em Cfg_Especificacoes."
        Me.lblModelo.Caption = "modelo: nao configurado"
    End If
End Sub

' Desenha os campos conforme o modelo da fonte.
Private Sub AjustarCampos()
    Dim modelo As String
    modelo = ModeloDaFonte(Me.cboFonte.Value)
    If modelo = "" Then
        Me.lblModelo.Caption = "modelo: nao configurado"
    Else
        Me.lblModelo.Caption = "modelo: " & modelo
    End If

    Me.lblEnt1.Visible = False: Me.txtEnt1.Visible = False
    Me.lblEnt2.Visible = False: Me.txtEnt2.Visible = False
    Me.lblEnt3.Visible = False: Me.cboRigor.Visible = False

    Select Case modelo
        Case "ETP_DIRETO"
            Me.lblEnt1.Caption = "ETp %": Me.lblEnt1.Visible = True: Me.txtEnt1.Visible = True
        Case "VB"
            Me.lblEnt1.Caption = "CVi %": Me.lblEnt1.Visible = True: Me.txtEnt1.Visible = True
            Me.lblEnt2.Caption = "CVg %": Me.lblEnt2.Visible = True: Me.txtEnt2.Visible = True
            Me.lblEnt3.Caption = "Rigor": Me.lblEnt3.Visible = True: Me.cboRigor.Visible = True
        Case "CV_BIAS_DIRETO"
            Me.lblEnt1.Caption = "CVtp %": Me.lblEnt1.Visible = True: Me.txtEnt1.Visible = True
            Me.lblEnt2.Caption = "BIAStp %": Me.lblEnt2.Visible = True: Me.txtEnt2.Visible = True
    End Select
    AtualizarPrevia
End Sub

' Lista os analitos do produto com o que ja esta resolvido para Ano + Fonte.
Private Sub CarregarLista()
    Dim c As Collection, i As Long, ano As Long, r As String, p() As String
    Dim guardado As Long
    guardado = Me.lstAnalitos.ListIndex
    Me.lstAnalitos.Clear
    If Not IsNumeric(Me.txtAno.Text) Then Exit Sub
    If Trim$(CStr(Me.cboFonte.Value)) = "" Then Exit Sub
    ano = CLng(Me.txtAno.Text)
    Set c = ListaAnalitos()
    For i = 1 To c.Count
        Me.lstAnalitos.AddItem c(i)
        r = ResolverEspec(CStr(c(i)), ano, CStr(Me.cboFonte.Value))
        If Left$(r, 2) = "OK" Then
            p = Split(r, "|")
            Me.lstAnalitos.List(i - 1, 1) = "ano " & p(5)
            Me.lstAnalitos.List(i - 1, 2) = FmtN(p(1))
            Me.lstAnalitos.List(i - 1, 3) = FmtN(p(2))
            Me.lstAnalitos.List(i - 1, 4) = FmtN(p(3))
        Else
            Me.lstAnalitos.List(i - 1, 1) = "-"
        End If
    Next i
    If guardado >= 0 And guardado < Me.lstAnalitos.ListCount Then Me.lstAnalitos.ListIndex = guardado
End Sub

' Val() e o par invariante de Str(): o motor devolve sempre com ponto.
Private Function FmtN(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then FmtN = "-" Else FmtN = Format$(Val(s), "0.00")
End Function

Private Sub AtualizarPrevia()
    Dim modelo As String, m As String, p() As String
    modelo = ModeloDaFonte(Me.cboFonte.Value)
    If modelo = "" Then Me.lblResultado.Caption = "": Exit Sub
    Select Case modelo
        Case "ETP_DIRETO"
            m = MetasDaLinha(modelo, NumOuVazio(Me.txtEnt1.Text), "", "", "", "", "")
        Case "VB"
            m = MetasDaLinha(modelo, "", NumOuVazio(Me.txtEnt1.Text), NumOuVazio(Me.txtEnt2.Text), _
                             CStr(Me.cboRigor.Value), "", "")
        Case "CV_BIAS_DIRETO"
            m = MetasDaLinha(modelo, "", "", "", "", NumOuVazio(Me.txtEnt1.Text), NumOuVazio(Me.txtEnt2.Text))
    End Select
    p = Split(m, "|")
    Me.lblResultado.Caption = "CVtp   " & FmtN(p(0)) & vbLf & _
                              "BIAStp " & FmtN(p(1)) & vbLf & _
                              "ETp    " & FmtN(p(2))
End Sub

' Aceita virgula ou ponto na DIGITACAO -- o usuario digita como quiser. O que
' nao pode depender de localidade e o protocolo interno, nao a tela.
Private Function NumOuVazio(ByVal s As String) As Variant
    Dim t As String
    t = Trim$(Replace(s, ",", "."))
    If t = "" Then NumOuVazio = "" Else NumOuVazio = Val(t)
End Function

Private Sub cboFonte_Change()
    If carregando Then Exit Sub
    AjustarCampos
    CarregarLista
End Sub

Private Sub txtAno_Change()
    If carregando Then Exit Sub
    If IsNumeric(Me.txtAno.Text) Then CarregarLista
End Sub

Private Sub txtEnt1_Change()
    If Not carregando Then AtualizarPrevia
End Sub
Private Sub txtEnt2_Change()
    If Not carregando Then AtualizarPrevia
End Sub
Private Sub cboRigor_Change()
    If Not carregando Then AtualizarPrevia
End Sub

Private Sub btnGravar_Click()
    GravarAtual
End Sub

' Ponto de entrada PUBLICO da gravacao.
'
' O manipulador de clique e Private -- e tem de ser, senao vira API. Sem um
' ponto publico, o teste automatizado nao consegue exercitar o mesmo caminho do
' botao e acabaria testando uma copia da logica, que e justamente o que este
' projeto evita.
Public Function GravarAtual() As Boolean
    GravarAtual = Gravar()
    If GravarAtual Then CarregarLista
End Function

Private Sub btnGravarLista_Click()
    Dim i As Long
    i = Me.lstAnalitos.ListIndex
    If Not GravarAtual() Then Exit Sub
    If i + 1 < Me.lstAnalitos.ListCount Then
        Me.lstAnalitos.ListIndex = i + 1
        Me.txtEnt1.SetFocus
    End If
End Sub

Private Function Gravar() As Boolean
    Dim modelo As String, ano As Long, analito As String
    Me.lblAviso.Caption = ""

    If Trim$(CStr(Me.cboFonte.Value)) = "" Then
        Me.lblAviso.Caption = "Cadastre e selecione uma fonte."
        Exit Function
    End If
    If Me.lstAnalitos.ListIndex < 0 Then
        Me.lblAviso.Caption = "Escolha um analito na lista."
        Exit Function
    End If
    If Not IsNumeric(Me.txtAno.Text) Then
        Me.lblAviso.Caption = "Ano invalido."
        Exit Function
    End If
    ano = CLng(Me.txtAno.Text)
    If ano < 1990 Or ano > 2200 Then
        Me.lblAviso.Caption = "Ano fora de faixa plausivel."
        Exit Function
    End If
    analito = CStr(Me.lstAnalitos.List(Me.lstAnalitos.ListIndex, 0))
    modelo = ModeloDaFonte(Me.cboFonte.Value)
    If modelo = "" Then
        Me.lblAviso.Caption = "A fonte selecionada nao possui modelo suportado."
        Exit Function
    End If
    If modelo = "VB" And Trim$(CStr(Me.cboRigor.Value)) = "" Then
        Me.lblAviso.Caption = "Selecione o rigor da fonte VB."
        Exit Function
    End If

    ' Recusa gravar sem entrada: linha sem valor viraria meta indefinida que o
    ' motor devolveria como "sem meta" -- pior que nao existir, porque parece
    ' cadastrada.
    Dim m As String, p() As String
    Select Case modelo
        Case "ETP_DIRETO"
            m = MetasDaLinha(modelo, NumOuVazio(Me.txtEnt1.Text), "", "", "", "", "")
        Case "VB"
            m = MetasDaLinha(modelo, "", NumOuVazio(Me.txtEnt1.Text), NumOuVazio(Me.txtEnt2.Text), CStr(Me.cboRigor.Value), "", "")
        Case "CV_BIAS_DIRETO"
            m = MetasDaLinha(modelo, "", "", "", "", NumOuVazio(Me.txtEnt1.Text), NumOuVazio(Me.txtEnt2.Text))
    End Select
    p = Split(m, "|")
    If Len(p(0)) = 0 And Len(p(2)) = 0 Then
        Me.lblAviso.Caption = "Preencha os campos desta fonte antes de gravar."
        Exit Function
    End If

    Select Case modelo
        Case "ETP_DIRETO"
            GravarEspec ano, CStr(Me.cboFonte.Value), analito, NumOuVazio(Me.txtEnt1.Text), "", "", "", "", ""
        Case "VB"
            GravarEspec ano, CStr(Me.cboFonte.Value), analito, "", NumOuVazio(Me.txtEnt1.Text), _
                        NumOuVazio(Me.txtEnt2.Text), CStr(Me.cboRigor.Value), "", ""
        Case "CV_BIAS_DIRETO"
            GravarEspec ano, CStr(Me.cboFonte.Value), analito, "", "", "", "", _
                        NumOuVazio(Me.txtEnt1.Text), NumOuVazio(Me.txtEnt2.Text)
    End Select
    Me.lblAviso.Caption = analito & " gravado para " & ano & "."
    Gravar = True
End Function

Private Sub btnFechar_Click()
    Unload Me
End Sub
'@
    $f.CodeModule.InsertLines(1, $codigo)
    "frmEspecificacoes: $($des.Controls.Count) controles, $($f.CodeModule.CountOfLines) linhas"

    # ---- ponto de entrada -----------------------------------------------
    $mod = $null
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq 'mEspecificacoes') { $mod = $c } }
    if ($mod -eq $null) { throw "mEspecificacoes nao esta no projeto" }
    $txt = $mod.CodeModule.Lines(1, $mod.CodeModule.CountOfLines)
    if ($txt -notmatch 'AbrirFormEspecificacoes') {
        $mod.CodeModule.InsertLines(($mod.CodeModule.CountOfLines + 1), @(
            "",
            "' Ponto de entrada do botao da aba Analitos.",
            "Public Sub AbrirFormEspecificacoes()",
            "    frmEspecificacoes.Show",
            "End Sub"
        ) -join "`r`n")
        "AbrirFormEspecificacoes acrescentado a mEspecificacoes"
    }

    # ---- botao na aba Analitos ------------------------------------------
    $an = $null
    foreach ($ws in @($wb.Worksheets)) { if ($ws.Name -like 'Analitos*') { $an = $ws } }
    if ($an -eq $null) { throw "aba Analitos nao encontrada" }
    $protAn = $an.ProtectContents
    if ($protAn) { try { $an.Unprotect($SENHA) } catch { } }
    foreach ($sh in @($an.Shapes)) { if ($sh.Name -eq 'btnEspec') { $sh.Delete() } }
    $b = $an.Shapes.AddShape(5, 520, 4, 190, 30)
    $b.Name = 'btnEspec'
    $b.TextFrame2.TextRange.Text = 'Especificacoes de Qualidade'
    $b.TextFrame2.TextRange.Font.Size = 10
    $b.TextFrame2.TextRange.Font.Bold = $true
    $b.Fill.ForeColor.RGB = 12419407
    $b.Line.Visible = $false
    $b.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    $b.OnAction = 'AbrirFormEspecificacoes'
    if ($protAn) { $an.Protect($SENHA, $true, $true, $true, $true) }
    "botao 'Especificacoes de Qualidade' na aba Analitos"

    if ($wb.ProtectStructure -eq $false) { $wb.Protect($SENHA, $true, $false) }
    $wb.Save()
    $salvou = $true
    "SALVO: $Workbook"
}
finally {
    try { if ($salvou) { $wb.Close($true) } else { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
