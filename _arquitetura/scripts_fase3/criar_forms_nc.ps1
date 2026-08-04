# criar_forms_nc.ps1 - Sprint NC: os dois formularios
#
#   frmResultadoNaoConforme  (aba Resultados)  registra a nao conformidade
#   frmExcluirRegistroNC     (aba Registros)   remove da vitrine, nunca do banco
#
# O NOME DO PRIMEIRO E "RESULTADO", NAO "RUN". Quem fica nao conforme e o
# resultado de um analito; a corrida continua valida. Se a RUN 258 tem PLT ruim
# e o resto bom, chamar a corrida inteira de nao conforme seria falso -- e
# apareceria assim no log e diante do auditor.
#
# O FORMULARIO NAO PEDE O QUE O SISTEMA JA SABE. Informado o RUN, ele preenche
# sozinho data e lote, e traz os resultados da corrida com o valor de cada
# analito. O usuario so escolhe o que esta errado e explica por que.
#
# PARECER TECNICO OBRIGATORIO, minimo de 5 palavras REAIS -- espacos multiplos
# e tokens sem letra nem digito nao contam. A contagem aparece na tela enquanto
# se digita, para o botao desabilitado nunca ser um misterio.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso:
#   .\criar_forms_nc.ps1 -Workbook <build.xlsm>

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath

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

# --------------------------------------------------------------- codigo -----
$codNC = @'
Option Explicit

Private analitos() As String
Private nAnal As Long
Private carregando As Boolean

Private Sub UserForm_Initialize()
    Dim t As Collection, i As Long
    Me.Caption = "Registrar Resultado Nao Conforme"
    Me.Width = 430
    Me.Height = 430

    Me.cboNivel.Clear
    For i = 1 To NLV
        Me.cboNivel.AddItem CStr(i)
    Next i
    Me.cboNivel.ListIndex = 0

    ' Tipos vem do Cfg_Status (estados NAO elegiveis). Nenhum escrito em codigo:
    ' acrescentar um tipo continua sendo digitar uma linha na tabela (ADR-006).
    Me.cboTipo.Clear
    Set t = TiposNaoConformidade()
    For i = 1 To t.Count
        Me.cboTipo.AddItem t(i)
    Next i
    If Me.cboTipo.ListCount > 0 Then Me.cboTipo.ListIndex = 0

    Me.lstAnalitos.MultiSelect = fmMultiSelectMulti
    Me.lblInfo.Caption = "Informe o RUN. Data, lote e resultados sao preenchidos pelo sistema."
    AtualizarContagem
End Sub

Private Sub txtRun_Change()
    CarregarCorrida
End Sub

Private Sub cboNivel_Change()
    CarregarCorrida
End Sub

' O RUN e a chave: com ele o sistema sabe data, lote e todos os resultados.
Private Sub CarregarCorrida()
    Dim run As Long, cab As String, p As Variant, res As Collection, i As Long, campos As Variant
    If carregando Then Exit Sub
    carregando = True
    On Error GoTo fim

    Me.lstAnalitos.Clear
    nAnal = 0
    Erase analitos

    run = Val(Trim$(Me.txtRun.Value))
    If run <= 0 Then
        Me.lblCabecalho.Caption = ""
        GoTo fim
    End If

    cab = CabecalhoDaCorrida(run)
    If cab = "" Then
        Me.lblCabecalho.Caption = "RUN " & run & " nao encontrado."
        GoTo fim
    End If
    p = Split(cab, "|")
    Me.lblCabecalho.Caption = "Data: " & p(0) & "     Lote: " & p(1)

    Set res = ResultadosDaCorrida(run, CLng(Me.cboNivel.Value))
    If res.Count > 0 Then ReDim analitos(1 To res.Count)
    For i = 1 To res.Count
        campos = Split(res(i), "|")
        nAnal = nAnal + 1
        analitos(nAnal) = CStr(campos(0))
        Me.lstAnalitos.AddItem CStr(campos(0)) & "   =  " & CStr(campos(2)) & _
                               IIf(CStr(campos(3)) <> "Ativo", "   [" & CStr(campos(3)) & "]", "")
    Next i
fim:
    carregando = False
End Sub

Private Sub txtParecer_Change()
    AtualizarContagem
End Sub

' A contagem na tela evita o pior tipo de bloqueio: o botao desabilitado sem
' explicacao.
Private Sub AtualizarContagem()
    Dim n As Long
    n = ContarPalavras(Me.txtParecer.Value)
    If n >= PARECER_MIN_PALAVRAS Then
        Me.lblParecer.Caption = "Parecer tecnico (" & n & " palavras) - ok"
    Else
        Me.lblParecer.Caption = "Parecer tecnico - obrigatorio, minimo " & _
                                PARECER_MIN_PALAVRAS & " palavras (" & n & " ate agora)"
    End If
End Sub

Private Sub btnTodos_Click()
    Dim i As Long
    For i = 0 To Me.lstAnalitos.ListCount - 1
        Me.lstAnalitos.Selected(i) = True
    Next i
End Sub

Private Sub btnSalvar_Click()
    Dim run As Long, nivel As Long, i As Long, n As Long, ids As String, idEv As String

    run = Val(Trim$(Me.txtRun.Value))
    If run <= 0 Then MsgBox "Informe o RUN.", vbExclamation: Exit Sub
    If CabecalhoDaCorrida(run) = "" Then MsgBox "RUN " & run & " nao encontrado.", vbExclamation: Exit Sub
    nivel = CLng(Me.cboNivel.Value)

    n = 0
    For i = 0 To Me.lstAnalitos.ListCount - 1
        If Me.lstAnalitos.Selected(i) Then n = n + 1
    Next i
    If n = 0 Then MsgBox "Selecione ao menos um analito.", vbExclamation: Exit Sub

    If Me.cboTipo.ListIndex = -1 Then MsgBox "Selecione o tipo de nao conformidade.", vbExclamation: Exit Sub

    If Not ParecerValido(Me.txtParecer.Value) Then
        MsgBox "O parecer tecnico precisa de pelo menos " & PARECER_MIN_PALAVRAS & _
               " palavras." & vbLf & vbLf & _
               "Ele explica a um auditor por que este resultado saiu dos calculos.", _
               vbExclamation, "Parecer insuficiente"
        Me.txtParecer.SetFocus
        Exit Sub
    End If

    If MsgBox("Marcar " & n & " resultado(s) do RUN " & run & ", nivel " & nivel & _
              ", como """ & Me.cboTipo.Value & """?" & vbLf & vbLf & _
              "Eles saem dos calculos estatisticos. O valor original permanece no banco " & _
              "e a operacao fica registrada na auditoria.", _
              vbYesNo + vbQuestion, "Confirmar") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    n = 0
    For i = 0 To Me.lstAnalitos.ListCount - 1
        If Me.lstAnalitos.Selected(i) Then
            idEv = MarcarNaoConforme(run, nivel, analitos(i + 1), Me.cboTipo.Value, Me.txtParecer.Value)
            If idEv <> "" Then
                n = n + 1
                If ids = "" Then ids = idEv
            End If
        End If
    Next i
    AtualizarOperacao
    Application.ScreenUpdating = True

    MsgBox n & " resultado(s) registrado(s) como nao conforme." & vbLf & vbLf & _
           "Primeiro evento de auditoria: " & ids, vbInformation, "Registrado"
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub
'@

$codEx = @'
Option Explicit

Private linhas() As Long
Private nLin As Long

Private Sub UserForm_Initialize()
    Dim c As Collection, i As Long, campos As Variant
    Me.Caption = "Excluir Registro de Nao Conformidade"
    Me.Width = 430
    Me.Height = 360
    Me.lstRegistros.Clear
    Set c = ListaRegistrosNC()
    nLin = 0
    If c.Count > 0 Then ReDim linhas(1 To c.Count)
    For i = 1 To c.Count
        campos = Split(c(i), "|")
        nLin = nLin + 1
        linhas(nLin) = CLng(campos(0))
        Me.lstRegistros.AddItem campos(1) & "   RUN " & campos(4) & "   " & _
                                campos(2) & "  N" & campos(3) & "  =  " & campos(5) & _
                                "   [" & campos(6) & "]"
    Next i
    Me.lblInfo.Caption = c.Count & " ocorrencia(s). A exclusao remove da aba Registros; " & _
                         "a linha permanece no banco e na auditoria."
    AtualizarContagem
End Sub

Private Sub txtParecer_Change()
    AtualizarContagem
End Sub

Private Sub AtualizarContagem()
    Dim n As Long
    n = ContarPalavras(Me.txtParecer.Value)
    If n >= PARECER_MIN_PALAVRAS Then
        Me.lblParecer.Caption = "Justificativa (" & n & " palavras) - ok"
    Else
        Me.lblParecer.Caption = "Justificativa - obrigatoria, minimo " & _
                                PARECER_MIN_PALAVRAS & " palavras (" & n & " ate agora)"
    End If
End Sub

Private Sub btnExcluir_Click()
    Dim i As Long, idEv As String
    If Me.lstRegistros.ListIndex = -1 Then MsgBox "Selecione o registro.", vbExclamation: Exit Sub
    If Not ParecerValido(Me.txtParecer.Value) Then
        MsgBox "A justificativa precisa de pelo menos " & PARECER_MIN_PALAVRAS & " palavras.", _
               vbExclamation, "Justificativa insuficiente"
        Me.txtParecer.SetFocus
        Exit Sub
    End If

    If MsgBox("Deseja realmente excluir este registro da aba Registros?" & vbLf & vbLf & _
              Me.lstRegistros.Value & vbLf & vbLf & _
              "O resultado PERMANECE no banco e a exclusao fica registrada na auditoria.", _
              vbYesNo + vbExclamation, "Confirmar exclusao") <> vbYes Then Exit Sub

    i = Me.lstRegistros.ListIndex + 1
    idEv = ExcluirRegistroNC(linhas(i), Me.txtParecer.Value)
    If idEv = "" Then
        MsgBox "Nao foi possivel excluir.", vbCritical
        Exit Sub
    End If
    MsgBox "Registro removido da aba Registros." & vbLf & vbLf & _
           "Evento de auditoria: " & idEv, vbInformation, "Excluido"
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub
'@

# ------------------------------------------------------------ construcao ----
$xl = Novo-Excel
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1

$wb = $xl.Workbooks.Open($Workbook)
try { $wb.EnableAutoRecover = $false } catch { }
if ($wb.ReadOnly) {
    $wb.Close($false); $xl.Quit()
    throw "Arquivo aberto em SOMENTE LEITURA: $Workbook"
}

try {
    $proj = $wb.VBProject

    # O VBA so libera um nome de formulario DEPOIS de salvar. Remover e recriar
    # na mesma sessao faz o segundo nascer como "frmX1" -- e o codigo que chama
    # pelo nome certo passa a nao encontrar nada.
    $removeu = $false
    foreach ($nome in @('frmResultadoNaoConforme', 'frmExcluirRegistroNC')) {
        foreach ($c in @($proj.VBComponents)) {
            if ($c.Name -eq $nome) { $proj.VBComponents.Remove($c); $removeu = $true }
        }
    }
    if ($removeu) {
        $wb.Save()
        $wb.Close($true)
        $xl.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        Start-Sleep -Seconds 2
        $xl = Novo-Excel
        $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
        $xl.AutomationSecurity = 1
        $wb = $xl.Workbooks.Open($Workbook)
        try { $wb.EnableAutoRecover = $false } catch { }
        $proj = $wb.VBProject
        "  formularios anteriores removidos (recriados apos salvar)"
    }

    # ---------------- frmResultadoNaoConforme ----------------
    $f = $proj.VBComponents.Add(3)          # vbext_ct_MSForm
    $f.Name = 'frmResultadoNaoConforme'
    # VBComponent.Properties nao e acessivel pelo PowerShell ("nao foi possivel
    # encontrar o membro"). Legenda e tamanho sao definidos no proprio
    # UserForm_Initialize, que roda no Excel e tem acesso direto a Me.
    $d = $f.Designer

    function Add-Ctl {
        param($Des, [string]$Tipo, [string]$Nome, [int]$L, [int]$T, [int]$W, [int]$H, [string]$Cap = $null)
        $c = $Des.Controls.Add($Tipo, $Nome, $true)
        $c.Left = $L; $c.Top = $T; $c.Width = $W; $c.Height = $H
        if ($Cap -ne $null) { try { $c.Caption = $Cap } catch { } }
        return $c
    }

    Add-Ctl $d 'Forms.Label.1' 'lblInfo' 12 8 400 14 '' | Out-Null
    Add-Ctl $d 'Forms.Label.1' 'lblRun' 12 30 40 14 'RUN:' | Out-Null
    Add-Ctl $d 'Forms.TextBox.1' 'txtRun' 52 28 60 18 | Out-Null
    Add-Ctl $d 'Forms.Label.1' 'lblNivelCap' 128 30 36 14 'Nivel:' | Out-Null
    Add-Ctl $d 'Forms.ComboBox.1' 'cboNivel' 166 28 50 18 | Out-Null
    Add-Ctl $d 'Forms.Label.1' 'lblCabecalho' 12 52 400 14 '' | Out-Null
    Add-Ctl $d 'Forms.Label.1' 'lblLista' 12 74 260 14 'Resultados da corrida (selecione os nao conformes):' | Out-Null
    Add-Ctl $d 'Forms.ListBox.1' 'lstAnalitos' 12 90 280 150 | Out-Null
    Add-Ctl $d 'Forms.CommandButton.1' 'btnTodos' 300 90 110 22 'Selecionar todos' | Out-Null
    Add-Ctl $d 'Forms.Label.1' 'lblTipoCap' 300 122 110 14 'Tipo de NC:' | Out-Null
    Add-Ctl $d 'Forms.ComboBox.1' 'cboTipo' 300 138 110 18 | Out-Null
    Add-Ctl $d 'Forms.Label.1' 'lblParecer' 12 248 400 14 '' | Out-Null
    $tp = Add-Ctl $d 'Forms.TextBox.1' 'txtParecer' 12 264 398 86
    $tp.MultiLine = $true
    $tp.ScrollBars = 2
    Add-Ctl $d 'Forms.CommandButton.1' 'btnSalvar' 232 362 88 26 'Salvar' | Out-Null
    Add-Ctl $d 'Forms.CommandButton.1' 'btnCancelar' 326 362 84 26 'Cancelar' | Out-Null

    $f.CodeModule.AddFromString($codNC)
    "  frmResultadoNaoConforme: $($d.Controls.Count) controles, $($f.CodeModule.CountOfLines) linhas"

    # ---------------- frmExcluirRegistroNC ----------------
    $g = $proj.VBComponents.Add(3)
    $g.Name = 'frmExcluirRegistroNC'
    $e = $g.Designer

    Add-Ctl $e 'Forms.Label.1' 'lblInfo' 12 8 400 14 '' | Out-Null
    Add-Ctl $e 'Forms.ListBox.1' 'lstRegistros' 12 28 398 150 | Out-Null
    Add-Ctl $e 'Forms.Label.1' 'lblParecer' 12 186 400 14 '' | Out-Null
    $tp2 = Add-Ctl $e 'Forms.TextBox.1' 'txtParecer' 12 202 398 76
    $tp2.MultiLine = $true
    $tp2.ScrollBars = 2
    Add-Ctl $e 'Forms.CommandButton.1' 'btnExcluir' 232 290 88 26 'Excluir' | Out-Null
    Add-Ctl $e 'Forms.CommandButton.1' 'btnCancelar' 326 290 84 26 'Cancelar' | Out-Null

    $g.CodeModule.AddFromString($codEx)
    "  frmExcluirRegistroNC: $($e.Controls.Count) controles, $($g.CodeModule.CountOfLines) linhas"

    $wb.Save()

    # conferir, nao confiar
    foreach ($nome in @('frmResultadoNaoConforme', 'frmExcluirRegistroNC')) {
        $achou = $false
        foreach ($c in $proj.VBComponents) { if ($c.Name -eq $nome) { $achou = $true } }
        if (-not $achou) { throw "$nome nao ficou no projeto apos salvar" }
    }
    "formularios da Sprint NC criados e conferidos"
}
finally {
    try { $wb.Close($true) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
