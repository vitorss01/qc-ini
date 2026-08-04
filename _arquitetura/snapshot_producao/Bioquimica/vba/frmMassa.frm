VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMassa 
   Caption         =   "Inserção em massa"
   ClientHeight    =   8832.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   14190
   OleObjectBlob   =   "frmMassa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMassa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private analNames() As String
Private analCount As Long

Private Sub UserForm_Initialize()
    Dim c As Collection, i As Long, h As String
    Set c = ListaAnalitos()
    analCount = c.Count
    ReDim analNames(1 To analCount)
    h = "Data" & vbTab & "Nível" & vbTab & "Lote"
    For i = 1 To analCount
        analNames(i) = c(i)
        h = h & vbTab & c(i)
    Next i
    Me.lblColunas.Caption = Replace(h, vbTab, "  |  ")
    Me.lblStatus.Caption = "Cole aqui (Ctrl+V) linhas copiadas do Excel — colunas separadas por TAB, na ordem acima."
    Me.barFundo.Visible = False
    Me.barProg.Visible = False
End Sub

' Divide o texto colado em matriz de campos.
Private Function LerGrade(ByRef linhas() As String, ByRef nLin As Long) As Boolean
    Dim raw As String, parts As Variant, i As Long, cnt As Long
    raw = Me.txtGrade.Value
    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    parts = Split(raw, vbLf)
    ReDim linhas(0 To UBound(parts))
    cnt = 0
    For i = 0 To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            linhas(cnt) = parts(i)
            cnt = cnt + 1
        End If
    Next i
    nLin = cnt
    LerGrade = (cnt > 0)
End Function

' Valida TUDO acumulando erros (nunca para no primeiro).
Private Function Validar(ByRef regs() As Variant, ByRef nReg As Long) As Boolean
    Dim linhas() As String, nLin As Long, i As Long, j As Long
    Dim campos As Variant, ok As Boolean, dt As Date, lvl As Long, lote As String
    Dim v As Double, erros As Long, mx As Long, run As Long
    Me.lstErros.Clear
    erros = 0: nReg = 0
    If Not LerGrade(linhas, nLin) Then
        Me.lstErros.AddItem "Nada para validar — a grade está vazia."
        Validar = False: Exit Function
    End If
    mx = nLin * analCount
    ReDim regs(1 To mx, 1 To 7)

    For i = 0 To nLin - 1
        campos = Split(linhas(i), vbTab)
        If UBound(campos) < 2 Then
            Me.lstErros.AddItem "Linha " & (i + 1) & " · estrutura · esperado ao menos Data, Nível e Lote"
            erros = erros + 1
            GoTo proxima
        End If
        dt = ParseData(CStr(campos(0)), ok)
        If Not ok Then
            Me.lstErros.AddItem "Linha " & (i + 1) & " · coluna 1 (Data) · data inválida: """ & campos(0) & """"
            erros = erros + 1
        End If
        If Not IsNumeric(Trim$(CStr(campos(1)))) Then
            Me.lstErros.AddItem "Linha " & (i + 1) & " · coluna 2 (Nível) · nível inválido: """ & campos(1) & """"
            erros = erros + 1
            lvl = 0
        Else
            lvl = CLng(Trim$(CStr(campos(1))))
            If lvl < 1 Or lvl > 2 Then
                Me.lstErros.AddItem "Linha " & (i + 1) & " · coluna 2 (Nível) · fora do intervalo 1..2: " & lvl
                erros = erros + 1
            End If
        End If
        lote = Trim$(CStr(campos(2)))
        If lote = "" Then
            Me.lstErros.AddItem "Linha " & (i + 1) & " · coluna 3 (Lote) · lote não preenchido"
            erros = erros + 1
        ElseIf Not LoteExiste(lote) Then
            Me.lstErros.AddItem "Linha " & (i + 1) & " · coluna 3 (Lote) · lote não cadastrado: """ & lote & """"
            erros = erros + 1
        End If

        If ok And lvl >= 1 And lvl <= 2 And lote <> "" Then
            run = NovoRUN(dt, lote)
            For j = 1 To analCount
                If (j + 2) <= UBound(campos) Then
                    Dim s As String
                    s = Trim$(CStr(campos(j + 2)))
                    If s <> "" Then
                        v = ParseNum(s, ok)
                        If Not ok Then
                            Me.lstErros.AddItem "Linha " & (i + 1) & " · coluna " & (j + 3) & " (" & analNames(j) & _
                                                ") · resultado não numérico: """ & s & """"
                            erros = erros + 1
                        Else
                            nReg = nReg + 1
                            regs(nReg, 1) = run
                            regs(nReg, 2) = dt
                            regs(nReg, 3) = lvl
                            regs(nReg, 4) = CodigoLote(lote, lvl)
                            regs(nReg, 5) = analNames(j)
                            regs(nReg, 6) = v
                            regs(nReg, 7) = ST_ATIVO
                        End If
                    End If
                End If
            Next j
        End If
proxima:
    Next i

    If erros = 0 Then
        Me.lstErros.AddItem "OK — " & nLin & " linha(s), " & nReg & " resultado(s) prontos para gravar."
        Me.lblStatus.Caption = "Validação concluída sem inconsistências."
    Else
        Me.lblStatus.Caption = erros & " inconsistência(s) encontrada(s). Nada foi gravado."
    End If
    Validar = (erros = 0 And nReg > 0)
End Function

Private Function LoteExiste(ByVal lote As String) As Boolean
    Dim c As Collection, i As Long
    Set c = ListaLotes()
    For i = 1 To c.Count
        If Trim$(CStr(c(i))) = lote Then LoteExiste = True: Exit Function
    Next i
End Function

Private Sub btnValidar_Click()
    Dim regs() As Variant, n As Long
    Validar regs, n
End Sub

Private Sub btnLimpar_Click()
    Me.txtGrade.Value = ""
    Me.lstErros.Clear
    Me.lblStatus.Caption = "Grade limpa."
End Sub

Private Sub btnSalvar_Click()
    Dim regs() As Variant, n As Long, res As String, final_() As Variant, i As Long, c As Long
    If Not Validar(regs, n) Then
        MsgBox "Corrija as inconsistências antes de gravar.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Gravar " & n & " resultado(s) em DB_Resultados?" & vbCrLf & vbCrLf & _
              "Registros já existentes (mesmo RUN + Nível + Analito) serão ATUALIZADOS, não duplicados.", _
              vbQuestion + vbYesNo, "Confirmar importação") <> vbYes Then Exit Sub

    Me.barFundo.Visible = True: Me.barProg.Visible = True
    Me.barProg.Width = 1
    Me.lblStatus.Caption = "Gravando..."
    Me.Repaint

    ReDim final_(1 To n, 1 To 7)
    For i = 1 To n
        For c = 1 To 7
            final_(i, c) = regs(i, c)
        Next c
        If (i Mod 50) = 0 Then
            Me.barProg.Width = Me.barFundo.Width * (i / n)
            Me.Repaint
        End If
    Next i
    Me.barProg.Width = Me.barFundo.Width: Me.Repaint

    res = UpsertResultados(final_)
    RegistrarLog "IMPORTACAO", n & " registros"
    AtualizarOperacao
    Me.barFundo.Visible = False: Me.barProg.Visible = False
    MsgBox "Importação concluída." & vbCrLf & _
           "Novos: " & Split(res, "|")(0) & "   ·   Atualizados: " & Split(res, "|")(1), vbInformation
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub


