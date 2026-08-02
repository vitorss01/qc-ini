# -*- coding: utf-8 -*-
"""FASE 2 / F2-3..F2-5: frmCorrida (atualizado), frmExcluir e frmMassa."""
import os, sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client as w

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
PWD = "qcini2025"
FILES = [("QC_Hematologia.xlsm", 3), ("QC_Bioquimica.xlsm", 2), ("QC_Imunologia.xlsm", 2)]

def rc(fn, t=14):
    last=None
    for _ in range(t):
        try: return fn()
        except Exception as e: last=e; time.sleep(0.4)
    raise last


def del_comp(vbp, nome, espera=25):
    for c in list(vbp.VBComponents):
        if c.Name == nome:
            try: vbp.VBComponents.Remove(c)
            except Exception: pass
    for _ in range(espera):
        if nome not in [c.Name for c in vbp.VBComponents]: return True
        time.sleep(0.3)
    return False

def safe_add(vbp, fn, nome, tries=5):
    for _ in range(tries):
        antes = set(c.Name for c in vbp.VBComponents)
        if nome in antes: del_comp(vbp, nome)
        try:
            return fn()
        except Exception as e:
            for c in list(vbp.VBComponents):
                if c.Name not in antes and c.Name != nome:
                    try: vbp.VBComponents.Remove(c)   # limpa orfao da tentativa
                    except Exception: pass
            last = e; time.sleep(1.2)
    raise RuntimeError("falha ao criar %s: %r" % (nome, last))


def remove_forms(fname):
    """Pass A: remove os forms antigos e SALVA. O VBA so libera o nome apos salvar."""
    xl = rc(lambda: w.DispatchEx("Excel.Application"))
    xl.Visible=False; xl.DisplayAlerts=False; xl.EnableEvents=False
    try: xl.AutomationSecurity=1
    except: pass
    wb = rc(lambda: xl.Workbooks.Open(os.path.join(BASE, fname)))
    try:
        vbp = wb.VBProject
        alvos = ("frmCorrida","frmExcluir","frmMassa","mEntrada")
        rem = []
        for nome in alvos:
            if del_comp(vbp, nome): rem.append(nome)
        # limpa qualquer orfao UserFormN de tentativas anteriores
        for c in list(vbp.VBComponents):
            if c.Type == 3 and c.Name.startswith("UserForm"):
                try: vbp.VBComponents.Remove(c); rem.append(c.Name)
                except Exception: pass
        rc(lambda: wb.Save())
        print("  [pass A] removidos e salvos:", rem, flush=True)
    finally:
        try: wb.Close(False)
        except: pass
        try: xl.Quit()
        except: pass
    time.sleep(1.5)

# ---------------- utilitarios compartilhados (modulo) ----------------
MOD_UTIL = '''Option Explicit
' ===== UTILITARIOS DE ENTRADA (Fase 2) =====

' Aceita "12,5" e "12.5"; devolve ok=False se nao for numero.
Public Function ParseNum(ByVal s As String, ByRef ok As Boolean) As Double
    Dim t As String
    ok = False
    t = Trim$(s)
    If t = "" Then Exit Function
    t = Replace(t, " ", "")
    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then
        t = Replace(t, ".", ""): t = Replace(t, ",", ".")
    ElseIf InStr(t, ",") > 0 Then
        t = Replace(t, ",", ".")
    End If
    If Not IsNumeric(t) Then Exit Function
    ParseNum = Val(t)
    ok = True
End Function

Public Function ParseData(ByVal s As String, ByRef ok As Boolean) As Date
    Dim t As String
    ok = False
    t = Trim$(s)
    If t = "" Then Exit Function
    On Error Resume Next
    If IsDate(t) Then
        ParseData = CDate(t)
        ok = (Err.Number = 0)
    End If
    On Error GoTo 0
End Function

' Codigo completo do lote a partir do nucleo de 6 digitos + nivel.
Public Function CodigoLote(ByVal loteCore As String, ByVal nivel As Long) As String
    CodigoLote = "QC-" & loteCore & Format(nivel, "00")
End Function

Public Function NucleoLote(ByVal codigo As String) As String
    NucleoLote = Mid$(Trim$(codigo), 4, 6)
End Function
'''

# ---------------- frmCorrida ----------------
FRM_CORRIDA = '''Option Explicit
Private analNames() As String
Private analCount As Long

Private Sub UserForm_Initialize()
    Dim c As Collection, i As Long, lote As String
    Set c = ListaAnalitos()
    analCount = c.Count
    If analCount > 0 Then
        ReDim analNames(1 To analCount)
        For i = 1 To analCount
            analNames(i) = c(i)
        Next i
    End If

    Me.txtData.Value = Format(Date, "dd/mm/yyyy")
    Me.cboNivel.Clear
{NIVEL_ITEMS}
    Me.cboNivel.ListIndex = 0

    Dim lotes As Collection
    Set lotes = ListaLotes()
    Me.cboLote.Clear
    For i = 1 To lotes.Count
        Me.cboLote.AddItem lotes(i)
    Next i
    lote = LoteAtivoCore()
    For i = 0 To Me.cboLote.ListCount - 1
        If Me.cboLote.List(i) = lote Then Me.cboLote.ListIndex = i
    Next i
    If Me.cboLote.ListIndex = -1 And Me.cboLote.ListCount > 0 Then Me.cboLote.ListIndex = 0

    Dim col As Long, rowi As Long, lbl As Object, txt As Object
    Dim baseTop As Long, rowH As Long, col1L As Long, col2L As Long, txtW As Long, lblW As Long, ti As Long
    baseTop = 74: rowH = 20: lblW = 88: txtW = 62: col1L = 12: col2L = 202: ti = 10
    For i = 1 To analCount
        rowi = (i - 1) \\ 2
        col = (i - 1) Mod 2
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblV" & i, True)
        lbl.Caption = analNames(i)
        lbl.Top = baseTop + rowi * rowH
        lbl.Left = IIf(col = 0, col1L, col2L)
        lbl.Width = lblW: lbl.Height = 17: lbl.Font.Size = 9
        Set txt = Me.Controls.Add("Forms.TextBox.1", "txtV" & i, True)
        txt.Top = baseTop + rowi * rowH
        txt.Left = IIf(col = 0, col1L + lblW + 2, col2L + lblW + 2)
        txt.Width = txtW: txt.Height = 18: txt.Font.Size = 9
        txt.TabIndex = ti: ti = ti + 1
    Next i

    Dim rows As Long
    rows = (analCount + 1) \\ 2
    If rows < 1 Then rows = 1
    Me.Height = baseTop + rows * rowH + 96
    Me.btnSalvar.Top = baseTop + rows * rowH + 14
    Me.btnCancelar.Top = Me.btnSalvar.Top
    Me.lblInfo.Top = Me.btnSalvar.Top - 20
    AtualizarRun
End Sub

Private Sub txtData_Change()
    AtualizarRun
End Sub
Private Sub cboLote_Change()
    AtualizarRun
End Sub

Private Sub AtualizarRun()
    Dim dt As Date, ok As Boolean, lote As String
    dt = ParseData(Me.txtData.Value, ok)
    If Not ok Or Me.cboLote.ListIndex = -1 Then
        Me.lblRun.Caption = ""
        Exit Sub
    End If
    lote = Me.cboLote.Value
    Me.lblRun.Caption = "RUN " & NovoRUN(dt, lote)
End Sub

Private Sub btnSalvar_Click()
    Dim dt As Date, ok As Boolean, lvl As Long, lote As String
    Dim i As Long, n As Long, run As Long, regs() As Variant, res As String, v As Double
    dt = ParseData(Me.txtData.Value, ok)
    If Not ok Then MsgBox "Data inválida.", vbExclamation: Exit Sub
    If Me.cboNivel.ListIndex = -1 Then MsgBox "Selecione o nível.", vbExclamation: Exit Sub
    If Me.cboLote.ListIndex = -1 Then MsgBox "Selecione o lote.", vbExclamation: Exit Sub
    lvl = CLng(Me.cboNivel.Value)
    lote = Me.cboLote.Value
    run = NovoRUN(dt, lote)

    n = 0
    ReDim regs(1 To analCount, 1 To 7)
    For i = 1 To analCount
        Dim s As String
        s = Trim$(CStr(Me.Controls("txtV" & i).Value))
        If s <> "" Then
            v = ParseNum(s, ok)
            If Not ok Then
                MsgBox "Valor inválido em " & analNames(i) & ": """ & s & """", vbExclamation
                Exit Sub
            End If
            n = n + 1
            regs(n, 1) = run
            regs(n, 2) = dt
            regs(n, 3) = lvl
            regs(n, 4) = CodigoLote(lote, lvl)
            regs(n, 5) = analNames(i)
            regs(n, 6) = v
            regs(n, 7) = ST_ATIVO
        End If
    Next i
    If n = 0 Then MsgBox "Nenhum valor informado.", vbExclamation: Exit Sub

    If MsgBox("Gravar " & n & " resultado(s) no RUN " & run & " (Nível " & lvl & ", lote " & lote & ")?", _
              vbQuestion + vbYesNo, "Confirmar") <> vbYes Then Exit Sub

    Dim final_() As Variant, c As Long
    ReDim final_(1 To n, 1 To 7)
    For i = 1 To n
        For c = 1 To 7
            final_(i, c) = regs(i, c)
        Next c
    Next i
    res = UpsertResultados(final_)
    RegistrarLog "LANCAMENTO", "RUN " & run & " N" & lvl
    AtualizarOperacao
    MsgBox "RUN " & run & " gravado." & vbCrLf & _
           "Novos: " & Split(res, "|")(0) & "   ·   Atualizados: " & Split(res, "|")(1), vbInformation
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub
'''

# ---------------- frmExcluir ----------------
FRM_EXCLUIR = '''Option Explicit

Private Sub UserForm_Initialize()
    Me.cboNivel.Clear
{NIVEL_ITEMS}
    Me.cboNivel.ListIndex = 0
    CarregarRuns
    CarregarAnalitos
End Sub

Private Sub cboNivel_Change()
    CarregarAnalitos
End Sub

Private Sub CarregarAnalitos()
    Dim c As Collection, i As Long
    Me.lstAnalitos.Clear
    Set c = ListaAnalitos()
    For i = 1 To c.Count
        Me.lstAnalitos.AddItem c(i)
    Next i
    Me.lblInfo.Caption = c.Count & " analito(s) do nível selecionado."
End Sub

Private Sub CarregarRuns()
    Dim runs As Collection, i As Long
    Me.cboRun.Clear
    Set runs = RunsDoLote(LoteAtivoCore())
    For i = 1 To runs.Count
        Me.cboRun.AddItem CStr(runs(i))
    Next i
    If Me.cboRun.ListCount > 0 Then Me.cboRun.ListIndex = Me.cboRun.ListCount - 1
End Sub

Private Sub btnTodos_Click()
    Dim i As Long
    For i = 0 To Me.lstAnalitos.ListCount - 1
        Me.lstAnalitos.Selected(i) = True
    Next i
End Sub

Private Sub btnLimpar_Click()
    Dim i As Long
    For i = 0 To Me.lstAnalitos.ListCount - 1
        Me.lstAnalitos.Selected(i) = False
    Next i
End Sub

Private Sub btnConfirmar_Click()
    Dim i As Long, alvos As Object, n As Long, lvl As Long, run As Long
    If Me.cboRun.ListIndex = -1 Then MsgBox "Selecione a corrida (RUN).", vbExclamation: Exit Sub
    If Me.cboNivel.ListIndex = -1 Then MsgBox "Selecione o nível.", vbExclamation: Exit Sub
    Set alvos = CreateObject("Scripting.Dictionary")
    For i = 0 To Me.lstAnalitos.ListCount - 1
        If Me.lstAnalitos.Selected(i) Then alvos(UCase$(Trim$(Me.lstAnalitos.List(i)))) = 1
    Next i
    If alvos.Count = 0 Then MsgBox "Marque pelo menos um analito.", vbExclamation: Exit Sub
    lvl = CLng(Me.cboNivel.Value)
    run = CLng(Me.cboRun.Value)

    If MsgBox("Excluir logicamente " & alvos.Count & " analito(s) do RUN " & run & _
              " (Nível " & lvl & ")?" & vbCrLf & vbCrLf & _
              "Os registros NÃO são apagados — apenas marcados como Excluído.", _
              vbExclamation + vbYesNo, "Confirmar exclusão") <> vbYes Then Exit Sub

    n = ExcluirLogico(run, lvl, alvos)
    RegistrarLog "EXCLUSAO", "RUN " & run & " N" & lvl & " (" & n & ")"
    AtualizarOperacao
    MsgBox n & " registro(s) marcado(s) como Excluído.", vbInformation
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub
'''

# ---------------- frmMassa ----------------
FRM_MASSA = '''Option Explicit
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
            If lvl < 1 Or lvl > {NLV} Then
                Me.lstErros.AddItem "Linha " & (i + 1) & " · coluna 2 (Nível) · fora do intervalo 1..{NLV}: " & lvl
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

        If ok And lvl >= 1 And lvl <= {NLV} And lote <> "" Then
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
'''

def build(fname, NLV):
    print("="*60); print("==", fname, flush=True)
    xl = rc(lambda: w.DispatchEx("Excel.Application"))
    xl.Visible=False; xl.DisplayAlerts=False; xl.EnableEvents=False
    try: xl.AutomationSecurity=1
    except: pass
    wb = rc(lambda: xl.Workbooks.Open(os.path.join(BASE, fname)))
    try:
        for sh in wb.Worksheets:
            try: sh.Unprotect(Password=PWD)
            except: pass
        # ordem das abas: view logo apos o banco
        # ATENCAO: Move(After=...) por argumento NOMEADO nao vincula no win32com —
        # o Excel entende Move() sem args e joga a aba para uma pasta NOVA (some daqui).
        # Usar SEMPRE posicional (1o arg = Before) e conferir depois.
        try:
            antes = [x.Name for x in wb.Worksheets]
            wb.Sheets("Resultados").Move(wb.Sheets("Painel"))     # Before=Painel (posicional)
            if "Resultados" not in [x.Name for x in wb.Worksheets]:
                raise RuntimeError("a aba Resultados sumiu no Move")
            print("  ordem das abas ajustada (Resultados antes de Painel)", flush=True)
        except Exception as e:
            print("  aviso ordem (ignorado):", repr(e)[:70], flush=True)

        vbp = wb.VBProject
        nivel_items = "".join(f'    Me.cboNivel.AddItem "{k}"\n' for k in range(1, NLV+1))

        def ins_util():
            m = vbp.VBComponents.Add(1); m.Name = "mEntrada"
            m.CodeModule.AddFromString(MOD_UTIL); return 1
        rc(ins_util); time.sleep(0.25)
        print("  + mEntrada (parsers/utilitarios)", flush=True)


        # ---------- frmCorrida ----------
        def mk_corrida():
            f = vbp.VBComponents.Add(3); f.Name = "frmCorrida"
            f.Properties("Caption").Value = "Nova Corrida"
            f.Properties("Width").Value = 430; f.Properties("Height").Value = 320
            d = f.Designer
            t = d.Controls.Add("Forms.Label.1"); t.Caption = "Lançar corrida — controle interno"
            t.Top=8; t.Left=12; t.Width=380; t.Font.Size=13; t.Font.Bold=True
            l1 = d.Controls.Add("Forms.Label.1"); l1.Caption="Data"; l1.Top=38; l1.Left=12; l1.Width=32; l1.Font.Size=10
            td = d.Controls.Add("Forms.TextBox.1"); td.Name="txtData"; td.Top=35; td.Left=46; td.Width=80; td.Height=20; td.Font.Size=10; td.TabIndex=0
            l2 = d.Controls.Add("Forms.Label.1"); l2.Caption="Nível"; l2.Top=38; l2.Left=136; l2.Width=34; l2.Font.Size=10
            cn = d.Controls.Add("Forms.ComboBox.1"); cn.Name="cboNivel"; cn.Top=35; cn.Left=172; cn.Width=46; cn.Height=20; cn.Font.Size=10; cn.Style=2; cn.TabIndex=1
            l3 = d.Controls.Add("Forms.Label.1"); l3.Caption="Lote"; l3.Top=38; l3.Left=228; l3.Width=28; l3.Font.Size=10
            cl = d.Controls.Add("Forms.ComboBox.1"); cl.Name="cboLote"; cl.Top=35; cl.Left=258; cl.Width=86; cl.Height=20; cl.Font.Size=10; cl.Style=2; cl.TabIndex=2
            lr = d.Controls.Add("Forms.Label.1"); lr.Name="lblRun"; lr.Caption=""; lr.Top=38; lr.Left=352; lr.Width=64; lr.Font.Size=10; lr.Font.Bold=True
            li = d.Controls.Add("Forms.Label.1"); li.Name="lblInfo"
            li.Caption="Preencha apenas os analitos desta corrida. Gravação vai para DB_Resultados."
            li.Top=250; li.Left=12; li.Width=400; li.Font.Size=8; li.Font.Italic=True; li.ForeColor=8421504
            bs = d.Controls.Add("Forms.CommandButton.1"); bs.Name="btnSalvar"; bs.Caption="SALVAR CORRIDA"
            bs.Top=272; bs.Left=12; bs.Width=180; bs.Height=32; bs.Font.Size=11; bs.Font.Bold=True; bs.Accelerator="S"
            bc = d.Controls.Add("Forms.CommandButton.1"); bc.Name="btnCancelar"; bc.Caption="Cancelar"
            bc.Top=272; bc.Left=202; bc.Width=120; bc.Height=32; bc.Font.Size=11; bc.Cancel=True
            f.CodeModule.AddFromString(FRM_CORRIDA.replace("{NIVEL_ITEMS}", nivel_items)); return 1
        safe_add(vbp, mk_corrida, "frmCorrida"); time.sleep(0.4); print("  + frmCorrida", flush=True)

        # ---------- frmExcluir ----------
        def mk_excluir():
            f = vbp.VBComponents.Add(3); f.Name = "frmExcluir"
            f.Properties("Caption").Value = "Excluir resultados (exclusão lógica)"
            f.Properties("Width").Value = 360; f.Properties("Height").Value = 400
            d = f.Designer
            t = d.Controls.Add("Forms.Label.1"); t.Caption="Excluir resultados"
            t.Top=8; t.Left=12; t.Width=320; t.Font.Size=13; t.Font.Bold=True
            s = d.Controls.Add("Forms.Label.1")
            s.Caption="Os registros não são apagados — apenas marcados como Excluído."
            s.Top=28; s.Left=12; s.Width=330; s.Font.Size=8; s.Font.Italic=True; s.ForeColor=8421504
            l1 = d.Controls.Add("Forms.Label.1"); l1.Caption="Nível"; l1.Top=52; l1.Left=12; l1.Width=34; l1.Font.Size=10
            cn = d.Controls.Add("Forms.ComboBox.1"); cn.Name="cboNivel"; cn.Top=49; cn.Left=48; cn.Width=52; cn.Height=20; cn.Font.Size=10; cn.Style=2; cn.TabIndex=0
            l2 = d.Controls.Add("Forms.Label.1"); l2.Caption="Corrida (RUN)"; l2.Top=52; l2.Left=118; l2.Width=76; l2.Font.Size=10
            cr = d.Controls.Add("Forms.ComboBox.1"); cr.Name="cboRun"; cr.Top=49; cr.Left=196; cr.Width=80; cr.Height=20; cr.Font.Size=10; cr.Style=2; cr.TabIndex=1
            ls = d.Controls.Add("Forms.ListBox.1"); ls.Name="lstAnalitos"
            ls.Top=80; ls.Left=12; ls.Width=232; ls.Height=250; ls.Font.Size=9
            ls.ListStyle = 1     # fmListStyleOption -> checkbox por item
            ls.MultiSelect = 1   # fmMultiSelectMulti
            bt = d.Controls.Add("Forms.CommandButton.1"); bt.Name="btnTodos"; bt.Caption="Selecionar todos"
            bt.Top=80; bt.Left=252; bt.Width=92; bt.Height=26; bt.Font.Size=9
            bl = d.Controls.Add("Forms.CommandButton.1"); bl.Name="btnLimpar"; bl.Caption="Limpar seleção"
            bl.Top=112; bl.Left=252; bl.Width=92; bl.Height=26; bl.Font.Size=9
            li = d.Controls.Add("Forms.Label.1"); li.Name="lblInfo"; li.Caption=""
            li.Top=336; li.Left=12; li.Width=330; li.Font.Size=8; li.ForeColor=8421504
            bc2 = d.Controls.Add("Forms.CommandButton.1"); bc2.Name="btnConfirmar"; bc2.Caption="EXCLUIR SELECIONADOS"
            bc2.Top=354; bc2.Left=12; bc2.Width=190; bc2.Height=30; bc2.Font.Size=10; bc2.Font.Bold=True
            bx = d.Controls.Add("Forms.CommandButton.1"); bx.Name="btnCancelar"; bx.Caption="Cancelar"
            bx.Top=354; bx.Left=212; bx.Width=132; bx.Height=30; bx.Font.Size=10; bx.Cancel=True
            f.CodeModule.AddFromString(FRM_EXCLUIR.replace("{NIVEL_ITEMS}", nivel_items)); return 1
        safe_add(vbp, mk_excluir, "frmExcluir"); time.sleep(0.4); print("  + frmExcluir", flush=True)

        # ---------- frmMassa ----------
        def mk_massa():
            f = vbp.VBComponents.Add(3); f.Name = "frmMassa"
            f.Properties("Caption").Value = "Inserção em massa"
            f.Properties("Width").Value = 720; f.Properties("Height").Value = 470
            d = f.Designer
            t = d.Controls.Add("Forms.Label.1"); t.Caption="Inserção em massa — colar da planilha"
            t.Top=8; t.Left=12; t.Width=500; t.Font.Size=13; t.Font.Bold=True
            lc = d.Controls.Add("Forms.Label.1"); lc.Name="lblColunas"; lc.Caption=""
            lc.Top=30; lc.Left=12; lc.Width=690; lc.Height=28; lc.Font.Size=8; lc.WordWrap=True; lc.ForeColor=4210752
            tg = d.Controls.Add("Forms.TextBox.1"); tg.Name="txtGrade"
            tg.Top=62; tg.Left=12; tg.Width=690; tg.Height=200
            tg.MultiLine=True; tg.EnterKeyBehavior=True; tg.ScrollBars=3; tg.Font.Name="Consolas"; tg.Font.Size=9
            tg.TabIndex=0
            bv = d.Controls.Add("Forms.CommandButton.1"); bv.Name="btnValidar"; bv.Caption="Validar"
            bv.Top=268; bv.Left=12; bv.Width=110; bv.Height=28; bv.Font.Size=10; bv.Accelerator="V"
            bl = d.Controls.Add("Forms.CommandButton.1"); bl.Name="btnLimpar"; bl.Caption="Limpar grade"
            bl.Top=268; bl.Left=130; bl.Width=110; bl.Height=28; bl.Font.Size=10
            ls = d.Controls.Add("Forms.Label.1"); ls.Name="lblStatus"; ls.Caption=""
            ls.Top=274; ls.Left=250; ls.Width=450; ls.Font.Size=9
            le = d.Controls.Add("Forms.ListBox.1"); le.Name="lstErros"
            le.Top=302; le.Left=12; le.Width=690; le.Height=96; le.Font.Name="Consolas"; le.Font.Size=8
            bf = d.Controls.Add("Forms.Label.1"); bf.Name="barFundo"; bf.Caption=""
            bf.Top=404; bf.Left=12; bf.Width=690; bf.Height=10; bf.BackColor=14737632; bf.SpecialEffect=2
            bp = d.Controls.Add("Forms.Label.1"); bp.Name="barProg"; bp.Caption=""
            bp.Top=404; bp.Left=12; bp.Width=1; bp.Height=10; bp.BackColor=6740804
            bs = d.Controls.Add("Forms.CommandButton.1"); bs.Name="btnSalvar"; bs.Caption="GRAVAR EM DB_RESULTADOS"
            bs.Top=420; bs.Left=12; bs.Width=250; bs.Height=30; bs.Font.Size=11; bs.Font.Bold=True; bs.Accelerator="G"
            bx = d.Controls.Add("Forms.CommandButton.1"); bx.Name="btnCancelar"; bx.Caption="Fechar"
            bx.Top=420; bx.Left=272; bx.Width=120; bx.Height=30; bx.Font.Size=11; bx.Cancel=True
            f.CodeModule.AddFromString(FRM_MASSA.replace("{NLV}", str(NLV))); return 1
        safe_add(vbp, mk_massa, "frmMassa"); time.sleep(0.4); print("  + frmMassa", flush=True)

        rc(lambda: xl.CalculateFullRebuild()); time.sleep(0.4)
        rc(lambda: wb.Save())
        print("  SALVO", flush=True)
    finally:
        try: wb.Close(False)
        except: pass
        try: xl.Quit()
        except: pass

if __name__ == "__main__":
    only = sys.argv[1] if len(sys.argv)>1 else None
    for fn,nlv in FILES:
        if only and only not in fn: continue
        remove_forms(fn)
        build(fn, nlv)
    print("FIM")
