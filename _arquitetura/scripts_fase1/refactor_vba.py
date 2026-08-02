# -*- coding: utf-8 -*-
"""FASE 1 / Etapas 6-7: Modo Desenvolvedor + refatoracao do VBA em camadas.
Divide o monolito mSystem.bas em modulos por responsabilidade e atualiza o
codigo que toca o schema novo de Resultados."""
import os, re, sys, io, time, hashlib
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client as w

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
REC  = os.path.join(BASE, "_arquitetura", "vba_recuperado")
PWD  = "qcini2025"
DEV_PASS = "QCDEV@2026"
DEV_HASH = hashlib.sha256(DEV_PASS.encode("utf-8")).hexdigest()
FILES = ["QC_Hematologia.xlsm", "QC_Bioquimica.xlsm", "QC_Imunologia.xlsm"]

# ---------- camadas: quem vai para onde ----------
MAP = {
 "mSeguranca": ["InitK","ToU","ToS","AddM","ShR","ShL","RotR","SHA256Hex","FindUserRow",
                "DoLogin","CadastrarUsuario","AssinarCom","CarimbarRubrica","Assinar",
                "LockApp","UnlockApp","UnprotectAll","ReprotectAll","Logout"],
 "mLotes":     ["BlocoDoLote","SalvarViewNoBloco","CarregarBlocoNaView","TrocarLote","FlushLoteAtual"],
 "mUI":        ["SystemLook","AbrirFormCorrida","AtualizarEixos","HookCharts"],
}
HEADER = {
 "mSeguranca": ('Option Explicit\n'
   "' CAMADA: Seguranca / Auditoria de acesso — hash SHA-256, login, papeis,\n"
   "' assinatura eletronica, protecao de planilhas e Modo Desenvolvedor.\n"
   'Public Const SHEET_LOGIN As String = "Login"\n'
   'Public Const SHEET_START As String = "Início"\n'
   f'Private Const DEV_HASH As String = "{DEV_HASH}"\n'
   'Private m_K(0 To 63) As Long\n'
   'Private m_ready As Boolean\n'),
 "mLotes": ('Option Explicit\n'
   "' CAMADA: Lotes — troca do lote ativo e persistencia das views por lote\n"
   "' (Analitos/LotesStore, Liberacao/LiberStore, Registros/RegistrosStore).\n"
   'Private Const CAP_ANALITOS As Long = 40\n'
   'Private Const NLB_LIBER As Long = 200\n'),
 "mUI": ('Option Explicit\n'
   "' CAMADA: Interface — aparencia de sistema, formularios, eixos e graficos.\n"
   'Private gHooks As Collection\n'),
}

def split_procs(src):
    lines = src.splitlines()
    procs, cur, name = {}, None, None
    for ln in lines:
        m = re.match(r'^\s*(?:Public |Private )?(?:Sub|Function)\s+(\w+)', ln)
        if m and cur is None:
            name = m.group(1); cur = [ln]; continue
        if cur is not None:
            cur.append(ln)
            if re.match(r'^\s*End (?:Sub|Function)\s*$', ln):
                procs[name] = "\n".join(cur); cur = None; name = None
    return procs

# ---------- codigo NOVO ----------
MOD_DADOS = '''Option Explicit
' CAMADA: Dados — acesso ao banco unico (aba Resultados / Single Source of Truth).
' Schema: A=RUN | B=Data | C=Nivel | D=Lote | E=Analito | F=Resultado | G=Status | H=NC
Public Const BANCO As String = "Resultados"
Public Const BANCO_R0 As Long = 4
Public Const COL_RUN As Long = 1
Public Const COL_DATA As Long = 2
Public Const COL_NIVEL As Long = 3
Public Const COL_LOTE As Long = 4
Public Const COL_ANALITO As Long = 5
Public Const COL_RESULT As Long = 6
Public Const COL_STATUS As Long = 7
Public Const COL_NC As Long = 8
Public Const ST_ATIVO As String = "Ativo"
Public Const ST_EXCLUIDO As String = "Excluído"

Public Function UltimaLinhaBanco() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(BANCO)
    UltimaLinhaBanco = ws.Cells(ws.Rows.Count, COL_RUN).End(xlUp).Row
    If UltimaLinhaBanco < BANCO_R0 Then UltimaLinhaBanco = BANCO_R0 - 1
End Function

Public Function LoteAtivoCore() As String
    On Error Resume Next
    LoteAtivoCore = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
End Function

' RUN = chave logica da corrida analitica. Unico por (Data + Lote de 6 digitos);
' todos os niveis/analitos da mesma corrida compartilham o mesmo RUN.
Public Function NovoRUN(ByVal dt As Date, ByVal loteCore As String) As Long
    Dim ws As Worksheet, lastRow As Long, dados As Variant, i As Long, mx As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    lastRow = UltimaLinhaBanco()
    If lastRow < BANCO_R0 Then NovoRUN = 1: Exit Function
    dados = ws.Range(ws.Cells(BANCO_R0, COL_RUN), ws.Cells(lastRow, COL_LOTE)).Value
    mx = 0
    For i = 1 To UBound(dados, 1)
        If IsNumeric(dados(i, COL_RUN)) And Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 Then
            If CLng(dados(i, COL_RUN)) > mx Then mx = CLng(dados(i, COL_RUN))
            If IsDate(dados(i, COL_DATA)) Then
                If CDate(dados(i, COL_DATA)) = dt Then
                    If Mid$(CStr(dados(i, COL_LOTE)), 4, 6) = loteCore Then
                        NovoRUN = CLng(dados(i, COL_RUN)): Exit Function
                    End If
                End If
            End If
        End If
    Next i
    NovoRUN = mx + 1
End Function

' Grava uma corrida no banco. valores() = array 1..n de Variant alinhado a nomes().
Public Function GravarCorrida(ByVal dt As Date, ByVal nivel As Long, _
        ByRef nomes() As String, ByRef valores() As Variant, ByVal n As Long) As Long
    Dim ws As Worksheet, r As Long, i As Long, run As Long, lote As String, gravadas As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    lote = LoteAtivoCore()
    If lote = "" Then GravarCorrida = -1: Exit Function
    run = NovoRUN(dt, lote)
    r = UltimaLinhaBanco() + 1
    gravadas = 0
    For i = 1 To n
        If Len(Trim$(CStr(valores(i)))) > 0 Then
            ws.Cells(r, COL_RUN).Value = run
            ws.Cells(r, COL_DATA).Value = dt
            ws.Cells(r, COL_NIVEL).Value = nivel
            ws.Cells(r, COL_LOTE).NumberFormat = "@"
            ws.Cells(r, COL_LOTE).Value = "QC-" & lote & Format(nivel, "00")
            ws.Cells(r, COL_ANALITO).Value = nomes(i)
            ws.Cells(r, COL_RESULT).Value = CDbl(valores(i))
            ws.Cells(r, COL_STATUS).Value = ST_ATIVO
            r = r + 1: gravadas = gravadas + 1
        End If
    Next i
    GravarCorrida = IIf(gravadas > 0, run, 0)
End Function

' Exclusao LOGICA — nunca apaga fisicamente (Fase 2 usara isto no formulario).
Public Sub MarcarStatusRUN(ByVal run As Long, ByVal novoStatus As String)
    Dim ws As Worksheet, i As Long, lastRow As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    lastRow = UltimaLinhaBanco()
    For i = BANCO_R0 To lastRow
        If ws.Cells(i, COL_RUN).Value = run Then ws.Cells(i, COL_STATUS).Value = novoStatus
    Next i
End Sub

Public Sub AtualizarBanco()
    Application.Calculate
End Sub

' Placeholder de trilha de auditoria — implementacao fica para a Fase 2.
Public Sub RegistrarLog(ByVal acao As String, ByVal detalhe As String)
    ' Fase 2: gravar em aba de auditoria (usuario, data/hora, acao, detalhe).
End Sub
'''

MOD_DEV = '''
' ===================== MODO DESENVOLVEDOR (Etapa 6) =====================
' Acionado por um Shape 100%% transparente na aba Painel (sem atalho de teclado).
Public Sub ModoDesenvolvedor()
    Dim s As String, ws As Worksheet
    frmDev.Confirmado = False
    frmDev.txtSenha.Value = ""
    frmDev.Show
    If Not frmDev.Confirmado Then
        Unload frmDev
        Exit Sub
    End If
    s = frmDev.txtSenha.Value
    Unload frmDev
    If LCase$(SHA256Hex(s)) <> LCase$(DEV_HASH) Then
        MsgBox "Senha incorreta.", vbExclamation, "Modo Desenvolvedor"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:="%s"
        ws.Visible = xlSheetVisible
    Next ws
    On Error GoTo 0
    SystemLook False
    Application.ScreenUpdating = True
    MsgBox "Modo Desenvolvedor ATIVO." & vbCrLf & vbCrLf & _
           "Todas as abas visiveis e desprotegidas." & vbCrLf & _
           "Use Sair (logout) para voltar ao modo normal.", vbInformation, "Modo Desenvolvedor"
End Sub
''' % PWD

MOD_UI_EXTRA = '''
' ===================== ORQUESTRACAO (Etapa 7) =====================
' Rotinas de responsabilidade unica; a cadeia completa e AtualizarTudo.
Public Sub AtualizarResultados()
    Application.Calculate
End Sub

Public Sub AtualizarEstatistica()
    On Error Resume Next
    ThisWorkbook.Sheets("Estatística").Calculate
End Sub

Public Sub AtualizarGraficos()
    AtualizarEixos
    HookCharts
End Sub

Public Sub AtualizarPainel()
    On Error Resume Next
    ThisWorkbook.Sheets("Painel").Calculate
    AtualizarGraficos
End Sub

Public Sub AtualizarTudo()
    Application.ScreenUpdating = False
    AtualizarBanco
    AtualizarResultados
    AtualizarEstatistica
    AtualizarPainel
    Application.ScreenUpdating = True
End Sub
'''

FRM_DEV_CODE = '''Option Explicit
Public Confirmado As Boolean
Private Sub btnOK_Click()
    Confirmado = True
    Me.Hide
End Sub
Private Sub btnCancel_Click()
    Confirmado = False
    Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Confirmado = False
End Sub
'''

RESULT_MOD = '''Option Explicit
' Camada de interface da aba Resultados (banco).
' Duplo-clique numa linha vazia pre-preenche RUN/Data/Nivel/Lote/Status.
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim r As Long, nivel As Variant, lote As String, dt As Date
    If Target.Row < BANCO_R0 Or Target.Column > COL_NC Then Exit Sub
    r = Target.Row
    Cancel = True
    Application.EnableEvents = False
    On Error GoTo fim
    If Trim$(CStr(Me.Cells(r, COL_DATA).Value)) = "" Then Me.Cells(r, COL_DATA).Value = Date
    If Trim$(CStr(Me.Cells(r, COL_NIVEL).Value)) = "" Then Me.Cells(r, COL_NIVEL).Value = UltimoNivel(r)
    dt = CDate(Me.Cells(r, COL_DATA).Value)
    nivel = Me.Cells(r, COL_NIVEL).Value
    lote = LoteAtivoCore()
    If lote <> "" And IsNumeric(nivel) Then
        If Trim$(CStr(Me.Cells(r, COL_LOTE).Value)) = "" Then
            Me.Cells(r, COL_LOTE).NumberFormat = "@"
            Me.Cells(r, COL_LOTE).Value = "QC-" & lote & Format(CLng(nivel), "00")
        End If
        If Trim$(CStr(Me.Cells(r, COL_RUN).Value)) = "" Then
            Me.Cells(r, COL_RUN).Value = NovoRUN(dt, lote)
        End If
    End If
    If Trim$(CStr(Me.Cells(r, COL_STATUS).Value)) = "" Then Me.Cells(r, COL_STATUS).Value = ST_ATIVO
    Me.Cells(r, COL_ANALITO).Select
fim:
    Application.EnableEvents = True
End Sub

Private Function UltimoNivel(ByVal r As Long) As Variant
    Dim i As Long
    For i = r - 1 To BANCO_R0 Step -1
        If Trim$(CStr(Me.Cells(i, COL_NIVEL).Value)) <> "" Then
            UltimoNivel = Me.Cells(i, COL_NIVEL).Value
            Exit Function
        End If
    Next i
    UltimoNivel = 1
End Function
'''

def rc(fn, t=12):
    last = None
    for _ in range(t):
        try: return fn()
        except Exception as e: last = e; time.sleep(0.4)
    raise last

def build_modules():
    src = io.open(os.path.join(REC, "mSystem.bas"), encoding="utf-8", errors="replace").read()
    procs = split_procs(src)
    print("  procedures encontradas no monolito:", len(procs), flush=True)
    mods = {}
    for mod, names in MAP.items():
        body = [HEADER[mod]]
        for nm in names:
            if nm in procs: body.append(procs[nm])
            else: print("    AVISO: proc ausente:", nm, flush=True)
        if mod == "mSeguranca": body.append(MOD_DEV)
        if mod == "mUI":        body.append(MOD_UI_EXTRA)
        mods[mod] = "\n\n".join(body)
    mods["mDados"] = MOD_DADOS
    return mods

def patch_frmcorrida(code):
    """Adapta o formulario ao schema novo (RUN/Status) e a camada de dados."""
    # SeqParaData -> NovoRUN via camada mDados
    code = re.sub(r'Private Function SeqParaData.*?End Function', '''Private Function RunParaData(ByVal dt As Date) As Long
    RunParaData = NovoRUN(dt, LoteAtivoCore())
End Function''', code, flags=re.S)
    code = code.replace("SeqParaData(", "RunParaData(")
    code = code.replace('Me.lblSeqInfo.Caption = "Corrida n. "', 'Me.lblSeqInfo.Caption = "RUN "')
    # gravacao: delega para mDados.GravarCorrida
    novo_salvar = '''Private Sub btnSalvar_Click()
    Dim i As Long, lvl As Long, dt As Date, n As Long, run As Long
    Dim nomes() As String, vals() As Variant
    On Error Resume Next
    dt = CDate(Me.txtData.Value)
    If Err.Number <> 0 Then
        MsgBox "Data invalida.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    If Me.cboNivel.ListIndex = -1 Then
        MsgBox "Selecione o nivel.", vbExclamation
        Exit Sub
    End If
    lvl = CLng(Me.cboNivel.Value)
    ReDim nomes(1 To analCount)
    ReDim vals(1 To analCount)
    n = 0
    For i = 1 To analCount
        nomes(i) = analNames(i)
        vals(i) = Me.Controls("txtV" & i).Value
        If Len(Trim$(CStr(vals(i)))) > 0 Then
            If Not IsNumeric(vals(i)) Then
                MsgBox "Valor invalido em " & analNames(i) & ".", vbExclamation
                Exit Sub
            End If
            n = n + 1
        End If
    Next i
    If n = 0 Then
        MsgBox "Nenhum valor informado.", vbExclamation
        Exit Sub
    End If
    run = GravarCorrida(dt, lvl, nomes, vals, analCount)
    If run = -1 Then
        MsgBox "Nenhum lote em uso selecionado (aba Configuracao).", vbExclamation
        Exit Sub
    End If
    RegistrarLog "LANCAMENTO", "RUN " & run & " nivel " & lvl & " (" & n & " resultados)"
    AtualizarTudo
    MsgBox n & " resultado(s) salvo(s) no RUN " & run & " (Nivel " & lvl & ").", vbInformation
    Unload Me
End Sub'''
    code = re.sub(r'Private Sub btnSalvar_Click\(\).*?\nEnd Sub', novo_salvar, code, flags=re.S)
    return code

def apply(fname, mods):
    print("=" * 60); print("==", fname, flush=True)
    xl = rc(lambda: w.DispatchEx("Excel.Application"))
    xl.Visible = False; xl.DisplayAlerts = False; xl.EnableEvents = False
    try: xl.AutomationSecurity = 1
    except: pass
    wb = rc(lambda: xl.Workbooks.Open(os.path.join(BASE, fname)))
    try:
        for sh in wb.Worksheets:
            try: sh.Unprotect(Password=PWD)
            except: pass
        vbp = wb.VBProject

        # 1) remover monolito
        for c in list(vbp.VBComponents):
            if c.Name == "mSystem":
                vbp.VBComponents.Remove(c); print("  mSystem removido", flush=True)
        time.sleep(0.3)

        # 2) inserir modulos em camadas
        for mod in ("mSeguranca", "mLotes", "mDados", "mUI"):
            def ins(mod=mod):
                m = vbp.VBComponents.Add(1); m.Name = mod
                m.CodeModule.AddFromString(mods[mod]); return 1
            rc(ins); time.sleep(0.25)
            print(f"  + {mod} ({len(mods[mod].splitlines())} linhas)", flush=True)

        # 3) frmDev (senha do Modo Desenvolvedor)
        for c in list(vbp.VBComponents):
            if c.Name == "frmDev": vbp.VBComponents.Remove(c)
        time.sleep(0.2)
        def mkdev():
            f = vbp.VBComponents.Add(3); f.Name = "frmDev"
            f.Properties("Caption").Value = "Acesso restrito"
            f.Properties("Width").Value = 250; f.Properties("Height").Value = 130
            d = f.Designer
            l = d.Controls.Add("Forms.Label.1"); l.Caption = "Senha do desenvolvedor"
            l.Top = 14; l.Left = 14; l.Width = 200; l.Font.Size = 10
            t = d.Controls.Add("Forms.TextBox.1"); t.Name = "txtSenha"
            t.Top = 38; t.Left = 14; t.Width = 210; t.Height = 22; t.Font.Size = 10; t.PasswordChar = "*"
            b1 = d.Controls.Add("Forms.CommandButton.1"); b1.Name = "btnOK"; b1.Caption = "Entrar"
            b1.Top = 72; b1.Left = 40; b1.Width = 80; b1.Height = 26
            b2 = d.Controls.Add("Forms.CommandButton.1"); b2.Name = "btnCancel"; b2.Caption = "Cancelar"
            b2.Top = 72; b2.Left = 130; b2.Width = 80; b2.Height = 26
            f.CodeModule.AddFromString(FRM_DEV_CODE); return 1
        rc(mkdev); time.sleep(0.3); print("  + frmDev", flush=True)

        # 4) frmCorrida -> schema novo
        for c in vbp.VBComponents:
            if c.Name == "frmCorrida":
                cm = c.CodeModule
                cur = cm.Lines(1, cm.CountOfLines)
                new = patch_frmcorrida(cur)
                cm.DeleteLines(1, cm.CountOfLines)
                cm.AddFromString(new)
                print("  frmCorrida adaptado ao schema novo", flush=True)
        time.sleep(0.3)

        # 5) modulo da aba Resultados
        rs = wb.Worksheets("Resultados")
        cm = vbp.VBComponents(rs.CodeName).CodeModule
        if cm.CountOfLines > 0: cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(RESULT_MOD)
        print("  modulo da aba Resultados atualizado", flush=True)

        # 6) shape invisivel do Modo Desenvolvedor no Painel
        pa = wb.Worksheets("Painel")
        for shp in list(pa.Shapes):
            if shp.Name == "btnDev":
                shp.Delete()
        anchor = pa.Range("R1")
        def mkshape():
            s = pa.Shapes.AddShape(1, anchor.Left + anchor.Width - 16, anchor.Top + 2, 14, 14)
            s.Name = "btnDev"
            s.Fill.Visible = False
            s.Line.Visible = False
            s.OnAction = "ModoDesenvolvedor"
            return 1
        rc(mkshape); print("  + shape invisivel 'btnDev' no Painel (R1, canto sup. dir.)", flush=True)

        rc(lambda: xl.CalculateFullRebuild()); time.sleep(0.4)
        rc(lambda: wb.Save())
        print("  SALVO", flush=True)
    finally:
        try: wb.Close(False)
        except: pass
        try: xl.Quit()
        except: pass

if __name__ == "__main__":
    mods = build_modules()
    only = sys.argv[1] if len(sys.argv) > 1 else None
    for f in FILES:
        if only and only not in f: continue
        apply(f, mods)
    print("FIM | senha do Modo Desenvolvedor:", DEV_PASS)
