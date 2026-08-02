# -*- coding: utf-8 -*-
"""FASE 2 / F2-1 + F2-2: DB_Resultados como banco operacional, aba Resultados como view
em 3 blocos, camada de dados com upsert em lote e refresh da view."""
import os, sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client as w

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
PWD = "qcini2025"
FILES = [("QC_Hematologia.xlsm", 3), ("QC_Bioquimica.xlsm", 2), ("QC_Imunologia.xlsm", 2)]
VIEW_ROWS = 3000

DARK="0F2A2E"; BAND="16383D"; MINT="5FE3C2"; MINT2="CFEEE7"; INK="1B3B40"; LIGHT="EAF6F3"

def rc(fn, t=14):
    last=None
    for _ in range(t):
        try: return fn()
        except Exception as e: last=e; time.sleep(0.4)
    raise last

def rgb(hexstr):  # "RRGGBB" -> BGR long do VBA
    r=int(hexstr[0:2],16); g=int(hexstr[2:4],16); b=int(hexstr[4:6],16)
    return r + g*256 + b*65536

# ---------------- camada de dados (substitui mDados) ----------------
MOD_DADOS = '''Option Explicit
' ===== CAMADA DE DADOS (Fase 2) =====
' DB_Resultados e a UNICA fonte operacional. Formato vertical normalizado.
' Schema: A=RUN | B=Data | C=Nivel | D=Lote | E=Analito | F=Resultado | G=Status
Public Const BANCO As String = "DB_Resultados"
Public Const VIEW As String = "Resultados"
Public Const BANCO_R0 As Long = 4
Public Const COL_RUN As Long = 1
Public Const COL_DATA As Long = 2
Public Const COL_NIVEL As Long = 3
Public Const COL_LOTE As Long = 4
Public Const COL_ANALITO As Long = 5
Public Const COL_RESULT As Long = 6
Public Const COL_STATUS As Long = 7
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

' Le o banco inteiro para memoria de uma vez (performance: sem loop celula a celula).
Public Function CarregarDB() As Variant
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    lastRow = UltimaLinhaBanco()
    If lastRow < BANCO_R0 Then
        CarregarDB = Empty
    Else
        CarregarDB = ws.Range(ws.Cells(BANCO_R0, COL_RUN), ws.Cells(lastRow, COL_STATUS)).Value
    End If
End Function

Public Function ChaveReg(ByVal run As Long, ByVal nivel As Long, ByVal analito As String) As String
    ' Chave de unicidade do registro. O Nivel ENTRA na chave: um mesmo RUN abrange
    ' todos os niveis da corrida, entao RUN+Analito sozinho colidiria entre niveis.
    ChaveReg = CStr(run) & "|" & CStr(nivel) & "|" & UCase$(Trim$(analito))
End Function

' RUN = chave logica da corrida. Unico por (Data + lote de 6 digitos).
Public Function NovoRUN(ByVal dt As Date, ByVal loteCore As String) As Long
    Dim dados As Variant, i As Long, mx As Long
    dados = CarregarDB()
    mx = 0
    If IsEmpty(dados) Then NovoRUN = 1: Exit Function
    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 And IsNumeric(dados(i, COL_RUN)) Then
            If CLng(dados(i, COL_RUN)) > mx Then mx = CLng(dados(i, COL_RUN))
            If IsDate(dados(i, COL_DATA)) Then
                If CDate(dados(i, COL_DATA)) = dt And Mid$(CStr(dados(i, COL_LOTE)), 4, 6) = loteCore Then
                    NovoRUN = CLng(dados(i, COL_RUN)): Exit Function
                End If
            End If
        End If
    Next i
    NovoRUN = mx + 1
End Function

' ===== UPSERT EM LOTE =====
' regs: array (1..n, 1..7) já no schema do banco. Atualiza o que existir
' (mesma chave RUN|Nivel|Analito) e acrescenta o resto — nunca duplica.
' Devolve "novos|atualizados".
Public Function UpsertResultados(ByRef regs As Variant) As String
    Dim ws As Worksheet, dados As Variant, idx As Object
    Dim i As Long, lastRow As Long, novos As Long, atual As Long, k As String, lin As Long
    Dim addBuf() As Variant, nAdd As Long
    If IsEmpty(regs) Then UpsertResultados = "0|0": Exit Function
    Set ws = ThisWorkbook.Sheets(BANCO)
    Set idx = CreateObject("Scripting.Dictionary")
    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
                k = ChaveReg(CLng(dados(i, COL_RUN)), CLng(dados(i, COL_NIVEL)), CStr(dados(i, COL_ANALITO)))
                If Not idx.Exists(k) Then idx.Add k, BANCO_R0 + i - 1
            End If
        Next i
    End If
    lastRow = UltimaLinhaBanco()
    ReDim addBuf(1 To UBound(regs, 1), 1 To COL_STATUS)
    nAdd = 0: novos = 0: atual = 0
    Application.ScreenUpdating = False
    For i = 1 To UBound(regs, 1)
        k = ChaveReg(CLng(regs(i, COL_RUN)), CLng(regs(i, COL_NIVEL)), CStr(regs(i, COL_ANALITO)))
        If idx.Exists(k) Then
            lin = idx(k)
            ws.Cells(lin, COL_RESULT).Value = regs(i, COL_RESULT)
            ws.Cells(lin, COL_STATUS).Value = ST_ATIVO
            ws.Cells(lin, COL_DATA).Value = regs(i, COL_DATA)
            ws.Cells(lin, COL_LOTE).Value = regs(i, COL_LOTE)
            atual = atual + 1
        Else
            nAdd = nAdd + 1
            Dim c As Long
            For c = 1 To COL_STATUS
                addBuf(nAdd, c) = regs(i, c)
            Next c
            idx.Add k, lastRow + nAdd
            novos = novos + 1
        End If
    Next i
    If nAdd > 0 Then
        ws.Range(ws.Cells(lastRow + 1, COL_LOTE), ws.Cells(lastRow + nAdd, COL_LOTE)).NumberFormat = "@"
        Dim outp() As Variant
        ReDim outp(1 To nAdd, 1 To COL_STATUS)
        For i = 1 To nAdd
            For c = 1 To COL_STATUS
                outp(i, c) = addBuf(i, c)
            Next c
        Next i
        ws.Range(ws.Cells(lastRow + 1, COL_RUN), ws.Cells(lastRow + nAdd, COL_STATUS)).Value = outp
    End If
    Application.ScreenUpdating = True
    UpsertResultados = CStr(novos) & "|" & CStr(atual)
End Function

' Exclusao LOGICA por RUN + Nivel + lista de analitos (Dictionary de nomes em UCase).
Public Function ExcluirLogico(ByVal run As Long, ByVal nivel As Long, ByRef alvos As Object) As Long
    Dim ws As Worksheet, dados As Variant, i As Long, n As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    dados = CarregarDB()
    If IsEmpty(dados) Then ExcluirLogico = 0: Exit Function
    Application.ScreenUpdating = False
    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
            If CLng(dados(i, COL_RUN)) = run And CLng(dados(i, COL_NIVEL)) = nivel Then
                If alvos.Exists(UCase$(Trim$(CStr(dados(i, COL_ANALITO))))) Then
                    ws.Cells(BANCO_R0 + i - 1, COL_STATUS).Value = ST_EXCLUIDO
                    n = n + 1
                End If
            End If
        End If
    Next i
    Application.ScreenUpdating = True
    ExcluirLogico = n
End Function

' Lista de RUNs distintos ATIVOS do lote em uso (para os combos dos formularios).
Public Function RunsDoLote(ByVal loteCore As String) As Collection
    Dim dados As Variant, i As Long, seen As Object, c As Collection
    Set c = New Collection: Set seen = CreateObject("Scripting.Dictionary")
    dados = CarregarDB()
    If IsEmpty(dados) Then Set RunsDoLote = c: Exit Function
    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 Then
            If Mid$(CStr(dados(i, COL_LOTE)), 4, 6) = loteCore Then
                If Not seen.Exists(CStr(dados(i, COL_RUN))) Then
                    seen.Add CStr(dados(i, COL_RUN)), 1
                    c.Add CLng(dados(i, COL_RUN))
                End If
            End If
        End If
    Next i
    Set RunsDoLote = c
End Function

' Nomes dos analitos cadastrados (aba Analitos).
Public Function ListaAnalitos() As Collection
    Dim ws As Worksheet, i As Long, nm As String, c As Collection
    Set c = New Collection
    Set ws = ThisWorkbook.Sheets("Analitos")
    For i = 4 To 43
        nm = Trim$(CStr(ws.Cells(i, 1).Value))
        If nm <> "" Then c.Add nm
    Next i
    Set ListaAnalitos = c
End Function

' Codigos completos de lote disponiveis no registro (Configuracao).
Public Function ListaLotes() As Collection
    Dim rng As Range, cel As Range, c As Collection, v As String
    Set c = New Collection
    On Error Resume Next
    Set rng = ThisWorkbook.Names("regLoteCol").RefersToRange
    If rng Is Nothing Then Set ListaLotes = c: Exit Function
    For Each cel In rng
        v = Trim$(CStr(cel.Value))
        If v <> "" Then c.Add v
    Next cel
    Set ListaLotes = c
End Function

Public Sub AtualizarBanco()
    Application.Calculate
End Sub

' Placeholder — trilha de auditoria e da Fase 5.
Public Sub RegistrarLog(ByVal acao As String, ByVal detalhe As String)
End Sub
'''

MOD_OPERACAO = '''Option Explicit
' ===== CAMADA DE OPERACAO (Fase 2) =====
' A aba Resultados e apenas VISUALIZACAO: reflete DB_Resultados em 3 blocos por nivel.
' Blocos: N1 = A:G | N2 = K:Q | N3 = U:AA   (colunas I:J, R:T e AB+ ficam livres p/ botoes)
Public Const VIEW_R0 As Long = 4
Public Const VIEW_ROWS As Long = %d
Public Const VIEW_NLV As Long = %d

Public Function ColunaBloco(ByVal t As Long) As Long
    ColunaBloco = 1 + t * 10          ' t=0 -> A(1), t=1 -> K(11), t=2 -> U(21)
End Function

' Reconstroi a view inteira em LOTE (uma escrita por bloco).
Public Sub AtualizarViewResultados()
    Dim ws As Worksheet, dados As Variant, lote As String
    Dim t As Long, i As Long, n As Long, c0 As Long
    Dim buf() As Variant, cont() As Long
    Set ws = ThisWorkbook.Sheets(VIEW)
    lote = LoteAtivoCore()
    Application.ScreenUpdating = False
    On Error GoTo fim

    For t = 0 To VIEW_NLV - 1
        c0 = ColunaBloco(t)
        ws.Range(ws.Cells(VIEW_R0, c0), ws.Cells(VIEW_R0 + VIEW_ROWS - 1, c0 + 5)).ClearContents
    Next t

    dados = CarregarDB()
    If IsEmpty(dados) Then GoTo fim

    ReDim buf(0 To VIEW_NLV - 1, 1 To VIEW_ROWS, 1 To 6)
    ReDim cont(0 To VIEW_NLV - 1)

    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
            If Mid$(CStr(dados(i, COL_LOTE)), 4, 6) = lote Then
                t = CLng(dados(i, COL_NIVEL)) - 1
                If t >= 0 And t <= VIEW_NLV - 1 Then
                    If cont(t) < VIEW_ROWS Then
                        cont(t) = cont(t) + 1
                        n = cont(t)
                        buf(t, n, 1) = dados(i, COL_RUN)
                        buf(t, n, 2) = dados(i, COL_DATA)
                        buf(t, n, 3) = dados(i, COL_LOTE)
                        buf(t, n, 4) = dados(i, COL_ANALITO)
                        buf(t, n, 5) = dados(i, COL_RESULT)
                        buf(t, n, 6) = dados(i, COL_STATUS)
                    End If
                End If
            End If
        End If
    Next i

    For t = 0 To VIEW_NLV - 1
        If cont(t) > 0 Then
            Dim outp() As Variant, r As Long, c As Long
            ReDim outp(1 To cont(t), 1 To 6)
            For r = 1 To cont(t)
                For c = 1 To 6
                    outp(r, c) = buf(t, r, c)
                Next c
            Next r
            c0 = ColunaBloco(t)
            ws.Range(ws.Cells(VIEW_R0, c0), ws.Cells(VIEW_R0 + cont(t) - 1, c0 + 5)).Value = outp
        End If
    Next t
fim:
    Application.ScreenUpdating = True
End Sub

' Cadeia unica chamada apos qualquer gravacao/exclusao.
Public Sub AtualizarOperacao()
    Application.ScreenUpdating = False
    AtualizarBanco
    AtualizarViewResultados
    AtualizarEstatistica
    AtualizarPainel
    Application.ScreenUpdating = True
End Sub

Public Sub AbrirFormMassa()
    frmMassa.Show
End Sub

Public Sub AbrirFormExcluir()
    frmExcluir.Show
End Sub
'''

def apply(fname, NLV):
    print("="*60); print("==", fname, "| niveis:", NLV, flush=True)
    xl = rc(lambda: w.DispatchEx("Excel.Application"))
    xl.Visible=False; xl.DisplayAlerts=False; xl.EnableEvents=False
    try: xl.AutomationSecurity=1
    except: pass
    wb = rc(lambda: xl.Workbooks.Open(os.path.join(BASE, fname)))
    try:
        for sh in wb.Worksheets:
            try: sh.Unprotect(Password=PWD)
            except: pass

        # 1) renomear Resultados -> DB_Resultados (Excel reaponta nomes/formulas sozinho)
        nomes = [s.Name for s in wb.Worksheets]
        if "DB_Resultados" not in nomes:
            wb.Sheets("Resultados").Name = "DB_Resultados"
            print("  Resultados -> DB_Resultados (nomes/formulas reapontados)", flush=True)
        db = wb.Sheets("DB_Resultados")

        # limpar coluna H (NC) do banco — removida na Fase 2
        db.Range("H3:H15003").ClearContents()
        db.Range("H3").Value = ""

        # 2) criar/limpar a nova aba Resultados (view)
        if "Resultados" in [s.Name for s in wb.Worksheets]:
            vw = wb.Sheets("Resultados"); vw.Cells.Clear()
            for shp in list(vw.Shapes): shp.Delete()
        else:
            vw = wb.Worksheets.Add(After=db); vw.Name = "Resultados"
        vw.Activate() if False else None

        vw.Cells.Interior.ColorIndex = -4142
        # titulo
        vw.Range("A1:AA1").Merge()
        vw.Range("A1").Value = "RESULTADOS · VISUALIZAÇÃO OPERACIONAL (reflete DB_Resultados)"
        vw.Range("A1").Font.Size = 16; vw.Range("A1").Font.Bold = True
        vw.Range("A1").Font.Color = rgb(MINT); vw.Range("A1").Interior.Color = rgb(DARK)
        vw.Rows(1).RowHeight = 34
        vw.Range("A2:AA2").Merge()
        vw.Range("A2").Formula = ('="Lote em uso: "&loteAtivo&"    ·    Esta aba é somente leitura — '
                                  'lance e exclua resultados pelos botões. O banco oficial é DB_Resultados."')
        vw.Range("A2").Font.Size = 10; vw.Range("A2").Font.Color = rgb(MINT2)
        vw.Range("A2").Interior.Color = rgb(BAND)
        vw.Rows(2).RowHeight = 20

        heads = ["RUN","Data","Lote","Analito","Resultado","Status"]
        widths = [7,12,14,16,12,11]
        for t in range(NLV):
            c0 = 1 + t*10
            # rotulo do bloco
            vw.Cells(3, c0).Value = f"NÍVEL {t+1}"
            for j,h in enumerate(heads):
                cel = vw.Cells(3, c0+j)
                if j>0: cel.Value = h
                cel.Font.Bold = True; cel.Font.Size = 10
                cel.Font.Color = rgb("FFFFFF"); cel.Interior.Color = rgb(INK)
                cel.HorizontalAlignment = -4108
                vw.Columns(c0+j).ColumnWidth = widths[j]
            vw.Cells(3, c0).Value = heads[0]
            # faixa com o nome do nivel acima do bloco
            vw.Range(vw.Cells(2, c0), vw.Cells(2, c0)).Value = ""
            vw.Range(vw.Cells(VIEW_ROWS+VIEW_ROWS, 1), vw.Cells(VIEW_ROWS+VIEW_ROWS, 1)).Value = ""
            # formatos
            vw.Range(vw.Cells(4, c0), vw.Cells(3+VIEW_ROWS, c0)).NumberFormat = "0"
            vw.Range(vw.Cells(4, c0+1), vw.Cells(3+VIEW_ROWS, c0+1)).NumberFormat = "dd/mm/yyyy"
            vw.Range(vw.Cells(4, c0+2), vw.Cells(3+VIEW_ROWS, c0+2)).NumberFormat = "@"
            vw.Range(vw.Cells(4, c0), vw.Cells(3+VIEW_ROWS, c0+5)).HorizontalAlignment = -4108
            vw.Range(vw.Cells(4, c0+3), vw.Cells(3+VIEW_ROWS, c0+3)).HorizontalAlignment = -4131
        vw.Rows(3).RowHeight = 24
        vw.Range("A4").Select() if False else None
        try: vw.Activate(); xl.ActiveWindow.FreezePanes = False; vw.Range("A4").Select(); xl.ActiveWindow.FreezePanes = True
        except: pass
        vw.Cells.Locked = True

        print(f"  view Resultados criada: {NLV} blocos, {VIEW_ROWS} linhas cada", flush=True)

        # 3) módulos de dados/operação
        vbp = wb.VBProject
        for c in list(vbp.VBComponents):
            if c.Name in ("mDados","mOperacao"):
                vbp.VBComponents.Remove(c)
        time.sleep(0.3)
        def ins(nome, code):
            m = vbp.VBComponents.Add(1); m.Name = nome
            m.CodeModule.AddFromString(code); return 1
        rc(lambda: ins("mDados", MOD_DADOS)); time.sleep(0.25)
        rc(lambda: ins("mOperacao", MOD_OPERACAO % (VIEW_ROWS, NLV))); time.sleep(0.25)
        print("  + mDados / mOperacao", flush=True)

        # 4) modulo da aba DB_Resultados: sem entrada manual (operacao e por UserForm)
        cm = vbp.VBComponents(db.CodeName).CodeModule
        if cm.CountOfLines > 0: cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString("Option Explicit\n' Banco operacional — gravacao exclusivamente pela camada de dados (mDados).\n")
        cmv = vbp.VBComponents(vw.CodeName).CodeModule
        if cmv.CountOfLines > 0: cmv.DeleteLines(1, cmv.CountOfLines)
        cmv.AddFromString("Option Explicit\n' Camada de visualizacao — somente leitura.\n")

        # 5) botoes nas colunas reservadas
        for shp in list(vw.Shapes): shp.Delete()
        botoes = [("Nova Corrida","AbrirFormCorrida",9),
                  ("Inserção em Massa","AbrirFormMassa",9),
                  ("Excluir Resultados","AbrirFormExcluir",9),
                  ("Atualizar View","AtualizarViewResultados",9)]
        top = vw.Rows(4).Top
        for i,(cap,act,col) in enumerate(botoes):
            a = vw.Cells(4+i*3, col)
            b = vw.Buttons().Add(a.Left, a.Top, 118, 30)
            b.Caption = cap; b.OnAction = act; b.Name = "btnOp%d" % i
        print("  + 4 botoes nas colunas reservadas (I:J)", flush=True)

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
        apply(fn, nlv)
    print("FIM")
