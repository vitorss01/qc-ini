# -*- coding: utf-8 -*-
"""FASE 2 / F2-6: regressao funcional completa."""
import os, sys, io, time, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client as w

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
CODE = {-2146826281:"#DIV/0!",-2146826246:"#N/A",-2146826259:"#NAME?",-2146826288:"#NULL!",
        -2146826252:"#NUM!",-2146826265:"#REF!",-2146826273:"#VALUE!"}
def iserr(v): return isinstance(v,int) and v in CODE
def rc(fn,t=14):
    last=None
    for _ in range(t):
        try: return fn()
        except Exception as e: last=e; time.sleep(0.4)
    raise last

TEST = '''Public Function T_Fase2() As String
    Dim r As String, dados As Variant, n0 As Long, n1 As Long, i As Long
    Dim regs(1 To 2, 1 To 7) As Variant, res As String, lote As String, run As Long
    On Error GoTo eh
    lote = LoteAtivoCore()
    dados = CarregarDB(): n0 = UBound(dados, 1)
    r = "linhas_iniciais=" & n0

    ' --- 1) INSERIR: RUN novo (data futura) ---
    run = NovoRUN(DateSerial(2032, 3, 3), lote)
    regs(1,1)=run: regs(1,2)=DateSerial(2032,3,3): regs(1,3)=1
    regs(1,4)=CodigoLote(lote,1): regs(1,5)="__TESTE_A": regs(1,6)=10.5: regs(1,7)=ST_ATIVO
    regs(2,1)=run: regs(2,2)=DateSerial(2032,3,3): regs(2,3)=2
    regs(2,4)=CodigoLote(lote,2): regs(2,5)="__TESTE_A": regs(2,6)=20.5: regs(2,7)=ST_ATIVO
    res = UpsertResultados(regs)
    r = r & " | insercao(novos|atual)=" & res

    ' --- 2) UPSERT: mesma chave, valor diferente -> deve ATUALIZAR, nao duplicar ---
    regs(1,6) = 99.9: regs(2,6) = 88.8
    res = UpsertResultados(regs)
    dados = CarregarDB(): n1 = UBound(dados, 1)
    r = r & " | reenvio(novos|atual)=" & res & " linhas_apos=" & n1
    r = r & " | delta_linhas=" & (n1 - n0) & " (esperado 2)"

    ' --- 3) valor foi mesmo atualizado? ---
    Dim achou As String
    For i = 1 To UBound(dados, 1)
        If CStr(dados(i, COL_ANALITO)) = "__TESTE_A" And CLng(dados(i, COL_NIVEL)) = 1 Then
            achou = CStr(dados(i, COL_RESULT))
        End If
    Next i
    r = r & " | valor_N1=" & achou & " (esperado 99,9)"

    ' --- 4) EXCLUSAO LOGICA ---
    Dim alvos As Object
    Set alvos = CreateObject("Scripting.Dictionary")
    alvos("__TESTE_A") = 1
    r = r & " | excluidos=" & ExcluirLogico(run, 1, alvos)

    ' --- 5) RunsDoLote / ListaAnalitos / ListaLotes ---
    r = r & " | runs=" & RunsDoLote(lote).Count
    r = r & " | analitos=" & ListaAnalitos().Count
    r = r & " | lotes=" & ListaLotes().Count
    T_Fase2 = r
    Exit Function
eh:
    T_Fase2 = "ERRO " & Err.Number & ": " & Err.Description
End Function

Public Function T_Limpar() As String
    ' remove fisicamente as linhas de teste (so o teste faz isso)
    Dim ws As Worksheet, i As Long, n As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    For i = UltimaLinhaBanco() To BANCO_R0 Step -1
        If CStr(ws.Cells(i, COL_ANALITO).Value) = "__TESTE_A" Then
            ws.Rows(i).Delete: n = n + 1
        End If
    Next i
    T_Limpar = "linhas de teste removidas: " & n
End Function

Public Function T_Forms() As String
    Dim r As String
    On Error GoTo eh
    Load frmCorrida: r = "frmCorrida=" & frmCorrida.Controls.Count: Unload frmCorrida
    Load frmExcluir: r = r & " frmExcluir=" & frmExcluir.Controls.Count & _
        " analitos=" & frmExcluir.lstAnalitos.ListCount & " runs=" & frmExcluir.cboRun.ListCount
    Unload frmExcluir
    Load frmMassa:   r = r & " frmMassa=" & frmMassa.Controls.Count: Unload frmMassa
    T_Forms = r
    Exit Function
eh:
    T_Forms = "ERRO " & Err.Number & ": " & Err.Description
End Function
'''

def check(fname, NLV):
    print("="*60); print("==", fname, flush=True)
    xl = rc(lambda: w.DispatchEx("Excel.Application"))
    xl.Visible=False; xl.DisplayAlerts=False; xl.EnableEvents=False
    try: xl.AutomationSecurity=1
    except: pass
    wb = rc(lambda: xl.Workbooks.Open(os.path.join(BASE, fname)))
    try:
        vbp = wb.VBProject
        print("  modulos:", sorted(c.Name for c in vbp.VBComponents if c.Type==1), flush=True)
        print("  forms  :", sorted(c.Name for c in vbp.VBComponents if c.Type==3), flush=True)
        rc(lambda: setattr(wb.Names("loginUser").RefersToRange,"Value","QCINI"))
        rc(lambda: setattr(wb.Names("loginPass").RefersToRange,"Value","HEMATO123"))
        rc(lambda: xl.Run("DoLogin")); time.sleep(0.3)

        pa=wb.Sheets("Painel"); vw=wb.Sheets("Resultados")
        rc(lambda: setattr(pa.Range("B3"),"Value",1)); rc(lambda: xl.CalculateFullRebuild())
        print("  [Fase 1] Painel n(N1) =", rc(lambda: pa.Range("B7").Value), "(esp 25)", flush=True)

        m = vbp.VBComponents.Add(1); m.Name="tstF2"; m.CodeModule.AddFromString(TEST)
        time.sleep(0.3)
        print("  [forms]", rc(lambda: xl.Run("T_Forms")), flush=True)
        print("  [dados]", rc(lambda: xl.Run("T_Fase2")), flush=True)
        rc(lambda: xl.Run("AtualizarViewResultados"))
        n1 = vw.Cells(vw.Rows.Count,1).End(-4162).Row-3
        print("  view bloco N1 linhas =", n1, flush=True)
        print("  [limpeza]", rc(lambda: xl.Run("T_Limpar")), flush=True)
        rc(lambda: xl.Run("AtualizarOperacao")); time.sleep(0.3)
        print("  [Fase 1 pos-teste] Painel n(N1) =", rc(lambda: pa.Range("B7").Value), "(esp 25)", flush=True)
        vbp.VBComponents.Remove(m)

        rc(lambda: xl.CalculateFullRebuild())
        rep={}
        for sht in wb.Worksheets:
            d=sht.UsedRange.Value
            if d is None: continue
            if not isinstance(d,tuple): d=((d,),)
            for r in d:
                if not isinstance(r,tuple): r=(r,)
                for v in r:
                    if iserr(v):
                        c=CODE[v]
                        if sht.Name!="Calc" or c!="#N/A": rep[(sht.Name,c)]=rep.get((sht.Name,c),0)+1
        print("  ERROS:", dict(rep) if rep else "NENHUM", flush=True)
        rc(lambda: xl.Run("SystemLook",False))
    finally:
        try: wb.Close(False)
        except: pass
        try: xl.Quit()
        except: pass

if __name__=="__main__":
    fs=[("QC_Hematologia.xlsm",3),("QC_Bioquimica.xlsm",2),("QC_Imunologia.xlsm",2)]
    only=sys.argv[1] if len(sys.argv)>1 else None
    for fn,nlv in fs:
        if only and only not in fn: continue
        check(fn,nlv)
    print("FIM")
