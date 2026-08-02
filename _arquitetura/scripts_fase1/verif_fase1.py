# -*- coding: utf-8 -*-
"""Verificacao Etapas 3-5: schema, RUN, Status/exclusao logica, consumidores."""
import os, sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client as w

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
CODE = {-2146826281:"#DIV/0!",-2146826246:"#N/A",-2146826259:"#NAME?",-2146826288:"#NULL!",
        -2146826252:"#NUM!",-2146826265:"#REF!",-2146826273:"#VALUE!"}
def iserr(v): return isinstance(v,int) and v in CODE

def rc(fn,t=12):
    last=None
    for _ in range(t):
        try: return fn()
        except Exception as e: last=e; time.sleep(0.4)
    raise last

def check(fname):
    print("="*60); print("==",fname,flush=True)
    xl=rc(lambda: w.DispatchEx("Excel.Application"))
    xl.Visible=False; xl.DisplayAlerts=False; xl.EnableEvents=False
    try: xl.AutomationSecurity=1
    except: pass
    wb=rc(lambda: xl.Workbooks.Open(os.path.join(BASE,fname)))
    try:
        rc(lambda: setattr(wb.Names("loginUser").RefersToRange,"Value","QCINI"))
        rc(lambda: setattr(wb.Names("loginPass").RefersToRange,"Value","HEMATO123"))
        rc(lambda: xl.Run("DoLogin")); time.sleep(0.3)
        rs=wb.Sheets("Resultados"); pa=wb.Sheets("Painel"); es=wb.Sheets("Estatística"); lb=wb.Sheets("Liberação")
        rc(lambda: setattr(pa.Range("B3"),"Value",1))   # selecionar o analito 1 (o mesmo que o teste exclui)
        rc(lambda: xl.CalculateFullRebuild())

        print("  cabecalho:", [rs.Cells(3,c).Value for c in range(1,9)], flush=True)
        print("  linha 4   :", [rs.Cells(4,c).Value for c in range(1,9)], flush=True)
        n0=rc(lambda: pa.Range("B7").Value)
        print("  Painel n (N1) =",n0,"(esperado 25)",flush=True)
        print("  Liberacao A4/B4 =",rc(lambda: lb.Range("A4").Value),"/",rc(lambda: lb.Range("B4").Value),flush=True)
        print("  Estatistica n(l7) =",rc(lambda: es.Range("C7").Value),flush=True)

        # ---- exclusao logica: marcar as 3 linhas do RUN 25 (um analito) como Excluido ----
        alvo=rc(lambda: rs.Range("E4").Value)
        marc=0
        for r in range(4,1600):
            if rs.Cells(r,1).Value==25 and rs.Cells(r,5).Value==alvo:
                rs.Cells(r,7).Value="Excluído"; marc+=1
        rc(lambda: xl.CalculateFullRebuild()); time.sleep(0.3)
        n1=rc(lambda: pa.Range("B7").Value)
        print(f"  excluidas {marc} linhas de '{alvo}' no RUN 25 -> Painel n =",n1,"(esperado 24)",flush=True)
        print("  Liberacao ultima linha preenchida =",
              rc(lambda: lb.Range("A28").Value),"(RUN 25 deve sumir se todo o RUN for excluido)",flush=True)
        e1=rc(lambda: es.Range("C7").Value)
        print("  Estatistica n apos exclusao =",e1,"(deve cair)",flush=True)
        # reverter
        for r in range(4,1600):
            if rs.Cells(r,1).Value==25 and rs.Cells(r,5).Value==alvo:
                rs.Cells(r,7).Value="Ativo"
        rc(lambda: xl.CalculateFullRebuild()); time.sleep(0.3)
        print("  revertido -> Painel n =",rc(lambda: pa.Range("B7").Value),"(esperado 25)",flush=True)

        rep={}
        for sht in wb.Worksheets:
            data=sht.UsedRange.Value
            if data is None: continue
            if not isinstance(data,tuple): data=((data,),)
            for row in data:
                if not isinstance(row,tuple): row=(row,)
                for v in row:
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
    for f in (sys.argv[1:] or ["QC_Hematologia.xlsm"]): check(f)
    print("FIM")
