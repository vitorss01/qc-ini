import os,sys,io,time
sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding="utf-8")
import win32com.client as w
BASE=r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
CODE={-2146826281:"#DIV/0!",-2146826246:"#N/A",-2146826259:"#NAME?",-2146826288:"#NULL!",-2146826252:"#NUM!",-2146826265:"#REF!",-2146826273:"#VALUE!"}
def iserr(v): return isinstance(v,int) and v in CODE
def rc(fn,t=14):
    last=None
    for _ in range(t):
        try: return fn()
        except Exception as e: last=e; time.sleep(0.4)
    raise last
f=sys.argv[1] if len(sys.argv)>1 else "QC_Hematologia.xlsm"
xl=rc(lambda: w.DispatchEx("Excel.Application")); xl.Visible=False; xl.DisplayAlerts=False; xl.EnableEvents=False
try: xl.AutomationSecurity=1
except: pass
wb=rc(lambda: xl.Workbooks.Open(os.path.join(BASE,f)))
try:
    print("==",f)
    print("  abas:",[s.Name for s in wb.Worksheets])
    for n in ("rRUN","rData","rStatus","rLote"):
        print(f"   {n:<9}",wb.Names(n).RefersTo)
    rc(lambda: setattr(wb.Names("loginUser").RefersToRange,"Value","QCINI"))
    rc(lambda: setattr(wb.Names("loginPass").RefersToRange,"Value","HEMATO123"))
    rc(lambda: xl.Run("DoLogin")); time.sleep(0.3)
    pa=wb.Sheets("Painel"); vw=wb.Sheets("Resultados")
    rc(lambda: setattr(pa.Range("B3"),"Value",1)); rc(lambda: xl.CalculateFullRebuild())
    print("  Painel n(N1) =",rc(lambda: pa.Range("B7").Value),"(Fase 1 intacta? esp 25)")
    t0=time.time(); rc(lambda: xl.Run("AtualizarViewResultados")); dt=time.time()-t0
    print(f"  AtualizarViewResultados OK ({dt:.1f}s)")
    for t,c0 in enumerate([1,11,21][: (3 if "Hema" in f else 2)]):
        hdr=[vw.Cells(3,c0+j).Value for j in range(6)]
        row=[vw.Cells(4,c0+j).Value for j in range(6)]
        n=vw.Cells(vw.Rows.Count,c0).End(-4162).Row-3
        print(f"   bloco N{t+1} col{c0}: linhas={n} hdr={hdr}")
        print(f"      1a linha: {row}")
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
    print("  ERROS:",dict(rep) if rep else "NENHUM")
    rc(lambda: xl.Run("SystemLook",False))
finally:
    try: wb.Close(False)
    except: pass
    try: xl.Quit()
    except: pass
