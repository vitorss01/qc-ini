# -*- coding: utf-8 -*-
"""Verificacao Etapas 6-8: compilacao do VBA, camadas, Modo Desenvolvedor, regressao."""
import os, sys, io, time, hashlib, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client as w

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
DEV_PASS = "QCDEV@2026"
DEV_HASH = hashlib.sha256(DEV_PASS.encode("utf-8")).hexdigest()
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
        vbp=wb.VBProject
        mods=sorted(c.Name for c in vbp.VBComponents if c.Type==1)
        forms=sorted(c.Name for c in vbp.VBComponents if c.Type==3)
        print("  modulos:",mods,flush=True)
        print("  formularios:",forms,flush=True)

        # --- compilacao real: SHA256Hex tem que bater com o hashlib ---
        h=rc(lambda: xl.Run("SHA256Hex",DEV_PASS))
        ok = (str(h).lower()==DEV_HASH)
        print("  SHA256Hex compila e confere:",ok,flush=True)

        # --- login ---
        rc(lambda: setattr(wb.Names("loginUser").RefersToRange,"Value","QCINI"))
        rc(lambda: setattr(wb.Names("loginPass").RefersToRange,"Value","HEMATO123"))
        rc(lambda: xl.Run("DoLogin")); time.sleep(0.3)
        print("  DoLogin -> papel:",rc(lambda: wb.Names("currentPapel").RefersToRange.Value),flush=True)

        pa=wb.Sheets("Painel"); rs=wb.Sheets("Resultados"); cf=wb.Sheets("Configuração")
        rc(lambda: setattr(pa.Range("B3"),"Value",1))
        rc(lambda: xl.Run("AtualizarTudo")); time.sleep(0.4)
        print("  AtualizarTudo OK -> n(N1) =",rc(lambda: pa.Range("B7").Value),flush=True)

        # --- camada de dados: NovoRUN reaproveita RUN existente e cria novo ---
        core=rc(lambda: wb.Names("loteAtivo").RefersToRange.Value)
        d0=rc(lambda: rs.Range("B4").Value)
        r_exist=rc(lambda: xl.Run("NovoRUN",d0,str(core)))
        r_novo =rc(lambda: xl.Run("NovoRUN",datetime.datetime(2031,5,5),str(core)))
        print(f"  NovoRUN(data existente)={r_exist} (esperado 1) | NovoRUN(data nova)={r_novo} (esperado 26)",flush=True)

        # --- shape invisivel do Modo Desenvolvedor ---
        found=None
        for shp in pa.Shapes:
            if shp.Name=="btnDev":
                found=(shp.OnAction, shp.Fill.Visible, shp.Line.Visible, round(shp.Width), round(shp.Height))
        print("  shape btnDev (acao, fill, line, w, h):",found,flush=True)

        # --- troca de lote continua funcionando ---
        rc(lambda: setattr(cf.Range("C27"),"NumberFormat","@"))
        rc(lambda: setattr(cf.Range("C27"),"Value","770000"))
        rc(lambda: setattr(cf.Range("C20"),"Value","770000"))
        rc(lambda: xl.Run("TrocarLote")); rc(lambda: xl.CalculateFullRebuild()); time.sleep(0.3)
        n_novo=rc(lambda: pa.Range("B7").Value)
        rc(lambda: setattr(cf.Range("C20"),"Value",str(core)))
        rc(lambda: xl.Run("TrocarLote")); rc(lambda: xl.CalculateFullRebuild()); time.sleep(0.3)
        n_volta=rc(lambda: pa.Range("B7").Value)
        print(f"  TrocarLote: lote novo n={n_novo} (esp 0) | volta n={n_volta} (esp 25)",flush=True)

        # --- formulario abre (Initialize sem erro) ---
        code='''Public Function TestForm() As String
    On Error GoTo eh
    Load frmCorrida
    TestForm = "controles=" & frmCorrida.Controls.Count
    Unload frmCorrida
    Exit Function
eh: TestForm = "ERRO " & Err.Number & ": " & Err.Description
End Function'''
        m=vbp.VBComponents.Add(1); m.Name="tstTmp"; m.CodeModule.AddFromString(code)
        print("  frmCorrida:",rc(lambda: xl.Run("TestForm")),flush=True)
        vbp.VBComponents.Remove(m)

        rc(lambda: xl.CalculateFullRebuild())
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
