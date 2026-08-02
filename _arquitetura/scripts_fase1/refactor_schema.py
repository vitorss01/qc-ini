# -*- coding: utf-8 -*-
"""FASE 1 / Etapas 3-5: schema RUN+Status, exclusao logica, consumidores.
Refatora IN PLACE os .xlsm (os geradores originais foram perdidos)."""
import os, sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client as w

BASE = r"C:\Users\vitor\OneDrive - MSFT\Desktop\QC_INI"
PWD = "qcini2025"
RR0, RMAX = 4, 15003
KC0, KCN = 3, 182          # Calc: linhas de dados
COL0, NFIELDS = 6, 22      # Calc: 1a coluna de nivel (F) e campos por nivel
E0 = 7                     # Estatistica: 1a linha
LB0, NLB = 4, 200          # Liberacao
CAP = 40

FILES = [("QC_Hematologia.xlsm", 3), ("QC_Bioquimica.xlsm", 2), ("QC_Imunologia.xlsm", 2)]

def rc(fn, t=12):
    last = None
    for _ in range(t):
        try: return fn()
        except Exception as e: last = e; time.sleep(0.4)
    raise last

def colletter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

def refactor(fname, NLV):
    print("=" * 60); print("==", fname, "| niveis:", NLV, flush=True)
    xl = rc(lambda: w.DispatchEx("Excel.Application"))
    xl.Visible = False; xl.DisplayAlerts = False; xl.EnableEvents = False
    try: xl.AutomationSecurity = 1
    except: pass
    wb = rc(lambda: xl.Workbooks.Open(os.path.join(BASE, fname)))
    try:
        for sh in wb.Worksheets:
            try: sh.Unprotect(Password=PWD)
            except: pass

        rs = wb.Sheets("Resultados")
        last = rs.Cells(rs.Rows.Count, 1).End(-4162).Row   # xlUp, coluna A = Data (schema antigo)
        print("  linhas de dados:", last - RR0 + 1, flush=True)

        # ---------- 1) ler schema ANTIGO: A=Data B=Nivel C=Seq D=Lote E=Analito F=Valor G=NC ----------
        old = rs.Range(f"A{RR0}:G{last}").Value

        # ---------- 2) RUN global unico por (Data, Lote) ----------
        runmap, nxt, rows = {}, 1, []
        for r in old:
            data, nivel, seq, lote, analito, valor, nc = r[0], r[1], r[2], r[3], r[4], r[5], r[6]
            core = str(lote)[3:9] if lote else ""   # 6 digitos: o nivel NAO entra na chave da corrida
            key = (str(data), core)
            if key not in runmap:
                runmap[key] = nxt; nxt += 1
            rows.append((runmap[key], data, nivel, lote, analito, valor, "Ativo", nc))
        print("  RUNs distintos:", len(runmap), "| max RUN:", nxt - 1, flush=True)

        # ---------- 3) gravar schema NOVO: A=RUN B=Data C=Nivel D=Lote E=Analito F=Resultado G=Status H=NC ----------
        rs.Range(f"A{RR0}:H{RMAX}").ClearContents()
        rs.Range(f"A{RR0}:H{last}").Value = rows
        hdr = ["RUN", "Data", "Nível", "Lote", "Analito", "Resultado", "Status", "NC (x)"]
        rs.Range("A3:H3").Value = [hdr]
        for col, wid in zip("ABCDEFGH", [8, 13, 8, 14, 16, 12, 10, 8]):
            rs.Columns(col).ColumnWidth = wid
        rs.Range(f"A{RR0}:A{RMAX}").NumberFormat = "0"        # RUN e inteiro (nao herdar formato de data)
        rs.Range(f"C{RR0}:C{RMAX}").NumberFormat = "0"
        rs.Range(f"G{RR0}:G{RMAX}").NumberFormat = "@"
        rs.Range(f"H{RR0}:H{RMAX}").NumberFormat = "@"
        rs.Range(f"B{RR0}:B{RMAX}").NumberFormat = "dd/mm/yyyy"
        rs.Range(f"D{RR0}:D{RMAX}").NumberFormat = "@"
        for col in ("A", "B", "C", "D", "F", "G", "H"):
            rs.Range(f"{col}{RR0}:{col}{RMAX}").HorizontalAlignment = -4108  # center

        # ---------- 4) validacoes ----------
        def dv(rng, f1):
            try: rs.Range(rng).Validation.Delete()
            except: pass
            rs.Range(rng).Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=f1)
        dv(f"C{RR0}:C{RMAX}", ",".join(str(i) for i in range(1, NLV + 1)))
        dv(f"E{RR0}:E{RMAX}", f"=Analitos!$A$4:$A${3+CAP}")
        dv(f"G{RR0}:G{RMAX}", "Ativo,Excluído")
        dv(f"H{RR0}:H{RMAX}", "x")
        lastcod = rs.Cells(rs.Rows.Count, 56).End(-4162).Row   # BD = codigos completos
        if lastcod >= 2: dv(f"D{RR0}:D{RMAX}", f"=$BD$2:$BD${lastcod}")

        # ---------- 5) colunas auxiliares (BA/BB/BC) ----------
        rs.Range("BA3").Value = "LoteCore"; rs.Range("BB3").Value = "1ªOc"; rs.Range("BC3").Value = "RunUnico"
        rs.Range(f"BA{RR0}:BA{RMAX}").Formula = '=IF($D4="","",MID($D4,4,6))'
        rs.Range(f"BB{RR0}:BB{RMAX}").Formula = (
            '=IF(OR($E4="",$G4<>"Ativo"),"",'
            f'IF(COUNTIFS($E${RR0}:$E4,$E4,$A${RR0}:$A4,$A4,$G${RR0}:$G4,"Ativo")=1,1,0))')
        rs.Range(f"BC{RR0}:BC{RMAX}").Formula = (
            '=IF(OR($A4="",$G4<>"Ativo"),"",'
            f'IF(COUNTIFS($A${RR0}:$A4,$A4,$G${RR0}:$G4,"Ativo")=1,1,0))')
        for c in ("BA", "BB", "BC"):
            rs.Columns(c).ColumnWidth = 7; rs.Columns(c).Hidden = True

        # ---------- 6) nomes definidos ----------
        def setname(nm, ref):
            try: wb.Names(nm).Delete()
            except: pass
            wb.Names.Add(Name=nm, RefersTo=ref)
        setname("rRUN",      f"=Resultados!$A${RR0}:$A${RMAX}")
        setname("rData",     f"=Resultados!$B${RR0}:$B${RMAX}")
        setname("rNivel",    f"=Resultados!$C${RR0}:$C${RMAX}")
        setname("rAnalito",  f"=Resultados!$E${RR0}:$E${RMAX}")
        setname("rValor",    f"=Resultados!$F${RR0}:$F${RMAX}")
        setname("rStatus",   f"=Resultados!$G${RR0}:$G${RMAX}")
        setname("rLote",     f"=Resultados!$BA${RR0}:$BA${RMAX}")
        setname("rFirst",    f"=Resultados!$BB${RR0}:$BB${RMAX}")
        setname("rRunUnico", f"=Resultados!$BC${RR0}:$BC${RMAX}")
        try: wb.Names("rSeq").Delete()
        except: pass
        print("  nomes definidos atualizados", flush=True)

        # ---------- 7) Calc: descoberta de RUN + data + valores (com Status=Ativo) ----------
        ca = wb.Sheets("Calc")
        ca.Range(f"B{KC0}:B{KCN}").Formula = (
            '=IF(selAnalito="","",IFERROR(AGGREGATE(15,6,'
            'rRUN/((rAnalito=selAnalito)*(rFirst=1)*((""&rLote)=(""&loteAtivo))),$A3),""))')
        ca.Range(f"C{KC0}:C{KCN}").Formula = (
            '=IF($B3="","",MAXIFS(rData,rAnalito,selAnalito,rRUN,$B3,rLote,loteAtivo,rStatus,"Ativo"))')
        for t in range(NLV):
            lvl = t + 1
            vc = colletter(COL0 + t * NFIELDS)
            ca.Range(f"{vc}{KC0}:{vc}{KCN}").Formula = (
                f'=IF($B3="","",IF(COUNTIFS(rAnalito,selAnalito,rNivel,{lvl},rRUN,$B3,rLote,loteAtivo,rStatus,"Ativo")=0,"",'
                f'SUMIFS(rValor,rAnalito,selAnalito,rNivel,{lvl},rRUN,$B3,rLote,loteAtivo,rStatus,"Ativo")))')
        print("  Calc atualizado (RUN + Status)", flush=True)

        # ---------- 8) Estatistica: n / media / DP com Status=Ativo ----------
        es = wb.Sheets("Estatística")
        EN = E0 + CAP * NLV - 1
        yf = ('rData,">="&DATE($B$3,1,1),rData,"<="&DATE($D$3,12,31),'
              'rLote,IF($B$4="","*",$B$4),rStatus,"Ativo"')
        es.Range(f"C{E0}:C{EN}").Formula = f'=IF($A{E0}="",0,COUNTIFS(rAnalito,$A{E0},rNivel,$B{E0},{yf}))'
        es.Range(f"D{E0}:D{EN}").Formula = f'=IF($C{E0}=0,"",AVERAGEIFS(rValor,rAnalito,$A{E0},rNivel,$B{E0},{yf}))'
        es.Range(f"E{E0}:E{EN}").Formula = (
            f'=IF($C{E0}<2,"",SQRT((SUMPRODUCT((rAnalito=$A{E0})*(rNivel=$B{E0})*'
            '(YEAR(rData)>=$B$3)*(YEAR(rData)<=$D$3)*(rStatus="Ativo")*'
            f'IF($B$4="",1,((""&rLote)=(""&$B$4)))*(rValor^2))-$C{E0}*$D{E0}^2)/($C{E0}-1)))')
        print("  Estatística atualizada (Status=Ativo)", flush=True)

        # ---------- 9) Liberacao: lista os RUNs distintos do lote ativo ----------
        lb = wb.Sheets("Liberação")
        lb.Range("A3").Value = "RUN"
        lb.Range(f"A{LB0}:A{LB0+NLB-1}").Formula = (
            f'=IFERROR(AGGREGATE(15,6,rRUN/((rRunUnico=1)*((""&rLote)=(""&loteAtivo))),ROW()-{LB0-1}),"")')
        lb.Range(f"B{LB0}:B{LB0+NLB-1}").Formula = (
            '=IF($A4="","",IFERROR(IF(MINIFS(rData,rRUN,$A4,rLote,loteAtivo,rStatus,"Ativo")=0,"",'
            'MINIFS(rData,rRUN,$A4,rLote,loteAtivo,rStatus,"Ativo")),""))')
        print("  Liberação atualizada (RUN distinto por lote)", flush=True)

        rc(lambda: xl.CalculateFullRebuild()); time.sleep(0.5)
        rc(lambda: wb.Save())
        print("  SALVO", flush=True)
    finally:
        try: wb.Close(False)
        except: pass
        try: xl.Quit()
        except: pass

if __name__ == "__main__":
    only = sys.argv[1] if len(sys.argv) > 1 else None
    for fn, nlv in FILES:
        if only and only not in fn: continue
        refactor(fn, nlv)
    print("FIM")
