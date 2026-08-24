# -*- coding: utf-8 -*-
"""Por que 2-2S/4-1S/8X nunca ficam verdes nos analitos reais.

Hipotese: nao e o realce que falhou -- e a violacao vencendo a recomendacao,
que e exatamente a prioridade projetada. Se for isso, cada celula que deixou
de ficar verde deve estar (a) recomendada pelo Sigma E (b) com contagem de
violacao > 0 E (c) pintada com a cor de violacao.

Aqui nao se conclui nada por raciocinio: le-se a bandeira da recomendacao em
Cfg!J10:N10, a contagem em G7:K7 / G8:K8, e a COR EFETIVA em DisplayFormat.
"""
import io, os, sys, time, subprocess
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

def cor(r,g,b): return r+g*256+b*65536
VERDE = cor(0x14,0x6C,0x43); VERM = cor(0xFD,0xE2,0xE2)
REGRAS = [('G6','1-3S'),('H6','2-2S'),('I6','R4S'),('J6','4-1S'),('K6','8X')]
TRANS = ('rejeitada','rejected','membro n','member not found','busy')

def tenta(fn, vezes=10):
    u=None
    for i in range(vezes):
        try: return fn()
        except Exception as e:
            u=e
            if not any(t in str(e).lower() for t in TRANS): raise
            time.sleep(1.0+0.8*i)
    raise u

def novo():
    for t in range(1,9):
        try:
            xl=w.DispatchEx('Excel.Application'); xl.Visible=False
            xl.DisplayAlerts=False; xl.EnableEvents=False; xl.AutomationSecurity=1
            return xl
        except Exception:
            if t in (1,3,5):
                subprocess.call(['powershell','-NoProfile','-Command',
                    'Start-Process excel.exe -WindowStyle Minimized'],stderr=subprocess.DEVNULL)
                time.sleep(10)
            time.sleep(2.5*t)
    raise RuntimeError('Excel COM nao subiu')

ALVOS = ['Cálcio','Bilirrubina total','Lactato','Capacidade de fixação do ferro','Ácido úrico']

def main(caminho):
    caminho=os.path.abspath(caminho)
    xl=novo(); wb=xl.Workbooks.Open(caminho, False, True)   # read-only: so mede
    try:
        pa=wb.Worksheets('Painel'); cfg=wb.Worksheets('Cfg_PlanoQC')
        aba=wb.Worksheets('Analitos')
        # B3 guarda o INDICE do analito, nao o nome
        idx={}
        for k in range(4,44):
            nm=tenta(lambda x=k: aba.Cells(x,1).Value)
            if nm: idx[str(nm).strip()]=k-3
        sel=tenta(lambda: pa.Range('B3').Value)
        tenta(lambda: xl.Application.__setattr__('Calculation', -4105))
        for an in ALVOS:
            if an not in idx:
                print('%s: ausente da aba Analitos' % an); continue
            tenta(lambda i=idx[an]: pa.Range('B3').__setattr__('Value', i))
            tenta(lambda: xl.CalculateFull())
            sg=tenta(lambda: cfg.Range('B1').Value)
            rec=[int(tenta(lambda c=c: cfg.Cells(10,c).Value) or 0) for c in range(10,15)]
            print()
            print('%-32s Sigma_Plano=%r' % (an, sg))
            print('   %-6s %-8s %-8s %-8s %-10s %s'
                  % ('regra','recom?','viol N1','viol N2','cor','veredito'))
            for i,(ref,rot) in enumerate(REGRAS):
                col=7+i
                v1=tenta(lambda c=col: pa.Cells(7,c).Value) or 0
                v2=tenta(lambda c=col: pa.Cells(8,c).Value) or 0
                c=int(tenta(lambda r=ref: pa.Range(r).DisplayFormat.Interior.Color))
                nome={VERDE:'VERDE',VERM:'VERMELHA'}.get(c,'#%06X'%c)
                viol=(v1 or 0)+(v2 or 0)>0
                if rec[i] and viol and c==VERM:
                    vd='recomendada E violada -> violacao venceu (esperado)'
                elif rec[i] and not viol and c==VERDE:
                    vd='recomendada, sem violacao -> verde (esperado)'
                elif not rec[i] and viol and c==VERM:
                    vd='nao recomendada mas violada -> vermelha (esperado)'
                elif not rec[i] and not viol:
                    vd='nem recomendada nem violada -> neutra (esperado)'
                else:
                    vd='*** INESPERADO ***'
                print('   %-6s %-8s %-8s %-8s %-10s %s'
                      % (rot,'sim' if rec[i] else 'nao', v1, v2, nome, vd))
        tenta(lambda: pa.Range('B3').__setattr__('Value', sel))
    finally:
        try: wb.Close(False)
        except Exception: pass
        try: xl.Quit()
        except Exception: pass

main(sys.argv[1])
