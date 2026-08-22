# -*- coding: utf-8 -*-
"""corrigir_nota.py - conserta uma afirmacao falsa que eu mesmo gravei

Eu escrevi no Painel que "4-1s e 8x sao reservas e permanecem em zero", com
base em Calc!N3 e Calc!O3, que sao IF(OR(FALSE;FALSE);1;0). Errado: a linha 3 e
a PRIMEIRA do bloco -- nao existe ponto anterior, entao os termos degeneram
para FALSE. Da linha 4 em diante as formulas crescem e as regras disparam: 45
vezes em 4-1s e 52 em 8x, nas 180 linhas do Calc.

Conferir a primeira linha de uma regra sequencial e conferir o unico lugar onde
ela nao pode disparar.

Tambem acrescenta ao eco do filtro de EP quantas linhas o filtro encontra. O
arquivo estava com o filtro em Controllab, provedor que nao tem uma linha
sequer na EQC_Dados -- e o resultado eram 80 "SEM EP" sem explicacao visivel.
"""
import io,os,sys,time,subprocess
sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8',write_through=True)
import win32com.client as w

NOTA = ('As cinco regras são avaliadas (1-3s, 2-2s, R-4s, 4-1s, 8x). '
        'As primeiras linhas de cada lote não têm histórico suficiente '
        'para 4-1s e 8x, e por isso não disparam ali.')
ECO = ('=ResumoFiltroEQ(eqProvedor,eqAnoEP,eqRodada)&" | "&'
       'TEXT(SUMPRODUCT((EQC_Dados!$A$4:$A$1003<>"")*'
       '(EQC_Dados!$E$4:$E$1003=eqProvedor)*'
       '(EQC_Dados!$B$4:$B$1003=eqAnoEP)*'
       'IF(eqRodada="TODAS",1,--(EQC_Dados!$C$4:$C$1003=eqRodada))),"0")&'
       '" linha(s) no banco de EP"')

def novo():
    for t in range(1,9):
        try:
            xl=w.DispatchEx('Excel.Application'); xl.Visible=False
            xl.DisplayAlerts=False; xl.EnableEvents=False; xl.AutomationSecurity=1
            return xl
        except Exception:
            if t in (1,3):
                subprocess.call(['powershell','-NoProfile','-Command','Start-Process excel.exe -WindowStyle Minimized'],stderr=subprocess.DEVNULL)
                time.sleep(10)
            time.sleep(2.5*t)
    raise RuntimeError('sem COM')

p=os.path.abspath(sys.argv[1])
xl=novo(); wb=xl.Workbooks.Open(p); salvou=False
if wb.ReadOnly:
    wb.Close(False); xl.Quit(); raise SystemExit('somente leitura')
try:
    pa=wb.Worksheets('Painel'); es=wb.Worksheets('Estatística')
    pa.Cells(9,12).Value = NOTA
    print('Painel L9 corrigido: %s' % NOTA[:60])
    es.Cells(5,11).Formula = ECO
    xl.CalculateFull()
    eco = es.Cells(5,11).Value
    print('eco do filtro: %r' % eco)
    if str(eco).startswith('#'):
        raise SystemExit('eco quebrou -- nada salvo')
    # prova: com CAP o eco tem de mudar de numero
    antes = es.Range('L4').Value
    es.Range('L4').Value='CAP'; xl.CalculateFull()
    print('com CAP      : %r' % es.Cells(5,11).Value)
    es.Range('L4').Value=antes; xl.CalculateFull()
    wb.Save(); salvou=True
    print('SALVO: %s' % p)
finally:
    try: wb.Close(salvou)
    except Exception: pass
    try: xl.Quit()
    except Exception: pass
