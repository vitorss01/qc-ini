# -*- coding: utf-8 -*-
"""restaurar_status_westgard.py - ADR-036: M7/M8 de volta, sem tocar no layout

O QUE FAZ, E SO ISSO

  1. escreve a FORMULA em Painel!M7 e Painel!M8
  2. cria Estatistica!AC e AD -- N de resultados e N de rodadas de EQA por
     analito
  3. prova, comparando o XML da aba antes e depois, que nenhum outro estilo
     do Painel mudou

O QUE NAO FAZ

Nao move celula, nao troca cor, nao redimensiona, nao reordena bloco, nao
reaplica layout. O Painel foi organizado manualmente pelo gestor e essa
organizacao e a base oficial.

A LOGICA ORIGINAL, LOCALIZADA E NAO INVENTADA

No commit 07f3ff1 o status vivia em Painel!S7 e S8:

    =IF($B7="","",IF($R7>0,"REPROVA — "&$R7&" violação(ões)","Sem violação"))

$B7 = n de resultados do nivel; $R7 = total de violacoes das cinco regras.
No layout novo o bloco Westgard foi para F6:M8 e o Total desceu de R para L.
A formula e a mesma; muda so a referencia da coluna.

DEFEITO ENCONTRADO NA AUDITORIA DA LOGICA ORIGINAL

O guarda `$B7=""` nunca dispara. B7 e AGGREGATE(2;6;...), que e uma CONTAGEM:
devolve ZERO quando nao ha resultado no periodo, nunca texto vazio. Com o
periodo vazio a celula exibia "Sem violação" -- que se le como "esta tudo
certo" onde na verdade nao ha dado nenhum. O guarda passa a testar o zero.

DE ONDE VEM O DADO

  Calc!K..O   x  Calc!D   ->  violacoes do NIVEL 1 dentro do filtro de periodo
  Calc!AG..AK x  Calc!D   ->  violacoes do NIVEL 2
  Painel!L7 e L8          ->  soma das cinco regras
  Painel!B7 e B8          ->  n de resultados do nivel

Nenhuma dessas fontes le Sigma, classificacao, DPM, rendimento, run size ou
margem de ETp. O status da corrida pertence ao CIQ, e so a ele.

Uso: python restaurar_status_westgard.py <arquivo.xlsm>
"""
import io
import os
import re
import sys
import time
import shutil
import zipfile
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

CRLF = chr(13) + chr(10)
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

EST_R0, EST_RN = 14, 93

# As duas colunas de rastreabilidade pedidas: quantos resultados e quantas
# rodadas de EQA entraram no |Bias| DAQUELE analito -- e nao o total global do
# filtro, que ja aparece no eco em K5.
NOVAS = [
    (29, 'N EQA Resultados',
     '=IF($A{0}="","",mCEQ.BiasEQ($A{0},eqAnoEP,"N",eqProvedor,eqRodada))'),
    (30, 'N EQA Rodadas',
     '=IF($A{0}="","",mCEQ.BiasEQ($A{0},eqAnoEP,"NRODADAS",eqProvedor,eqRodada))'),
]

FORMULA_M = ('=IF(OR(NOT(ISNUMBER($B{0})),$B{0}=0),"sem dados no período",'
             'IF($L{0}>0,"REPROVA — "&$L{0}&" violação(ões)","Sem violação"))')


def importar_bas(vbp, caminho):
    import tempfile
    texto = io.open(caminho, encoding='utf-8', errors='replace').read()
    tmp = os.path.join(tempfile.gettempdir(), 'ansi_' + os.path.basename(caminho))
    with io.open(tmp, 'w', encoding='cp1252', errors='replace',
                 newline=CRLF) as fh:
        fh.write(texto)
    vbp.VBComponents.Import(tmp)


def novo_excel():
    for t in range(1, 9):
        try:
            xl = w.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.EnableEvents = False
            xl.AutomationSecurity = 1
            return xl
        except Exception:
            if t in (1, 3, 5):
                subprocess.call(['powershell', '-NoProfile', '-Command',
                                 'Start-Process excel.exe -WindowStyle Minimized'],
                                stderr=subprocess.DEVNULL)
                time.sleep(10)
            time.sleep(2.5 * t)
    raise RuntimeError('Excel COM nao subiu')


TRANS = ('rejeitada', 'rejected', 'membro n', 'member not found', 'busy')


def tenta(fn, vezes=10):
    ult = None
    for i in range(vezes):
        try:
            return fn()
        except Exception as e:
            ult = e
            if not any(t in str(e).lower() for t in TRANS):
                raise
            time.sleep(1.0 + 0.8 * i)
    raise ult


def estilos_do_painel(caminho):
    """Mapa {celula: indice de estilo} lido do XML da aba Painel.

    Comparar isto antes e depois e a unica forma barata de provar que nenhuma
    cor, borda, fonte, alinhamento ou formato mudou: sao milhares de celulas, e
    ler propriedade por propriedade via COM levaria dezenas de milhares de
    chamadas.
    """
    z = zipfile.ZipFile(caminho)
    wbx = z.read('xl/workbook.xml').decode('utf-8', 'replace')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'replace')
    rid = {m.group(1): m.group(2)
           for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)}
    alvo = None
    for m in re.finditer(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx):
        if m.group(1) == 'Painel':
            alvo = 'xl/' + rid[m.group(2)].lstrip('/').replace('xl/', '')
    if alvo is None:
        z.close()
        return {}
    sx = z.read(alvo).decode('utf-8', 'replace')
    z.close()
    d = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*)/?>', sx):
        ref, attrs = m.group(1), m.group(2)
        s = re.search(r's="(\d+)"', attrs)
        d[ref] = s.group(1) if s else '0'
    # largura de coluna e altura de linha tambem entram na comparacao
    for m in re.finditer(r'<col min="(\d+)" max="(\d+)"[^>]*width="([\d.]+)"', sx):
        for c in range(int(m.group(1)), int(m.group(2)) + 1):
            d['__col%d' % c] = m.group(3)
    for m in re.finditer(r'<row r="(\d+)"[^>]*ht="([\d.]+)"', sx):
        d['__row%s' % m.group(1)] = m.group(2)
    return d


def main(caminho):
    caminho = os.path.abspath(caminho)
    antes_arq = os.path.join(os.environ.get('TEMP', '.'), 'painel_antes.xlsm')
    shutil.copy(caminho, antes_arq)
    antes = estilos_do_painel(antes_arq)
    print('estilos do Painel antes: %d celulas mapeadas' % len(antes))

    xl = novo_excel()
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura: %s' % caminho)
    try:
        xl.Calculation = -4135
        pa = wb.Worksheets('Painel')
        es = wb.Worksheets('Estatística')

        vbp = wb.VBProject
        for nome in ('mCEQ', 'mPlanoQC'):
            for c in list(vbp.VBComponents):
                if c.Name == nome:
                    vbp.VBComponents.Remove(c)
            importar_bas(vbp, os.path.join(RAIZ, 'src_producao', nome + '.bas'))

        # ---- 0. confere o pressuposto ANTES de escrever -------------------
        print()
        print('=== o bloco Westgard esta onde eu penso? ===')
        for ref, esperado in (('F6', 'Nível'), ('G6', '1-3S'), ('L6', 'Total'),
                              ('M6', 'Status'), ('F7', 'N1'), ('F8', 'N2')):
            v = str(tenta(lambda r=ref: pa.Range(r).Value) or '')
            ok = v.strip() == esperado
            print('   %-4s = %-10s %s' % (ref, v[:10], 'ok' if ok else
                                          'ESPERADO %r' % esperado))
            if not ok:
                raise SystemExit('o bloco nao esta na posicao suposta -- '
                                 'nada escrito')
        for ref in ('M7', 'M8'):
            v = tenta(lambda r=ref: pa.Range(r).Formula)
            print('   %-4s antes: %r' % (ref, v))

        # ---- 1. so M7 e M8, so a formula ---------------------------------
        for lin in (7, 8):
            tenta(lambda x=lin: pa.Cells(x, 13).__setattr__(
                'Formula', FORMULA_M.format(x)))
        print()
        print('Painel!M7 e M8: formula restaurada (nada mais foi tocado)')

        # ---- 2. rastreabilidade por analito na Estatistica ---------------
        for c, tit, f in NOVAS:
            cel = es.Cells(13, c)
            cel.Value = tit
            cel.Font.Bold = True
            es.Columns(c).ColumnWidth = 16
            tenta(lambda cc=c, ff=f:
                  es.Range(es.Cells(EST_R0, cc), es.Cells(EST_RN, cc))
                  .__setattr__('Formula', ff.format(EST_R0)))
            tenta(lambda cc=c:
                  es.Range(es.Cells(EST_R0, cc), es.Cells(EST_RN, cc))
                  .__setattr__('NumberFormatLocal', '0'))
        print('Estatistica AC e AD: N de resultados e N de rodadas por analito')

        xl.Calculation = -4105
        tenta(lambda: xl.CalculateFullRebuild())

        # ---- 3. conferencia rapida ---------------------------------------
        print()
        print('=== o que M7/M8 exibem agora ===')
        for lin in (7, 8):
            print('   M%d = %r   (n=%s, total de violacoes=%s)'
                  % (lin, str(tenta(lambda x=lin: pa.Cells(x, 13).Text)),
                     tenta(lambda x=lin: pa.Cells(x, 2).Value),
                     tenta(lambda x=lin: pa.Cells(x, 12).Value)))

        erros = []
        for r in range(1, 40):
            for c in range(1, 30):
                t = str(tenta(lambda x=r, y=c: pa.Cells(x, y).Text))
                if t.startswith('#') and t.strip('#') != '':
                    erros.append('Painel!%d,%d=%s' % (r, c, t))
        for r in range(13, 94):
            for c in (29, 30):
                t = str(tenta(lambda x=r, y=c: es.Cells(x, y).Text))
                if t.startswith('#') and t.strip('#') != '':
                    erros.append('Estatistica!%d,%d=%s' % (r, c, t))
        print('celulas em erro: %d %s' % (len(erros), erros[:5]))
        if erros:
            raise SystemExit('erro de formula -- nada salvo')

        wb.Save()
        salvou = True
    finally:
        try:
            wb.Close(salvou)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    # ---- 4. o layout do Painel sobreviveu? -------------------------------
    depois = estilos_do_painel(caminho)
    print()
    print('=== DIFF DE ESTILO DO PAINEL (antes x depois) ===')
    print('celulas mapeadas: %d -> %d' % (len(antes), len(depois)))
    mudou, sumiu, nasceu = [], [], []
    for k, v in antes.items():
        if k not in depois:
            sumiu.append(k)
        elif depois[k] != v:
            mudou.append('%s: estilo %s -> %s' % (k, v, depois[k]))
    for k in depois:
        if k not in antes:
            nasceu.append(k)
    print('   estilos alterados : %s' % (mudou if mudou else 'nenhum'))
    print('   celulas sumidas   : %s' % (sumiu if sumiu else 'nenhuma'))
    print('   celulas novas     : %s' % (nasceu if nasceu else 'nenhuma'))

    # M7 e M8 podem aparecer como novas (antes estavam vazias) -- e o unico
    # nascimento legitimo desta intervencao.
    inesperadas = [k for k in nasceu if k not in ('M7', 'M8')]
    if mudou or sumiu or inesperadas:
        print()
        print('LAYOUT DO PAINEL FOI ALTERADO ALEM DE M7/M8 -- revise')
        sys.exit(1)
    print()
    print('LAYOUT DO PAINEL PRESERVADO: so M7 e M8 mudaram')
    print('SALVO: %s' % caminho)


if __name__ == '__main__':
    main(sys.argv[1])
