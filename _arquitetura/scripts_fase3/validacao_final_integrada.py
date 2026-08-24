# -*- coding: utf-8 -*-
"""validacao_final_integrada.py - a prova de ponta a ponta

Trabalha numa COPIA em TEMP. NAO altera implementacao.

O INSTRUMENTO QUE FAZ ESTA VALIDACAO VALER

Range.DisplayFormat devolve o formato EFETIVO da celula -- o que o Excel
realmente pinta depois de resolver toda a formatacao condicional. Ate agora as
provas liam a FORMULA da condicao e a prioridade dela, e concluiam por
raciocinio qual venceria. DisplayFormat mede o resultado.

E a diferenca entre "a regra de violacao tem prioridade menor, entao deve
vencer" e "a celula esta pintada de vermelho".

Blocos: A fronteiras, B pior nivel, C analitos reais, D realce efetivo,
E violacao > recomendacao, F independencia M7/M8, H ET x ETp, I plano de CQ.

Uso: python validacao_final_integrada.py <arquivo.xlsm>
"""
import io
import os
import sys
import time
import shutil
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

RESULTADOS = []          # (bloco, teste, entrada, esperado, obtido, passou)

CFG = 'Cfg_PlanoQC'
MAT_C0, MAT_N = 10, 5
REGRAS = [('G6', '1-3S'), ('H6', '2-2S'), ('I6', 'R4S'),
          ('J6', '4-1S'), ('K6', '8X')]

VERDE = 0x14 + 0x6C * 256 + 0x43 * 65536
BRANCO = 255 + 255 * 256 + 255 * 65536
VERM_FUNDO = 0xFD + 0xE2 * 256 + 0xE2 * 65536

TODAS5 = ['1-3S', '2-2S', 'R4S', '4-1S', '8X']

# Sigma -> (classe, regras, N, runsize)
ESCADA = [
    (2.99, 'Desempenho inadequado', TODAS5, '', ''),
    (3.00, 'Marginal', TODAS5, 6, 45),
    (3.99, 'Marginal', TODAS5, 6, 45),
    (4.00, 'Bom', TODAS5[:4], 4, 200),
    (4.99, 'Bom', TODAS5[:4], 4, 200),
    (5.00, 'Excelente', TODAS5[:3], 2, 450),
    (5.99, 'Excelente', TODAS5[:3], 2, 450),
    (6.00, 'Classe mundial', TODAS5[:1], 2, 1000),
    (6.01, 'Classe mundial', TODAS5[:1], 2, 1000),
]

# (N1, N2, Sigma_Plano esperado, nivel governante)
PIOR = [
    (6.50, 5.50, 5.50, 'nível 2'),
    (5.50, 4.50, 4.50, 'nível 2'),
    (4.50, 3.50, 3.50, 'nível 2'),
    (6.50, 2.99, 2.99, 'nível 2'),
    (2.99, 6.50, 2.99, 'nível 1'),
    (4.20, None, 4.20, 'nível 1'),
    (None, 4.20, 4.20, 'nível 2'),
    (None, None, 'SEM DADOS', ''),
]

ANALITOS = ['Ácido úrico', 'Cálcio', 'Bilirrubina total', 'Lactato',
            'Capacidade de fixação do ferro']


def reg(bloco, teste, entrada, esperado, obtido, passou):
    RESULTADOS.append((bloco, teste, str(entrada), str(esperado),
                       str(obtido), bool(passou)))
    print('  %s  %-46s %-26s %s'
          % ('OK  ' if passou else 'FALHA', teste[:46],
             ('esp ' + str(esperado))[:26], 'obt ' + str(obtido)[:44]))


def txt(v):
    return '' if v is None else str(v).strip()


def perto(a, b, tol=1e-9):
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return False


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


def main(caminho):
    copia = os.path.join(os.environ.get('TEMP', '.'),
                         'vfi_' + os.path.basename(caminho))
    shutil.copy(caminho, copia)
    xl = novo_excel()
    wb = xl.Workbooks.Open(copia)
    try:
        pa = wb.Worksheets('Painel')
        es = wb.Worksheets('Estatística')
        an = wb.Worksheets('Analitos')
        cfg = wb.Worksheets(CFG)
        cfg.Visible = -1
        xl.Calculation = -4105

        fV10 = tenta(lambda: pa.Range('V10').Formula)
        fV11 = tenta(lambda: pa.Range('V11').Formula)
        fG7 = tenta(lambda: pa.Range('G7').Formula)

        def poe(n1, n2):
            for ref, v in (('V10', n1), ('V11', n2)):
                if v is None:
                    tenta(lambda r=ref: pa.Range(r).ClearContents())
                else:
                    tenta(lambda r=ref, x=v: pa.Range(r).__setattr__('Value', x))
            tenta(lambda: xl.CalculateFull())

        def acesas():
            rot = [txt(tenta(lambda c=c: cfg.Cells(3, c).Value))
                   for c in range(MAT_C0, MAT_C0 + MAT_N)]
            return [r for r, c in zip(rot, range(MAT_C0, MAT_C0 + MAT_N))
                    if tenta(lambda cc=c: cfg.Cells(10, cc).Value) == 1]

        def efetivo(ref):
            """O formato que o Excel REALMENTE pinta, ja resolvida a CF."""
            d = tenta(lambda r=ref: pa.Range(r).DisplayFormat)
            return (int(tenta(lambda x=d: x.Interior.Color)),
                    int(tenta(lambda x=d: x.Font.Color)),
                    bool(tenta(lambda x=d: x.Font.Bold)),
                    bool(tenta(lambda x=d: x.Font.Italic)))

        def zeraViolacoes():
            for lin in (7, 8):
                for c in range(7, 12):
                    tenta(lambda x=lin, y=c:
                          pa.Cells(x, y).__setattr__('Value', 0))
            tenta(lambda: xl.CalculateFull())

        def devolveViolacoes():
            for lin, letra in ((7, 'K'), (8, 'AG')):
                pass
            for c, orig in ((7, fG7),):
                pass

        # ================================================================
        print('=' * 78)
        print('A. FRONTEIRAS EXATAS DO SIGMA')
        print('=' * 78)
        zeraViolacoes()
        for sg, classe, regras, n, run in ESCADA:
            poe(sg, sg)
            zeraViolacoes()
            gc = txt(tenta(lambda s=sg: xl.Run('PlanoQC', s, 'CLASSE')))
            gr = txt(tenta(lambda: pa.Range('T22').Value))
            gn = txt(tenta(lambda: pa.Range('U22').Text))
            gu = txt(tenta(lambda: pa.Range('V22').Text))
            gf = txt(tenta(lambda: pa.Range('W22').Value))
            ac = acesas()
            sp = tenta(lambda: cfg.Range('B1').Value)
            base = txt(tenta(lambda: pa.Range('S23').Value))
            esperaTexto = ' / '.join(
                {'1-3S': '1_3s', '2-2S': '2_2s', 'R4S': 'R_4s',
                 '4-1S': '4_1s', '8X': '8x'}[r] for r in regras)

            reg('A', 'Sigma %.2f -> Sigma_Plano' % sg, sg, sg, sp, perto(sp, sg))
            reg('A', 'Sigma %.2f -> classificacao' % sg, sg, classe, gc,
                gc == classe)
            reg('A', 'Sigma %.2f -> regras (texto)' % sg, sg, esperaTexto, gr,
                gr == esperaTexto)
            reg('A', 'Sigma %.2f -> regras (realce G6:K6)' % sg, sg, regras, ac,
                ac == regras)
            reg('A', 'Sigma %.2f -> N' % sg, sg, n if n != '' else '(vazio)',
                gn if gn else '(vazio)',
                (perto(gn.replace('.', '').replace(',', '.'), n)
                 if n != '' else gn == ''))
            reg('A', 'Sigma %.2f -> Run Size' % sg, sg,
                run if run != '' else '(vazio)', gu if gu else '(vazio)',
                (perto(gu.replace('.', '').replace(',', '.'), run)
                 if run != '' else gu == ''))
            reg('A', 'Sigma %.2f -> frequencia preenchida' % sg, sg,
                'texto nao vazio', (gf[:30] + '...') if gf else '(vazio)',
                gf != '')
            if sg < 3:
                reg('A', 'Sigma %.2f -> REAVALIAR METODO visivel' % sg, sg,
                    'REAVALIAR MÉTODO na base', base[-40:],
                    'REAVALIAR' in base.upper())

        # ================================================================
        print()
        print('=' * 78)
        print('B. PIOR NIVEL GOVERNA')
        print('=' * 78)
        for n1, n2, esp, nivel in PIOR:
            poe(n1, n2)
            sp = tenta(lambda: cfg.Range('B1').Value)
            gov = txt(tenta(lambda: cfg.Range('D2').Value))
            ok = (perto(sp, esp) if isinstance(esp, float) else txt(sp) == esp)
            reg('B', 'N1=%s N2=%s -> Sigma_Plano' % (n1, n2),
                '%s / %s' % (n1, n2), esp, sp, ok)
            if nivel:
                reg('B', 'N1=%s N2=%s -> nivel governante' % (n1, n2),
                    '%s / %s' % (n1, n2), nivel, gov, gov == nivel)
        # zero e vazio nao sao Sigma valido
        poe(0.0, None)
        reg('B', 'zero NAO e tratado como ausencia', '0 / vazio', 0.0,
            tenta(lambda: cfg.Range('B1').Value),
            perto(tenta(lambda: cfg.Range('B1').Value), 0.0))
        poe(None, None)
        reg('B', 'sem nenhum nivel valido -> SEM DADOS', 'vazio / vazio',
            'SEM DADOS', txt(tenta(lambda: cfg.Range('B1').Value)),
            txt(tenta(lambda: cfg.Range('B1').Value)) == 'SEM DADOS')
        reg('B', 'e o plano NAO inventa regra', 'vazio / vazio', '(vazio)',
            txt(tenta(lambda: pa.Range('T22').Value)) or '(vazio)',
            txt(tenta(lambda: pa.Range('T22').Value)) == '')

        # ================================================================
        print()
        print('=' * 78)
        print('D. REALCE EFETIVO DE G6:K6 (DisplayFormat, nao a formula)')
        print('=' * 78)
        for sg, classe, regras, _n, _u in [ESCADA[i] for i in (0, 2, 4, 6, 8)]:
            poe(sg, sg)
            zeraViolacoes()
            for ref, rot in REGRAS:
                fundo, fonte, negrito, italico = efetivo(ref)
                deveAcender = rot in regras
                if deveAcender:
                    ok = (fundo == VERDE and fonte == BRANCO
                          and negrito and italico)
                    reg('D', 'Sigma %.2f: %s recomendada = verde/branca/N+I'
                        % (sg, rot), sg, '#146C43 branca N+I',
                        '#%06X f#%06X N=%s I=%s'
                        % (fundo, fonte, negrito, italico), ok)
                else:
                    ok = fundo != VERDE
                    reg('D', 'Sigma %.2f: %s neutra (nao verde)' % (sg, rot),
                        sg, 'fundo != #146C43', '#%06X' % fundo, ok)

        # ================================================================
        print()
        print('=' * 78)
        print('E. VIOLACAO VENCE RECOMENDACAO (demonstrado, nao inferido)')
        print('=' * 78)
        poe(5.50, 5.50)          # 1-3S recomendada
        zeraViolacoes()
        f0 = efetivo('G6')
        reg('E', 'Sigma 5,50 sem violacao: 1-3S verde', '5,50 / 0 violacoes',
            '#%06X' % VERDE, '#%06X' % f0[0], f0[0] == VERDE)
        tenta(lambda: pa.Range('G7').__setattr__('Value', 3))
        tenta(lambda: xl.CalculateFull())
        f1 = efetivo('G6')
        reg('E', 'com violacao real: 1-3S deixa de ser verde',
            'G7 = 3 violacoes', 'fundo != #146C43', '#%06X' % f1[0],
            f1[0] != VERDE)
        reg('E', 'e assume a aparencia de VIOLACAO', 'G7 = 3 violacoes',
            '#%06X' % VERM_FUNDO, '#%06X' % f1[0], f1[0] == VERM_FUNDO)
        reg('E', 'a regra continua recomendada pelo Sigma', 'G7 = 3',
            '1-3S entre as acesas', acesas()[:1], '1-3S' in acesas())
        tenta(lambda: pa.Range('G7').__setattr__('Value', 0))
        tenta(lambda: xl.CalculateFull())
        f2 = efetivo('G6')
        reg('E', 'removida a violacao, volta ao verde', 'G7 = 0',
            '#%06X' % VERDE, '#%06X' % f2[0], f2[0] == VERDE)

        # ================================================================
        print()
        print('=' * 78)
        print('F. M7/M8 SAO INDEPENDENTES DO PLANO SIGMA')
        print('=' * 78)
        poe(3.50, 3.50)
        zeraViolacoes()
        reg('F', 'Sigma 3,50 recomenda as cinco regras', '3,50', TODAS5,
            acesas(), acesas() == TODAS5)
        m7 = txt(tenta(lambda: pa.Range('M7').Text))
        m8 = txt(tenta(lambda: pa.Range('M8').Text))
        reg('F', 'sem violacao: M7 nao reprova', '3,50 / 0 violacoes',
            'Sem violação', m7, m7 == 'Sem violação')
        reg('F', 'sem violacao: M8 nao reprova', '3,50 / 0 violacoes',
            'Sem violação', m8, m8 == 'Sem violação')
        tenta(lambda: pa.Range('H7').__setattr__('Value', 4))
        tenta(lambda: xl.CalculateFull())
        m7b = txt(tenta(lambda: pa.Range('M7').Text))
        m8b = txt(tenta(lambda: pa.Range('M8').Text))
        reg('F', 'violacao no N1: M7 muda', 'H7 = 4',
            'REPROVA — 4 violação(ões)', m7b,
            m7b == 'REPROVA — 4 violação(ões)')
        reg('F', 'violacao no N1: M8 NAO muda', 'H7 = 4', 'Sem violação',
            m8b, m8b == 'Sem violação')
        tenta(lambda: pa.Range('H7').__setattr__('Value', 0))
        tenta(lambda: pa.Range('J8').__setattr__('Value', 2))
        tenta(lambda: xl.CalculateFull())
        reg('F', 'violacao no N2: so M8 muda', 'J8 = 2',
            'M7 sem violação / M8 reprova',
            '%s | %s' % (txt(tenta(lambda: pa.Range('M7').Text)),
                         txt(tenta(lambda: pa.Range('M8').Text))),
            txt(tenta(lambda: pa.Range('M7').Text)) == 'Sem violação'
            and 'REPROVA' in txt(tenta(lambda: pa.Range('M8').Text)))
        # Sigma sozinho nao reprova
        poe(1.50, 1.50)
        zeraViolacoes()
        reg('F', 'Sigma 1,50 e DPM alto NAO reprovam a corrida', '1,50',
            'Sem violação nos dois',
            '%s | %s' % (txt(tenta(lambda: pa.Range('M7').Text)),
                         txt(tenta(lambda: pa.Range('M8').Text))),
            txt(tenta(lambda: pa.Range('M7').Text)) == 'Sem violação'
            and txt(tenta(lambda: pa.Range('M8').Text)) == 'Sem violação')

        # ---- devolve as formulas mexidas -------------------------------
        tenta(lambda: pa.Range('V10').__setattr__('Formula', fV10))
        tenta(lambda: pa.Range('V11').__setattr__('Formula', fV11))
        for lin, cols in ((7, ('K', 'L', 'M', 'N', 'O')),
                          (8, ('AG', 'AH', 'AI', 'AJ', 'AK'))):
            for k, cl in enumerate(cols):
                tenta(lambda x=lin, y=7 + k, c=cl: pa.Cells(x, y).__setattr__(
                    'Formula',
                    '=SUMPRODUCT(Calc!${0}$3:${0}$182,Calc!$D$3:$D$182)'.format(c)))
        tenta(lambda: xl.CalculateFull())
        reg('F', 'formulas originais do bloco Westgard devolvidas',
            'restauracao', fG7[:30], txt(tenta(lambda: pa.Range('G7').Formula))[:30],
            txt(tenta(lambda: pa.Range('G7').Formula)) == txt(fG7))

        # ================================================================
        print()
        print('=' * 78)
        print('C / H / I. ANALITOS REAIS: cadeia, ET x ETp e plano')
        print('=' * 78)
        tenta(lambda: es.Range('L4').__setattr__('Value', 'CAP'))
        tenta(lambda: es.Range('N4').__setattr__('Value', 2025))
        tenta(lambda: es.Range('P4').__setattr__('Value', 'TODAS'))
        idx = {}
        for k in range(4, 44):
            nm = txt(tenta(lambda x=k: an.Cells(x, 1).Value))
            if nm:
                idx[nm] = k - 3
        for nome in ANALITOS:
            if nome not in idx:
                reg('C', '%s existe na Analitos' % nome[:22], nome, 'presente',
                    'AUSENTE', False)
                continue
            tenta(lambda i=idx[nome]: pa.Range('B3').__setattr__('Value', i))
            tenta(lambda: xl.CalculateFull())
            s1 = tenta(lambda: pa.Range('V10').Value)
            s2 = tenta(lambda: pa.Range('V11').Value)
            sp = tenta(lambda: cfg.Range('B1').Value)
            gov = txt(tenta(lambda: cfg.Range('D2').Value))
            cls = txt(tenta(lambda: xl.Run('PlanoQC', sp, 'CLASSE')))
            ac = acesas()
            gn = txt(tenta(lambda: pa.Range('U22').Text))
            gu = txt(tenta(lambda: pa.Range('V22').Text))
            gf = txt(tenta(lambda: pa.Range('W22').Value))
            print()
            print('   %s' % nome)
            print('      Sigma N1=%s  N2=%s  ->  Sigma_Plano=%s  (%s)'
                  % (('%.3f' % s1) if isinstance(s1, (int, float)) else '—',
                     ('%.3f' % s2) if isinstance(s2, (int, float)) else '—',
                     ('%.3f' % sp) if isinstance(sp, (int, float)) else sp, gov))
            print('      classificação=%s   regras=%s' % (cls, '/'.join(ac)))
            print('      G6..K6 = %s' % ' '.join(
                '%s:%s' % (rot, {VERDE: 'VERDE', VERM_FUNDO: 'VERMELHA'}.get(
                    efetivo(ref)[0], 'neutra'))
                for ref, rot in REGRAS))
            print('      N=%s  Run=%s  freq=%s' % (gn or '—', gu or '—', gf[:44]))

            esperado = min([x for x in (s1, s2) if isinstance(x, (int, float))])
            reg('C', '%s: Sigma_Plano = pior nivel' % nome[:22],
                'N1 %s / N2 %s' % (s1, s2), esperado, sp, perto(sp, esperado))
            # o realce EFETIVO acompanha a escada -- descontada a violacao
            #
            # Com dado real, "recomendada" e "verde" NAO sao a mesma coisa. A
            # violacao tem prioridade maior que a recomendacao (ADR-037), entao
            # uma regra recomendada E violada aparece VERMELHA, nao verde. So o
            # bloco D ve todas as recomendadas verdes, porque zera as contagens
            # de proposito.
            #
            # A expectativa honesta e: verde onde recomendada e sem violacao,
            # vermelha onde violada, neutra no resto.
            faixa = ([r for r in ESCADA if r[0] <= sp] or [ESCADA[0]])[-1]
            espRegras = faixa[2] if sp >= 3 else TODAS5
            viol = {}
            for i, (ref, rot) in enumerate(REGRAS):
                v1 = tenta(lambda c=7 + i: pa.Cells(7, c).Value) or 0
                v2 = tenta(lambda c=7 + i: pa.Cells(8, c).Value) or 0
                viol[rot] = (v1 + v2) > 0
            espVerdes = [r for r in espRegras if not viol[r]]
            verdes = [rot for ref, rot in REGRAS if efetivo(ref)[0] == VERDE]
            reg('C', '%s: verde = recomendada e sem violacao' % nome[:22],
                'Sigma %s / violadas %s'
                % (sp, [r for r in espRegras if viol[r]] or 'nenhuma'),
                espVerdes, verdes, verdes == espVerdes)
            # e toda regra violada tem de estar vermelha, recomendada ou nao
            vermelhas = [rot for ref, rot in REGRAS
                         if efetivo(ref)[0] == VERM_FUNDO]
            espVerm = [rot for rot in
                       [r for _, r in REGRAS] if viol[rot]]
            reg('C', '%s: violada -> vermelha' % nome[:22],
                'violadas %s' % (espVerm or 'nenhuma'),
                espVerm, vermelhas, vermelhas == espVerm)
            # H. ET x ETp
            vals = [txt(tenta(lambda c=c: pa.Cells(17, c).Text))
                    for c in range(19, 24)]
            reg('H', '%s: ET x ETp do N1 preenchido' % nome[:22], nome,
                '5 campos com valor', vals, all(v != '' for v in vals))
            vals2 = [txt(tenta(lambda c=c: pa.Cells(18, c).Text))
                     for c in range(19, 24)]
            reg('H', '%s: ET x ETp do N2 preenchido' % nome[:22], nome,
                '5 campos com valor', vals2, all(v != '' for v in vals2))
            # I. plano coerente com o pior nivel
            if sp >= 3:
                reg('I', '%s: plano tem N e Run Size' % nome[:22], sp,
                    'ambos preenchidos', 'N=%s Run=%s' % (gn, gu),
                    gn != '' and gu != '')
            else:
                reg('I', '%s: abaixo de 3 Sigma, sem N e sem Run Size'
                    % nome[:22], sp, 'ambos vazios',
                    'N=%r Run=%r' % (gn, gu), gn == '' and gu == '')
            reg('I', '%s: frequencia/orientacao preenchida' % nome[:22], nome,
                'texto nao vazio', gf[:26] or '(vazio)', gf != '')
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass

    # ================================================================
    print()
    print('=' * 78)
    print('MATRIZ FINAL DE VALIDACAO')
    print('=' * 78)
    blocos = {}
    for b, t, e, esp, obt, ok in RESULTADOS:
        blocos.setdefault(b, [0, 0])
        blocos[b][0] += 1
        if ok:
            blocos[b][1] += 1
    rotulos = {'A': 'Fronteiras Sigma', 'B': 'Pior nivel',
               'C': 'Analitos reais', 'D': 'Realce G6:K6',
               'E': 'Violacao > recomendacao', 'F': 'Independencia M7/M8',
               'H': 'ET x ETp', 'I': 'Plano QC'}
    for b in sorted(blocos):
        tot, ok = blocos[b]
        print('   %s. %-26s %3d/%3d  %s'
              % (b, rotulos.get(b, b), ok, tot,
                 'PASS' if ok == tot else 'FAIL'))
    falhas = [r for r in RESULTADOS if not r[5]]
    print()
    print('   TOTAL: %d de %d' % (len(RESULTADOS) - len(falhas), len(RESULTADOS)))
    if falhas:
        print()
        print('FALHAS:')
        for b, t, e, esp, obt, _ in falhas:
            print('   [%s] %s | entrada=%s | esperado=%s | obtido=%s'
                  % (b, t, e, esp, obt))
        sys.exit(1)
    print()
    print('VALIDACAO FINAL: PASS')


if __name__ == '__main__':
    main(sys.argv[1])
