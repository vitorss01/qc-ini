# -*- coding: utf-8 -*-
"""t_cinco_anos.py -- CINCO ANOS DE USO, com o numero de lotes do laboratorio.

Uso: python t_cinco_anos.py bio  <QC_Bioquimica_instalada.xlsm>
     python t_cinco_anos.py hema <QC_Hematologia_instalada.xlsm>

Premissas (pedido do gestor, 19/09/2026):
  Bioquimica ... 4 lotes por ano = 20 lotes; 31 analitos x 2 niveis; 1 corrida
                 por dia util e uma 2a corrida em 40% dos dias (~110 mil
                 resultados -- o pior caso do ADR-025).
  Hematologia .. 8 lotes por ano = 40 lotes; 28 analitos x 3 niveis; 1 corrida
                 TODO dia, inclusive fim de semana (~153 mil resultados).

Monta o historico a partir do arquivo instalado (dados operacionais zerados),
digita media/DP de CADA lote pela aba Analitos (o caminho do usuario), grava o
banco em bloco + AtualizarFlagsBanco, salva, REABRE e mede o que o usuario sente:
abrir, clicar no spinner, trocar de lote (em analise e em uso), lancar a
corrida do dia, motor completo. E confere, em lotes do inicio, do meio e do fim,
que grafico, z, limites e Painel sao DESTE lote.
"""
import datetime as dt
import json
import os
import random
import shutil
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import xlh  # noqa: E402

AQUI = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(os.environ['TMP'], 'qcw', 'cinco')
os.makedirs(TMP, exist_ok=True)
NA = -2146826246
CRLF, LF = chr(13) + chr(10), chr(10)

P = {
    'bio': dict(nlv=2, lotes_ano=4, todo_dia=False, seg=0.40, cel_lote='H3', faixa='A39',
                val=('F', 'AB'), z=('G', 'AC'), st=('P', 'AL'),
                lim=(('Q', 'T', 'W'), ('AM', 'AP', 'AS')),
                par=(('AX', 'AY'), ('AZ', 'BA')), inicio=dt.date(2026, 1, 5)),
    'hema': dict(nlv=3, lotes_ano=8, todo_dia=True, seg=0.0, cel_lote='H3', faixa='O3',
                 val=('F', 'AB', 'AX'), z=('G', 'AC', 'AY'), st=('P', 'AL', 'BH'),
                 lim=(('Q', 'T', 'W'), ('AM', 'AP', 'AS'), ('BI', 'BL', 'BO')),
                 par=(('BT', 'BU'), ('BV', 'BW'), ('BX', 'BY')), inicio=dt.date(2026, 1, 1)),
}
RES = []
MED = {}


def chk(nome, ok, det=''):
    RES.append((nome, bool(ok), str(det)[:300]))
    print(f'   [{"PASS" if ok else "FAIL"}] {nome}' + ('' if ok else f'  -> {str(det)[:300]}'), flush=True)
    return ok


def medir(k, v):
    MED[k] = v
    print(f'   MEDIDA {k:40s} {v}', flush=True)


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and v > -2146000000


def igual(a, b, tol=1e-7):
    return num(a) and num(b) and abs(a - b) <= tol * max(1.0, abs(b))


def cl(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n - 1          # 0-based


# ---------------------------------------------------------------------------
def gerar(cfg, analitos):
    rnd = random.Random(20260919)
    n_lotes = cfg['lotes_ano'] * 5
    dias = []
    d = cfg['inicio']
    fim = dt.date(d.year + 5, d.month, d.day)
    while d < fim:
        if cfg['todo_dia'] or d.weekday() < 5:
            dias.append(d)
        d += dt.timedelta(days=1)
    por_lote = len(dias) // n_lotes
    lotes = []
    for i in range(n_lotes):
        ds = dias[i * por_lote:(i + 1) * por_lote] if i < n_lotes - 1 else dias[i * por_lote:]
        lotes.append((f'L{i + 1:02d}', ds))
    params = {}
    for li, (nome, _) in enumerate(lotes):
        for ai, an in enumerate(analitos):
            p = []
            for t in range(cfg['nlv']):
                m = 10.0 * (ai + 1) * (1 + t) * (1 + 0.015 * (li % 7))
                p += [round(m, 4), round(m * (0.02 + 0.005 * t), 5)]
            params[(nome, an)] = tuple(p)
    linhas, run = [], 0
    for nome, ds in lotes:
        for dd in ds:
            k_max = 2 if rnd.random() < cfg['seg'] else 1
            for k in range(k_max):
                run += 1
                quando = dt.datetime(dd.year, dd.month, dd.day, 8 + 6 * k)
                for an in analitos:
                    p = params[(nome, an)]
                    for t in range(cfg['nlv']):
                        z = max(-2.4, min(2.4, rnd.gauss(0, 0.8)))
                        linhas.append((run, quando, t + 1, f'QC-{nome}{t + 1:02d}', an,
                                       round(p[2 * t] + z * p[2 * t + 1], 5), 'Ativo'))
    return lotes, params, linhas


# ---------------------------------------------------------------------------
class S:
    def __init__(self, caminho):
        self.caminho = caminho
        t0 = time.time()
        self.ex = xlh.Excel()
        self.wb = self.ex.abrir(caminho)
        self.t_abrir = time.time() - t0
        self.xl = self.ex.xl
        self.sh = {s.Name: s for s in self.wb.Worksheets}

    def run(self, m, *a, teto=900):
        r = self.ex.run(m, *a, teto=teto)
        self.ex.esperar()
        return r

    def fechar(self):
        self.ex.fechar()


def instalar_teste(s):
    vbp = s.wb.VBProject
    with open(os.path.join(AQUI, 'mTesteEtapa1.bas'), encoding='utf-8') as f:
        cod = CRLF.join(l for l in f.read().replace(CRLF, LF).split(LF) if not l.startswith('Attribute '))
    comp = None
    for c in vbp.VBComponents:
        if c.Name == 'mTesteEtapa1':
            comp = c
    if comp is None:
        comp = vbp.VBComponents.Add(1)
        comp.Name = 'mTesteEtapa1'
    cm = comp.CodeModule
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(cod)
    if 'T_Entrada' not in s.sh:
        try:
            s.wb.Unprotect('qcini2025')           # estrutura protegida: Add falharia
        except Exception:
            pass
        ws = s.wb.Worksheets.Add()
        ws.Name = 'T_Entrada'
        s.sh['T_Entrada'] = ws


def montar(prod, src):
    cfg = P[prod]
    dst = os.path.join(TMP, f'{prod}_5anos.xlsm')
    shutil.copyfile(src, dst)
    s = S(dst)
    try:
        for w in s.wb.Worksheets:
            if w.ProtectContents:
                w.Unprotect('qcini2025')
        instalar_teste(s)
        pai = s.sh['Painel']
        pai.Range('G3').Value = ''
        pai.Range('G4').Value = ''
        if prod == 'bio':
            pai.Range('I3').Value = ''
        pai.Range('M3:N4').Value = False
        s.run('mTesteEtapa1.T_LimparBase')
        ana = s.sh['Analitos']
        analitos = [ana.Cells(r, 1).Value for r in range(4, 44) if ana.Cells(r, 1).Value]
        lotes, params, linhas = gerar(cfg, analitos)
        print(f'{prod}: {len(lotes)} lotes, {len(analitos)} analitos x {cfg["nlv"]} niveis, '
              f'{len(linhas)} resultados, {linhas[-1][0]} corridas, '
              f'{linhas[0][1]:%d/%m/%Y} a {linhas[-1][1]:%d/%m/%Y}', flush=True)
        for nome, _ in lotes:
            s.run('mTesteEtapa1.T_Cadastrar', nome)
        s.run('mTesteEtapa1.T_LoteEmUso', lotes[-1][0])
        t0 = time.time()
        ncol = 2 * cfg['nlv']
        ultc = chr(ord('E') + ncol - 1)
        for nome, _ in lotes:
            s.run('mTesteEtapa1.T_Selecionar', nome)
            ana.Range(f'E4:{ultc}{3 + len(analitos)}').Value = [list(params[(nome, an)]) for an in analitos]
            s.run('mLotes.ParametroEditado', ana.Range(f'E4:{ultc}{3 + len(analitos)}'))
        medir('digitar_parametros_todos_os_lotes_s', round(time.time() - t0, 1))
        db = s.sh['DB_Resultados']
        t0 = time.time()
        s.xl.Calculation = -4135
        for i in range(0, len(linhas), 10000):
            parte = linhas[i:i + 10000]
            r0 = 4 + i
            db.Range(f'D{r0}:D{r0 + len(parte) - 1}').NumberFormat = '@'
            db.Range(f'A{r0}:G{r0 + len(parte) - 1}').Value = parte
        s.run('mBanco.AtualizarFlagsBanco', teto=1800)
        s.run('mDados.AtualizarListasAno', teto=600)
        s.xl.Calculation = -4105
        s.run('mEstatistica.AtualizarEstatistica', teto=1800)
        medir('gravar_banco_5anos_s', round(time.time() - t0, 1))
        s.wb.Save()
    finally:
        s.fechar()
    medir('tamanho_arquivo_MB', round(os.path.getsize(dst) / 1e6, 2))
    return dst, lotes, params, analitos


# ---------------------------------------------------------------------------
def conferir(prod, s, lote, analito, p, rot):
    cfg = P[prod]
    s.xl.Calculate()
    calc = s.sh['Calc']
    L = calc.Range('A1:CI182').Value
    ws = s.sh['DB_Resultados']
    ult = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    dbv = ws.Range(f'A4:G{ult}').Value
    db = [r for r in dbv if r[4] == analito and str(r[3])[3:-2] == lote and r[6] == 'Ativo']
    dia = lambda d: d.replace(tzinfo=None)
    corr = sorted({(dia(r[1]), int(r[0])) for r in db})
    janela = {rn for _, rn in corr[-180:]}
    val = {(int(r[0]), int(r[2])): r[5] for r in db}
    linhas = [r for r in L[2:] if r[1] not in (None, '')]
    runs = [int(r[1]) for r in linhas]
    pre = f'[{rot}] {lote} {analito}: '
    ok = chk(pre + f'corridas no grafico = as {len(janela)} mais recentes do lote ({len(corr)} no total)',
             set(runs) == janela and len(runs) == len(set(runs)), (len(runs), len(janela)))
    datas = [r[2] for r in linhas]
    ok &= chk(pre + 'ordem cronologica', datas == sorted(datas))
    par = []
    for (cm, cs) in cfg['par']:
        par += [L[0][cl(cm)], L[0][cl(cs)]]
    ok &= chk(pre + f'media/DP do Calc = os DESTE lote {p}', all(igual(a, b) for a, b in zip(par, p)), par)
    ev, ez, el = [], [], []
    for r in linhas:
        rn = int(r[1])
        for t in range(cfg['nlv']):
            v = r[cl(cfg['val'][t])]
            if not igual(v, val.get((rn, t + 1))):
                ev.append((rn, t + 1, v, val.get((rn, t + 1))))
                continue
            m, sd = p[2 * t], p[2 * t + 1]
            if not igual(r[cl(cfg['z'][t])], (v - m) / sd, 1e-6):
                ez.append((rn, t + 1, r[cl(cfg['z'][t])]))
            if r[3] == 1:
                for col, esp in zip(cfg['lim'][t], (m - 3 * sd, m, m + 3 * sd)):
                    if not igual(r[cl(col)], esp):
                        el.append((rn, t + 1, col))
    ok &= chk(pre + 'valores plotados = banco (todos os niveis)', not ev, ev[:3])
    ok &= chk(pre + 'z com a media/DP deste lote', not ez, ez[:3])
    ok &= chk(pre + 'linhas -3s/media/+3s deste lote', not el, el[:3])
    o3 = str(s.sh['Painel'].Range(cfg['faixa']).Value)
    ok &= chk(pre + 'faixa de status do Painel mostra o lote', o3.startswith(f'Lote {lote} · alvo'), o3[:80])
    return ok


def cliques(s, lote, n=10):
    pai = s.sh['Painel']
    s.run('mTesteEtapa1.T_Selecionar', lote)
    tt = []
    idx = [3, 7, 1, 12, 5, 20, 2, 15, 9, 4, 18, 6][:n]
    for i in idx:
        t0 = time.time()
        pai.Range('B3').Value = i
        s.run('PainelMudou')
        tt.append(round(time.time() - t0, 3))
    return tt


def usar(prod, arq, lotes, params, analitos):
    cfg = P[prod]
    s = S(arq)
    medir('abrir_arquivo_s', round(s.t_abrir, 1))
    try:
        t0 = time.time(); s.xl.CalculateFull(); s.ex.esperar()
        medir('recalculo_completo_s', round(time.time() - t0, 1))
        # ---- spinner ----
        ult = lotes[-1][0]
        tt = cliques(s, ult)
        medir(f'spinner_lote_{ult}_cliques_s', tt)
        medir('spinner_pior_clique_s', max(tt))
        chk('spinner: TODO clique abaixo de 3 s (pedido do gestor)', max(tt) < 3.0, max(tt))
        chk('spinner: todo clique abaixo de 0,5 s (meta interna)', max(tt) < 0.5, max(tt))
        meio = lotes[len(lotes) // 2][0]
        tt2 = cliques(s, meio, 6)
        medir(f'spinner_lote_{meio}_cliques_s', tt2)
        # ---- troca de lote em analise ----
        tl = []
        for lote, _ in (lotes[0], lotes[len(lotes) // 3], lotes[-1], lotes[len(lotes) // 2], lotes[-2]):
            t0 = time.time(); s.run('mTesteEtapa1.T_Selecionar', lote); tl.append(round(time.time() - t0, 2))
        medir('trocar_lote_em_analise_s', tl)
        chk('troca de lote em analise abaixo de 3 s', max(tl) < 3.0, tl)
        # ---- lote em uso ----
        t0 = time.time(); s.run('mTesteEtapa1.T_LoteEmUso', lotes[-2][0])
        medir('trocar_lote_em_uso_s', round(time.time() - t0, 2))
        s.run('mTesteEtapa1.T_LoteEmUso', lotes[-1][0])
        # ---- conferencias ----
        s.run('mTesteEtapa1.T_Analito', analitos[0])
        for lote, rot in ((lotes[-1][0], 'ultimo lote'), (lotes[len(lotes) // 2][0], 'lote do meio'),
                          (lotes[0][0], 'primeiro lote')):
            s.run('mTesteEtapa1.T_Selecionar', lote)
            conferir(prod, s, lote, analitos[0], params[(lote, analitos[0])], rot)
        s.run('mTesteEtapa1.T_Analito', analitos[-1])
        s.run('mTesteEtapa1.T_Selecionar', lotes[-1][0])
        conferir(prod, s, lotes[-1][0], analitos[-1], params[(lotes[-1][0], analitos[-1])], 'ultimo analito')
        # ---- corrida do dia (lancamento) ----
        te = s.sh['T_Entrada']
        dia_novo = lotes[-1][1][-1] + dt.timedelta(days=1)
        lin = [['Data', 'Nivel', 'Lote', 'Analito', 'Valor']]
        for an in analitos:
            p = params[(lotes[-1][0], an)]
            for t in range(cfg['nlv']):
                lin.append([dt.datetime(dia_novo.year, dia_novo.month, dia_novo.day), t + 1, lotes[-1][0], an, p[2 * t]])
        te.Range(f'A1:E{len(lin)}').Value = lin
        t0 = time.time()
        r = s.run('mTesteEtapa1.T_GravarOperacao')
        medir(f'lancar_corrida_do_dia_{len(lin) - 1}_resultados_s', round(time.time() - t0, 2))
        chk(f'corrida do dia gravada: {r}', str(r).startswith(str(len(lin) - 1)), r)
        # ---- motor completo ----
        t0 = time.time(); s.run('mEstatistica.AtualizarEstatistica')
        medir('motor_completo_s', round(time.time() - t0, 2))
        db = s.sh['DB_Resultados']
        medir('linhas_no_banco', db.Cells(db.Rows.Count, 1).End(-4162).Row - 3)
        medir('linhas_livres_ate_o_teto', s.run('mBanco.LinhasLivres'))
        r = s.run('mBanco.TestarCapacidade', 1)
        chk(f'banco ainda aceita gravacao: {r}', str(r).upper().startswith('PERMIT'), r)
        r = s.run('mLotes.ConferirLotesStore')
        chk(f'LotesStore integra com todos os lotes: {r}', str(r).startswith('ok'), r)
    finally:
        s.fechar()


def main(prod, src):
    t00 = time.time()
    arq, lotes, params, analitos = montar(prod, os.path.abspath(src))
    usar(prod, arq, lotes, params, analitos)
    npass = sum(1 for r in RES if r[1])
    print(f'\nMEDIDAS ({prod}):')
    for k, v in MED.items():
        print(f'   {k:40s} {v}')
    print(f'\nTOTAL {prod}: {npass}/{len(RES)} PASS em {time.time() - t00:.0f}s')
    out = os.path.join(AQUI, '..', 'resultados', f'cinco_anos_{prod}_{time.strftime("%Y%m%d_%H%M%S")}.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump({'medidas': MED, 'verificacoes': RES}, f, ensure_ascii=False, indent=1, default=str)
    print('resultado:', os.path.abspath(out))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
