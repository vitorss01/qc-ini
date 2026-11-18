# -*- coding: utf-8 -*-
"""repro_OC-0084.py -- Tester/QA, ciclo 2026-11-18_reteste08, modo VALIDAR (ADR-098, OC-0084).

Roteiro proprio (nao reaproveita o script do Usuario Simulado nem o qa_adr097). Usa so a infraestrutura qa_final.QA.

Uso: python -u repro_OC-0084.py <Bioquimica|Hematologia> <pasta_trabalho> <testes_do_commit> <calc_meta.json>

FASE A (copia nova da producao, eventos LIGADOS): digita 4 manuais celula a celula no lote/equipamento/janela da
  Estatistica (M1, P, Q no nivel 1; S no nivel 2 com SUBSTITUI_ID de um interfaceamento real); troca o dono e a hora
  do carimbo (eventos desligados, como se outro usuario os tivesse lancado: TESTE-AUDITORIA_U1..U4); ATUALIZAR #1;
  confere integracao; salva a COPIA e fecha.
FASE B (reabre a copia: sem retrato; o 1o retrato vem do SelectionChange, como na mao do usuario):
  K1 M1 recortado (Selection.Cut + ActiveSheet.Paste) para linha vazia
  K2 bloco P,Q sobe uma linha (Selection.Cut + Paste sobre linha ocupada); carimbo de P anterior a DATA/HORA de Q
  K3 S (com SUBSTITUI_ID) recortado com Range.Cut(destino)
  K4 redigitar o ID de S (mesmo texto e "man_N"); manual N nao integrado com carimbo 2020 redigitado
  ATUALIZAR #2; K6 numeros (Estatistica x oraculo proprio); vizinhos K5 (sem ATUALIZAR).
Dados sinteticos: valores = resultado real x fator, 3 decimais; donos TESTE-AUDITORIA_Ux. Nada e salvo na producao.
"""
import datetime as dt
import json
import os
import shutil
import statistics as stt
import sys
import time
import traceback

PROD, W, TESTES, META = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
sys.path.insert(0, TESTES)
sys.path.insert(0, os.path.join(TESTES, '..'))
sys.path.insert(0, os.path.join(TESTES, '..', '..', 'etapa1_multilote'))
from qa_final import QA, sha256_arquivo  # noqa: E402
import instalar_integracao as ii  # noqa: E402
import camada_dados as cd  # noqa: E402

RAIZ = os.path.abspath(os.path.join(TESTES, '..', '..', '..'))
REPO = os.environ.get('QC_RAIZ_PROD')
ARQ = 'QC_Bioquimica.xlsm' if PROD.startswith('Bio') else 'QC_Hematologia.xlsm'
TB = 'tblResultados_Manuais'
OUT = {'produto': PROD, 'casos': {}, 'log': []}
SAIDA = os.path.join(W, f'repro_OC-0084_{PROD}.json')


def log(*a):
    s = ' '.join(str(x) for x in a)
    print(s, flush=True)
    OUT['log'].append(s)


def caso(cid, ok, evid):
    OUT['casos'][cid] = {'resultado': 'APROVADO' if ok else 'REPROVADO', 'evidencia': evid}
    log(('APROVADO ' if ok else 'REPROVADO ') + cid, json.dumps(evid, ensure_ascii=False, default=str)[:900])


def gravar():
    with open(SAIDA, 'w', encoding='utf-8') as f:
        json.dump(OUT, f, ensure_ascii=False, indent=1, default=str)


def txt(v):
    return '' if v is None else str(v).strip()


def ing(v):
    return v.replace(tzinfo=None) if getattr(v, 'tzinfo', None) else v


def nlote(v):
    s = txt(v)
    return s[:-2] if s.endswith('.0') and s[:-2].isdigit() else s


def base_id(v):
    return txt(v).split('~')[0]


meta = json.load(open(META, encoding='utf-8'))['meta']
LOTE, EQUIP = nlote(meta['lote']), txt(meta['equipamento'])
if 'janela' in meta and meta['janela'][0]:
    J0, J1 = [dt.datetime.fromisoformat(x) for x in meta['janela']]
else:
    J0, J1 = dt.datetime(int(meta['ano_de']), 1, 1), dt.datetime(int(meta['ano_ate']), 12, 31)
J1 = J1 + dt.timedelta(days=1) - dt.timedelta(seconds=1)
log('lote', LOTE, 'equip', EQUIP, 'janela', J0, J1)


class Aba:
    def __init__(self, q):
        self.q, self.xl = q, q.ex.xl
        self.lo = q.lo(TB)
        self.ws = self.lo.Parent
        self.cols = [c.Name for c in self.lo.ListColumns]
        self.c_mot = self.cols.index('MOTIVO') + 1

    def cel(self, r, col):
        return self.lo.ListColumns(col).DataBodyRange.Cells(r, 1)

    def bloco(self, r1, r2=None):
        return self.ws.Range(self.cel(r1, 'ID_REGISTRO'), self.lo.DataBodyRange.Cells(r2 or r1, self.c_mot))

    def car(self, r):
        return txt(self.cel(r, 'USUARIO').Value), ing(self.cel(r, 'REGISTRADO_EM').Value)

    def sub(self, r):
        return txt(self.cel(r, 'SUBSTITUI_ID').Value) if 'SUBSTITUI_ID' in self.cols else None

    def idr(self, r):
        return txt(self.cel(r, 'ID_REGISTRO').Value)

    def estado(self, r):
        c = self.car(r)
        return {'linha': r, 'id': self.idr(r), 'usuario': c[0], 'registrado_em': str(c[1]), 'subst': self.sub(r),
                'resultado': self.cel(r, 'RESULTADO').Value}

    def eventos_off_set(self, r, col, v):
        self.xl.EnableEvents = False
        if isinstance(v, dt.datetime):      # serial do Excel: o pywin32 desloca o datetime ingenuo pelo fuso (+3 h)
            self.cel(r, col).Value2 = ser(v)
        else:
            self.cel(r, col).Value = v
        self.xl.EnableEvents = True

    def digitar(self, r, campos):
        self.xl.EnableEvents = True
        for col in ('DATA', 'HORA', 'EQUIPAMENTO', 'MATRIZ', 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO', 'MOTIVO'):
            if col in campos and campos[col] is not None and col in self.cols:
                v = campos[col]
                if isinstance(v, dt.datetime):
                    self.cel(r, col).Value2 = ser(v)
                else:
                    self.cel(r, col).Value = v
                self.q.ex.esperar()
        return self.idr(r)


def ser(d):
    return (d - dt.datetime(1899, 12, 30)).total_seconds() / 86400.0


def mesmo_inst(a, b):
    return hasattr(a, 'year') and hasattr(b, 'year') and abs((a - b).total_seconds()) <= 1


def estat(q):
    v = q.ws('Estatística').Range('A14:F133').Value
    out = {}
    for l in v:
        if l[0] and isinstance(l[1], (int, float)):
            k = f'{txt(l[0])}|{int(l[1])}'
            out.setdefault(k, (l[2], l[3], l[4]))
    return out


def oraculo(fin, an, nv):
    vals = []
    for r in fin:
        if txt(r.get('PARTICIPA_ESTATISTICA')).upper() != 'SIM':
            continue
        if txt(r.get('ANALITO')).upper() != an.upper() or nlote(r.get('LOTE')) != LOTE:
            continue
        if int(r.get('NIVEL') or 0) != nv or txt(r.get('EQUIPAMENTO')).upper() != EQUIP.upper():
            continue
        d = r.get('DATA_HORA') or r.get('DATA')
        if not hasattr(d, 'year') or not (J0 <= d <= J1):
            continue
        if isinstance(r.get('RESULTADO'), (int, float)):
            vals.append(float(r['RESULTADO']))
    return {'n': len(vals), 'media': stt.fmean(vals) if vals else None,
            'dp': stt.stdev(vals) if len(vals) > 1 else None, 'soma': sum(vals)}


def audit_rows(q, ids):
    v = q.ws('Audit_Log').UsedRange.Value
    out = []
    for l in v:
        s = ' | '.join(txt(x) for x in l)
        if any(i and i in s for i in ids):
            out.append(s[:300])
    return out


def abrir(caminho):
    q = QA(PROD, caminho)
    OUT['log'].append(f'EXCEL_PID {q.ex.pid}')
    return q


def fase_a(copia):
    q = abrir(copia)
    try:
        a = Aba(q)
        xl = a.xl
        fin = q.ler('tblCQ_Final')
        e0 = estat(q)
        an_est = {k.split('|')[0].upper() for k in e0}

        def elegivel(r, nv):
            d = r.get('DATA_HORA')
            return (txt(r.get('ORIGEM_RESULTADO')).upper().startswith('INTERF') and r.get('STATUS_ANALITICO') == 'ATIVO'
                    and txt(r.get('PARTICIPA_ESTATISTICA')).upper() == 'SIM' and nlote(r.get('LOTE')) == LOTE
                    and txt(r.get('EQUIPAMENTO')).upper() == EQUIP.upper() and int(r.get('NIVEL') or 0) == nv
                    and txt(r.get('ANALITO')).upper() in an_est and hasattr(d, 'year')
                    and J0 + dt.timedelta(days=3) <= d <= J1 and isinstance(r.get('RESULTADO'), (int, float)))
        c1 = sorted([r for r in fin if elegivel(r, 1)], key=lambda r: r['DATA_HORA'], reverse=True)
        ref = ref2 = None
        for r in c1:
            r2s = sorted([x for x in fin if elegivel(x, 2) and txt(x['ANALITO']) == txt(r['ANALITO'])],
                         key=lambda x: x['DATA_HORA'], reverse=True)
            if r2s:
                ref, ref2 = r, r2s[0]
                break
        if not ref:
            raise RuntimeError('sem interfaceamento elegivel no lote/equipamento/janela da Estatistica')
        AN = txt(ref['ANALITO'])
        dia = dt.datetime.combine(ref['DATA_HORA'].date() - dt.timedelta(days=2), dt.time())
        v0 = float(ref['RESULTADO'])
        base = {'EQUIPAMENTO': ref['EQUIPAMENTO'], 'MATRIZ': ref.get('MATRIZ'), 'LOTE': LOTE, 'NIVEL': 1,
                'ANALITO': AN, 'MOTIVO': cd.MOTIVOS_MANUAL[0]}
        st = {'analito': AN, 'dia': dia, 'ref': txt(ref['ID_REGISTRO']), 'ref2': txt(ref2['ID_REGISTRO']),
              'ref2_valor': ref2['RESULTADO'], 'ref2_dh': ref2['DATA_HORA']}
        a.ws.Unprotect(ii.SENHA)
        ids = [txt(x[0]) for x in a.lo.ListColumns('ID_REGISTRO').DataBodyRange.Value]
        r0 = max([k for k, v in enumerate(ids, 1) if v] or [0]) + 3
        if r0 + 40 > len(ids):
            xl.EnableEvents = False
            a.lo.Resize(a.lo.Range.Resize(a.lo.Range.Rows.Count + 60))
        st['r0'] = r0
        # linhas: M1 r0 (-> r0+1); vazia r0+4, P r0+5, Q r0+6; S r0+9 (-> r0+10); N r0+16; vizinhos r0+20..
        M = {'M1': (r0, 3 * 60 + 10, round(v0 * 0.997, 3), 'TESTE-AUDITORIA_U1', 3 * 60 + 40),
             'P': (r0 + 5, 3 * 60 + 20, round(v0 * 0.995, 3), 'TESTE-AUDITORIA_U2', 3 * 60 + 25),
             'Q': (r0 + 6, 3 * 60 + 30, round(v0 * 0.993, 3), 'TESTE-AUDITORIA_U3', 5 * 60)}
        for k, (r, h, v, dono, hc) in M.items():
            i = a.digitar(r, dict(base, DATA=dia, HORA=h / 1440.0, RESULTADO=v))
            a.eventos_off_set(r, 'USUARIO', dono)
            a.eventos_off_set(r, 'REGISTRADO_EM', dia + dt.timedelta(minutes=hc))
            st[k] = {'linha': r, 'id': i, 'valor': v, 'dono': dono, 'carimbo': a.car(r)}
        # S: substitui o ref2 (nivel 2) -- SUBSTITUI_ID digitado preenche a linha (ADR-082)
        rS = r0 + 9
        xl.EnableEvents = True
        a.cel(rS, 'SUBSTITUI_ID').Value = st['ref2']
        q.ex.esperar()
        preench = a.estado(rS)
        vS = round(float(ref2['RESULTADO']) * 1.012, 3)
        a.cel(rS, 'RESULTADO').Value = vS
        if not txt(a.cel(rS, 'MOTIVO').Value):
            a.cel(rS, 'MOTIVO').Value = base['MOTIVO']
        q.ex.esperar()
        a.eventos_off_set(rS, 'USUARIO', 'TESTE-AUDITORIA_U4')
        a.eventos_off_set(rS, 'REGISTRADO_EM', ing(ref2['DATA_HORA']) + dt.timedelta(minutes=2))
        st['S'] = {'linha': rS, 'id': a.idr(rS), 'valor': vS, 'dono': 'TESTE-AUDITORIA_U4', 'carimbo': a.car(rS),
                   'subst': a.sub(rS), 'preenchida_pelo_SUBSTITUI_ID': preench}
        log('fase A digitada', json.dumps(st, default=str)[:1500])
        st['estat0'] = {k: v for k, v in e0.items() if k.split('|')[0].upper() == AN.upper()}
        st['oraculo0'] = {nv: oraculo(fin, AN, nv) for nv in (1, 2)}
        msg, seg = q.atualizar()
        log('ATUALIZAR #1', round(seg), msg[-300:])
        fin1 = q.ler('tblCQ_Final')
        fb = {base_id(r['ID_REGISTRO']): r for r in fin1}
        st['integ1'] = {k: {c: fb.get(st[k]['id'], {}).get(c) for c in
                            ('USUARIO_MANUAL', 'STATUS_ANALITICO', 'RESULTADO', 'DATA_HORA', 'RECEBIDO_EM',
                             'MOTIVO_EXCLUSAO_AUTOMATICA')} for k in ('M1', 'P', 'Q', 'S')}
        r2 = fb.get(st['ref2'], {})
        st['ref2_1'] = {c: r2.get(c) for c in ('STATUS_ANALITICO', 'ID_RELACIONADO')}
        st['estat1'] = {k: v for k, v in estat(q).items() if k.split('|')[0].upper() == AN.upper()}
        st['oraculo1'] = {nv: oraculo(fin1, AN, nv) for nv in (1, 2)}
        st['estat1_full'] = estat(q)
        ok = all(txt(st['integ1'][k]['USUARIO_MANUAL']) == st[k]['dono'] and st['integ1'][k]['STATUS_ANALITICO'] == 'ATIVO'
                 for k in ('M1', 'P', 'Q', 'S')) and st['ref2_1']['STATUS_ANALITICO'] == 'SUBSTITUIDO_MANUAL'
        caso('A0 Pre-condicao: 4 manuais integrados ATIVO no nome dos donos; ref2 SUBSTITUIDO_MANUAL', ok,
             {'integ1': st['integ1'], 'ref2': st['ref2_1']})
        t, m0 = time.time(), os.path.getmtime(copia)
        xl.EnableEvents = False          # o BeforeSave do QC cancela e grava a propria (trancada); na copia, direto
        q.wb.Save()
        xl.EnableEvents = True
        if os.path.getmtime(copia) == m0:
            raise RuntimeError('a copia nao foi gravada')
        log('salvou a copia em', round(time.time() - t), 's')
        return st
    finally:
        q.fechar()


def fase_b(copia, st):
    q = abrir(copia)
    try:
        a = Aba(q)
        xl = a.xl
        xl.EnableEvents = True
        AN = st['analito']
        q.run('mSeguranca.ReprotectAll', teto=300)        # o estado depois do login: abas protegidas (UI only)
        prot0 = bool(a.ws.ProtectContents)
        log('Digitar Resultados protegida na fase B:', prot0)
        a.ws.Activate()
        a.cel(st['r0'] + 2, 'ANALITO').Select()                  # SelectionChange: 1o retrato (como o usuario)
        q.ex.esperar()
        protegido_no_corte = prot0

        def cortar_colar(rng, destino):
            rng.Select()
            xl.Selection.Cut()
            destino.Select()
            a.ws.Paste()
            q.ex.esperar()
            xl.CutCopyMode = False

        # K1
        m1 = st['M1']
        rD1 = m1['linha'] + 1
        try:
            cortar_colar(a.bloco(m1['linha']), a.cel(rD1, 'ID_REGISTRO'))
        except Exception as e:
            log('corte com a aba protegida falhou:', repr(e)[:200], '-> desprotege')
            a.ws.Unprotect(ii.SENHA)
            protegido_no_corte = False
            cortar_colar(a.bloco(m1['linha']), a.cel(rD1, 'ID_REGISTRO'))
        k1d, k1o = a.estado(rD1), a.estado(m1['linha'])
        ok1 = (k1d['id'] == m1['id'] and k1d['usuario'] == m1['dono'] and mesmo_inst(a.car(rD1)[1], m1['carimbo'][1])
               and k1o['id'] == '' and k1o['usuario'] == '' and k1o['registrado_em'] == 'None')
        st['K1'] = {'destino': k1d, 'origem': k1o, 'ok_tela': ok1}
        # K2
        p, qq = st['P'], st['Q']
        rV = p['linha'] - 1
        cortar_colar(a.bloco(p['linha'], qq['linha']), a.cel(rV, 'ID_REGISTRO'))
        k2 = [a.estado(r) for r in (rV, p['linha'], qq['linha'])]
        ok2 = (k2[0]['id'] == p['id'] and k2[0]['usuario'] == p['dono'] and mesmo_inst(a.car(rV)[1], p['carimbo'][1])
               and k2[1]['id'] == qq['id'] and k2[1]['usuario'] == qq['dono']
               and mesmo_inst(a.car(p['linha'])[1], qq['carimbo'][1])
               and k2[2]['id'] == '' and k2[2]['usuario'] == '')
        st['K2'] = {'linhas': k2, 'ok_tela': ok2}
        # K3 (Range.Cut com destino, eventos ligados)
        s = st['S']
        rSd = s['linha'] + 1
        a.bloco(s['linha']).Cut(a.bloco(rSd))
        q.ex.esperar()
        k3d, k3o = a.estado(rSd), a.estado(s['linha'])
        ok3 = (k3d['id'] == s['id'] and k3d['subst'] == s['subst'] and k3d['usuario'] == s['dono']
               and mesmo_inst(a.car(rSd)[1], s['carimbo'][1]) and k3o['id'] == '' and k3o['usuario'] == ''
               and not k3o['subst'])
        st['K3'] = {'destino': k3d, 'origem': k3o, 'ok_tela': ok3}
        # K4a: redigitar o ID de S (mesmo texto e man_N)
        antes = a.estado(rSd)
        a.cel(rSd, 'ID_REGISTRO').Value = s['id']
        q.ex.esperar()
        d1 = a.estado(rSd)
        dig = ''.join(ch for ch in s['id'] if ch.isdigit()).lstrip('0')
        a.cel(rSd, 'ID_REGISTRO').Value = 'man_' + dig
        q.ex.esperar()
        d2 = a.estado(rSd)
        ok4a = all(d['id'] == s['id'] and d['usuario'] == antes['usuario'] and d['registrado_em'] == antes['registrado_em']
                   and d['subst'] == antes['subst'] for d in (d1, d2))
        st['K4a'] = {'antes': antes, 'mesmo_texto': d1, 'man_N': d2, 'ok_tela': ok4a}
        # K4b: manual NAO integrado com carimbo 2020 -> redigitar o ID nao lava o carimbo
        if a.ws.ProtectContents:
            a.ws.Unprotect(ii.SENHA)
        rN = st['r0'] + 16
        base = {'EQUIPAMENTO': a.cel(rD1, 'EQUIPAMENTO').Value, 'MATRIZ': a.cel(rD1, 'MATRIZ').Value
                if 'MATRIZ' in a.cols else None, 'LOTE': LOTE, 'NIVEL': 1, 'ANALITO': AN, 'MOTIVO': cd.MOTIVOS_MANUAL[0]}
        dia = st['dia']
        idN = a.digitar(rN, dict(base, DATA=dia, HORA=(3 * 60 + 50) / 1440.0, RESULTADO=round(m1['valor'] * 1.004, 3)))
        a.eventos_off_set(rN, 'USUARIO', 'TESTE-AUDITORIA_U5')
        a.eventos_off_set(rN, 'REGISTRADO_EM', dt.datetime(2020, 1, 1, 6, 0, 0))
        q.run('mSeguranca.ReprotectAll', teto=300)
        log('protegida de novo para o resto da fase B:', bool(a.ws.ProtectContents))
        nb = a.estado(rN)
        a.cel(rN, 'ID_REGISTRO').Value = idN
        q.ex.esperar()
        a.cel(rN, 'ID_REGISTRO').Value = 'man_' + ''.join(ch for ch in idN if ch.isdigit()).lstrip('0')
        q.ex.esperar()
        na = a.estado(rN)
        st['K4b'] = {'id': idN, 'antes': nb, 'depois': na,
                     'ok_tela': na['usuario'] == 'TESTE-AUDITORIA_U5' and na['registrado_em'] == nb['registrado_em']
                     and na['id'] == idN}
        ids = [m1['id'], p['id'], qq['id'], s['id'], st['ref2']]
        al0 = audit_rows(q, ids)
        msg, seg = q.atualizar()
        log('ATUALIZAR #2', round(seg), msg[-300:])
        fin2 = q.ler('tblCQ_Final')
        fb = {base_id(r['ID_REGISTRO']): r for r in fin2}
        integ2 = {k: {c: fb.get(st[k]['id'], {}).get(c) for c in
                      ('USUARIO_MANUAL', 'STATUS_ANALITICO', 'RESULTADO', 'DATA_HORA', 'RECEBIDO_EM',
                       'MOTIVO_EXCLUSAO_AUTOMATICA')} for k in ('M1', 'P', 'Q', 'S')}
        fN = fb.get(idN, {})
        r2 = fb.get(st['ref2'], {})
        al1 = audit_rows(q, ids)
        al_novas = [x for x in al1 if x not in al0]
        i1 = st['integ1']

        def igual(k):
            return all(txt(integ2[k][c]) == txt(i1[k][c]) for c in ('USUARIO_MANUAL', 'STATUS_ANALITICO', 'RESULTADO'))
        caso('K1 Recortar M1 (integrado, dono U1) para linha vazia (Ctrl+X/Ctrl+V via Selection.Cut/Paste): destino com '
             'o carimbo do dono, origem sem carimbo; apos ATUALIZAR USUARIO_MANUAL e status iguais',
             st['K1']['ok_tela'] and igual('M1'), {'tela': st['K1'], 'antes': i1['M1'], 'depois': integ2['M1'],
                                                    'protegida_no_corte': protegido_no_corte})
        okq = all(integ2[k]['STATUS_ANALITICO'] == 'ATIVO'
                  and 'posterior' not in txt(integ2[k]['MOTIVO_EXCLUSAO_AUTOMATICA']).lower() for k in ('P', 'Q'))
        caso('K2 Bloco P(U2),Q(U3) sobe uma linha (recorte sobre linha ocupada): cada um com o proprio carimbo; '
             'nenhum recusado por "posterior" apos ATUALIZAR; USUARIO_MANUAL iguais',
             st['K2']['ok_tela'] and okq and igual('P') and igual('Q'),
             {'tela': st['K2'], 'antes': {k: i1[k] for k in 'PQ'}, 'depois': {k: integ2[k] for k in 'PQ'}})
        rel = base_id(r2.get('ID_RELACIONADO'))
        caso('K3 Recortar S (integrado, com SUBSTITUI_ID): SUBSTITUI_ID e carimbo vao junto; apos ATUALIZAR o '
             'interfaceamento continua SUBSTITUIDO_MANUAL por S',
             st['K3']['ok_tela'] and igual('S') and r2.get('STATUS_ANALITICO') == 'SUBSTITUIDO_MANUAL' and rel == s['id'],
             {'tela': st['K3'], 'S_depois': integ2['S'], 'ref2': {'id': st['ref2'], 'status': r2.get('STATUS_ANALITICO'),
                                                                  'id_relacionado': r2.get('ID_RELACIONADO')}})
        recN = txt(fN.get('MOTIVO_EXCLUSAO_AUTOMATICA'))
        caso('K4 Tocar o ID da propria linha (mesmo texto e man_N) nao muda carimbo nem SUBSTITUI_ID; manual nao '
             'integrado com carimbo 2020 continua recusado por "posterior"',
             st['K4a']['ok_tela'] and st['K4b']['ok_tela'] and fN.get('STATUS_ANALITICO') != 'ATIVO'
             and 'posterior' in recN.lower(),
             {'K4a': st['K4a'], 'K4b': st['K4b'], 'N_final': {'status': fN.get('STATUS_ANALITICO'), 'motivo': recN,
                                                              'usuario': fN.get('USUARIO_MANUAL')}})
        # K6 numeros
        e2full = estat(q)
        e2 = {k: v for k, v in e2full.items() if k.split('|')[0].upper() == AN.upper()}
        o2 = {nv: oraculo(fin2, AN, nv) for nv in (1, 2)}
        e1full = {k: tuple(v) for k, v in st['estat1_full'].items()}
        dif_full = [k for k in e1full if tuple(e2full.get(k, ())) != e1full[k]]
        o0, o1 = st['oraculo0'], st['oraculo1']
        k1v = sum(st[k]['valor'] for k in ('M1', 'P', 'Q'))
        delta_ok = (o1[1]['n'] == o0[1]['n'] + 3 and abs(o1[1]['soma'] - o0[1]['soma'] - k1v) < 1e-6
                    and o1[2]['n'] == o0[2]['n']
                    and abs(o1[2]['soma'] - o0[2]['soma'] - (s['valor'] - float(st['ref2_valor']))) < 1e-6)

        def bate(e, o):
            return (e is not None and o['n'] == int(e[0] or 0) and o['media'] is not None
                    and abs(float(e[1]) - o['media']) <= 1e-9 * abs(o['media']))
        keyn = {nv: next((k for k in e2 if k.endswith(f'|{nv}')), None) for nv in (1, 2)}
        bate1 = all(bate(st['estat1'].get(keyn[nv]), o1[nv]) and bate(e2.get(keyn[nv]), o2[nv]) for nv in (1, 2))
        caso('K6 Numeros: n/media do analito (N1 e N2) iguais antes e depois do recorte+ATUALIZAR; Estatistica = '
             'oraculo proprio (1e-9); delta do ATUALIZAR #1 = +3 valores (N1) e troca 1 por 1 (N2); nenhuma linha da '
             'Estatistica mudou com o recorte',
             delta_ok and bate1 and o1 == o2 and not dif_full,
             {'analito': AN, 'oraculo0': o0, 'oraculo1': o1, 'oraculo2': o2, 'estat0': st['estat0'],
              'estat1': st['estat1'], 'estat2': e2, 'linhas_estat_mudaram': dif_full[:10]})
        caso('K7 Audit_Log: nenhuma linha nova sobre os IDs movidos no ATUALIZAR #2 (USUARIO_MANUAL nao mudou)',
             not al_novas, {'novas': al_novas[:10], 'total_antes': len(al0)})
        q.wb.SaveCopyAs(os.path.join(W, f'snapB2_{ARQ}'))
        # ---------------------------------------------------------------- K5 vizinhos (sem ATUALIZAR)
        quem = None
        rX = st['r0'] + 20
        idX = a.digitar(rX, dict(base, DATA=dia, HORA=(4 * 60 + 5) / 1440.0, RESULTADO=round(m1['valor'] * 1.006, 3)))
        cx = a.car(rX)
        quem = cx[0]
        ok5a = (idX.startswith('MAN_') and quem and not quem.startswith('TESTE-AUDITORIA')
                and hasattr(cx[1], 'year') and abs((dt.datetime.now() - cx[1]).total_seconds()) < 900)
        # Delete ID..MOTIVO de S (com SUBSTITUI_ID)
        a.bloco(rSd).ClearContents()
        q.ex.esperar()
        dS = a.estado(rSd)
        ok5b = dS['id'] == '' and dS['usuario'] == '' and dS['registrado_em'] == 'None' and not dS['subst']
        # colar (Copy) a linha de P (ID repetido) numa linha vazia
        rC = st['r0'] + 24
        a.bloco(rV).Copy(a.bloco(rC))
        xl.CutCopyMode = False
        q.ex.esperar()
        dC = a.estado(rC)
        ok5c = (dC['id'].startswith('MAN_') and dC['id'] != p['id'] and dC['usuario'] == quem
                and a.estado(rV)['usuario'] == p['dono'])
        # Ctrl+D a partir de rX
        cx0 = a.estado(rX)
        a.ws.Range(a.cel(rX, 'ID_REGISTRO'), a.lo.DataBodyRange.Cells(rX + 2, a.c_mot)).FillDown()
        q.ex.esperar()
        dD = [a.estado(r) for r in (rX, rX + 1, rX + 2)]
        ok5d = (dD[0]['id'] == cx0['id'] and dD[0]['usuario'] == cx0['usuario']
                and dD[0]['registrado_em'] == cx0['registrado_em']
                and len({d['id'] for d in dD}) == 3 and all(d['id'].startswith('MAN_') for d in dD))
        # editar na propria linha (Q, agora na linha de P)
        e0q = a.estado(p['linha'])
        a.cel(p['linha'], 'RESULTADO').Value = round(qq['valor'] * 1.001, 3)
        q.ex.esperar()
        e1q = a.estado(p['linha'])
        ok5e = e1q['usuario'] == e0q['usuario'] == qq['dono'] and e1q['registrado_em'] == e0q['registrado_em'] \
            and e1q['id'] == qq['id']
        # SUBSTITUI_ID digitado numa linha vazia preenche a linha
        rY = st['r0'] + 28
        a.cel(rY, 'SUBSTITUI_ID').Value = st['ref']
        q.ex.esperar()
        dY = a.estado(rY)
        ok5f = (dY['subst'] == st['ref'] and txt(a.cel(rY, 'ANALITO').Value).upper() == AN.upper()
                and a.cel(rY, 'DATA').Value is not None and int(a.cel(rY, 'NIVEL').Value or 0) == 1)
        caso('K5 Vizinhos: manual novo (carimbo novo); Delete ID..MOTIVO limpa a linha e o SUBSTITUI_ID; colar linha com '
             'ID repetido (ID novo, carimbo de quem colou); Ctrl+D (origem mantem); editar na propria linha; '
             'SUBSTITUI_ID digitado preenche a linha',
             ok5a and ok5b and ok5c and ok5d and ok5e and ok5f,
             {'a_novo': [ok5a, idX, quem, str(cx[1])], 'b_delete': [ok5b, dS], 'c_colar': [ok5c, dC],
              'd_ctrlD': [ok5d, dD], 'e_editar': [ok5e, e0q, e1q], 'f_substitui': [ok5f, dY]})
    finally:
        q.fechar()


def main():
    raiz = os.environ['QC_RAIZ_PROD']
    prod = os.path.join(raiz, ARQ)
    h0 = sha256_arquivo(prod)
    OUT['sha_producao_antes'] = h0
    pasta = os.path.join(W, 'repro_' + PROD)
    os.makedirs(pasta, exist_ok=True)
    copia = os.path.join(pasta, ARQ)
    shutil.copy2(prod, copia)
    t0 = time.time()
    try:
        st = fase_a(copia)
        shutil.copy2(copia, os.path.join(W, f'snapA1_{ARQ}'))
        gravar()
        fase_b(copia, st)
    except Exception:
        OUT['erro'] = traceback.format_exc()
        log(OUT['erro'])
    OUT['segundos'] = round(time.time() - t0)
    OUT['sha_producao_depois'] = sha256_arquivo(prod)
    OUT['producao_intacta'] = OUT['sha_producao_depois'] == h0
    log('producao intacta', OUT['producao_intacta'], h0[:16])
    gravar()


if __name__ == '__main__':
    main()
