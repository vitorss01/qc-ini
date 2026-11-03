# -*- coding: utf-8 -*-
"""
conferencia_cruzada_hema.py -- segunda implementacao, minima e escrita do zero, para conferir a propria
calculadora independente (calc_independente.py) na Hematologia.

Autor: AGENTE_USUARIO_SIMULADO (F3/Hematologia, ciclo 2026-10-31_ciclo01, reexecucao). NAO importa
calc_independente.py: se as duas implementacoes concordarem entre si e com o QC_INI, o erro de uma delas
nao passa despercebido.

O QUE FAZ
    Le uma COPIA de QC_Hematologia.xlsm (openpyxl, read_only, data_only) e, para WBC, RBC, HCT, MCV e PLT
    (N1..N3), recalcula em Python puro:
      - n, media, DP (n-1), CV do lote 522611 (LOTE), equipamento XN1000 (EQ), ANO(DATA_HORA) = 2025 (ANO),
        PARTICIPA_ESTATISTICA = SIM, RESULTADO numerico  (src/hema/mEstatistica.EstatBasica);
      - bias do CEQ em duas etapas: CAP 2025, Uso_Analitico <> NAO, por (Ano|Rodada) media dos
        |(Resultado - Valor_Alvo)/Valor_Alvo*100|, depois media das medias  (src/comum/mCEQ.BiasEQ);
      - CV pooled de 6 meses: janela [EDATE(MU_Fim,-6)+1, MU_Fim], MU_Fim = maior DATA_HORA ativa; lotes
        com n >= 20; CVp = raiz(soma((n-1)CV^2)/soma(n-1))  (src/comum/mIncerteza.IncertezaCIQ "CV");
      - ET = 1,65*CV + |bias|; Sigma = (ETp - |bias|)/CVp, com o ETp lido da Estatistica K (o ETp e
        conferido de forma independente pela calc_independente.py, a partir da biblioteca).
    e imprime lado a lado com a Estatistica (C, F, G, H, K, L) e a diferenca relativa.

ARMADILHA CONHECIDA
    A aba EQA_Base tem, a direita (W:Y), uma tabela de/para com outro cabecalho "Provedor". Leia so A:U
    (as 21 primeiras colunas), senao o indice da coluna "Provedor" aponta para a tabela errada.

USO
    set PYTHONIOENCODING=utf-8
    python -I conferencia_cruzada_hema.py <copia.xlsm>
    Os parametros LOTE/EQ/ANO/ANALITOS abaixo correspondem aos filtros gravados no ciclo 2026-10-31_ciclo01;
    ajuste-os aos filtros da copia (Painel H3, Cfg_Integracao B11, Estatistica B3:D3).
"""
import calendar
import datetime as dt
import math
import statistics
import sys
from collections import defaultdict

import openpyxl

LOTE, EQ, ANO, PROVEDOR, ANO_EP, MESES = "522611", "XN1000", 2025, "CAP", 2025, 6
ANALITOS = ["WBC", "RBC", "HCT", "MCV", "PLT"]


def edate(d, m):
    y, mo = d.year, d.month + m
    while mo <= 0:
        mo += 12
        y -= 1
    while mo > 12:
        mo -= 12
        y += 1
    return dt.date(y, mo, min(d.day, calendar.monthrange(y, mo)[1]))


def rel(a, b):
    if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
        return "n/a"
    return abs(a - b) / max(abs(b), 1e-300)


def main(arq):
    wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)

    ws = wb["Principal - Resultados"]
    it = ws.iter_rows(min_row=6, values_only=True)          # cabecalho da tblCQ_Final na linha 6
    cab = [str(c).strip() if c is not None else "" for c in next(it)]
    ix = {c: i for i, c in enumerate(cab) if c}
    cq = []
    for r in it:
        if r is None or len(r) <= ix["PARTICIPA_ESTATISTICA"]:
            continue
        if str(r[ix["PARTICIPA_ESTATISTICA"]] or "").strip().upper() != "SIM":
            continue
        res, d = r[ix["RESULTADO"]], r[ix["DATA_HORA"]]
        if not isinstance(res, (int, float)) or isinstance(res, bool) or not isinstance(d, dt.datetime):
            continue
        cq.append((str(r[ix["ANALITO"]]).strip().upper(), int(float(r[ix["NIVEL"]])), str(r[ix["LOTE"]]).strip(),
                   str(r[ix["EQUIPAMENTO"]]).strip().upper(), float(res), d))
    print("CIQ ativos numericos:", len(cq))

    we = wb["EQA_Base"]
    it = we.iter_rows(min_row=1, values_only=True)
    ce = [str(c).strip() if c is not None else "" for c in next(it)]
    ie = {c: i for i, c in enumerate(ce[:21]) if c}          # so A:U (W:Y e o de/para)
    eqa = [r for r in it if r and any(v is not None for v in r[:21])]
    print("EQA_Base linhas:", len(eqa))

    mu_fim = max(x[5] for x in cq).date()
    ini = edate(mu_fim, -MESES) + dt.timedelta(days=1)
    print("janela do Sigma:", ini, "a", mu_fim)

    def bias2(an):
        rod_abs, rod_sig = defaultdict(list), defaultdict(list)
        for r in eqa:
            if str(r[ie["Analito_Canonico"]] or "").strip().upper() != an:
                continue
            if str(r[ie["Uso_Analitico"]] or "").strip().upper() == "NAO":
                continue
            if str(r[ie["Provedor"]] or "").strip().upper() != PROVEDOR:
                continue
            if int(float(r[ie["Ano"]] or 0)) != ANO_EP:
                continue
            x, a = r[ie["Resultado"]], r[ie["Valor_Alvo"]]
            if not isinstance(x, (int, float)) or not isinstance(a, (int, float)) or a == 0:
                continue
            b = (x - a) / a * 100
            chave = f"{r[ie['Ano']]}|{r[ie['Rodada']]}"
            rod_abs[chave].append(abs(b))
            rod_sig[chave].append(b)
        if not rod_abs:
            return None, None
        return (sum(statistics.fmean(v) for v in rod_abs.values()) / len(rod_abs),
                sum(statistics.fmean(v) for v in rod_sig.values()) / len(rod_sig))

    est = list(wb["Estatística"].iter_rows(min_row=14, max_row=133, max_col=12, values_only=True))
    pior = 0.0
    for an in ANALITOS:
        b_abs, _ = bias2(an)
        for nv in (1, 2, 3):
            v = [x[4] for x in cq if x[0] == an and x[1] == nv and x[2] == LOTE and x[3] == EQ and x[5].year == ANO]
            n = len(v)
            m = statistics.fmean(v)
            cv = statistics.stdev(v) / m * 100
            por = defaultdict(list)
            for x in cq:
                if x[0] == an and x[1] == nv and x[3] == EQ and ini <= x[5].date() <= mu_fim:
                    por[x[2]].append(x[4])
            num = den = 0.0
            for vv in por.values():
                if len(vv) >= 20 and statistics.fmean(vv) != 0:
                    c = statistics.stdev(vv) / statistics.fmean(vv) * 100
                    num += (len(vv) - 1) * c * c
                    den += len(vv) - 1
            cvp = math.sqrt(num / den) if den else None
            linha = next((i for i, row in enumerate(est) if str(row[0]).strip().upper() == an
                          and str(row[1]).strip() in (str(nv), f"{nv}.0")), None)
            row = est[linha]
            etp = row[10]
            et = 1.65 * cv + b_abs if b_abs is not None else None
            sig = (etp - b_abs) / cvp if isinstance(etp, (int, float)) and b_abs is not None and cvp else "-"
            difs = [rel(cv, row[5]), rel(b_abs, row[6]), rel(et, row[7]), rel(sig, row[11])]
            pior = max([pior] + [d for d in difs if isinstance(d, float)])
            print(f"{an} N{nv} (Estatistica linha {linha + 14}): n {n}/{row[2]} | CV {cv!r}/{row[5]!r} "
                  f"| bias {b_abs!r}/{row[6]!r} | ET {et!r}/{row[7]!r} | ETp {etp!r} | CVp {cvp!r} "
                  f"| Sigma {sig!r}/{row[11]!r} | dif.rel {['%.1e' % d if isinstance(d, float) else d for d in difs]}")
    print("MAIOR DIFERENCA RELATIVA:", "%.2e" % pior, "(tolerancia 1e-9)", "OK" if pior <= 1e-9 else "DIVERGENTE")


if __name__ == "__main__":
    main(sys.argv[1])
