# -*- coding: utf-8 -*-
"""
segunda_opiniao_hema.py -- terceira implementacao, escrita do zero, para conferir a Estatistica da Hematologia
(e, por tabela, a calc_independente.py e a conferencia_cruzada_hema.py).

Autor: AGENTE_USUARIO_SIMULADO (F3/Hematologia, ciclo 2026-10-31_ciclo01, terceira passada de 01/11/2026).
NAO importa calc_independente.py nem conferencia_cruzada_hema.py. Diferente da conferencia_cruzada_hema.py, que
le o ETp da propria Estatistica, esta calcula o ETp a partir da BIBLIOTECA anual (Especificacoes, bloco do ano de
referencia) e usa erfc no DPM (mais exato que 1 - Phi quando o Sigma e alto).

O QUE FAZ
    Le uma COPIA de QC_Hematologia.xlsm (openpyxl, read_only, data_only) e, para os analitos de ANALITOS (N1..N3),
    recalcula em Python puro (math/statistics):
      C:F  n, media, DP (n-1), CV do LOTE, equipamento EQ, ANO(DATA_HORA) entre ANO_DE e ANO_ATE,
           PARTICIPA_ESTATISTICA = SIM, RESULTADO numerico           (src/hema/mEstatistica.EstatBasica)
      G/T  bias do CEQ em duas etapas: PROVEDOR, ANO_EP exato (ADR-072), Uso_Analitico <> NAO; por (Ano|Rodada)
           media dos |(Resultado - Valor_Alvo)/Valor_Alvo*100| (e do bias com sinal), depois media das medias;
           AC = n amostras, AD = n rodadas                            (src/comum/mCEQ.BiasEQ)
      K/J/I ETp e CVTp pela fonte cadastrada (Especificacoes, coluna "ESPQ FONTE") e pelo desempenho VB
           (coluna "Desemp."): CLIA -> "CLIA ETp%" da biblioteca do ano CLIA; CVTp = ETp/3.
           VB -> 1,65*f*CVi + f/2*raiz(CVi^2 + CVg^2), f = 0,25/0,50/0,75 (OTI/DES/MIN), CVi/CVg da biblioteca do
           ano VB; CVTp = f*CVi. FAB -> "FAB ETp%" / "FAB CVTp%". Vazio ou <= 0 -> "-"      (ADR-068/069/075)
      H    ET = 1,65*CV + |bias| (vazio sem bias)
      L    Sigma = (ETp - |bias|)/CVp, CVp de MESES_SIGMA meses: janela [EDATE(MU_Fim,-m)+1, MU_Fim], MU_Fim =
           maior DATA_HORA ativa (dia), mesmo equipamento, lotes com n >= 20 e media <> 0,
           CVp = raiz(soma((n_i-1)*CV_i^2)/soma(n_i-1))                (src/comum/mIncerteza.IncertezaCIQ "CV")
      N/O  margem = ETp - ET; % = margem/ETp*100
      V    DPM = 0,5*erfc((Sigma-1,5)/raiz(2))*1e6
      AF   u(Rw) = o mesmo CVp numa janela de 12 meses (comparado so quando a celula tem numero)
    e compara com a Estatistica (Value2 gravado). Tolerancia: 1e-9 relativo; DPM com Sigma > 6 tambem aceita
    1e-9 DPM absoluto (o "1 - Phi" do Excel perde digitos relativos).

ARMADILHAS
    - A EQA_Base tem, em W:Y, uma tabela de/para com outro cabecalho "Provedor": leia so A:U.
    - Com python -I o PYTHONIOENCODING e ignorado: o script reconfigura o stdout para UTF-8.

USO
    python -I segunda_opiniao_hema.py <copia.xlsm>
    Os parametros abaixo correspondem aos filtros gravados na producao no ciclo 2026-10-31_ciclo01 (Painel H3 =
    522611, Cfg_Integracao B11 = XN1000, Estatistica B3:D3 = 2025, R4 = CAP, N4 = 2025, L11 = 6, bloco do lote na
    aba Analitos = CLIA 2025 / VB 2025). Ajuste-os aos filtros da copia que for auditar.
"""
import calendar
import datetime as dt
import math
import statistics
import sys
from collections import defaultdict

import openpyxl

LOTE, EQ, ANO_DE, ANO_ATE = "522611", "XN1000", 2025, 2025
PROVEDOR, ANO_EP, MESES_SIGMA = "CAP", 2025, 6
ANO_CLIA, ANO_VB = 2025, 2025
ANALITOS = ["HGB", "MCH", "MCHC", "NEUT%", "MONO%", "LYMPH#", "EO#", "MPV", "RDW-CV"]
F_VB = {"ÓTI": 0.25, "OTI": 0.25, "DES": 0.50, "MIN": 0.75}
TOL = 1e-9

# Estatistica: coluna (letra) -> indice 0-based na linha lida com max_col=32
COLS = {"C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9, "K": 10, "L": 11, "N": 13, "O": 14,
        "T": 19, "V": 21, "AC": 28, "AD": 29, "AF": 31}


def edate(d, m):
    y, mo = d.year, d.month + m
    while mo <= 0:
        mo += 12
        y -= 1
    while mo > 12:
        mo -= 12
        y += 1
    return dt.date(y, mo, min(d.day, calendar.monthrange(y, mo)[1]))


def ehnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def comparar(ind, qc, sigma=None):
    """(ok, dif) -- numeros por diferenca relativa; textos/vazios por igualdade."""
    if ind is None and (qc is None or qc == ""):
        return True, 0.0
    if ehnum(ind) and ehnum(qc):
        d = abs(ind - qc) / max(abs(qc), 1e-300)
        ok = d <= TOL or (sigma is not None and sigma > 6 and abs(ind - qc) <= 1e-9)
        return ok, d
    return str(ind).strip() == str(qc).strip(), 0.0


def main(arq):
    sys.stdout.reconfigure(encoding="utf-8")
    wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)

    # ---- CIQ (tblCQ_Final; cabecalho na linha 6) ----
    it = wb["Principal - Resultados"].iter_rows(min_row=6, values_only=True)
    cab = [str(c).strip() if c is not None else "" for c in next(it)]
    ix = {c: i for i, c in enumerate(cab) if c}
    cq = []
    for r in it:
        if not r or len(r) <= ix["PARTICIPA_ESTATISTICA"]:
            continue
        if str(r[ix["PARTICIPA_ESTATISTICA"]] or "").strip().upper() != "SIM":
            continue
        x, d = r[ix["RESULTADO"]], r[ix["DATA_HORA"]]
        if not ehnum(x) or not isinstance(d, dt.datetime):
            continue
        cq.append({"an": str(r[ix["ANALITO"]]).strip().upper(), "nv": int(float(r[ix["NIVEL"]])),
                   "lote": str(r[ix["LOTE"]]).strip(), "eq": str(r[ix["EQUIPAMENTO"]]).strip().upper(),
                   "x": float(x), "d": d})
    mu_fim = max(c["d"] for c in cq).date()

    # ---- CEQ (EQA_Base, so A:U) ----
    it = wb["EQA_Base"].iter_rows(min_row=1, values_only=True)
    ce = [str(c).strip() if c is not None else "" for c in next(it)][:21]
    ie = {c: i for i, c in enumerate(ce) if c}
    eqa = [r[:21] for r in it if r and any(v is not None for v in r[:21])]

    # ---- Especificacoes: cadastro (fonte, desempenho) e biblioteca anual ----
    esp = list(wb["Especificações"].iter_rows(min_row=1, max_row=120, max_col=40, values_only=True))
    cab_esp = [" ".join(str(c).split()) if c is not None else "" for c in esp[2]]  # linha 3 = cabecalho (espacos normalizados)
    i_fonte = cab_esp.index("ESPQ FONTE")
    i_desemp = cab_esp.index("Desemp.")                                            # 1a ocorrencia (coluna Q)
    cad = {str(r[0]).strip().upper(): r for r in esp[3:43] if r[0]}

    def biblioteca(ano):
        for k, r in enumerate(esp):
            if str(r[0] or "").strip().upper() == "ANO" and r[1] == ano:
                nomes = [str(c).strip().upper() if c else "" for c in esp[k + 1]]
                tab = {}
                for lin in esp[k + 2:k + 8]:
                    rot = str(lin[0] or "").strip()
                    tab[rot] = {nomes[j]: lin[j] for j in range(1, len(nomes)) if nomes[j]}
                return tab
        raise SystemExit(f"biblioteca {ano} nao encontrada")

    bib_clia, bib_vb = biblioteca(ANO_CLIA), biblioteca(ANO_VB)

    def etp_cvtp(an):
        fonte = str(cad[an][i_fonte] or "").strip().upper()
        if fonte == "CLIA":
            e = bib_clia["CLIA ETp%"].get(an)
            return fonte, (float(e) if ehnum(e) and e > 0 else None), (e / 3 if ehnum(e) and e > 0 else None)
        if fonte == "VB":
            f = F_VB[str(cad[an][i_desemp]).strip().upper()]
            cvi, cvg = bib_vb["VB-CVI"].get(an), bib_vb["VB-CVG"].get(an)
            if not (ehnum(cvi) and ehnum(cvg)):
                return fonte, None, None
            e = 1.65 * f * cvi + f / 2 * math.sqrt(cvi ** 2 + cvg ** 2)
            return fonte, (e if e > 0 else None), f * cvi
        if fonte == "FAB":
            e, c = bib_clia["FAB ETp%"].get(an), bib_clia["FAB CVTp%"].get(an)
            return fonte, (float(e) if ehnum(e) and e > 0 else None), c
        return fonte, None, None

    def bias(an):
        abs_r, sig_r = defaultdict(list), defaultdict(list)
        for r in eqa:
            if str(r[ie["Analito_Canonico"]] or "").strip().upper() != an:
                continue
            if str(r[ie["Uso_Analitico"]] or "").strip().upper() == "NAO":
                continue
            if str(r[ie["Provedor"]] or "").strip().upper() != PROVEDOR:
                continue
            if not ehnum(r[ie["Ano"]]) or int(r[ie["Ano"]]) != ANO_EP:
                continue
            x, a = r[ie["Resultado"]], r[ie["Valor_Alvo"]]
            if not (ehnum(x) and ehnum(a)) or a == 0:
                continue
            b = (x - a) / a * 100.0
            k = f"{r[ie['Ano']]}|{r[ie['Rodada']]}"
            abs_r[k].append(abs(b))
            sig_r[k].append(b)
        if not abs_r:
            return None, None, None, None
        n_am = sum(len(v) for v in abs_r.values())
        return (statistics.fmean([statistics.fmean(v) for v in abs_r.values()]),
                statistics.fmean([statistics.fmean(v) for v in sig_r.values()]), n_am, len(abs_r))

    def cvp(an, nv, meses):
        ini = edate(mu_fim, -meses) + dt.timedelta(days=1)
        por = defaultdict(list)
        for c in cq:
            if c["an"] == an and c["nv"] == nv and c["eq"] == EQ and ini <= c["d"].date() <= mu_fim:
                por[c["lote"]].append(c["x"])
        num = den = 0.0
        for v in por.values():
            m = statistics.fmean(v)
            if len(v) >= 20 and m != 0:
                num += (len(v) - 1) * (statistics.stdev(v) / m * 100) ** 2
                den += len(v) - 1
        return (math.sqrt(num / den) if den else None), int(den), ini

    est = list(wb["Estatística"].iter_rows(min_row=14, max_row=133, max_col=32, values_only=True))
    linha_de = {(str(r[0]).strip().upper(), int(float(r[1]))): i for i, r in enumerate(est)
                if r[0] and ehnum(r[1])}

    print(f"MU_Fim = {mu_fim}; lote {LOTE}; {EQ}; ano {ANO_DE}-{ANO_ATE}; {PROVEDOR} {ANO_EP}; "
          f"CLIA {ANO_CLIA} / VB {ANO_VB}")
    tot = ok_tot = 0
    pior = 0.0
    for an in ANALITOS:
        b_abs, b_sig, n_am, n_rod = bias(an)
        fonte, etp, cvtp = etp_cvtp(an)
        for nv in (1, 2, 3):
            v = [c["x"] for c in cq if c["an"] == an and c["nv"] == nv and c["lote"] == LOTE and c["eq"] == EQ
                 and ANO_DE <= c["d"].year <= ANO_ATE]
            n = len(v)
            m = statistics.fmean(v) if n else None
            dp = statistics.stdev(v) if n >= 2 else None
            cv = dp / m * 100 if dp is not None and m else None
            cvp6, gl6, ini6 = cvp(an, nv, MESES_SIGMA)
            cvp12, gl12, _ = cvp(an, nv, 12)
            et = 1.65 * cv + b_abs if cv is not None and b_abs is not None else None
            if etp is None:
                sig = margem = pct = dpm = "-"
                cvtp_s = "-"
            else:
                cvtp_s = cvtp
                sig = (etp - b_abs) / cvp6 if b_abs is not None and cvp6 else None
                margem = etp - et if et is not None else None
                pct = margem / etp * 100 if margem is not None else None
                dpm = 0.5 * math.erfc((sig - 1.5) / math.sqrt(2)) * 1e6 if ehnum(sig) else None
            ind = {"C": n, "D": m, "E": dp, "F": cv,
                   "G": b_abs if b_abs is not None else "SEM EP", "T": b_sig if b_sig is not None else "SEM EP",
                   "AC": n_am if n_am is not None else "SEM EP", "AD": n_rod if n_rod is not None else "SEM EP",
                   "H": et, "I": fonte, "J": cvtp_s, "K": etp if etp is not None else "-", "L": sig,
                   "N": margem, "O": pct, "V": dpm, "AF": cvp12}
            row = est[linha_de[(an, nv)]]
            difs = []
            for col, val in ind.items():
                qc = row[COLS[col]]
                if col == "AF" and not ehnum(qc):
                    continue                       # u(Rw) vazio = validade insuficiente: nao comparado
                if col == "C" and ehnum(qc):
                    qc = int(qc)
                ok, d = comparar(val, qc, sig if ehnum(sig) else None)
                tot += 1
                ok_tot += ok
                pior = max(pior, d if not (col == "V" and ehnum(sig) and sig > 6) else 0.0)
                difs.append(f"{col}{'' if ok else '!'}")
                if not ok:
                    print(f"   DIVERGE {an} N{nv} {col}: independente {val!r} x QC_INI {qc!r} (dif {d:.2e})")
            print(f"{an} N{nv} (linha {linha_de[(an, nv)] + 14}): n={n} CV={cv!r} |bias|={b_abs!r} "
                  f"ET={et!r} fonte={fonte} ETp={etp!r} CVp6={cvp6!r} (gl {gl6}) Sigma={sig!r} DPM={dpm!r} "
                  f"CVp12={cvp12!r} | QC_INI F={row[5]!r} G={row[6]!r} H={row[7]!r} K={row[10]!r} "
                  f"L={row[11]!r} V={row[21]!r} AF={row[31]!r}")
    print(f"COMPARACOES: {ok_tot}/{tot} iguais; maior dif. relativa (fora V com Sigma > 6) = {pior:.2e}")
    print(f"janela do Sigma: {ini6} a {mu_fim}")


if __name__ == "__main__":
    main(sys.argv[1])
