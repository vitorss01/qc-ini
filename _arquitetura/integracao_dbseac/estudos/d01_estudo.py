# -*- coding: utf-8 -*-
"""
D-01 -- Estudo comparativo (SOMENTE LEITURA) dos estimadores de vies do CEQ
usados no Sigma e no Erro Total do QC_INI.

Le COPIAS dos .xlsm (openpyxl, data_only, read_only -- sem Excel, nada e gravado
nas pastas). Para cada analito x nivel da aba Estatistica compara:

    A  = Mean(|bias|)  em duas etapas por rodada  (BiasEQ "ABS" -- o de hoje)
    B  = |Mean(bias)|  em duas etapas por rodada  (BiasEQ "SIGNED", em modulo)
         + media com sinal por amostra, DP, EP e IC95% (amostra e rodada)
    C  = RMS = raiz(media dos bias^2)              (Nordtest / Ercan 2022)
    D  = regressao OLS  Xlab = a + b.Xref  por analito; vies% no nivel do
         controle = (a + (b-1).Xc)/Xc.100, com IC95% (so se interpolado)
    E  = bootstrap estratificado por rodada (estabilidade de A e de B)
    F  = (informativo) B se o IC95% (amostras) exclui zero, senao 0
         -- a leitura "vies nao significativo pode ser desprezado"
    G  = (acrescentado pelo gestor) RMS das medias de rodada = raiz(media das
         medias de rodada^2): guarda o vies que muda entre rodadas (deriva de
         calibracao / lote) e dilui o erro aleatorio de cada amostra por raiz(m);
         + ANOVA de 1 via entre rodadas (heterogeneidade do vies no tempo)

Sigma = (TEa - |vies|)/CV, com o MESMO TEa (coluna ETp %) e o MESMO CV
(coluna CV %) que a planilha usa hoje; classe, regras de Westgard, N, run size
e frequencia saem da tblPlanoQC_Sigma (aba Cfg_PlanoQC) da propria pasta.

Duas selecoes de amostras:
    S1 = exatamente a do BiasEQ (mCEQ.bas): canonico, Uso_Analitico <> "NAO",
         provedor, rodada, ano vigente (maior ano <= Ano EP), Bias_Abs
         "numerico" no sentido do VBA (celula vazia conta como 0).
    S2 = a do ViesEQ / ADR-064: S1 + avaliada pelo provedor + resultado e
         alvo numericos e <> 0 + bias nao vazio + chave unica.

Uso:
    python d01_estudo.py --bio <copia QC_Bioquimica.xlsm> \
                         --hema <copia QC_Hematologia.xlsm> \
                         [--saida <pasta dos CSV>] [--boot 2000] [--seed 20261006]

Saidas: d01_resultados_bioquimica.csv, d01_resultados_hematologia.csv e um
resumo no stdout (validacao, matrizes de mudanca de classe, casos criticos).
Nenhum caminho de rede, senha ou nome de arquivo-fonte do CEQ e gravado.
"""
import argparse
import csv
import math
import os
import sys
from collections import OrderedDict, defaultdict

import numpy as np
import openpyxl
from scipy import stats

ABA_EST = "Estatística"
ABA_EQA = "EQA_Base"
ABA_PLANO = "Cfg_PlanoQC"
ABA_AN = "Analitos"

# colunas da EQA_Base (mEQA.bas, base 1)
B_PROV, B_ANO, B_ROD, B_ANALITO, B_CANON, B_AMOSTRA = 1, 2, 3, 4, 5, 6
B_XLAB, B_XREF, B_SD, B_SDI = 7, 8, 9, 10
B_STATUS, B_UNID, B_BIAS, B_BIASABS, B_USO, B_CHAVE = 14, 15, 16, 17, 20, 21

ORDEM_CLASSE = ["Desempenho inadequado", "Marginal", "Bom", "Excelente", "Classe mundial"]
EST = ["A", "B", "C", "D", "F", "G"]


# ---------------------------------------------------------------- utilidades
def vba_isnumeric(v):
    """IsNumeric do VBA: Empty (celula vazia) e numerico (vale 0)."""
    if v is None:
        return True
    if isinstance(v, bool):
        return True
    if isinstance(v, (int, float)):
        return not (isinstance(v, float) and math.isnan(v))
    s = str(v).strip()
    if s == "":
        return False
    try:
        float(s.replace(",", "."))
        return True
    except ValueError:
        return False


def vba_num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).strip().replace(",", "."))


def num_real(v):
    """Numero de verdade (nao vazio)."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip()
    if not s:
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def up(v):
    return ("" if v is None else str(v)).strip().upper()


def livre(f):
    return up(f) in ("", "TODOS", "TODAS")


def nome_definido(wb, nome):
    dn = wb.defined_names.get(nome)
    if dn is None:
        return None
    for aba, ref in dn.destinations:
        ref = ref.replace("$", "")
        return wb[aba][ref].value
    return None


def tcrit(df):
    return float(stats.t.ppf(0.975, df)) if df >= 1 else float("nan")


# ------------------------------------------------------------- plano de CQ
def ler_plano(wb):
    ws = wb[ABA_PLANO]
    faixas = []
    for row in ws.iter_rows(min_row=3, max_row=20, max_col=8, values_only=True):
        smin, smax = num_real(row[0]), num_real(row[1])
        if smin is None or smax is None:
            continue
        faixas.append(dict(smin=smin, smax=smax, classe=row[2], regras=row[3] or "",
                           n=row[4] if row[4] is not None else "",
                           runsize=row[5] if row[5] is not None else "",
                           freq=row[6] or ""))
    return faixas


def plano(faixas, s):
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return None
    for f in faixas:
        if f["smin"] <= s < f["smax"]:
            return f
    return None


def idx_classe(c):
    return ORDEM_CLASSE.index(c) if c in ORDEM_CLASSE else None


# ------------------------------------------------------------ leitura base
def ler_eqa(wb):
    ws = wb[ABA_EQA]
    linhas = []
    for row in ws.iter_rows(min_row=2, max_row=5001, max_col=21, values_only=True):
        r = (None,) + tuple(row)          # base 1, como no VBA
        if up(r[B_CANON]):
            linhas.append(r)
    return linhas


def casa(r, analito, prov, rod):
    if up(r[B_CANON]) != up(analito):
        return False
    if up(r[B_USO]) == "NAO":
        return False
    if not livre(prov) and up(r[B_PROV]) != up(prov):
        return False
    if not livre(rod) and up(r[B_ROD]) != up(rod):
        return False
    return True


def ano_vigente(linhas, analito, ano_ref, prov, rod, col):
    aref = 32767
    if num_real(ano_ref) is not None:
        aref = int(num_real(ano_ref))
    melhor = None
    for r in linhas:
        if not casa(r, analito, prov, rod):
            continue
        if not vba_isnumeric(r[B_ANO]) or not vba_isnumeric(r[col]):
            continue
        a = int(vba_num(r[B_ANO]))
        if a <= aref and (melhor is None or a > melhor):
            melhor = a
    return melhor


def selecionar(linhas, analito, ano_ref, prov, rod):
    """Devolve (ano, S1, S2, n_vazio_S1). Cada amostra: dict."""
    anov = ano_vigente(linhas, analito, ano_ref, prov, rod, B_BIASABS)
    if anov is None:
        return None, [], [], 0
    s1, s2, vistos, n_vazio = [], [], set(), 0
    for r in linhas:
        if not casa(r, analito, prov, rod):
            continue
        if not vba_isnumeric(r[B_ANO]) or int(vba_num(r[B_ANO])) != anov:
            continue
        if not vba_isnumeric(r[B_BIASABS]):
            continue
        b_abs = vba_num(r[B_BIASABS])
        b_sig = vba_num(r[B_BIAS]) if vba_isnumeric(r[B_BIAS]) else 0.0
        if r[B_BIAS] is None or r[B_BIASABS] is None:
            n_vazio += 1
        am = dict(rod=up(r[B_ROD]), b=b_sig, babs=b_abs, xlab=num_real(r[B_XLAB]),
                  xref=num_real(r[B_XREF]), status=up(r[B_STATUS]), chave=up(r[B_CHAVE]),
                  amostra=r[B_AMOSTRA])
        s1.append(am)
        # S2 -- regra do ViesEQ (ADR-064)
        if am["status"] == "NAO AVALIADO":
            continue
        if num_real(r[B_BIAS]) is None or am["xlab"] is None or am["xref"] is None:
            continue
        if am["xlab"] == 0 or am["xref"] == 0:
            continue
        if am["chave"]:
            if am["chave"] in vistos:
                continue
            vistos.add(am["chave"])
        s2.append(am)
    return anov, s1, s2, n_vazio


# ------------------------------------------------------------- estimadores
def por_rodada(amostras):
    g = OrderedDict()
    for a in amostras:
        g.setdefault(a["rod"], []).append(a)
    return g


def est_A(amostras):
    g = por_rodada(amostras)
    return float(np.mean([np.mean([a["babs"] for a in v]) for v in g.values()])) if g else None


def media_signed_2etapas(amostras):
    g = por_rodada(amostras)
    return float(np.mean([np.mean([a["b"] for a in v]) for v in g.values()])) if g else None


def resumo_vies(amostras):
    """Todos os estimadores de vies agregado (independem do nivel)."""
    out = {}
    n = len(amostras)
    g = por_rodada(amostras)
    k = len(g)
    out["n"], out["k"] = n, k
    if n == 0:
        return out
    b = np.array([a["b"] for a in amostras], float)
    out["A"] = est_A(amostras)
    out["Bsig2"] = media_signed_2etapas(amostras)          # com sinal, 2 etapas
    out["B"] = abs(out["Bsig2"])
    out["Bsig1"] = float(b.mean())                          # com sinal, por amostra
    out["C"] = float(math.sqrt((b ** 2).mean()))
    out["mediana"] = float(np.median(b))
    if n >= 2:
        dp = float(b.std(ddof=1))
        ep = dp / math.sqrt(n)
        t = tcrit(n - 1)
        out.update(DP=dp, EP=ep, IC_lo=out["Bsig1"] - t * ep, IC_hi=out["Bsig1"] + t * ep)
        out["sig_am"] = not (out["IC_lo"] <= 0 <= out["IC_hi"])
    mr = np.array([np.mean([a["b"] for a in v]) for v in g.values()], float)
    out["medias_rodada"] = [float(x) for x in mr]
    if k >= 2:
        dpr = float(mr.std(ddof=1))
        epr = dpr / math.sqrt(k)
        t = tcrit(k - 1)
        out.update(EProd=epr, ICr_lo=out["Bsig2"] - t * epr, ICr_hi=out["Bsig2"] + t * epr)
        out["sig_rod"] = not (out["ICr_lo"] <= 0 <= out["ICr_hi"])
        # componentes: DP dentro da rodada (agrupado) e DP entre medias de rodada
        ss_w = sum(((np.array([a["b"] for a in v]) - np.mean([a["b"] for a in v])) ** 2).sum()
                   for v in g.values())
        gl_w = n - k
        out["DPw"] = float(math.sqrt(ss_w / gl_w)) if gl_w > 0 else None
        out["DPentre_medias"] = dpr
        out["inversao_sinal"] = bool((mr > 0).any() and (mr < 0).any())
        # G -- RMS das medias de rodada: mantem o vies que muda de rodada para
        # rodada e dilui o erro aleatorio de cada amostra (divide por raiz(m)).
        out["G"] = float(math.sqrt((mr ** 2).mean()))
        # heterogeneidade entre rodadas (ANOVA de 1 via sobre o bias com sinal)
        grupos = [np.array([a["b"] for a in v]) for v in g.values() if len(v) >= 2]
        if len(grupos) >= 2 and gl_w > 0 and ss_w > 0:
            fstat, pval = stats.f_oneway(*grupos)
            out["anova_F"], out["anova_p"] = float(fstat), float(pval)
            out["heterog"] = bool(pval < 0.05)
    # F -- "se nao significativo, desprezar"
    out["F"] = out["B"] if out.get("sig_am") else 0.0
    # outliers robustos (|z| > 3,5 pela MAD)
    mad = float(np.median(np.abs(b - np.median(b)))) * 1.4826
    out["n_outliers"] = int((np.abs(b - np.median(b)) > 3.5 * mad).sum()) if mad > 0 else 0
    return out


def regressao(amostras):
    """OLS Xlab = a + b.Xref (Xref = media do grupo, erro pequeno frente ao Xlab)."""
    pts = [(a["xref"], a["xlab"]) for a in amostras
           if a["xref"] is not None and a["xlab"] is not None]
    r = {"reg_n": len(pts)}
    if len(pts) < 8:
        r["reg_motivo"] = "menos de 8 amostras com Xlab e Xref"
        return r
    x = np.array([p[0] for p in pts], float)
    y = np.array([p[1] for p in pts], float)
    nconc = len(set(np.round(x, 6)))
    r["reg_nconc"] = nconc
    r["reg_xmin"], r["reg_xmax"] = float(x.min()), float(x.max())
    if nconc < 3 or x.min() <= 0 or x.max() / x.min() < 1.5:
        r["reg_motivo"] = "faixa de concentracao insuficiente (<3 niveis ou max/min < 1,5)"
        return r
    n = len(x)
    xm = x.mean()
    sxx = ((x - xm) ** 2).sum()
    b = ((x - xm) * (y - y.mean())).sum() / sxx
    a = y.mean() - b * xm
    res = y - (a + b * x)
    s2 = (res ** 2).sum() / (n - 2)
    se_b = math.sqrt(s2 / sxx)
    se_a = math.sqrt(s2 * (1 / n + xm ** 2 / sxx))
    t = tcrit(n - 2)
    r.update(reg_a=float(a), reg_b=float(b), reg_a_lo=a - t * se_a, reg_a_hi=a + t * se_a,
             reg_b_lo=b - t * se_b, reg_b_hi=b + t * se_b,
             reg_r2=float(1 - (res ** 2).sum() / ((y - y.mean()) ** 2).sum()),
             _s2=s2, _xm=xm, _sxx=sxx, _n=n, _t=t)
    r["reg_prop"] = not (r["reg_b_lo"] <= 1 <= r["reg_b_hi"])
    r["reg_const"] = not (r["reg_a_lo"] <= 0 <= r["reg_a_hi"])
    return r


def vies_regressao_no_nivel(rg, xc):
    """(vies%, lo, hi, motivo) no nivel do controle."""
    if "reg_a" not in rg:
        return None, None, None, rg.get("reg_motivo", "sem regressao")
    if xc is None or xc <= 0:
        return None, None, None, "media do controle ausente"
    if not (rg["reg_xmin"] / 1.25 <= xc <= rg["reg_xmax"] * 1.25):
        return None, None, None, ("extrapolacao: controle %.4g fora de %.4g-%.4g (unidade/faixa)"
                                  % (xc, rg["reg_xmin"], rg["reg_xmax"]))
    yhat = rg["reg_a"] + rg["reg_b"] * xc
    se = math.sqrt(rg["_s2"] * (1 / rg["_n"] + (xc - rg["_xm"]) ** 2 / rg["_sxx"]))
    v = (yhat - xc) / xc * 100
    d = rg["_t"] * se / xc * 100
    return v, v - d, v + d, ""


def bootstrap(amostras, nboot, rng):
    """Reamostra amostras DENTRO de cada rodada (preserva a estrutura)."""
    g = por_rodada(amostras)
    grupos = [(np.array([a["babs"] for a in v]), np.array([a["b"] for a in v])) for v in g.values()]
    A = np.empty(nboot)
    B = np.empty(nboot)
    G = np.empty(nboot)
    for i in range(nboot):
        ma, mb = [], []
        for babs, bs in grupos:
            idx = rng.integers(0, len(bs), len(bs))
            ma.append(babs[idx].mean())
            mb.append(bs[idx].mean())
        A[i] = np.mean(ma)
        B[i] = abs(np.mean(mb))
        G[i] = math.sqrt(np.mean(np.square(mb)))
    return A, B, G


# ------------------------------------------------------------- por produto
def ler_estatistica(wb):
    ws = wb[ABA_EST]
    linhas = []
    for i, row in enumerate(ws.iter_rows(min_row=14, max_row=400, max_col=47, values_only=True), 14):
        an = row[0]
        nivel_ok = isinstance(row[1], int) and not isinstance(row[1], bool) and 1 <= row[1] <= 3
        if an not in (None, "", 0) and not nivel_ok:
            break                 # fim da tabela principal (comeca o RESUMO / outro bloco)
        if not isinstance(an, str) or not an.strip():
            continue
        linhas.append(dict(lin=i, analito=an.strip(), nivel=row[1], n=row[2], media=num_real(row[3]),
                           cv=num_real(row[5]), G=num_real(row[6]), tea=num_real(row[10]),
                           L=num_real(row[11]), M=row[12], T=num_real(row[19]), X=row[23],
                           AC=num_real(row[28]), AD=num_real(row[29]), AR=num_real(row[43])))
    return linhas


def provedores_bio(wb):
    ws = wb[ABA_AN]
    m = {}
    for row in ws.iter_rows(min_row=4, max_row=43, max_col=44, values_only=True):
        if isinstance(row[0], str) and row[0].strip():
            m[up(row[0])] = row[43]
    return m


def sigma(tea, vies, cv):
    if tea is None or cv in (None, 0) or vies is None:
        return None
    return (tea - abs(vies)) / cv


def processar(rotulo, caminho, por_analito_prov, nboot, seed):
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ano_ep = nome_definido(wb, "eqAnoEP")
    prov_glob = nome_definido(wb, "eqProvedor")
    rod = nome_definido(wb, "eqRodada")
    faixas = ler_plano(wb)
    eqa = ler_eqa(wb)
    est = ler_estatistica(wb)
    provmap = provedores_bio(wb) if por_analito_prov else {}
    wb.close()
    rng = np.random.default_rng(seed)

    cache = {}
    linhas_out = []
    valid = defaultdict(int)
    for e in est:
        an = e["analito"]
        prov = (provmap.get(up(an)) or "CAP") if por_analito_prov else prov_glob
        if an not in cache:
            anov, s1, s2, nvazio = selecionar(eqa, an, ano_ep, prov, rod)
            c = {"ano": anov, "prov": prov, "nvazio": nvazio}
            for nome, s in (("S1", s1), ("S2", s2)):
                rv = resumo_vies(s)
                rv.update(regressao(s))
                if s:
                    rv["bootA"], rv["bootB"], rv["bootG"] = bootstrap(s, nboot, rng)
                c[nome] = rv
            cache[an] = c
        c = cache[an]
        out = OrderedDict(produto=rotulo, analito=an, nivel=e["nivel"], provedor=c["prov"],
                          ano_ceq=c["ano"], cv_pct=e["cv"], tea_pct=e["tea"], media_controle=e["media"],
                          sigma_planilha=e["L"], classe_planilha=e["M"])
        for sel in ("S1", "S2"):
            rv = c[sel]
            p = sel + "_"
            out[p + "n_amostras"] = rv.get("n")
            out[p + "n_rodadas"] = rv.get("k")
            for k in ("A", "Bsig2", "B", "Bsig1", "DP", "EP", "IC_lo", "IC_hi", "EProd", "ICr_lo",
                      "ICr_hi", "C", "G", "F", "mediana", "DPw", "DPentre_medias", "anova_F", "anova_p"):
                out[p + k] = rv.get(k)
            out[p + "sig_amostras"] = rv.get("sig_am")
            out[p + "sig_rodadas"] = rv.get("sig_rod")
            out[p + "inversao_sinal"] = rv.get("inversao_sinal")
            out[p + "heterog_rodadas"] = rv.get("heterog")
            out[p + "n_outliers"] = rv.get("n_outliers")
            out[p + "medias_rodada"] = " | ".join("%.3f" % x for x in rv.get("medias_rodada", []))
            out[p + "razao_DPw_CV"] = (rv["DPw"] / e["cv"]) if rv.get("DPw") and e["cv"] else None
            for k in ("reg_n", "reg_nconc", "reg_a", "reg_a_lo", "reg_a_hi", "reg_b", "reg_b_lo",
                      "reg_b_hi", "reg_r2", "reg_prop", "reg_const"):
                out[p + k] = rv.get(k)
            vd, vlo, vhi, mot = vies_regressao_no_nivel(rv, e["media"])
            out[p + "D"], out[p + "D_lo"], out[p + "D_hi"], out[p + "D_motivo"] = vd, vlo, vhi, mot
            vals = {"A": rv.get("A"), "B": rv.get("B"), "C": rv.get("C"), "D": vd, "F": rv.get("F"),
                    "G": rv.get("G")}
            for k, v in vals.items():
                s = sigma(e["tea"], v, e["cv"])
                f = plano(faixas, s)
                out[p + "sigma_" + k] = s
                out[p + "classe_" + k] = f["classe"] if f else ""
                out[p + "regras_" + k] = f["regras"] if f else ""
                out[p + "N_" + k] = f["n"] if f else ""
                out[p + "runsize_" + k] = f["runsize"] if f else ""
                out[p + "freq_" + k] = f["freq"] if f else ""
                out[p + "ET_" + k] = (1.65 * e["cv"] + abs(v)) if (v is not None and e["cv"] is not None) else None
            # Sigma_B no limite desfavoravel do IC95% do vies com sinal
            for nm, lo, hi in (("am", rv.get("IC_lo"), rv.get("IC_hi")), ("rod", rv.get("ICr_lo"), rv.get("ICr_hi"))):
                if lo is None or e["tea"] is None or not e["cv"]:
                    out[p + "sigma_B_pior_IC_" + nm] = None
                else:
                    out[p + "sigma_B_pior_IC_" + nm] = (e["tea"] - max(abs(lo), abs(hi))) / e["cv"]
            # E -- bootstrap: IC do Sigma e estabilidade da classe
            for k, arr in (("A", rv.get("bootA")), ("B", rv.get("bootB")), ("G", rv.get("bootG"))):
                s0 = out[p + "sigma_" + k]
                if arr is None or s0 is None:
                    for nm in ("E_sig%s_lo", "E_sig%s_hi", "E_estab%s", "E_bias%s_lo", "E_bias%s_hi"):
                        out[p + nm % k] = None
                    continue
                sb = (e["tea"] - arr) / e["cv"]
                cl0 = out[p + "classe_" + k]
                same = np.mean([(plano(faixas, x) or {}).get("classe") == cl0 for x in sb])
                out[p + "E_sig%s_lo" % k] = float(np.percentile(sb, 2.5))
                out[p + "E_sig%s_hi" % k] = float(np.percentile(sb, 97.5))
                out[p + "E_estab%s" % k] = float(same)
                out[p + "E_bias%s_lo" % k] = float(np.percentile(arr, 2.5))
                out[p + "E_bias%s_hi" % k] = float(np.percentile(arr, 97.5))
        # validacao contra a planilha (so S1 = regra do BiasEQ)
        if e["G"] is not None:
            valid["G_total"] += 1
            if out["S1_A"] is not None and abs(out["S1_A"] - e["G"]) < 1e-9:
                valid["G_ok"] += 1
            else:
                valid["G_falha"] += 1
                print("  [!] A difere de G: %s N%s A=%s G=%s" % (an, e["nivel"], out["S1_A"], e["G"]))
        if e["T"] is not None and out["S1_Bsig2"] is not None:
            valid["T_total"] += 1
            valid["T_ok"] += abs(out["S1_Bsig2"] - e["T"]) < 1e-9
        if e["AR"] is not None and out["S2_Bsig1"] is not None:
            valid["AR_total"] += 1
            valid["AR_ok"] += abs(out["S2_Bsig1"] - e["AR"]) < 1e-9
        if e["AC"] is not None:
            valid["N_total"] += 1
            valid["N_ok"] += (out["S1_n_amostras"] == e["AC"] and out["S1_n_rodadas"] == e["AD"])
        if e["L"] is not None:
            valid["L_total"] += 1
            if out["S1_sigma_A"] is not None and abs(out["S1_sigma_A"] - e["L"]) < 1e-9:
                valid["L_ok"] += 1
            valid["M_ok"] += (out["S1_classe_A"] == e["M"])
            valid["X_ok"] += (out["S1_regras_A"] == (e["X"] or ""))
        linhas_out.append(out)
    meta = dict(ano_ep=ano_ep, provedor=prov_glob, rodada=rod, n_linhas=len(linhas_out),
                n_analitos=len(cache), vazios=sum(c["nvazio"] for c in cache.values()))
    return linhas_out, dict(valid), meta


# ------------------------------------------------------------- relatorio
def matriz(linhas, sel, k):
    m = defaultdict(int)
    sobe = desce = igual = 0
    for o in linhas:
        ca, ck = o[sel + "_classe_A"], o[sel + "_classe_" + k]
        if not ca or not ck:
            continue
        m[(ca, ck)] += 1
        ia, ik = idx_classe(ca), idx_classe(ck)
        if ik > ia:
            sobe += 1
        elif ik < ia:
            desce += 1
        else:
            igual += 1
    return m, sobe, desce, igual


def fmt(v, d=2):
    if v is None or v == "":
        return "-"
    if isinstance(v, bool):
        return "sim" if v else "nao"
    if isinstance(v, float):
        return ("%." + str(d) + "f") % v
    return str(v)


def imprimir_resumo(rotulo, linhas, valid, meta):
    print("\n" + "=" * 78)
    print("%s  | Ano EP=%s  provedor=%s  rodada=%s  | %d linhas, %d analitos"
          % (rotulo, meta["ano_ep"], meta["provedor"], meta["rodada"], meta["n_linhas"], meta["n_analitos"]))
    print("Validacao: A=G %d/%d | Bsig2=T %d/%d | n/rodadas=AC/AD %d/%d | Sigma_A=L %d/%d | "
          "classe=M %d/%d | regras=X %d/%d | S2 Bsig1=ViesEQ MEDIA %d/%d | bias vazio contado como 0 (S1): %d"
          % (valid.get("G_ok", 0), valid.get("G_total", 0), valid.get("T_ok", 0), valid.get("T_total", 0),
             valid.get("N_ok", 0), valid.get("N_total", 0), valid.get("L_ok", 0), valid.get("L_total", 0),
             valid.get("M_ok", 0), valid.get("L_total", 0), valid.get("X_ok", 0), valid.get("L_total", 0),
             valid.get("AR_ok", 0), valid.get("AR_total", 0), meta["vazios"]))
    for sel in ("S1", "S2"):
        print("\n-- selecao %s --" % sel)
        com_sigma = sum(1 for o in linhas if o[sel + "_classe_A"])
        print("linhas com Sigma (A): %d" % com_sigma)
        for k in ("B", "C", "D", "F", "G"):
            m, s, d, i = matriz(linhas, sel, k)
            tot = s + d + i
            print("A->%s: comparaveis=%d  sobe=%d  desce=%d  igual=%d" % (k, tot, s, d, i))
            for (ca, ck), v in sorted(m.items(), key=lambda x: (idx_classe(x[0][0]), idx_classe(x[0][1]))):
                if ca != ck:
                    print("     %-22s -> %-22s %d" % (ca, ck, v))
        # distribuicao de classes
        for k in ("A", "B", "C", "D", "F", "G"):
            cnt = defaultdict(int)
            for o in linhas:
                if o[sel + "_classe_" + k]:
                    cnt[o[sel + "_classe_" + k]] += 1
            print("  classes %s: %s" % (k, ", ".join("%s=%d" % (c, cnt[c]) for c in ORDEM_CLASSE)))
        # estabilidade bootstrap
        for k in ("A", "B", "G"):
            v = [o[sel + "_E_estab" + k] for o in linhas if o.get(sel + "_E_estab" + k) is not None
                 and o[sel + "_classe_A"]]
            if v:
                print("  bootstrap %s: estabilidade da classe mediana=%.2f; linhas com <80%% na mesma classe=%d/%d"
                      % (k, float(np.median(v)), sum(1 for x in v if x < 0.8), len(v)))
        dok = sum(1 for o in linhas if o[sel + "_D"] is not None)
        print("  D calculavel no nivel: %d/%d linhas" % (dok, len(linhas)))
        mot = defaultdict(int)
        for o in linhas:
            if o[sel + "_D"] is None:
                mot[o[sel + "_D_motivo"].split(":")[0]] += 1
        for k, v in mot.items():
            print("     sem D: %s = %d" % (k, v))

    # efeito da SELECAO (S1 -> S2) com o mesmo estimador
    for k in ("A", "B", "C", "G"):
        dif = [(o["analito"], o["nivel"], o["S1_classe_" + k], o["S2_classe_" + k]) for o in linhas
               if o["S1_classe_" + k] and o["S2_classe_" + k] and o["S1_classe_" + k] != o["S2_classe_" + k]]
        print("S1->S2 com %s: %d linha(s) mudam de classe %s" % (k, len(dif), dif if dif else ""))
    # linhas que sobem com B: quantas tem IC95 incluindo zero / inversao de sinal / heterogeneidade
    sob = [o for o in linhas if o["S1_classe_A"] and o["S1_classe_B"]
           and idx_classe(o["S1_classe_B"]) > idx_classe(o["S1_classe_A"])]
    if sob:
        print("Sobem com B (S1): %d | IC95 amostra inclui 0: %d | IC95 rodada inclui 0: %d | inversao de sinal: %d | "
              "rodadas heterogeneas (p<0,05): %d" % (
                  len(sob), sum(1 for o in sob if o["S1_IC_lo"] <= 0 <= o["S1_IC_hi"]),
                  sum(1 for o in sob if o["S1_ICr_lo"] <= 0 <= o["S1_ICr_hi"]),
                  sum(1 for o in sob if o["S1_inversao_sinal"]), sum(1 for o in sob if o["S1_heterog_rodadas"])))
    an = {}
    for o in linhas:
        if o["S1_n_amostras"]:
            an[o["analito"]] = o
    if an:
        print("Analitos com CEQ: %d | rodadas heterogeneas: %d | inversao de sinal entre rodadas: %d | "
              "IC95 amostra inclui 0: %d | IC95 rodada inclui 0: %d | vies proporcional: %d | constante: %d | "
              "com outlier: %d | A >= 2B: %d | mediana A/B=%.2f C/A=%.2f G/A=%.2f" % (
                  len(an), sum(1 for o in an.values() if o["S1_heterog_rodadas"]),
                  sum(1 for o in an.values() if o["S1_inversao_sinal"]),
                  sum(1 for o in an.values() if o["S1_IC_lo"] is not None and o["S1_IC_lo"] <= 0 <= o["S1_IC_hi"]),
                  sum(1 for o in an.values() if o["S1_ICr_lo"] is not None and o["S1_ICr_lo"] <= 0 <= o["S1_ICr_hi"]),
                  sum(1 for o in an.values() if o["S1_reg_prop"]), sum(1 for o in an.values() if o["S1_reg_const"]),
                  sum(1 for o in an.values() if o["S1_n_outliers"]),
                  sum(1 for o in an.values() if o["S1_B"] is not None and o["S1_A"] >= 2 * o["S1_B"]),
                  float(np.median([o["S1_A"] / o["S1_B"] for o in an.values() if o["S1_B"]])),
                  float(np.median([o["S1_C"] / o["S1_A"] for o in an.values() if o["S1_A"]])),
                  float(np.median([o["S1_G"] / o["S1_A"] for o in an.values() if o["S1_A"]]))))

    # por analito (S1): flags
    print("\n-- por analito (S1 | S2) --")
    print("analito | n,k | A | Bsig2 [IC95 amostra] [IC95 rodada] | C | G | medias rodada | DPw | ANOVA p | "
          "outl | prop/const | S2: n A Bsig1 C G")
    vistos = set()
    for o in linhas:
        if o["analito"] in vistos:
            continue
        vistos.add(o["analito"])
        print("%s | %s,%s | %s | %s [%s;%s] [%s;%s] | %s | %s | %s | %s | %s | %s | %s/%s | S2: %s %s %s %s %s" % (
            o["analito"], o["S1_n_amostras"], o["S1_n_rodadas"], fmt(o["S1_A"]), fmt(o["S1_Bsig2"]),
            fmt(o["S1_IC_lo"]), fmt(o["S1_IC_hi"]), fmt(o["S1_ICr_lo"]), fmt(o["S1_ICr_hi"]), fmt(o["S1_C"]),
            fmt(o["S1_G"]), o["S1_medias_rodada"], fmt(o["S1_DPw"]), fmt(o["S1_anova_p"], 4), o["S1_n_outliers"],
            fmt(o["S1_reg_prop"]), fmt(o["S1_reg_const"]), o["S2_n_amostras"], fmt(o["S2_A"]), fmt(o["S2_Bsig1"]),
            fmt(o["S2_C"]), fmt(o["S2_G"])))

    print("\n-- linhas: Sigma por estimador (S1); * = classe A != classe B --")
    print("analito N | CV | TEa | sA clA | sB clB | sC clC | sD clD | sF | sG clG | E: sigA IC estabA | "
          "sigB IC estabB | DPw/CV | flags")
    for o in linhas:
        if not o["S1_classe_A"]:
            continue
        flag = "*" if o["S1_classe_A"] != o["S1_classe_B"] else " "
        fl = []
        if o["S1_IC_lo"] is not None and o["S1_IC_lo"] <= 0 <= o["S1_IC_hi"]:
            fl.append("IC95(am) inclui 0")
        if o["S1_ICr_lo"] is not None and o["S1_ICr_lo"] <= 0 <= o["S1_ICr_hi"]:
            fl.append("IC95(rod) inclui 0")
        if o["S1_heterog_rodadas"]:
            fl.append("rodadas heterogeneas")
        if o["S1_inversao_sinal"]:
            fl.append("inversao de sinal")
        if o["S1_reg_prop"]:
            fl.append("vies proporcional")
        if o["S1_n_outliers"]:
            fl.append("%d outlier(s)" % o["S1_n_outliers"])
        if o["S1_A"] and o["S1_B"] is not None and o["S1_A"] >= 2 * o["S1_B"]:
            fl.append("A>=2B")
        print("%s%s N%s | %s | %s | %s %s | %s %s | %s %s | %s %s | %s | %s %s | [%s;%s] %s | [%s;%s] %s | %s | %s" % (
            flag, o["analito"], o["nivel"], fmt(o["cv_pct"]), fmt(o["tea_pct"]),
            fmt(o["S1_sigma_A"]), abrev(o["S1_classe_A"]), fmt(o["S1_sigma_B"]), abrev(o["S1_classe_B"]),
            fmt(o["S1_sigma_C"]), abrev(o["S1_classe_C"]), fmt(o["S1_sigma_D"]),
            abrev(o["S1_classe_D"]) or o["S1_D_motivo"][:14], fmt(o["S1_sigma_F"]), fmt(o["S1_sigma_G"]),
            abrev(o["S1_classe_G"]), fmt(o["S1_E_sigA_lo"]), fmt(o["S1_E_sigA_hi"]), fmt(o["S1_E_estabA"]),
            fmt(o["S1_E_sigB_lo"]), fmt(o["S1_E_sigB_hi"]), fmt(o["S1_E_estabB"]), fmt(o["S1_razao_DPw_CV"]),
            "; ".join(fl)))


def abrev(c):
    return {"Desempenho inadequado": "Inadeq", "Marginal": "Marg", "Bom": "Bom", "Excelente": "Exc",
            "Classe mundial": "Mund"}.get(c or "", c or "")


def gravar_csv(caminho, linhas):
    cols = list(linhas[0].keys())
    with open(caminho, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(cols)
        for o in linhas:
            row = []
            for c in cols:
                v = o[c]
                if isinstance(v, (bool, np.bool_)):
                    v = "sim" if v else "nao"
                elif isinstance(v, (float, np.floating)):
                    v = repr(float(v)).replace(".", ",")
                elif v is None:
                    v = ""
                row.append(v)
            w.writerow(row)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--bio", help="COPIA de QC_Bioquimica.xlsm")
    ap.add_argument("--hema", help="COPIA de QC_Hematologia.xlsm")
    ap.add_argument("--saida", default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("--boot", type=int, default=2000)
    ap.add_argument("--seed", type=int, default=20261006)
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print("D-01 estudo de estimadores de vies | bootstrap=%d seed=%d" % (a.boot, a.seed))
    for rotulo, cam, porprov in (("bioquimica", a.bio, True), ("hematologia", a.hema, False)):
        if not cam:
            continue
        linhas, valid, meta = processar(rotulo, cam, porprov, a.boot, a.seed)
        imprimir_resumo(rotulo.upper(), linhas, valid, meta)
        destino = os.path.join(a.saida, "d01_resultados_%s.csv" % rotulo)
        gravar_csv(destino, linhas)
        print("\nCSV: %s" % os.path.basename(destino))


if __name__ == "__main__":
    main()
