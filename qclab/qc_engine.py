"""
qc_engine.py
------------
Núcleo de cálculo do Controle de Qualidade Interno (independente de banco/UI).

Reproduz a lógica da planilha QC-INI e implementa as Regras de Westgard
diferenciadas por setor:

  HEMATOLOGIA  (3 níveis: L1, L2, L3)
  ─────────────────────────────────────────────────────────────────────
  • 1-3S    → |Z| > 3 em QUALQUER nível na corrida t            (Bloqueio)
  • 2of3-2S → 2 dos 3 níveis na MESMA corrida t com Z > +2
              ou Z < −2 (mesma direção)                          (Bloqueio)
  • R4S     → Amplitude Z (max−min) > 4 na mesma corrida t      (Bloqueio)
  • 3-1S    → 3 resultados consecutivos do MESMO nível > 1 DP
              (mesmo lado)                                        (Bloqueio)
  • 6X      → 6 resultados consecutivos do MESMO nível
              no mesmo lado da média (Z>0 ou Z<0)                (Bloqueio)

  BIOQUÍMICA   (2 níveis: L1, L2)
  ─────────────────────────────────────────────────────────────────────
  • 1-3S    → |Z| > 3 em QUALQUER nível na corrida t            (Bloqueio)
  • 2-2S    → Intra-corrida: ambos os níveis t com Z>+2 ou Z<−2
              OU Inter-corrida: mesmo nível t e t-1 com Z>+2
              ou Z<−2                                             (Bloqueio)
  • R4S     → |Z_L1 − Z_L2| > 4 E sinais opostos na corrida t  (Bloqueio)
  • 4-1S    → Mesmo nível: 4 consecutivos > 1 DP mesmo lado
              OU Inter-nível: últimos 2 de L1 E últimos 2 de L2
              > 1 DP mesmo lado (janela simultânea)              (Bloqueio)
  • 8X      → Mesmo nível: 8 consecutivos mesmo lado da média
              OU Inter-nível: últimos 4 de L1 E últimos 4 de L2
              todos mesmo lado da média                           (Bloqueio)

Estatísticas:
  * Média, DP (amostral), CV%
  * Métricas Sigma / Erro Total:
        ETP = TEa CLIA (preferência) ou Erro Total por VB (EFLM)
        ES  = |viés| do controle externo
        EA  = z * CV%   (z = 1.96 → 95 %)
        ET  = ES + EA
        Sigma = (ETP − |viés|) / CV%

Todas as frações (TEa, CV, viés) são tratadas como proporção (0.07 == 7 %).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from statistics import mean as _mean, stdev as _stdev
from typing import Optional


# --------------------------------------------------------------------------- #
# Especificações de qualidade (multiplicadores EFLM por nível de desempenho)
# --------------------------------------------------------------------------- #
PERF_FACTORS = {
    "OTI": (0.25, 0.125),
    "ÓTI": (0.25, 0.125),
    "DES": (0.50, 0.250),
    "MIN": (0.75, 0.375),
}
EA_Z = 1.65   # fator do componente aleatório (padronizado em 1,65 em todo o
              # sistema: ET = CV% × 1,65 + Bias). Antes era 1,96 (bicaudal).


# --------------------------------------------------------------------------- #
# Estruturas de resultado
# --------------------------------------------------------------------------- #
@dataclass
class PointResult:
    seq: int
    value: float
    z: Optional[float]                        # (valor − média) / DP
    flags: list = field(default_factory=list)  # ex.: ["1-3S", "3-1S"]
    status: str = "OK"                         # OK | ALERTA | REJEITADO
    is_nc: bool = False                        # marcado manualmente como NC


@dataclass
class AnalyteLevelStats:
    analyte: str
    level: int
    n: int = 0
    mean_obs: Optional[float] = None
    sd_obs: Optional[float] = None
    cv_obs: Optional[float] = None
    es: Optional[float] = None
    ea: Optional[float] = None
    et: Optional[float] = None
    etp: Optional[float] = None
    etp_fonte: str = "-"
    etp_clia: Optional[float] = None       # ETp pela TEa CLIA
    etp_vb: Optional[float] = None          # ETp pela Variação Biológica
    sigma: Optional[float] = None
    cv_limite_clia: Optional[float] = None
    cv_limite_vb: Optional[float] = None
    cv_limite_fb: Optional[float] = None    # limite de CV% do fabricante
    cv_status: str = "-"
    rule_summary: dict = field(default_factory=dict)


# Regras que implicam REJEIÇÃO (as demais são só alerta)
REJECT_RULES = {"1-3S", "2of3-2S", "2-2S", "R4S", "3-1S", "4-1S", "6X", "8X"}


# --------------------------------------------------------------------------- #
# Estatística básica
# --------------------------------------------------------------------------- #
def compute_stats(values: list[float]):
    """Retorna (n, media, dp_amostral, cv_fracao)."""
    vals = [v for v in values if v is not None]
    n = len(vals)
    if n == 0:
        return 0, None, None, None
    m = _mean(vals)
    sd = _stdev(vals) if n >= 2 else 0.0
    cv = (sd / m) if m not in (0, None) else None
    return n, m, sd, cv


# --------------------------------------------------------------------------- #
# Z-score auxiliar
# --------------------------------------------------------------------------- #
def _zscore(value, ref_mean, ref_sd):
    if ref_sd in (None, 0) or value is None or ref_mean is None:
        return None
    return (value - ref_mean) / ref_sd


# =========================================================================== #
#  HEMATOLOGIA — 3 NÍVEIS (L1, L2, L3)
# =========================================================================== #

def _evaluate_hematologia(
    series_by_level: dict,
    ref_by_level: dict,
) -> dict[int, list[PointResult]]:
    """
    Regras para Hematologia (3 níveis simultâneos por corrida):
      1-3S, 2of3-2S, R4S, 3-1S, 6X
    """

    # ── 1. Calcular z-scores ─────────────────────────────────────────────── #
    results: dict[int, list[PointResult]] = {}
    for level, serie in series_by_level.items():
        ref_mean, ref_sd = ref_by_level.get(level, (None, None))
        pts: list[PointResult] = []
        for seq, val in serie:
            z = _zscore(val, ref_mean, ref_sd)
            pts.append(PointResult(seq=seq, value=val, z=z))
        results[level] = pts

    # ── 2. Regras INTRA-NÍVEL (avaliadas por nível ao longo do tempo) ────── #
    for level, pts in results.items():
        zs = [p.z for p in pts]

        for i, p in enumerate(pts):
            if p.z is None:
                continue

            # 1-3S ─────────────────────────────────────────────────────────
            if abs(p.z) > 3:
                p.flags.append("1-3S")

            # 3-1S ─────────────────────────────────────────────────────────
            # 3 resultados consecutivos do mesmo nível > 1 DP, mesmo lado
            if i >= 2:
                ult3 = zs[i - 2 : i + 1]
                if all(z is not None and z > 1 for z in ult3):
                    p.flags.append("3-1S")
                elif all(z is not None and z < -1 for z in ult3):
                    p.flags.append("3-1S")

            # 6X ───────────────────────────────────────────────────────────
            # 6 resultados consecutivos do mesmo nível no mesmo lado da média
            if i >= 5:
                ult6 = zs[i - 5 : i + 1]
                if all(z is not None and z > 0 for z in ult6):
                    p.flags.append("6X")
                elif all(z is not None and z < 0 for z in ult6):
                    p.flags.append("6X")

    # ── 3. Regras INTER-NÍVEL por corrida (seq) ──────────────────────────── #
    # Mapeia seq → {level: PointResult} para consulta rápida
    seq_map: dict[int, dict[int, PointResult]] = {}
    for level, pts in results.items():
        for p in pts:
            seq_map.setdefault(p.seq, {})[level] = p

    for seq, level_map in seq_map.items():
        pts_run = list(level_map.values())
        zvals = [p.z for p in pts_run if p.z is not None]
        if not zvals:
            continue

        # 2of3-2S ─────────────────────────────────────────────────────────
        # 2 dos 3 níveis na MESMA corrida com Z > +2 OU Z < −2
        zs_run = {lvl: p.z for lvl, p in level_map.items() if p.z is not None}
        levels_pos2 = [lvl for lvl, z in zs_run.items() if z > 2]
        levels_neg2 = [lvl for lvl, z in zs_run.items() if z < -2]

        if len(levels_pos2) >= 2 or len(levels_neg2) >= 2:
            for p in pts_run:
                if "2of3-2S" not in p.flags:
                    p.flags.append("2of3-2S")

        # R4S ─────────────────────────────────────────────────────────────
        # Um nível ultrapassa +2 DP E outro nível ultrapassa −2 DP na mesma
        # corrida (lados opostos cruzando a linha de 2 DP → amplitude > 4 DP).
        if len(zvals) >= 2:
            z_max, z_min = max(zvals), min(zvals)
            if z_max > 2 and z_min < -2:
                for p in pts_run:
                    if p.z is not None and "R4S" not in p.flags:
                        p.flags.append("R4S")

    # ── 4. Status final ──────────────────────────────────────────────────── #
    for pts in results.values():
        for p in pts:
            if set(p.flags) & REJECT_RULES:
                p.status = "REJEITADO"
            else:
                p.status = "OK"

    return results


# =========================================================================== #
#  BIOQUÍMICA — 2 NÍVEIS (L1, L2)
# =========================================================================== #

def _evaluate_bioquimica(
    series_by_level: dict,
    ref_by_level: dict,
) -> dict[int, list[PointResult]]:
    """
    Regras para Bioquímica (2 níveis simultâneos por corrida):
      1-3S, 2-2S (intra + inter corrida), R4S, 4-1S, 8X
    """

    # ── 1. Calcular z-scores ─────────────────────────────────────────────── #
    results: dict[int, list[PointResult]] = {}
    for level, serie in series_by_level.items():
        ref_mean, ref_sd = ref_by_level.get(level, (None, None))
        pts: list[PointResult] = []
        for seq, val in serie:
            z = _zscore(val, ref_mean, ref_sd)
            pts.append(PointResult(seq=seq, value=val, z=z))
        results[level] = pts

    # ── 2. Regras INTRA-NÍVEL (avaliadas por nível ao longo do tempo) ────── #
    for level, pts in results.items():
        zs = [p.z for p in pts]

        for i, p in enumerate(pts):
            if p.z is None:
                continue

            # 1-3S ─────────────────────────────────────────────────────────
            if abs(p.z) > 3:
                p.flags.append("1-3S")

            # 2-2S Inter-corrida (mesmo nível, corrida t e t-1) ────────────
            if i >= 1 and zs[i - 1] is not None:
                z_prev = zs[i - 1]
                if (p.z > 2 and z_prev > 2) or (p.z < -2 and z_prev < -2):
                    p.flags.append("2-2S")

            # 4-1S intra-nível: 4 consecutivos > 1 DP mesmo lado ──────────
            if i >= 3:
                ult4 = zs[i - 3 : i + 1]
                if all(z is not None and z > 1 for z in ult4):
                    p.flags.append("4-1S")
                elif all(z is not None and z < -1 for z in ult4):
                    p.flags.append("4-1S")

            # 8X intra-nível: 8 consecutivos mesmo lado da média ──────────
            if i >= 7:
                ult8 = zs[i - 7 : i + 1]
                if all(z is not None and z > 0 for z in ult8):
                    p.flags.append("8X")
                elif all(z is not None and z < 0 for z in ult8):
                    p.flags.append("8X")

    # ── 3. Regras INTER-NÍVEL por corrida (seq) ──────────────────────────── #
    # Mapeia seq → {level: (index_no_array, PointResult)}
    seq_to_pts: dict[int, dict[int, tuple[int, PointResult]]] = {}
    for level, pts in results.items():
        for i, p in enumerate(pts):
            seq_to_pts.setdefault(p.seq, {})[level] = (i, p)

    # Obtemos seqs ordenadas para poder percorrer janelas inter-nível
    all_seqs = sorted(seq_to_pts.keys())

    for seq_idx, seq in enumerate(all_seqs):
        level_map = seq_to_pts[seq]  # {level: (i, pt)}

        # 2-2S Intra-corrida: ambos os níveis na mesma corrida t > +2 ou < −2
        zs_run = {lvl: data[1].z for lvl, data in level_map.items()
                  if data[1].z is not None}
        if len(zs_run) >= 2:
            zlist = list(zs_run.values())
            if all(z > 2 for z in zlist) or all(z < -2 for z in zlist):
                for _, (_, p) in level_map.items():
                    if "2-2S" not in p.flags:
                        p.flags.append("2-2S")

        # R4S: um nível ultrapassa +2 DP E o outro ultrapassa −2 DP ────────
        # (lados opostos cruzando 2 DP → distância total > 4 DP)
        if len(zs_run) == 2:
            z_vals = list(zs_run.values())
            z_max, z_min = max(z_vals), min(z_vals)
            if z_max > 2 and z_min < -2:
                for _, (_, p) in level_map.items():
                    if "R4S" not in p.flags:
                        p.flags.append("R4S")

        # 4-1S Inter-nível: últimos 2 de L1 E últimos 2 de L2 > 1 DP
        # mesmo lado da média (janela das 2 corridas mais recentes de cada nível)
        if len(results) == 2:
            levels = list(results.keys())
            l1_pts = results[levels[0]]
            l2_pts = results[levels[1]]

            # índices dos pontos de cada nível na corrida atual
            idx_l1 = level_map.get(levels[0], (None,))[0]
            idx_l2 = level_map.get(levels[1], (None,))[0]

            if idx_l1 is not None and idx_l1 >= 1 and \
               idx_l2 is not None and idx_l2 >= 1:

                zs_l1_2 = [l1_pts[idx_l1 - 1].z, l1_pts[idx_l1].z]
                zs_l2_2 = [l2_pts[idx_l2 - 1].z, l2_pts[idx_l2].z]

                if all(z is not None for z in zs_l1_2 + zs_l2_2):
                    todos_pos = all(z > 1 for z in zs_l1_2 + zs_l2_2)
                    todos_neg = all(z < -1 for z in zs_l1_2 + zs_l2_2)
                    if todos_pos or todos_neg:
                        for lvl in levels:
                            _, pt = level_map.get(lvl, (None, None))
                            if pt is not None and "4-1S" not in pt.flags:
                                pt.flags.append("4-1S")

        # 8X Inter-nível: últimos 4 de L1 E últimos 4 de L2 mesmo lado ────
        if len(results) == 2:
            levels = list(results.keys())
            l1_pts = results[levels[0]]
            l2_pts = results[levels[1]]

            idx_l1 = level_map.get(levels[0], (None,))[0]
            idx_l2 = level_map.get(levels[1], (None,))[0]

            if idx_l1 is not None and idx_l1 >= 3 and \
               idx_l2 is not None and idx_l2 >= 3:

                zs_l1_4 = [l1_pts[j].z for j in range(idx_l1 - 3, idx_l1 + 1)]
                zs_l2_4 = [l2_pts[j].z for j in range(idx_l2 - 3, idx_l2 + 1)]

                if all(z is not None for z in zs_l1_4 + zs_l2_4):
                    todos_pos = all(z > 0 for z in zs_l1_4 + zs_l2_4)
                    todos_neg = all(z < 0 for z in zs_l1_4 + zs_l2_4)
                    if todos_pos or todos_neg:
                        for lvl in levels:
                            _, pt = level_map.get(lvl, (None, None))
                            if pt is not None and "8X" not in pt.flags:
                                pt.flags.append("8X")

    # ── 4. Status final ──────────────────────────────────────────────────── #
    for pts in results.values():
        for p in pts:
            if set(p.flags) & REJECT_RULES:
                p.status = "REJEITADO"
            else:
                p.status = "OK"

    return results


# =========================================================================== #
#  API PÚBLICA — ponto de entrada único
# =========================================================================== #

def evaluate_levels(
    series_by_level: dict,
    ref_by_level: dict,
    setor: str = "",
    nc_by_level: dict | None = None,
) -> dict[int, list[PointResult]]:
    """
    Avalia as Regras de Westgard para um analito em múltiplos níveis.

    Parâmetros
    ----------
    series_by_level : {nivel: [(seq, value), ...]}  ordenado por seq.
                      IMPORTANTE: passe a série CRONOLÓGICA COMPLETA
                      (incluindo pontos marcados como NC). As regras de
                      tendência (3-1S, 6X, 4-1S, 8X, 2-2S inter-corrida)
                      dependem de resultados realmente consecutivos — excluir
                      um ponto do meio faria a contagem "pular" e juntar
                      resultados não adjacentes.
    ref_by_level    : {nivel: (media_ref, dp_ref)}
    setor           : "Hematologia" | "Bioquimica" (case-insensitive)
                      — se omitido, infere pelo número de níveis
    nc_by_level     : {nivel: set(seq, ...)} dos pontos marcados como Não
                      Conforme. Esses pontos CONTAM na janela de
                      consecutividade, mas são marcados (is_nc=True) para que
                      a UI não os reporte como rejeições novas.

    Retorna
    -------
    {nivel: [PointResult, ...]}
    """
    setor_norm = setor.strip().upper()

    # Inferência automática quando setor não informado
    n_levels = len(series_by_level)
    if "HEMAT" in setor_norm or n_levels == 3:
        results = _evaluate_hematologia(series_by_level, ref_by_level)
    else:
        # Bioquímica ou qualquer cenário com 2 níveis
        results = _evaluate_bioquimica(series_by_level, ref_by_level)

    # Marca pontos NC (não altera flags/contagem — só sinaliza para a UI)
    if nc_by_level:
        for level, pts in results.items():
            ncset = nc_by_level.get(level, set())
            for p in pts:
                if p.seq in ncset:
                    p.is_nc = True

    return results


# --------------------------------------------------------------------------- #
# Resumo por analito/nível (para tabela de regras no painel)
# --------------------------------------------------------------------------- #
def summarize_rules(points: list[PointResult]) -> dict:
    """
    Retorna dict {código_regra: "OK" | código_regra} para exibição
    na tabela de status das regras.
    """
    # Todas as regras possíveis na ordem de exibição
    all_rules = ["1-3S", "2of3-2S", "2-2S", "R4S", "3-1S", "4-1S", "6X", "8X"]
    out = {}
    all_flags = {flag for p in points for flag in p.flags}
    for code in all_rules:
        out[code] = code if code in all_flags else "OK"
    return out


# --------------------------------------------------------------------------- #
# Sigma / Erro Total
# --------------------------------------------------------------------------- #
def vb_specs(cvw, cvg, perf):
    """Imprecisão e erro total desejáveis por Variação Biológica (frações)."""
    if cvw in (None, "", "-"):
        return None, None, None
    fator_i, fator_b = PERF_FACTORS.get((perf or "MIN").upper(), PERF_FACTORS["MIN"])
    cvg = 0.0 if cvg in (None, "", "-") else float(cvg)
    impr = (fator_i * float(cvw)) / 100.0
    vies = (fator_b * ((float(cvw) ** 2 + cvg ** 2) ** 0.5)) / 100.0
    te_vb = vies + 1.65 * impr
    return impr, vies, te_vb


def compute_sigma(cv_obs, bias_obs, clia_tea, cvw, cvg, perf,
                  etp_pct=None, etp_source=None, fab_cv=None):
    """
    Retorna dict com etp, etp_fonte, es, ea, et, sigma,
    cv_limite_clia, cv_limite_vb, cv_status.
    bias_obs : viés observado (fração) — single CQ externo OU Indicador de Bias
               Anual já convertido p/ fração; 0 se ausente.
    etp_pct  : Erro Total Permitido configurado no cadastro do ensaio (em %).
               Quando informado, tem prioridade sobre CLIA/VB.
    etp_source : origem do ETp configurado (CLIA | VB | Fabricante | ...).
    """
    bias_obs = abs(bias_obs) if bias_obs not in (None, "", "-") else 0.0
    impr_vb, _vies_vb, te_vb = vb_specs(cvw, cvg, perf)

    if etp_pct not in (None, "", "-"):
        etp, fonte = float(etp_pct) / 100.0, (etp_source or "Configurado")
    elif clia_tea not in (None, "", "-"):
        etp, fonte = float(clia_tea), "CLIA"
    elif te_vb is not None:
        etp, fonte = te_vb, "VB"
    else:
        etp, fonte = None, "-"

    cv_lim_clia = (float(clia_tea) / 3.0) if clia_tea not in (None, "", "-") else None
    cv_lim_vb = impr_vb
    cv_lim_fb = float(fab_cv) if fab_cv not in (None, "", "-") else None

    # ETp por fonte (frações) — exibidos lado a lado no painel
    etp_clia = float(clia_tea) if clia_tea not in (None, "", "-") else None
    etp_vb = te_vb

    es = bias_obs
    ea = (EA_Z * cv_obs) if cv_obs not in (None, "") else None
    et = (es + ea) if ea is not None else None
    sigma = ((etp - es) / cv_obs) if (etp is not None and cv_obs not in (None, 0, "")) else None

    limites = [x for x in (cv_lim_clia, cv_lim_vb) if x is not None]
    if cv_obs in (None, "") or not limites:
        cv_status = "-"
    else:
        cv_status = "OK" if cv_obs <= min(limites) else "NÃO"

    return dict(etp=etp, etp_fonte=fonte, etp_clia=etp_clia, etp_vb=etp_vb,
                es=es, ea=ea, et=et, sigma=sigma,
                cv_limite_clia=cv_lim_clia, cv_limite_vb=cv_lim_vb,
                cv_limite_fb=cv_lim_fb, cv_status=cv_status)


def build_stats(analyte, level, values, bias_obs, specs) -> AnalyteLevelStats:
    """Monta o objeto de estatísticas + sigma para um analito/nível."""
    n, m, sd, cv = compute_stats(values)
    s = AnalyteLevelStats(analyte=analyte, level=level, n=n,
                          mean_obs=m, sd_obs=sd, cv_obs=cv)
    if cv is not None:
        sig = compute_sigma(cv, bias_obs,
                            specs.get("clia_tea"), specs.get("cvw"),
                            specs.get("cvg"), specs.get("perf"),
                            etp_pct=specs.get("etp_pct"),
                            etp_source=specs.get("etp_source"),
                            fab_cv=specs.get("fab_cv"))
        for k, v in sig.items():
            setattr(s, k, v)
    return s


# =========================================================================== #
#  CONTROLE EXTERNO DA QUALIDADE (EQC) — cálculos puros, em PERCENTUAL
# =========================================================================== #
# Convenção de unidades: tudo em % (CV%=2,0; bias=0,93; ET=4,23; Sigma=4,535),
# idêntico ao exemplo da especificação. Fator do erro aleatório = 1,65 (EA_Z).
RANDOM_ERROR_FACTOR = EA_Z  # 1,65 — mesma constante usada no IQC


def _num(x):
    """Converte para float aceitando vírgula decimal; None se inválido/vazio/NaN."""
    if x in (None, "", "-"):
        return None
    if isinstance(x, str):
        x = x.replace(",", ".").strip()
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    if v != v:          # NaN (ex.: célula vazia de DataFrame)
        return None
    return v


def sample_bias_pct(lab_value, peer_mean):
    """
    Bias percentual de uma amostra (espécime) frente à média do grupo comparativo:
        Bias_% = ((Resultado_Lab − Média_Grupo) / Média_Grupo) × 100
    Retorna None (sem cálculo) se faltar dado ou a média do grupo for 0 — evita
    divisão por zero (regra de negócio 10).
    """
    lab = _num(lab_value)
    peer = _num(peer_mean)
    if lab is None or peer in (None, 0) or peer == 0.0:
        return None
    return ((lab - peer) / peer) * 100.0


def round_abs_bias(sample_biases):
    """
    |Bias| da RODADA = média aritmética dos |bias| das amostras válidas.
    Aceita lista de bias percentuais (com sinal) ou None. Retorna None se não
    houver nenhuma amostra válida.
    """
    abs_vals = [abs(b) for b in sample_biases if b is not None and b == b]  # b==b exclui NaN
    if not abs_vals:
        return None
    return sum(abs_vals) / len(abs_vals)


def annual_bias_indicator(round_abs_biases):
    """
    Indicador de Bias Anual Robusto = média dos |bias| das rodadas do ano.
        Indicador = média(|Bias_R1|, |Bias_R2|, ..., |Bias_Rn|)
    Aceita lista de |bias| por rodada (já em módulo) ignorando None.
    Retorna None se não houver rodadas válidas.
    """
    vals = [b for b in round_abs_biases if b is not None and b == b]  # exclui None/NaN
    if not vals:
        return None
    return sum(vals) / len(vals)


def external_total_error(cv_pct, bias_anual):
    """Erro Total (%) = CV% × 1,65 + Indicador_Bias_Anual. None se faltar entrada."""
    cv = _num(cv_pct)
    bias = _num(bias_anual)
    if cv is None or bias is None:
        return None
    return cv * RANDOM_ERROR_FACTOR + bias


def external_sigma(etp_pct, bias_anual, cv_pct):
    """
    Métrica Sigma = (ETp − Indicador_Bias_Anual) / CV%.
    None se CV% for 0/None ou faltarem ETp/bias (evita divisão por zero).
    """
    etp = _num(etp_pct)
    bias = _num(bias_anual)
    cv = _num(cv_pct)
    if etp is None or bias is None or cv in (None, 0) or cv == 0.0:
        return None
    return (etp - bias) / cv


def annual_performance(cv_pct, bias_anual, etp_pct, etp_source="-"):
    """
    Consolida o desempenho analítico anual de um ensaio integrando IQC (CV%) e
    EQC (Bias Anual): retorna dict com cv_pct, bias_anual, etp, etp_source, et,
    sigma e listas de pendências ('faltas') para indicar dados insuficientes.

    Todas as entradas/saídas em PERCENTUAL.
    """
    cv = _num(cv_pct)
    bias = _num(bias_anual)
    etp = _num(etp_pct)

    faltas = []
    if cv is None:
        faltas.append("CV% (Controle Interno)")
    if bias is None:
        faltas.append("Bias Anual (Controle Externo)")
    if etp is None:
        faltas.append("ETp (cadastro do ensaio)")

    et = external_total_error(cv, bias)
    sigma = external_sigma(etp, bias, cv)

    return dict(cv_pct=cv, bias_anual=bias, etp=etp,
                etp_source=etp_source if etp is not None else "-",
                et=et, sigma=sigma, faltas=faltas)


# --------------------------------------------------------------------------- #
# Limites do gráfico (Levey-Jennings)
# --------------------------------------------------------------------------- #
def control_limits(ref_mean, ref_sd):
    """Retorna dict com média e ±1/2/3 DP."""
    if ref_mean is None or ref_sd is None:
        return {}
    return {
        "-3s": ref_mean - 3 * ref_sd,
        "-2s": ref_mean - 2 * ref_sd,
        "-1s": ref_mean - ref_sd,
        "media": ref_mean,
        "+1s": ref_mean + ref_sd,
        "+2s": ref_mean + 2 * ref_sd,
        "+3s": ref_mean + 3 * ref_sd,
    }
