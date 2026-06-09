"""
seed_data.py
------------
Dados reais extraídos da planilha QC-INI (Instituto Nacional de Infectologia)
para popular o protótipo, mais um gerador de pontos de CQ de demonstração.
"""

import random
from datetime import datetime, timedelta

# Ordem oficial do painel hematológico (Sysmex XN-1000)
ANALYTES = ["WBC", "RBC", "HGB", "HCT", "MCV", "MCH", "MCHC", "PLT",
            "NEUT%", "LYMPH%", "MONO%", "EO%", "BASO%",
            "NEUT#", "LYMPH#", "MONO#", "EO#", "BASO#",
            "IG%", "IG#", "NRBC%", "NRBC#", "RDW-SD", "RDW-CV",
            "MPV", "RET%", "RET#", "IRF"]

# Média e DP de referência por (analito, nível) — valores reais do laboratório
REFERENCE = {
    ("WBC", 1): (3.0037931034482748, 0.052875446171940406), ("WBC", 2): (6.805172413793103, 0.091167847153916), ("WBC", 3): (16.523793103448273, 0.16925073946036923),
    ("RBC", 1): (2.306551724137931, 0.021921265290151055), ("RBC", 2): (4.376896551724138, 0.034959512120233176), ("RBC", 3): (5.122758620689655, 0.03452706513158902),
    ("HGB", 1): (5.5, 0.059761430466719466), ("HGB", 2): (12.27586206896552, 0.08304547985374039), ("HGB", 3): (15.565517241379315, 0.08139788549800775),
    ("HCT", 1): (16.63448275862069, 0.3096955441223865), ("HCT", 2): (35.8344827586207, 0.4521976986017179), ("HCT", 3): (44.9103448275862, 0.49881633292239813),
    ("MCV", 1): (72.14827586206897, 0.9006429011922422), ("MCV", 2): (81.87586206896552, 0.7390414508062798), ("MCV", 3): (87.66551724137935, 0.7102694971774365),
    ("MCH", 1): (23.83793103448276, 0.27440037627974), ("MCH", 2): (28.05517241379311, 0.20973828401903993), ("MCH", 3): (30.393103448275863, 0.23744327816307542),
    ("MCHC", 1): (33.01724137931035, 0.5782170659617875), ("MCHC", 2): (34.27241379310345, 0.37975231751774535), ("MCHC", 3): (34.65862068965517, 0.36986217643421004),
    ("PLT", 1): (70.62068965517241, 4.68595007216442), ("PLT", 2): (227.93103448275863, 10.875697135234134), ("PLT", 3): (516.2758620689655, 14.851398567436902),
    ("NEUT#", 1): (1.2193103448275862, 0.03788431964384275), ("NEUT#", 2): (2.9737931034482763, 0.06038056647533018), ("NEUT#", 3): (7.845517241379308, 0.1775746564707072),
    ("LYMPH#", 1): (1.0048275862068967, 0.04702740004167578), ("LYMPH#", 2): (1.9882758620689651, 0.04855548854640736), ("LYMPH#", 3): (4.037586206896552, 0.17793769319651906),
    ("MONO#", 1): (0.3462068965517241, 0.02920540472283565), ("MONO#", 2): (0.7727586206896552, 0.060762802985058095), ("MONO#", 3): (1.9575862068965515, 0.19830811474668464),
    ("EO#", 1): (0.2927586206896552, 0.016668308621910943), ("EO#", 2): (0.7279310344827586, 0.04074128393593861), ("EO#", 3): (1.8631034482758626, 0.12764637111611632),
    ("BASO#", 1): (0.1444827586206897, 0.005723514714723381), ("BASO#", 2): (0.3251724137931035, 0.008709883407113863), ("BASO#", 3): (0.7917241379310347, 0.021723746966534337),
    ("IG#", 1): (10.28965517241379, 0.3488192610038767), ("IG#", 2): (11.051724137931034, 0.32472041782692795), ("IG#", 3): (11.875862068965516, 0.3345094194849197),
    ("NRBC%", 1): (5.176206896551725, 0.5741292400117304), ("NRBC%", 2): (6.213793103448276, 0.26955537190374723), ("NRBC%", 3): (6.6068965517241365, 0.21369376366787565),
    ("NRBC#", 1): (0.1430206896551724, 0.029590133241026263), ("NRBC#", 2): (0.42241379310344834, 0.017658501071679954), ("NRBC#", 3): (1.0889655172413792, 0.032551588460934575),
    ("RDW-SD", 1): (49.97931034482758, 0.6043764852809799), ("RDW-SD", 2): (14.779310344827586, 0.08185051856379796), ("RDW-SD", 3): (47.2896551724138, 0.5582784939860129),
    ("RDW-CV", 1): (19.43448275862069, 0.16534781969482565), ("RDW-CV", 2): (9.83448275862069, 0.3003282276692877), ("RDW-CV", 3): (14.944827586206898, 0.09097176522946811),
    ("RET%", 1): (4.630689655172414, 0.2666449088003793), ("RET%", 2): (2.343793103448276, 0.06126137172007034), ("RET%", 3): (1.236206896551724, 0.051989957347786254),
    ("RET#", 1): (0.10683448275862067, 0.006742046947651357), ("RET#", 2): (0.10258275862068966, 0.0027378392112755705), ("RET#", 3): (0.06333103448275865, 0.0027089934145724764),
    ("IRF", 1): (36.46666666666667, 6.267565791627644), ("IRF", 2): (39.303448275862074, 7.061337528027445), ("IRF", 3): (33.52068965517241, 4.655747547029924),
}

# Especificações de qualidade por analito (TEa CLIA, CV fabricante, CVi/CVg EFLM, nível de desempenho)
SPECS = {
    "WBC": {"clia_tea": 0.07, "fab_cv": 0.03, "cvw": 10.8, "cvg": 16.4, "perf": "DES"},
    "RBC": {"clia_tea": 0.04, "fab_cv": 0.015, "cvw": 2.6, "cvg": 6.4, "perf": "MIN"},
    "HGB": {"clia_tea": 0.04, "fab_cv": 0.01, "cvw": 2.7, "cvg": 5.9, "perf": "ÓTI"},
    "HCT": {"clia_tea": 0.04, "fab_cv": 0.015, "cvw": 2.8, "cvg": 5.4, "perf": "MIN"},
    "MCV": {"clia_tea": None, "fab_cv": 0.01, "cvw": 0.8, "cvg": 3.7, "perf": "MIN"},
    "MCH": {"clia_tea": None, "fab_cv": 0.02, "cvw": 0.8, "cvg": 4.3, "perf": "MIN"},
    "MCHC": {"clia_tea": None, "fab_cv": 0.02, "cvw": 0.9, "cvg": 1.4, "perf": "MIN"},
    "PLT": {"clia_tea": 0.15, "fab_cv": 0.04, "cvw": 7.5, "cvg": 18.6, "perf": "MIN"},
    "NEUT%": {"clia_tea": None, "fab_cv": 0.08, "cvw": 14.0, "cvg": 23.5, "perf": "MIN"},
    "LYMPH%": {"clia_tea": None, "fab_cv": 0.08, "cvw": 10.8, "cvg": 22.6, "perf": "MIN"},
    "MONO%": {"clia_tea": None, "fab_cv": 0.2, "cvw": 13.3, "cvg": 22.2, "perf": "MIN"},
    "EO%": {"clia_tea": None, "fab_cv": 0.25, "cvw": 14.9, "cvg": 65.3, "perf": "MIN"},
    "BASO%": {"clia_tea": None, "fab_cv": 0.4, "cvw": 12.4, "cvg": 26.3, "perf": "MIN"},
    "NEUT#": {"clia_tea": None, "fab_cv": 0.12, "cvw": 14.0, "cvg": 23.5, "perf": "MIN"},
    "LYMPH#": {"clia_tea": None, "fab_cv": 0.08, "cvw": 10.8, "cvg": 22.6, "perf": "MIN"},
    "MONO#": {"clia_tea": None, "fab_cv": 0.2, "cvw": 13.3, "cvg": 22.2, "perf": "MIN"},
    "EO#": {"clia_tea": None, "fab_cv": 0.25, "cvw": 14.9, "cvg": 65.3, "perf": "MIN"},
    "BASO#": {"clia_tea": None, "fab_cv": 0.4, "cvw": 12.4, "cvg": 26.3, "perf": "MIN"},
    "IG%": {"clia_tea": None, "fab_cv": 0.12, "cvw": None, "cvg": None, "perf": "MIN"},
    "IG#": {"clia_tea": None, "fab_cv": None, "cvw": None, "cvg": None, "perf": "MIN"},
    "NRBC%": {"clia_tea": None, "fab_cv": 0.25, "cvw": None, "cvg": None, "perf": "MIN"},
    "NRBC#": {"clia_tea": None, "fab_cv": None, "cvw": None, "cvg": None, "perf": "MIN"},
    "RDW-SD": {"clia_tea": None, "fab_cv": 0.481, "cvw": 1.6, "cvg": 4.2, "perf": "MIN"},
    "RDW-CV": {"clia_tea": None, "fab_cv": 0.151, "cvw": None, "cvg": None, "perf": "MIN"},
    "MPV": {"clia_tea": None, "fab_cv": 0.094, "cvw": 2.2, "cvg": 7.0, "perf": "MIN"},
    "RET%": {"clia_tea": None, "fab_cv": 0.012, "cvw": 9.7, "cvg": 27.1, "perf": "MIN"},
    "RET#": {"clia_tea": None, "fab_cv": None, "cvw": None, "cvg": None, "perf": "MIN"},
    "IRF": {"clia_tea": None, "fab_cv": None, "cvw": None, "cvg": None, "perf": "MIN"},
}

CONFIG = {
    "instituicao": "Instituto Nacional de Infectologia Evandro Chagas",
    "setor": "Hematologia",
    "responsavel_tecnico": "EDSON BEYKER DE MENDONÇA",
    "analista_supervisor": "EDSON BEYKER DE MENDONÇA / VÍTOR SANTOS DA SILVA",
    "equipamento": "XN-1000 / SYSMEX",
    "numero_serie": "18567",
    "controle": "XN-CHECK",
    "autor": "VÍTOR SANTOS DA SILVA",
}

# Usuários de demonstração. access_level: 1 = TÉCNICO, 2 = ANALISTA (pode assinar/validar NC)
# (senha em sha256; o login de teste é INILAB / TESTE05)
USERS = [
    {"login": "INILAB", "name": "Usuário de Teste", "role": "ANALISTA", "access_level": 2, "password": "TESTE05"},
    {"login": "admin", "name": "Administrador", "role": "ANALISTA", "access_level": 2, "password": "admin"},
]

# Lotes de controle de exemplo por nível
LOTES = {1: "QC-52261101", 2: "QC-52261102", 3: "QC-52261103"}
VALIDADE = "2026-03-31"


def gerar_dados_demo(seed: int = 7, n: int = 32):
    """
    Gera pontos de CQ de demonstração para os 3 níveis a partir das médias/DP reais.
    Injeta propositalmente violações em alguns analitos para que o motor de Westgard
    e o painel mostrem alertas e rejeições (efeito didático).

    Retorna lista de dicts prontos para inserir na tabela `results`.
    """
    rng = random.Random(seed)
    linhas = []
    data0 = datetime(2025, 11, 1, 8, 0, 0)

    for level in (1, 2, 3):
        for i in range(n):
            seq = i + 1
            dt = data0 + timedelta(days=i)
            for a in ANALYTES:
                ref = REFERENCE.get((a, level))
                if not ref:
                    continue
                m, sd = ref
                val = rng.gauss(m, sd if sd else m * 0.01)

                # --- violações didáticas ---
                # WBC nível 2: um pico 1-3s no ponto 20
                if a == "WBC" and level == 2 and seq == 20:
                    val = m + 3.4 * sd
                # RBC nível 1: deriva 3-1s (pontos 14-16 acima de +1s)
                if a == "RBC" and level == 1 and seq in (14, 15, 16):
                    val = m + 1.6 * sd
                # PLT nível 3: tendência 9x (pontos 5-13 acima da média)
                if a == "PLT" and level == 3 and 5 <= seq <= 13:
                    val = m + rng.uniform(0.3, 0.9) * sd
                # R4s entre níveis na rodada 25 do HGB (N1 alto, N2 baixo)
                if a == "HGB" and seq == 25 and level == 1:
                    val = m + 2.3 * sd
                if a == "HGB" and seq == 25 and level == 2:
                    val = m - 2.3 * sd

                linhas.append({
                    "level": level, "lote": LOTES[level], "seq": seq,
                    "run_date": dt.strftime("%Y-%m-%d"),
                    "run_time": dt.strftime("%H:%M:%S"),
                    "analyte": a, "value": round(val, 4),
                })
    return linhas


# Controle externo (proficiência/CAP) de exemplo -> alimenta o viés/Sigma
# (analito, nível): (valor_INI, valor_par, dp_par)
EXTERNAL_QC = {
    ("WBC", 1): (3.05, 3.00, 0.10), ("WBC", 2): (6.95, 6.80, 0.20),
    ("HGB", 1): (5.58, 5.50, 0.08), ("HGB", 2): (12.40, 12.28, 0.15),
    ("PLT", 2): (231.0, 228.0, 8.0), ("RBC", 1): (2.33, 2.31, 0.05),
}
