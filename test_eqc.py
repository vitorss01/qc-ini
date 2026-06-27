"""
test_eqc.py — Testes do módulo de Controle Externo (EQC) + integração com o IQC.
Usa um banco SQLite temporário (via QCLAB_DB) para não tocar o banco real.
Execute:  python test_eqc.py
"""
import sys
import os
import tempfile

sys.path.insert(0, "qclab")
os.environ["QCLAB_DB"] = os.path.join(tempfile.gettempdir(), "qctest_eqc_suite.db")

import database as db          # noqa: E402
import qc_engine as qc         # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")   # evita erro de console no Windows
except Exception:
    pass

total = passed = 0


def test(label, cond):
    global total, passed
    total += 1
    status = "PASS" if cond else "FAIL !!!"
    print(f"  [{status}]  {label}")
    if cond:
        passed += 1


def approx(a, b, tol=1e-6):
    return a is not None and b is not None and abs(a - b) <= tol


# Banco limpo, SEM demo (controle total dos dados de teste)
db.init_db(force=True, demo=False)


# ===========================================================================
print("=" * 64)
print("FUNÇÕES PURAS (motor de cálculos, em %)")
print("=" * 64)

test("sample_bias_pct(102, 100) = +2.0", approx(qc.sample_bias_pct(102, 100), 2.0))
test("sample_bias_pct(98, 100) = -2.0", approx(qc.sample_bias_pct(98, 100), -2.0))
test("sample_bias_pct(5, 0) = None (div/zero)", qc.sample_bias_pct(5, 0) is None)
test("sample_bias_pct(5, None) = None", qc.sample_bias_pct(5, None) is None)
test("sample_bias_pct(None, 100) = None", qc.sample_bias_pct(None, 100) is None)

# |bias| da rodada = média dos |bias| das amostras (5 amostras, sinais variados)
biases5 = [qc.sample_bias_pct(lab, 100) for lab in (100.3, 99.5, 100.4, 100.0, 99.8)]
# |0.3|,|−0.5|,|0.4|,|0|,|−0.2| -> média = 0.28
test("round_abs_bias de 5 amostras = 0.28", approx(qc.round_abs_bias(biases5), 0.28))
test("round_abs_bias([0.3,-0.5,0.4]) = 0.4", approx(qc.round_abs_bias([0.3, -0.5, 0.4]), 0.4))
test("round_abs_bias([]) = None", qc.round_abs_bias([]) is None)
test("round_abs_bias ignora NaN", approx(qc.round_abs_bias([float("nan"), 0.5]), 0.5))

test("annual_bias_indicator([0.4,1.8,0.6]) = 0.9333",
     approx(qc.annual_bias_indicator([0.4, 1.8, 0.6]), 2.8 / 3))
test("annual_bias_indicator([]) = None", qc.annual_bias_indicator([]) is None)

test("external_total_error(2.0, 0.93) = 4.23", approx(qc.external_total_error(2.0, 0.93), 4.23))
test("external_total_error(None, 0.93) = None", qc.external_total_error(None, 0.93) is None)

test("external_sigma(10, 0.93, 2.0) = 4.535", approx(qc.external_sigma(10, 0.93, 2.0), 4.535))
test("external_sigma cv=0 -> None (div/zero)", qc.external_sigma(10, 0.93, 0) is None)
test("external_sigma sem ETp -> None", qc.external_sigma(None, 0.93, 2.0) is None)

perf = qc.annual_performance(2.0, 0.93, 10, "CLIA")
test("annual_performance ET = 4.23", approx(perf["et"], 4.23))
test("annual_performance Sigma = 4.535", approx(perf["sigma"], 4.535))
test("annual_performance sem faltas", perf["faltas"] == [])
test("annual_performance sem CV% -> falta + sigma None",
     qc.annual_performance(None, 0.93, 10)["sigma"] is None
     and "CV% (Controle Interno)" in qc.annual_performance(None, 0.93, 10)["faltas"])
test("annual_performance sem ETp -> falta + sigma None",
     qc.annual_performance(2.0, 0.93, None)["sigma"] is None)


# ===========================================================================
print("=" * 64)
print("INTEGRAÇÃO COM BANCO (rodadas, limite, filtros, auditoria)")
print("=" * 64)


def amostras_alvo(peer, biases):
    return [{"sample_label": f"{i:02d}", "lab_value": peer * (1 + b / 100.0),
             "peer_mean": peer} for i, b in enumerate(biases, 1)]


# Cadastro de uma rodada
rid, err = db.eqc_add_round(
    {"area": "Bioquímica", "analyte": "GLI_TESTE", "year": 2024, "round_number": 1,
     "round_date": "2024-03-01", "provider": "Controllab", "unit": "mg/dL",
     "lote": "L1", "status": "Aceitável"},
    amostras_alvo(100.0, [0.3, -0.5, 0.4]))
test("cadastro de rodada retorna id", rid is not None and err is None)
test("rodada cadastrada tem 3 amostras", len(db.eqc_get_samples(rid)) == 3)
test("amostras guardam bias_pct", db.eqc_get_samples(rid)[0]["bias_pct"] is not None)

# |bias| da rodada salvo bate com o cálculo (0.4)
lst = db.eqc_list_rounds(year=2024, analyte="GLI_TESTE")
test("|bias| da rodada salva = 0.4", approx(lst[0]["round_abs_bias"], 0.4))

# Até 6 rodadas/ano e bloqueio na 7ª
for n in range(2, 7):
    db.eqc_add_round(
        {"area": "Bioquímica", "analyte": "LIM_TESTE", "year": 2024, "round_number": n - 1,
         "provider": "CAP"}, amostras_alvo(100.0, [1.0]))
# já temos rodadas 1..5; adiciona a 6ª
db.eqc_add_round({"area": "Bioquímica", "analyte": "LIM_TESTE", "year": 2024,
                  "round_number": 6, "provider": "CAP"}, amostras_alvo(100.0, [1.0]))
test("6 rodadas cadastradas", db.eqc_round_count("LIM_TESTE", 2024) == 6)
rid7, err7 = db.eqc_add_round({"area": "Bioquímica", "analyte": "LIM_TESTE", "year": 2024,
                               "round_number": 3, "provider": "CAP"}, amostras_alvo(100.0, [1.0]))
test("7ª rodada bloqueada (limite 6)", rid7 is None and err7 is not None)
test("rodada nº 7 rejeitada (1..6)",
     db.eqc_add_round({"analyte": "OUTRO", "year": 2024, "round_number": 7},
                      amostras_alvo(100.0, [1.0]))[0] is None)
test("rodada nº duplicado rejeitado",
     db.eqc_add_round({"analyte": "GLI_TESTE", "year": 2024, "round_number": 1},
                      amostras_alvo(100.0, [1.0]))[0] is None)

# Limpa as tabelas EQC/IQC (sem apagar o arquivo — evita lock no Windows) para
# os testes de indicador/filtros partirem de um estado conhecido.
_c = db.get_conn()
_c.executescript("DELETE FROM eqc_samples; DELETE FROM eqc_rounds; "
                 "DELETE FROM eqc_audit; DELETE FROM results;")
_c.commit()

# Indicador anual robusto: 3 rodadas com |bias| 0.4 / 1.8 / 0.6 -> 0.9333
for n, bs in zip((1, 2, 3), ([0.3, -0.5, 0.4], [2.0, -1.5, 1.9], [0.5, -0.7, 0.6])):
    db.eqc_add_round({"area": "Hematologia", "analyte": "WBC", "year": 2025,
                      "round_number": n, "provider": "CAP", "lote": f"L{n}"},
                     amostras_alvo(6.8, bs))
ind = db.eqc_annual_indicator("WBC", 2025)
test("Indicador de Bias Anual (WBC/2025) = 0.9333", approx(ind["indicador"], 2.8 / 3))
test("Indicador usa 3 rodadas", ind["n_rodadas"] == 3)

# Atribuição ao ensaio correto (outra analito não interfere)
db.eqc_add_round({"area": "Bioquímica", "analyte": "Creatinina", "year": 2025,
                  "round_number": 1, "provider": "CAP"}, amostras_alvo(1.0, [5.0]))
test("ensaio correto: WBC continua com 3 rodadas", db.eqc_round_count("WBC", 2025) == 3)
test("ensaio correto: Creatinina com 1 rodada", db.eqc_round_count("Creatinina", 2025) == 1)

# Filtros
test("filtro por ano", len(db.eqc_list_rounds(year=2025)) == 4)
test("filtro por analito", len(db.eqc_list_rounds(analyte="WBC")) == 3)
test("filtro por área", len(db.eqc_list_rounds(area="Bioquímica")) == 1)
test("filtro por rodada", len(db.eqc_list_rounds(round_number=1)) == 2)
test("filtro por lote", len(db.eqc_list_rounds(lote="L2")) == 1)
test("filtro por provedor", len(db.eqc_list_rounds(provider="CAP")) == 4)

# Exclusão
algum = db.eqc_list_rounds(analyte="Creatinina")[0]["round_id"]
db.eqc_delete_round(algum)
test("exclusão remove a rodada", db.eqc_round_count("Creatinina", 2025) == 0)
test("auditoria registrou ações",
     db.get_conn().execute("SELECT COUNT(*) c FROM eqc_audit").fetchone()["c"] > 0)


# ===========================================================================
print("=" * 64)
print("INTEGRAÇÃO IQC↔EQC (CV% do interno + ET/Sigma)")
print("=" * 64)

# Insere resultados IQC para WBC nível 1 e calcula CV%
for i, v in enumerate([6.78, 6.80, 6.82, 6.79, 6.81, 6.83], 1):
    db.insert_result(1, "L", i, "2025-01-0%d" % i, "08:00:00", "WBC", v)
vals = [r["value"] for r in db.get_results("WBC", 1)]
_n, _m, _sd, cv = qc.compute_stats(vals)
cv_pct = cv * 100.0
test("CV% do IQC calculado (>0)", cv_pct is not None and cv_pct > 0)

# Cenário do exemplo: força bias=0.93, etp=10, cv=2.0
perf2 = qc.annual_performance(2.0, ind["indicador"], 10, "CLIA")
test("ET com bias anual real (cv=2,0) > 0", perf2["et"] is not None and perf2["et"] > 0)

# Sem CV% (analito sem resultados IQC)
test("sem resultados IQC -> CV% None",
     not [r["value"] for r in db.get_results("Creatinina", 1)])

# peer_mean inválido não quebra (rodada só com amostras válidas)
rid_z, _ = db.eqc_add_round(
    {"area": "Bioquímica", "analyte": "ZERO_TESTE", "year": 2025, "round_number": 1},
    [{"sample_label": "01", "lab_value": 5, "peer_mean": 0},
     {"sample_label": "02", "lab_value": 102, "peer_mean": 100}])
lst_z = db.eqc_list_rounds(analyte="ZERO_TESTE")
test("peer_mean=0 ignorado, rodada usa só amostra válida (|bias|=2.0)",
     approx(lst_z[0]["round_abs_bias"], 2.0))


# ===========================================================================
print()
print("=" * 64)
print(f"RESULTADO FINAL:  {passed}/{total} testes passaram")
print("=" * 64)
print("Tudo certo!" if passed == total else "ATENÇÃO: há falhas a corrigir.")
