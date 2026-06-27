"""
test_westgard.py  –  Testes unitários das Regras de Westgard
Execute:  python test_westgard.py
"""
import sys
sys.path.insert(0, "qclab")
import qc_engine as qc

PASS = "PASS"
FAIL = "FAIL !!!"

def chk(label, condition):
    status = PASS if condition else FAIL
    print(f"  [{status}]  {label}")
    return condition

total = 0
passed = 0

def test(label, condition):
    global total, passed
    total += 1
    if chk(label, condition):
        passed += 1

# ===========================================================================
# HEMATOLOGIA – 3 níveis
# ===========================================================================
print("=" * 60)
print("HEMATOLOGIA – 3 Níveis")
print("=" * 60)

refs3 = {1: (100.0, 10.0), 2: (100.0, 10.0), 3: (100.0, 10.0)}

# ── 1-3S ──────────────────────────────────────────────────────────────────
s = {1: [(1, 102.0)], 2: [(1, 135.0)], 3: [(1, 98.0)]}  # L2: Z=+3.5
r = qc.evaluate_levels(s, refs3, setor="Hematologia")
test("1-3S  – L2 Z=+3.5 levanta flag em L2",   "1-3S" in r[2][0].flags)
test("1-3S  – L2 Z=+3.5 status REJEITADO",      r[2][0].status == "REJEITADO")
test("1-3S  – L1 Z=+0.2 NÃO levanta flag",      "1-3S" not in r[1][0].flags)

# ── 2of3-2S (intra-corrida, mesma direção) ────────────────────────────────
# Dois positivos
s = {1: [(1, 125.0)], 2: [(1, 123.0)], 3: [(1, 101.0)]}  # L1=+2.5, L2=+2.3, L3=+0.1
r = qc.evaluate_levels(s, refs3, setor="Hematologia")
test("2of3-2S – L1=+2.5, L2=+2.3 (pos) levanta flag L1",  "2of3-2S" in r[1][0].flags)
test("2of3-2S – L1=+2.5, L2=+2.3 (pos) levanta flag L2",  "2of3-2S" in r[2][0].flags)
test("2of3-2S – L3=+0.1 também sinalizado (corrida)",      "2of3-2S" in r[3][0].flags)

# Dois negativos
s = {1: [(1, 75.0)], 2: [(1, 77.0)], 3: [(1, 99.0)]}   # L1=-2.5, L2=-2.3, L3=-0.1
r = qc.evaluate_levels(s, refs3, setor="Hematologia")
test("2of3-2S – L1=-2.5, L2=-2.3 (neg) levanta flag",     "2of3-2S" in r[1][0].flags)

# Sem violação: direções opostas
s = {1: [(1, 125.0)], 2: [(1, 77.0)], 3: [(1, 100.0)]}   # +2.5 e -2.3 (R4S, não 2of3)
r = qc.evaluate_levels(s, refs3, setor="Hematologia")
test("2of3-2S – direções opostas NÃO viola 2of3-2S", "2of3-2S" not in r[1][0].flags)

# ── R4S Hematologia ───────────────────────────────────────────────────────
s = {1: [(1, 125.0)], 2: [(1, 79.0)], 3: [(1, 100.0)]}   # amp=4.6
r = qc.evaluate_levels(s, refs3, setor="Hematologia")
test("R4S   – amp=4.6 (L1=+2.5, L2=-2.1) viola R4S",  "R4S" in r[1][0].flags)

s = {1: [(1, 121.0)], 2: [(1, 81.0)], 3: [(1, 100.0)]}   # amp=4.0 (não viola: <=4)
r = qc.evaluate_levels(s, refs3, setor="Hematologia")
test("R4S   – amp=4.0 NÃO viola R4S",                 "R4S" not in r[1][0].flags)

# ── 3-1S Hematologia ──────────────────────────────────────────────────────
s = {1: [(1,111.0),(2,112.0),(3,113.0)],
     2: [(1,100.0),(2,100.0),(3,100.0)],
     3: [(1,100.0),(2,100.0),(3,100.0)]}
r = qc.evaluate_levels(s, refs3, setor="Hematologia")
test("3-1S  – 3x Z>1 em L1, ponto 3 tem flag",        "3-1S" in r[1][2].flags)
test("3-1S  – ponto 2 ainda sem flag (só 2 na janela)","3-1S" not in r[1][1].flags)

# ── 6X Hematologia ────────────────────────────────────────────────────────
s = {1: [(i, 100.5 + i*0.1) for i in range(1,7)],  # sempre Z>0
     2: [(i, 100.0) for i in range(1,7)],
     3: [(i, 100.0) for i in range(1,7)]}
r = qc.evaluate_levels(s, refs3, setor="Hematologia")
test("6X    – 6x Z>0 em L1, ponto 6 tem flag",        "6X" in r[1][5].flags)
test("6X    – ponto 5 ainda sem flag (só 5 na janela)","6X" not in r[1][4].flags)

# ===========================================================================
# BIOQUÍMICA – 2 níveis
# ===========================================================================
print()
print("=" * 60)
print("BIOQUÍMICA – 2 Níveis")
print("=" * 60)

refs2 = {1: (100.0, 10.0), 2: (100.0, 10.0)}

# ── 1-3S Bioquímica ───────────────────────────────────────────────────────
s = {1: [(1, 100.0)], 2: [(1, 132.0)]}   # L2 Z=+3.2
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("1-3S  – L2 Z=+3.2 viola 1-3S",     "1-3S" in r[2][0].flags)

# ── 2-2S intra-corrida ────────────────────────────────────────────────────
s = {1: [(1, 122.0)], 2: [(1, 123.0)]}   # ambos >+2
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("2-2S  – intra (L1=+2.2, L2=+2.3 mesmo lado)",  "2-2S" in r[1][0].flags)

s = {1: [(1, 78.0)], 2: [(1, 77.0)]}     # ambos <-2
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("2-2S  – intra negativo (L1=-2.2, L2=-2.3)",     "2-2S" in r[1][0].flags)

s = {1: [(1, 122.0)], 2: [(1, 78.0)]}    # lados opostos → não viola 2-2S
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("2-2S  – intra lados opostos NÃO viola 2-2S",    "2-2S" not in r[1][0].flags)

# ── 2-2S inter-corrida (mesmo nível, t e t-1) ─────────────────────────────
s = {1: [(1, 125.0), (2, 123.0)], 2: [(1, 100.0), (2, 100.0)]}
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("2-2S  – inter (L1 seq1=+2.5, seq2=+2.3)",       "2-2S" in r[1][1].flags)
test("2-2S  – inter NÃO marca ponto anterior (seq1)",  "2-2S" not in r[1][0].flags)

# ── R4S Bioquímica ────────────────────────────────────────────────────────
s = {1: [(1, 125.0)], 2: [(1, 79.0)]}   # +2.5 e -2.1, amp=4.6, opostos
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("R4S   – |Z_L1-Z_L2|=4.6 e sinais opostos",      "R4S" in r[1][0].flags)

s = {1: [(1, 125.0)], 2: [(1, 125.0)]}  # mesma direção, amp=0 → não viola
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("R4S   – mesma direção NÃO viola R4S",            "R4S" not in r[1][0].flags)

# ── 4-1S intra-nível ─────────────────────────────────────────────────────
s = {1: [(i, 111.0 + i*0.5) for i in range(1,5)],  # Z>1 por 4x
     2: [(i, 100.0) for i in range(1,5)]}
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("4-1S  – 4x Z>1 em L1, ponto 4 tem flag",        "4-1S" in r[1][3].flags)
test("4-1S  – ponto 3 ainda sem flag",                 "4-1S" not in r[1][2].flags)

# ── 4-1S inter-nível (últimos 2 de L1 E últimos 2 de L2 >1, mesmo lado) ──
s = {1: [(1, 112.0), (2, 113.0)],   # Z=+1.2, +1.3
     2: [(1, 111.0), (2, 114.0)]}   # Z=+1.1, +1.4
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("4-1S  – inter-nivel 2+2 pontos >+1 mesmo lado", "4-1S" in r[1][1].flags)

# ── 8X intra-nível ────────────────────────────────────────────────────────
s = {1: [(i, 100.5 + i*0.05) for i in range(1,9)],   # Z>0 por 8x
     2: [(i, 100.0) for i in range(1,9)]}
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("8X    – 8x Z>0 em L1, ponto 8 tem flag",        "8X" in r[1][7].flags)
test("8X    – ponto 7 ainda sem flag",                 "8X" not in r[1][6].flags)

# ── 8X inter-nível (últimos 4 de L1 E últimos 4 de L2, mesmo lado) ────────
s = {1: [(i, 100.5 + i*0.05) for i in range(1,5)],   # Z>0 por 4x
     2: [(i, 100.3 + i*0.02) for i in range(1,5)]}   # Z>0 por 4x
r = qc.evaluate_levels(s, refs2, setor="Bioquimica")
test("8X    – inter-nivel 4+4 pontos Z>0 mesmo lado", "8X" in r[1][3].flags)

# ── Inferência automática por nº de níveis ────────────────────────────────
print()
print("=" * 60)
print("INFERÊNCIA AUTOMÁTICA DE SETOR")
print("=" * 60)
s3 = {1: [(1,125.0)], 2: [(1,123.0)], 3: [(1,101.0)]}
r_auto = qc.evaluate_levels(s3, refs3)   # setor omitido, 3 níveis → Hematologia
test("Inferência – 3 níveis usa regra 2of3-2S (Hemat)", "2of3-2S" in r_auto[1][0].flags)

s2 = {1: [(1,125.0)], 2: [(1,123.0)]}
r_auto2 = qc.evaluate_levels(s2, refs2)  # setor omitido, 2 níveis → Bioquímica
test("Inferência – 2 níveis usa regra 2-2S (Bioq)",    "2-2S" in r_auto2[1][0].flags)

# ===========================================================================
print()
print("=" * 60)
print(f"RESULTADO FINAL:  {passed}/{total} testes passaram")
print("=" * 60)
if passed == total:
    print("Tudo certo! Lógica de Westgard validada.")
else:
    print("ATENÇÃO: Há falhas a corrigir.")
