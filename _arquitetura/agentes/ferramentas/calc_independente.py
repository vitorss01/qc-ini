# -*- coding: utf-8 -*-
"""
calc_independente.py -- calculadora INDEPENDENTE de bias, ET, Sigma, media, DP e CV do QC_INI.

Autor: AGENTE_USUARIO_SIMULADO (frente F3, ciclo 2026-10-31_ciclo01). Reutilizavel nos proximos ciclos.

O QUE FAZ
    Le uma COPIA de QC_Bioquimica.xlsm (ou outro produto com o mesmo layout da Bioquimica) com openpyxl,
    SEM Excel e SEM ler as formulas do QC_INI para calcular. Usa so:
      - dados brutos: tblCQ_Final (aba "Principal - Resultados"), EQA_Base, catalogo/biblioteca da aba
        "Especificacoes", anos de referencia do bloco do lote na aba "Analitos";
      - celulas de ENTRADA do usuario (filtros): Estatistica B4/B5/D5 (visao e periodo), H4 (lote),
        N4 (ano EP), P4 (rodada), L11 (meses do Sigma), V8:W12 (exclusoes), Painel L4 (equipamento),
        Especificacoes AR (provedor do CEQ por analito).
    Recalcula em Python puro (math/statistics) e compara com o valor que o QC_INI calculou e GRAVOU
    (data_only=True = valor em cache da ultima recalculacao do Excel, precisao total do Value2).

DEFINICAO USADA (conferida no codigo e nos ADR em 31/10/2026 -- reconfirme a cada ciclo):
    CIQ do periodo (Estatistica C:F; mEstatPeriodo.EstatPeriodo/Agregar):
        PARTICIPA_ESTATISTICA = SIM; EQUIPAMENTO = Painel!L4; DATA_HORA truncada ao dia dentro da janela e
        fora das exclusoes; RESULTADO numerico; LOTE = Estatistica!H4; DP amostral (n-1); CV = DP/media*100.
    Bias do CEQ (Estatistica G/T/AC/AD; mCEQ.BiasEQ, ADR-035/071/072):
        EQA_Base, Analito_Canonico = analito, Uso_Analitico <> NAO, Provedor = Especificacoes!AR do analito,
        ano: com provedor e ano explicitos, modo TODAS/rotulo usa EXATAMENTE o ano (sem dado = SEM EP);
        Ano = TODOS -> todas as rodadas de todos os anos. bias_i = (Resultado - Alvo)/Alvo*100 (aqui
        RECALCULADO de Resultado e Valor_Alvo, nao lido da coluna Bias). Duas etapas: media dos |bias| de
        cada rodada (ano|rodada), depois media das medias. Sinal: mesma consolidacao com o bias com sinal.
    ETp (Especificacoes S/T, ADR-069): CLIA -> ETp CLIA do ano de referencia do lote; VB -> 1,65*f*CVi +
        f/2*raiz(CVi^2+CVg^2), f = 0,25 (OTI) / 0,50 (DES) / 0,75 (MIN); FAB -> ETp FAB. Vazio ou <= 0 = "-".
        CVTp: CLIA -> ETp/3; VB -> f*CVi; FAB -> CVTp FAB.
    ET (Estatistica H) = 1,65*CV(periodo, lote) + |bias|.
    CV pooled do Sigma (mIncerteza.IncertezaCIQ "CV", ADR-064/071): janela [EDATE(MU_Fim,-meses)+1, MU_Fim],
        MU_Fim = maior DATA com PARTICIPA=SIM (toda a tabela); mesmo equipamento; lotes com n >= 20 e media
        <> 0; CVp = raiz(soma((n_i-1)*CV_i^2)/soma(n_i-1)).
    Sigma (Estatistica L) = (ETp - |bias|)/CVp.  Margem N = ETp - ET; O = N/ETp*100.
    DPM (V) = (1 - Phi(Sigma - 1,5))*1e6; Rendimento (W) = 100 - DPM/1e4 (ADR-035).
    Classe (M): >=6 Classe mundial; 5-6 Excelente; 4-5 Bom; 3-4 Marginal; <3 Desempenho inadequado.
    Vies com sinal e RMS do CEQ (AR/AT; mCEQ.ViesEQ, ADR-064): so amostras avaliadas (Status_Padronizado <>
        NAO AVALIADO), resultado e alvo <> 0, cada chave uma vez; ano vigente, e o anterior se nao houver
        >= 6 amostras de >= 2 rodadas.
    u(Rw) 12 m (AF): o mesmo CV pooled numa janela de 12 meses.

HEMATOLOGIA (acrescentado na F3/Hematologia do mesmo ciclo; classes DadosHema/CalculadoraHema, auditar_hema)
    Detectada pelo titulo da aba Estatistica (A1 contem "Hematologia") ou por --produto hema. Diferencas
    conferidas no codigo (src/hema/mEstatistica.EstatBasica/AtualizarEstatisticaAba, formulas da Estatistica):
      - Estatistica C:F: lote = loteAnalise (Painel!H3, ou Configuracao!C20 com H3 vazio), equipamento =
        Cfg_Integracao!B11 (selEquipamento), PARTICIPA = SIM, ANO de DATA_HORA entre B3 e D3 (sem janela de dias
        e sem exclusoes);
      - provedor do CEQ unico em Estatistica!R4; Ano EP N4; Rodada P4;
      - ETp/CVTp pela aba Especificacoes da Hematologia (K/M/N/O/P da biblioteca do ano CLIA/VB do lote, R = VB,
        T = ETp em uso, U = CVTp, V = K/3, W = fi(Q)*CVi); ETp ausente = "-";
      - colunas comparadas: C..W, AC, AD, AF..AI (u(Rw) 12 m, gl, lotes, dias), AR, AU; e o Painel do analito
        selecionado (B7:I9, C45:C47, J/M/O45:47, C/H61:63, E74:E76).

CONFERENCIAS DESTA CALCULADORA (implementacoes escritas do zero, que NAO a importam)
    - conferencia_cruzada_hema.py: WBC, RBC, HCT, MCV, PLT (C, F, G, H, L; ETp lido da Estatistica).
    - segunda_opiniao_hema.py (01/11/2026): HGB, MCH, MCHC, NEUT%, MONO%, LYMPH#, EO#, MPV, RDW-CV (C..H, I..L, N, O,
      T, V, AC, AD, AF), com o ETp recalculado da biblioteca anual e o DPM por erfc.
    Na terceira passada da F3/Hematologia (01/11/2026, producao 0b45fa270d2a) esta calculadora deu de novo 84 x 27
    e 45 celulas do Painel sem divergencia, com JSON identico ao da segunda passada.
    Com python -I o PYTHONIOENCODING e ignorado: se a saida der UnicodeEncodeError (nome de analito com acento no
    console cp1252), rode sem -I com PYTHONIOENCODING=utf-8 ou redirecione para arquivo.

USO
    set PYTHONIOENCODING=utf-8
    python calc_independente.py <copia.xlsm> [--json saida.json] [--tol 1e-9] [--produto auto|bio|hema]
    Sem argumento extra imprime o resumo e a tabela de comparacao. Nunca abra a producao: copie antes
    (shutil.copy2 para %TEMP%\\qcw\\auditoria\\<ciclo>\\).

    Como modulo: from calc_independente import Dados, Calculadora; d = Dados(caminho); c = Calculadora(d)
    c.ciq_stats(...), c.bias_ceq(...), c.cv_pooled(...), c.etp(...), c.linha_estatistica(...).

LIMITACOES CONHECIDAS
    - Visoes MENSAL/TRIMESTRAL/SEMESTRAL/ANUAL do periodo seguem o calendario civil (mPeriodo); so
      PERSONALIZADO foi conferido contra o QC_INI no ciclo 2026-10-31_ciclo01.
    - Os modos de rodada ADR-071 (conjunto, ULTIMAS n, ACUMULADAS) estao implementados, mas so TODAS foi
      exercitado contra o QC_INI neste ciclo.
    - Le o valor em cache: se o Excel nao recalculou depois da ultima atualizacao, a divergencia aparece
      aqui como diferenca (e e isso mesmo que se quer ver).
    - Hematologia: o status SDI (R) so foi implementado para Rodada = TODAS; a coluna S (limites), o plano de CQ
      (X:AA), a incerteza expandida (AJ:AQ), a triagem (AS) e o texto AT nao sao comparados. O Painel e conferido
      so para o analito selecionado em C3 e para o filtro De/Ate gravado (G3/G4). A coluna V (DPM) usa tolerancia
      absoluta de 1e-9 DPM (TOL_ABS_HEMA), porque 1 - Phi perde digitos relativos com Sigma alto.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import statistics
import sys
from collections import OrderedDict, defaultdict

import openpyxl

ABA_CQ = "Principal - Resultados"
ABA_EQA = "EQA_Base"
ABA_ESP = "Especificações"
ABA_EST = "Estatística"
ABA_PAINEL = "Painel"
ABA_ANALITOS = "Analitos"
FATOR_VB = {"ÓTI": 0.25, "OTI": 0.25, "DES": 0.50, "MIN": 0.75}
SEM_EP = "SEM EP"


# --------------------------------------------------------------------------- utilidades
def _num(v):
    """Numero de verdade (int/float, nao bool) ou texto numerico; senao None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return None
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        try:
            return float(s.replace(",", ".")) if s.count(",") == 1 and "." not in s else float(s)
        except ValueError:
            return None
    return None


def _txt(v):
    return "" if v is None else str(v).strip()


def _dia(v):
    """Data truncada ao dia (date) ou None."""
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    n = _num(v)
    if n is not None and n > 0:  # serial do Excel
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(n))).date()
    return None


def edate(d: dt.date, meses: int) -> dt.date:
    """EDATE do Excel: soma meses e prende no ultimo dia do mes."""
    m = d.month - 1 + meses
    a = d.year + m // 12
    m = m % 12 + 1
    for dia in (d.day, 30, 29, 28):
        try:
            return dt.date(a, m, min(d.day, dia))
        except ValueError:
            continue
    raise ValueError(d)


def phi(z: float) -> float:
    return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))


def dif_rel(a, b):
    """Diferenca relativa |a-b|/max(|a|,|b|,1e-300); None se algum nao for numero."""
    if a is None or b is None or isinstance(a, str) or isinstance(b, str):
        return None
    m = max(abs(a), abs(b), 1e-300)
    return abs(a - b) / m


# --------------------------------------------------------------------------- leitura
class Dados:
    """Le o que a calculadora precisa de uma copia do .xlsm (openpyxl, read_only, data_only)."""

    def __init__(self, caminho: str):
        self.caminho = caminho
        self.wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True, keep_vba=False)
        self._ler_cq()
        self._ler_eqa()
        self._ler_esp()
        self._ler_filtros()
        self._ler_anos_lote()

    # tblCQ_Final: cabecalho na linha 6
    def _ler_cq(self):
        ws = self.wb[ABA_CQ]
        it = ws.iter_rows(min_row=6, values_only=True)
        cab = [(_txt(c)) for c in next(it)]
        self.cq_cab = cab
        ix = {n: i for i, n in enumerate(cab) if n}
        self.cq = []
        for r in it:
            if r is None or all(v is None for v in r):
                continue
            self.cq.append({k: r[i] if i < len(r) else None for k, i in ix.items()})

    # EQA_Base: cabecalho na linha 1, dados 2..5001, colunas A:U
    def _ler_eqa(self):
        ws = self.wb[ABA_EQA]
        it = ws.iter_rows(min_row=1, max_row=5001, max_col=21, values_only=True)
        cab = [_txt(c) for c in next(it)]
        self.eqa_cab = cab
        self.eqa = []
        for i, r in enumerate(it, start=2):
            if r is None or all(v is None for v in r):
                continue
            d = {cab[j]: r[j] for j in range(min(len(cab), len(r))) if cab[j]}
            d["_linha"] = i
            self.eqa.append(d)

    # Especificacoes: catalogo A4:AR43 e biblioteca anual (linhas 46..89, blocos de 9 linhas)
    def _ler_esp(self):
        ws = self.wb[ABA_ESP]
        linhas = list(ws.iter_rows(min_row=1, max_row=100, max_col=46, values_only=True))
        self.esp_ano_padrao = _num(linhas[1][18])  # S2
        self.catalogo = OrderedDict()
        for r in range(3, 43):  # linhas 4..43
            row = linhas[r]
            nome = _txt(row[0])
            if not nome or nome == "0":
                continue
            self.catalogo[nome] = {
                "linha": r + 1, "unidade": _txt(row[2]),
                "Q": _txt(row[16]), "S": _txt(row[18]).upper(), "AR": _txt(row[43]),
                "T_cache": row[19], "U_cache": row[20],
            }
        # biblioteca: procura "ANO" na coluna A a partir da linha 46
        self.biblio = {}
        r = 45
        while r < len(linhas):
            row = linhas[r]
            if _txt(row[0]).upper() == "ANO" and _num(row[1]) is not None:
                ano = int(_num(row[1]))
                nomes = linhas[r + 1]
                tab = {}
                for k in range(2, 9):
                    if r + k >= len(linhas):
                        break
                    rot = _txt(linhas[r + k][0]).upper()
                    tab[rot] = linhas[r + k]
                bloco = {}
                for c in range(1, len(nomes)):
                    an = _txt(nomes[c])
                    if not an or an == "0":
                        continue
                    def g(rot):
                        for kk, vv in tab.items():
                            if kk.startswith(rot):
                                return _num(vv[c]) if c < len(vv) else None
                        return None
                    bloco[an] = {"CLIA": g("CLIA ETP"), "CVI": g("VB-CVI"), "CVG": g("VB-CVG"),
                                 "FAB": g("FAB ETP"), "FABCV": g("FAB CVTP")}
                self.biblio[ano] = bloco
                r += 9
            else:
                r += 1

    def _ler_filtros(self):
        ws = self.wb[ABA_EST]
        L = {}
        for r, row in enumerate(ws.iter_rows(min_row=1, max_row=12, max_col=47, values_only=True), 1):
            L[r] = row
        self.f_visao = _txt(L[4][1]).upper()        # B4
        self.f_ano = L[4][3]                        # D4
        self.f_parte = L[4][5]                      # F4
        self.f_custIni = _dia(L[5][1])              # B5
        self.f_custFim = _dia(L[5][3])              # D5
        self.f_lote = _txt(L[4][7])                 # H4
        self.f_anoEP = L[4][13]                     # N4
        self.f_rodada = _txt(L[4][15])              # P4
        self.f_meses = _num(L[11][11])              # L11
        self.f_excl = []
        for r in range(8, 13):                      # V8:W12
            a, b = _dia(L[r][21]), _dia(L[r][22])
            if a and b and b >= a:
                self.f_excl.append((a, b))
        # cache para comparar janela
        self.cache_ini = _dia(L[5][7]); self.cache_fim = _dia(L[5][9])
        self.cache_sigIni = _dia(L[11][13]); self.cache_muFim = _dia(L[11][15])
        wp = self.wb[ABA_PAINEL]
        P = {r: row for r, row in enumerate(wp.iter_rows(min_row=1, max_row=8, max_col=20, values_only=True), 1)}
        self.f_equip = _txt(P[4][11]).upper()       # L4 (selEquipamento)
        self.painel_analito = _txt(P[3][2])         # C3
        self.painel_lote = _txt(P[3][7])            # H3
        self.painel_ini = _dia(P[3][6]); self.painel_fim = _dia(P[4][6])  # G3/G4

    def _ler_anos_lote(self):
        """Ano de referencia CLIA/VB de cada lote: bloco do lote na aba Analitos."""
        self.anos_lote = {}
        ws = self.wb[ABA_ANALITOS]
        lote = None
        for row in ws.iter_rows(min_row=1, max_col=12, values_only=True):
            a = _txt(row[0]).upper()
            if a == "LOTE":
                lote = _txt(row[1])
            elif a.startswith("CLIA - ANO DE REFER") and lote:
                self.anos_lote[lote] = (_num(row[3]), _num(row[9]))

    # ------------------------------------------------------------------ saida gravada pelo QC_INI
    def estatistica(self):
        ws = self.wb[ABA_EST]
        out = []
        for r, row in enumerate(ws.iter_rows(min_row=14, max_row=93, max_col=47, values_only=True), 14):
            out.append((r, row))
        return out

    def painel(self):
        wp = self.wb[ABA_PAINEL]
        return {r: row for r, row in enumerate(wp.iter_rows(min_row=1, max_row=70, max_col=20, values_only=True), 1)}


# --------------------------------------------------------------------------- calculo
class Calculadora:
    def __init__(self, d: Dados):
        self.d = d

    # ---- janelas
    def janela(self):
        d = self.d
        if d.f_visao == "PERSONALIZADO":
            return d.f_custIni, d.f_custFim
        ano = int(_num(d.f_ano) or 0)
        parte = int(_num(d.f_parte) or 1)
        if not ano:
            return None, None
        if d.f_visao == "ANUAL":
            return dt.date(ano, 1, 1), dt.date(ano, 12, 31)
        meses = {"MENSAL": 1, "TRIMESTRAL": 3, "SEMESTRAL": 6}.get(d.f_visao, 12)
        ini = dt.date(ano, (parte - 1) * meses + 1, 1)
        return ini, edate(ini, meses) - dt.timedelta(days=1)

    def mu_fim(self):
        m = None
        for r in self.d.cq:
            if _txt(r.get("PARTICIPA_ESTATISTICA")).upper() != "SIM":
                continue
            x = _dia(r.get("DATA"))
            if x and (m is None or x > m):
                m = x
        return m

    def janela_sigma(self, meses=None):
        fim = self.mu_fim()
        meses = int(meses if meses is not None else (self.d.f_meses or 6))
        if fim is None or meses < 1:
            return None, None
        return edate(fim, -meses) + dt.timedelta(days=1), fim

    # ---- CIQ
    def _elegivel(self, r, equip, ini, fim, excl):
        if _txt(r.get("PARTICIPA_ESTATISTICA")).upper() != "SIM":
            return None
        if equip and _txt(r.get("EQUIPAMENTO")).upper() != equip:
            return None
        dia = r.get("DATA_HORA")
        if not isinstance(dia, (dt.datetime, dt.date)):
            return None
        dia = _dia(dia)
        v = _num(r.get("RESULTADO"))
        if v is None:
            return None
        if ini and dia < ini:
            return None
        if fim and dia > fim:
            return None
        for a, b in excl:
            if a <= dia <= b:
                return None
        return dia, v

    def valores(self, analito, nivel, lote=None, ini=None, fim=None, equip=None, excl=None):
        """Lista de (ID, dia, valor, lote) que entram. lote=None -> todos os lotes."""
        equip = self.d.f_equip if equip is None else equip.upper()
        excl = self.d.f_excl if excl is None else excl
        an = analito.strip().upper()
        out = []
        for r in self.d.cq:
            if _txt(r.get("ANALITO")).upper() != an:
                continue
            if _txt(r.get("NIVEL")) != str(nivel) and _num(r.get("NIVEL")) != float(nivel):
                continue
            e = self._elegivel(r, equip, ini, fim, excl)
            if not e:
                continue
            lt = _txt(r.get("LOTE"))
            if lote is not None and lt != str(lote).strip():
                continue
            out.append((_txt(r.get("ID_REGISTRO")), e[0], e[1], lt))
        return out

    def ciq_stats(self, analito, nivel, lote=None, ini=None, fim=None, equip=None, excl=None):
        if ini is None and fim is None:
            ini, fim = self.janela()
        vals = [v for _, _, v, _ in self.valores(analito, nivel, lote, ini, fim, equip, excl)]
        n = len(vals)
        res = {"n": n, "media": None, "dp": None, "cv": None}
        if n >= 1:
            res["media"] = statistics.fmean(vals)
        if n >= 2:
            res["dp"] = statistics.stdev(vals)
            if res["media"] != 0:
                res["cv"] = res["dp"] / res["media"] * 100.0
        return res

    def cv_pooled(self, analito, nivel, ini=None, fim=None, equip=None, excl=None, nmin=20):
        if ini is None or fim is None:
            ini, fim = self.janela_sigma()
        por = defaultdict(list)
        dias = defaultdict(list)
        for _id, dia, v, lt in self.valores(analito, nivel, None, ini, fim, equip, excl):
            if not lt:
                continue
            por[lt].append(v)
            dias[lt].append(dia)
        gl = 0.0; sq = 0.0; lotes = []
        for lt, vs in por.items():
            n = len(vs)
            m = statistics.fmean(vs)
            info = {"lote": lt, "n": n, "media": m, "ini": min(dias[lt]).isoformat(), "fim": max(dias[lt]).isoformat()}
            if n >= nmin and m != 0:
                s = statistics.stdev(vs)
                cv = s / abs(m) * 100.0
                info.update(dp=s, cv=cv, entra=True)
                gl += n - 1
                sq += (n - 1) * cv * cv
            else:
                info.update(entra=False)
            lotes.append(info)
        cvp = math.sqrt(sq / gl) if gl > 0 else None
        return {"cvp": cvp, "gl": gl, "lotes": sorted(lotes, key=lambda x: x["ini"]),
                "ini": ini.isoformat() if ini else None, "fim": fim.isoformat() if fim else None}

    # ---- CEQ
    def _linhas_eqa(self, analito, provedor):
        an = analito.strip().upper()
        pv = (provedor or "").strip().upper()
        out = []
        for r in self.d.eqa:
            if _txt(r.get("Analito_Canonico")).upper() != an:
                continue
            if _txt(r.get("Uso_Analitico")).upper() == "NAO":
                continue
            if pv not in ("", "TODOS", "TODAS") and _txt(r.get("Provedor")).upper() != pv:
                continue
            ano = _num(r.get("Ano"))
            if ano is None:
                continue
            x = _num(r.get("Resultado")); h = _num(r.get("Valor_Alvo"))
            b = None
            if x is not None and h is not None and h != 0:
                b = (x - h) / h * 100.0
            out.append({"ano": int(ano), "rodada": _txt(r.get("Rodada")).upper(), "b": b, "x": x, "h": h,
                        "status": _txt(r.get("Status_Padronizado")).upper(), "chave": _txt(r.get("Chave")).upper(),
                        "linha": r["_linha"], "amostra": _txt(r.get("Amostra")),
                        "b_gravado": _num(r.get("Bias")), "babs_gravado": _num(r.get("Bias_Abs"))})
        return out

    def bias_ceq(self, analito, provedor=None, ano_ep=None, rodada=None):
        """Bias do ET/Sigma (BiasEQ "ABS" e "SIGNED"), recalculado de Resultado e Valor_Alvo."""
        d = self.d
        provedor = (d.catalogo.get(analito, {}).get("AR") or "CAP") if provedor is None else provedor
        ano_ep = d.f_anoEP if ano_ep is None else ano_ep
        rodada = (d.f_rodada if rodada is None else rodada).strip().upper()
        lin = [x for x in self._linhas_eqa(analito, provedor) if x["b"] is not None]
        ano_txt = _txt(ano_ep).upper()
        todos = ano_txt in ("TODOS", "TODAS", "TODOS OS ANOS")
        teto = 32767 if todos or _num(ano_ep) is None else int(_num(ano_ep))
        livre = rodada in ("", "TODAS", "TODOS")
        if todos and livre:
            sel = lin  # ADR-072 opcao 3
        elif livre or "|" not in rodada:
            cand = [x for x in lin if (livre or x["rodada"] == rodada) and x["ano"] <= teto]
            if not cand:
                return None
            vig = max(x["ano"] for x in cand)
            if provedor and _num(ano_ep) is not None and vig != teto:
                return None  # ADR-072: ano explicito sem dado = SEM EP
            sel = [x for x in cand if x["ano"] == vig]
        else:  # conjunto ano|rodada; ...
            itens = {t.strip() for t in rodada.split(";") if t.strip()}
            sel = [x for x in lin if f'{x["ano"]}|{x["rodada"]}' in itens]
        if not sel:
            return None
        por = OrderedDict()
        for x in sel:
            por.setdefault(f'{x["ano"]}|{x["rodada"]}', []).append(x)
        m_abs = [statistics.fmean(abs(x["b"]) for x in v) for v in por.values()]
        m_sig = [statistics.fmean(x["b"] for x in v) for v in por.values()]
        return {"abs": statistics.fmean(m_abs), "signed": statistics.fmean(m_sig), "n": len(sel),
                "nrod": len(por), "rodadas": {k: (len(v), statistics.fmean(abs(x["b"]) for x in v)) for k, v in por.items()},
                "amostras": sel}

    def vies_nordtest(self, analito, provedor=None, ano_ep=None):
        """ViesEQ MEDIA/RMS/N/NRODADAS (incerteza; ADR-064). Janela: ano vigente (+ anterior se preciso)."""
        d = self.d
        provedor = (d.catalogo.get(analito, {}).get("AR") or "CAP") if provedor is None else provedor
        ano_ep = d.f_anoEP if ano_ep is None else ano_ep
        teto = 32767 if _num(ano_ep) is None else int(_num(ano_ep))
        lin = self._linhas_eqa(analito, provedor)
        cand = [x for x in lin if x["b"] is not None and x["ano"] <= teto]
        if not cand:
            return None
        vig = max(x["ano"] for x in cand)
        if provedor and _num(ano_ep) is not None and vig != teto:
            return None
        for passo in (0, 1):
            vistos = set(); sel = []
            for x in lin:
                if x["ano"] > vig or x["ano"] < vig - passo:
                    continue
                if x["status"] == "NAO AVALIADO" or x["b"] is None:
                    continue
                if x["x"] in (None, 0) or x["h"] in (None, 0):
                    continue
                if x["chave"]:
                    if x["chave"] in vistos:
                        continue
                    vistos.add(x["chave"])
                sel.append(x)
            rods = {f'{x["ano"]}|{x["rodada"]}' for x in sel}
            if len(sel) >= 6 and len(rods) >= 2:
                break
        if not sel:
            return None
        bs = [x["b"] for x in sel]
        return {"media": statistics.fmean(bs), "rms": math.sqrt(statistics.fmean(b * b for b in bs)),
                "n": len(bs), "nrod": len(rods)}

    # ---- especificacao
    def etp(self, analito, lote=None):
        d = self.d
        c = d.catalogo.get(analito)
        if not c:
            return None, None, None
        lote = d.f_lote if lote is None else lote
        anos = d.anos_lote.get(str(lote), (d.esp_ano_padrao, d.esp_ano_padrao))
        aclia = int(anos[0] or d.esp_ano_padrao); avb = int(anos[1] or d.esp_ano_padrao)
        bc = d.biblio.get(aclia, {}).get(analito, {}); bv = d.biblio.get(avb, {}).get(analito, {})
        fonte = c["S"]
        etp = cvtp = None
        if fonte == "CLIA":
            v = bc.get("CLIA")
            if v and v > 0:
                etp, cvtp = v, v / 3.0
        elif fonte == "VB":
            cvi, cvg = bv.get("CVI"), bv.get("CVG")
            f = FATOR_VB.get(c["Q"].upper())
            if cvi and cvg and cvi > 0 and cvg > 0 and f:
                etp = 1.65 * f * cvi + f / 2.0 * math.sqrt(cvi * cvi + cvg * cvg)
                cvtp = f * cvi
        elif fonte == "FAB":
            v, w = bc.get("FAB"), bc.get("FABCV")
            if v and v > 0:
                etp, cvtp = v, w
        return (etp if etp else "-"), cvtp, fonte

    @staticmethod
    def classe(s):
        if s is None or isinstance(s, str):
            return s
        if s >= 6: return "Classe mundial"
        if s >= 5: return "Excelente"
        if s >= 4: return "Bom"
        if s >= 3: return "Marginal"
        return "Desempenho inadequado"

    # ---- linha completa da Estatistica
    def linha_estatistica(self, analito, nivel):
        ini, fim = self.janela()
        st = self.ciq_stats(analito, nivel, self.d.f_lote, ini, fim)
        b = self.bias_ceq(analito)
        etp, cvtp, fonte = self.etp(analito)
        si, sf = self.janela_sigma()
        cp = self.cv_pooled(analito, nivel, si, sf)
        r = OrderedDict(analito=analito, nivel=nivel, n=st["n"], media=st["media"], dp=st["dp"], cv=st["cv"],
                        bias_abs=b["abs"] if b else SEM_EP, bias_sig=b["signed"] if b else SEM_EP,
                        ceq_n=b["n"] if b else SEM_EP, ceq_nrod=b["nrod"] if b else SEM_EP,
                        fonte=fonte, cvtp=cvtp, etp=etp, cvp=cp["cvp"], cvp_gl=cp["gl"])
        r["et"] = (1.65 * st["cv"] + abs(b["abs"])) if (st["cv"] is not None and b) else None
        if etp == "-":
            r["sigma"] = "-"
        elif b and cp["cvp"]:
            r["sigma"] = (etp - abs(b["abs"])) / cp["cvp"]
        else:
            r["sigma"] = None
        r["classe"] = self.classe(r["sigma"])
        if etp != "-" and r["et"] is not None:
            r["margem_pp"] = etp - r["et"]; r["margem_pct"] = r["margem_pp"] / etp * 100.0
        else:
            r["margem_pp"] = r["margem_pct"] = ("-" if etp == "-" else None)
        if isinstance(r["sigma"], float):
            r["dpm"] = (1.0 - phi(r["sigma"] - 1.5)) * 1e6
            r["rend"] = 100.0 - r["dpm"] / 1e4
        else:
            r["dpm"] = r["rend"] = r["sigma"]
        ui, uf = self.janela_sigma(12)
        r["urw12"] = self.cv_pooled(analito, nivel, ui, uf)["cvp"]
        v = self.vies_nordtest(analito)
        r["vies_media"] = v["media"] if v else SEM_EP
        r["vies_rms"] = v["rms"] if v else None
        return r


# --------------------------------------------------------------------------- comparacao
# Estatistica: coluna (indice 0-based) -> chave do resultado independente
MAPA = [("C", 2, "n"), ("D", 3, "media"), ("E", 4, "dp"), ("F", 5, "cv"), ("G", 6, "bias_abs"),
        ("H", 7, "et"), ("J", 9, "cvtp"), ("K", 10, "etp"), ("L", 11, "sigma"), ("M", 12, "classe"),
        ("N", 13, "margem_pp"), ("O", 14, "margem_pct"), ("T", 19, "bias_sig"), ("V", 21, "dpm"),
        ("W", 22, "rend"), ("AC", 28, "ceq_n"), ("AD", 29, "ceq_nrod"), ("AF", 31, "urw12"), ("AR", 43, "vies_media")]


def _vazio(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def comparar(qc, ind, tol):
    if _vazio(qc) and _vazio(ind):
        return "OK", 0.0
    if isinstance(qc, str) or isinstance(ind, str):
        a = _txt(qc); b = _txt(ind)
        return ("OK" if a.upper() == b.upper() else "DIF"), None
    if qc is None or ind is None:
        return "DIF", None
    dr = dif_rel(float(qc), float(ind))
    return ("OK" if dr <= tol else "DIF"), dr


def auditar(caminho, tol=1e-9):
    d = Dados(caminho)
    c = Calculadora(d)
    ini, fim = c.janela()
    si, sf = c.janela_sigma()
    meta = OrderedDict(arquivo=caminho, lote=d.f_lote, equipamento=d.f_equip, visao=d.f_visao,
                       janela=(ini.isoformat() if ini else None, fim.isoformat() if fim else None),
                       janela_cache=(d.cache_ini.isoformat() if d.cache_ini else None,
                                     d.cache_fim.isoformat() if d.cache_fim else None),
                       ano_ep=d.f_anoEP, rodada=d.f_rodada, meses_sigma=d.f_meses,
                       janela_sigma=(si.isoformat() if si else None, sf.isoformat() if sf else None),
                       janela_sigma_cache=(d.cache_sigIni.isoformat() if d.cache_sigIni else None,
                                           d.cache_muFim.isoformat() if d.cache_muFim else None),
                       exclusoes=[(a.isoformat(), b.isoformat()) for a, b in d.f_excl],
                       anos_ref_lote=d.anos_lote.get(d.f_lote), n_cq=len(d.cq), n_eqa=len(d.eqa))
    linhas = []
    for r, row in d.estatistica():
        an = _txt(row[0])
        if not an or an == "0":
            continue
        nivel = int(_num(row[1]) or 0)
        ind = c.linha_estatistica(an, nivel)
        comp = OrderedDict()
        for col, i, k in MAPA:
            st, dr = comparar(row[i], ind[k], tol)
            comp[col] = {"qc": row[i], "ind": ind[k], "status": st, "dif_rel": dr}
        linhas.append({"linha": r, "analito": an, "nivel": nivel, "ind": ind, "comp": comp})
    return d, c, meta, linhas


def _fmt(v):
    if isinstance(v, float):
        return f"{v:.10g}"
    return str(v)


# =========================================================================== HEMATOLOGIA
# Acrescentado no ciclo 2026-10-31_ciclo01 (frente F3, Hematologia). O layout da Hematologia difere do da
# Bioquimica em quatro pontos, todos conferidos no codigo em 31/10/2026:
#   1. Estatistica C:F (n, media, DP, CV) = mEstatistica.EstatBasica (src/hema) publicada em Eng_Saida 190..309:
#      PARTICIPA_ESTATISTICA = SIM; EQUIPAMENTO = EquipFiltro() (Cfg_Integracao!B11 = selEquipamento, senao
#      EQUIPAMENTO_PADRAO); LOTE = Estatistica!B4 (= loteAnalise = Painel!H3, ou Configuracao!C20 com H3 vazio);
#      filtro por ANO de DATA_HORA entre B3 (Ano De) e D3 (Ano Ate) -- Ano Ate < Ano De vira Ano De; Ano De = 0
#      = sem filtro. DP amostral (n-1); n < 2 -> DP/CV vazios; media 0 -> CV vazio.
#   2. Provedor do CEQ UNICO, em Estatistica!R4 (eqProvedor); a Bioquimica usa Especificacoes!AR por analito.
#   3. Especificacoes: K ETp CLIA, M ETp FAB, N CVTp FAB, O CVi, P CVg, Q desempenho, R ETp VB, S fonte,
#      T ETp em uso, U CVTp (cvtpOficial), V = K/3, W = fi(Q)*CVi com fi da tabela X4:Y6 (padrao 0,75).
#      K, M, N, O, P vem da biblioteca anual (linhas 46..89) do ano CLIA/VB em uso do lote (celula vazia = 0).
#   4. Estatistica tem 120 linhas (14..133, 40 analitos x 3 niveis) e as colunas AF..AU de incerteza.
#   Bias (G/T/AC/AD), ET (H), Sigma (L), margem (N/O/P), DPM (V), rendimento (W), u(Rw) (AF) e vies (AR) seguem as
#   mesmas definicoes da Bioquimica (mCEQ.BiasEQ, mIncerteza.IncertezaCIQ, mPlanoQC). O Sigma da Hematologia NAO
#   passa exclusoes de periodo para o CV pooled (a formula L chama IncertezaCIQ sem o 6o argumento).
MARGEM_CRITICA_PCT = 10.0
ABA_CFG_INT = "Cfg_Integracao"
ABA_CONFIG = "Configuração"


def _biblioteca(linhas, r0=45):
    """Biblioteca anual (blocos de 9 linhas a partir de 'ANO' na coluna A). Celula vazia -> None."""
    biblio = {}
    r = r0
    while r < len(linhas):
        row = linhas[r]
        if _txt(row[0]).upper() == "ANO" and _num(row[1]) is not None:
            ano = int(_num(row[1]))
            nomes = linhas[r + 1]
            tab = {}
            for k in range(2, 9):
                if r + k < len(linhas):
                    tab[_txt(linhas[r + k][0]).upper()] = linhas[r + k]
            bloco = {}
            for c in range(1, len(nomes)):
                an = _txt(nomes[c])
                if not an or an == "0":
                    continue

                def g(rot, c=c):
                    for kk, vv in tab.items():
                        if kk.startswith(rot):
                            return _num(vv[c]) if c < len(vv) else None
                    return None
                bloco[an] = {"CLIA": g("CLIA ETP"), "CVI": g("VB-CVI"), "CVG": g("VB-CVG"),
                             "FAB": g("FAB ETP"), "FABCV": g("FAB CVTP")}
            biblio[ano] = bloco
            r += 9
        else:
            r += 1
    return biblio


class DadosHema(Dados):
    """Leitura de uma COPIA de QC_Hematologia.xlsm (openpyxl, read_only, data_only)."""
    produto = "hema"

    def __init__(self, caminho: str):
        self.caminho = caminho
        self.wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True, keep_vba=False)
        self._ler_cq()
        self._ler_eqa()
        self._ler_esp_hema()
        self._ler_filtros_hema()
        self._ler_anos_lote()

    def _ler_esp_hema(self):
        ws = self.wb[ABA_ESP]
        linhas = list(ws.iter_rows(min_row=1, max_row=100, max_col=46, values_only=True))
        self.esp_ano_padrao = _num(linhas[1][18])  # S2
        self.fi = {}
        for r in range(3, 6):  # X4:Y6
            self.fi[_txt(linhas[r][23]).upper()] = _num(linhas[r][24])
        self.catalogo = OrderedDict()
        for r in range(3, 43):  # linhas 4..43
            row = linhas[r]
            nome = _txt(row[0])
            if not nome or nome == "0":
                continue
            self.catalogo[nome] = {"linha": r + 1, "unidade": _txt(row[2]), "Q": _txt(row[16]).upper(),
                                   "S": _txt(row[18]).upper(), "T_cache": row[19], "U_cache": row[20],
                                   "K_cache": row[10], "R_cache": row[17]}
        self.biblio = _biblioteca(linhas)

    def _ler_filtros_hema(self):
        ws = self.wb[ABA_EST]
        L = {r: row for r, row in enumerate(ws.iter_rows(min_row=1, max_row=12, max_col=47, values_only=True), 1)}
        self.f_anoDe = int(_num(L[3][1]) or 0)       # B3
        self.f_anoAte = int(_num(L[3][3]) or 0)      # D3
        self.cache_lote = _txt(L[4][1])              # B4 (=loteAnalise, formula)
        self.f_anoEP = L[4][13]                      # N4
        self.f_rodada = _txt(L[4][15])               # P4
        self.f_provedor = _txt(L[4][17])             # R4
        self.f_meses = _num(L[11][11])               # L11
        self.f_excl = []                             # a Hematologia nao tem exclusoes no Sigma
        self.cache_sigIni = _dia(L[11][13]); self.cache_muFim = _dia(L[11][15])
        self.cache_muIni = _dia(L[11][32]); self.cache_muFim12 = _dia(L[11][34])
        self.cache_k5 = _txt(L[5][10])
        wp = self.wb[ABA_PAINEL]
        P = {r: row for r, row in enumerate(wp.iter_rows(min_row=1, max_row=8, max_col=20, values_only=True), 1)}
        self.painel_analito = _txt(P[3][2])          # C3
        self.painel_lote = _txt(P[3][7])             # H3
        self.painel_de = _dia(P[3][6]); self.painel_ate = _dia(P[4][6])   # G3/G4 (filtroDe/filtroAte)
        self.painel_ano = P[3][10]; self.painel_periodo = _txt(P[4][10])  # K3/K4
        wc = self.wb[ABA_CONFIG]
        self.config_lote = _txt(list(wc.iter_rows(min_row=20, max_row=20, min_col=3, max_col=3, values_only=True))[0][0])
        # loteAnalise = IF(Painel!H3="", Configuracao!C20, Painel!H3) -- recomposto das ENTRADAS
        self.f_lote = self.painel_lote if self.painel_lote else self.config_lote
        wi = self.wb[ABA_CFG_INT]
        eq = ""; padrao = ""
        for r, row in enumerate(wi.iter_rows(min_row=1, max_row=40, max_col=2, values_only=True), 1):
            if r == 11:
                eq = _txt(row[1]).upper()            # selEquipamento = Cfg_Integracao!B11
            if _txt(row[0]).upper() == "EQUIPAMENTO_PADRAO":
                padrao = _txt(row[1]).upper()
        self.f_equip = eq or padrao

    def estatistica(self):
        ws = self.wb[ABA_EST]
        return [(r, row) for r, row in enumerate(ws.iter_rows(min_row=14, max_row=133, max_col=47, values_only=True), 14)]

    def painel(self):
        wp = self.wb[ABA_PAINEL]
        return {r: row for r, row in enumerate(wp.iter_rows(min_row=1, max_row=80, max_col=23, values_only=True), 1)}


class CalculadoraHema(Calculadora):
    # ---- CIQ por ano (EstatBasica)
    def valores_ano(self, analito, nivel, lote, anoDe, anoAte, equip=None):
        equip = self.d.f_equip if equip is None else equip.upper()
        if anoAte < anoDe:
            anoAte = anoDe
        an = analito.strip().upper()
        out = []
        for r in self.d.cq:
            if _txt(r.get("PARTICIPA_ESTATISTICA")).upper() != "SIM":
                continue
            if _txt(r.get("ANALITO")).upper() != an:
                continue
            if equip and _txt(r.get("EQUIPAMENTO")).upper() != equip:
                continue
            if _num(r.get("NIVEL")) != float(nivel):
                continue
            if lote and _txt(r.get("LOTE")) != str(lote).strip():
                continue
            v = _num(r.get("RESULTADO"))
            if v is None:
                continue
            if anoDe:
                dh = r.get("DATA_HORA")
                if not isinstance(dh, (dt.datetime, dt.date)) or not (anoDe <= dh.year <= anoAte):
                    continue
            out.append((_txt(r.get("ID_REGISTRO")), _dia(r.get("DATA_HORA")), v, _txt(r.get("LOTE"))))
        return out

    def valores_periodo(self, analito, nivel, lote, ini=None, fim=None, equip=None):
        """Para o Painel (motor por corrida, filtroDe/filtroAte por dia)."""
        return [x for x in self.valores_ano(analito, nivel, lote, 0, 0, equip)
                if (ini is None or (x[1] and x[1] >= ini)) and (fim is None or (x[1] and x[1] <= fim))]

    @staticmethod
    def _stats(vals):
        n = len(vals)
        res = {"n": n, "media": None, "dp": None, "cv": None}
        if n >= 1:
            res["media"] = statistics.fmean(vals)
        if n >= 2:
            res["dp"] = statistics.stdev(vals)
            if res["media"] != 0:
                res["cv"] = res["dp"] / res["media"] * 100.0
        return res

    def ciq_ano(self, analito, nivel, lote=None, anoDe=None, anoAte=None):
        lote = self.d.f_lote if lote is None else lote
        anoDe = self.d.f_anoDe if anoDe is None else anoDe
        anoAte = self.d.f_anoAte if anoAte is None else anoAte
        return self._stats([v for _, _, v, _ in self.valores_ano(analito, nivel, lote, anoDe, anoAte)])

    # ---- CV pooled com as metricas auxiliares (IncertezaCIQ CV/GL/LOTES/DIAS/MEDIA/VALIDADE)
    def cv_pooled_ext(self, analito, nivel, ini, fim, nmin=20):
        por = defaultdict(list); dias = defaultdict(list)
        for _id, dia, v, lt in self.valores(analito, nivel, None, ini, fim, None, []):
            if lt:
                por[lt].append(v); dias[lt].append(dia)
        gl = sq = 0.0; ntot = 0; somax = 0.0; nl = 0; dmin = dmax = None; lotes = []
        for lt, vs in por.items():
            n = len(vs); m = statistics.fmean(vs)
            info = {"lote": lt, "n": n, "media": m, "ini": min(dias[lt]).isoformat(), "fim": max(dias[lt]).isoformat()}
            if n >= nmin and m != 0:
                cv = statistics.stdev(vs) / abs(m) * 100.0
                info.update(cv=cv, entra=True)
                gl += n - 1; sq += (n - 1) * cv * cv; ntot += n; somax += sum(vs); nl += 1
                a, b = min(dias[lt]), max(dias[lt])
                dmin = a if dmin is None or a < dmin else dmin
                dmax = b if dmax is None or b > dmax else dmax
            else:
                info.update(entra=False)
            lotes.append(info)
        dd = (dmax - dmin).days if nl else None
        if nl == 0 or dd < 90 or gl < 30:
            val = "INSUFICIENTE"
        elif dd < 180 or gl < 100:
            val = "PROVISORIA"
        else:
            val = "VALIDA"
        return {"cvp": math.sqrt(sq / gl) if gl > 0 else None, "gl": gl, "lotes_n": nl, "dias": dd,
                "media": somax / ntot if ntot else None, "validade": val,
                "lotes": sorted(lotes, key=lambda x: x["ini"]),
                "ini": ini.isoformat() if ini else None, "fim": fim.isoformat() if fim else None}

    # ---- especificacao (Especificacoes K/M/N/O/P/R/S/T/U/V/W)
    def etp(self, analito, lote=None):
        d = self.d
        c = d.catalogo.get(analito)
        if not c:
            return "-", "-", None
        lote = d.f_lote if lote is None else lote
        anos = d.anos_lote.get(str(lote), (d.esp_ano_padrao, d.esp_ano_padrao))
        aclia = int(anos[0] or d.esp_ano_padrao); avb = int(anos[1] or d.esp_ano_padrao)
        bc = d.biblio.get(aclia, {}).get(analito, {}); bv = d.biblio.get(avb, {}).get(analito, {})
        K = bc.get("CLIA") or 0.0; M = bc.get("FAB") or 0.0; N = bc.get("FABCV") or 0.0
        O = bv.get("CVI") or 0.0; P = bv.get("CVG") or 0.0
        q = c["Q"]
        R = None
        if O > 0 and P > 0 and q in ("ÓTI", "DES", "MIN"):
            f = {"ÓTI": 0.25, "DES": 0.50, "MIN": 0.75}[q]
            R = 1.65 * (f * O) + (f / 2.0) * math.sqrt(O * O + P * P)
        V = K / 3.0
        W = (d.fi.get(q) if d.fi.get(q) is not None else 0.75) * O
        fonte = c["S"]
        if fonte == "CLIA":
            etp, cvtp = (K if K > 0 else "-"), (V if V > 0 else "-")
        elif fonte == "VB":
            etp, cvtp = (R if R and R > 0 else "-"), (W if W > 0 else "-")
        elif fonte == "FAB":
            etp, cvtp = (M if M > 0 else "-"), (N if N > 0 else "-")
        else:
            etp, cvtp = "-", "-"
        return etp, cvtp, fonte

    @staticmethod
    def classe_margem(o):
        if not isinstance(o, float):
            return o
        if o < 0:
            return "ETp excedido"
        if o <= MARGEM_CRITICA_PCT:
            return "Margem critica"
        return "Dentro do orcamento"

    def status_sdi(self, analito, provedor, ano_ep):
        """StatusSDIeq (TODAS): pior |SDI| (coluna SDI do provedor) das amostras do ano vigente."""
        teto = 32767 if _num(ano_ep) is None else int(_num(ano_ep))
        an = analito.upper(); pv = provedor.upper()
        lin = [r for r in self.d.eqa if _txt(r.get("Analito_Canonico")).upper() == an
               and _txt(r.get("Uso_Analitico")).upper() != "NAO" and _txt(r.get("Provedor")).upper() == pv
               and _num(r.get("Ano")) is not None and _num(r.get("SDI")) is not None and int(_num(r.get("Ano"))) <= teto]
        if not lin:
            return SEM_EP
        vig = max(int(_num(r.get("Ano"))) for r in lin)
        if _num(ano_ep) is not None and vig != teto:
            return SEM_EP
        s = [_num(r.get("SDI")) for r in lin if int(_num(r.get("Ano"))) == vig]
        m = max(abs(x) for x in s)
        if m <= 2.0:
            return f"OK (|SDI| max {m:.2f} em {len(s)} amostra(s))".replace(".", ",")
        return f"FORA (|SDI| max {m:.2f} > 2)".replace(".", ",")

    def linha_estatistica(self, analito, nivel):
        d = self.d
        st = self.ciq_ano(analito, nivel)
        b = self.bias_ceq(analito, d.f_provedor, d.f_anoEP, d.f_rodada)
        etp, cvtp, fonte = self.etp(analito)
        si, sf = self.janela_sigma()
        cp = self.cv_pooled_ext(analito, nivel, si, sf)
        r = OrderedDict(analito=analito, nivel=nivel, n=st["n"], media=st["media"], dp=st["dp"], cv=st["cv"],
                        bias_abs=b["abs"] if b else SEM_EP, bias_sig=b["signed"] if b else SEM_EP,
                        ceq_n=b["n"] if b else SEM_EP, ceq_nrod=b["nrod"] if b else SEM_EP,
                        fonte=fonte, cvtp=cvtp, etp=etp, cvp=cp["cvp"], cvp_gl=cp["gl"])
        r["et"] = (1.65 * st["cv"] + abs(b["abs"])) if (st["cv"] is not None and b) else None
        if etp == "-":
            r["sigma"] = "-"
        elif b and cp["cvp"]:
            r["sigma"] = (etp - abs(b["abs"])) / cp["cvp"]
        else:
            r["sigma"] = None
        r["classe"] = "-" if etp == "-" else self.classe(r["sigma"])
        if etp != "-" and r["et"] is not None:
            r["margem_pp"] = etp - r["et"]; r["margem_pct"] = r["margem_pp"] / etp * 100.0
        else:
            r["margem_pp"] = r["margem_pct"] = ("-" if etp == "-" else None)
        r["margem_cls"] = "-" if etp == "-" else self.classe_margem(r["margem_pct"])
        if st["cv"] is not None and isinstance(cvtp, float):
            r["status_cv"] = "OK" if st["cv"] <= cvtp else "CV acima da meta"
        else:
            r["status_cv"] = None
        r["status_sdi"] = self.status_sdi(analito, d.f_provedor, d.f_anoEP)
        if isinstance(r["sigma"], float):
            r["dpm"] = (1.0 - phi(r["sigma"] - 1.5)) * 1e6
            r["rend"] = 100.0 - r["dpm"] / 1e4
        else:
            r["dpm"] = r["rend"] = r["sigma"]
        ui, uf = self.janela_sigma(12)
        u = self.cv_pooled_ext(analito, nivel, ui, uf)
        r["urw12"] = u["cvp"] if u["validade"] != "INSUFICIENTE" else None
        r["urw_gl"] = u["gl"]; r["urw_lotes"] = u["lotes_n"]; r["urw_dias"] = u["dias"] if u["dias"] is not None else ""
        v = self.vies_nordtest(analito, d.f_provedor, d.f_anoEP)
        r["vies_media"] = v["media"] if v else SEM_EP
        r["vies_rms"] = v["rms"] if v else None
        r["vies_n_rod"] = f'{v["n"]} / {v["nrod"]}' if v else f"{SEM_EP} / {SEM_EP}"
        r["_cvp_lotes"] = cp["lotes"]
        return r


# DPM = (1 - Phi(Sigma - 1,5)) * 1e6. Com Sigma > ~6,5, Phi fica a menos de 1e-11 de 1 e o "1 - Phi" perde os digitos
# relativos (cancelamento): o QC_INI (Phi do Excel) e esta calculadora (math.erf) erram ~1e-16 em Phi, ou ~1e-10 DPM.
# Para V vale entao uma tolerancia ABSOLUTA de 1e-9 DPM (1e-15 em probabilidade); o metodo e o mesmo dos dois lados.
TOL_ABS_HEMA = {"V": 1e-9}

MAPA_HEMA = [("C", 2, "n"), ("D", 3, "media"), ("E", 4, "dp"), ("F", 5, "cv"), ("G", 6, "bias_abs"),
             ("H", 7, "et"), ("I", 8, "fonte"), ("J", 9, "cvtp"), ("K", 10, "etp"), ("L", 11, "sigma"),
             ("M", 12, "classe"), ("N", 13, "margem_pp"), ("O", 14, "margem_pct"), ("P", 15, "margem_cls"),
             ("Q", 16, "status_cv"), ("R", 17, "status_sdi"), ("T", 19, "bias_sig"), ("V", 21, "dpm"),
             ("W", 22, "rend"), ("AC", 28, "ceq_n"), ("AD", 29, "ceq_nrod"), ("AF", 31, "urw12"),
             ("AG", 32, "urw_gl"), ("AH", 33, "urw_lotes"), ("AI", 34, "urw_dias"), ("AR", 43, "vies_media"),
             ("AU", 46, "vies_n_rod")]


def detectar_produto(caminho):
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True, keep_vba=False)
    try:
        a1 = _txt(next(wb[ABA_EST].iter_rows(min_row=1, max_row=1, max_col=1, values_only=True))[0]).upper()
    finally:
        wb.close()
    return "hema" if "HEMATOLOGIA" in a1 else "bio"


def auditar_hema(caminho, tol=1e-9):
    d = DadosHema(caminho)
    c = CalculadoraHema(d)
    si, sf = c.janela_sigma()
    ui, uf = c.janela_sigma(12)
    meta = OrderedDict(arquivo=caminho, produto="hema", lote=d.f_lote, lote_cache_B4=d.cache_lote,
                       equipamento=d.f_equip, ano_de=d.f_anoDe, ano_ate=d.f_anoAte, provedor=d.f_provedor,
                       ano_ep=d.f_anoEP, rodada=d.f_rodada, meses_sigma=d.f_meses,
                       janela_sigma=(si.isoformat() if si else None, sf.isoformat() if sf else None),
                       janela_sigma_cache=(d.cache_sigIni.isoformat() if d.cache_sigIni else None,
                                           d.cache_muFim.isoformat() if d.cache_muFim else None),
                       janela_12m=(ui.isoformat() if ui else None, uf.isoformat() if uf else None),
                       janela_12m_cache=(d.cache_muIni.isoformat() if d.cache_muIni else None,
                                         d.cache_muFim12.isoformat() if d.cache_muFim12 else None),
                       k5=d.cache_k5, anos_ref_lote=d.anos_lote.get(d.f_lote), n_cq=len(d.cq), n_eqa=len(d.eqa),
                       painel=dict(analito=d.painel_analito, lote=d.painel_lote, de=str(d.painel_de),
                                   ate=str(d.painel_ate), ano=d.painel_ano, periodo=d.painel_periodo))
    linhas = []
    for r, row in d.estatistica():
        an = _txt(row[0])
        if not an or an == "0":
            continue
        nivel = int(_num(row[1]) or 0)
        ind = c.linha_estatistica(an, nivel)
        comp = OrderedDict()
        for col, i, k in MAPA_HEMA:
            st, dr = comparar(row[i], ind[k], tol)
            if st != "OK" and col in TOL_ABS_HEMA and isinstance(row[i], float) and isinstance(ind[k], float)                     and abs(row[i] - ind[k]) <= TOL_ABS_HEMA[col]:
                st = "OK"  # tolerancia ABSOLUTA justificada (ver TOL_ABS_HEMA)
            comp[col] = {"qc": row[i], "ind": ind[k], "status": st, "dif_rel": dr}
        linhas.append({"linha": r, "analito": an, "nivel": nivel, "ind": ind, "comp": comp})
    # Painel (analito selecionado): motor por corrida no filtro De/Ate do Painel + leituras da Estatistica
    P = d.painel()
    an = d.painel_analito
    pain = []
    for t in (1, 2, 3):
        st = c._stats([v for _, _, v, _ in c.valores_periodo(an, t, d.f_lote, d.painel_de, d.painel_ate)])
        ind = next((L["ind"] for L in linhas if L["analito"] == an and L["nivel"] == t), None)
        if ind is None:
            continue
        cp = c.cv_pooled_ext(an, t, si, sf)
        et_painel = (1.65 * st["cv"] + (abs(ind["bias_abs"]) if isinstance(ind["bias_abs"], float) else 0.0)) \
            if st["cv"] is not None else None
        row7 = P[6 + t]; row45 = P[44 + t]; row61 = P[60 + t]
        itens = [("B%d" % (6 + t), row7[1], st["n"]), ("C%d" % (6 + t), row7[2], st["media"]),
                 ("D%d" % (6 + t), row7[3], st["dp"]), ("E%d" % (6 + t), row7[4], st["cv"]),
                 ("F%d" % (6 + t), row7[5], ind["etp"]), ("G%d" % (6 + t), row7[6], ind["bias_abs"]),
                 ("H%d" % (6 + t), row7[7], et_painel), ("I%d" % (6 + t), row7[8], ind["sigma"]),
                 ("C%d" % (44 + t), row45[2], cp["cvp"]), ("J%d" % (44 + t), row45[9], ind["sigma"]),
                 ("M%d" % (44 + t), row45[12], ind["dpm"]), ("O%d" % (44 + t), row45[14], ind["rend"]),
                 ("C%d" % (60 + t), row61[2], ind["et"]), ("H%d" % (60 + t), row61[7], ind["margem_pp"])]
        for cel, qc, iv in itens:
            s, dr = comparar(qc, iv, tol)
            pain.append({"celula": cel, "nivel": t, "qc": qc, "ind": iv, "status": s, "dif_rel": dr})
    cont = defaultdict(int)
    for L in linhas:
        cont[L["ind"]["margem_cls"]] += 1
    for cel, rot in (("E74", "ETp excedido"), ("E75", "Margem critica"), ("E76", "Dentro do orcamento")):
        qc = P[int(cel[1:])][4]
        s, dr = comparar(qc, float(cont[rot]), tol)
        pain.append({"celula": cel, "nivel": None, "qc": qc, "ind": cont[rot], "status": s, "dif_rel": dr})
    meta["painel_comparacoes"] = pain
    return d, c, meta, linhas



def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("xlsm")
    ap.add_argument("--json")
    ap.add_argument("--tol", type=float, default=1e-9)
    ap.add_argument("--produto", choices=("auto", "bio", "hema"), default="auto",
                    help="layout da pasta (auto = pelo titulo da aba Estatistica)")
    a = ap.parse_args(argv)
    prod = detectar_produto(a.xlsm) if a.produto == "auto" else a.produto
    if prod == "hema":
        d, c, meta, linhas = auditar_hema(a.xlsm, a.tol)
    else:
        d, c, meta, linhas = auditar(a.xlsm, a.tol)
    pain = meta.pop("painel_comparacoes", None)
    print("META", json.dumps(meta, ensure_ascii=False, default=str))
    if pain is not None:
        meta["painel_comparacoes"] = pain
        for x in pain:
            if x["status"] != "OK":
                print(f'DIF Painel {x["celula"]}: QC={_fmt(x["qc"])} IND={_fmt(x["ind"])} dr={x["dif_rel"]}')
        print("PAINEL (OK, DIF):", (sum(x["status"] == "OK" for x in pain), sum(x["status"] != "OK" for x in pain)))
    tot = defaultdict(lambda: [0, 0])
    for L in linhas:
        for col, x in L["comp"].items():
            tot[col][0 if x["status"] == "OK" else 1] += 1
            if x["status"] != "OK":
                print(f'DIF L{L["linha"]} {L["analito"]} N{L["nivel"]} {col}: QC={_fmt(x["qc"])} IND={_fmt(x["ind"])} dr={x["dif_rel"]}')
    print("RESUMO por coluna (OK, DIF):", {k: tuple(v) for k, v in tot.items()})
    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            json.dump({"meta": meta, "linhas": linhas}, f, ensure_ascii=False, indent=1, default=str)
    return 0


if __name__ == "__main__":
    sys.exit(main())
