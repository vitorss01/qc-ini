"""
QC-INI  —  Controle de Qualidade Interno (protótipo SQL + Python)
Aplicação web responsiva: login, painel com gráfico de Levey-Jennings,
regras de Westgard, métricas Sigma/Erro Total, lançamento de resultados,
gestão de não conformidades com assinatura eletrônica.

Executar:  streamlit run app.py
Login de teste:  INILAB / TESTE05
"""

import io
from datetime import datetime, date

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import database as db
import qc_engine as qc
import seed_data as seed

st.set_page_config(page_title="QC-INI · Controle de Qualidade",
                   page_icon="🧪", layout="wide", initial_sidebar_state="expanded")

# garante schema + dados de demonstração (só popula se vazio)
db.init_db(force=False, demo=True)

# --------------------------------------------------------------------------- #
# Estilo
# --------------------------------------------------------------------------- #
CSS = """
<style>
:root{ --mint:#5fe3c2; --mint2:#7af0d4; --bg:#0f2a2e; --card:#1b3f45; --card2:#16383d;
       --ok:#46d39a; --warn:#f4c430; --bad:#ff5d6c; --txt:#e7f3f1; --muted:#8fb3ad; }
.block-container{ padding-top:3rem !important; max-width:1400px; }
/* Barra superior do Streamlit com fundo sólido para nunca cobrir o título */
[data-testid="stHeader"]{ background:var(--bg) !important; }
/* Cabeçalhos: altura de linha e espaço suficientes para não cortar o texto */
h1,h2,h3,h4{ letter-spacing:.3px; line-height:1.6 !important; overflow:visible !important;
             white-space:normal !important; padding-top:.12em; padding-bottom:.18em;
             margin-top:.2rem; margin-bottom:.35rem; }
[data-testid="stMarkdownContainer"] h1,
[data-testid="stMarkdownContainer"] h2,
[data-testid="stMarkdownContainer"] h3,
[data-testid="stMarkdownContainer"] h4{ overflow:visible !important; white-space:normal !important;
             line-height:1.5 !important; }
/* Barra lateral mais larga (SÓ em tablet/desktop; no celular o Streamlit a
   transforma em overlay e forçar largura quebra o layout) */
@media (min-width: 769px){
  section[data-testid="stSidebar"]{ min-width:300px !important; width:300px !important; }
}
section[data-testid="stSidebar"] .stRadio label{ white-space:normal !important; }
section[data-testid="stSidebar"] .stRadio label p{ font-size:15px; line-height:1.5;
             white-space:normal !important; }
.kpi-grid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:12px; margin:6px 0 4px; }
.kpi{ background:linear-gradient(160deg,var(--card),var(--card2)); border:1px solid #265055;
      border-radius:16px; padding:14px 16px; box-shadow:0 6px 18px rgba(0,0,0,.25); }
.kpi .v{ font-size:26px; font-weight:800; line-height:1.1; color:var(--txt); }
.kpi .l{ font-size:11px; letter-spacing:1.2px; color:var(--muted); text-transform:uppercase; margin-top:4px; }
.kpi.ok .v{ color:var(--ok);} .kpi.warn .v{ color:var(--warn);} .kpi.bad .v{ color:var(--bad);}
.badge{ display:inline-block; padding:2px 10px; border-radius:999px; font-size:12px; font-weight:700; }
.badge.ok{ background:rgba(70,211,154,.15); color:var(--ok); border:1px solid var(--ok);}
.badge.warn{ background:rgba(244,196,48,.15); color:var(--warn); border:1px solid var(--warn);}
.badge.bad{ background:rgba(255,93,108,.15); color:var(--bad); border:1px solid var(--bad);}
.wrlist{ background:var(--card2); border:1px solid #265055; border-radius:14px; padding:8px 12px;
         max-height:430px; overflow:auto; font-size:13px; }
.wrlist .row{ padding:4px 2px; border-bottom:1px dashed #2c5258; }
.logo{ font-size:42px; font-weight:900; letter-spacing:-2px; color:var(--mint); }
.small{ color:var(--muted); font-size:12px; }
.stButton>button{ border-radius:10px; border:1px solid var(--mint); color:var(--bg);
                  background:var(--mint); font-weight:700; }
.stButton>button:hover{ background:var(--mint2); border-color:var(--mint2); }

/* ===================== RESPONSIVIDADE ===================== */
/* Tablet (769–1024px): grade de KPIs mais compacta */
@media (min-width: 769px) and (max-width: 1024px){
  .kpi-grid{ grid-template-columns:repeat(auto-fit,minmax(130px,1fr)); gap:10px; }
  .kpi .v{ font-size:22px; }
  section[data-testid="stSidebar"]{ min-width:270px !important; width:270px !important; }
}

/* Smartphone (até 768px): colunas empilham, fontes e espaçamentos reduzidos */
@media (max-width: 768px){
  /* Trava o overflow horizontal que cortava o texto na esquerda */
  html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"]{
       overflow-x:hidden !important; max-width:100vw !important; }
  .block-container{ padding-left:.8rem !important; padding-right:.8rem !important;
       padding-top:3rem !important; max-width:100% !important; }
  /* 2 cartões por linha — evita quebra de palavra (ex.: "REJEITADO") */
  .kpi-grid{ grid-template-columns:repeat(2,1fr) !important; gap:8px; }
  .kpi{ padding:10px 12px; border-radius:13px; }
  .kpi .v{ font-size:18px; word-break:keep-all; overflow-wrap:normal; }
  .kpi .l{ font-size:10px; letter-spacing:.5px; }
  .logo{ font-size:28px; }
  /* Empilha colunas (gráfico + resumo, filtros, etc.) em tela estreita */
  [data-testid="stHorizontalBlock"]{ flex-wrap:wrap !important; gap:.4rem !important; }
  [data-testid="stHorizontalBlock"] > [data-testid="column"]{
       flex:1 1 100% !important; width:100% !important; min-width:100% !important; }
  /* Sidebar (overlay no celular): largura confortável sem empurrar o conteúdo */
  section[data-testid="stSidebar"][aria-expanded="true"]{ min-width:78vw !important; width:78vw !important; }
  /* Tabelas/editores roláveis horizontalmente sem quebrar o layout */
  [data-testid="stDataFrame"], [data-testid="stDataEditor"]{ overflow-x:auto; }
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

DECIMALS = db.get_decimals()


def fmt(analyte, v, nd=None):
    if v is None or v == "":
        return "—"
    nd = DECIMALS.get(analyte, 2) if nd is None else nd
    return f"{v:.{nd}f}"


def pct(v, nd=2):
    return "—" if v in (None, "") else f"{v*100:.{nd}f}%"


# --------------------------------------------------------------------------- #
# Login
# --------------------------------------------------------------------------- #
def login_screen():
    c1, c2, c3 = st.columns([1, 1.1, 1])
    with c2:
        st.markdown("<div style='text-align:center;margin-top:6vh'>"
                    "<div class='logo'>QC-INI</div>"
                    "<div class='small'>Quality Control · Controle de Qualidade Interno</div></div>",
                    unsafe_allow_html=True)
        st.write("")
        with st.form("login"):
            login = st.text_input("Usuário", placeholder="INILAB")
            senha = st.text_input("Senha", type="password", placeholder="••••••")
            ok = st.form_submit_button("ENTRAR", use_container_width=True)
        if ok:
            user = db.authenticate(login.strip(), senha)
            if user:
                st.session_state["user"] = user
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")
        # Dica de credenciais só em ambiente local (NUNCA em produção/nuvem)
        if __import__("os").environ.get("QCLAB_SHOW_DEMO_HINT"):
            st.markdown("<div class='small' style='text-align:center'>Acesso de teste: "
                        "<b>INILAB</b> / <b>TESTE05</b></div>", unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Coleta de séries + avaliação
# --------------------------------------------------------------------------- #
def carregar_series(analyte, lote=None, months=None, year=None):
    """
    Retorna (series_validas, series_completas, refs, mapa_result_id) por nível.
    Filtros opcionais:
      lote   : restringe a um lote específico
      months : conjunto de meses 'AAAA-MM' (mês único ou combinados)
      year   : ano 'AAAA'
    """
    series_valid, series_all, refs, ids = {}, {}, {}, {}

    def _mantem(r):
        d = r["run_date"] or ""
        if lote and r["lote"] != lote:
            return False
        if months and d[:7] not in months:
            return False
        if year and d[:4] != year:
            return False
        return True

    for lvl in (1, 2, 3):
        rf = db.get_reference(analyte, lvl)
        if rf:
            refs[lvl] = (rf[0]["mean"], rf[0]["sd"])
        rows_all = [r for r in db.get_results(analyte, lvl, only_valid=False) if _mantem(r)]
        if rows_all:
            series_all[lvl] = [(r["seq"], r["value"]) for r in rows_all]
            series_valid[lvl] = [(r["seq"], r["value"]) for r in rows_all if not r["is_nc"]]
            ids[lvl] = {r["seq"]: r for r in rows_all}
    return series_valid, series_all, refs, ids


# --------------------------------------------------------------------------- #
# Painel (dashboard)
# --------------------------------------------------------------------------- #
def page_dashboard():
    cfg = db.get_config()
    st.markdown(f"### Painel · {cfg.get('setor','')}")
    st.markdown(f"<span class='small'>{cfg.get('instituicao','')} · "
                f"{cfg.get('equipamento','')} · Controle {cfg.get('controle','')}</span>",
                unsafe_allow_html=True)

    analytes = db.get_analytes(with_reference=True) or db.get_analytes()
    c1, c2, c3 = st.columns([2, 1, 1])
    analyte = c1.selectbox("Analito (ensaio)", analytes, index=0)
    level = c2.selectbox("Nível de controle", [1, 2, 3],
                         format_func=lambda x: f"Nível {x}")
    eixo = c3.selectbox("Eixo X", ["Sequência", "Data"])

    # ---- Filtros de Período e Lote ----
    todos_rows = db.query_results(analyte=analyte)
    datas = sorted({r["run_date"] for r in todos_rows if r["run_date"]})
    meses_disp = sorted({d[:7] for d in datas}, reverse=True)
    anos_disp = sorted({d[:4] for d in datas}, reverse=True)
    lotes_disp = sorted({r["lote"] for r in todos_rows if r["lote"]}, reverse=True)

    f1, f2, f3 = st.columns([1.1, 1.5, 1.1])
    periodo = f1.selectbox("Período", ["Todos", "Por mês", "Meses combinados", "Por ano"])
    months_sel = year_sel = None
    if periodo == "Por mês":
        m = f2.selectbox("Mês (AAAA-MM)", meses_disp or ["—"])
        months_sel = {m} if meses_disp else None
    elif periodo == "Meses combinados":
        sel = f2.multiselect("Meses (AAAA-MM)", meses_disp,
                             default=meses_disp[:2] if meses_disp else [])
        months_sel = set(sel) if sel else None
    elif periodo == "Por ano":
        year_sel = f2.selectbox("Ano", anos_disp or ["—"])
    lote_sel = f3.selectbox("Lote", ["Todos"] + lotes_disp)
    lote = None if lote_sel == "Todos" else lote_sel

    series_valid, series_all, refs, ids = carregar_series(
        analyte, lote=lote, months=months_sel, year=year_sel)
    if level not in refs:
        st.warning("Sem média/DP de referência cadastrados para este analito/nível.")
        return
    if not ids.get(level):
        st.info("Nenhum resultado para os filtros selecionados (período/lote). "
                "Ajuste os filtros acima.")
        return

    ref_mean, ref_sd = refs[level]
    # Conjunto de seqs marcadas como NC (por nível) — contam na janela mas
    # não geram rejeição nova.
    nc_by_level = {lvl: {seq for seq, row in ids.get(lvl, {}).items() if row["is_nc"]}
                   for lvl in ids}
    # avaliação Westgard sobre a SÉRIE CRONOLÓGICA COMPLETA (inclui NC), para
    # que as regras de tendência (6X, 3-1S, etc.) contem resultados realmente
    # consecutivos.
    aval = qc.evaluate_levels(series_all, refs, setor=cfg.get("setor", ""),
                              nc_by_level=nc_by_level)
    pts = aval.get(level, [])

    # estatística + sigma
    vals = [v for _, v in series_valid.get(level, [])]
    # Viés: Indicador de Bias Anual (Controle Externo) quando houver rodadas;
    # senão, fallback para o viés isolado antigo (external_qc).
    bias_anual, bias_year, bias_nrod = annual_bias_for(analyte)
    if bias_anual is not None:
        bias = bias_anual / 100.0                  # % -> fração
        bias_origem = f"Indicador Anual {bias_year} · {bias_nrod} rodada(s)"
    else:
        bias = abs(db.get_external_bias(analyte, level) or 0.0)
        bias_origem = "CQ externo (valor isolado)"
    stats = qc.build_stats(analyte, level, vals, bias, db.get_spec(analyte))

    # --------- KPIs ----------
    def card(label, value, cls=""):
        return f"<div class='kpi {cls}'><div class='v'>{value}</div><div class='l'>{label}</div></div>"

    n_rej = sum(1 for p in pts if p.status == "REJEITADO" and not p.is_nc)
    status_geral = ("REJEITADO", "bad") if n_rej else \
                   ("ALERTA", "warn") if any(p.status == "ALERTA" and not p.is_nc for p in pts) else ("OK", "ok")
    sigma_cls = "ok" if (stats.sigma or 0) >= 4 else "warn" if (stats.sigma or 0) >= 3 else "bad"

    html = "<div class='kpi-grid'>"
    html += card("Status Westgard", status_geral[0], status_geral[1])
    html += card("Média observada", fmt(analyte, stats.mean_obs))
    html += card("CV%", pct(stats.cv_obs), stats.cv_status.lower().replace("não", "bad").replace("ok", "ok"))
    html += card("Six Sigma", f"{stats.sigma:.2f}" if stats.sigma else "—", sigma_cls)
    html += card("Erro Total Permitido", f"{pct(stats.etp)} · {stats.etp_fonte}")
    html += card("Erro Sistemático - Bias", pct(stats.es))
    html += card("Erro Aleatório", pct(stats.ea))
    html += card("Erro Total", pct(stats.et))
    html += card("Limite CV CLIA", pct(stats.cv_limite_clia))
    html += card("Limite CV VB", pct(stats.cv_limite_vb))
    html += card("Pontos (n)", str(stats.n))
    html += card("Rejeições", str(n_rej), "bad" if n_rej else "ok")
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)
    st.caption(f"Erro Sistemático (viés) · origem: {bias_origem} · "
               f"ETp: {stats.etp_fonte} · fator do erro aleatório = 1,65.")

    st.write("")
    gcol, scol = st.columns([2.4, 1])

    # --------- Gráfico de Levey-Jennings ----------
    with gcol:
        st.markdown(f"**Gráfico de Levey-Jennings — {analyte} · Nível {level}**")
        lim = qc.control_limits(ref_mean, ref_sd)
        # usa os resultados já filtrados (período/lote)
        rows = [ids[level][s] for s in sorted(ids[level])]
        seqs = [r["seq"] for r in rows]
        x = seqs if eixo == "Sequência" else [r["run_date"] for r in rows]
        status_by_seq = {p.seq: p for p in pts}

        colors, texts = [], []
        for r in rows:
            if r["is_nc"]:
                colors.append("#9aa7a4")
            else:
                p = status_by_seq.get(r["seq"])
                s = p.status if p else "OK"
                colors.append({"OK": "#46d39a", "ALERTA": "#f4c430", "REJEITADO": "#ff5d6c"}.get(s, "#46d39a"))
            p = status_by_seq.get(r["seq"])
            flagtxt = ", ".join(p.flags) if (p and p.flags) else "OK"
            ztxt = f"{p.z:+.2f}s" if (p and p.z is not None) else ""
            texts.append(f"#{r['seq']} · {r['run_date']}<br>valor {fmt(analyte, r['value'])} ({ztxt})<br>{flagtxt}")

        fig = go.Figure()
        # (chave, cor, rótulo legível)
        band = [("+3s", "#ff5d6c", "+3 DP"), ("+2s", "#f4c430", "+2 DP"),
                ("+1s", "#46d39a", "+1 DP"), ("media", "#5fe3c2", "Média"),
                ("-1s", "#46d39a", "−1 DP"), ("-2s", "#f4c430", "−2 DP"),
                ("-3s", "#ff5d6c", "−3 DP")]
        for key, col, rotulo in band:
            valor = fmt(analyte, lim[key])
            fig.add_hline(y=lim[key], line=dict(color=col, width=2 if key == "media" else 1,
                          dash="solid" if key in ("media", "+3s", "-3s") else "dash"),
                          annotation_text=f"{rotulo} = {valor}", annotation_position="right",
                          annotation_font_color=col, annotation_font_size=12, opacity=.85)
        fig.add_trace(go.Scatter(x=x, y=[r["value"] for r in rows], mode="lines+markers",
                      line=dict(color="#cfeee7", width=1.4),
                      marker=dict(size=10, color=colors, line=dict(color="#0f2a2e", width=1)),
                      text=texts, hoverinfo="text", name=analyte))
        fig.update_layout(height=430, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                          font_color="#e7f3f1", margin=dict(l=10, r=120, t=10, b=10),
                          showlegend=False)
        fig.update_yaxes(showgrid=False, zeroline=False)
        fig.update_xaxes(showgrid=False, zeroline=False, title=eixo)
        st.plotly_chart(fig, use_container_width=True)

    # --------- Resumo WR + regras ----------
    with scol:
        st.markdown("**Resumo de Westgard (todos os níveis)**")
        items = ""
        for lvl in (1, 2, 3):
            for p in aval.get(lvl, []):
                if p.status != "OK" and not p.is_nc:
                    cls = "bad" if p.status == "REJEITADO" else "warn"
                    items += (f"<div class='row'>{analyte}-N{lvl} #{p.seq} "
                              f"<span class='badge {cls}'>{', '.join(p.flags)}</span></div>")
        if not items:
            items = "<div class='row'>Sem violações nos níveis avaliados ✅</div>"
        st.markdown(f"<div class='wrlist'>{items}</div>", unsafe_allow_html=True)

    # --------- Tabela de regras por nível ----------
    st.markdown("**Análise das regras de Westgard por nível**")

    # Colunas (código interno → rótulo) conforme o setor
    is_hemat = "HEMAT" in cfg.get("setor", "").upper() or len(refs) == 3
    if is_hemat:
        cols_regras = [("1-3S", "1-3s"), ("2of3-2S", "2 de 3-2s"),
                       ("R4S", "R4s"), ("3-1S", "3-1s"), ("6X", "6x")]
        niveis = sorted(refs.keys())
    else:
        cols_regras = [("1-3S", "1-3s"), ("2-2S", "2-2s"),
                       ("R4S", "R4s"), ("4-1S", "4-1s"), ("8X", "8x")]
        niveis = sorted(refs.keys())

    linhas = []
    for lvl in niveis:
        rs = qc.summarize_rules(aval.get(lvl, []))
        linha = {"Nível": f"N{lvl}"}
        for code, label in cols_regras:
            linha[label] = rs.get(code, "OK")
        linhas.append(linha)
    dfr = pd.DataFrame(linhas).set_index("Nível")

    def color_cell(v):
        return "color:#ff5d6c;font-weight:700" if v != "OK" else "color:#46d39a"
    st.dataframe(dfr.style.map(color_cell), use_container_width=True)


# --------------------------------------------------------------------------- #
# Lançar resultados
# --------------------------------------------------------------------------- #
def page_lancar():
    st.markdown("### Lançar / interfacear resultados")
    analytes = db.get_analytes(with_reference=True) or db.get_analytes()
    tab1, tab2, tab3 = st.tabs(["➕ Nova corrida (manual)", "📥 Importar arquivo",
                                "🗂️ Gerenciar / excluir"])

    with tab1:
        level = st.selectbox("Nível de controle", [1, 2, 3], format_func=lambda x: f"Nível {x}", key="lvl_man")
        seq = db.next_seq(level)
        c1, c2 = st.columns(2)
        d = c1.date_input("Data", value=date.today())
        t = c2.time_input("Hora", value=datetime.now().time())
        st.markdown(f"<span class='small'>Próxima sequência para o Nível {level}: "
                    f"<b>#{seq}</b></span>", unsafe_allow_html=True)

        with st.form("nova_corrida"):
            st.write("Informe os valores do controle (deixe em branco o que não se aplica):")
            cols = st.columns(4)
            valores = {}
            for i, a in enumerate(analytes):
                valores[a] = cols[i % 4].text_input(a, key=f"v_{a}")
            ok = st.form_submit_button("Salvar corrida")
        if ok:
            n = 0
            for a, v in valores.items():
                v = (v or "").replace(",", ".").strip()
                if v:
                    try:
                        rf = db.get_reference(a, level)
                        lote = rf[0]["lote"] if rf else ""
                        db.insert_result(level, lote, seq, d.isoformat(),
                                         t.strftime("%H:%M:%S"), a, float(v))
                        n += 1
                    except ValueError:
                        st.warning(f"Valor inválido em {a}: {v}")
            st.success(f"Corrida #{seq} (Nível {level}) salva — {n} analitos.")
            st.rerun()

    with tab2:
        st.write("Importe um CSV no formato longo com as colunas: "
                 "`level, seq, run_date, analyte, value` (run_time é opcional).")
        modelo = pd.DataFrame({"level": [1, 1], "seq": [33, 33],
                               "run_date": ["2025-12-01", "2025-12-01"],
                               "analyte": ["WBC", "RBC"], "value": [3.01, 2.30]})
        st.download_button("Baixar modelo CSV", modelo.to_csv(index=False).encode(),
                           "modelo_qc.csv", "text/csv")
        up = st.file_uploader("Arquivo CSV / Excel", type=["csv", "xlsx"])
        if up:
            df = pd.read_csv(up) if up.name.endswith(".csv") else pd.read_excel(up)
            st.dataframe(df.head(20), use_container_width=True)
            if st.button("Confirmar importação"):
                n = 0
                for _, r in df.iterrows():
                    try:
                        db.insert_result(int(r["level"]), str(r.get("lote", "")), int(r["seq"]),
                                         str(r["run_date"]), str(r.get("run_time", "00:00:00")),
                                         str(r["analyte"]).strip(), float(r["value"]))
                        n += 1
                    except Exception as e:  # noqa
                        st.warning(f"Linha ignorada: {e}")
                st.success(f"{n} resultados importados.")
                st.rerun()

    # ----------------------------------------------------------------- #
    # Aba 3: Gerenciar / excluir resultados (espelho do banco de dados)
    # ----------------------------------------------------------------- #
    with tab3:
        st.markdown("#### Gerenciar resultados — visualização e exclusão")
        st.markdown("<span class='small'>Tabela espelho do banco. Use os filtros, "
                    "marque os resultados na coluna <b>Excluir</b> e confirme. "
                    "A exclusão é permanente.</span>", unsafe_allow_html=True)

        # Datas/anos/meses disponíveis (a partir de todos os resultados)
        todos = db.query_results()
        datas = sorted({r["run_date"] for r in todos if r["run_date"]})
        anos = sorted({d[:4] for d in datas}, reverse=True)
        meses = sorted({d[:7] for d in datas}, reverse=True)

        # ---- Linha 1 de filtros: período ----
        f1, f2, f3 = st.columns([1.2, 1, 1])
        periodo = f1.selectbox("Período",
                               ["Todos", "Por mês", "Por semestre", "Por ano", "Personalizado"])

        date_from = date_to = None
        meses_sel = None
        if periodo == "Por mês":
            meses_sel = st.multiselect("Selecione o(s) mês(es) (AAAA-MM)",
                                       meses, default=meses[:1] if meses else [])
        elif periodo == "Por semestre":
            s1, s2 = st.columns(2)
            ano_s = s1.selectbox("Ano", anos or [str(date.today().year)], key="del_ano_sem")
            sem = s2.selectbox("Semestre", ["1º semestre (Jan–Jun)", "2º semestre (Jul–Dez)"])
            if sem.startswith("1"):
                date_from, date_to = f"{ano_s}-01-01", f"{ano_s}-06-30"
            else:
                date_from, date_to = f"{ano_s}-07-01", f"{ano_s}-12-31"
        elif periodo == "Por ano":
            ano_a = st.selectbox("Ano", anos or [str(date.today().year)], key="del_ano")
            date_from, date_to = f"{ano_a}-01-01", f"{ano_a}-12-31"
        elif periodo == "Personalizado":
            p1, p2 = st.columns(2)
            df_ini = p1.date_input("De", value=date.today().replace(day=1), key="del_de")
            df_fim = p2.date_input("Até", value=date.today(), key="del_ate")
            date_from, date_to = df_ini.isoformat(), df_fim.isoformat()

        # ---- Linha 2 de filtros: lote / analito / nível ----
        lotes = db.get_lotes()
        lote_sel = f2.selectbox("Lote", ["Todos"] + lotes, key="del_lote")
        nivel_sel = f3.selectbox("Nível", ["Todos", 1, 2, 3], key="del_nivel",
                                 format_func=lambda x: x if x == "Todos" else f"Nível {x}")
        analito_sel = st.selectbox("Analito", ["Todos"] + analytes, key="del_analito")

        # ---- Consulta com filtros ----
        rows = db.query_results(
            analyte=None if analito_sel == "Todos" else analito_sel,
            level=None if nivel_sel == "Todos" else nivel_sel,
            lote=None if lote_sel == "Todos" else lote_sel,
            date_from=date_from, date_to=date_to,
        )
        if meses_sel is not None:
            mset = set(meses_sel)
            rows = [r for r in rows if (r["run_date"] or "")[:7] in mset]

        st.markdown(f"<span class='small'>{len(rows)} resultado(s) encontrado(s).</span>",
                    unsafe_allow_html=True)

        if not rows:
            st.info("Nenhum resultado para os filtros selecionados.")
        else:
            df_show = pd.DataFrame([{
                "Excluir": False,
                "ID": r["result_id"],
                "Data": r["run_date"],
                "Hora": r["run_time"],
                "Nível": r["level"],
                "Seq": r["seq"],
                "Analito": r["analyte"],
                "Valor": r["value"],
                "Lote": r["lote"],
                "NC": "Sim" if r["is_nc"] else "",
            } for r in rows])

            edited = st.data_editor(
                df_show,
                use_container_width=True, hide_index=True, height=380,
                column_config={
                    "Excluir": st.column_config.CheckboxColumn("Excluir", default=False),
                    "ID": st.column_config.NumberColumn("ID", disabled=True),
                    "Data": st.column_config.TextColumn("Data", disabled=True),
                    "Hora": st.column_config.TextColumn("Hora", disabled=True),
                    "Nível": st.column_config.NumberColumn("Nível", disabled=True),
                    "Seq": st.column_config.NumberColumn("Seq", disabled=True),
                    "Analito": st.column_config.TextColumn("Analito", disabled=True),
                    "Valor": st.column_config.NumberColumn("Valor", disabled=True),
                    "Lote": st.column_config.TextColumn("Lote", disabled=True),
                    "NC": st.column_config.TextColumn("NC", disabled=True),
                },
                key="editor_excluir",
            )

            sel_ids = edited.loc[edited["Excluir"], "ID"].tolist()
            cdel, cinfo = st.columns([1, 2])
            if cdel.button(f"🗑️ Excluir selecionados ({len(sel_ids)})",
                           disabled=not sel_ids, key="btn_excluir"):
                st.session_state["pending_delete"] = sel_ids

            # ---- Confirmação ----
            pend = st.session_state.get("pending_delete")
            if pend:
                st.warning(f"⚠️ Você confirma a exclusão **permanente** de "
                           f"{len(pend)} resultado(s)? Esta ação não pode ser desfeita.")
                cc1, cc2 = st.columns(2)
                if cc1.button("✅ Sim, excluir definitivamente", key="confirm_del"):
                    n = db.delete_results(pend)
                    st.session_state.pop("pending_delete", None)
                    st.success(f"{n} resultado(s) excluído(s).")
                    st.rerun()
                if cc2.button("❌ Cancelar", key="cancel_del"):
                    st.session_state.pop("pending_delete", None)
                    st.rerun()


# --------------------------------------------------------------------------- #
# Não conformidades
# --------------------------------------------------------------------------- #
def page_nc():
    st.markdown("### Não conformidades")
    st.markdown("<span class='small'>Pontos sinalizados pelo motor de Westgard. "
                "Marque como não conforme com assinatura eletrônica (login).</span>",
                unsafe_allow_html=True)
    analytes = db.get_analytes(with_reference=True) or db.get_analytes()
    analyte = st.selectbox("Analito", analytes)
    cfg = db.get_config()
    series_valid, series_all, refs, ids = carregar_series(analyte)
    nc_by_level = {lvl: {seq for seq, row in ids.get(lvl, {}).items() if row["is_nc"]}
                   for lvl in ids}
    aval = qc.evaluate_levels(series_all, refs, setor=cfg.get("setor", ""),
                              nc_by_level=nc_by_level)
    flagged = []
    for lvl in (1, 2, 3):
        for p in aval.get(lvl, []):
            if p.status != "OK" and not p.is_nc:
                row = ids[lvl][p.seq]
                flagged.append((lvl, p, row))
    # também já marcados como NC
    nc_marked = []
    for lvl in (1, 2, 3):
        for seq, row in ids.get(lvl, {}).items():
            if row["is_nc"]:
                nc_marked.append((lvl, row))

    st.markdown("#### Sinalizados pelas regras")
    if not flagged:
        st.info("Nenhum ponto sinalizado para este analito.")
    for lvl, p, row in flagged:
        cls = "bad" if p.status == "REJEITADO" else "warn"
        c1, c2 = st.columns([3, 1])
        c1.markdown(f"N{lvl} · corrida #{p.seq} · {row['run_date']} · valor "
                    f"**{fmt(analyte, p.value)}** ({p.z:+.2f}s) "
                    f"<span class='badge {cls}'>{', '.join(p.flags)}</span>",
                    unsafe_allow_html=True)
        if row["is_nc"]:
            c2.markdown("<span class='badge bad'>JÁ É NC</span>", unsafe_allow_html=True)
        else:
            with c2.popover("Marcar NC"):
                motivo = st.text_input("Motivo / ação", key=f"m_{row['result_id']}")
                pw = st.text_input("Assinatura (senha)", type="password", key=f"p_{row['result_id']}")
                if st.button("Confirmar NC", key=f"b_{row['result_id']}"):
                    if db.authenticate(st.session_state["user"]["login"], pw):
                        db.set_nc(row["result_id"], True, st.session_state["user"]["login"], motivo)
                        st.success("Marcado como não conforme.")
                        st.rerun()
                    else:
                        st.error("Assinatura inválida.")

    st.markdown("#### Já marcados como não conformes")
    if nc_marked:
        df = pd.DataFrame([{"Nível": f"N{l}", "Corrida": r["seq"], "Data": r["run_date"],
                            "Valor": fmt(analyte, r["value"]), "Motivo": r["nc_reason"],
                            "Assinado por": r["nc_by"], "Em": r["nc_at"]} for l, r in nc_marked])
        st.dataframe(df, use_container_width=True, hide_index=True)
        rid = st.selectbox("Reverter NC (corrida)", [r["result_id"] for _, r in nc_marked],
                           format_func=lambda x: f"id {x}")
        if st.button("Reverter para conforme"):
            db.set_nc(rid, False, st.session_state["user"]["login"])
            st.rerun()
    else:
        st.caption("Nenhum.")


# --------------------------------------------------------------------------- #
# Média / DP de referência
# --------------------------------------------------------------------------- #
def page_referencia():
    st.markdown("### Média e DP de referência")
    st.markdown("<span class='small'>Premissa: média e DP exigem no mínimo 20 dados. "
                "Limites de controle = média ± 1/2/3 DP.</span>", unsafe_allow_html=True)
    level = st.selectbox("Nível", [1, 2, 3], format_func=lambda x: f"Nível {x}")
    refs = db.get_reference(level=level)
    by = {r["analyte"]: r for r in refs}
    rows = []
    for a in db.get_analytes():
        r = by.get(a)
        if r:
            rows.append({"Analito": a, "Média": r["mean"], "DP": r["sd"],
                         "LI (-3DP)": r["mean"] - 3 * r["sd"], "LS (+3DP)": r["mean"] + 3 * r["sd"]})
    df = pd.DataFrame(rows)
    edited = st.data_editor(df, use_container_width=True, hide_index=True,
                            disabled=["Analito", "LI (-3DP)", "LS (+3DP)"], key="ref_ed")
    if st.button("Salvar alterações de referência"):
        for _, r in edited.iterrows():
            db.update_reference(r["Analito"], level, float(r["Média"]), float(r["DP"]))
        st.success("Referências atualizadas.")
        st.rerun()


# --------------------------------------------------------------------------- #
# Especificação da qualidade / Sigma
# --------------------------------------------------------------------------- #
def page_specs():
    st.markdown("### Especificação da qualidade · Sigma e Erro Total")
    st.markdown("<span class='small'>ETp = valor configurado por ensaio (ou TEa CLIA/Variação "
                "Biológica). Viés = Indicador de Bias Anual (Controle Externo) quando houver "
                "rodadas. ET = CV%×1,65 + viés · Sigma = (ETp − viés) / CV%.</span>",
                unsafe_allow_html=True)
    level = st.selectbox("Nível", [1, 2, 3], format_func=lambda x: f"Nível {x}", key="spec_lvl")
    rows = []
    for a in db.get_analytes(with_reference=True):
        vals = [r["value"] for r in db.get_results(a, level)]
        if not vals:
            continue
        bias_anual, _yr, _nr = annual_bias_for(a)
        if bias_anual is not None:
            bias = bias_anual / 100.0
            origem = "Bias Anual"
        else:
            bias = abs(db.get_external_bias(a, level) or 0.0)
            origem = "CQ isolado"
        s = qc.build_stats(a, level, vals, bias, db.get_spec(a))
        rows.append({"Analito": a, "n": s.n, "Média": round(s.mean_obs, 4) if s.mean_obs else None,
                     "CV%": pct(s.cv_obs), "ETP": pct(s.etp), "Fonte": s.etp_fonte,
                     "ES (viés)": pct(s.es), "Origem viés": origem, "EA": pct(s.ea), "ET": pct(s.et),
                     "Sigma": round(s.sigma, 2) if s.sigma else None,
                     "CV status": s.cv_status})
    df = pd.DataFrame(rows)

    def sig_color(v):
        try:
            v = float(v)
        except (TypeError, ValueError):
            return ""
        return "color:#46d39a;font-weight:700" if v >= 4 else \
               "color:#f4c430;font-weight:700" if v >= 3 else "color:#ff5d6c;font-weight:700"
    sty = df.style.map(sig_color, subset=["Sigma"])
    st.dataframe(sty, use_container_width=True, hide_index=True)
    st.caption("Sigma ≥ 6 excelente · ≥ 4 bom · 3–4 marginal · < 3 inaceitável.")


# --------------------------------------------------------------------------- #
# Integração IQC↔EQC — helpers de CV% e ETp (reusados no painel/Sigma)
# --------------------------------------------------------------------------- #
def cv_pct_iqc(analyte, level):
    """CV% acumulado do Controle Interno (em %) para o analito/nível; None se sem dados."""
    vals = [r["value"] for r in db.get_results(analyte, level)]
    if not vals:
        return None
    _n, _m, _sd, cv = qc.compute_stats(vals)
    return cv * 100.0 if cv is not None else None


def etp_pct_for(analyte):
    """ETp (%) e origem: usa o configurado; senão deriva de CLIA/VB. (None,'-') se nada."""
    etp_pct, src = db.get_etp(analyte)
    if etp_pct is not None:
        return float(etp_pct), (src or "Configurado")
    spec = db.get_spec(analyte)
    if spec.get("clia_tea") not in (None, "", "-"):
        return float(spec["clia_tea"]) * 100.0, "CLIA"
    _i, _v, te_vb = qc.vb_specs(spec.get("cvw"), spec.get("cvg"), spec.get("perf"))
    if te_vb is not None:
        return te_vb * 100.0, "VB"
    return None, "-"


def annual_bias_for(analyte, year=None):
    """
    Indicador de Bias Anual (%) do EQC para o analito. Se year=None, usa o ano
    EQC mais recente com dados. Retorna (indicador_pct, year, n_rodadas) ou
    (None, None, 0).
    """
    if year is None:
        anos = [y for y in db.eqc_years()
                if db.eqc_annual_indicator(analyte, y)["indicador"] is not None]
        if not anos:
            return None, None, 0
        year = anos[0]
    info = db.eqc_annual_indicator(analyte, year)
    return info["indicador"], year, info["n_rodadas"]


# --------------------------------------------------------------------------- #
# Controle Externo da Qualidade (EQC)
# --------------------------------------------------------------------------- #
PROVIDERS = ["CAP", "Controllab", "PNCQ", "SBPC/ML", "ControlLab", "Outro"]
EQC_STATUS = ["Aceitável", "Alerta", "Inaceitável", "Fora de critério"]
ETP_SOURCES = ["CLIA", "Variação biológica", "Fabricante",
               "Especificação interna validada", "Outra"]


def _samples_from_editor(edited):
    """Converte o DataFrame do editor em (lista_amostras, lista_bias, |bias|_rodada)."""
    amostras, biases = [], []
    for _, row in edited.iterrows():
        lab = row.get("Resultado lab")
        peer = row.get("Média grupo")
        b = qc.sample_bias_pct(lab, peer)
        if qc._num(lab) is None and qc._num(peer) is None:
            continue  # linha vazia
        amostras.append({
            "sample_label": row.get("Amostra") or f"{len(amostras)+1:02d}",
            "lab_value": qc._num(lab), "peer_mean": qc._num(peer),
            "peer_sd": qc._num(row.get("DP grupo")),
            "lower_limit": qc._num(row.get("Lim. inferior")),
            "upper_limit": qc._num(row.get("Lim. superior")),
        })
        biases.append(b)
    return amostras, biases, qc.round_abs_bias(biases)


def _editor_default(n=5):
    return pd.DataFrame([
        {"Amostra": f"{i:02d}", "Resultado lab": None, "Média grupo": None,
         "DP grupo": None, "Lim. inferior": None, "Lim. superior": None}
        for i in range(1, n + 1)])


def page_controle_externo():
    cfg = db.get_config()
    st.markdown("### Controle Externo da Qualidade")
    st.markdown("<span class='small'>Ensaios de proficiência (CAP, Controllab, PNCQ…). "
                "Até 6 rodadas/ano por ensaio; cada rodada com várias amostras. O "
                "<b>Indicador de Bias Anual</b> alimenta o Erro Total e a Métrica Sigma.</span>",
                unsafe_allow_html=True)

    tab_cad, tab_ger, tab_des, tab_ens = st.tabs(
        ["➕ Cadastrar rodada", "🗂️ Consultar / gerenciar",
         "📊 Desempenho anual", "🧬 Ensaios & ETp"])

    # ===================================================================== #
    # Aba 1 — Cadastrar rodada
    # ===================================================================== #
    with tab_cad:
        areas = db.get_areas() or seed_areas()
        c1, c2, c3 = st.columns(3)
        area = c1.selectbox("Área técnica", areas, key="eqc_area")
        analitos_area = db.get_analytes(area=area)
        if not analitos_area:
            c2.info("Nenhum analito nesta área. Cadastre em '🧬 Ensaios & ETp'.")
            analito = None
        else:
            analito = c2.selectbox("Ensaio / analito", analitos_area, key="eqc_analito")
        ano_default = (db.eqc_years() or [date.today().year])[0]
        ano = c3.number_input("Ano de referência", min_value=2000, max_value=2100,
                              value=int(ano_default), step=1, key="eqc_ano")

        if analito:
            taken = sorted({rd["round_number"] for rd in
                            db.eqc_list_rounds(year=int(ano), analyte=analito)})
            n_rod = db.eqc_round_count(analito, int(ano))
            livres = [n for n in range(1, db.MAX_ROUNDS_POR_ANO + 1) if n not in taken]

            if n_rod >= db.MAX_ROUNDS_POR_ANO:
                st.warning(f"⚠️ Limite de {db.MAX_ROUNDS_POR_ANO} rodadas/ano já atingido "
                           f"para {analito} em {ano}. Exclua uma rodada para cadastrar outra.")
            else:
                st.caption(f"Rodadas já cadastradas em {ano}: "
                           f"{taken or '—'} · disponíveis: {livres}")
                d1, d2, d3 = st.columns(3)
                rnum = d1.selectbox("Nº da rodada", livres, key="eqc_rnum")
                rdata = d2.date_input("Data da rodada", value=date.today(), key="eqc_data")
                prov = d3.selectbox("Provedor", PROVIDERS, key="eqc_prov")
                if prov == "Outro":
                    prov = d3.text_input("Qual provedor?", key="eqc_prov_outro") or "Outro"

                e1, e2, e3 = st.columns(3)
                ainfo = db.get_analyte_info(analito)
                unidade = e1.text_input("Unidade", value=ainfo.get("unit", ""), key="eqc_unid")
                lote = e2.text_input("Lote / amostra", key="eqc_lote")
                status = e3.selectbox("Status de desempenho", EQC_STATUS, key="eqc_status")

                report_ref = st.text_input("Referência ao relatório (PDF)", key="eqc_pdf")
                notes = st.text_area("Observações técnicas", key="eqc_notes", height=70)

                st.markdown("**Amostras da rodada** (o CAP costuma trazer ~5 espécimes). "
                            "Preencha resultado do laboratório e média do grupo:")
                edited = st.data_editor(
                    _editor_default(5), num_rows="dynamic", hide_index=True,
                    use_container_width=True, key="eqc_samples_editor",
                    column_config={
                        "Amostra": st.column_config.TextColumn("Amostra"),
                        "Resultado lab": st.column_config.NumberColumn("Resultado lab", format="%.4f"),
                        "Média grupo": st.column_config.NumberColumn("Média grupo", format="%.4f"),
                        "DP grupo": st.column_config.NumberColumn("DP grupo", format="%.4f"),
                        "Lim. inferior": st.column_config.NumberColumn("Lim. inferior", format="%.4f"),
                        "Lim. superior": st.column_config.NumberColumn("Lim. superior", format="%.4f"),
                    })
                amostras, biases, round_bias = _samples_from_editor(edited)

                # Prévia dos cálculos (ao vivo)
                if amostras:
                    prev = pd.DataFrame([
                        {"Amostra": a["sample_label"],
                         "Resultado": a["lab_value"], "Média grupo": a["peer_mean"],
                         "Bias %": None if b is None else round(b, 3),
                         "|Bias| %": None if b is None else round(abs(b), 3)}
                        for a, b in zip(amostras, biases)])
                    cprev, cmet = st.columns([2, 1])
                    cprev.dataframe(prev, use_container_width=True, hide_index=True)
                    cmet.metric("|Bias| da rodada",
                                "—" if round_bias is None else f"{round_bias:.3f}%")
                    cmet.caption("Média dos |bias| das amostras válidas.")

                if st.button("💾 Salvar rodada", key="eqc_save", disabled=not amostras):
                    rid, err = db.eqc_add_round(
                        {"area": area, "analyte": analito, "year": int(ano),
                         "round_number": int(rnum), "round_date": rdata.isoformat(),
                         "provider": prov, "unit": unidade, "lote": lote,
                         "status": status, "notes": notes, "report_ref": report_ref},
                        amostras, user=st.session_state["user"]["login"])
                    if err:
                        st.error(err)
                    else:
                        st.success(f"Rodada {rnum} de {analito}/{ano} salva "
                                   f"(|bias| = {round_bias:.3f}%).")
                        st.rerun()

    # ===================================================================== #
    # Aba 2 — Consultar / gerenciar
    # ===================================================================== #
    with tab_ger:
        g1, g2, g3 = st.columns(3)
        f_ano = g1.selectbox("Ano", ["Todos"] + db.eqc_years(), key="ger_ano")
        f_area = g2.selectbox("Área", ["Todas"] + db.get_areas(), key="ger_area")
        analitos_f = (db.get_analytes(area=None if f_area == "Todas" else f_area))
        f_analito = g3.selectbox("Ensaio", ["Todos"] + analitos_f, key="ger_analito")
        g4, g5, g6 = st.columns(3)
        f_prov = g4.selectbox("Provedor", ["Todos"] + db.eqc_providers(), key="ger_prov")
        f_rod = g5.selectbox("Rodada", ["Todas"] + list(range(1, db.MAX_ROUNDS_POR_ANO + 1)),
                             key="ger_rod")
        f_lote = g6.selectbox("Lote", ["Todos"] + db.eqc_lotes(), key="ger_lote")

        rounds = db.eqc_list_rounds(
            year=None if f_ano == "Todos" else f_ano,
            area=None if f_area == "Todas" else f_area,
            analyte=None if f_analito == "Todos" else f_analito,
            provider=None if f_prov == "Todos" else f_prov,
            round_number=None if f_rod == "Todas" else f_rod,
            lote=None if f_lote == "Todos" else f_lote)

        st.caption(f"{len(rounds)} rodada(s) encontrada(s).")
        if rounds:
            tabela = pd.DataFrame([{
                "ID": rd["round_id"], "Ano": rd["year"], "Área": rd["area"],
                "Ensaio": rd["analyte"], "Rodada": rd["round_number"],
                "Data": rd["round_date"], "Provedor": rd["provider"],
                "Lote": rd["lote"], "Amostras": rd["n_samples"],
                "|Bias| %": None if rd["round_abs_bias"] is None else round(rd["round_abs_bias"], 3),
                "Status": rd["status"]} for rd in rounds])
            st.dataframe(tabela, use_container_width=True, hide_index=True)

            rotulos = {f'#{rd["round_id"]} · {rd["analyte"]} {rd["year"]} R{rd["round_number"]}':
                       rd["round_id"] for rd in rounds}
            sel = st.selectbox("Selecione uma rodada para ver amostras / excluir",
                               ["—"] + list(rotulos.keys()), key="ger_sel")
            if sel != "—":
                rid = rotulos[sel]
                samples = db.eqc_get_samples(rid)
                dfp = pd.DataFrame([{
                    "Amostra": s["sample_label"], "Resultado lab": s["lab_value"],
                    "Média grupo": s["peer_mean"], "DP grupo": s["peer_sd"],
                    "Bias %": None if s["bias_pct"] is None else round(s["bias_pct"], 3),
                    "|Bias| %": None if s["bias_pct"] is None else round(abs(s["bias_pct"]), 3)}
                    for s in samples])
                st.dataframe(dfp, use_container_width=True, hide_index=True)
                rd = db.eqc_get_round(rid)
                st.caption(f"Última alteração por {rd.get('updated_by') or rd.get('created_by')} "
                           f"em {rd.get('updated_at') or rd.get('created_at')}.")

                if st.button("🗑️ Excluir esta rodada", key="ger_del"):
                    st.session_state["eqc_pending_del"] = rid
                if st.session_state.get("eqc_pending_del") == rid:
                    st.warning("⚠️ Confirma a exclusão **permanente** desta rodada e suas amostras?")
                    cc1, cc2 = st.columns(2)
                    if cc1.button("✅ Sim, excluir", key="ger_del_yes"):
                        db.eqc_delete_round(rid, user=st.session_state["user"]["login"])
                        st.session_state.pop("eqc_pending_del", None)
                        st.success("Rodada excluída.")
                        st.rerun()
                    if cc2.button("❌ Cancelar", key="ger_del_no"):
                        st.session_state.pop("eqc_pending_del", None)
                        st.rerun()

    # ===================================================================== #
    # Aba 3 — Desempenho anual (integração IQC↔EQC)
    # ===================================================================== #
    with tab_des:
        anos = db.eqc_years() or [date.today().year]
        h1, h2 = st.columns(2)
        ano_d = h1.selectbox("Ano de referência", anos, key="des_ano")
        nivel_cv = h2.selectbox("Nível do CV% (Controle Interno)", [1, 2, 3],
                                index=1, format_func=lambda x: f"Nível {x}", key="des_nivel")
        st.caption("ET = CV% × 1,65 + Bias Anual · Sigma = (ETp − Bias Anual) / CV%")

        analitos_ano = db.eqc_analytes_with_rounds(int(ano_d))
        linhas, faltas_msgs = [], []
        for a in analitos_ano:
            bias, _yr, n_rod = annual_bias_for(a, int(ano_d))
            cvp = cv_pct_iqc(a, int(nivel_cv))
            etpp, etps = etp_pct_for(a)
            perf = qc.annual_performance(cvp, bias, etpp, etps)
            linhas.append({
                "Ensaio": a, "Rodadas": n_rod,
                "Bias Anual %": None if bias is None else round(bias, 3),
                "CV% (IQC)": None if cvp is None else round(cvp, 3),
                "ETp %": None if etpp is None else round(etpp, 2),
                "Origem ETp": etps,
                "ET %": None if perf["et"] is None else round(perf["et"], 3),
                "Sigma": None if perf["sigma"] is None else round(perf["sigma"], 2)})
            if perf["faltas"]:
                faltas_msgs.append(f"**{a}**: faltam " + ", ".join(perf["faltas"]))

        if not linhas:
            st.info("Sem rodadas de Controle Externo neste ano.")
        else:
            dfd = pd.DataFrame(linhas)

            def sig_color(v):
                try:
                    v = float(v)
                except (TypeError, ValueError):
                    return ""
                return ("color:#46d39a;font-weight:700" if v >= 4 else
                        "color:#f4c430;font-weight:700" if v >= 3 else
                        "color:#ff5d6c;font-weight:700")
            st.dataframe(dfd.style.map(sig_color, subset=["Sigma"]),
                         use_container_width=True, hide_index=True)
            st.caption("Sigma ≥ 6 excelente · ≥ 4 bom · 3–4 marginal · < 3 inaceitável.")

            # Melhor / pior desempenho (por Sigma)
            com_sigma = [l for l in linhas if l["Sigma"] is not None]
            if com_sigma:
                melhor = max(com_sigma, key=lambda l: l["Sigma"])
                pior = min(com_sigma, key=lambda l: l["Sigma"])
                m1, m2 = st.columns(2)
                m1.success(f"🏆 Melhor: **{melhor['Ensaio']}** · Sigma {melhor['Sigma']}")
                m2.error(f"⚠️ Pior: **{pior['Ensaio']}** · Sigma {pior['Sigma']}")

            if faltas_msgs:
                with st.expander("Ensaios com dados insuficientes"):
                    for m in faltas_msgs:
                        st.markdown("- " + m)

            # Evolução do bias por rodada
            st.markdown("**Evolução do bias por rodada**")
            a_sel = st.selectbox("Ensaio", analitos_ano, key="des_evo")
            info = db.eqc_annual_indicator(a_sel, int(ano_d))
            pr = [p for p in info["por_rodada"] if p["abs_bias"] is not None]
            if pr:
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=[f"R{p['round_number']}" for p in pr],
                    y=[p["abs_bias"] for p in pr], mode="lines+markers",
                    line=dict(color="#5fe3c2", width=2),
                    marker=dict(size=11, color="#5fe3c2"), name="|Bias| rodada"))
                if info["indicador"] is not None:
                    fig.add_hline(y=info["indicador"], line=dict(color="#f4c430", dash="dash"),
                                  annotation_text=f"Indicador Anual = {info['indicador']:.3f}%",
                                  annotation_position="right", annotation_font_color="#f4c430")
                fig.update_layout(height=320, paper_bgcolor="rgba(0,0,0,0)",
                                  plot_bgcolor="rgba(0,0,0,0)", font_color="#e7f3f1",
                                  margin=dict(l=10, r=130, t=10, b=10), showlegend=False)
                fig.update_yaxes(title="|Bias| %", gridcolor="#24474c")
                fig.update_xaxes(title="Rodada")
                st.plotly_chart(fig, use_container_width=True)

    # ===================================================================== #
    # Aba 4 — Ensaios & ETp (extensibilidade)
    # ===================================================================== #
    with tab_ens:
        st.markdown("#### Cadastrar novo ensaio / analito")
        st.caption("O sistema é extensível: cadastre novos parâmetros a qualquer momento.")
        n1, n2, n3, n4 = st.columns([1.2, 1.4, 1, 1])
        areas_existentes = db.get_areas() or seed_areas()
        nova_area_opt = n1.selectbox("Área", areas_existentes + ["➕ Nova área"], key="ens_area")
        if nova_area_opt == "➕ Nova área":
            nova_area_opt = n1.text_input("Nome da nova área", key="ens_area_nova")
        novo_nome = n2.text_input("Nome do ensaio", key="ens_nome")
        nova_unid = n3.text_input("Unidade", key="ens_unid")
        nova_dec = n4.number_input("Casas decimais", 0, 6, 2, key="ens_dec")
        if st.button("➕ Adicionar ensaio", key="ens_add"):
            if novo_nome.strip() and (nova_area_opt or "").strip():
                db.add_analyte(novo_nome, nova_area_opt, nova_unid, int(nova_dec))
                st.success(f"Ensaio '{novo_nome}' cadastrado em {nova_area_opt}.")
                st.rerun()
            else:
                st.error("Informe ao menos a área e o nome do ensaio.")

        st.divider()
        st.markdown("#### Erro Total Permitido (ETp) por ensaio")
        todos = db.get_analytes()
        p1, p2, p3 = st.columns([1.6, 1, 1.2])
        a_etp = p1.selectbox("Ensaio", todos, key="etp_analito")
        etp_atual, src_atual = db.get_etp(a_etp)
        val_etp = p2.number_input("ETp (%)", min_value=0.0, max_value=100.0,
                                  value=float(etp_atual) if etp_atual else 0.0,
                                  step=0.5, key="etp_val")
        idx_src = ETP_SOURCES.index(src_atual) if src_atual in ETP_SOURCES else 0
        src_etp = p3.selectbox("Origem do ETp", ETP_SOURCES, index=idx_src, key="etp_src")
        if st.button("💾 Salvar ETp", key="etp_save"):
            db.set_etp(a_etp, val_etp, src_etp)
            st.success(f"ETp de {a_etp} definido em {val_etp:.2f}% ({src_etp}).")
            st.rerun()

        st.divider()
        st.markdown("#### Ensaios cadastrados")
        linhas_ens = []
        for a in todos:
            info = db.get_analyte_info(a)
            etpp, etps = db.get_etp(a)
            linhas_ens.append({"Ensaio": a, "Área": info.get("area", ""),
                               "Unidade": info.get("unit", ""),
                               "ETp %": etpp, "Origem ETp": etps or "—"})
        st.dataframe(pd.DataFrame(linhas_ens), use_container_width=True, hide_index=True)


def seed_areas():
    """Áreas sugeridas quando o banco ainda não tem nenhuma."""
    try:
        return list(seed.EQC_AREAS)
    except Exception:
        return ["Hematologia", "Bioquímica", "Imunologia", "Coagulação"]


# --------------------------------------------------------------------------- #
# Configurações / Sobre
# --------------------------------------------------------------------------- #
def page_config():
    st.markdown("### Configurações do serviço")
    cfg = db.get_config()
    for k, v in cfg.items():
        st.text_input(k.replace("_", " ").title(), value=v, disabled=True)
    st.divider()
    st.markdown("#### Administração")
    if st.button("↺ Recriar banco de demonstração (apaga dados lançados)"):
        db.init_db(force=True, demo=True)
        st.success("Banco recriado com os dados de demonstração.")
        st.rerun()


# --------------------------------------------------------------------------- #
# App
# --------------------------------------------------------------------------- #
if "user" not in st.session_state:
    login_screen()
    st.stop()

user = st.session_state["user"]
with st.sidebar:
    st.markdown("<div class='logo' style='font-size:30px'>QC-INI</div>", unsafe_allow_html=True)
    st.markdown(f"<span class='small'>{user['name']} · {user['role']}</span>", unsafe_allow_html=True)
    st.divider()
    pagina = st.radio("Navegação", [
        "Painel", "Lançar resultados", "Não conformidades",
        "Controle Externo", "Média e DP", "Especificação da qualidade",
        "Configurações"],
        label_visibility="collapsed")
    st.divider()
    if st.button("Sair", use_container_width=True):
        del st.session_state["user"]
        st.rerun()

{
    "Painel": page_dashboard,
    "Lançar resultados": page_lancar,
    "Não conformidades": page_nc,
    "Controle Externo": page_controle_externo,
    "Média e DP": page_referencia,
    "Especificação da qualidade": page_specs,
    "Configurações": page_config,
}[pagina]()
