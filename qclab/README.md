# QC-INI · Controle de Qualidade Interno (protótipo SQL + Python)

Reconstrução em **Python + SQL** do seu sistema de CQ que hoje vive no Excel/VBA + Power BI.
Aplicação **web responsiva** (abre no navegador, funciona no celular) com login, gráfico de
Levey-Jennings, regras de Westgard automáticas, métricas Sigma / Erro Total e gestão de não
conformidades com assinatura eletrônica.

> **Login de teste:** `INILAB` / `TESTE05`  (perfil ANALISTA, pode assinar NC)
> Também existe `admin` / `admin`.

---

## O que já está implementado

| Módulo da planilha | No sistema |
|---|---|
| Login (macro VBA) | Tela de login com usuários no banco (senha em SHA-256), níveis de acesso 1=Técnico / 2=Analista |
| Aba "CONTROLE INTERNO" + análise sob cada ensaio | **Painel**: regras de Westgard (1-3s, 2 de 3-2s, R4s **entre níveis**, 3-1s, 9x) calculadas automaticamente por analito/nível |
| Gráficos no Power BI | **Levey-Jennings** interativo (Plotly) com média e ±1/2/3 DP e pontos coloridos por status (verde/amarelo/vermelho), além do "Resumo WR" |
| INSERIR M e DP | **Média e DP de referência** editável por nível (limites = média ± 3 DP) |
| INTERFACEAR / lançamento horizontal | **Lançar resultados**: nova corrida manual (todos os analitos) **+ importação CSV/Excel** |
| RES. NC | **Não conformidades**: pontos sinalizados pelas regras → marcar como NC com **assinatura (senha)**; reverter |
| ERRO SIST. (controle externo → bias) | Viés calculado a partir do controle externo (tabela `external_qc`) e usado no Sigma |
| ESP. QUAL. (CLIA / VB / fabricante → Sigma) | **Especificação da qualidade**: ETP, ES, EA, ET, **Six Sigma**, limites de CV (CLIA e VB) por analito/nível |

Os **28 analitos** do painel hematológico (Sysmex XN-1000), as **médias/DP reais** dos 3 níveis e as
**especificações** (TEa CLIA, CVi/CVg da variação biológica EFLM, CV do fabricante) já vêm carregados,
extraídos da sua planilha. Há ainda **dados de demonstração** (≈32 corridas por nível) com violações
propositais para você ver o motor de Westgard e o painel reagindo (WBC com 1-3s, RBC com 3-1s, PLT com 9x,
HGB com R4s entre níveis).

---

## Como rodar no seu computador (5 minutos)

1. Instale o **Python 3.10+** (https://www.python.org/downloads/ — marque "Add Python to PATH").
2. Abra o terminal na pasta do projeto e instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Inicie o sistema:
   ```bash
   streamlit run app.py
   ```
4. O navegador abre em `http://localhost:8501`. Entre com **INILAB / TESTE05**.

O banco `qclab.db` (SQLite) é criado e populado automaticamente na primeira execução.
Para zerar e voltar aos dados de demonstração: menu **Configurações → Recriar banco**.

---

## Como disponibilizar um link para um testador (grátis)

Para alguém testar **sem instalar nada**, publique no **Streamlit Community Cloud** (gratuito):

1. Crie uma conta no GitHub e suba esta pasta como um repositório.
2. Acesse https://share.streamlit.io → **New app** → aponte para o repositório e o arquivo `app.py`.
3. Em poucos minutos você recebe um link público. Passe o link + **INILAB / TESTE05** para o testador.

> ⚠️ **Importante:** no plano gratuito o armazenamento é **efêmero** — quando o app "dorme" e reinicia,
> o `qclab.db` volta ao estado de demonstração. Isso é ótimo para **demonstrar**, mas para uso real
> (dados que persistem, vários laboratórios) o próximo passo é trocar o SQLite por um banco gerenciado
> (ex.: **PostgreSQL** no Neon/Supabase/Render). A camada de dados está isolada em `database.py`,
> então a migração é direta.

---

## Estrutura dos arquivos

```
qclab/
├── app.py            # interface (Streamlit): login, painel, páginas
├── qc_engine.py      # regras de Westgard, estatística, Sigma/Erro Total (Python puro, testável)
├── database.py       # SQLite: schema (SQL), seed e consultas
├── seed_data.py      # dados reais do laboratório (médias/DP, specs) + gerador de demonstração
├── requirements.txt
└── .streamlit/config.toml   # tema escuro (verde-petróleo, como seu painel)
```

---

## O que ainda falta para virar produto (e não só protótipo)

Isto é um **protótipo funcional** — excelente para validar e demonstrar a laboratórios, mas ainda **não**
é um SaaS pronto para vender. Para isso, os próximos blocos (em ordem de impacto):

1. **Interfaceamento com o analisador / LIS** (HL7 ou middleware como o do próprio Sysmex). É aqui que
   está o verdadeiro diferencial competitivo — e o maior trabalho. Hoje os resultados entram manualmente
   ou por importação.
2. **Banco persistente + multiempresa** (PostgreSQL, cada laboratório isolado).
3. **Segurança/conformidade**: hashing mais forte (bcrypt/argon2) + política de senha, trilha de auditoria
   completa (quem fez o quê e quando), e adequação à **LGPD**.
4. **Refinos do CQ**: R4s e regras combinadas configuráveis por analito, múltiplos lotes em paralelo,
   relatórios de calibração e exportação assinada (PDF).

A matemática de CQ (Westgard, Sigma, Erro Total) — que é o miolo técnico — já está reproduzida e testada.
