# Relatório Técnico — Módulo de Controle Externo (EQC) + Integração com o Controle Interno (IQC)

## 1. Objetivo
Adicionar ao QC-INI um módulo completo de **Controle Externo da Qualidade** (ensaios de
proficiência: CAP, Controllab, PNCQ…) e integrá-lo ao **Controle Interno** já existente, gerando
indicadores robustos de desempenho analítico por ensaio: **CV%, Indicador de Bias Anual, Erro
Total e Métrica Sigma**.

## 2. O que foi implementado

### 2.1 Banco de dados (`qclab/database.py`)
- **Novas tabelas:** `eqc_rounds` (rodadas), `eqc_samples` (amostras de cada rodada),
  `eqc_audit` (trilha de auditoria).
- **Colunas novas** (via `_migrate`, sem perder dados de bancos existentes):
  `analytes.area`, `analytes.unit`, `specs.etp_pct`, `specs.etp_source`.
- **Regras no banco:** `UNIQUE(analyte, year, round_number)` + validação em `eqc_add_round`
  (máx. **6 rodadas/ano**, número de rodada único, faixa 1–6).
- **Serviços:** `eqc_add_round`, `eqc_update_round`, `eqc_delete_round`, `eqc_list_rounds`
  (com filtros), `eqc_get_round/samples`, `eqc_round_count`, `eqc_annual_indicator`,
  `eqc_years/providers/lotes`, `eqc_analytes_with_rounds`, `eqc_log` (auditoria).
- **Registro de ensaios (extensível):** `add_analyte`, `get_analytes(area, with_reference)`,
  `get_areas`, `get_analyte_info`, `set_etp/get_etp`.

### 2.2 Motor de cálculos (`qclab/qc_engine.py`) — em **percentual**
- `sample_bias_pct(lab, peer)` = `((lab−peer)/peer)×100` (None se peer 0/nulo → evita divisão por zero).
- `round_abs_bias(biases)` = média dos **|bias|** das amostras (ignora None/NaN).
- `annual_bias_indicator(biases_rodada)` = média dos |bias| das rodadas → **Indicador de Bias Anual**.
- `external_total_error(cv, bias)` = `cv×1,65 + bias`.
- `external_sigma(etp, bias, cv)` = `(etp − bias)/cv` (None se cv 0/nulo).
- `annual_performance(...)` consolida ET/Sigma e lista pendências ("dados insuficientes").
- **Padronização:** fator do erro aleatório unificado em **1,65** em todo o sistema (`EA_Z`).
- `compute_sigma` aceita **ETp configurável** (`etp_pct`/`etp_source`).

### 2.3 Interface (`qclab/app.py`)
- Nova página **"Controle Externo"** no menu, com 4 abas:
  1. **Cadastrar rodada** — área → ensaio → ano → nº rodada (com checagem 1–6 e bloqueio do
     limite), provedor, unidade, lote, status, PDF, observações e **amostras** (editor) com
     **bias por amostra e |bias| da rodada ao vivo**.
  2. **Consultar / gerenciar** — filtros (ano, área, ensaio, provedor, rodada, lote), tabela,
     detalhe de amostras, exclusão com confirmação e registro de auditoria.
  3. **Desempenho anual** — tabela integrada por ensaio (Bias Anual, CV% do IQC, ETp, ET, Sigma),
     melhor/pior desempenho, selos de dados insuficientes e gráfico de evolução do bias.
  4. **Ensaios & ETp** — cadastro de novos ensaios/áreas (extensibilidade) e configuração do ETp.
- **Integração no Painel e na página de Sigma:** o viés isolado antigo foi **substituído** pelo
  Indicador de Bias Anual quando há rodadas (com fallback ao valor isolado); ETp configurável;
  fator 1,65. Telas IQC passam a listar apenas analitos com referência.

### 2.4 Dados de demonstração (`qclab/seed_data.py`)
- Áreas e analitos extras (Bioquímica, Imunologia, Coagulação) — extensíveis.
- ETp por ensaio. Rodadas EQC demo: **WBC/2025** reproduz o exemplo (|bias| 0,4 / 1,8 / 0,6 →
  **Indicador 0,9333**) e Glicose (Bioquímica).

## 3. Validação numérica (exemplo da especificação)
| Cálculo | Resultado | Esperado |
|---|---|---|
| Indicador Anual (0,4; 1,8; 0,6) | **0,9333%** | 0,93% |
| ET (CV%=2,0; bias=0,93) | **4,23%** | 4,23% |
| Sigma (ETp=10; bias=0,93; CV%=2,0) | **4,535** | 4,5 |
| Painel demo WBC/2025 (CV%=1,424) | ET 3,283% · **Sigma 4,26** | conferido na UI |

## 4. Testes
- `test_eqc.py` — **45/45** (funções puras, limite de 6 rodadas, bloqueio da 7ª, duplicidade,
  filtros, auditoria, indicador anual, integração com CV%, divisão por zero, peer nulo).
- `test_westgard.py` — **30/30** (Controle Interno, sem regressão).
- Executar: `python test_eqc.py` e `python test_westgard.py`.

## 5. Arquivos modificados/criados
- `qclab/database.py` (schema, migração, serviços EQC, registro de ensaios/ETp, auditoria)
- `qclab/qc_engine.py` (funções EQC em %, fator 1,65, ETp configurável)
- `qclab/app.py` (página Controle Externo + integração no painel/Sigma + menu)
- `qclab/seed_data.py` (áreas/analitos, ETp, rodadas demo)
- `test_eqc.py` (novo)
- `RELATORIO_EQC.md` (este relatório)

## 6. Como executar
```
streamlit run qclab/app.py          # app (login INILAB / TESTE05)
python test_eqc.py                  # testes do Controle Externo
python test_westgard.py             # testes do Controle Interno
```
No app: menu **Controle Externo** → cadastrar rodadas, consultar, ver **Desempenho anual**.
Para um banco limpo com os dados demo: Configurações → "Recriar banco de demonstração".

## 7. Compatibilidade e evolução
- **Migração automática:** bancos existentes ganham as colunas/tabelas novas sem perder dados.
- **Multiárea e extensível:** novos ensaios/áreas podem ser cadastrados a qualquer momento.
- **Mobile/tablet/desktop:** a página reusa o CSS responsivo já validado (abas roláveis, colunas
  empilhadas, tabelas com rolagem horizontal).
- **Play Store (futuro):** arquitetura web + banco central não impede o empacotamento posterior
  (TWA/PWA) — não implementado nesta etapa, conforme combinado.
