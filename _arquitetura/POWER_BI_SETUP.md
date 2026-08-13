# QC_INI — Montagem do relatório Power BI

## Antes de tudo: o que é automatizável e o que não é

Verificado nesta máquina:

| Item | Estado |
|---|---|
| Power BI Desktop | ✔ instalado — `2.156.951.0` (`C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe`) |
| `pbi-tools` | ✘ não instalado |
| Tabular Editor | ✘ não instalado |
| módulo PowerShell `MicrosoftPowerBIMgmt` | ✘ não instalado |
| Power BI REST API | ✘ exigiria tenant, registro de aplicativo e token |

**Conclusão honesta: não há caminho suportado para gerar o `.pbix` por script
aqui, e nenhum `.pbix` foi fabricado.** O `.pbix` é contêiner binário
proprietário; escrevê-lo à mão não seria reprodutível nem manutenível, e um
arquivo que abre hoje e corrompe no próximo update do Desktop é pior do que não
ter arquivo.

O que **foi** automatizado: a camada de dados inteira (`BI_Data` + `tblBI_Fato`),
com validação de chave e reconciliação contra o motor a cada build. É a parte
onde erro custa caro. A montagem visual abaixo leva ~40 minutos e é feita uma vez.

**Para automatizar depois:** instalar `pbi-tools` (`winget install pbi-tools`) ou
Tabular Editor 2 (gratuito) torna o modelo — tabelas, relações, medidas —
gerável por script a partir dos arquivos deste diretório. Só o layout visual
continuaria manual.

---

## Passo a passo

### 1. Parâmetro do caminho
Página Inicial → Transformar dados → Gerenciar Parâmetros → Novo
`pCaminhoQC` · Texto · valor = caminho completo do `QC_Bioquimica.xlsm`.

### 2. Consultas
Cole as cinco consultas de [`POWER_QUERY.md`](POWER_QUERY.md), nesta ordem:
`Fato_QC`, `Dim_Data`, `Dim_Analito`, `Dim_Lote`, `Dim_Nivel`. Fechar e Aplicar.

### 3. Modelo
Criar as quatro relações da tabela em `POWER_QUERY.md §6`. Todas
muitos‑para‑um, **direção simples**. Marcar `Dim_Data` como tabela de datas.
Ocultar as colunas de chave.

### 4. Medidas
Criar a tabela `_Medidas` e colar as medidas de [`DAX_MEASURES.md`](DAX_MEASURES.md).

### 5. Páginas

#### Página 1 — **Painel Operacional**
- Barra superior: segmentações `Analito`, `Nível`, `Lote`, `Competencia`
- Cartões: `n Resultados`, `n Corridas`, `% Aceitáveis`, `Viol Total`, `Status Global`
- **Gráfico Levey-Jennings** (ver §6 abaixo)
- Tabela de detalhe: `Data`, `RUN`, `Resultado`, `Z`, `Veredito` — com formatação
  condicional em `Veredito`

#### Página 2 — **Análise Estatística**
- Segmentações: `Analito`, `Nível`, `Lote`, período
- Matriz por analito: `n Resultados`, `Media Observada`, `DP Observado`, `CV%`,
  `Bias%`, `Sigma`, `% Fora de Controle`, `Viol Total`
- Barras: `Sigma` por analito, ordenado crescente — os piores primeiro
- Linha: `Media Movel 7 Corridas` e `Deslocamento vs Alvo (SD)`

#### Página 3 — **Monitoramento Executivo**
- Cartão grande: `Status Global`
- `% Aceitáveis`, `Analitos Críticos`, `Rejeitados Hoje`
- Barras: violações por regra (`Viol 1_3s`, `Viol 2_2s`, `Viol R_4s`)
- Tabela: analitos com `Sigma < 4`, ordenada crescente

#### Página 4 — **Qualidade do Dia**
Feita para o celular: `Data Mais Recente`, `n Hoje`, `Rejeitados Hoje`,
`Status Global`, e a lista dos analitos rejeitados hoje.
Objetivo: o gestor abre e entende em menos de 30 segundos.

### 6. Levey-Jennings no Power BI

Não existe visual nativo de L‑J. A montagem que funciona:

1. Visual **Gráfico de linhas e colunas agrupadas** — ou linhas com múltiplas séries.
2. Eixo X: `RUN` (não a data — corridas do mesmo dia se sobreporiam).
3. Valores: `Resultado`.
4. Adicionar como séries de linha constante: `LJ Média`, `LJ +1s`, `LJ +2s`,
   `LJ +3s`, `LJ -1s`, `LJ -2s`, `LJ -3s`.
5. Cores: média sólida; ±1s discreto; ±2s âmbar; ±3s vermelho.
6. Cor do ponto por `Veredito` (formatação condicional): `OK` neutro,
   `REJEITADO` vermelho.
7. Colocar o cartão `LJ Aviso` sobreposto — ele aparece quando o filtro não está
   restrito a um analito/nível/lote, situação em que o gráfico não teria sentido.

### 7. Navegação por botões
Inserir → Botões → Navegador → **Navegador de páginas**. Ele se atualiza sozinho
quando páginas são criadas ou renomeadas — ao contrário de bookmarks manuais,
que quebram calados quando alguém renomeia uma página.

### 8. Layout mobile
Exibir → Layout do celular, por página. Manter apenas: `Status Global`,
`% Aceitáveis`, `Rejeitados Hoje`, o Levey‑Jennings e a lista de críticos.
**Não** reduzir o layout de desktop — reorganizar.

### 9. Identidade visual
Exibir → Temas → Personalizar. Reaproveitar as cores do Painel do Excel para que
o usuário reconheça o mesmo sistema.

---

## Verificação obrigatória antes de publicar

Rodar o [`BI_TEST_PLAN.md`](BI_TEST_PLAN.md), em especial a reconciliação
Excel × Power BI. Publicar um painel que discorda do Excel cria duas verdades
sobre o mesmo controle — e o laboratório não teria como saber qual seguir.
