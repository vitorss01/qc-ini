# Triagem dos 64 achados da FASE3A (QUALITY_GATE 6.4) — 03/10/2026

Fonte dos achados: `FASE3A_achados_estaticos.md` (análise estática de 04/08/2026, antes do hardening e do ADR-057).
Cada achado foi conferido **contra o que está hoje dentro dos `.xlsm` entregues** — VBA exportado dos arquivos,
fórmulas e estado das abas lidos por COM —, não contra a documentação. Onde a prova é um teste, ele é citado.

**Legenda:** ✅ resolvido · ⛔ obsoleto (o código ou a aba deixou de existir) · ◐ parcial · ⏳ aberto · ⚖ decisão registrada

## Resumo

| | Crítica | Alta | Média | Baixa | Total |
|---|---|---|---|---|---|
| ✅ resolvido | 7 | 12 | 12 | 3 | **34** |
| ⛔ obsoleto | 3 | 5 | 4 | 4 | **16** |
| ◐ parcial | 2 | 2 | 2 | 0 | **6** |
| ⏳ aberto | 0 | 2 | 2 | 2 | **6** |
| ⚖ decisão | 0 | 0 | 1 | 1 | **2** |
| **Total** | 12 | 21 | 21 | 10 | **64** |

Os **6 abertos** são de experiência de uso ou de defesa adicional; nenhum deixa resultado errado passar como certo:
#24 (base de conhecimento de Westgard inalcançável), #27 (título do gráfico não diz o analito), #41 (histórico de
eventos sem filtro), #46 (senha da tela de login em célula), #60 (tooltip chama de RUN a posição do ponto), #61 (Sair só no Início).
Dos **6 parciais**, quatro dependem de decisão ou de passo manual — senha do projeto VBA (#2, gate 3.4), sal nos
hashes (#13), assinatura com vínculo ao conteúdo (#14), histórico de Westgard append-only (#5) — e dois são limites
de exibição: janela de 180 corridas e Registros com 200 linhas (#37), alerta 12s sem contador no Painel (#44).

## Crítica (12)

| # | Achado | | Situação hoje e prova |
|---|---|---|---|
| 1 | Sem trilha de auditoria | ✅ | Audit_Log append-only com cadeia de hash (gate 3.1/3.2). Inativação, reativação, manual, plotagem (ADR-057, T21); acesso e cadastro (ADR-059, S03–S17); verificação `OK|n` (S17) |
| 2 | Nenhuma proteção persistida no arquivo | ◐ | Proteção gravada em toda aba, só Login visível, estrutura travada — **tinha regredido** com os instaladores do ADR-057 e voltou com `blindar_entrega.py` (S01, lido do XML). Falta a senha do projeto VBA (gate 3.4, passo manual) |
| 3 | Reenvio de Data+lote apaga reprovado e ressuscita excluído | ⛔ | Upsert e NovoRUN saíram (ADR-057). O recebido é imutável (DB_RECEBIMENTO; divergência vira A03), o RUN identifica a corrida (bloco por GAP_CORRIDA_MIN, T06) e reenvio não desfaz inativação |
| 4 | Cfg_Status redefine o histórico com uma célula | ⛔ | Cfg_Status removida (ADR-057). Participação é decidida por resultado na DB_CQ_FINAL, com justificativa obrigatória (E01) |
| 5 | Histórico de Westgard derivado e reescrevível | ◐ | `RegistrarEventosWestgard` ainda limpa e recalcula a aba do lote. Mitigação: toda alteração de média/DP vai ao Audit_Log (`PARAMETRO_LOTE_ALTERADO`), então o passado é reconstituível. **Decisão pendente:** gravar eventos append-only com os parâmetros congelados |
| 6 | Escalação de privilégio e sequestro de identidade | ✅ | ADR-059. Ainda aberto em 03/10 (na Bioquímica, N1/N2/hash destravados). Papel vem da tabela; ninguém se promove; só ADM altera terceiros (S05–S10, S13, S14) |
| 7 | Painel rotula regras que o motor não calcula | ✅ | Rótulos 1-3S · 2-2S · R4S · 4-1S · 8X = os cinco campos que o motor grava (ADR-045; conferido na aba) |
| 8 | Exclusão lógica sem justificativa, assinatura, rastro, valor | ✅ | Aba Inativar (ADR-057): justificativa obrigatória (E01), autor e data carimbados, trilha (`RESULTADO_INATIVADO`), valor preservado na Principal - Resultados e X vermelho no LJ (T07–T14) |
| 9 | Coluna Status sem validação | ⛔ | Não existe status digitado: STATUS_ANALITICO é calculado pelo Power Query |
| 10 | (diagnóstico) RegistrarEventosWestgard não tem laço infinito | ✅ | Confirmado; o congelamento era o recálculo (#11). 0,098 s (gate 1.2) |
| 11 | Recálculo nunca desligado + carga de fórmulas | ✅ | `mApp.InicioUsuario` (ADR-050), gate 1.2, 4.5 (estresse), 4.7 (estado restaurado) |
| 12 | AtualizarCalc apaga 12.600 fórmulas | ✅ | Motor escreve só em `Eng_Saida` (gate 2.1, 2.1b) |

## Alta (21)

| # | Achado | | Situação hoje e prova |
|---|---|---|---|
| 13 | Credenciais fracas e expostas | ◐ | Coluna do hash oculta e travada (ADR-059, S02). Continuam: SHA-256 **sem sal** (salgar exige recadastrar senhas) e senha de proteção legível no VBA enquanto o projeto não tiver senha (3.4) |
| 14 | Assinatura sem vínculo com o conteúdo; corridas sem assinatura | ◐ | Assinatura e recusa agora na trilha, com signatário e célula (ADR-059). Sem vínculo criptográfico com o conteúdo. Corridas sem assinatura são processo do laboratório |
| 15 | Modo Desenvolvedor sem marca | ✅ | `MODO_DESENVOLVEDOR_ATIVADO` / `_RECUSADO` (ADR-059, S15) |
| 16 | Três implementações de elegibilidade | ✅ | Uma só: `PARTICIPA_ESTATISTICA` da DB_CQ_FINAL (ADR-057, T18, T22 fonte única) |
| 17 | Status sem validação de dados | ⛔ | Idem #9 |
| 18 | Bias/ET/Sigma dos níveis 2–3 usam o bias do nível 1 (EQC) | ✅ | `AtualizarPainelEng` calcula bias por nível contra o alvo do próprio nível (`AlvoAnalito(analito, t+1, …)`) |
| 19 | DP alvo ausente: Westgard anulado e Painel diz "OK" | ✅ | Corrida sem parâmetro sai "SEM PARAMETROS" no Calc; o status do Painel mostra **"SEM MÉDIA/DP"** com o mesmo critério do motor (média e DP numéricos, DP > 0; `Calc!AX1`) — ADR-049 |
| 20 | Troca de lote herda média/DP | ✅ | `mLotes`: "LOTE SEM PARAMETRO NAO HERDA" |
| 21 | Aba Resultados salva vazia | ⛔ | Aba removida (ADR-057) |
| 22 | Validações de DB_Resultados no schema antigo | ⛔ | DB_Resultados removida (ADR-057) |
| 23 | Sem aviso de fora de faixa na digitação (frmCorrida) | ⛔ | frmCorrida removido; o dado chega pelo interfaceamento |
| 24 | Base de conhecimento de Westgard é código morto | ⏳ | A **classificação** chega ao Painel e aos eventos; `DetalheViolacao` (interpretação, causas, sugestões) **continua sem chamador** — nenhum botão ou célula chega nela |
| 25 | frmExcluir não filtra por nível | ⛔ | frmExcluir removido (ADR-057) |
| 26 | Nenhuma trilha em ponto de escrita | ✅ | Idem #1 |
| 27 | Painel sem congelamento; título do gráfico não diz o analito | ⏳ | Títulos fixos ("Nível 1 — Bioquímica"); o analito só aparece no seletor acima. Congelamento de painéis não verificado |
| 28 | Sem macros: abre sem login e com o banco à vista | ✅ | Arquivo entregue blindado (S01, XML); regressão do ADR-057 corrigida |
| 29 | Dois motores escrevendo Painel/Estatística | ✅ | Motor único → `Eng_Saida` (ADR-019; gate 2.1b) |
| 30 | priR/priRun não reiniciam por grupo | ✅ | Gate 1.8: 0 grupos com estado vazado |
| 31 | ScreenUpdating não restaurado em erro | ✅ | Gate 4.7; `mApp.Reset` (ADR-050) |
| 32 | RegistrarEventosWestgard roda 2× | ✅ | Cache por lote/equipamento: a segunda chamada sai na hora (ADR-050) |
| 33 | Nível fora de faixa em AlvoAnalito | ✅ | `ParametrosLote` recusa nível fora de 1..3 |

## Média (21)

| # | Achado | | Situação hoje e prova |
|---|---|---|---|
| 34 | Motor sobrescreve fórmulas; rótulos trocados | ✅ | Gate 2.1 e #7 |
| 35 | Sem carimbo imutável de digitação | ✅ | `REGISTRADO_EM` (manual) e `RECEBIDO_EM` (interface) carimbados pelo sistema; data futura recusada (gate 10.3, X01) |
| 36 | Registro de NC incompleto e destrutível | ✅ | NC = Inativar + COMENTARIOS_TECNICOS (ADR-057): justificativa obrigatória, reativação registrada, comentário preservado (A05) |
| 37 | Truncamentos silenciosos | ◐ | Buffer de eventos dinâmico (4.5); DB_Resultados e VIEW_ROWS saíram. Ficam: janela do Calc (NK = 180 corridas, por desenho) e Registros (200 linhas, hoje só calibração) |
| 38 | DP da Estatística por fórmula instável | ✅ | Estatística vem do motor (gate 2.1b: 0 fórmulas de parâmetro sobre o banco) |
| 39 | Lançar corrida custa ~20 cliques | ⛔ | frmCorrida removido; interfaceamento automático + Digitar Resultados |
| 40 | Trocar analito exige 27 cliques de spinner | ✅ | `selAnalito` com lista suspensa (`=lstAnalitos`); ▲▼ continua |
| 41 | "Ver Histórico" sem filtro nem ordenação | ⏳ | Eventos_Westgard: 6.179 linhas na Bioquímica, sem AutoFiltro |
| 42 | frmMassa: barra de progresso | ⛔ | frmMassa removido |
| 43 | frmMassa: contrato ilegível | ⛔ | Idem |
| 44 | ALERTA 12s nunca chega ao usuário | ◐ | Calculado por corrida (`Eng_Saida`, campo 6); o resumo do Painel conta só as 5 regras de rejeição |
| 45 | "REJEITADO" sem RUN/regra e sem expirar | ✅ | Painel: "Últ. violação: 1_3s · RUN 26050401", classificação e contagem (ADR-045); o veredito vale para o período filtrado |
| 46 | Senha de login em texto branco na célula | ⏳ | `DoLogin` apaga ao clicar Entrar; se ninguém clicar, fica na célula até salvar. Correção indicada: limpar também em `LockApp` |
| 47 | Registros pede "Corrida (Seq)" | ✅ | Registros só calibração, marcada pelo dia (ADR-057); "Seq" → "RUN" (gate 6.2) |
| 48 | Nenhuma aba com AutoFiltro | ✅ | Abas de dados são tabelas do Excel (DB, Principal - Resultados, QA, Inativar, Digitar Resultados, Audit_Log com filtro). Exceção: #41 |
| 49 | Início ensina caminho que não existe | ✅ | Texto atual: ATUALIZAR DADOS → Inativar → Painel. Detalhe: chama de "Resultados manuais" a aba "Digitar Resultados" |
| 50 | Contador `nv` dentro da guarda `nEv < 5000` | ✅ | Buffer dinâmico, sem teto (gate 4.5) |
| 51 | Collection indexada por posição (quadrático) | ⚖ | `For j … col(j)` continua, mas a coleção é por analito/lote e o motor leva 2,4–2,5 s com 32–44 mil resultados (03/10/2026). Sem efeito medido — não mexido |
| 52 | Coluna J da Estatística muda de bias EQC para interno | ✅ | Cadeia do bias EQC mantida por decisão (gate 2.1b) |
| 53 | EnableEvents nunca desligado | ✅ | `mApp.InicioUsuario` desliga e devolve (ADR-050) |
| 54 | AvaliarWestgard1N sem guarda | ⛔ | Rotina removida (ADR-045) |

## Baixa (10)

| # | Achado | | Situação hoje e prova |
|---|---|---|---|
| 55 | Área "DADOS INTERFACEADOS" dentro do banco | ⛔ | DB_Resultados removida; o interfaceamento entra pela camada DB (ADR-057) |
| 56 | Eixo X "Corrida (RUN)" / "Seq" | ✅ | Resolvido em 05/08/2026 (gate 6.2/6.3) |
| 57 | frmCorrida em ziguezague | ⛔ | Removido |
| 58 | Mensagens inconsistentes entre formulários | ⛔ | Formulários removidos |
| 59 | Legenda com 14 séries e "Repetição" 3× | ✅ | Série única "Não conforme (X)", X vermelho (ADR-057, T11) |
| 60 | Tooltip sem z e sem regra | ⏳ | Continua; e mais: mostra `"RUN " & <posição do ponto>` (ex.: "RUN 37"), não o RUN do eixo (26050401). Correção indicada: ler o RUN da coluna B do Calc. Não automatizável (evento de mouse) |
| 61 | Sair só no Início | ⏳ | Continua só no Início (a aba Resultados saiu) |
| 62 | `grupos(chave).Add` encadeado | ⚖ | Idioma válido em VBA (item do dicionário é Collection); coberto pelas provas 1.5/1.8 |
| 63 | Nomes fixos em 15.003 linhas | ⛔ | Nomes `r*` removidos (ADR-057) |
| 64 | `Dim stR` sombreia `Str()` | ✅ | Renomeado (`dstR`) |

## Achados de dados do laboratório (não são de software), vistos nesta triagem

- **Bioquímica, aba Início:** a identificação diz "Equipamento XN-1000 / SYSMEX" e "Controle XN-CHECK", que são da Hematologia.
- **Bioquímica:** lote em uso 8974 aparece **vencido** desde 30/06/2026; o lote 8977 chega pelo interfaceamento sem cadastro (QA I03).
- **Hematologia:** 8 lotes recebidos sem cadastro, de 31/10/2025 até o atual 625311 (~35.700 resultados fora do Painel) — QUALITY_GATE seção 10.

## Adendo — achados novos da 2ª rodada (03/10/2026), todos corrigidos e provados

| Achado | Onde | Correção / prova |
|---|---|---|
| Classificação de Westgard **nunca casava** ("Não classificada" para toda regra) | mWestgardKnowledge × códigos do motor (`1_3s`, `2_2s`, `2of3_2s`, `R_4s`, `3_1s`, `4_1s`, `Nx`) | `Familia()` traduz para a regra-mãe (ADR-060) |
| Bioquímica: validade do lote em uso só nos 25 primeiros lotes | Configuração!C21 | 100 lotes (ADR-060, A11) |
| Bioquímica: aba Configuração inteira destravada | travas | Editáveis só C5:C16, C20, D26:D125 (A12) |
| Bioquímica: identificação do equipamento/controle da Hematologia | Configuração!C9:C11, Início!C7:C8 | DIMENSION EXL-200 / SIEMENS; série e controle "A INFORMAR"; Início lê da Configuração |
| Lote gravado como número no Audit_Log | mAuditoria.Auditar | Coluna LOTE como texto (cadeia de hash inalterada) |
| `UnprotectAll` e `UnlockApp` executáveis por qualquer um pelo Alt+F8 | mSeguranca | Só sessão ADM / sem sessão volta ao login (ADR-059 revisto, S11–S12) |
| Spinner ia até 40 com 31/28 analitos ("0.0" no nome) | Painel | Limitado ao cadastro (ADR-062, P04) |
| **N3 (Hematologia) e N2 (Bioquímica) fora da tela**; cabeçalho A:U cortado | Painel (ADR-054: só a largura acompanhava a janela) | Encaixe ao entrar/redimensionar (ADR-063, G01/G05) |
| Legenda do gráfico sob a barra de rolagem | `AjustarGraficos` (`UsableWidth` inclui os cabeçalhos) | Cabeçalhos descontados (ADR-063) |
| Tique de OnTime reabre a pasta fechada **por automação** | vigia de zoom (ADR-061) | Laço ocupado no `BeforeClose` + sinal de fechamento (ADR-063, G08; a prova antiga era falso positivo) |
| "Quer salvar?" ao fechar sem nada do usuário mudado | `Workbook_BeforeClose` (`SystemLook False`) | Devolve o estado "salvo" (ADR-063) |
| Ctrl+Shift+G é atalho nativo do Excel 365 | proposta do Gemini | Atalho removido junto com o foco por nível (ADR-063) |

## Adendo — ciclo autônomo de 04/10/2026 (ADR-064 e auditoria final)

| Achado | Onde | Correção / prova |
|---|---|---|
| Hematologia: bias com sinal, nº de amostras/rodadas e status SDI/limites do CEQ **sempre "SEM EP"** | Estatística R, S, T, AC, AD (liam `Analitos!AR`, que só existe na Bio) | Provedor do filtro, como a coluna G (ADR-064, I07) |
| Bioquímica: rótulos acentuados corrompidos na padronização do CEQ | `mEQA.PadronizarStatus` | Texto correto, versionado em `src/comum` (I08) |
| Texto corrompido visível ("nÃ£o aplicÃ¡vel") | `mPlanoQC` (os dois) | Corrigido |
| Datas como "01/01/yyyy" | Hematologia, Estatística AF3:AF4 | Formato nos códigos do Excel instalado |
| Títulos dos gráficos sobre as linhas / sobre o eixo | Painel (IncludeInLayout = False) | Incluídos no layout |
| Frequência de CQ cortada | Painel, Plano de CQ | Alinhada à esquerda, com quebra |
| Bias EQC% = média de \|bias\| superestima o viés | ET% e Sigma (ADR-035) | **Pendência D-01** (decisão do RT), impacto medido no ADR-064 |
| **#2 (P1) completado**: Ctrl+S numa sessão gravava o arquivo aberto (abas visíveis; ADM/Modo Desenvolvedor desprotegido) | `Workbook_BeforeSave`/`AfterSave` | ADR-065: grava trancado e sem a sessão (S19, S20) |
| Sessão ADM editava o Audit_Log sem rastro | `mSeguranca.UnprotectAll` | Trilha sempre protegida (S18) |
| Auditoria final do cálculo/CEQ: 11 achados | ADR-064, "Auditoria final" | Corrigidos e provados (I00–I13) |

