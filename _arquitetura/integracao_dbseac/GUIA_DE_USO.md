# QC_INI — Guia de uso (implantação, outubro/2026)

QC_Hematologia.xlsm e QC_Bioquimica.xlsm · origem dos resultados: DB_SEAC (interfaceamento SIPEC), que a tarefa
agendada atualiza a cada 2 minutos.

## 1. Rotina diária

1. Abra o QC e faça o login.
2. Clique em **⟳ ATUALIZAR DADOS**. Leva cerca de 2,5 min porque lê o DB_SEAC pela rede.
3. No **Painel**, escolha o analito pelas setas ou pela lista. Os gráficos de Levey-Jennings mostram os níveis lado a
   lado.
4. Passe o mouse num ponto do gráfico. Aparecem o **ID**, a corrida (RUN), a data/hora e o valor. Nos pontos já
   inativados (X vermelho) aparece também **INATIVADO**.
5. Confira o Status da corrida, as regras de Westgard violadas e o histórico do analito.
6. Olhe a aba **QA_INTEGRACAO**:
   - **ERRO**: precisa de ação;
   - **ALERTA**: o sistema protegeu o dado, mas alguém deve olhar;
   - **INFO**: só registro.

## 2. Inativar um resultado (não conforme / repetição)

1. No gráfico, passe o mouse no ponto e anote o **ID**. Basta o número, ex. 314216.
2. Na aba **Inativar**, na próxima linha livre, preencha:
   - **ID**;
   - **Analito** (lista);
   - **Registrar-LJ**: marcado = o ponto aparece como X vermelho; desmarcado = não aparece no gráfico;
   - **Motivo**, que é obrigatório e vale como justificativa.
3. Clique em ⟳ **ATUALIZAR DADOS**.
- O resultado não é apagado: continua na base, com PARTICIPA = NÃO, e a trilha de auditoria registra data e usuário.
- Se o ID não for do analito informado, o QA acusa e a inativação não é aplicada.
- **Para reativar**, apague a linha e atualize.

## 3. Digitar um resultado (falha do interfaceamento)

1. Na aba **Digitar Resultados**, preencha data, hora, lote, nível, analito, resultado e motivo. O ID (MAN_0001…) é
   gerado sozinho.
2. Atualize e confira o STATUS:
   - **ATIVO** = entrou na estatística;
   - **CONFLITO_MANUAL** = o mesmo resultado já veio pelo interfaceamento. O manual fica guardado, fora do cálculo.
     Manual **nunca** substitui automático.

## 4. Especificação do ETp (aba Analitos)

- **Bioquímica:** a coluna S escolhe a especificação de cada analito:
  - **CLIA** usa a coluna K;
  - **VB** (variação biológica: CVI e CVG nas colunas O e P, desempenho na Q) usa a coluna R;
  - **FAB** usa a coluna M.
  
  A coluna T mostra o ETp em uso. "-" = especificação não cadastrada, e nesse caso o Sigma fica "-".
- **Hematologia:** a fonte segue a prioridade ETp configurado (P) → CLIA → VB, e a coluna Q mostra qual valeu.

## 5. Controle externo (CEQ) e Sigma (aba Estatística)

**Bias do CEQ.** O programa é escolhido em R4 na Hematologia e por analito em Analitos!AR na Bioquímica. CAP e
Controllab nunca se misturam. As opções:

| Ano | Rodada | Bias |
|---|---|---|
| qualquer | uma rodada da lista | só aquela rodada |
| um ano (ex. 2025) | TODAS | as rodadas daquele ano juntas (sem dado no ano = SEM EP) |
| TODOS | TODAS | as rodadas de todos os anos juntas |

A célula K5 diz o que entrou: rodadas e linhas.

**CV do Sigma.**
- É o **CV pooled** do CIQ: lotes com n ≥ 20, combinados pelos graus de liberdade.
- A janela padrão é de **6 meses**, terminando no último resultado. Ela fica na linha 11 da Estatística e pode ser
  mudada para 7, 8, 9 meses ou mais.

**Sigma.** Sigma = (ETp − |Bias|) / CV pooled. O "Sigma do plano" no Painel usa esse valor e avisa quando algum nível
ficou sem Sigma.

**Incerteza de medição** (Nordtest, ADR-067).
- Usa só o CIQ (12 meses) e o CEQ, sem certificado de calibrador.
- Sem CEQ suficiente para o analito, o U não é estimado, e a coluna "Situação" diz o motivo.

## 6. O que o laboratório mantém

- **CEQ:** digite cada rodada nas abas EQA.CAP_Dados / EQA.Controllab_Dados, com o **DP do grupo e o nº de
  laboratórios** de cada amostra. Sem esses dois campos, a incerteza não sai.
- **Cadastro (aba Analitos):** ETp, CVI e CVG de analito novo.
- **Investigação do CV:** revisar os IDs de `estudos/cv_suspeitos_<produto>.csv` (relatório
  `estudos/CV_INVESTIGACAO_2026-10-13.md`) e inativar os não conformes com motivo.

## 7. Itens que ficam com o laboratório (senhas não ficam com o desenvolvimento)

- **Senha do projeto VBA:** no Excel, Alt+F11 → Ferramentas → Propriedades de VBAProject → Proteção → "Bloquear projeto
  para exibição" e a senha. Salve e reabra.
- **Senha de proteção das abas:** a senha atual está no código público do repositório desde as versões anteriores.
  Para trocá-la:
  1. altere a constante `SENHA` em `instalar_integracao.py` e nos testes;
  2. faça uma entrega (`entregar.py`);
  3. não publique a senha nova no GitHub. Use uma variável de ambiente; o desenvolvimento pode preparar isso se for
     pedido.

## 8. Se algo der errado

- **"DB_SEAC não encontrado" / ATUALIZAR DADOS falha:** veja o log diário em `AUTOMACAO_DB_SEAC\logs`. Desde
  13/10/2026, a automação restaura sozinha o último backup se o arquivo sumir; procure "RESTAURADO" no log.
- **Voltar uma versão do QC:** os backups de cada entrega ficam em `_backup_pre_integracao_<data>`, e os hashes em
  `MANIFESTO_PRODUCAO.md`.
- **Atualização sem rede (fora do laboratório):** MODO_FONTE = HISTORICO na aba Cfg_Integracao. O sistema reprocessa só
  o que já foi recebido. Volte para SEAC na rede.
