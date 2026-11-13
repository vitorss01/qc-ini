# Entrega 39 (ADR-092): acompanhamento

Este arquivo é atualizado a cada etapa e enviado ao GitHub (`_arquitetura/agentes/relatorios/STATUS_ENTREGA_39.md`), para você acompanhar pelo celular.

## Situação: entrega EM ANDAMENTO desde 13/11/2026
- **Código entregue:** o commit `055e7ff`, feito em casa e baixado com git pull no trabalho.
  - A conferência automática do VBA passou.
  - Os scripts novos carregam sem erro.
- **Produção de referência (11/11):** Bioquímica `8b1cda13…`, Hematologia `f1907656…`.
- **Etapas:**
  1. Instala os cerca de 29 instaladores, incluindo o `instalar_adr092`, numa cópia de cada produto.
  2. Roda cerca de 32 suítes por produto.
  3. Troca a produção só se tudo passar, com backup e conferência de sha.
- **Previsão:** de 3 a 4 h para os dois produtos.

## Depois da entrega
1. ATUALIZAR na produção (`aplicar_atualizar.py`).
2. Usuário Simulado (reteste): as 19 ocorrências do ADR-090, OC-0040/0043/0049/0056..0059 e o ADR-091.
3. Tester VALIDAR do que mudou.
4. Confirmação do CQ para a regra do N-01 (manual em lote de urina com analito de soro: recusar).
