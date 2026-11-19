---
name: adr086-087-telas
description: "ADR-086 (H3 vazio = todos os lotes) e ADR-087 (tblRegistros, filtros, frmCEQ, ajuste visual guiado pela auditoria); armadilhas do Excel pt-BR achadas em 29/10/2026"
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-30T00:03:14.497Z
---

ADR-086/087 (commit 301f41f) estão em PRODUÇÃO nos dois QCs. A Bio foi entregue em 30/10 às 20h37 (com o 085) e a Hema em 31/10 às 00h16, depois de o S6 do qa_sync (refresh do Power Query, motor reserva) passar do teto de 1500 s; o S6 ganhou teto próprio de 3600 s. Ver [[adr085-painel-tela-unica]].

O que não é óbvio pelo código:
- **Names.Add:** com literal booleano (`=X=TRUE`), o Excel pt-BR grava `='TRUE'` e o .xlsm salvo **não abre mais**. Use `=N(X)<>0`. O `instalar_adr086.nome()` confere a volta do Excel.
- **Palavra reservada como variável:** `Dim aS()` não compila (aS = As). O `lint_vba.py` agora pega isso, mas só checa os arquivos passados como argumento: sem argumento, ele diz OK sem checar nada.
- **Salvar com eventos ligados:** o Workbook_BeforeSave cancela o Save e engole o erro, então o log diz "salvo" e nada é gravado. O instalador mantém os eventos desligados e confere o mtime.
- **Aba oculta no arquivo blindado:** Select e FreezePanes falham. Mostre a aba só durante a operação e use SplitRow.
- **Ajuste visual:** o `visual_telas.py` é guiado pelo `testes/auditoria_abas.py`. O V02 do qa_adr087 é o gate da entrega para as telas principais.
- **Inspeção visual:** é feita pelos PNG (exportados para PDF e renderizados), nunca por captura de tela. Ela achou o que a medição não achava: gráfico sobre o bloco, eixo fixo cortando barra, limites ausentes por `#VALOR!`.
- **Fontes VBA em UTF-8:** os `src/*.bas` são UTF-8, e o lint e o codigo_atual leem assim. Editar como cp1252 e
  inserir acento grava 0xE9 solto, e o lint quebra. Em 30/10 isso aconteceu com "Média" e "Nº".
- **Revisão adversarial antes da entrega:** achou o que 40+ testes verdes não pegavam. Com filtro, a linha de
  digitação some; assinatura em linha oculta sai no tamanho original; o N02 do qa_inativar falharia na entrega.
- **Tempo do S6 (Power Query) na Hema:** varia de 170 a 1130 s, com o motor em ~20 s. Não é sinal de regressão; o que se confere é o resultado.
- **qa_sync:** usa uma cópia congelada do DB_SEAC. Na entrega 33, a produção recebeu dados no meio do teste e a Bio foi barrada.

**Why:** cada item custou uma rodada de testes; o do nome corrompeu o arquivo.
**How to apply:** valem para qualquer instalador novo que crie nomes, salve ou mexa em formato.
