---
name: linhas-alvo-adr078
description: ADR-078 (15/10/2026) — linhas de limite do LJ são séries de 2 pontos em Calc 184:185; troca de lote nunca grava na LotesStore
metadata:
  node_type: memory
  type: project
  originSessionId: cc5cf662-a93c-4a60-a8b5-23064798503c
  modified: 2026-10-15T22:12:33.369Z
---

ADR-078 resolveu o "gráfico vazio" ao trocar de lote (8974 -> 8973). A Média/DP sempre chegou ao Calc. O problema
era que as 7 séries de limite usavam como X as corridas (Calc!E3:E182, NA sem corrida).

Hoje é assim:
- **Séries de limite** (-3s..+3s) leem as linhas 184:185 do Calc. O X vem de xrunmin/xrunmax, que vale 0..21 sem
  corrida. O Y é M + k·S. As linhas 3..182 de Q:W/AM:AS/BI:BO não são mais plotadas.
- **Status** em J7..: DESATUALIZADO > SEM MÉDIA/DP > AGUARDANDO 1ª CORRIDA > SEM CORRIDA DESTE ANALITO >
  SEM CORRIDA NO PERÍODO > SEM PONTO > veredito. O sufixo " (alvo TESTE)" vem de LotesStore X:Z.
- **Contagem do lote na base:** Eng_Saida!W1 (`engNLoteBase`), escrita pela VBA `mEstatistica.NResultadosLote`.
- **Aviso:** Eng_Saida!Y1 (`engAviso`). O filtro da rotina fica em AA1 (`engFiltroRotina`).
- **`SalvarParametrosView` foi removida.** Trocar de lote só lê a store. A gravação é célula a célula, por
  `GravarParametroStore`.
- **Instalador:** `instalar_adr078.py` roda no `entregar.py` depois do `blocos_lote`. O `instalar_adr074.py` chama
  `mLotesBlocos.ResgatarDigitados` antes de `MontarBlocos`.

Commit de6b36a (16/10/2026, main). Entrega 20 lançada em %TEMP%\qcw\entrega20.log.

**Why:** o usuário ameaçou reduzir o projeto a um lote por arquivo, o que não resolveria nada. Ele pediu uma análise
profunda e uma implantação orquestrada.

**How to apply:**
- Não voltar a plotar limites sobre a grade de corridas.
- Não gravar na LotesStore em troca de lote.
- Na Hematologia, `selEquipamento` = Cfg_Integracao!B11. Nunca escrever nele.
- Teste: `testes/qa_linhas_alvo.py` (A01 = o caso do usuário).

**Armadilha da ferramenta:** em heredoc no Bash, `\\n` vira uma quebra de linha real. Patches com barra invertida
vão em arquivo .py escrito pelo Write. Em Windows, `open(p,'w')` em modo texto troca LF por CRLF; use `newline=''`.

Ver [[abas-especificacoes-blocos]] e [[qc-ini-contexto]].
