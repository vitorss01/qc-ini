# Parecer UX/UI: ciclo 2026-11-15_reteste05 (reteste)

**Avaliador:** UX/UI (ADR-081) · **Data:** 15/11/2026 · **Entrada:** 1 ocorrência de usabilidade da frente F1
(`Bioquimica_F1.md`, R5-01) · **Produção avaliada:** Bio b20fe00f7e51 · Não abri Excel nem editei código. Usei a
evidência do relatório (`Bioquimica_F1.json`, cenário S8_so_De, NOVA e BASE) e a leitura das fontes e do instalador.

## R5-01: Estatística J5 mostra "00/01/1900" quando o Painel tem só o De

**Veredito:** PROCEDE · **Prioridade final:** Baixo (mantida) · **Precisa de autorização:** não (só formato).

### Avaliação do ponto de vista do usuário

- É um defeito, não uma preferência. O cabeçalho "PERÍODO ANALISADO" existe para o usuário conferir a janela que
  o sistema aplicou. Hoje ele mostra "Janela 01/09/2026 a 00/01/1900", uma data impossível e anterior ao início. Quem
  confere ou imprime a Estatística para o registro da qualidade fica sem saber até quando os dados foram usados, e a
  data de 1900 tira a confiança no documento.
- O cenário é legítimo, e o Painel não deve proibi-lo. "Do dia X em diante" é um uso real: por exemplo, depois de
  uma manutenção ou da troca de reagente. Por isso não recomendo avisar ou bloquear no Painel. O certo é a
  Estatística mostrar o que a janela é.
- O que a janela é de fato: no motor do Painel, `fAte = 0` não aplica limite superior (`mEstatistica.PassaFiltro`,
  linha 1832, `If fAte > 0 Then ...`), e a janela das corridas vai até a última (`fimJan = nTodas` quando o
  `filtroAte` não é data, linha 1462). Logo, "só o De" quer dizer **do De até o último dado**. Isso responde à dúvida
  1 do relatório F1 quanto ao Painel. A confirmação final cabe ao Tester (ver aceite).
- Mantenho a prioridade Baixo: os números não mudam, e o caso só aparece quando alguém apaga uma das datas. É
  acabamento de apresentação (REGRAS_COMUNS §4).

### Correção proposta (só formato; fórmula, Value2 e JanelaFim intactos)

Dono: `instalar_adr092.py`, `estatistica_bio`, no mesmo laço que já formata B5/D5 pelo `cd.fmt_data` (nunca
`NumberFormat = 'yyyy'`, por causa da armadilha do ano literal, ADR-095 §5 e `qa_adr095` L04).

| Célula | Formato novo (forma passada ao `cd.fmt_data`) | Mostra com valor 0 | Caso |
|---|---|---|---|
| J5 | `dd/mm/yyyy;;"último dado"` | `último dado` | só o De (Até vazio) |
| H5 | `dd/mm/yyyy;;"primeiro dado"` | `primeiro dado` | só o Até (De vazio), o caso simétrico |

- A linha fica assim: "Janela 01/09/2026 a último dado". Com só o Até: "primeiro dado a 30/09/2026".
- Com o Painel em "Tudo", nada muda: H5 continua com o texto "histórico inteiro do lote" (um formato de 3 seções
  mostra o texto como está), e I5/J5 continuam "".
- Nas datas normais (mês, trimestre, ano, De e Até), a exibição continua `dd/mm/aaaa`.
- O instalador deve conferir, antes de formatar, que as fórmulas de H5/J5 são as do ADR-092 (`H5`/`J5` do próprio
  instalador). Se não forem, ele para e não salva nada, como já faz com B5/D5.
- Antes de mudar, confira se `fmt_data` só troca o `y`: os textos entre aspas não têm `y`.
- **Recuo, se a largura não der:** se a auditoria mostrar `####` no J5 ou no H5 (um número formatado não transborda
  para a célula vizinha), use `dd/mm/yyyy;;` (zero em branco, igual ao B5/D5). Não alargue a coluna J sem rodar a
  auditoria visual da aba inteira.
- B5/D5 ficam como estão (`;;`). Eles espelham o que foi digitado no Painel, e o vazio ali está certo.
- Opcional, no GUIA_DE_USO (seção do Painel/período), uma linha: "Deixar só o De = do De até o último resultado;
  só o Até = do primeiro resultado até o Até."

### Critério de aceite (reteste)

1. **S8 (só o De, 01/09/2026):** J5.Text = "último dado", J5.Value2 = 0, fórmula do J5 igual à do ADR-092,
   H5.Text = "01/09/2026".
2. **S9 (só o Até, 30/09/2026), novo, não foi executado neste ciclo:** H5.Text = "primeiro dado", H5.Value2 = 0,
   J5.Text = "30/09/2026".
3. **Regressão:** S6 (Out/2026), S7 (atravessa o ano) e "Tudo". H5/J5 com o mesmo texto da entrega anterior;
   em "Tudo", H5 = "histórico inteiro do lote" e I5/J5 vazios.
4. **Busca "1900":** nenhum "1900" no texto das linhas 1–12 da Estatística em S6, S7, S8, S9 e "Tudo".
5. **`auditoria_abas.py` na Estatística:** zero achados `numero_####` e `texto_cortado` em H5:J5, em S8 e S9.
   Se houver `####`, aplique o recuo `;;`, e o aceite do item 1/2 passa a ser Text = "".
6. **`qa_adr095` L04:** nenhum formato com ano literal (o H5/J5 novos inclusive).
7. **Números:** n, média, DP, CV, bias, ET, Sigma e status da Estatística iguais aos da base em S8 e S9 (diferença 0).
8. **Tester:** confirme, em S8, que o Painel e a Estatística usam os resultados de 01/09/2026 até o último dado
   (os resultados de antes do De ficam fora, nenhum limite superior). Se não for assim, o texto "último dado" seria
   falso: use o recuo `;;` e leve a diferença ao Especialista em CQ como ocorrência funcional.

### Observação para o Orquestrador (fora do escopo de UX)

Na leitura do `AtualizarEstatisticaAba` (`mEstatistica.bas`, cerca da linha 2130), com B4 = PERSONALIZADO o motor
desse bloco filtra por **ano** (`anoDe`/`anoAte` a partir de B5/D5) e não por data. Com só o De, isso vira
2026 a 2026. O relatório F1 diz que a Estatística lê o `engPainel`, igual ao Painel. Peço que o Tester confirme
qual bloco alimenta a grade exibida. Se for esse, há diferença entre a janela mostrada no H5:J5 e a usada no
cálculo, e o caso vai para o Especialista em CQ. Não é parte da R5-01 e não muda a correção de formato acima.
