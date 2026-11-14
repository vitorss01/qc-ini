# Parecer de UX: OC-0060 (tempo de inativar ou reativar pelo gráfico)

**Agente UX/UI, 14/11/2026, só leitura.** Registrado pelo Orquestrador.

## Veredito sobre a ocorrência: PROCEDE (Médio)
A lentidão é um problema real, não preferência. A prioridade sobe para Alto se a Hematologia passar de 30 s ou se inativar a corrida inteira virar rotina.

Na prática, depois de uma falha de calibração ou de equipamento, o técnico inativa vários pontos seguidos. Na Hematologia, com 3 níveis, uma corrida inteira leva de 75 a 85 s com a tela presa.

**Achado novo:** a estimativa do "aguarde" é fixa (`src/bio/mUI.bas:1015`, `src/hema/mUI.bas:1016`: `IIf(mEstatistica.NLV >= 3, 25, 15)`). A Bioquímica promete "cerca de 15 s" e leva de 19 a 20 s. Promessa não cumprida aumenta a espera que o usuário sente.

## 1. O teto provisório de 22/30 s serve como critério de UX?
**Sim, mas só como proteção contra regressão, com prazo e com condições.**

**Por que serve:** a ação é rara e deliberada. Durante a espera há janela de etapas, a mensagem "⏳ … aguarde" e a confirmação "✓ … entre as corridas k e k+1". Por isso, 20 a 27 s são toleráveis.

**Por que é só provisório:**
- passa muito dos cerca de 10 s em que a pessoa ainda fica atenta à tarefa;
- os 30 s da Hematologia são o limite máximo tolerável, e a produção atual já usa 84% dele;
- o teto absoluto sozinho não pega a piora de uma entrega.

**Condições:**
- vale até a entrega da otimização, no máximo até 30/11/2026;
- não pode subir de novo sem novo parecer de UX;
- exige o ajuste da estimativa do "aguarde" (item 3a).

## 2. Critério duplo (recomendado)
Os dois critérios valem separadamente para inativar e para reativar. Se qualquer um falhar, o U13 reprova.

- **(a) Critério relativo:**
  - medição: 3 vezes na produção e 3 na cópia, na mesma bateria e máquina, alternadas (P, C, P, C, P, C);
  - alerta (não reprova): mediana da cópia maior que **1,10 ×** a mediana da produção;
  - reprova: mediana da cópia maior que **1,20 ×** a mediana da produção.
- **(b) Teto absoluto (mediana de 3):**
  - provisório: Bio 22 s, Hema 30 s;
  - depois da otimização: Bio 12 s, Hema 15 s.

## 3. Meta da otimização e melhorias baratas
**Meta:** mediana até **8 s na Bio e 10 s na Hema** por clique. Para chegar lá, o cálculo tem de ser incremental, só do lote e do analito do ponto.

A otimização não precisa de autorização se o resultado for idêntico ao de um ATUALIZAR completo na mesma cópia: Eng_Saida, Calc, Westgard_Fotos, Incerteza e tblQA_Integracao. Adiar o QA, as fotos do Westgard ou a incerteza para depois precisa de autorização, e não é recomendado.

Melhorias baratas:
- **a) Estimativa honesta no "aguarde" (obrigatória).**
  - Usar a mediana das últimas 5 durações "SO_ENTRADAS" do Sync_Log, arredondada para cima até o múltiplo de 5 s. Sem histórico, usar 20 s na Bio e 30 s na Hema.
  - Aceite: o tempo prometido é maior ou igual ao medido em 3 de 3 execuções, e no máximo 10 s acima dele.
- **b) Mostrar o gráfico assim que estiver pronto (opcional).**
  - A janela de etapas continua aberta e o controle continua bloqueado.
  - Ganho esperado: de 1 a 6 s a menos de espera sentida.
- **c) Corrigir o GUIA_DE_USO (Baixo).** A linha 40 diz "Etapa k de 7", mas o código abre 8 etapas.
- **Não recomendado:** devolver o controle e terminar o cálculo depois, via `Application.OnTime`.

## 4. Veredito: APROVADO_PROVISORIO
> APROVADO_PROVISORIO (UX 14/11): teto U13 Bio 22 s / Hema 30 s aceito só como proteção contra regressão, até a otimização (no máximo até 30/11/2026); não sobe de novo sem parecer de UX. Condições: (1) U13 com critério duplo: mediana de 3 na cópia ≤ 1,20 × a mediana de 3 da produção na mesma bateria (alerta em 1,10), mais o teto absoluto; (2) "aguarde (cerca de N s)" tirado do tempo medido, não fixo em 15/25 s. Meta da otimização: ≤ 8 s Bio / ≤ 10 s Hema; teto depois da otimização: 12/15 s, com resultado idêntico ao ATUALIZAR completo. A OC continua ABERTA.
