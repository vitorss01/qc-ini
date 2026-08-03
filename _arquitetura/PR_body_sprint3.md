# Sprint HARDENING 3 — trilha de auditoria e conserto do pipeline de build

> **Abrir como DRAFT.** O `QUALITY_GATE.md` está em **15 ✅ · 27 ⏳ · 4 ❌**, e a regra do
> projeto é explícita: nada entra na `main` enquanto houver ⏳ ou ❌.

---

## O achado que muda o estado do projeto

**O motor corrigido nunca entrou no `.xlsm`** — nem nesta máquina, nem na do trabalho.

Verificado lendo o VBA de dentro dos três artefatos:

| | versionado (fonte) | dentro do `.xlsm` |
|---|---|---|
| `mEstatistica` | 1.049 linhas | 1.028 linhas |
| ocorrências de `aS` | 0 | **8** |
| referencia `Eng_Saida` | sim | **não** |

As 4.362 fórmulas redirecionadas apontavam para uma camada que ninguém abastecia. O item
**2.1 estava ✅ com evidência verdadeira sobre as planilhas e falsa sobre o sistema**.

Os itens **2.1 e 2.1b voltam para ⏳**, e o entregável
`_entregas/QC_Hematologia_hardening1.xlsm` precisa ser regerado.

## Quatro defeitos no pipeline — nenhum em VBA, nenhum em lógica de CQI

1. **`Select-Object -First` encerra o script produtor.** Comportamento documentado do
   PowerShell (`StopUpstreamCommandsException`). As etapas morriam depois de imprimir as
   primeiras mensagens e **antes** do `Save()`. `criar_eng_saida` usava `-Last`, que
   bufferiza e não interrompe — por isso só ela persistia, e a aba `Corridas` nunca
   aparecia no artefato.
2. **`$xl.Quit()` não mata o processo** enquanto o PowerShell segura referências COM. O
   Excel órfão travava o arquivo e a etapa seguinte o abria em **somente leitura**, sem
   aviso, porque `DisplayAlerts = $false`.
3. **Nenhum dos 12 scripts tinha `$ErrorActionPreference = 'Stop'`.** Erro de COM é
   não-terminante: falha virava mensagem de sucesso.
4. **`aplicar_vba.ps1` imprimia "VBA aplicado" sem conferir nada.**

**Lição estrutural:** havia lint de VBA, diff célula a célula e varredura de arquitetura —
e nenhum controle verificava se a etapa anterior tinha acontecido. O pipeline reportava
**intenção, não resultado**.

## Trilha de auditoria (itens 3.1 e 3.2)

`mAuditoria.bas` — log append-only **encadeado por hash**: cada linha carrega o SHA-256 da
anterior somada ao próprio conteúdo.

A proteção de planilha do Excel não resiste a fraude deliberada — a senha está em texto no
projeto VBA e a marcação sai descompactando o arquivo. Sem a cadeia, o log seria "confie em
mim". Com ela, adulteração vira **evidência verificável**: `VerificarIntegridadeLog` aponta
em que linha a cadeia quebrou.

- **Identidade de registro** = login do sistema (`Usuarios`, senha em hash). Windows,
  Office e nome da máquina entram como **corroboração** — vêm de variável de ambiente.
- **Limite declarado no código:** a hora vem do relógio da máquina, que o usuário pode
  mudar. Não há solução dentro do Excel; fica registrado, não escondido.
- **Parecer Técnico** com mínimo de 5 palavras reais.

## Item 2.3 fechado no código

`UpsertResultados` forçava `Status = Ativo` em toda atualização — reenviar a mesma chave
**ressuscitava registro excluído**, sem rastro. Era vetor de fraude. Agora a linha não ativa
é preservada e a tentativa é auditada como `UPSERT_BLOQUEADO`.

## O que está provado e o que não está

| | |
|---|---|
| Pipeline consertado e verificável | ✅ 20 scripts checados por analisador de sintaxe |
| Artefato com o motor correto | ✅ 21 abas, `aS=0`, `Eng_Saida` populada |
| Motor executando | ✅ chave `ANALITO\|RUN` publicada |
| Painel/Estatística lendo a nova camada | ❌ não verificado |
| Diff das 4.362 fórmulas refeito | ❌ não refeito |
| Auditoria funcionando dentro do sistema | ❌ não testada |

## Para revisar

A produção (`QC_Hematologia.xlsm` na raiz) **não foi tocada** — segue idêntica ao backup,
61.416 fórmulas. Bioquímica e Imunologia intactas na Fase 2.

A pasta de build saiu para `%USERPROFILE%\QCINI_build_hardening1`, fora do OneDrive.
O nome precisa conter `build_hardening` porque `rodar_motor.ps1` recusa executar fora de
uma cópia de build — trava que existe para o motor nunca rodar sobre a produção.
