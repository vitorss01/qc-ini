# Decisões do gestor técnico e da qualidade, 06/10/2026

Decido com autonomia o que é de gestão técnica e de configuração. O que é decisão clínica formal (mudança que relaxa o controle estatístico de linhas de exame) fica **preparado para o RT e não implementado**.

---

## Decisão 1: D-01 (viés usado no Sigma e no Erro Total)

**Plano:** aprovado com ajustes, e executado.

O roteiro recebido (estudo somente leitura antes de mudar a fórmula; estimadores A, B, C, D e E) é o correto: não se troca uma métrica que governa regras de Westgard sem medir o impacto em dados reais. Acrescentei quatro itens:

1. **Estimador G** (RMS das médias de rodada) e **ANOVA entre rodadas**. Com 3 rodadas, o que separa A de B é a variação do viés de uma rodada para outra, e isso precisava ser medido.
2. **IC95% pela rodada**, além do IC por amostra. As amostras da mesma rodada não são independentes.
3. **Duas seleções de amostras**: a do Sigma de hoje (`BiasEQ`) e a do ADR-064 (`ViesEQ`), porque hoje elas diferem.
4. **D** só por interpolação, e **F** (zerar viés não significativo) apenas informativo.

**Resultado** (detalhe em `D01_estudo_estimadores_bias.md`):

- O estudo reproduz a produção em 100% das linhas (Bias, Sigma, classe e regras).
- O impacto do ADR-064 se confirma: com B, Bio 5/58 e Hema 15/57 linhas sobem; MCH N3 vai de 1,74σ para 3,25σ.
- As 20 subidas com B ocorrem todas com IC95% incluindo zero e inversão de sinal entre rodadas. No MCH, B apaga uma deriva comprovada entre rodadas (ANOVA p = 5×10⁻⁵).
- G muda Bio 3↑/1↓ e Hema 8↑/2↓. MCH N3 continua inadequado (1,71σ).

**Status:**

- **Produção inalterada: continua A.** É o estado atual, conservador, até a decisão do RT. Isso **não** é "manter A definitivamente".
- A pendência D-01 segue aberta, agora com recomendação: G com transição controlada, A rebaixado a indicador, B e o IC como indicador de direção, C só na incerteza, D na fase 2.
- A minuta de decisão do RT está na seção 12 do relatório.
- **Não** alterei fórmulas, VBA, ADRs nem regras de Westgard. O texto proposto de correção do ADR-064 (atribuição errada a Coskun 2022) e do QUALITY_GATE 14.9 está na seção 13.

**Encaminhamentos de gestão** (não dependem do RT e entram na implementação, quando houver):

- Unificar a seleção de amostras do `BiasEQ` com a do `ViesEQ` (ADR-064). Hoje não muda nenhuma classe; muda números, como Bilirrubina direta (A de 30,77% para 19,99%).
- Lipase sem ETp cadastrado (TEa = 0): o Sigma negativo é artefato de cadastro.
- Bioquímica com CV do CIQ muito alto no nível 1 (HDL N2 68%, Amilase N1 62%, AST N1 33%...): investigar em item próprio. O CV domina o Sigma da Bio, com qualquer estimador de viés.

---

## Decisão 2: `QC_Hematologia.xlsm` e `QC_Bioquimica.xlsm` no GitHub público

**Opção escolhida: (a) tirar os dois arquivos de produção do controle de versão a partir de agora (`git rm --cached` + `.gitignore`), sem reescrever o histórico, e registrar no repositório só a impressão digital (SHA-256) da versão entregue.**

### Justificativa

1. O remoto é **público**, e as versões modificadas trazem o **caminho UNC interno do servidor do laboratório** (tblConfigIntegracao). Conferi: 1 ocorrência em cada arquivo modificado e 0 na versão do HEAD. Publicar isso expõe a infraestrutura interna, por isso a opção (b) está descartada.
2. Os `.xlsm` de produção também carregam **dados do laboratório**: CEQ, CIQ, trilha de auditoria e usuários. O próprio `.gitignore` já declara que "`.xlsm` é ARTEFATO, não fonte" e que dados e cópias do laboratório não se versionam (`INTERFACEAMENTO_DB_COPIA/`, `_entrega_*/`, `_backup_pre_integracao_*/`). Os dois arquivos da raiz eram a exceção que sobrou.
3. A fonte oficial continua versionada: VBA em `src/`, snapshots e instaladores. A pasta de produção é reconstruível a partir dela. A rastreabilidade da versão em uso (ISO 15189, controle de documentos) fica garantida pelo hash registrado e pelo arquivo guardado no servidor do laboratório.
4. **Não reescrever o histórico:** o caminho UNC **nunca foi commitado**, então não há segredo novo no histórico. As versões antigas contêm só caminho local de usuário, que já é público. Um force-push num repositório público com vários ramos remotos causaria mais dano que benefício. Isso pode ser reavaliado depois, em separado, se o laboratório quiser purgar o histórico dos binários.
5. A cópia sanitizada (c) foi descartada: manter um segundo binário "limpo" em sincronia cria a duplicidade de artefato que o projeto evita, e ainda publicaria os dados do laboratório.

### Passos exatos (a executar pelo agente principal; eu não executei nada no git)

```bash
cd "<raiz do repo qc-ini>"

# 0) conferir o estado (esperado: M QC_Bioquimica.xlsm / M QC_Hematologia.xlsm)
git status --short

# 1) parar de rastrear os dois arquivos (os arquivos LOCAIS ficam intactos)
git rm --cached QC_Hematologia.xlsm QC_Bioquimica.xlsm

# 2) acrescentar ao .gitignore (ao final):
# --- Pastas de PRODUCAO do laboratorio (dados reais + configuracao interna,
# --- inclusive o caminho UNC do servidor). O remoto e publico: nunca versionar.
# --- A fonte e src/ + instaladores; a versao em uso e rastreada pelo SHA-256.
/QC_Hematologia.xlsm
/QC_Bioquimica.xlsm

# 3) registrar a impressão digital da versão entregue em 05/10/2026 (só hash e tamanho, nada de conteúdo),
#    por exemplo em _arquitetura/integracao_dbseac/MANIFESTO_PRODUCAO.md:
#    QC_Hematologia.xlsm  13167580 bytes  sha256 2af4fed57df9846dfe0cea92ebd71eba252d0f41439efbb8f834652c58e99fc9
#    QC_Bioquimica.xlsm   15348763 bytes  sha256 db53726deb3fe515d981ba0b34e5faabc863f49ed66a914778ec4d8ea7a0a390
#    (medidos em 06/10/2026 sobre os arquivos da raiz; o hash muda a cada gravação no laboratório)

# 4) conferir o que vai no commit: os dois .xlsm devem aparecer como "deleted",
#    e nada no diff staged pode conter caminho de rede ("\\servidor\...") nem senha
git status --short
git diff --cached --stat
git diff --cached -- . ':(exclude)*.xlsm' | grep -n '\\\\' || echo "ok: sem caminho UNC no texto staged"

# 5) commit (junto com os entregáveis do estudo D-01) e push para origin/main
```

**Atenção depois disso:**

- `git add -A` e `git commit -a` não pegam mais esses dois arquivos. É o comportamento desejado.
- Os scripts (`entregar.py`, instaladores) usam os arquivos locais da raiz e continuam funcionando.
- O GitHub deixa de mostrar os `.xlsm` na ponta da `main`. As versões antigas continuam no histórico (já eram públicas e não têm o UNC).
- Fora do escopo desta decisão, mas recomendo conferir no mesmo espírito os outros `.xlsm` ainda rastreados: `QC_Imunologia.xlsm` (raiz), `_entregas/*_hardening1.xlsm`, `_arquitetura/fase3a_corrigido/QC_Hematologia.xlsm` e `_arquitetura/etapa1_multilote/ref/QC_Hematologia_build_h1.xlsm`. Se tiverem dados reais ou configuração interna, devem seguir o mesmo caminho.
