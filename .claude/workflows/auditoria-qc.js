export const meta = {
  name: 'auditoria-qc',
  description: 'Ciclo de auditoria do QC_INI: Usuário Simulado testa (cópias), especialistas avaliam, Tester reproduz; devolve ao Orquestrador',
  whenToUse: 'Chamado pela skill /auditoria-qc. args: {ciclo, modo: "completo"|"reteste", produtos: ["Bioquimica","Hematologia"], frentes: ["F1","F2","F3","F4"], ocorrencias: [...] (reteste)}',
  phases: [
    { title: 'Testar', detail: 'Usuário Simulado: frentes Excel em série (F1, F2, F4) e F3 (matemática, só leitura) em paralelo' },
    { title: 'Consolidar', detail: 'Usuário Simulado junta os Relatórios A–D' },
    { title: 'Analisar', detail: 'Especialista em CQ, UX/UI e Tester (reprodução das críticas/altas)' },
  ],
}

// Regras que o script garante (ver _arquitetura/agentes/README.md):
// - um único Excel por vez: frentes com Excel e o Tester rodam em série;
// - a F3 lê a cópia com openpyxl (sem Excel) e por isso corre em paralelo;
// - nenhum agente altera src/, instaladores ou produção; quem corrige é o Orquestrador, depois.

const A = args || {}
const CICLO = A.ciclo || 'ciclo-sem-nome'
const MODO = A.modo || 'completo'
const PRODUTOS = A.produtos || ['Bioquimica', 'Hematologia']
const FRENTES = A.frentes || ['F1', 'F2', 'F3', 'F4']
const OCORR = A.ocorrencias || []
const PASTA = `_arquitetura/agentes/relatorios/${CICLO}`
// ensaio: confere só a ligação (registro dos agentes, schemas, roteamento). Ninguém abre Excel nem copia arquivo;
// cada frente devolve uma ocorrência [ENSAIO] de tipo e prioridade diferentes para exercitar o roteamento.
const ENSAIO = !!A.ensaio
const ENSAIO_OC = { F1: ['funcional', 'Alto'], F2: ['usabilidade', 'Medio'], F3: ['calculo', 'Critico'], F4: ['duvida_metodo', 'Baixo'] }
const ENSAIO_TXT = 'ENSAIO DE INTEGRAÇÃO: NÃO abra Excel, NÃO copie arquivos, NÃO rode suítes. Leia a sua definição e '
  + 'REGRAS_COMUNS.md, grave o relatório pedido com 5 a 10 linhas dizendo o que você FARIA no ciclo real, e devolva o '
  + 'formato estruturado com testes em NAO_EXECUTADO (motivo: ensaio). '

const BASE = (ENSAIO ? ENSAIO_TXT : '') + `Ciclo: ${CICLO}. Modo: ${MODO}. Leia _arquitetura/agentes/REGRAS_COMUNS.md antes de tudo e siga-o à risca `
  + `(só cópias em %TEMP%\\qcw\\auditoria\\${CICLO}\\, um Excel por vez, matar só o seu PID, nunca capturar a tela, `
  + `nada de senha ou caminho UNC no que escrever, não alterar src/ nem fórmulas). Relatórios em ${PASTA}/. `
  + `Diretório do repositório: a raiz qc-ini (os .xlsm de produção estão nela; NUNCA os abra para escrita).`

const OCORRENCIA = {
  type: 'object',
  properties: {
    titulo: { type: 'string' },
    produto: { type: 'string' },
    modulo: { type: 'string' },
    tipo: { type: 'string', enum: ['calculo', 'dados', 'funcional', 'usabilidade', 'duvida_metodo'] },
    prioridade: { type: 'string', enum: ['Critico', 'Alto', 'Medio', 'Baixo'] },
    acao: { type: 'string' },
    esperado: { type: 'string' },
    observado: { type: 'string' },
    consequencia: { type: 'string' },
    sugestao: { type: 'string' },
    evidencia: { type: 'string' },
    id_anterior: { type: 'string', description: 'no reteste: o ID OC-xxxx retestado' },
  },
  required: ['titulo', 'produto', 'modulo', 'tipo', 'prioridade', 'observado', 'evidencia'],
}

const FRENTE = {
  type: 'object',
  properties: {
    produto: { type: 'string' },
    frente: { type: 'string' },
    relatorio: { type: 'string', description: 'caminho do .md gravado' },
    sha_producao: { type: 'string' },
    testes: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          id: { type: 'string' },
          descricao: { type: 'string' },
          estado: { type: 'string', enum: ['APROVADO', 'REPROVADO', 'NAO_EXECUTADO', 'BLOQUEADO', 'RESOLVIDA', 'PERSISTE', 'EFEITO_COLATERAL'] },
          evidencia: { type: 'string' },
        },
        required: ['id', 'descricao', 'estado', 'evidencia'],
      },
    },
    calculos: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          indicador: { type: 'string' }, analito: { type: 'string' }, equipamento: { type: 'string' },
          lote: { type: 'string' }, nivel: { type: 'string' }, periodo: { type: 'string' },
          entradas: { type: 'string' }, formula: { type: 'string' },
          independente: { type: 'string' }, qcini: { type: 'string' }, diferenca: { type: 'string' },
          status: { type: 'string', enum: ['APROVADO', 'REPROVADO', 'NAO_EXECUTADO'] },
        },
        required: ['indicador', 'analito', 'independente', 'qcini', 'status'],
      },
    },
    ocorrencias: { type: 'array', items: OCORRENCIA },
    duvidas: { type: 'array', items: { type: 'string' } },
    impressoes: { type: 'string' },
  },
  required: ['produto', 'frente', 'relatorio', 'testes', 'ocorrencias'],
}

const PARECER = {
  type: 'object',
  properties: {
    ocorrencias: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          titulo: { type: 'string' },
          produto: { type: 'string' },
          veredito: { type: 'string' },
          justificativa: { type: 'string' },
          prioridade_final: { type: 'string', enum: ['Critico', 'Alto', 'Medio', 'Baixo'] },
          correcao: { type: 'string' },
          precisa_autorizacao: { type: 'boolean' },
          criterio_aceite: { type: 'string' },
          evidencia: { type: 'string' },
        },
        required: ['titulo', 'produto', 'veredito', 'justificativa', 'prioridade_final'],
      },
    },
    relatorio: { type: 'string' },
  },
  required: ['ocorrencias', 'relatorio'],
}

const CONSOLIDADO = {
  type: 'object',
  properties: {
    relatorio: { type: 'string' },
    parecer_final: {
      type: 'object',
      properties: {
        rotina_funciona: { type: 'string' }, bias: { type: 'string' }, erro_total: { type: 'string' },
        sigma: { type: 'string' }, ciq_ceq: { type: 'string' }, filtros_periodo_lote: { type: 'string' },
        entrada_manual_automatica: { type: 'string' }, divergencias_matematicas: { type: 'string' },
        problemas_operacionais: { type: 'string' }, pendencias: { type: 'string' },
      },
      required: ['rotina_funciona', 'bias', 'erro_total', 'sigma', 'ciq_ceq', 'filtros_periodo_lote',
        'entrada_manual_automatica', 'divergencias_matematicas', 'problemas_operacionais', 'pendencias'],
    },
  },
  required: ['relatorio', 'parecer_final'],
}

// Usa o tipo de agente registrado; se a sessão não o carregou (agente criado depois de abrir a sessão, ou sessão
// aberta fora da pasta onde instalar_agentes.py o colocou), roda um agente comum que carrega a mesma definição.
async function rodar(tipo, prompt, opts) {
  if (ENSAIO) opts = { ...opts, effort: 'low' }
  try {
    return await agent(prompt, { ...opts, agentType: tipo })
  } catch (e) {
    log(`agente ${tipo} fora do registro (${String(e).slice(0, 80)}); usando a definição de .claude/agents/${tipo}.md`)
    return await agent(
      `Você é o agente definido em qc-ini/.claude/agents/${tipo}.md. Leia esse arquivo INTEIRO agora e siga-o como as `
      + `suas instruções de sistema (o cabeçalho YAML só dá o nome e as ferramentas; respeite a lista de ferramentas: `
      + `os agentes especialista-cq e ux-ui não editam código nem abrem Excel).\n\n${prompt}`, opts)
  }
}

function promptFrente(produto, frente) {
  const reteste = OCORR.filter(o => !o.produto || o.produto === produto)
  const alvo = MODO === 'reteste'
    ? `RETESTE. Refaça exatamente o fluxo de cada ocorrência abaixo numa cópia NOVA da produção atual e diga RESOLVIDA, `
      + `PERSISTE ou EFEITO_COLATERAL (estado do teste com id = ID da ocorrência), refazendo também o fluxo vizinho. `
      + `Ocorrências desta frente:\n${JSON.stringify(reteste.filter(o => !o.frente || o.frente === frente), null, 1)}`
    : `Execute a frente ${frente} completa, conforme a sua definição (seção "Frentes de teste").`
  const f3 = frente === 'F3'
    ? ` IMPORTANTE: nesta frente NÃO abra Excel (outra frente o está usando). Copie a produção para a pasta do ciclo e `
      + `leia com openpyxl (data_only=True para os valores que o QC_INI calculou e gravou; os dados brutos das tabelas). `
      + `A verificação "o valor muda quando incluo ou corrijo um CEQ/CIQ" fica com as frentes F1/F2. Grave a calculadora `
      + `independente em _arquitetura/agentes/ferramentas/calc_independente.py (reutilizável, documentada).`
    : ''
  const f2 = frente === 'F2'
    ? ` Inclua nesta frente: incluir e corrigir um resultado de CEQ TESTE-AUDITORIA na cópia e conferir se o bias, o ET e `
      + `o Sigma do analito certo (e só dele) mudam como deveriam.`
    : ''
  const ens = ENSAIO && ENSAIO_OC[frente]
    ? ` Devolva exatamente UMA ocorrência fictícia com título começando por "[ENSAIO]", tipo "${ENSAIO_OC[frente][0]}" e `
      + `prioridade "${ENSAIO_OC[frente][1]}".`
    : ''
  return `${BASE}\n\nProduto: ${produto}. Frente: ${frente}.\n${alvo}${f3}${f2}${ens}\n\n`
    + `Grave o relatório em ${PASTA}/${produto}_${frente}.md e responda no formato estruturado.`
}

// ---------------------------------------------------------------- Testar
phase('Testar')
const comExcel = []
for (const p of PRODUTOS) for (const f of FRENTES) if (f !== 'F3') comExcel.push([p, f])
const semExcel = []
for (const p of PRODUTOS) if (FRENTES.includes('F3')) semExcel.push([p, 'F3'])

async function serieExcel() {
  const out = []
  for (const [p, f] of comExcel) {
    log(`Excel: ${p} ${f}`)
    out.push(await rodar('usuario-simulado', promptFrente(p, f), { label: `usuario:${p}:${f}`, phase: 'Testar', schema: FRENTE }))
  }
  return out
}
const [rExcel, rF3] = await parallel([
  () => serieExcel(),
  () => parallel(semExcel.map(([p, f]) => () =>
    rodar('usuario-simulado', promptFrente(p, f), { label: `usuario:${p}:${f}`, phase: 'Testar', schema: FRENTE }))),
])
const frentes = [...(rExcel || []), ...(rF3 || [])].filter(Boolean)
const faltaram = comExcel.length + semExcel.length - frentes.length
if (faltaram) log(`ATENÇÃO: ${faltaram} frente(s) sem resultado (agente interrompido ou falhou)`)
const todas = frentes.flatMap(r => (r.ocorrencias || []).map(o => ({ ...o, frente: r.frente })))
log(`${frentes.length} frentes, ${todas.length} ocorrências`)

// ---------------------------------------------------------------- Consolidar
phase('Consolidar')
const consolidado = await rodar('usuario-simulado', 
  `${BASE}\n\nPARECER CONSOLIDADO do ciclo. Leia os relatórios das frentes abaixo e escreva ${PASTA}/00_PARECER.md com os `
  + `Relatórios A (experiência), B (validação matemática, a tabela completa), C (ocorrências, com ID provisório P-01, P-02...; `
  + `o ID definitivo OC-xxxx é dado pelo Orquestrador) e D (parecer final). Não abra Excel. Responda "validado" só onde houve `
  + `teste executado e aprovado.\n\nResultados das frentes:\n${JSON.stringify(frentes, null, 1)}`,
  { label: 'usuario:parecer', phase: 'Consolidar', schema: CONSOLIDADO })

// ---------------------------------------------------------------- Analisar
phase('Analisar')
const paraCQ = todas.filter(o => ['calculo', 'dados', 'duvida_metodo'].includes(o.tipo))
const paraUX = todas.filter(o => o.tipo === 'usabilidade' || (o.tipo === 'funcional' && o.prioridade !== 'Critico'))
const paraTester = todas.filter(o => o.prioridade === 'Critico' || o.prioridade === 'Alto')

const [cq, ux, tester] = await parallel([
  () => paraCQ.length ? rodar('especialista-cq', 
    `${BASE}\n\nDê o parecer técnico sobre estas ocorrências (relatórios das frentes em ${PASTA}/). Grave ${PASTA}/parecer_cq.md.\n`
    + JSON.stringify(paraCQ, null, 1),
    { label: 'especialista-cq', phase: 'Analisar', schema: PARECER }) : null,
  () => paraUX.length ? rodar('ux-ui', 
    `${BASE}\n\nAvalie estas ocorrências de usabilidade (relatórios das frentes em ${PASTA}/). Grave ${PASTA}/parecer_ux.md.\n`
    + JSON.stringify(paraUX, null, 1),
    { label: 'ux-ui', phase: 'Analisar', schema: PARECER }) : null,
  // o Tester usa Excel, mas a fase Testar já terminou: é o único Excel agora
  () => paraTester.length ? rodar('tester-qa', 
    `${BASE}\n\nModo REPRODUZIR. Reproduza de forma independente cada ocorrência crítica/alta abaixo, em cópias novas. `
    + `Grave ${PASTA}/tester_reproduzir.md. Veredito por ocorrência: REPRODUZIDA, NAO_REPRODUZIDA ou PARCIAL.\n`
    + JSON.stringify(paraTester, null, 1),
    { label: 'tester-qa', phase: 'Analisar', schema: PARECER }) : null,
])

return {
  ciclo: CICLO, modo: MODO, ensaio: ENSAIO, pasta: PASTA,
  roteamento: { cq: paraCQ.map(o => o.titulo), ux: paraUX.map(o => o.titulo), tester: paraTester.map(o => o.titulo) },
  frentes: frentes.map(r => ({ produto: r.produto, frente: r.frente, relatorio: r.relatorio, sha: r.sha_producao,
    testes: r.testes, calculos: r.calculos })),
  ocorrencias: todas,
  consolidado, parecer_cq: cq, parecer_ux: ux, tester,
  frentes_sem_resultado: faltaram,
}
