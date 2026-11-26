# -*- coding: utf-8 -*-
"""QA do FLUXO DO CEQ pela janela, de ponta a ponta: o controle externo inserido como o ANALISTA faz
(mCEQForm.ConfirmacoesCEQ + SalvarCEQ / CorrigirCEQ / ExcluirCEQ, login pelo DoLogin) gera bias, Erro Total e Sigma
corretos. Roda numa COPIA instalada (em casa: modo HISTORICO, sem DB_SEAC); nada e salvo no arquivo.

Uso: python qa_fluxo_ceq.py <Bioquimica|Hematologia> <copia.xlsm> <saida.json>

Nenhum valor esperado e fixo: o analito, o ano (eqAnoEP), a rodada (eqRodada), o provedor da linha, a EQA_Base de
antes, o CV pooled (tblCQ_Final) e o ETp sao lidos da copia em tempo de execucao; os oraculos sao os das suites
vizinhas (qa_rodadas_cv.Oraculo / k5_esperado, qa_incerteza.esperado_ciq). So os valores DIGITADOS sao do teste
(AMOSTRAS, CORRECAO) e o esperado e calculado deles aqui.

  Preparo  ATUALIZAR so entradas + consolidacao da EQA_Base; usuario ANALISTA de teste + DoLogin; analito escolhido
           em tempo de execucao: nome no mapa do CAP, provedor da linha = CAP (Bio: Especificacoes AR; Hema: R4),
           CV do periodo (F), ETp (K) e CV pooled da janela do Sigma em TODOS os niveis (de preferencia com G numerico
           antes); Painel nesse analito; P4 que nao deixaria a rodada nova entrar (rotulo, conjunto, ULTIMAS n,
           texto invalido) e fixado em TODAS SO NESTA COPIA (como o qa_rodadas_cv; volta no fim); FOTO (Estatistica
           A14:AU, K5, Painel A7:J, EQA_Base, tabelas, listas da Configuracao AF:AI, N4/P4)
  F01 ConfirmacoesCEQ vazia e SalvarCEQ 'OK|linha|' sem aviso nas 3 amostras; O e P da linha (formula) =
      (F-G)/G*100 e |O|
  F02 EQA_Base: cada chave nova com canonico, ACEITO, Uso SIM e bias; o resto da base intacto; a rodada entra na
      lstRodadasCAP; Audit_Log com 3 CEQ_REGISTRADO e 3 EQA_BASE_CONSOLIDADA
  F03 G, T, AC, AD do analito = Oraculo(base de antes + 3 linhas novas) (e a FOTO = Oraculo(base de antes)); as 3
      amostras ENTRARAM (AC = antes + 3, AD = antes + 1: sem isso o oraculo concordaria com um G parado); K5 =
      k5_esperado antes e depois; as linhas dos outros analitos nao mudam
  F04 isolamento: P4 = '<ano>|QA-FLUXO 1' da os valores fechados das AMOSTRAS (media de |b|, media de b, 3, 1), H e L
      com eles; os outros analitos = Oraculo (sem dado nessa rodada); P4 volta e o analito volta
  F05 ET: H = 1,65*F + |G| em cada nivel (na FOTO e depois)
  F06 Sigma: L = (K - |G|)/CVp (esperado_ciq na janela Sigma_Ini..MU_Fim, com Sigma_Ini = EDATE(MU_Fim,-meses)+1),
      classe M pela escada (na FOTO e depois)
  F07 Painel: G = Estatistica G; H = 1,65*E + |G| (vazio sem E); I = Estatistica L (na FOTO e depois)
  F08 CorrigirCEQ da amostra 1 (CORRECAO): mesma linha (a da lista RegistrosDaRodada), EQA_Base, G/T/AC/AD pelo
      oraculo, H, L, Painel, CEQ_ALTERADO "antes -> depois"
  F09 ExcluirCEQ das 3: Estatistica, K5, Painel, EQA_Base, tabelas (so as linhas com analito) e listas voltam
      EXATAMENTE a FOTO (nunca ListRows.Count: na Hema a tabela do CAP ganha uma linha vazia); 3 CEQ_EXCLUIDO
  F10 ALERTA avaliacao nao reconhecida ('Adequado') -> NAO AVALIADO, fora do bias, sem aviso
  F11 ALERTA rodada com ano = Ano EP + 1 nao muda G (com Ano EP = TODOS e P4 = TODAS passa a entrar)
  F12 ALERTA Controllab mapeado num analito cujo provedor do bias e o CAP (OC-0069): nao entra, sem aviso
      F10-F12 ficam com resultado 'ALERTA' (nao FAIL) quando o risco se confirma; PASS se o modulo passar a avisar
      (ConfirmacoesCEQ, ou o aviso do OK, com "NAO entra no bias", "fora do bias" ou "Ano EP"); FAIL se o calculo
      divergir do oraculo ou a limpeza nao voltar a FOTO
  F13 limpeza: nenhum registro nem linha de mapa do teste fica na copia; N4/P4 como estavam; tudo igual a FOTO
  F00 seguranca do teste: arquivo de entrada intacto (SHA-256 no inicio e no fim)
Pre-condicao que falta = FAIL explicito "pre-condicao ausente" (nunca pula calado).
"""
import collections
import json
import os
import sys
import time
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo, sem_acento  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login, USUARIO_QA  # noqa: E402
from qa_rodadas_cv import (Oraculo, k5_esperado, ler_base, linhas_tabela, ultima_linha, provedores_linha,  # noqa: E402
                           recalc, num, perto, igual_valor, ci, U, modo, LISTAS)
from qa_incerteza import esperado_ciq, ler_ciq, dia, dia_menos_meses  # noqa: E402

SETA = '→'
PISO0 = [False]                  # ADR-101: a copia ja tem Sigma com piso 0 e 6 faixas (lido da formula de L14)
TOL = 1e-9
TOL_SIGMA = 1e-6                 # CV pooled recalculado por fora: a mesma tolerancia do qa_incerteza (I14/I16)
ROD = 'QA-FLUXO 1'
ARQ = 'TESTE qa_fluxo_ceq'       # nunca "SIMULA" (Uso = NAO tiraria a linha do bias)
NOME_CTL_TMP = 'TESTE qa_fluxo_ceq CTL'
# valores DIGITADOS (amostra, resultado, alvo): bias +5 %, -4 %, +5 %; todo esperado sai deles
AMOSTRAS = [('QAFC-01', 105, 100), ('QAFC-02', 96, 100), ('QAFC-03', 210, 200)]
CORRECAO = 110                   # amostra 1 corrigida: +10 %
TERMO_F10 = 'Adequado'           # avaliacao que o PadronizarStatus nao reconhece (ADR: vira NAO AVALIADO)
DESVIOS_F10 = (12, 17, 23)       # bias (%) da amostra do F10: o primeiro que MUDARIA o G se fosse avaliada
AVISOS = ('NAO ENTRA NO BIAS', 'FORA DO BIAS', 'ANO EP')
COLS = ('F', 'G', 'H', 'K', 'L', 'M', 'T', 'AC', 'AD')
METRICAS = (('G', 'ABS'), ('T', 'SIGNED'), ('AC', 'N'), ('AD', 'NRODADAS'))


class PreCondicao(Exception):
    def __init__(self, motivo, evid=None, codigo='F01'):
        super().__init__(motivo)
        self.motivo, self.evid, self.codigo = motivo, evid or {}, codigo


def txt(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return '' if v is None else str(v).strip()


def s(v):
    return '' if v is None else str(v)


def letra(j):
    out = ''
    j += 1
    while j:
        j, r = divmod(j - 1, 26)
        out = chr(65 + r) + out
    return out


def alerta(teste, evid):
    """Risco conhecido confirmado com evidencia: nao reprova a entrega, mas aparece na tabela e no log."""
    RES.append({'teste': teste, 'resultado': 'ALERTA', 'evidencia': evid})
    print('ALERTA ' + teste + ' -- ' + json.dumps(evid, ensure_ascii=False, default=str)[:600], flush=True)


def classe_sigma(x, seis=True):
    """ADR-103 (26/11/2026): REGRA INSTITUCIONAL da Metrica Sigma, operadores literais (sigma = 4 Marginal, sigma = 5 Bom)."""
    if x >= 6:
        return 'Excelente — Classe mundial'
    if 5 < x < 6:
        return 'Excelente'
    if 4 < x <= 5:
        return 'Bom'
    if 3 <= x <= 4:
        return 'Marginal'
    return 'Insatisfatório'


def bias_pct(res, alvo):
    return (res - alvo) / alvo * 100.0


def campos(prov, rod, ano, nome, am, res, alvo, aval='Acceptable'):
    """Os 16 textos da janela (frmCEQ.Campos): Provedor, Rodada, Ano, Analito, Amostra, Resultado, Alvo, SD, SDI,
    LimInf, LimSup, Avaliacao, NLabs, Unidade, Pagina, Arquivo."""
    return [prov, rod, str(ano), nome, am, str(res), str(alvo), '', '', '', '', aval, '', '', '', ARQ]


def linha_base(prov, ano, rod, nome, canon, am, res, alvo, aval, status):
    """Linha que a consolidacao TEM de gravar na EQA_Base (indices 0..20), escrita do que foi digitado:
    provedor literal, ano, rodada, nome do provedor, canonico do mapa, ..., status padronizado, bias, |bias|, Uso SIM
    (arquivo sem SIMULA), chave provedor|ano|rodada|analito|amostra."""
    b = bias_pct(res, alvo)
    return [prov, float(ano), rod, nome, canon, am, float(res), float(alvo), None, None, None, None, aval, status, '',
            b, abs(b), None, ARQ, 'SIM', f'{prov}|{ano}|{rod}|{nome}|{am}']


def conf_linha_base(got, esp):
    """Campos da linha consolidada que importam ao bias (1e-9 nos numeros, texto sem diferenca de maiusculas)."""
    if got is None:
        return ['chave ausente da EQA_Base']
    dif = []
    for i, nome in ((0, 'provedor'), (2, 'rodada'), (3, 'analito'), (4, 'canonico'), (5, 'amostra'), (12, 'avaliacao'),
                    (13, 'status'), (18, 'arquivo'), (19, 'uso'), (20, 'chave')):
        if U(got[i]) != U(esp[i]):
            dif.append((nome, got[i], esp[i]))
    for i, nome in ((1, 'ano'), (6, 'resultado'), (7, 'alvo'), (15, 'bias'), (16, '|bias|')):
        if not perto(got[i], esp[i], TOL):
            dif.append((nome, got[i], esp[i]))
    return dif


def avisou(*textos):
    """Marcas de aviso do modulo (sem acento, sem diferenca de maiusculas)."""
    t = ' '.join(sem_acento(x).upper() for x in textos if x)
    return [a for a in AVISOS if a in t]


def aviso_ok(retorno):
    """Terceiro campo de 'OK|<linha>|<aviso>'."""
    p = s(retorno).split('|', 2)
    return p[2] if len(p) > 2 else ''


def linha_ok(retorno):
    return int(s(retorno).split('|')[1]) if s(retorno).startswith('OK|') else 0


def chave_rep(r):
    return tuple(repr(x) for x in r)


def cheias(rows):
    """Linhas da EQA_Base com algum valor (ler_base devolve a linha 2 vazia quando a base esta vazia)."""
    return [r for r in rows if any(x not in (None, '') for x in r)]


def k5_confere(partes, tipo, texto):
    """K5 = ResumoFiltroEQ: contagem do recorte; provedor vazio nos modos novos / texto invalido tem o aviso proprio
    (mesmas palavras do R06 do qa_rodadas_cv)."""
    if tipo == 'PROVEDOR':
        return 'provedor nao definido' in texto
    if tipo == 'INVALIDA':
        return 'INVALIDA' in texto
    return tipo is None and all(pp in texto for pp in partes)


def dif_lista(a, b, n=6):
    out = []
    for i, (x, y) in enumerate(zip(a, b)):
        if x != y:
            out.append((i, x, y))
            if len(out) >= n:
                break
    if len(a) != len(b):
        out.append(('tamanho', len(a), len(b)))
    return out


def dif_foto(f0, f1):
    """Diferencas EXATAS entre duas fotos (vazio = voltou bit a bit)."""
    out = {}
    d, mx = [], 0.0
    for i, (r0, r1) in enumerate(zip(f0['est'], f1['est'])):
        for j, (a, b) in enumerate(zip(r0, r1)):
            if a != b:
                d.append((14 + i, letra(j), a, b))
                if num(a) and num(b):
                    mx = max(mx, abs(a - b))
    if len(f0['est']) != len(f1['est']):
        d.append(('linhas', len(f0['est']), len(f1['est'])))
    if d:
        out['Estatistica'] = {'n': len(d), 'primeiras': d[:8], 'maior_dif_numerica': mx}
    for k in ('K5', 'N4', 'P4'):
        if f0[k] != f1[k]:
            out[k] = (f0[k], f1[k])
    dp = [(7 + i, letra(j), a, b) for i, (r0, r1) in enumerate(zip(f0['painel'], f1['painel']))
          for j, (a, b) in enumerate(zip(r0, r1)) if a != b]
    if dp:
        out['Painel'] = dp[:8]
    # so as linhas com conteudo: a exclusao pode deixar linha vazia no fim da EQA_Base (sem efeito no bias)
    b0, b1 = cheias(f0['base']), cheias(f1['base'])
    if b0 != b1:
        c0 = collections.Counter(chave_rep(r) for r in b0)
        c1 = collections.Counter(chave_rep(r) for r in b1)
        out['EQA_Base'] = {'n_antes': len(b0), 'n_depois': len(b1), 'mesmo_conjunto': c0 == c1,
                           'diferencas': dif_lista(b0, b1, 3), 'linhas_vazias': [len(f0['base']) - len(b0),
                                                                                  len(f1['base']) - len(b1)]}
    for pv in f0['tab']:
        if f0['tab'][pv] != f1['tab'].get(pv):
            out['tabela ' + pv] = {'n_antes': len(f0['tab'][pv]), 'n_depois': len(f1['tab'].get(pv) or []),
                                   'diferencas': dif_lista(f0['tab'][pv], f1['tab'].get(pv) or [], 3)}
    for nm in f0['listas']:
        if f0['listas'][nm] != f1['listas'].get(nm):
            out[nm] = {'antes': f0['listas'][nm], 'depois': f1['listas'].get(nm)}
    return out


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    bio = q.bio
    nlv = q.nlv
    incluidos = []                   # (provedor, linha da aba) ainda gravados por este teste
    mapa_tmp = []                    # linhas do mapa W:Y criadas pelo F12
    estado = {'foto0': None, 'f13': False, 'b3': None, 'nomes': {}, 'p4_forcado': False, 'p4_original': None}
    try:
        xl = q.ex.xl
        e = q.ws('Estatística')
        PISO0[0] = 'MAX(0,' in str(e.Range('L14').Formula).replace(' ', '')   # ADR-101
        p = q.ws('Painel')
        base_ws = q.ws('EQA_Base')
        al = q.ws('Audit_Log')

        def run(macro, *a, teto=900):
            return q.run(macro, *a, teto=teto)

        def aud0():
            return al.Cells(al.Rows.Count, 1).End(-4162).Row

        def acoes(r0, acao):
            ult_ = aud0()
            return [txt(al.Cells(r, 24).Value) for r in range(r0 + 1, ult_ + 1) if txt(al.Cells(r, 7).Value) == acao]

        def tabela(nome):
            return next((t for w in q.wb.Worksheets for t in w.ListObjects if t.Name == nome), None)
        LO = {'CAP': tabela('tblEQA_CAP_Dados'), 'Controllab': tabela('tblEQA_Controllab_Dados')}

        def linhas_lo(prov):
            lo = LO[prov]
            if lo is None or lo.ListRows.Count == 0:
                return []
            r0 = lo.DataBodyRange.Row
            return [(r0 + i, list(row[:18])) for i, row in enumerate(lo.DataBodyRange.Value) if txt(row[3])]

        def do_teste():
            return [(pv, r) for pv in LO for r, row in linhas_lo(pv) if txt(row[17]) == ARQ]

        try:
            cfg_ws = q.ws('Configuração')
        except Exception:                            # noqa: BLE001
            cfg_ws = None

        def lista(nome):
            """A coluna que o mDados.AtualizarListasAno grava (Configuracao AF..AI, linhas 2..101) -- a mesma leitura do
            R03 do qa_rodadas_cv. Pelo nome definido, um nome ausente dava None antes e depois e a comparacao passava
            sem testar nada."""
            col = LISTAS[nome]
            return [txt(v[0]) for v in cfg_ws.Range(f'{col}2:{col}101').Value if txt(v[0])]

        def definir_nome(nome, valor):
            rg = q.nome(nome)
            if nome not in estado['nomes']:
                estado['nomes'][nome] = rg.Value
            try:
                rg.Worksheet.Unprotect(ii.SENHA)
            except Exception:                        # noqa: BLE001
                pass
            xl.EnableEvents = False
            try:
                rg.Value = valor
            finally:
                xl.EnableEvents = True
            recalc(q)

        def excluir(prov, lin):
            r_ = s(run('mCEQForm.ExcluirCEQ', prov, lin))
            if r_.startswith(f'OK|{lin}|') and (prov, lin) in incluidos:
                incluidos.remove((prov, lin))
            return r_

        def salvar(c):
            cf = s(run('mCEQForm.ConfirmacoesCEQ', c, 0))
            r_ = s(run('mCEQForm.SalvarCEQ', c))
            lin = linha_ok(r_)
            if lin:
                incluidos.append(('Controllab' if U(c[0]) == 'CONTROLLAB' else 'CAP', lin))
            return cf, r_, lin

        def limpar():
            ev = {'excluidos': {}, 'mapa_limpo': [], 'nomes_restaurados': {}}
            for pv, lin in list(incluidos) + [x for x in do_teste() if x not in incluidos]:
                ev['excluidos'][f'{pv}|{lin}'] = excluir(pv, lin)[:120]
            if mapa_tmp:
                base_ws.Unprotect(ii.SENHA)
                xl.EnableEvents = False
                try:
                    for r in list(mapa_tmp):
                        base_ws.Range(f'W{r}:Y{r}').ClearContents()
                        mapa_tmp.remove(r)
                        ev['mapa_limpo'].append(r)
                finally:
                    xl.EnableEvents = True
            for nome, v0 in estado['nomes'].items():
                if q.nome(nome).Value != v0:
                    definir_nome(nome, v0)
                    ev['nomes_restaurados'][nome] = v0
            recalc(q)
            return ev

        def foto():
            recalc(q)
            return {'est': [list(v) for _, v in linhas_tabela(e, ult)], 'K5': s(e.Range('K5').Text),
                    'painel': [list(x) for x in p.Range(f'A7:J{6 + nlv}').Value], 'base': ler_base(q),
                    'tab': {pv: [row for _, row in linhas_lo(pv)] for pv in LO},
                    'listas': {n: lista(n) for n in ('lstRodadasCAP', 'lstRodadasCTL', 'lstAnosCAP', 'lstAnosCTL')},
                    'N4': q.nome('eqAnoEP').Value, 'P4': q.nome('eqRodada').Value}

        # ============================================================ preparo
        if LO['CAP'] is None:
            raise PreCondicao('tabela tblEQA_CAP_Dados não encontrada')
        if cfg_ws is None:
            raise PreCondicao('aba Configuração (listas lstRodadas*/lstAnos* do mDados) não encontrada')
        r_at = s(run('mIntegracao.AtualizarSoEntradas', teto=1500))
        if not r_at.startswith('OK|'):
            raise PreCondicao('ATUALIZAR só entradas não terminou com OK', {'retorno': r_at[:400]})
        run('mEQA.ConsolidarEQABaseDe', 'qa_fluxo_ceq: preparo')
        recalc(q)
        criar_usuario(q, USUARIO_QA, 'ANALISTA')
        usuario = fazer_login(q, USUARIO_QA)
        if run('mSeguranca.SessaoAtiva') is not True:
            raise PreCondicao('login ANALISTA de teste (DoLogin) sem sessão ativa', {'currentUser': usuario})
        sobra0 = do_teste() + [txt(r[20]) for r in ler_base(q) if txt(r[18]) == ARQ]
        if sobra0:
            raise PreCondicao('a cópia já tem registros deste teste (execução anterior interrompida?)',
                              {'registros': sobra0[:6]})

        ano_val = q.nome('eqAnoEP').Value
        if not (num(ano_val) and float(ano_val).is_integer() and 1990 <= ano_val <= 2100):
            raise PreCondicao('Ano EP (Estatística N4, eqAnoEP) não é um ano: o registro usa ano = eqAnoEP',
                              {'eqAnoEP': ano_val})
        ano = int(ano_val)
        rod_val = q.nome('eqRodada').Value
        # Com P4 = rotulo, conjunto, ULTIMAS n ou texto invalido a rodada 'QA-FLUXO 1' nao entraria no G e o F03 passaria
        # sem testar a cadeia (oraculo e planilha concordando com o G parado). So TODAS/vazio (ano exato = N4) e
        # ACUMULADAS deixam entrar com certeza: os outros modos sao fixados em TODAS NESTA COPIA (nada e salvo; o
        # valor original volta no fim), como o forcar_todas do qa_rodadas_cv.
        if modo(rod_val)[0] not in (0, 3):
            p4 = q.nome('eqRodada')
            estado['p4_original'], estado['p4_forcado'] = p4.Formula, True
            try:
                p4.Worksheet.Unprotect(ii.SENHA)
            except Exception:                        # noqa: BLE001
                pass
            xl.EnableEvents = False
            try:
                p4.Value = 'TODAS'
            finally:
                xl.EnableEvents = True
            recalc(q)
            print(f'   P4 da cópia era {rod_val!r}: fixado em TODAS (só nesta cópia; volta no fim)', flush=True)
            rod_val = q.nome('eqRodada').Value
        prov_k5 = q.nome('eqProvedor').Value
        prov_de = provedores_linha(q, bio)

        # janela do Sigma e CV pooled independente (qa_incerteza.esperado_ciq)
        try:
            fim = dia(q.nome('MU_Fim').Value)
            meses = q.nome('Sigma_Meses').Value
            ini = dia(q.nome('Sigma_Ini').Value)
        except Exception as ex_:                     # noqa: BLE001
            raise PreCondicao('janela do Sigma ilegível (MU_Fim, Sigma_Meses, Sigma_Ini)', {'erro': repr(ex_)[:200]})
        janela = {'Sigma_Ini': s(q.nome('Sigma_Ini').Text), 'MU_Fim': s(q.nome('MU_Fim').Text), 'meses': meses,
                  'Sigma_Ini_esperado': dia_menos_meses(fim, int(meses)) if num(meses) else None, 'Sigma_Ini_lido': ini}
        janela_ok = num(meses) and janela['Sigma_Ini_esperado'] == ini
        exc = []
        # so a Bio tem Estat_Exclusoes; quando o NOME existe ele entra no Sigma (IncertezaCIQ) e tem de ser lido --
        # erro de leitura vira pre-condicao, nunca "sem exclusoes" calado (o CVp por fora divergiria)
        if any(n_.Name == 'Estat_Exclusoes' for n_ in q.wb.Names):
            try:
                for r_ in q.nome('Estat_Exclusoes').Value:
                    if r_[0] and r_[1]:
                        x_, y_ = dia(r_[0]), dia(r_[1])
                        if y_ >= x_:
                            exc.append((x_, y_))
            except Exception as ex_:                 # noqa: BLE001
                raise PreCondicao('Estat_Exclusoes existe mas não foi lida', {'erro': repr(ex_)[:200]})
        try:
            eqf = s(q.nome('selEquipamento').Value).strip().upper()
        except Exception:                            # noqa: BLE001
            eqf = ''
        esp_cv = esperado_ciq(q, eqf, ini, fim, exc, ler_ciq(q))

        # analito escolhido em tempo de execucao
        ult = ultima_linha(e)
        E0 = linhas_tabela(e, ult)
        E0d = dict(E0)
        por_an = collections.OrderedDict()
        for r, v in E0:
            an_ = txt(v[0])
            if an_ and an_ != '0' and '|' in txt(v[ci('AB')]):
                por_an.setdefault(an_, []).append((r, txt(v[1])))
        mapa_cap = run('mEQA.MapaDoProvedor', 'CAP') or ()
        nomes_cap = collections.defaultdict(list)
        for m in mapa_cap:
            nomes_cap[U(m[1])].append((txt(m[0]), txt(m[1])))
        uso_nome = collections.Counter(U(row[3]) for _, row in linhas_lo('CAP'))
        sel0 = txt(q.nome('selAnalito').Value)
        cand, recusados = [], {}
        for ordem, (an_, rows) in enumerate(por_an.items()):
            motivo = []
            if not nomes_cap.get(U(an_)):
                motivo.append('sem nome no mapa do CAP')
            if U(prov_de(an_)) != 'CAP':
                motivo.append(f'provedor da linha {prov_de(an_)!r}')
            for r, nv in rows:
                v = E0d[r]
                if not (num(v[ci('K')]) and v[ci('K')] > 0):
                    motivo.append(f'N{nv} sem ETp')
                if not num(v[ci('F')]):
                    motivo.append(f'N{nv} sem CV do período')
                if not (esp_cv.get((U(an_), nv)) or {}).get('cv'):
                    motivo.append(f'N{nv} sem CV pooled')
            if motivo:
                recusados[an_] = motivo[:3]
                continue
            g_num = all(num(E0d[r][ci('G')]) for r, _ in rows)
            cand.append(((not g_num, U(an_) != U(sel0), ordem), an_))
        if not cand:
            raise PreCondicao('nenhum analito com nome no mapa do CAP, provedor da linha CAP, CV do período, ETp e CV '
                              'pooled em todos os níveis', {'recusados': dict(list(recusados.items())[:12])})
        an = min(cand)[1]
        linhas_an = por_an[an]
        nome_cap, canon = max(nomes_cap[U(an)], key=lambda x: (uso_nome[U(x[0])], -nomes_cap[U(an)].index(x)))
        pv_an = prov_de(an)

        # Painel no analito (so troca se ja nao estiver nele)
        if U(sel0) != U(an):
            raw = [txt(v[0]) for v in q.espec().Range('A4:A43').Value]
            pos = next((i for i, x in enumerate(raw, 1) if U(x) == U(an)), None)
            if pos is None:
                raise PreCondicao('analito fora do catálogo (Especificações A4:A43)', {'analito': an})
            estado['b3'] = p.Range('B3').Value
            p.Unprotect(ii.SENHA)
            xl.EnableEvents = False
            try:
                p.Range('B3').Value = pos
            finally:
                xl.EnableEvents = True
            run('mEstatistica.PainelMudou', teto=600)
        recalc(q)
        if U(q.nome('selAnalito').Value) != U(an):
            raise PreCondicao('o Painel não passou a mostrar o analito escolhido',
                              {'analito': an, 'selAnalito': txt(q.nome('selAnalito').Value)})

        def vals_an():
            recalc(q)
            out = {}
            for r, nv in linhas_an:
                v = e.Range(f'A{r}:AU{r}').Value[0]
                out[nv] = {c: v[ci(c)] for c in COLS}
            return out

        def painel():
            return [list(x) for x in p.Range(f'A7:J{6 + nlv}').Value]

        def conf_g(rows_base, L, rod=None, ano_=None):
            """G/T/AC/AD do analito = Oraculo (qa_rodadas_cv) sobre a base dada, com o N4/P4 dados (ou os da copia)."""
            orc = Oraculo(rows_base)
            a_ = ano_val if ano_ is None else ano_
            r_ = rod_val if rod is None else rod
            esp = {c: orc.bias(an, a_, pv_an, r_, met) for c, met in METRICAS}
            dif = [(nv, c, v[c], esp[c]) for nv, v in L.items() for c in esp if not igual_valor(v[c], esp[c])]
            return esp, dif

        def conf_h(L):
            dif = []
            for nv, v in L.items():
                if num(v['F']) and num(v['G']):
                    if not perto(v['H'], 1.65 * v['F'] + abs(v['G']), TOL):
                        dif.append((nv, v['F'], v['G'], v['H'], 1.65 * v['F'] + abs(v['G'])))
                elif v['H'] not in ('', None):
                    dif.append((nv, 'H deveria ser vazio', v['F'], v['G'], v['H']))
            return dif

        def conf_l(L, g_fixo=None):
            dif = []
            for nv, v in L.items():
                cv = (esp_cv.get((U(an), nv)) or {}).get('cv')
                g = v['G'] if g_fixo is None else g_fixo
                sg = None
                if v['K'] == '-':
                    ok = v['L'] == '-'
                elif num(g) and num(v['K']) and cv:
                    sg = (v['K'] - abs(g)) / cv
                    if PISO0[0]:
                        sg = max(0.0, sg)                  # ADR-101: Sigma negativo = 0
                    ok = perto(v['L'], sg, TOL_SIGMA) and v['M'] == classe_sigma(v['L'] if num(v['L']) else sg,
                                                                              PISO0[0])
                else:
                    sg, ok = None, v['L'] in ('', None)
                if not ok:
                    dif.append({'nivel': nv, 'K': v['K'], 'G': g, 'CVp': cv, 'L': v['L'], 'L_esperado': sg,
                                'M': v['M']})
            return dif

        def conf_painel(L, pn):
            dif = []
            for k in range(1, nlv + 1):
                E_, G_, H_, I_ = pn[k - 1][4], pn[k - 1][6], pn[k - 1][7], pn[k - 1][8]
                v = L.get(str(k))
                if v is None:
                    if G_ not in ('', None) or I_ not in ('', None):
                        dif.append((k, 'nível sem linha na Estatística', G_, I_))
                    continue
                if not igual_valor(G_, v['G']):
                    dif.append((k, 'G', G_, v['G']))
                if num(E_) and num(G_):
                    if not perto(H_, 1.65 * E_ + abs(G_), TOL):
                        dif.append((k, 'H', H_, 1.65 * E_ + abs(G_)))
                elif H_ not in ('', None):
                    dif.append((k, 'H deveria ser vazio', E_, G_, H_))
                if not igual_valor(I_, v['L']):
                    dif.append((k, 'I', I_, v['L']))
            return dif

        def k5_ok(rows_base, rod, texto):
            partes, (tipo, *_r) = k5_esperado(Oraculo(rows_base), prov_k5, ano_val, rod)
            return k5_confere(partes, tipo, texto), {'K5': texto[:160], 'esperado': partes, 'tipo': tipo}

        rows_est = {r for r, _ in linhas_an}

        def outros(est_a, est_b):
            # a coluna U ("ordem critico") e um contador acumulado das linhas criticas: muda nas linhas de baixo quando o
            # analito do teste muda de margem -- nao e "outro analito mudou"
            return [(14 + i, letra(j), a, b) for i, (r0, r1) in enumerate(zip(est_a, est_b)) if 14 + i not in rows_est
                    for j, (a, b) in enumerate(zip(r0, r1)) if a != b and letra(j) != 'U']

        # ============================================================ FOTO
        foto0 = foto()
        estado['foto0'] = foto0
        L0, pn0 = vals_an(), painel()
        base0 = foto0['base']
        G0 = {nv: v['G'] for nv, v in L0.items()}
        prep = {'analito': an, 'nome_CAP': nome_cap, 'canonico_mapa': canon, 'provedor_linha': pv_an,
                'eqProvedor_K5': prov_k5, 'ano_EP': ano_val, 'rodada_P4': rod_val, 'usuario': usuario,
                'painel_trocado': estado['b3'] is not None, 'candidatos': len(cand), 'G_antes': G0,
                'P4_fixado_em_TODAS': estado['p4_original'] if estado['p4_forcado'] else None,
                'equipamento_filtro': eqf, 'exclusoes': len(exc), 'janela_sigma': janela}
        print('   preparo ' + json.dumps(prep, ensure_ascii=False, default=str)[:600], flush=True)

        # ============================================================ F01 incluir pela janela
        a0 = aud0()
        ins = []
        for am, res, alvo in AMOSTRAS:
            cf, r_, lin = salvar(campos('CAP', ROD, ano, nome_cap, am, res, alvo))
            ins.append({'amostra': am, 'confirmacao': cf, 'retorno': r_[:160], 'linha': lin})
        recalc(q)
        wsc, c0 = LO['CAP'].Parent, LO['CAP'].Range.Column
        ok1 = len({x['linha'] for x in ins}) == len(AMOSTRAS)
        for x, (am, res, alvo) in zip(ins, AMOSTRAS):
            lin = x['linha']
            if not lin:
                ok1 = False
                continue
            F_, G_ = wsc.Cells(lin, c0 + 5).Value, wsc.Cells(lin, c0 + 6).Value
            co, cp = wsc.Cells(lin, c0 + 14), wsc.Cells(lin, c0 + 15)
            b_ = bias_pct(res, alvo)
            x.update({'F': F_, 'G': G_, 'O': co.Value, 'P': cp.Value, 'O_formula': bool(co.HasFormula),
                      'P_formula': bool(cp.HasFormula), 'bias_esperado': b_})
            ok1 = (ok1 and x['confirmacao'] == '' and aviso_ok(x['retorno']) == '' and perto(F_, res) and perto(G_, alvo)
                   and x['O_formula'] and x['P_formula'] and perto(co.Value, b_, TOL)
                   and num(F_) and num(G_) and perto(co.Value, (F_ - G_) / G_ * 100, TOL)
                   and perto(cp.Value, abs(co.Value), TOL))
        reg('F01 Inclusão pela janela (ANALISTA): ConfirmacoesCEQ vazia e SalvarCEQ "OK|linha|" sem aviso nas 3 amostras; '
            'O e P da linha são fórmula e valem (F−G)/G×100 e |O| (1e-9)', ok1, dict(prep, amostras=ins))

        # ============================================================ F02 EQA_Base + trilha
        novas = [linha_base('CAP', ano, ROD, nome_cap, canon, am, res, alvo, 'Acceptable', 'ACEITO')
                 for am, res, alvo in AMOSTRAS]
        base1 = ler_base(q)
        por_ch = {U(r[20]): r for r in base1}
        dif2 = {nv[5]: conf_linha_base(por_ch.get(U(nv[20])), nv) for nv in novas}
        dif2 = {k: v for k, v in dif2.items() if v}
        ch_novas = {U(nv[20]) for nv in novas}
        resto = collections.Counter(chave_rep(r) for r in cheias(base1) if U(r[20]) not in ch_novas)
        resto_ok = resto == collections.Counter(chave_rep(r) for r in cheias(base0))
        rod_lista = f'{ano}|{ROD}'
        lst1 = lista('lstRodadasCAP') or []
        reg_ = acoes(a0, 'CEQ_REGISTRADO')
        cons = acoes(a0, 'EQA_BASE_CONSOLIDADA')
        trilha_ok = (len(reg_) == len(AMOSTRAS) and len(cons) == len(AMOSTRAS)
                     and all(any(am in d and ROD in d for d in reg_) for am, _, _ in AMOSTRAS))
        reg('F02 EQA_Base consolidada na hora: cada chave nova com o canônico do mapa, ACEITO, Uso SIM e o bias; o resto '
            'da base intacto; a rodada entra na lstRodadasCAP; Audit_Log com 3 CEQ_REGISTRADO e 3 EQA_BASE_CONSOLIDADA',
            not dif2 and resto_ok and len(cheias(base1)) == len(cheias(base0)) + len(AMOSTRAS)
            and U(rod_lista) in [U(x) for x in lst1] and trilha_ok,
            {'divergencias': dif2, 'resto_intacto': resto_ok, 'linhas_base': [len(cheias(base0)), len(cheias(base1))],
             'lstRodadasCAP_tem': U(rod_lista) in [U(x) for x in lst1], 'CEQ_REGISTRADO': [d[:160] for d in reg_],
             'EQA_BASE_CONSOLIDADA': len(cons)})

        # ============================================================ F03 G/T/AC/AD/K5 pelo oraculo
        L1, pn1 = vals_an(), painel()
        est1 = [list(v) for _, v in linhas_tabela(e, ult)]
        esp0, dif0 = conf_g(base0, L0)
        esp1, dif1 = conf_g(base0 + novas, L1)
        k5a, ev_k5a = k5_ok(base0, rod_val, foto0['K5'])
        k5b, ev_k5b = k5_ok(base0 + novas, rod_val, s(e.Range('K5').Text))
        out1 = outros(foto0['est'], est1)

        def conta(x):
            return int(x) if num(x) else 0
        # as 3 amostras ENTRARAM no recorte: sem isto, um G parado (rodada fora do N4/P4, consolidacao que nao chegou)
        # passaria, porque o oraculo tambem nao as contaria
        entrou = (num(esp1['AC']) and conta(esp1['AC']) == conta(esp0['AC']) + len(AMOSTRAS)
                  and num(esp1['AD']) and conta(esp1['AD']) == conta(esp0['AD']) + 1)
        reg('F03 G, T, AC e AD do analito = oráculo independente (qa_rodadas_cv) sobre a EQA_Base de antes + as 3 linhas '
            'novas (e a FOTO = oráculo da base de antes); as 3 amostras entram (AC +3, AD +1); K5 conta pela mesma '
            'regra; os outros analitos não mudam',
            not dif0 and not dif1 and entrou and k5a and k5b and not out1,
            {'esperado_antes': esp0, 'esperado_depois': esp1, 'entraram': entrou, 'divergencias_antes': dif0[:6],
             'divergencias_depois': dif1[:6], 'lido_depois': {nv: {c: v[c] for c in ('G', 'T', 'AC', 'AD')}
                                                             for nv, v in L1.items()},
             'K5_antes': ev_k5a, 'K5_depois': ev_k5b, 'outros_analitos_mudaram': out1[:6]})

        # ============================================================ F04 isolamento (valores fechados)
        bs = [bias_pct(res, alvo) for _, res, alvo in AMOSTRAS]
        fech = {'G': sum(abs(x) for x in bs) / len(bs), 'T': sum(bs) / len(bs), 'AC': len(bs), 'AD': 1}
        rod_iso = f'{ano}|{ROD}'
        definir_nome('eqRodada', rod_iso)
        Li = vals_an()
        dif_f = [(nv, c, v[c], fech[c]) for nv, v in Li.items() for c in fech if not igual_valor(v[c], fech[c])]
        esp_i, dif_i = conf_g(base0 + novas, Li, rod=rod_iso)
        dif_hl = conf_h(Li) + conf_l(Li, g_fixo=fech['G'])
        orc1 = Oraculo(base0 + novas)
        oth = []
        for r, v in linhas_tabela(e, ult):
            an_ = txt(v[0])
            if r in rows_est or not an_ or an_ == '0' or '|' not in txt(v[ci('AB')]):
                continue
            esp_o = orc1.bias(an_, ano_val, prov_de(an_), rod_iso, 'ABS')
            if not igual_valor(v[ci('G')], esp_o):
                oth.append((r, an_, v[ci('G')], esp_o))
        k5i, ev_k5i = k5_ok(base0 + novas, rod_iso, s(e.Range('K5').Text))
        definir_nome('eqRodada', rod_val)
        Lr = vals_an()
        volta4 = all(igual_valor(Lr[nv][c], L1[nv][c]) for nv in L1 for c in COLS)
        reg('F04 Isolamento: P4 = "<ano>|QA-FLUXO 1" dá os valores fechados das amostras digitadas (G = média de |b|, '
            'T = média de b, AC = 3, AD = 1) e o oráculo concorda; H e L com esse G; os outros analitos = oráculo (sem '
            'dado nessa rodada); K5 do conjunto; P4 volta e o analito volta',
            not dif_f and not dif_i and not dif_hl and not oth and k5i and volta4,
            {'P4': rod_iso, 'fechado': fech, 'oraculo': esp_i, 'lido': {nv: {c: v[c] for c in ('G', 'T', 'AC', 'AD', 'H', 'L')}
                                                                       for nv, v in Li.items()},
             'divergencias_fechado': dif_f[:6], 'divergencias_oraculo': dif_i[:6], 'divergencias_H_L': dif_hl[:4],
             'outros_analitos_divergentes': oth[:6], 'K5': ev_k5i, 'P4_restaurado': rod_val, 'voltou': volta4})

        # ============================================================ F05 ET
        h0, h1 = conf_h(L0), conf_h(L1)
        reg('F05 Erro Total: Estatística H = 1,65 × F + |G| em cada nível, na FOTO e depois da inclusão (1e-9); depois '
            'da inclusão H é número em todos os níveis',
            not h0 and not h1 and all(num(v['H']) for v in L1.values()),
            {'antes': {nv: (v['F'], v['G'], v['H']) for nv, v in L0.items()},
             'depois': {nv: (v['F'], v['G'], v['H']) for nv, v in L1.items()}, 'divergencias': (h0 + h1)[:6]})

        # ============================================================ F06 Sigma
        l0, l1 = conf_l(L0), conf_l(L1)
        cvs = {nv: (esp_cv.get((U(an), nv)) or {}).get('cv') for nv in L1}
        reg('F06 Sigma: Estatística L = (K − |G|)/CVp, com o CV pooled recalculado por fora (qa_incerteza.esperado_ciq) '
            'na janela Sigma_Ini..MU_Fim (Sigma_Ini = EDATE(MU_Fim; −meses)+1); classe M pela escada; na FOTO e depois',
            janela_ok and not l0 and not l1 and all(num(v['L']) for v in L1.values()),
            {'janela': janela, 'janela_ok': janela_ok, 'CVp': cvs,
             'antes': {nv: (v['K'], v['G'], v['L'], v['M']) for nv, v in L0.items()},
             'depois': {nv: (v['K'], v['G'], v['L'], v['M']) for nv, v in L1.items()}, 'divergencias': (l0 + l1)[:6]})

        # ============================================================ F07 Painel
        p0, p1 = conf_painel(L0, pn0), conf_painel(L1, pn1)
        reg('F07 Painel do analito: G = Estatística G; H = 1,65 × E + |G| (vazio sem E); I = Estatística L; na FOTO e '
            'depois da inclusão', not p0 and not p1,
            {'antes': [r[4:9] for r in pn0], 'depois': [r[4:9] for r in pn1], 'divergencias': (p0 + p1)[:6]})

        # ============================================================ F08 corrigir
        am1, res1, alvo1 = AMOSTRAS[0]
        lin1 = ins[0]['linha']
        lst = run('mCEQForm.RegistrosDaRodada', 'CAP', ROD, str(ano)) or ()
        lin_lista = {txt(x[1]): int(float(x[4])) for x in lst if len(x) > 4}
        c8 = campos('CAP', ROD, ano, nome_cap, am1, CORRECAO, alvo1)
        cf8 = s(run('mCEQForm.ConfirmacoesCEQ', c8, lin1)) if lin1 else ''
        a8 = aud0()
        r8 = s(run('mCEQForm.CorrigirCEQ', lin1, c8)) if lin1 else 'sem linha'
        L2, pn2 = vals_an(), painel()
        novas2 = [linha_base('CAP', ano, ROD, nome_cap, canon, am1, CORRECAO, alvo1, 'Acceptable', 'ACEITO')] + novas[1:]
        esp2, dif8 = conf_g(base0 + novas2, L2)
        lb2 = conf_linha_base({U(r[20]): r for r in ler_base(q)}.get(U(novas2[0][20])), novas2[0])
        alt = acoes(a8, 'CEQ_ALTERADO')
        cons8 = acoes(a8, 'EQA_BASE_CONSOLIDADA')
        dif8b = conf_h(L2) + conf_l(L2) + conf_painel(L2, pn2)
        mudou = any(not igual_valor(L2[nv]['G'], L1[nv]['G']) for nv in L1)
        reg('F08 Corrigir selecionado (CorrigirCEQ) da amostra 1: a mesma linha da lista "Registros desta rodada", '
            'confirmação vazia, EQA_Base com o valor novo, G/T/AC/AD = oráculo, H, L e Painel fecham, G muda; '
            f'CEQ_ALTERADO "{res1} {SETA} {CORRECAO}" e 1 consolidação',
            bool(lin1) and lin_lista.get(am1) == lin1 and cf8 == '' and r8.startswith(f'OK|{lin1}|')
            and aviso_ok(r8) == '' and perto(wsc.Cells(lin1, c0 + 5).Value, CORRECAO) and not lb2 and not dif8
            and not dif8b and mudou and len(alt) == 1 and f'{res1} {SETA} {CORRECAO}' in alt[0] and len(cons8) == 1,
            {'linha': lin1, 'linha_da_lista': lin_lista.get(am1), 'confirmacao': cf8, 'retorno': r8[:160],
             'base': lb2, 'esperado': esp2, 'divergencias': dif8[:6], 'H_L_Painel': dif8b[:6], 'G_mudou': mudou,
             'G': {nv: (L1[nv]['G'], L2[nv]['G']) for nv in L1}, 'CEQ_ALTERADO': [d[:300] for d in alt],
             'consolidacoes': len(cons8)})

        # ============================================================ F09 excluir -> FOTO
        a9 = aud0()
        n_lr = LO['CAP'].ListRows.Count
        exc9 = {f'CAP|{x["linha"]}': excluir('CAP', x['linha'])[:120] for x in ins if x['linha']}
        f9 = foto()
        d9 = dif_foto(foto0, f9)
        excl = acoes(a9, 'CEQ_EXCLUIDO')
        reg('F09 Excluir selecionado (ExcluirCEQ) das 3 amostras: Estatística, K5, Painel, EQA_Base, tabelas (linhas com '
            'analito) e listas voltam EXATAMENTE à FOTO; 3 CEQ_EXCLUIDO (ListRows.Count não é comparado)',
            all(v.startswith('OK|') for v in exc9.values()) and len(exc9) == len(AMOSTRAS) and not d9
            and len(excl) == len(AMOSTRAS),
            {'exclusoes': exc9, 'diferencas_da_foto': d9, 'CEQ_EXCLUIDO': len(excl),
             'ListRows_CAP_info': n_lr, 'G_depois': {nv: f9['est'][r - 14][ci('G')] for r, nv in linhas_an}})

        # ============================================================ F10 avaliacao nao reconhecida
        st10 = txt(run('mEQA.PadronizarStatus', TERMO_F10))
        rod10, am10 = 'QA-FLUXO A10', 'QAFC-A10'
        g_se, res10 = None, None
        for d_ in DESVIOS_F10:
            teste_ = linha_base('CAP', ano, rod10, nome_cap, canon, am10, 100 + d_, 100, TERMO_F10, 'ACEITO')
            g_se = Oraculo(base0 + [teste_]).bias(an, ano_val, pv_an, rod_val, 'ABS')
            res10 = 100 + d_
            if not all(igual_valor(g_se, g) for g in G0.values()):
                break
        cf10, r10, lin10 = salvar(campos('CAP', rod10, ano, nome_cap, am10, res10, 100, aval=TERMO_F10))
        L10 = vals_an()
        row10 = linha_base('CAP', ano, rod10, nome_cap, canon, am10, res10, 100, TERMO_F10, st10)
        lb10 = conf_linha_base({U(r[20]): r for r in ler_base(q)}.get(U(row10[20])), row10)
        esp10, dif10 = conf_g(base0 + [row10], L10)
        x10 = excluir('CAP', lin10) if lin10 else 'sem linha'
        L10x = vals_an()
        volta10 = L10x == L0
        g_igual = all(igual_valor(L10[nv]['G'], G0[nv]) for nv in G0)
        av10 = avisou(cf10, aviso_ok(r10))
        ev10 = {'termo': TERMO_F10, 'status_padronizado': st10, 'confirmacao': cf10[:200], 'retorno': r10[:160],
                'base': lb10, 'G_antes': G0, 'G_depois': {nv: v['G'] for nv, v in L10.items()}, 'oraculo': esp10,
                'G_se_fosse_reconhecida': g_se, 'divergencias': dif10[:4], 'avisos': av10, 'exclusao': x10[:120],
                'voltou_a_foto': volta10}
        calc10 = bool(lin10) and not lb10 and not dif10 and x10.startswith('OK|') and volta10
        mudaria10 = not all(igual_valor(g_se, g) for g in G0.values())
        t10 = ('F10 Avaliação não reconhecida ("%s") vira %s na EQA_Base e fica fora do bias, do ET e do Sigma'
               % (TERMO_F10, st10 or '?'))
        if not calc10:
            reg(t10 + ': cálculo/limpeza conferem com o oráculo', False, ev10)
        elif U(st10) != 'NAO AVALIADO':
            # o termo passou a ser reconhecido: o oraculo (dif10 vazio) ja conferiu que a amostra entra no bias
            reg(t10 + ': termo agora reconhecido pelo PadronizarStatus, a amostra entra no bias pelo oráculo', True, ev10)
        elif not mudaria10:
            reg(f'F10 Pré-condição ausente: nenhum desvio {DESVIOS_F10} mudaria o G se a avaliação fosse reconhecida '
                '(o caso não distingue "fora do bias" de "dentro")', False, ev10)
        elif av10:
            reg(t10 + ': o módulo avisa antes/depois de gravar', True, ev10)
        elif g_igual:
            alerta(t10 + ' SEM AVISO (ConfirmacoesCEQ vazia, OK sem aviso): o G não muda, mas mudaria se a avaliação '
                   'fosse reconhecida (risco alto 2 do mapa)', ev10)
        else:
            reg(t10 + ': o G mudou com a amostra NAO AVALIADO', False, ev10)

        # ============================================================ F11 ano diferente do Ano EP
        ano11, rod11, am11, res11 = ano + 1, 'QA-FLUXO A11', 'QAFC-A11', 112
        cf11, r11, lin11 = salvar(campos('CAP', rod11, ano11, nome_cap, am11, res11, 100))
        L11 = vals_an()
        row11 = linha_base('CAP', ano11, rod11, nome_cap, canon, am11, res11, 100, 'Acceptable', 'ACEITO')
        lb11 = conf_linha_base({U(r[20]): r for r in ler_base(q)}.get(U(row11[20])), row11)
        esp11, dif11 = conf_g(base0 + [row11], L11)
        g_igual11 = all(igual_valor(L11[nv]['G'], G0[nv]) for nv in G0)
        definir_nome('eqAnoEP', 'TODOS')
        definir_nome('eqRodada', 'TODAS')
        Lt = vals_an()
        esp_t, dif_t = conf_g(base0 + [row11], Lt, rod='TODAS', ano_='TODOS')
        g_t_sem = Oraculo(base0).bias(an, 'TODOS', pv_an, 'TODAS', 'ABS')
        entra = not igual_valor(esp_t['G'], g_t_sem)
        definir_nome('eqAnoEP', estado['nomes']['eqAnoEP'])
        definir_nome('eqRodada', estado['nomes']['eqRodada'])
        Lr11 = vals_an()
        volta_n = all(igual_valor(Lr11[nv][c], L11[nv][c]) for nv in L11 for c in COLS)
        x11 = excluir('CAP', lin11) if lin11 else 'sem linha'
        volta11 = vals_an() == L0
        av11 = avisou(cf11, aviso_ok(r11))
        ev11 = {'ano_EP': ano_val, 'ano_digitado': ano11, 'confirmacao': cf11[:200], 'retorno': r11[:160],
                'base': lb11, 'G_antes': G0, 'G_com_Ano_EP': {nv: v['G'] for nv, v in L11.items()}, 'oraculo': esp11,
                'divergencias': dif11[:4], 'G_com_TODOS': {nv: v['G'] for nv, v in Lt.items()}, 'oraculo_TODOS': esp_t,
                'G_TODOS_sem_a_linha': g_t_sem, 'entra_com_TODOS': entra, 'divergencias_TODOS': dif_t[:4],
                'N4_P4_restaurados': volta_n, 'avisos': av11, 'exclusao': x11[:120], 'voltou_a_foto': volta11}
        calc11 = (bool(lin11) and not lb11 and not dif11 and not dif_t and entra and volta_n and x11.startswith('OK|')
                  and volta11)
        t11 = (f'F11 Rodada com ano {ano11} ≠ Ano EP {ano}: o G não muda (ADR-072, ano vigente exato); com Ano EP = TODOS '
               'e P4 = TODAS ela entra')
        if not calc11:
            reg(t11 + ': cálculo/limpeza conferem com o oráculo', False, ev11)
        elif av11:
            reg(t11 + ': o módulo avisa (Ano EP) ao gravar', True, ev11)
        elif g_igual11:
            alerta(t11 + ' -- a janela diz "Gravado" SEM AVISO e o usuário vê "inseri e nada mudou" (risco alto 1 do '
                   'mapa)', ev11)
        else:
            # com dif11 vazio o oraculo diria o mesmo G de antes; chegar aqui = a FOTO ja divergia do oraculo (F03)
            reg(t11 + ': o G mudou com o Ano EP da cópia (contradiz o ano vigente exato do ADR-072)', False, ev11)

        # ============================================================ F12 Controllab num analito do CAP (OC-0069)
        if LO['Controllab'] is None:
            raise PreCondicao('tabela tblEQA_Controllab_Dados não encontrada (o F12 registra pelo Controllab)',
                              codigo='F12')
        mapa_ctl = run('mEQA.MapaDoProvedor', 'Controllab') or ()
        nome_ctl = next((txt(m[0]) for m in mapa_ctl if U(m[1]) == U(an)), None)
        mapa_real = nome_ctl is not None
        if not mapa_real:
            wy = base_ws.Range('W2:Y201').Value
            livre = next((i for i, l_ in enumerate(wy, 2) if not any(txt(x) for x in l_)), None)
            if livre is None:
                raise PreCondicao('mapa W2:Y201 cheio: sem linha para o nome Controllab de teste do F12', codigo='F12')
            base_ws.Unprotect(ii.SENHA)
            xl.EnableEvents = False
            try:
                base_ws.Range(f'W{livre}:Y{livre}').Value = (('Controllab', NOME_CTL_TMP, canon),)
            finally:
                xl.EnableEvents = True
            mapa_tmp.append(livre)
            nome_ctl = NOME_CTL_TMP
        aviso_an = s(run('mCEQForm.AvisoAnalitoCEQ', 'Controllab', nome_ctl))
        rod12, am12, res12 = 'QA-FLUXO A12', 'QAFC-A12', 130
        cf12, r12, lin12 = salvar(campos('Controllab', rod12, ano, nome_ctl, am12, res12, 100))
        L12 = vals_an()
        row12 = linha_base('Controllab', ano, rod12, nome_ctl, canon, am12, res12, 100, 'Acceptable', 'ACEITO')
        lb12 = conf_linha_base({U(r[20]): r for r in ler_base(q)}.get(U(row12[20])), row12)
        esp12, dif12 = conf_g(base0 + [row12], L12)
        g_se_ctl = Oraculo(base0 + [row12]).bias(an, ano_val, 'Controllab', rod_val, 'ABS')
        g_igual12 = all(igual_valor(L12[nv]['G'], G0[nv]) for nv in G0)
        x12 = excluir('Controllab', lin12) if lin12 else 'sem linha'
        if mapa_tmp:
            xl.EnableEvents = False
            try:
                for r in list(mapa_tmp):
                    base_ws.Range(f'W{r}:Y{r}').ClearContents()
                    mapa_tmp.remove(r)
            finally:
                xl.EnableEvents = True
        volta12 = vals_an() == L0
        av12 = avisou(cf12, aviso_ok(r12))
        ev12 = {'provedor_da_linha': pv_an, 'nome_Controllab': nome_ctl, 'mapa_real': mapa_real,
                'aviso_do_analito': aviso_an[:160], 'confirmacao': cf12[:200], 'retorno': r12[:160], 'base': lb12,
                'G_antes': G0, 'G_depois': {nv: v['G'] for nv, v in L12.items()}, 'oraculo': esp12,
                'G_se_o_provedor_da_linha_fosse_Controllab': g_se_ctl, 'divergencias': dif12[:4], 'avisos': av12,
                'exclusao': x12[:120], 'voltou_a_foto': volta12}
        calc12 = bool(lin12) and not lb12 and not dif12 and x12.startswith('OK|') and volta12
        t12 = (f'F12 Controllab mapeado para {an}, cujo provedor do bias é {pv_an} (OC-0069): a amostra não entra no G')
        if not calc12:
            reg(t12 + ': cálculo/limpeza conferem com o oráculo', False, ev12)
        elif av12:
            reg(t12 + ' e o módulo avisa ao gravar', True, ev12)
        elif g_igual12:
            alerta(t12 + ' SEM AVISO (a janela mostra o canônico como se entrasse no bias)', ev12)
        else:
            reg(t12 + ': o G mudou', False, ev12)

        # ============================================================ F13 limpeza final
        caso_f13(q, estado, limpar, foto, do_teste, base_ws, mapa_tmp)
    except PreCondicao as pc:
        resto_ = ' (F01..F13 não executados)' if pc.codigo == 'F01' else ' (os casos seguintes não foram executados)'
        reg(f'{pc.codigo} Pré-condição ausente: {pc.motivo}{resto_}', False, dict(pc.evid, pre_condicao=pc.motivo))
    except Exception as ex_:                          # noqa: BLE001
        reg('F99 Execução da suíte sem exceção', False, {'erro': repr(ex_), 'trace': traceback.format_exc()[-2000:]})
    finally:
        try:
            if estado['foto0'] is not None and not estado['f13']:
                caso_f13(q, estado, limpar, foto, do_teste, base_ws, mapa_tmp)
        except Exception as ex_:                      # noqa: BLE001
            reg('F13 Limpeza: nada do teste fica na cópia e tudo volta à FOTO', False,
                {'erro': repr(ex_), 'trace': traceback.format_exc()[-1500:], 'ainda_incluidos': incluidos})
        try:
            if estado['p4_forcado']:                  # o P4 original (so a tela; a copia nunca e salva)
                q.ex.xl.EnableEvents = False
                try:
                    q.nome('eqRodada').Formula = estado['p4_original']
                finally:
                    q.ex.xl.EnableEvents = True
        except Exception:                             # noqa: BLE001
            pass
        try:
            if estado['b3'] is not None:
                p.Unprotect(ii.SENHA)
                q.ex.xl.EnableEvents = False
                try:
                    p.Range('B3').Value = estado['b3']
                finally:
                    q.ex.xl.EnableEvents = True
                q.run('mEstatistica.PainelMudou', teto=600)
        except Exception:                             # noqa: BLE001 -- so a tela; a copia nunca e salva
            pass
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('F00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    alertas = [r for r in RES if r['resultado'] == 'ALERTA']
    n_pass = sum(1 for r in RES if r['resultado'] == 'PASS')
    print(f'\n=== {produto} fluxo do CEQ pela janela (bias, ET, Sigma): {n_pass} PASS / {len(falhas)} FAIL / '
          f'{len(alertas)} ALERTA ===', flush=True)
    return falhas


def caso_f13(q, estado, limpar, foto, do_teste, base_ws, mapa_tmp):
    estado['f13'] = True
    ev = limpar()
    f13 = foto()
    d13 = dif_foto(estado['foto0'], f13)
    sobra = do_teste()
    sobra_base = [txt(r[20]) for r in f13['base'] if txt(r[18]) == ARQ]
    wy = base_ws.Range('W2:Y201').Value
    mapa_sobra = [i for i, l_ in enumerate(wy, 2) if any(txt(x) == NOME_CTL_TMP for x in l_)]
    nomes_ok = {n: (q.nome(n).Value, v0) for n, v0 in estado['nomes'].items() if q.nome(n).Value != v0}
    reg('F13 Limpeza: nenhum registro (tabelas, EQA_Base) nem linha de mapa do teste fica na cópia; N4/P4 como '
        'estavam; Estatística, K5, Painel, EQA_Base, tabelas e listas EXATAMENTE iguais à FOTO',
        not d13 and not sobra and not sobra_base and not mapa_sobra and not mapa_tmp and not nomes_ok,
        {'limpeza': ev, 'diferencas_da_foto': d13, 'registros_que_sobraram': sobra, 'base_que_sobrou': sobra_base,
         'mapa_que_sobrou': mapa_sobra, 'nomes_diferentes': nomes_ok})


if __name__ == '__main__':
    t0 = time.time()
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    print(f'tempo {time.time() - t0:.0f}s')
    sys.exit(1 if falhas else 0)
