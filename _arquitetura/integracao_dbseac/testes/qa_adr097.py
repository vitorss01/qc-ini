# -*- coding: utf-8 -*-
"""QA do ADR-097 (16/11/2026, reteste 06): carimbo da linha manual apagada ou movida (OC-0078), DETALHE do ID fora do
padrao (OC-0080), A01 e mensagem do ATUALIZAR para o manual repetido por chave (OC-0081), filtros do Registros com a
lista vazia e o lote digitado (OC-0082/0083) e o aviso do Ano De/Ate da Estatistica da Hema (OC-0048). Roda numa COPIA
instalada; nada e salvo.

Uso: python qa_adr097.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  C01 apagar ID..MOTIVO de uma linha manual (eventos ligados): USUARIO e REGISTRADO_EM ficam vazios
  C02 digitar outro manual nessa linha: carimbo NOVO (nao o de quem apagou)
  C03 corrigir o RESULTADO de um manual ja carimbado: o carimbo NAO muda
  C04 apagar so o ID de uma linha com dado: ID novo e o carimbo fica
  C05 carimbo orfao numa linha vazia (arquivo antigo): o PrepararEntradas (passada do ATUALIZAR) o limpa; nenhuma
      linha com dado muda de USUARIO/REGISTRADO_EM
  C06 colar ID..MOTIVO de outra linha sobre uma linha com carimbo orfao: ID novo e carimbo novo
  C07 digitar numa linha com carimbo orfao (origem de recorte sem evento): carimbo novo
  C08 recortar ID..MOTIVO para outra linha e digitar outro manual na origem: carimbo novo na origem; o ID movido fica
  C09 apagar DATA..MOTIVO deixando o ID (revisao): o ID e o carimbo saem; o manual digitado depois tem ID e carimbo novos
  C10 preencher para baixo (FillDown) a partir de uma linha carimbada: a linha de origem MANTEM ID e carimbo; as de baixo
      ganham ID novo
  C11 apagar ID..MOTIVO de um manual com SUBSTITUI_ID: o SUBSTITUI_ID que sobrou na linha sai junto
  C12 linha so com o SUBSTITUI_ID (ainda sem dado): digitar o EQUIPAMENTO nao apaga o SUBSTITUI_ID (fluxo do ADR-082)
  M01 depois do ATUALIZAR: os manuais digitados nas linhas apagada/orfa NAO sao recusados por "DATA/HORA posterior ao
      lancamento" e o USUARIO_MANUAL nao e o do carimbo antigo
  Q01 dois manuais com a mesma data, hora, nivel e analito: o 2o continua CONFLITO_MANUAL "resultado manual repetido"
      (regra do motor intacta); o A01 dele manda corrigir a HORA ou apagar a linha e NAO inativar o 1o; a mensagem do
      ATUALIZAR conta a linha
  D01 DETALHE: a formula = camada_dados.FORMULA_DETALHE; ID fora do padrao -> "ID fora do padrão: apague-o (um novo é
      gerado)"; MAN_ valido ainda nao atualizado (inclusive " man_7 ") -> "atualize os dados"
  F01 filtros do Registros (Lote, Analito, Equipamento, Status): lista nunca vazia; texto fora da lista recusado; o
      item "(nenhum registrado)", escolhido, nao filtra nada
  F02 filtro de Lote: o lote digitado como numero e aceito (celula em texto)
  E01 (Hema) Estatistica: D3 recusa Ano Ate < Ano De; A5 avisa quando o lote em analise nao tem resultado no
      intervalo e fica vazio quando o intervalo cobre o lote; B3 acima do D3 tambem e avisado; sem resultado do lote
      em analise o teste REPROVA (nao exercitado)
  Y00 seguranca do teste: arquivo de entrada intacto
"""
import datetime as dt
import json
import os
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import instalar_integracao as ii  # noqa: E402
import camada_dados as cd  # noqa: E402

TB = 'tblResultados_Manuais'
VELHO = dt.datetime(2020, 1, 1, 6, 0, 0)
TXT_INVALIDO = 'ID fora do padrão: apague-o (um novo é gerado)'
ITEM_VAZIO = '(nenhum registrado)'


def txt(v):
    return '' if v is None else str(v).strip()


def norm_lote(v):
    s = txt(v)
    return s[:-2] if s.endswith('.0') and s[:-2].isdigit() else s


def ingenuo(v):
    return v.replace(tzinfo=None) if hasattr(v, 'tzinfo') and getattr(v, 'tzinfo', None) else v


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    bio = produto.startswith('Bio')
    q = QA(produto, caminho)
    try:
        xl, wb = q.ex.xl, q.wb
        xl.EnableEvents = False
        lo = q.lo(TB)
        ws = lo.Parent
        ws.Unprotect(ii.SENHA)
        nomes = [c.Name for c in lo.ListColumns]
        c_id, c_mot = nomes.index('ID_REGISTRO') + 1, nomes.index('MOTIVO') + 1

        def cel(r, col):
            return lo.ListColumns(col).DataBodyRange.Cells(r, 1)

        def bloco(r):                                   # ID..MOTIVO da linha r (o "A:K" do usuario)
            return ws.Range(cel(r, 'ID_REGISTRO'), lo.DataBodyRange.Cells(r, c_mot))

        def carimbo(r):
            return txt(cel(r, 'USUARIO').Value), ingenuo(cel(r, 'REGISTRADO_EM').Value)

        def carimbar_velho(r, quem):                    # eventos DESLIGADOS: como um carimbo antigo/orfao
            xl.EnableEvents = False
            cel(r, 'USUARIO').Value = quem
            cel(r, 'REGISTRADO_EM').Value = VELHO
            return carimbo(r)                           # como o COM le de volta (com o fuso)

        def novo(c, quem):
            u, d = c
            return bool(u) and u != quem and hasattr(d, 'year') and d.year >= 2026

        # dados de um resultado real (o motor aceita): o mais recente ATIVO do interfaceamento
        fin = q.final_por_id()
        reais = [r for r in fin.values() if r.get('ORIGEM_RESULTADO') == 'INTERFACEAMENTO'
                 and r.get('STATUS_ANALITICO') == 'ATIVO' and hasattr(r.get('DATA_HORA'), 'year')]
        ref = max(reais, key=lambda r: r['DATA_HORA'])
        dia = dt.datetime.combine(ref['DATA_HORA'].date() - dt.timedelta(days=2), dt.time())
        base = {'EQUIPAMENTO': ref['EQUIPAMENTO'], 'MATRIZ': ref.get('MATRIZ'), 'LOTE': norm_lote(ref['LOTE']),
                'NIVEL': ref['NIVEL'], 'ANALITO': ref['ANALITO'], 'MOTIVO': cd.MOTIVOS_MANUAL[0]}

        def digitar(r, hora_min, valor, extra=None):     # eventos LIGADOS, celula a celula (como a mao do usuario)
            xl.EnableEvents = True
            vals = dict(base, DATA=dia, HORA=hora_min / 1440.0, RESULTADO=valor, **(extra or {}))
            for col in ('DATA', 'HORA', 'EQUIPAMENTO', 'MATRIZ', 'LOTE', 'NIVEL', 'ANALITO', 'RESULTADO', 'MOTIVO'):
                if vals.get(col) is not None:
                    cel(r, col).Value = vals[col]
            xl.EnableEvents = False
            return txt(cel(r, 'ID_REGISTRO').Value)

        vids = [txt(v[0]) for v in lo.ListColumns('ID_REGISTRO').DataBodyRange.Value]
        r0 = max([k for k, v in enumerate(vids, start=1) if v] or [0]) + 3
        if r0 + 20 > len(vids):
            lo.Resize(lo.Range.Resize(lo.Range.Rows.Count + 25))
        valor = round(float(ref['RESULTADO']), 3)
        ev = {}

        # ------------------------------------------------------------ C01 / C02
        rA = r0
        idA = digitar(rA, 3 * 60 + 1, valor)
        carimbar_velho(rA, 'QA97_ANTIGO')
        xl.EnableEvents = True
        bloco(rA).ClearContents()                       # a tecla Delete em ID..MOTIVO
        xl.EnableEvents = False
        c1 = carimbo(rA)
        reg('C01 Apagar ID..MOTIVO de uma linha manual: USUARIO e REGISTRADO_EM ficam vazios (OC-0078)',
            c1 == ('', None), {'linha': rA, 'id_apagado': idA, 'carimbo_depois': [c1[0], str(c1[1])]})
        idB = digitar(rA, 3 * 60 + 3, valor)
        c2 = carimbo(rA)
        reg('C02 Outro manual digitado na linha apagada: carimbo novo (quem digita, agora), não o de quem apagou',
            novo(c2, 'QA97_ANTIGO') and idB.startswith('MAN_') and idB != idA,
            {'id': idB, 'carimbo': [c2[0], str(c2[1])]})

        # ------------------------------------------------------------ C03 / C04
        rE = r0 + 1
        idE = digitar(rE, 3 * 60 + 5, valor)
        velhoE = carimbar_velho(rE, 'QA97_EDIT')
        xl.EnableEvents = True
        cel(rE, 'RESULTADO').Value = round(valor * 1.001, 3)
        xl.EnableEvents = False
        c3 = carimbo(rE)
        reg('C03 Corrigir o RESULTADO de um manual já carimbado não muda o carimbo (a edição continua sendo dele)',
            c3 == velhoE and txt(cel(rE, 'ID_REGISTRO').Value) == idE, {'carimbo': [c3[0], str(c3[1])]})
        xl.EnableEvents = True
        cel(rE, 'ID_REGISTRO').ClearContents()
        xl.EnableEvents = False
        idE2 = txt(cel(rE, 'ID_REGISTRO').Value)
        c4 = carimbo(rE)
        reg('C04 Apagar só o ID de uma linha com dado: ganha ID novo e mantém o carimbo',
            idE2.startswith('MAN_') and idE2 != idE and c4 == velhoE,
            {'id_antes': idE, 'id_depois': idE2, 'carimbo': [c4[0], str(c4[1])]})

        # ------------------------------------------------------------ C05
        rO = r0 + 2
        carimbar_velho(rO, 'QA97_ORFAO')
        n = lo.ListRows.Count
        us = lo.ListColumns('USUARIO').DataBodyRange.Value
        rg = lo.ListColumns('REGISTRADO_EM').DataBodyRange.Value
        ids = lo.ListColumns('ID_REGISTRO').DataBodyRange.Value
        antes = {k: (txt(us[k - 1][0]), str(ingenuo(rg[k - 1][0]))) for k in range(1, n + 1) if txt(ids[k - 1][0])}
        q.run('mIntegracao.PrepararEntradas')
        us = lo.ListColumns('USUARIO').DataBodyRange.Value
        rg = lo.ListColumns('REGISTRADO_EM').DataBodyRange.Value
        depois = {k: (txt(us[k - 1][0]), str(ingenuo(rg[k - 1][0]))) for k in antes}
        c5 = carimbo(rO)
        mudou = [k for k in antes if antes[k] != depois[k]]
        reg('C05 Carimbo órfão numa linha vazia (arquivo antigo): a passada do ATUALIZAR o limpa; nenhuma linha com dado '
            'muda de USUARIO/REGISTRADO_EM', c5 == ('', None) and not mudou,
            {'linhas_com_dado': len(antes), 'mudaram': mudou[:10], 'orfa_depois': [c5[0], str(c5[1])]})

        # ------------------------------------------------------------ C06
        rP = r0 + 3
        carimbar_velho(rP, 'QA97_ORFAO')
        xl.EnableEvents = True
        bloco(rA).Copy(bloco(rP))                       # colar ID..MOTIVO da linha rA (MAN_ repetido) na linha orfa
        xl.CutCopyMode = False
        xl.EnableEvents = False
        idP = txt(cel(rP, 'ID_REGISTRO').Value)
        c6 = carimbo(rP)
        reg('C06 Colar ID..MOTIVO de outra linha sobre uma linha com carimbo órfão: ID novo (ADR-096) e carimbo novo',
            idP.startswith('MAN_') and idP != idB and novo(c6, 'QA97_ORFAO'),
            {'id_origem': idB, 'id_colado': idP, 'carimbo': [c6[0], str(c6[1])]})

        # ------------------------------------------------------------ C07
        rQ = r0 + 4
        carimbar_velho(rQ, 'QA97_ORFAO')
        idQ = digitar(rQ, 3 * 60 + 9, valor)
        c7 = carimbo(rQ)
        reg('C07 Digitar numa linha com carimbo órfão (origem de um recorte, sem evento): carimbo novo',
            novo(c7, 'QA97_ORFAO') and idQ.startswith('MAN_'), {'id': idQ, 'carimbo': [c7[0], str(c7[1])]})

        # ------------------------------------------------------------ C08
        rM, rD = r0 + 5, r0 + 6
        idM = digitar(rM, 3 * 60 + 11, valor)
        carimbar_velho(rM, 'QA97_MOVE')
        xl.EnableEvents = True
        bloco(rM).Cut(bloco(rD))
        xl.EnableEvents = False
        orig_logo = carimbo(rM)
        idM2 = digitar(rM, 3 * 60 + 13, valor)
        c8 = carimbo(rM)
        reg('C08 Recortar ID..MOTIVO para outra linha e digitar outro manual na origem: carimbo novo na origem; o ID movido '
            'continua o mesmo', novo(c8, 'QA97_MOVE') and txt(cel(rD, 'ID_REGISTRO').Value) == idM and idM2 != idM,
            {'id_movido': idM, 'destino': txt(cel(rD, 'ID_REGISTRO').Value), 'origem_logo_apos_recorte':
             [orig_logo[0], str(orig_logo[1])], 'id_novo_na_origem': idM2, 'carimbo_origem': [c8[0], str(c8[1])]})

        # ------------------------------------------------------------ C09 (revisao: Delete sem a coluna do ID)
        rN = r0 + 10
        idN = digitar(rN, 3 * 60 + 15, valor)
        carimbar_velho(rN, 'QA97_ANTIGO')
        xl.EnableEvents = True
        ws.Range(cel(rN, 'DATA'), lo.DataBodyRange.Cells(rN, c_mot)).ClearContents()
        xl.EnableEvents = False
        c9a = (txt(cel(rN, 'ID_REGISTRO').Value),) + carimbo(rN)
        idN2 = digitar(rN, 3 * 60 + 17, valor)
        c9 = carimbo(rN)
        reg('C09 Apagar DATA..MOTIVO deixando o ID: o ID e o carimbo saem (um ID sem dado não é registro e não é '
            'reaproveitado); o manual digitado depois tem ID e carimbo novos',
            c9a == ('', '', None) and idN2.startswith('MAN_') and idN2 != idN and novo(c9, 'QA97_ANTIGO'),
            {'id_antes': idN, 'depois_do_delete': [str(x) for x in c9a], 'id_novo': idN2, 'carimbo': [c9[0], str(c9[1])]})

        # ------------------------------------------------------------ C10 (revisao: FillDown com a origem no Target)
        rF = r0 + 11
        idF = digitar(rF, 3 * 60 + 19, valor)
        velhoF = carimbar_velho(rF, 'QA97_FILL')
        xl.EnableEvents = True
        ws.Range(cel(rF, 'ID_REGISTRO'), lo.DataBodyRange.Cells(rF + 2, c_mot)).FillDown()
        xl.EnableEvents = False
        ids_f = [txt(cel(rF + k, 'ID_REGISTRO').Value) for k in range(3)]
        c10 = carimbo(rF)
        reg('C10 Preencher para baixo a partir de uma linha carimbada: a linha de origem mantém o ID e o carimbo (não muda '
            'de dono); as de baixo ganham ID novo',
            ids_f[0] == idF and c10 == velhoF and all(x.startswith('MAN_') and x != idF for x in ids_f[1:])
            and ids_f[1] != ids_f[2], {'ids': ids_f, 'carimbo_origem': [c10[0], str(c10[1])]})

        # ------------------------------------------------------------ C11 / C12 (revisao: SUBSTITUI_ID)
        if 'SUBSTITUI_ID' in nomes:
            rS = r0 + 14
            digitar(rS, 3 * 60 + 21, valor)
            xl.EnableEvents = False
            cel(rS, 'SUBSTITUI_ID').Value = 'QA97-SUBST'        # eventos desligados: nao procura o resultado
            carimbar_velho(rS, 'QA97_ANTIGO')
            xl.EnableEvents = True
            bloco(rS).ClearContents()                           # Delete de ID..MOTIVO (o SUBSTITUI_ID fica no fim)
            xl.EnableEvents = False
            c11 = [txt(cel(rS, c).Value) for c in ('ID_REGISTRO', 'SUBSTITUI_ID', 'USUARIO')]
            reg('C11 Apagar ID..MOTIVO de um manual com SUBSTITUI_ID: o SUBSTITUI_ID que sobrou sai junto (o próximo '
                'manual da linha não substitui nada que ninguém escolheu)', c11 == ['', '', ''], {'linha': c11})
            rT = r0 + 15
            xl.EnableEvents = False
            cel(rT, 'SUBSTITUI_ID').Value = 'QA97-SUBST'
            xl.EnableEvents = True
            cel(rT, 'EQUIPAMENTO').Value = base['EQUIPAMENTO']
            xl.EnableEvents = False
            c12 = txt(cel(rT, 'SUBSTITUI_ID').Value)
            cel(rT, 'SUBSTITUI_ID').ClearContents()
            cel(rT, 'EQUIPAMENTO').ClearContents()
            reg('C12 Linha só com o SUBSTITUI_ID (sem dado ainda): digitar o EQUIPAMENTO não apaga o SUBSTITUI_ID '
                '(fluxo do ADR-082)', c12 == 'QA97-SUBST', {'subst_depois': c12})

        # ------------------------------------------------------------ D01 (antes do ATUALIZAR)
        rI = r0 + 7
        fd = str(lo.ListColumns('DETALHE').DataBodyRange.Cells(1, 1).Formula)
        okf = fd.replace('[@[ID_REGISTRO]]', '[@ID_REGISTRO]').replace(' ', '') == cd.FORMULA_DETALHE.replace(' ', '')
        det = {}
        for idx in ('TESTEAUD97', 'MAN_9997', ' man_7 ', 'MAN_1,5', 'MAN_12%', 'MAN_0012'):
            cel(rI, 'ID_REGISTRO').Value = idx          # eventos desligados: o texto fica como digitado
            xl.Calculate()
            det[idx] = txt(cel(rI, 'DETALHE').Text)
        cel(rI, 'ID_REGISTRO').ClearContents()
        reg('D01 DETALHE: ID fora do padrão diz "ID fora do padrão: apague-o (um novo é gerado)"; MAN_ válido ainda não '
            'atualizado continua "atualize os dados" (OC-0080)',
            okf and det['TESTEAUD97'] == TXT_INVALIDO and det['MAN_9997'] == 'atualize os dados'
            and det[' man_7 '] == 'atualize os dados' and det['MAN_0012'] == 'atualize os dados'
            and det['MAN_1,5'] == TXT_INVALIDO and det['MAN_12%'] == TXT_INVALIDO, {'formula_igual': okf, 'detalhe': det})

        # ------------------------------------------------------------ Q01 (dados) + ATUALIZAR
        r8, r9 = r0 + 8, r0 + 9
        id8 = digitar(r8, 2 * 60 + 13, valor)
        id9 = digitar(r9, 2 * 60 + 13, round(valor * 1.002, 3))
        msg, seg = q.atualizar()
        fin = q.final_por_id()
        qa = q.ler('tblQA_Integracao')

        # ------------------------------------------------------------ M01
        m1 = {}
        for i in (idB, idQ, idM2):
            r = fin.get(i) or {}
            m1[i] = {'status': r.get('STATUS_ANALITICO'), 'motivo': txt(r.get('MOTIVO_EXCLUSAO_AUTOMATICA')),
                     'usuario': txt(r.get('USUARIO_MANUAL'))}
        okm = all(v['status'] and 'posterior ao lancamento' not in v['motivo'] and not v['usuario'].startswith('QA97')
                  for v in m1.values())
        reg('M01 Depois do ATUALIZAR: os manuais digitados na linha apagada, na órfã e na origem do recorte não são '
            'recusados por "DATA/HORA posterior ao lançamento" e o USUARIO_MANUAL não é o do carimbo antigo',
            okm, {'manuais': m1, 'atualizar_s': round(seg, 1)})

        # ------------------------------------------------------------ Q01
        f9 = fin.get(id9) or {}
        a01 = [txt(x.get('DETALHE')) for x in qa if x.get('CODIGO') == 'A01' and txt(x.get('ID_REGISTRO')) == id9]
        okq = (txt(f9.get('MOTIVO_EXCLUSAO_AUTOMATICA')).startswith('resultado manual repetido')
               and f9.get('STATUS_ANALITICO') == 'CONFLITO_MANUAL' and len(a01) == 1
               and f'corrija a HORA da linha {id9}' in a01[0] and f'Nao inative o {id8}' in a01[0]
               and 'inative um deles' not in a01[0]
               and 'repete(m) data, hora, nível e analito' in msg)
        reg('Q01 Manual com a mesma data, hora, nível e analito de outro: continua CONFLITO_MANUAL (regra do motor); o A01 '
            'manda corrigir a HORA ou apagar a linha e não inativar o que vale; a mensagem do ATUALIZAR conta a linha '
            '(OC-0081)', okq, {'id1': id8, 'id2': id9, 'status2': f9.get('STATUS_ANALITICO'),
                               'motivo2': txt(f9.get('MOTIVO_EXCLUSAO_AUTOMATICA')), 'a01': a01, 'mensagem': msg[-700:]})

        # ------------------------------------------------------------ F01 / F02
        wr = q.ws('Registros')
        wr.Unprotect(ii.SENHA)
        lr = q.lo('tblRegistros')                       # forca a lista vazia: nenhum registro com equipamento
        if lr.ListRows.Count > 0:
            lr.ListColumns(12).DataBodyRange.ClearContents()
        q.run('mCalibracao.AtualizarListasFiltro')
        f1 = {}
        for col, nome in ((2, 'lstRegLote'), (3, 'lstRegAnalito'), (7, 'lstRegEquip'), (10, 'lstRegStatus')):
            rgl = q.nome(nome)
            itens = [txt(v[0]) for v in rgl.Value] if rgl.Count > 1 else [txt(rgl.Value)]
            c = wr.Cells(4, col)
            antes_f = c.Formula
            c.Formula = 'QA97-NAO-EXISTE'
            fora = bool(c.Validation.Value)
            c.Formula = antes_f
            f1[nome] = {'itens': itens[:5], 'n': len(itens), 'texto_fora_aceito': fora}
        okf1 = all(v['n'] >= 1 and '' not in v['itens'] and v['texto_fora_aceito'] is False for v in f1.values())
        vazias = [k for k, v in f1.items() if v['itens'] == [ITEM_VAZIO]]
        if vazias:                                      # o item da lista vazia, escolhido, nao filtra nada
            col = {'lstRegLote': 2, 'lstRegAnalito': 3, 'lstRegEquip': 7, 'lstRegStatus': 10}[vazias[0]]
            c = wr.Cells(4, col)
            antes_f = c.Formula
            c.Value = ITEM_VAZIO
            r = txt(q.run('mCalibracao.AplicarFiltrosRegistros'))
            c.Formula = antes_f
            q.run('mCalibracao.AplicarFiltrosRegistros')
            vis, tot = (r.split('|') + ['', ''])[:2]
            f1['item_vazio_escolhido'] = {'lista': vazias[0], 'visiveis_total': r}
            okf1 = okf1 and vis == tot
        okf1 = okf1 and 'lstRegEquip' in vazias           # o caso foi exercitado (a coluna foi esvaziada)
        reg('F01 Filtros do Registros: nenhuma lista vazia (sem dado = "(nenhum registrado)"); texto fora da lista '
            'recusado; o item da lista vazia não filtra nada (OC-0082)', okf1, f1)
        rgl = q.nome('lstRegLote')
        lotes = [txt(v[0]) for v in rgl.Value] if rgl.Count > 1 else [txt(rgl.Value)]
        lote_num = next((x for x in lotes if x.isdigit()), None)
        c = wr.Cells(4, 2)
        antes_f = c.Formula
        if lote_num:
            c.Formula = lote_num                        # como digitar 8977 na celula
            aceito = bool(c.Validation.Value)
            tipo = type(c.Value).__name__
            c.Formula = antes_f
        else:
            aceito, tipo = None, None
        reg('F02 Filtro de Lote: o lote digitado como número é aceito (a célula é texto, como os itens da lista) '
            '(OC-0083)', str(c.NumberFormat) == '@' and (lote_num is None or aceito is True),
            {'formato': str(c.NumberFormat), 'lote': lote_num, 'aceito': aceito, 'tipo_gravado': tipo})

        # ------------------------------------------------------------ E01 (Hema)
        if not bio:
            e = q.ws('Estatística')
            e.Unprotect(ii.SENHA)
            v = e.Range('D3').Validation
            b3, d3 = e.Range('B3').Formula, e.Range('D3').Formula
            lote = txt(e.Range('B4').Text)
            equip = txt(q.run('mEstatistica.EquipFiltro'))          # o mesmo filtro do indice da Estatistica
            anos = sorted({r['DATA_HORA'].year for r in fin.values()
                           if r.get('STATUS_ANALITICO') == 'ATIVO' and norm_lote(r.get('LOTE')) == norm_lote(lote)
                           and hasattr(r.get('DATA_HORA'), 'year')
                           and (not equip or txt(r.get('EQUIPAMENTO')).upper() == equip.upper())})
            ev_e = {'lote': lote, 'equip': equip, 'anos_do_lote': anos}
            try:
                e.Range('B3').Value = 2026
                e.Range('D3').Value = 2025
                ev_e['D3_menor_aceito'] = bool(v.Value)
                ev_e['D3_validacao'] = [int(v.Type), str(v.Formula1), str(v.ErrorMessage)]
                fora = (min(anos) - 1) if anos else 2000
                e.Range('B3').Value = fora
                e.Range('D3').Value = fora
                q.run('mEstatistica.AtualizarEstatisticaAba')
                ev_e['A5_fora'] = txt(e.Range('A5').Value)
                if anos:
                    e.Range('B3').Value = anos[0]
                    e.Range('D3').Value = anos[-1]
                    q.run('mEstatistica.AtualizarEstatisticaAba')
                    ev_e['A5_cobre'] = txt(e.Range('A5').Value)
                    e.Range('B3').Value = anos[-1] + 1          # B3 acima do D3 (a validacao so protege o D3)
                    q.run('mEstatistica.AtualizarEstatisticaAba')
                    ev_e['A5_invertido'] = txt(e.Range('A5').Value)
            finally:
                e.Range('B3').Formula = b3
                e.Range('D3').Formula = d3
            oke = (ev_e['D3_menor_aceito'] is False and ev_e['D3_validacao'][0] == 1
                   and ev_e['D3_validacao'][1].replace('$', '') == '=B3'
                   and bool(anos) and ev_e['A5_fora'].startswith(f'⚠ O lote {lote} não tem resultado em {fora} (tem em ')
                   and ev_e.get('A5_cobre') == ''
                   and ev_e.get('A5_invertido', '').startswith(f'⚠ Ano Até ({anos[-1]}) menor que o Ano De ({anos[-1] + 1})'))
            reg('E01 Estatística da Hema: Ano Até menor que o Ano De é recusado; A5 avisa quando o lote em análise não tem '
                'resultado no Ano De/Até e fica vazio quando o intervalo cobre o lote (OC-0048)', oke, ev_e)
    except Exception as e:                                     # noqa: BLE001
        reg('F99 Execução interrompida: os testes seguintes a este ponto não rodaram', False,
            {'erro': f'{type(e).__name__}: {e}', 'traceback': traceback.format_exc()[-1500:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('Y00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)', hash0 == hash1,
        {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-097 (carimbo, DETALHE, A01, filtros, ano da Estatística): {len(RES) - len(falhas)} PASS / '
          f'{len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
