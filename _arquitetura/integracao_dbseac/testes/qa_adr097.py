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
  R01 (ADR-098, OC-0084) recortar um manual JA INTEGRADO para outra linha: o destino recebe de volta o carimbo de quem
      lancou (nao o de quem moveu); depois de outro ATUALIZAR o USUARIO_MANUAL continua o de quem lancou
  R02 (ADR-098) o mesmo ID colado numa linha nova (ID..MOTIVO de uma vez): com OUTRO resultado ou OUTRA data, carimbo
      novo (nao recupera o de outro lancamento); com o mesmo resultado e a mesma data, o carimbo do lancamento volta
  R03 (revisao) subir um bloco de 2 manuais integrados uma linha (recorte sobre linha ocupada): cada um fica com o
      carimbo do proprio lancamento; depois do ATUALIZAR nenhum e recusado e os USUARIO_MANUAL nao mudam
  R04 (2a/4a revisao) tocar o ID da propria linha (redigitar o mesmo ID, e com outra grafia "man_N" e NBSP) de um
      manual nao integrado com carimbo anterior a DATA/HORA: o carimbo NAO muda e o motor continua recusando
      "DATA/HORA posterior ao lancamento" (regra intacta)
  R05 (revisao) recortar um manual integrado que substitui um interfaceamento (ADR-082): o SUBSTITUI_ID vai junto e,
      depois do ATUALIZAR, o interfaceamento continua SUBSTITUIDO_MANUAL por ele
  R06 (2a revisao) subir um bloco de 2 manuais integrados com o MESMO carimbo, so o 1o com SUBSTITUI_ID: o SUBSTITUI_ID
      vai com o 1o e nao fica com o 2o; depois do ATUALIZAR o interfaceamento continua substituido pelo 1o e o 2o ATIVO
  R07 (3a revisao) recortar um manual NAO integrado para uma linha com carimbo alheio (orfao): o ID chegou (retrato dos
      IDs) -> carimbo novo; depois do ATUALIZAR fica ATIVO
  R08 (3a/4a revisao) apagar o SUBSTITUI_ID de um manual integrado e redigitar o ID dele (F2+Enter, tambem com outra
      grafia "man_N"): o SUBSTITUI_ID continua vazio; recortar a linha depois disso: o SUBSTITUI_ID continua vazio no
      destino (vale o do retrato, nao o da tblCQ_Final)
  R09 (5a revisao) o recorte com os eventos da ORIGEM e do DESTINO em ordem controlada (eventos desligados no Cut e
      EntradaMudou chamado a mao): origem antes -> o destino (vazio) recebe o carimbo e o SUBSTITUI_ID do lancamento;
      destino antes -> sobre uma linha ocupada, o destino recebe o carimbo do lancamento que chegou; R09c: a origem
      perto do fim da tabela faz a tabela crescer (FOLGA) no evento dela, e o destino ainda ve o ID chegar
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

        def carimbar_velho(r, quem, quando=VELHO):      # eventos DESLIGADOS: como um carimbo antigo/orfao
            xl.EnableEvents = False
            cel(r, 'USUARIO').Value = quem
            cel(r, 'REGISTRADO_EM').Value = quando
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
        if r0 + 50 > len(vids):
            lo.Resize(lo.Range.Resize(lo.Range.Rows.Count + 60))
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
        # R01 (dados): um manual com dono e hora conhecidos (eventos desligados, como um lancamento de outra pessoa)
        rDono, rDest, rNovo = r0 + 16, r0 + 17, r0 + 18
        valor_dono = round(valor * 0.997, 3)
        idDono = digitar(rDono, 3 * 60 + 25, valor_dono)
        cel(rDono, 'USUARIO').Value = 'QA98_DONO'
        cel(rDono, 'REGISTRADO_EM').Value = dia + dt.timedelta(hours=4)
        car_dono = carimbo(rDono)
        # R03 (dados): dois manuais de donos e horas diferentes; o carimbo do P e ANTERIOR a DATA/HORA do Q
        rVazia, rP, rQq = r0 + 21, r0 + 22, r0 + 23
        idP = digitar(rP, 3 * 60 + 31, round(valor * 0.995, 3))
        idQq = digitar(rQq, 3 * 60 + 35, round(valor * 0.993, 3))
        carP = carimbar_velho(rP, 'QA98_P', dia + dt.timedelta(minutes=3 * 60 + 33))
        carQ = carimbar_velho(rQq, 'QA98_Q', dia + dt.timedelta(hours=4))
        # R05 (dados): um manual que substitui o interfaceamento de referencia (ADR-082)
        rS5, rS5d = r0 + 26, r0 + 27
        idS5, subst5, ref_id = '', '', txt(ref.get('ID_REGISTRO'))
        if 'SUBSTITUI_ID' in nomes:
            xl.EnableEvents = True
            cel(rS5, 'SUBSTITUI_ID').Value = ref_id     # a linha se preenche com os dados do resultado indicado
            cel(rS5, 'RESULTADO').Value = round(valor * 1.01, 3)
            if not txt(cel(rS5, 'MOTIVO').Value):
                cel(rS5, 'MOTIVO').Value = base['MOTIVO']
            xl.EnableEvents = False
            idS5, subst5 = txt(cel(rS5, 'ID_REGISTRO').Value), txt(cel(rS5, 'SUBSTITUI_ID').Value)
            carimbar_velho(rS5, 'QA98_S', ingenuo(ref['DATA_HORA']) + dt.timedelta(minutes=1))
        car_s5 = carimbo(rS5)
        # R06 (dados): A substitui um 2o interfaceamento, B nao; os dois com o MESMO carimbo (colados juntos)
        r6v, r6a, r6b = r0 + 29, r0 + 30, r0 + 31
        reais_ord = sorted(reais, key=lambda r: r['DATA_HORA'], reverse=True)
        ref6 = next((r for r in reais_ord if txt(r.get('ID_REGISTRO')) != ref_id), None)
        id6a = id6b = subst6 = ''
        ref6_id = txt((ref6 or {}).get('ID_REGISTRO'))
        if 'SUBSTITUI_ID' in nomes and ref6:
            xl.EnableEvents = True
            cel(r6a, 'SUBSTITUI_ID').Value = ref6_id
            cel(r6a, 'RESULTADO').Value = round(float(ref6['RESULTADO']) * 1.01, 3)
            if not txt(cel(r6a, 'MOTIVO').Value):
                cel(r6a, 'MOTIVO').Value = base['MOTIVO']
            xl.EnableEvents = False
            id6a, subst6 = txt(cel(r6a, 'ID_REGISTRO').Value), txt(cel(r6a, 'SUBSTITUI_ID').Value)
            id6b = digitar(r6b, 3 * 60 + 45, round(valor * 0.989, 3))
            quando6 = max(ingenuo(ref6['DATA_HORA']), dia + dt.timedelta(minutes=3 * 60 + 45)) + dt.timedelta(minutes=2)
            car6a = carimbar_velho(r6a, 'QA98_AB', quando6)
            car6b = carimbar_velho(r6b, 'QA98_AB', quando6)
        # R09 (dados): M9 substitui um 3o interfaceamento; N9 e O9 sem SUBSTITUI_ID, donos diferentes (o carimbo do O9 e
        # anterior a DATA/HORA do N9: no lugar dele, o N9 seria recusado)
        rM9, rM9d, rN9, rO9 = r0 + 40, r0 + 42, r0 + 44, r0 + 46
        ref9 = next((r for r in reais_ord if txt(r.get('ID_REGISTRO')) not in (ref_id, ref6_id)), None)
        ref9_id = txt((ref9 or {}).get('ID_REGISTRO'))
        idM9 = subst9 = ''
        car_m9 = ('', None)
        if 'SUBSTITUI_ID' in nomes and ref9:
            xl.EnableEvents = True
            cel(rM9, 'SUBSTITUI_ID').Value = ref9_id
            cel(rM9, 'RESULTADO').Value = round(float(ref9['RESULTADO']) * 1.01, 3)
            if not txt(cel(rM9, 'MOTIVO').Value):
                cel(rM9, 'MOTIVO').Value = base['MOTIVO']
            xl.EnableEvents = False
            idM9, subst9 = txt(cel(rM9, 'ID_REGISTRO').Value), txt(cel(rM9, 'SUBSTITUI_ID').Value)
            car_m9 = carimbar_velho(rM9, 'QA98_M9', ingenuo(ref9['DATA_HORA']) + dt.timedelta(minutes=3))
        idN9 = digitar(rN9, 4 * 60 + 5, round(valor * 0.985, 3))
        idO9 = digitar(rO9, 3 * 60 + 55, round(valor * 0.983, 3))
        car_n9 = carimbar_velho(rN9, 'QA98_N9', dia + dt.timedelta(hours=5))
        car_o9 = carimbar_velho(rO9, 'QA98_O9', dia + dt.timedelta(minutes=3 * 60 + 57))
        # R09c (dados): X sera a origem perto do fim da tabela; Y o ocupante (carimbo anterior a DATA/HORA do X)
        rYc, rXc = r0 + 38, r0 + 48
        idXc = digitar(rXc, 4 * 60 + 15, round(valor * 0.981, 3))
        idYc = digitar(rYc, 4 * 60 + 9, round(valor * 0.979, 3))
        car_xc = carimbar_velho(rXc, 'QA98_XC', dia + dt.timedelta(hours=5))
        carimbar_velho(rYc, 'QA98_YC', dia + dt.timedelta(minutes=4 * 60 + 10))
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

        # ------------------------------------------------------------ R01..R05 (ADR-098, OC-0084)
        def mesmo_instante(a, b):
            return hasattr(a, 'year') and hasattr(b, 'year') and abs((a - b).total_seconds()) <= 1

        def mesmo_carimbo(c, ref_c):
            return c[0] == ref_c[0] and mesmo_instante(c[1], ref_c[1])

        def integrado(i, quem):
            return txt((fin.get(i) or {}).get('USUARIO_MANUAL')) == quem

        # R01: destino vazio
        f_dono = fin.get(idDono) or {}
        pre_r1 = integrado(idDono, 'QA98_DONO')                         # o caso foi exercitado
        xl.EnableEvents = True
        bloco(rDono).Cut(bloco(rDest))
        xl.EnableEvents = False
        c_dest, c_orig = carimbo(rDest), carimbo(rDono)
        id_dest = txt(cel(rDest, 'ID_REGISTRO').Value)
        # R03: sobe o bloco P,Q uma linha (o Q cai na linha que era do P, com o carimbo dele)
        pre_r3 = integrado(idP, 'QA98_P') and integrado(idQq, 'QA98_Q')
        xl.EnableEvents = True
        ws.Range(cel(rP, 'ID_REGISTRO'), lo.DataBodyRange.Cells(rQq, c_mot)).Cut(cel(rVazia, 'ID_REGISTRO'))
        xl.EnableEvents = False
        r3 = {'ids': [txt(cel(k, 'ID_REGISTRO').Value) for k in (rVazia, rP, rQq)],
              'carimbos': [carimbo(k) for k in (rVazia, rP, rQq)]}
        # R04: redigitar o ID da propria linha (F2+Enter) de um manual nao integrado com carimbo invalido (2020)
        rN4 = r0 + 24
        idN4 = digitar(rN4, 3 * 60 + 41, round(valor * 0.991, 3))
        car4 = carimbar_velho(rN4, 'QA98_PROPRIO')
        grafias4 = [idN4, 'man_' + str(int(idN4[4:])), idN4 + chr(160)]
        c4s = []
        for g in grafias4:
            xl.EnableEvents = True
            cel(rN4, 'ID_REGISTRO').Value = g
            xl.EnableEvents = False
            c4s.append((txt(cel(rN4, 'ID_REGISTRO').Value), carimbo(rN4)))
        c4d = carimbo(rN4)
        id4d = txt(cel(rN4, 'ID_REGISTRO').Value)
        # R07: manual NAO integrado recortado para uma linha vazia com carimbo orfao alheio (2020)
        rN7, rN7d = r0 + 33, r0 + 34
        idN7 = digitar(rN7, 3 * 60 + 51, round(valor * 0.987, 3))
        carimbar_velho(rN7d, 'QA98_ALHEIO')
        xl.EnableEvents = True
        bloco(rN7).Cut(bloco(rN7d))
        xl.EnableEvents = False
        c7d, id7d = carimbo(rN7d), txt(cel(rN7d, 'ID_REGISTRO').Value)
        # R06: sobe o bloco A,B uma linha
        pre_r6 = bool(id6a) and integrado(id6a, 'QA98_AB') and integrado(id6b, 'QA98_AB') and \
            (fin.get(ref6_id) or {}).get('STATUS_ANALITICO') == 'SUBSTITUIDO_MANUAL'
        r6 = {}
        if id6a:
            xl.EnableEvents = True
            ws.Range(cel(r6a, 'ID_REGISTRO'), lo.DataBodyRange.Cells(r6b, c_mot)).Cut(cel(r6v, 'ID_REGISTRO'))
            xl.EnableEvents = False
            r6 = {'ids': [txt(cel(k, 'ID_REGISTRO').Value) for k in (r6v, r6a, r6b)],
                  'subst': [txt(cel(k, 'SUBSTITUI_ID').Value) for k in (r6v, r6a, r6b)],
                  'carimbos': [carimbo(k) for k in (r6v, r6a, r6b)]}
        # R05: manual com SUBSTITUI_ID recortado
        pre_r5 = bool(idS5) and integrado(idS5, 'QA98_S') and \
            (fin.get(ref_id) or {}).get('STATUS_ANALITICO') == 'SUBSTITUIDO_MANUAL'
        r5 = {}
        if idS5:
            xl.EnableEvents = True
            bloco(rS5).Cut(bloco(rS5d))
            xl.EnableEvents = False
            r5 = {'id_dest': txt(cel(rS5d, 'ID_REGISTRO').Value), 'subst_dest': txt(cel(rS5d, 'SUBSTITUI_ID').Value),
                  'subst_orig': txt(cel(rS5, 'SUBSTITUI_ID').Value), 'carimbo_dest': carimbo(rS5d)}

        xl.EnableEvents = False
        ws.Unprotect(ii.SENHA)
        lo.Resize(ws.Range(lo.Range.Cells(1, 1), lo.Range.Cells(lo.HeaderRowRange.Row - lo.Range.Row + 1 + rXc + 3,
                                                                  lo.Range.Columns.Count)))
        n_curta = lo.ListRows.Count
        msg2, seg2 = q.atualizar()                       # o PrepararEntradas tira o retrato inteiro da tabela curta
        fin2 = q.final_por_id()
        # R09c: logo depois do ATUALIZAR, antes de qualquer outro evento perto do fim
        pre_r9c = integrado(idXc, 'QA98_XC') and integrado(idYc, 'QA98_YC') and n_curta == rXc + 3
        xl.EnableEvents = False
        bloco(rXc).Cut(bloco(rYc))                       # o X (perto do fim) por cima do Y
        q.run('mIntegracao.EntradaMudou', ws, bloco(rXc))   # origem primeiro: a FOLGA cresce a tabela
        xl.EnableEvents = False
        n_depois = lo.ListRows.Count
        q.run('mIntegracao.EntradaMudou', ws, bloco(rYc))
        xl.EnableEvents = False
        r9c = (txt(cel(rYc, 'ID_REGISTRO').Value), carimbo(rYc))
        reg('R09c Recorte cuja origem está perto do fim da tabela (a tabela cresce no evento da origem): o destino ainda vê o '
            'ID chegar e recebe o carimbo do lançamento',
            pre_r9c and n_depois > n_curta and r9c[0] == idXc and mesmo_carimbo(r9c[1], car_xc),
            {'id': idXc, 'ocupante': idYc, 'linhas': [n_curta, n_depois], 'destino': [r9c[0], r9c[1][0], str(r9c[1][1])],
             'esperado': [car_xc[0], str(car_xc[1])]})
        f2 = fin2.get(idDono) or {}
        reg('R01 Recortar um manual já integrado para outra linha: o destino recebe de volta o carimbo de quem lançou (não '
            'o de quem moveu) e, depois de outro ATUALIZAR, o USUARIO_MANUAL continua o de quem lançou (OC-0084)',
            pre_r1 and id_dest == idDono and mesmo_carimbo(c_dest, car_dono)
            and c_orig == ('', None) and txt(f2.get('USUARIO_MANUAL')) == 'QA98_DONO'
            and f2.get('STATUS_ANALITICO') == f_dono.get('STATUS_ANALITICO'),
            {'id': idDono, 'usuario_integrado_antes': txt(f_dono.get('USUARIO_MANUAL')),
             'carimbo_lancamento': [car_dono[0], str(car_dono[1])], 'destino': [c_dest[0], str(c_dest[1])],
             'origem': [c_orig[0], str(c_orig[1])], 'usuario_integrado_depois': txt(f2.get('USUARIO_MANUAL')),
             'status': [f_dono.get('STATUS_ANALITICO'), f2.get('STATUS_ANALITICO')], 'atualizar2_s': round(seg2, 1)})

        def estado2(i):
            r = fin2.get(i) or {}
            return {'status': r.get('STATUS_ANALITICO'), 'usuario': txt(r.get('USUARIO_MANUAL')),
                    'motivo': txt(r.get('MOTIVO_EXCLUSAO_AUTOMATICA'))}

        e3 = {idP: estado2(idP), idQq: estado2(idQq)}
        ok3 = (pre_r3 and r3['ids'] == [idP, idQq, ''] and mesmo_carimbo(r3['carimbos'][0], carP)
               and mesmo_carimbo(r3['carimbos'][1], carQ) and r3['carimbos'][2] == ('', None)
               and e3[idP]['usuario'] == 'QA98_P' and e3[idQq]['usuario'] == 'QA98_Q'
               and all(v['status'] != 'MANUAL_INCOMPLETO' and 'posterior' not in v['motivo'] for v in e3.values()))
        reg('R03 Subir um bloco de 2 manuais integrados uma linha (recorte sobre linha ocupada): cada um fica com o carimbo '
            'do próprio lançamento; depois do ATUALIZAR nenhum é recusado por "posterior" e os USUARIO_MANUAL não mudam',
            ok3, {'ids': r3['ids'], 'carimbos': [[c[0], str(c[1])] for c in r3['carimbos']],
                  'esperado': [[carP[0], str(carP[1])], [carQ[0], str(carQ[1])]], 'depois': e3})
        e4 = estado2(idN4)
        reg('R04 Redigitar o ID da própria linha (manual não integrado, carimbo anterior à DATA/HORA): o carimbo NÃO muda e '
            'o motor continua recusando "DATA/HORA posterior ao lançamento" (a regra não muda)',
            id4d == idN4 and all(x[0] == idN4 and x[1] == car4 for x in c4s)
            and e4['status'] == 'MANUAL_INCOMPLETO' and 'posterior' in e4['motivo'],
            {'id': idN4, 'carimbo_antes': [car4[0], str(car4[1])],
             'por_grafia': [[repr(g), x[0], x[1][0], str(x[1][1])] for g, x in zip(grafias4, c4s)], 'depois': e4})
        if 'SUBSTITUI_ID' in nomes:
            fr6 = fin2.get(ref6_id) or {}
            rel6 = txt(fr6.get('ID_RELACIONADO')).split('~')[0]
            ok6 = (pre_r6 and r6.get('ids') == [id6a, id6b, ''] and r6['subst'] == [subst6, '', '']
                   and mesmo_carimbo(r6['carimbos'][0], car6a) and mesmo_carimbo(r6['carimbos'][1], car6b)
                   and fr6.get('STATUS_ANALITICO') == 'SUBSTITUIDO_MANUAL' and rel6 == id6a
                   and estado2(id6a)['status'] == 'ATIVO' and estado2(id6b)['status'] != 'MANUAL_INCOMPLETO'
                   and 'SUBSTITUI_ID' not in estado2(id6b)['motivo'].upper())
            reg('R06 Subir um bloco de 2 manuais integrados com o mesmo carimbo, só o 1º com SUBSTITUI_ID: o SUBSTITUI_ID vai '
                'com o 1º e não fica com o 2º; depois do ATUALIZAR o interfaceamento continua substituído pelo 1º e o 2º não é '
                'recusado nem acusado de SUBSTITUI_ID',
                ok6, {'ids': [id6a, id6b], 'substitui': subst6, 'interfaceamento': ref6_id, 'pre': pre_r6,
                      'movido': {'ids': r6.get('ids'), 'subst': r6.get('subst'),
                                 'carimbos': [[c[0], str(c[1])] for c in r6.get('carimbos', [])]},
                      'interf_depois': [fr6.get('STATUS_ANALITICO'), rel6],
                      'manuais_depois': [estado2(id6a), estado2(id6b)]})
        if 'SUBSTITUI_ID' in nomes:
            fr = fin2.get(ref_id) or {}
            rel = txt(fr.get('ID_RELACIONADO')).split('~')[0]
            ok5 = (pre_r5 and r5.get('id_dest') == idS5 and r5.get('subst_dest') == subst5 and r5.get('subst_orig') == ''
                   and mesmo_carimbo(r5['carimbo_dest'], car_s5)
                   and fr.get('STATUS_ANALITICO') == 'SUBSTITUIDO_MANUAL' and rel == idS5
                   and estado2(idS5)['status'] == 'ATIVO' and estado2(idS5)['usuario'] == 'QA98_S')
            reg('R05 Recortar um manual integrado que substitui um interfaceamento (ADR-082): o SUBSTITUI_ID e o carimbo '
                'vão junto; depois do ATUALIZAR o interfaceamento continua SUBSTITUIDO_MANUAL por ele',
                ok5, {'id': idS5, 'substitui': subst5, 'interfaceamento': ref_id, 'pre': pre_r5,
                      'movido': {k: (v if k != 'carimbo_dest' else [v[0], str(v[1])]) for k, v in r5.items()},
                      'interf_depois': [fr.get('STATUS_ANALITICO'), rel], 'manual_depois': estado2(idS5)})

        e7 = estado2(idN7)
        reg('R07 Recortar um manual não integrado para uma linha com carimbo alheio (órfão): o ID chegou, então o carimbo é '
            'novo; depois do ATUALIZAR não é recusado por "posterior" e o dono não é o do carimbo alheio',
            id7d == idN7 and novo(c7d, 'QA98_ALHEIO') and e7['status'] not in ('', None, 'MANUAL_INCOMPLETO')
            and 'posterior' not in e7['motivo'] and e7['usuario'] != 'QA98_ALHEIO',
            {'id': idN7, 'destino': id7d, 'carimbo': [c7d[0], str(c7d[1])], 'depois': e7})
        if 'SUBSTITUI_ID' in nomes and idS5:
            s8_antes = txt(cel(rS5d, 'SUBSTITUI_ID').Value)      # tem de ser o SUBSTITUI_ID do R05 (o caso existe)
            xl.EnableEvents = True
            cel(rS5d, 'SUBSTITUI_ID').ClearContents()   # desfazer a substituicao (GUIA: apague o SUBSTITUI_ID)
            xl.EnableEvents = False
            s8_apagado = txt(cel(rS5d, 'SUBSTITUI_ID').Value)
            s8 = []
            for g in (idS5, 'man_' + str(int(idS5[4:]))):
                xl.EnableEvents = True
                cel(rS5d, 'ID_REGISTRO').Value = g       # F2+Enter no ID da propria linha (e outra grafia)
                xl.EnableEvents = False
                s8.append(txt(cel(rS5d, 'SUBSTITUI_ID').Value))
            rS8 = r0 + 36                                # e depois recorta a linha (o retrato manda: vazio)
            xl.EnableEvents = True
            bloco(rS5d).Cut(bloco(rS8))
            xl.EnableEvents = False
            s8_rec = (txt(cel(rS8, 'ID_REGISTRO').Value), txt(cel(rS8, 'SUBSTITUI_ID').Value), carimbo(rS8))
            reg('R08 Apagar o SUBSTITUI_ID de um manual integrado e redigitar o ID dele (também como "man_N"): o SUBSTITUI_ID '
                'continua vazio; recortar a linha depois disso: o destino fica sem SUBSTITUI_ID e com o carimbo do lançamento',
                pre_r5 and s8_antes == subst5 and s8_apagado == '' and s8 == ['', ''] and s8_rec[0] == idS5
                and s8_rec[1] == '' and mesmo_carimbo(s8_rec[2], car_s5),
                {'id': idS5, 'subst_antes': s8_antes, 'apos_apagar': s8_apagado, 'apos_tocar': s8,
                 'recorte': [s8_rec[0], s8_rec[1], s8_rec[2][0], str(s8_rec[2][1])]})

        # R09: ordem dos eventos controlada (eventos desligados no Cut; o EntradaMudou chamado a mao)
        def evento(r):
            q.run('mIntegracao.EntradaMudou', ws, bloco(r))
            xl.EnableEvents = False

        r9 = {}
        pre_r9a = bool(idM9) and integrado(idM9, 'QA98_M9') and \
            (fin.get(ref9_id) or {}).get('STATUS_ANALITICO') == 'SUBSTITUIDO_MANUAL'
        if 'SUBSTITUI_ID' in nomes and not idM9:
            reg('R09a Recorte com o evento da origem ANTES do destino: NÃO EXERCITADO (sem um 3º interfaceamento para '
                'substituir)', False, {'pre': pre_r9a})
        if idM9:
            xl.EnableEvents = False
            bloco(rM9).Cut(bloco(rM9d))
            evento(rM9)                                  # origem primeiro
            evento(rM9d)
            r9['a'] = {'id': txt(cel(rM9d, 'ID_REGISTRO').Value), 'subst': txt(cel(rM9d, 'SUBSTITUI_ID').Value),
                       'carimbo': carimbo(rM9d), 'origem': [txt(cel(rM9, c).Value) for c in ('ID_REGISTRO', 'SUBSTITUI_ID', 'USUARIO')]}
            reg('R09a Recorte com o evento da origem ANTES do destino (vazio): o destino recebe o carimbo e o SUBSTITUI_ID '
                'do lançamento; a origem fica vazia',
                pre_r9a and r9['a']['id'] == idM9 and r9['a']['subst'] == subst9 and mesmo_carimbo(r9['a']['carimbo'], car_m9)
                and r9['a']['origem'] == ['', '', ''],
                {'id': idM9, 'substitui': subst9, 'pre': pre_r9a, 'destino': [r9['a']['id'], r9['a']['subst'],
                 r9['a']['carimbo'][0], str(r9['a']['carimbo'][1])], 'esperado': [car_m9[0], str(car_m9[1])],
                 'origem': r9['a']['origem']})
        pre_r9b = integrado(idN9, 'QA98_N9') and integrado(idO9, 'QA98_O9')
        xl.EnableEvents = False
        bloco(rN9).Cut(bloco(rO9))                       # o N9 por cima do O9
        evento(rO9)                                      # destino primeiro
        evento(rN9)
        r9['b'] = {'id': txt(cel(rO9, 'ID_REGISTRO').Value), 'carimbo': carimbo(rO9),
                   'origem': [txt(cel(rN9, c).Value) for c in ('ID_REGISTRO', 'USUARIO')]}
        reg('R09b Recorte por cima de outro manual com o evento do destino ANTES da origem: o destino recebe o carimbo do '
            'lançamento que chegou (não o do que estava ali); a origem fica vazia',
            pre_r9b and r9['b']['id'] == idN9 and mesmo_carimbo(r9['b']['carimbo'], car_n9)
            and r9['b']['origem'] == ['', ''],
            {'id': idN9, 'ocupante': idO9, 'pre': pre_r9b, 'destino': [r9['b']['id'], r9['b']['carimbo'][0],
             str(r9['b']['carimbo'][1])], 'esperado': [car_n9[0], str(car_n9[1])], 'ocupante_carimbo': [car_o9[0], str(car_o9[1])],
             'origem': r9['b']['origem']})

        # R02: o mesmo ID colado numa linha nova, ID..MOTIVO de uma vez (o ID so existe na tblCQ_Final)
        xl.EnableEvents = True
        bloco(rDest).ClearContents()
        xl.EnableEvents = False
        cols_bloco = nomes[c_id - 1:c_mot]
        rR2 = r0 + 18

        def colar_linha(res, data):
            v = dict(base, ID_REGISTRO=idDono, DATA=data, HORA=(3 * 60 + 25) / 1440.0, RESULTADO=res)
            xl.EnableEvents = True
            bloco(rR2).Value = (tuple(v.get(c) for c in cols_bloco),)
            xl.EnableEvents = False
            out = (txt(cel(rR2, 'ID_REGISTRO').Value), carimbo(rR2))
            xl.EnableEvents = True
            bloco(rR2).ClearContents()
            xl.EnableEvents = False
            return out

        r2a = colar_linha(round(valor * 1.004, 3), dia)                         # outro resultado
        r2b = colar_linha(valor_dono, dia - dt.timedelta(days=1))               # outra data
        r2c = colar_linha(valor_dono, dia)                                      # o mesmo lancamento
        reg('R02 O mesmo ID colado numa linha nova: com outro resultado ou outra data, carimbo novo (não recupera o de '
            'outro lançamento); com o mesmo resultado e a mesma data, o carimbo do lançamento volta',
            pre_r1 and r2a[0] == idDono and novo(r2a[1], 'QA98_DONO') and r2b[0] == idDono
            and novo(r2b[1], 'QA98_DONO') and r2c[0] == idDono and mesmo_carimbo(r2c[1], car_dono),
            {k: [v[0], v[1][0], str(v[1][1])] for k, v in (('outro_resultado', r2a), ('outra_data', r2b),
                                                           ('mesmo_lancamento', r2c))})

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
