# -*- coding: utf-8 -*-
"""QA do ADR-087 (29/10/2026): aba Registros como Tabela do Excel que cresce sozinha, filtros por lote e demais,
registro do CEQ pela janela frmCEQ e auditoria das janelas. Roda numa COPIA instalada; nada e salvo.

Uso: python qa_adr087.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  R01 tblRegistros existe (A..M, coluna Lote); os registros de cada lote da RegistrosStore (fora o carregado) estao na
      tabela com o seu lote; a RegistrosStore ficou intacta (arquivo historico)
  R02 crescimento: a ultima linha esta sempre vazia; registrar pela ID usa/acrescenta linha; tres registros digitados
      em sequencia (como o usuario) entram com o Lote em uso, o Nº calculado, a validacao e o formato da tabela, e a
      proxima linha aparece sozinha; os registros anteriores nao mudam
  R03 filtros: Lote (dos dados reais) mostra so o lote; combinado com Analito; sem resultado = aviso claro; Limpar volta
      tudo; filtrar NAO altera nenhum registro (valores da tabela iguais antes e depois)
  R04 listas dos filtros acompanham os dados (lote novo aparece) e incluem os lotes cadastrados
  R05 icone de calibracao do grafico: so as linhas do lote em uso, como a vista por lote de antes
  R06 assinatura numa linha nova da tabela (login e senha) e o desenho pelas linhas da tabela
  R07 com filtro ativo a linha vazia de digitacao continua visivel; assinatura so nas linhas visiveis; filtro de
      Analito acha o registro com espaco sobrando
  C01 CEQ pela janela (nucleo SalvarCEQ): grava na tabela do provedor (todos os campos), consolida a EQA_Base (a chave
      aparece), Audit_Log CEQ_REGISTRADO
  C02 CEQ duplicado (mesma chave provedor|ano|rodada|analito|amostra) recusado, nada gravado
  C03 CEQ invalido recusado com o campo certo: sem analito, numero invalido, limite inferior > superior, ano errado
  C04 reaproveitamento: rodada existente traz o ano; analito existente traz a unidade; lista da rodada
  V02 telas principais: auditoria visual sem achado (nada cortado nem por cima de texto)
  V01 frmCEQ instalado igual a especificacao (controles e propriedades); auditoria das janelas sem corte/sobreposicao
  R00 seguranca do teste: arquivo de entrada intacto
"""
import json
import os
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402
from qa_inativar_ux import criar_usuario, fazer_login, Leitor, SENHA_QA, USUARIO_QA  # noqa: E402
import auditoria_janelas  # noqa: E402
import auditoria_abas  # noqa: E402

TELAS_PRINCIPAIS = ['Início', 'Painel', 'Registros', 'Inativar', 'COMENTARIOS_TECNICOS', 'Liberação',
                    'EQA.CAP_Resumo', 'EQA.Controllab_Resumo', 'Digitar Resultados', 'Usuarios']

TB = 'tblRegistros'


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    L = Leitor(q)

    def run(macro, *a, teto=300):
        return str(q.run(macro, *a, teto=teto))

    try:
        xl = q.ex.xl
        # ---------------------------------------------------------------- V02 telas (antes de qualquer mudanca do teste)
        achados_telas = {}
        for aba in TELAS_PRINCIPAIS:
            try:
                ws_ = q.wb.Worksheets(aba)
            except Exception:                            # noqa: BLE001
                continue
            ws_.Unprotect(ii.SENHA)
            achados_telas[aba] = auditoria_abas.auditar_aba(ws_)
        q.run('mSeguranca.ReprotectAll')                 # as abas voltam protegidas como o usuario as ve
        reg('V02 Telas principais sem texto cortado, botao/barra por cima de texto, grafico por cima de celulas ou '
            'formas sobrepostas (auditoria visual medida com a fonte real)',
            not any(achados_telas.values()),
            {k: v[:4] for k, v in achados_telas.items() if v} or {'telas': list(achados_telas)})
        rw = q.ws('Registros')
        lo = rw.ListObjects(TB)
        criar_usuario(q, USUARIO_QA, 'ANALISTA')
        fazer_login(q, USUARIO_QA)
        cab = [txt(c.Name) for c in lo.ListColumns]

        def corpo():
            return [tuple(r) for r in lo.DataBodyRange.Value] if lo.ListRows.Count else []

        # ---------------------------------------------------------------- R01
        st = q.ws('RegistrosStore')
        lotes = [txt(r[0]) for r in q.nome('regLoteCol').Value]
        carregado = txt(q.nome('loteCarregado').Value)
        faltam, conferidos = [], 0
        dados = corpo()
        chaves = {(txt(r[12]),) + tuple(txt(x) for x in r[1:12]) for r in dados}
        for i, lote in enumerate(lotes):
            if not lote or lote == carregado:
                continue
            r0 = 2 + i * 200
            for ln in st.Range(st.Cells(r0, 1), st.Cells(r0 + 199, 11)).Value:
                if any(txt(x) for x in ln):
                    conferidos += 1
                    if (lote,) + tuple(txt(x) for x in ln) not in chaves:
                        faltam.append((lote, [txt(x) for x in ln[:4]]))
        reg('R01 Aba Registros = Tabela do Excel tblRegistros (A..M, coluna Lote); os registros de cada lote da '
            'RegistrosStore estão na tabela com o seu lote; a RegistrosStore ficou como arquivo histórico',
            cab[0] == 'Nº' and cab[-1] == 'Lote' and len(cab) == 13 and not faltam,
            {'colunas': cab, 'registros_da_store_conferidos': conferidos, 'faltando': faltam[:5],
             'linhas_na_tabela': lo.ListRows.Count})

        # ---------------------------------------------------------------- R02 crescimento
        antes = corpo()
        ult_vazia = not any(txt(x) for x in antes[-1][1:12]) if antes else False
        n0 = lo.ListRows.Count
        lote_uso = txt(q.nome('loteAtivo').Value)
        an = txt(q.espec().Range('A4').Value)
        rw.Unprotect(ii.SENHA)
        xl.EnableEvents = True
        linhas_novas = []
        for k in range(3):                              # como o usuario: digita na ultima linha (eventos ligados)
            r = lo.Range.Row + lo.Range.Rows.Count - 1
            rw.Cells(r, 3).Value = an
            rw.Cells(r, 7).Value = f'QA ADR-087: registro em sequência {k + 1}'
            linhas_novas.append(r)
        depois = corpo()
        ok_nov = True
        detalhes = []
        for r in linhas_novas:
            i = r - lo.DataBodyRange.Row
            lin = depois[i]
            v_niv = rw.Cells(r, 4).Validation
            try:
                f_niv = str(v_niv.Formula1)
            except Exception:                           # noqa: BLE001
                f_niv = ''
            okl = (txt(lin[12]) == lote_uso and isinstance(lin[0], (int, float)) and int(lin[0]) == i + 1
                   and f_niv != '' and str(rw.Cells(r, 2).NumberFormat) == str(lo.ListColumns(2).DataBodyRange.Cells(1, 1).NumberFormat))
            ok_nov = ok_nov and okl
            detalhes.append({'linha': r, 'lote': lin[12], 'n': lin[0], 'validacao_nivel': f_niv, 'ok': okl})
        ultima = depois[-1]
        reg('R02 Tabela que cresce sozinha: a última linha fica sempre vazia; 3 registros digitados em sequência entram '
            'com o Lote em uso, o Nº calculado, a validação e o formato da tabela; os anteriores não mudam',
            ult_vazia and ok_nov and lo.ListRows.Count == n0 + 3 and not any(txt(x) for x in ultima[1:12])
            and depois[:n0 - 1] == antes[:n0 - 1],
            {'linhas_antes': n0, 'linhas_depois': lo.ListRows.Count, 'novas': detalhes,
             'ultima_vazia_antes': ult_vazia, 'ultima_vazia_depois': not any(txt(x) for x in ultima[1:12])})

        # ---------------------------------------------------------------- R03 filtros
        valores0 = corpo()
        lotes_tab = sorted({txt(r[12]) for r in valores0 if txt(r[12])})
        alvo = lote_uso if lote_uso in lotes_tab else (lotes_tab[0] if lotes_tab else '')
        rw.Range('B4').Value = "'" + alvo                # Worksheet_Change aplica
        vis = [lo.DataBodyRange.Rows(i + 1).EntireRow.Hidden is False for i in range(lo.ListRows.Count)]
        so_lote = all((txt(r[12]) == alvo) for r, v in zip(valores0, vis) if v and any(txt(x) for x in r[1:12]))
        info1 = txt(rw.Range('K4').Value)
        rw.Range('C4').Value = an
        vis2 = [lo.DataBodyRange.Rows(i + 1).EntireRow.Hidden is False for i in range(lo.ListRows.Count)]
        comb = all(txt(r[12]) == alvo and txt(r[2]) == an for r, v in zip(valores0, vis2) if v and any(txt(x) for x in r[1:12]))
        info2 = txt(rw.Range('K4').Value)
        rw.Range('C4').Value = 'ANALITO_QUE_NAO_EXISTE'
        info3 = txt(rw.Range('K4').Value)
        cor3 = rw.Range('K4').Font.Color
        run('mCalibracao.LimparFiltrosRegistros')
        vis4 = [lo.DataBodyRange.Rows(i + 1).EntireRow.Hidden is False for i in range(lo.ListRows.Count)]
        info4 = txt(rw.Range('K4').Value)
        valores1 = corpo()
        reg('R03 Filtros: Lote (dados reais) mostra só o lote; combinado com Analito; sem resultado = aviso claro; Limpar '
            'filtros mostra tudo; filtrar não altera nenhum registro',
            alvo and so_lote and 'Lote ' + alvo in info1 and comb and 'de' in info2 and 'Nenhum registro' in info3
            and all(vis4) and info4.startswith('Sem filtro') and valores0 == valores1 and rw.Range('B4').Value in (None, ''),
            {'lote': alvo, 'info_lote': info1, 'info_lote_analito': info2, 'info_sem_resultado': info3,
             'cor_sem_resultado': cor3, 'info_limpo': info4, 'registros_iguais': valores0 == valores1})

        # ---------------------------------------------------------------- R04 listas
        r_novo = lo.Range.Row + lo.Range.Rows.Count - 1
        rw.Cells(r_novo, 3).Value = an
        rw.Cells(r_novo, 13).Value = "'QA_LOTE_NOVO"
        run('mCalibracao.AtualizarListasFiltro')
        lst = [txt(v[0]) for v in q.nome('lstRegLote').Value]
        reg('R04 As listas dos filtros acompanham os dados (lote novo aparece) e incluem os lotes cadastrados',
            'QA_LOTE_NOVO' in lst and all(l in lst for l in lotes if l), {'lista_lote': lst})

        # ---------------------------------------------------------------- R05 icone de calibracao
        calc = q.ws('Calc')
        linha3 = calc.Range('A3:EZ3').Formula[0]
        cols = [k + 1 for k, f in enumerate(linha3) if 'regCalib' in str(f or '')]
        forms = [str(calc.Cells(3, k).Formula) for k in cols]
        reg('R05 Ícone de calibração: as fórmulas contam só as linhas do lote em uso (regLote = loteAtivo), como a vista '
            'por lote de antes, e os nomes reg* apontam para as colunas da tabela',
            cols and all(f.count('regLote,loteAtivo') == 2 for f in forms)
            and str(q.wb.Names('regCalib').RefersTo).replace(' ', '') == f'={TB}[Foicalibrado?]'.replace(' ', '')
            or (cols and all(f.count('regLote,loteAtivo') == 2 for f in forms) and TB in str(q.wb.Names('regCalib').RefersTo)),
            {'colunas': cols, 'formula': forms[:1], 'regCalib': str(q.wb.Names('regCalib').RefersTo)})

        # ---------------------------------------------------------------- R06 assinatura numa linha nova
        base = L.eng_por_run()
        pts = [d['id'][0] for d in base.values() if d['id'][0]]
        r6 = run('mCalibracao.RegistrarCalibracaoPorId', pts[len(pts) // 2], 0, teto=120)
        l6 = int(r6.split('|')[1]) if r6.startswith('OK|') else 0
        a6 = run('mCalibracao.AssinarRegistroCom', l6, USUARIO_QA, SENHA_QA, teto=120)
        h6 = txt(rw.Cells(l6, 8).Value) if l6 else ''
        ult6 = corpo()[-1]
        reg('R06 Calibração pela ID numa linha da tabela e assinatura com login e senha; a próxima linha continua pronta',
            r6.startswith('OK|') and a6.startswith('OK|') and f'login {USUARIO_QA}' in h6 and not any(txt(x) for x in ult6[1:12]),
            {'registro': r6, 'assinatura': a6, 'H': h6})

        # ---------------------------------------------------------------- R07 filtro x linha de entrada e assinaturas
        # (revisao de 30/10/2026: com qualquer filtro a linha vazia sumia -- a troca de lote sempre filtra; e a assinatura
        # de linha oculta era desenhada no tamanho original por cima da proxima linha visivel)
        r7 = lo.Range.Row + lo.Range.Rows.Count - 1
        rw.Cells(r7, 3).Value = an + ' '                # analito digitado com espaco sobrando
        rw.Cells(r7, 7).Value = 'QA ADR-087: analito com espaço no fim'

        def entrada():
            u = lo.Range.Row + lo.Range.Rows.Count - 1
            vazia = not any(txt(x) for x in rw.Range(rw.Cells(u, 2), rw.Cells(u, 12)).Value[0])
            return vazia and not rw.Rows(u).Hidden
        rw.Range('B4').Value = "'QA_LOTE_NOVO"          # outro lote: a linha assinada do R06 fica oculta
        ent1 = entrada()
        oculta6 = bool(rw.Rows(l6).Hidden) if l6 else False
        assin = [(str(s.Name), int(s.TopLeftCell.Row)) for s in rw.Shapes if str(s.Name).startswith('qcAssin_')]
        assin_ocultas = [a for a in assin if rw.Rows(a[1]).Hidden]
        run('mCalibracao.LimparFiltrosRegistros')
        rw.Range('C4').Value = an
        vis_esp = not rw.Rows(r7).Hidden
        ent2 = entrada()
        info7 = txt(rw.Range('K4').Value)
        run('mCalibracao.LimparFiltrosRegistros')
        reg('R07 Com filtro ativo a linha vazia de digitação continua visível; assinatura só nas linhas visíveis; o filtro '
            'de Analito acha o registro digitado com espaço sobrando',
            ent1 and oculta6 and not assin_ocultas and vis_esp and ent2,
            {'entrada_com_lote': ent1, 'linha_R06_oculta': oculta6, 'assinaturas': assin, 'em_linha_oculta': assin_ocultas,
             'registro_com_espaco_visivel': vis_esp, 'entrada_com_analito': ent2, 'info': info7})

        # ---------------------------------------------------------------- C01..C04 CEQ
        def campos(**kw):
            base_c = {'prov': 'CAP', 'rod': 'QA-087 Rodada', 'ano': '2026', 'an': 'QA Analito 087', 'am': 'QA-01',
                      'res': '12,5', 'alvo': '12', 'sd': '0,5', 'sdi': '1', 'linf': '10', 'lsup': '14',
                      'aval': 'Acceptable', 'nlabs': '100', 'unid': 'mg/dL', 'pag': '3', 'arq': 'qa_087.pdf'}
            base_c.update(kw)
            return [base_c[k] for k in ('prov', 'rod', 'ano', 'an', 'am', 'res', 'alvo', 'sd', 'sdi', 'linf', 'lsup',
                                        'aval', 'nlabs', 'unid', 'pag', 'arq')]

        def salvar(cs):
            return str(q.ex.xl.Run(f"'{q.wb.Name}'!mCEQForm.SalvarCEQ", cs))
        aud = L.audit_ult()
        c1 = salvar(campos())
        lo_cap = None
        for ws in q.wb.Worksheets:
            for t in ws.ListObjects:
                if t.Name == 'tblEQA_CAP_Dados':
                    lo_cap = t
        linha_cap = None
        if c1.startswith('OK|'):
            r = int(c1.split('|')[1])
            linha_cap = [txt(v) for v in lo_cap.Parent.Range(lo_cap.Parent.Cells(r, 1), lo_cap.Parent.Cells(r, 18)).Value[0]]
        base_eqa = q.ws('EQA_Base')
        chaves_b = [txt(r[0]) for r in base_eqa.Range('U2:U5001').Value]
        acoes = [a['acao'] for a in L.audit_desde(aud)]
        reg('C01 CEQ pela janela (SalvarCEQ): grava todos os campos na tabela do provedor, consolida a EQA_Base (a chave '
            'aparece) e registra CEQ_REGISTRADO',
            c1.startswith('OK|') and linha_cap and linha_cap[3] == 'QA Analito 087' and linha_cap[5] == '12.5'
            and linha_cap[11] == 'Acceptable' and 'CAP|2026|QA-087 Rodada|QA Analito 087|QA-01' in chaves_b
            and 'CEQ_REGISTRADO' in acoes,
            {'retorno': c1, 'linha': linha_cap, 'trilha': acoes})
        n_cap = lo_cap.ListRows.Count
        corpo_cap = [tuple(r) for r in lo_cap.DataBodyRange.Value]
        c2 = salvar(campos(res='13'))
        reg('C02 CEQ duplicado (mesma chave provedor|ano|rodada|analito|amostra) recusado; nada gravado',
            c2.startswith('ERRO|') and 'Já existe' in c2 and lo_cap.ListRows.Count == n_cap
            and [tuple(r) for r in lo_cap.DataBodyRange.Value] == corpo_cap, {'retorno': c2})
        c3 = {'sem_analito': salvar(campos(an='', am='QA-02')), 'numero': salvar(campos(am='QA-03', res='abc')),
              'limites': salvar(campos(am='QA-04', linf='20', lsup='10')), 'ano': salvar(campos(am='QA-05', ano='26')),
              'ano_virgula': salvar(campos(am='QA-06', ano='2,02')), 'ano_exp': salvar(campos(am='QA-07', ano='9e99'))}
        reg('C03 CEQ inválido recusado apontando o campo: sem analito, número inválido, limite inferior > superior, ano',
            c3['sem_analito'].startswith('ERRO|4|') and c3['numero'].startswith('ERRO|6|')
            and c3['limites'].startswith('ERRO|10|') and c3['ano'].startswith('ERRO|3|')
            and c3['ano_virgula'].startswith('ERRO|3|') and c3['ano_exp'].startswith('ERRO|3|')
            and [tuple(r) for r in lo_cap.DataBodyRange.Value] == corpo_cap, c3)
        ano = run('mCEQForm.ValorRegistrado', 'CAP', 2, 'QA-087 Rodada', 3)
        un = run('mCEQForm.ValorRegistrado', 'CAP', 4, 'QA Analito 087', 14)
        lista = q.ex.xl.Run(f"'{q.wb.Name}'!mCEQForm.RegistrosDaRodada", 'CAP', 'QA-087 Rodada', '2026')
        reg('C04 Reaproveitamento: rodada existente traz o ano; analito existente traz a unidade; registros da rodada',
            ano == '2026' and un == 'mg/dL' and lista and len(lista) == 1 and lista[0][0] == 'QA Analito 087',
            {'ano': ano, 'unidade': un, 'lista': lista})

        # ---------------------------------------------------------------- V01 janelas
        import instalar_frm_inativar as fi
        import ceq_form_spec as cq
        comp = q.wb.VBProject.VBComponents('frmCEQ')
        dif = fi._diferencas(comp, __import__('instalar_integracao').codigo(cq.FONTE_CODIGO), cq.FORM, cq.CONTROLES)
        aj = {'frmInativar': auditoria_janelas.auditar('frmInativar', fi.FORM, fi.CONTROLES),
              'frmCEQ': auditoria_janelas.auditar('frmCEQ', cq.FORM, cq.CONTROLES)}
        reg('V01 frmCEQ instalado igual à especificação; auditoria das janelas (texto medido com a fonte real, '
            'retângulos): nada cortado nem sobreposto',
            not dif and not any(aj.values()), {'diferencas_frmCEQ': dif, 'auditoria': aj})
    except Exception as e:                           # noqa: BLE001
        reg('R99 Execução da suíte sem exceção', False, {'erro': repr(e), 'trace': traceback.format_exc()[-2000:]})
    finally:
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('R00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-087 (tabela de registros, filtros, CEQ): {len(RES) - len(falhas)} PASS / '
          f'{len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
