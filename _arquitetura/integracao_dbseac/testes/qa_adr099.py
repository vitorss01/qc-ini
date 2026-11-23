# -*- coding: utf-8 -*-
"""QA do ADR-099 (22/11/2026): coluna MATERIAL no cadastro de lotes (Configuração!F25:F125), entrada amigavel da regra
MATRIZ da tblLotesMaterial (ADR-091). Roda numa COPIA instalada (instalar_adr099.py); nada e salvo.

Uso: python qa_adr099.py <Bioquimica|Hematologia> <copia_instalada.xlsm> <saida.json>

  M01 F25 "Material" travada; F26:F125 livres, com a lista das matrizes do de/para (MatrizesPermitidas); C26/D25 travadas
  M02 F = a 1a regra MATRIZ de 100 % dos lotes do cadastro (carga inicial = espelho; nenhuma regra nova)
  M04 digitar o material de um lote SEM regra (eventos ligados, sem dialogo): cria 1 regra MATRIZ (LOTE texto, NOTA
      ADR-099) e 1 LOTE_MATERIAL_ALTERADO; depois do ATUALIZAR so das entradas, todas as linhas do lote ficam com essa
      MATRIZ e nada fora do lote muda
  M05 apagar o material: as regras MATRIZ do lote saem; o lote volta exatamente ao de antes (STATUS/PARTICIPA/MATRIZ/
      ANALITO/RUN); SO_ANALITOS e FORA_DO_CQ intactas
  M06 colar um material fora da lista ('LIQUOR', 'xyz'): volta ao que era; a tabela nao muda
  M07 editar a tblLotesMaterial direto (regra MATRIZ num lote) e ATUALIZAR: a coluna F do lote passa a mostrar a regra
      (espelho no ATUALIZAR) com LOTE_MATERIAL_SINCRONIZADO
  M09 PreviaMaterial = varredura independente da tblCQ_Final
  Y00 seguranca do teste: arquivo de entrada intacto
"""
import collections
import json
import os
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from qa_final import QA, reg, RES, sha256_arquivo  # noqa: E402
sys.path.insert(0, os.path.join(AQUI, '..'))
import instalar_integracao as ii  # noqa: E402

CFG = 'Configuração'
R0, R1 = 26, 125
CAMPOS = ('STATUS_ANALITICO', 'PARTICIPA_ESTATISTICA', 'MATRIZ', 'ANALITO', 'RUN')


def txt(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return '' if v is None else str(v).strip()


def up(v):
    return txt(v).upper()


def executar(produto, caminho, saida):
    hash0 = sha256_arquivo(caminho)
    q = QA(produto, caminho)
    try:
        xl = q.ex.xl
        ws = q.ws(CFG)
        nome = q.wb.Name

        def xrun(m, *a):
            return xl.Run(f"'{nome}'!{m}", *a)

        def regras():
            lo = q.lo('tblLotesMaterial')
            return [tuple(txt(c) for c in l) for l in (lo.DataBodyRange.Value if lo.ListRows.Count else ())]

        def cad():
            return [(r, txt(ws.Cells(r, 3).Value), up(ws.Cells(r, 6).Value)) for r in range(R0, R1 + 1)
                    if txt(ws.Cells(r, 3).Value)]

        def foto_lote(fin, lote):
            return {txt(x['ID_REGISTRO']): tuple(txt(x[c]) for c in CAMPOS) for x in fin if up(x['LOTE']) == up(lote)}

        def foto_fora(fin, lote):
            return {txt(x['ID_REGISTRO']): tuple(txt(x[c]) for c in CAMPOS) for x in fin if up(x['LOTE']) != up(lote)}

        def aud0():
            return int(xrun('mAuditoria.UltimaLinhaAudit'))

        def acoes(r0):
            al = q.ws('Audit_Log')
            r1 = aud0()
            if r1 <= r0:
                return []
            v = al.Range(al.Cells(r0 + 1, 1), al.Cells(r1, 15)).Value
            return [(txt(l[6]), txt(l[10]), txt(l[11]), txt(l[12])) for l in v]

        def editar(cel, valor):
            xl.EnableEvents = True
            cel.Value = valor
            xl.EnableEvents = False

        xl.Interactive = False                     # o handler aplica direto (sem MsgBox) sob automacao
        lista = txt(xrun('mMaterialLote.MatrizesPermitidas'))
        # ---------------------------------------------------------------- M01
        f1 = ''
        try:
            f1 = txt(ws.Range('F26').Validation.Formula1)
        except Exception:                          # noqa: BLE001
            pass
        sep = xl.International[4]
        reg('M01 F25 "Material" travada; F26:F125 livres com a lista do de/para (%s); C26 travada' % lista,
            txt(ws.Range('F25').Value) == 'Material' and ws.Range('F25').Locked is True and ws.Range('F26').Locked is False
            and ws.Range('F125').Locked is False and ws.Range('C26').Locked is True and bool(lista)
            and f1.split(sep) == lista.split(';'),
            {'F25': ws.Range('F25').Value, 'lista': lista, 'validacao': f1})
        # ---------------------------------------------------------------- M02
        tab0 = regras()
        primeira = {}
        for l in tab0:
            if up(l[1]) == 'MATRIZ' and up(l[0]) and up(l[2]) and up(l[0]) not in primeira:
                primeira[up(l[0])] = up(l[2])
        dif = [(r, lt, f, primeira.get(up(lt), '')) for r, lt, f in cad() if f != primeira.get(up(lt), '')]
        reg('M02 coluna Material = 1a regra MATRIZ de cada lote do cadastro (%d lotes, %d com regra)'
            % (len(cad()), sum(1 for _, lt, _ in cad() if up(lt) in primeira)), not dif and bool(cad()),
            {'diferentes': dif[:10]})
        # ---------------------------------------------------------------- escolha do lote do teste
        # a referencia e a tblCQ_Final DEPOIS de um ATUALIZAR so das entradas: uma copia em que o ATUALIZAR ainda nao
        # aplicou as regras ja existentes (ex.: montada sem o --atualizar do ADR-091) mudaria lotes alheios no 1o ATUALIZAR
        q.run('mIntegracao.AtualizarSoEntradas', '', teto=1800)
        fin0 = q.ler('tblCQ_Final')
        por_lote = collections.Counter(up(x['LOTE']) for x in fin0)
        sem_regra = [(r, lt) for r, lt, f in cad() if up(lt) not in primeira and por_lote.get(up(lt), 0) > 0
                     and not any(up(l[0]) == up(lt) for l in tab0)]
        if not sem_regra:
            reg('M04 pre-condicao ausente: lote do cadastro sem regra e com resultados', False, {'cadastro': cad()[:10]})
        else:
            r, lt = min(sem_regra, key=lambda x: por_lote[up(x[1])])         # o menor lote: o efeito fica localizado
            mz_atual = collections.Counter(up(x['MATRIZ']) for x in fin0 if up(x['LOTE']) == up(lt))
            opc = lista.split(';')
            novo = next((m for m in opc if m not in mz_atual), opc[0])
            # M09 previa
            prev = txt(xrun('mMaterialLote.PreviaMaterial', lt, novo)).split('|')
            n_l = sum(1 for x in fin0 if up(x['LOTE']) == up(lt))
            n_mu = sum(1 for x in fin0 if up(x['LOTE']) == up(lt) and up(x['MATRIZ']) != novo)
            n_mp = sum(1 for x in fin0 if up(x['LOTE']) == up(lt) and up(x['MATRIZ']) != novo
                       and up(x['PARTICIPA_ESTATISTICA']) == 'SIM')
            reg('M09 PreviaMaterial = varredura independente da tblCQ_Final', prev == [str(n_l), str(n_mu), str(n_mp)],
                {'lote': lt, 'material': novo, 'previa': prev, 'esperado': [n_l, n_mu, n_mp]})
            base_lote, base_fora = foto_lote(fin0, lt), foto_fora(fin0, lt)
            a0 = aud0()
            editar(ws.Cells(r, 6), novo.lower())
            tab1 = regras()
            novas = [l for l in tab1 if l not in tab0]
            tr = [a for a in acoes(a0) if a[0] == 'LOTE_MATERIAL_ALTERADO']
            r_at = txt(q.run('mIntegracao.AtualizarSoEntradas', '', teto=1800))
            fin1 = q.ler('tblCQ_Final')
            lote1 = foto_lote(fin1, lt)
            mz1 = {up(x['MATRIZ']) for x in fin1 if up(x['LOTE']) == up(lt)}
            fora_mudou = [(k, dict(zip(CAMPOS, base_fora.get(k) or ())), dict(zip(CAMPOS, v)))
                          for k, v in foto_fora(fin1, lt).items() if base_fora.get(k) != v]
            reg('M04 material "%s" no lote %s (sem regra): 1 regra MATRIZ nova (LOTE texto, NOTA ADR-099), 1 trilha, '
                'o lote inteiro com essa MATRIZ depois do ATUALIZAR e nada fora do lote muda' % (novo, lt),
                r_at.startswith('OK|') and len(novas) == 1 and up(novas[0][0]) == up(lt) and up(novas[0][1]) == 'MATRIZ'
                and up(novas[0][2]) == novo and novas[0][4].startswith('ADR-099') and len(tab1) == len(tab0) + 1
                and len(tr) == 1 and up(ws.Cells(r, 6).Value) == novo and mz1 == {novo} and not fora_mudou,
                {'regra_nova': novas, 'trilha': tr, 'F': ws.Cells(r, 6).Value, 'matrizes_do_lote': sorted(mz1),
                 'linhas_do_lote': len(lote1), 'fora_do_lote_mudou': fora_mudou[:5], 'atualizar': r_at[:100]})
            # ---------------------------------------------------------------- M05
            a1 = aud0()
            editar(ws.Cells(r, 6), None)
            tab2 = regras()
            r_at2 = txt(q.run('mIntegracao.AtualizarSoEntradas', '', teto=1800))
            fin2 = q.ler('tblCQ_Final')
            volta = foto_lote(fin2, lt)
            difv = [k for k in set(volta) | set(base_lote) if volta.get(k) != base_lote.get(k)]
            tr2 = [a for a in acoes(a1) if a[0] == 'LOTE_MATERIAL_ALTERADO']
            reg('M05 apagar o material: a regra sai, as SO_ANALITOS/FORA_DO_CQ ficam e o lote volta exatamente ao de antes',
                r_at2.startswith('OK|') and tab2 == tab0 and not difv and len(tr2) == 1 and not txt(ws.Cells(r, 6).Value),
                {'tabela_igual': tab2 == tab0, 'linhas_diferentes': difv[:5], 'trilha': tr2})
            # ---------------------------------------------------------------- M06
            res6 = []
            for ruim in ('LIQUOR', 'xyz'):
                if ruim.upper() in lista.split(';'):
                    continue
                editar(ws.Cells(r, 6), ruim)
                res6.append((ruim, ws.Cells(r, 6).Value))
            reg('M06 material fora da lista (colado por cima da validacao) e desfeito; a tabela nao muda',
                all(v in (None, '') for _, v in res6) and regras() == tab0 and bool(res6), {'tentativas': res6})
            # ---------------------------------------------------------------- M07
            lo = q.lo('tblLotesMaterial')
            prot = lo.Parent.ProtectContents
            lo.Parent.Unprotect(ii.SENHA)
            lr = lo.ListRows.Add()
            lr.Range.NumberFormat = '@'
            cols = [str(c.Name) for c in lo.ListColumns]
            for c_, v_ in (('LOTE', lt), ('REGRA', 'MATRIZ'), ('MATRIZ', novo), ('NOTA', 'TESTE-ADR099 M07')):
                lr.Range.Cells(1, cols.index(c_) + 1).Value = v_
            a2 = aud0()
            r_at3 = txt(q.run('mIntegracao.AtualizarSoEntradas', '', teto=1800))
            tr3 = [a for a in acoes(a2) if a[0] == 'LOTE_MATERIAL_SINCRONIZADO']
            reg('M07 regra gravada direto na tabela: o ATUALIZAR refaz a coluna Material do lote (espelho) e audita',
                r_at3.startswith('OK|') and up(ws.Cells(r, 6).Value) == novo and len(tr3) >= 1,
                {'F': ws.Cells(r, 6).Value, 'trilha': tr3[:3], 'protegida_antes': prot})
    except Exception as ex_:                         # noqa: BLE001
        reg('Z99 Execução da suíte sem exceção', False, {'erro': repr(ex_), 'trace': traceback.format_exc()[-2000:]})
    finally:
        try:
            q.ex.xl.Interactive = True
        except Exception:                            # noqa: BLE001
            pass
        q.fechar(salvar=False)
    hash1 = sha256_arquivo(caminho)
    reg('Y00 Segurança do teste: arquivo de entrada intacto (SHA-256 igual antes e depois; nada salvo nele)',
        hash0 == hash1, {'arquivo': caminho, 'sha256_antes': hash0, 'sha256_depois': hash1})
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(RES, f, ensure_ascii=False, indent=1, default=str)
    falhas = [r for r in RES if r['resultado'] == 'FAIL']
    print(f'\n=== {produto} ADR-099: {len(RES) - len(falhas)} PASS / {len(falhas)} FAIL ===', flush=True)
    return falhas


if __name__ == '__main__':
    falhas = executar(sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3] if len(sys.argv) > 3 else None)
    sys.exit(1 if falhas else 0)
