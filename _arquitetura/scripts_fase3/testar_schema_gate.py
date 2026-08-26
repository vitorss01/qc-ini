# -*- coding: utf-8 -*-
"""testar_schema_gate.py - o schema gate reprova o que deve

POR QUE

testar_schema_bi.py devolveu 0 FAIL nos dois produtos. "0 FAIL" tem duas
explicacoes que se parecem exatamente iguais na tela: o schema esta certo, ou
o gate esta cego. Ja aconteceu de eu tomar cegueira por aprovacao duas vezes
nesta obra -- o analisador de VBA que nao via "Private gTrace As Object", e a
auditoria inteira rodando sobre modulos com o dobro de linhas.

COMO

Muta o CONTRATO (nao o artefato) e exige que o gate acuse. Mutar o contrato e
seguro e reversivel; mutar o .xlsm exigiria reconstruir o artefato a cada caso.
Como o gate compara contrato x aba, uma divergencia introduzida de qualquer um
dos lados tem de ser detectada.

CASOS

  A. campo do contrato que a aba nao tem .............. reprova
  B. coluna da aba que o contrato nao declara ......... reprova
  C. W_10x de volta na lista de campos ................ reprova
  D. tipo trocado (Sigma como texto) .................. reprova
  E. obrigatorio marcado onde ha vazio ................ reprova
  F. chave trocada por coluna nao unica ............... reprova
  G. ordem das regras divergindo da matriz do motor ... reprova
  H. contagem de colunas errada ....................... reprova
  I. contrato intacto ................................. aprova

Uso: python testar_schema_gate.py <arquivo.xlsm>
"""
import copy
import io
import json
import os
import subprocess
import sys
import tempfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)

AQUI = os.path.dirname(os.path.abspath(__file__))
GATE = os.path.join(AQUI, 'testar_schema_bi.py')
CONTRATO = os.path.join(os.path.dirname(AQUI), 'CONTRATO_BI.json')

CASOS = []


def caso(nome, espera_reprova=True):
    def deco(fn):
        CASOS.append((nome, espera_reprova, fn))
        return fn
    return deco


@caso('campo exigido que a aba nao tem')
def _a(c):
    c['campos'].append({'nome': 'Campo_Que_Nao_Existe', 'tipo': 'texto',
                        'obrigatorio': False, 'produto': 'ambos',
                        'origem': 'isca', 'descricao': 'isca'})
    c['n_colunas_por_produto'] += 1


@caso('coluna da aba que o contrato deixou de declarar')
def _b(c):
    c['campos'] = [x for x in c['campos'] if x['nome'] != 'Veredito']
    c['n_colunas_por_produto'] -= 1


@caso('W_10x de volta como campo declarado')
def _c(c):
    c['campos'].append({'nome': 'W_10x', 'tipo': 'inteiro',
                        'obrigatorio': False, 'produto': 'ambos',
                        'origem': 'isca', 'descricao': 'isca'})
    c['proibidos']['nomes'] = [n for n in c['proibidos']['nomes']
                               if n != 'W_10x']
    c['n_colunas_por_produto'] += 1


@caso('tipo trocado: Sigma declarado como texto')
def _d(c):
    for x in c['campos']:
        if x['nome'] == 'Sigma':
            x['tipo'] = 'logico'


@caso('obrigatorio marcado onde existe vazio')
def _e(c):
    for x in c['campos']:
        if x['nome'] == 'Sigma':
            x['obrigatorio'] = True


@caso('chave trocada por coluna nao unica')
def _f(c):
    c['chave'] = ['Analito']


@caso('ordem das regras divergindo do motor')
def _g(c):
    c['ordem_das_regras']['hematologia'] = ['1_3s', '2_2s', 'R_4s', '4_1s', '8x']
    c['ordem_das_regras']['bioquimica'] = ['1_3s', '2of3_2s', 'R_4s', '3_1s', '6x']


@caso('contagem de colunas errada')
def _h(c):
    c['n_colunas_por_produto'] = 99


@caso('contrato intacto', espera_reprova=False)
def _i(c):
    pass


def main():
    artefato = os.path.abspath(sys.argv[1])
    base = json.load(io.open(CONTRATO, encoding='utf-8'))
    print('=' * 78)
    print('O SCHEMA GATE REPROVA O QUE DEVE?  (%s)' % os.path.basename(artefato))
    print('=' * 78)
    print()

    falhas = 0
    for nome, espera, mutar in CASOS:
        c = copy.deepcopy(base)
        mutar(c)
        fd, tmp = tempfile.mkstemp(suffix='.json')
        os.close(fd)
        io.open(tmp, 'w', encoding='utf-8').write(
            json.dumps(c, ensure_ascii=False, indent=1))
        try:
            r = subprocess.run([sys.executable, '-u', GATE, artefato,
                                '--contrato', tmp], capture_output=True)
            reprovou = (r.returncode != 0)
        finally:
            os.remove(tmp)
        ok = (reprovou == espera)
        if not ok:
            falhas += 1
        print('   %-50s %-8s %s'
              % (nome[:50], 'reprova' if espera else 'aprova',
                 'PASS' if ok else 'FAIL'))
        if not ok:
            saida = (r.stdout or b'').decode('utf-8', 'replace')
            for l in [x for x in saida.split('\n') if 'FAIL' in x][:3]:
                print('        %s' % l.strip())

    print()
    print('TOTAL: %d PASS, %d FAIL' % (len(CASOS) - falhas, falhas))
    return 1 if falhas else 0


if __name__ == '__main__':
    sys.exit(main())
