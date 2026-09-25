# -*- coding: utf-8 -*-
"""comparar_fotos.py -- o que mudou para o usuario entre antes e depois da Etapa 1.

Uso: python comparar_fotos.py <antes.json> <depois.json>

Por analito: corridas plotadas (e ordem), valores, z, veredictos, linhas de
limite, indicadores do Painel e da Estatistica. Diferenca esperada e
EXPLICADA (ordem cronologica; nenhuma outra); o resto tem de ser igual.
"""
import json
import sys


def carregar(p):
    with open(p, encoding='utf-8') as f:
        return json.load(f)


def linhas_calc(f):
    """(RUN, data, N1 val, N1 z, N1 veredicto, N2 val, N2 z, N2 veredicto) por linha com corrida."""
    out = []
    for a, b in zip(f['calc_B_AB'], f['calc_AC_AW']):
        if a[0] in (None, ''):
            continue
        out.append({'run': a[0], 'data': a[1], 'v1': a[4], 'z1': a[5], 'st1': a[14],
                    'lim1': (a[15], a[18], a[21]), 'v2': a[26], 'z2': b[0], 'st2': b[9],
                    'lim2': (b[10], b[13], b[16]), 'flags1': tuple(a[9:14]), 'flags2': tuple(b[4:9])})
    return out


def main(pa, pd):
    A, D = carregar(pa), carregar(pd)
    for an in A:
        if an not in D:
            print(an, 'ausente depois'); continue
        la, ld = linhas_calc(A[an]), linhas_calc(D[an])
        ra = {l['run']: l for l in la}
        rd = {l['run']: l for l in ld}
        print(f'\n== {an}: corridas antes {len(la)} / depois {len(ld)}')
        print('   mesmo conjunto de corridas:', set(ra) == set(rd),
              '| so antes:', sorted(set(ra) - set(rd))[:6], '| so depois:', sorted(set(rd) - set(ra))[:6])
        ordem_a = [l['run'] for l in la]
        ordem_d = [l['run'] for l in ld]
        print('   mesma ordem no eixo:', ordem_a == ordem_d)
        datas_d = [l['data'] for l in ld]
        print('   ordem depois e cronologica:', datas_d == sorted(datas_d))
        dif_v = [r for r in set(ra) & set(rd) if (ra[r]['v1'], ra[r]['v2']) != (rd[r]['v1'], rd[r]['v2'])]
        dif_z = [r for r in set(ra) & set(rd) if (ra[r]['z1'], ra[r]['z2']) != (rd[r]['z1'], rd[r]['z2'])]
        dif_l = [r for r in set(ra) & set(rd) if (ra[r]['lim1'], ra[r]['lim2']) != (rd[r]['lim1'], rd[r]['lim2'])]
        dif_s = [r for r in set(ra) & set(rd) if (ra[r]['st1'], ra[r]['st2'], ra[r]['flags1'], ra[r]['flags2'])
                 != (rd[r]['st1'], rd[r]['st2'], rd[r]['flags1'], rd[r]['flags2'])]
        print(f'   valores diferentes: {len(dif_v)} | z diferentes: {len(dif_z)} | limites diferentes: {len(dif_l)}')
        print(f'   veredicto/flags diferentes: {len(dif_s)}')
        for r in sorted(dif_s, key=lambda x: str(rd[x]['data']))[:8]:
            print(f'      RUN {r} {rd[r]["data"]}: antes {ra[r]["st1"]}/{ra[r]["st2"]} {ra[r]["flags1"]}{ra[r]["flags2"]}'
                  f' -> depois {rd[r]["st1"]}/{rd[r]["st2"]} {rd[r]["flags1"]}{rd[r]["flags2"]}')
        print('   parametros AX1:BA1 antes', A[an]['calc_AX_BG1'][0][:4], 'depois', D[an]['calc_AX_BG1'][0][:4])
        print('   Painel A7:M8 igual:', A[an]['painel_A6_M8'][1:] == D[an]['painel_A6_M8'][1:])
        if A[an]['painel_A6_M8'][1:] != D[an]['painel_A6_M8'][1:]:
            print('      antes ', A[an]['painel_A6_M8'][1:])
            print('      depois', D[an]['painel_A6_M8'][1:])
        print('   Estatistica igual:', A[an]['estat'] == D[an]['estat'])
        if A[an]['estat'] != D[an]['estat']:
            print('      antes ', A[an]['estat'])
            print('      depois', D[an]['estat'])
        print('   eventos antes/depois:', A[an]['eventos_total'], D[an]['eventos_total'])


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
