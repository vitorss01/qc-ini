# -*- coding: utf-8 -*-
"""t10_historico.py -- TESTE 10: cinco anos de uso, ~100 mil resultados, 10 lotes.

Uso: python t10_historico.py <base.xlsm>

Monta, a partir da base limpa (Etapa 1 instalada), um historico de uso
prolongado e mede o que o usuario sente e o que pode quebrar com o volume:

  Lotes  L01..L09 : ~6 meses cada, 1 corrida por dia util (~125 corridas)
         L10      : 1 ano, 2 corridas por dia util (~500 corridas) -- passa do
                    teto de 180 pontos do grafico, de proposito
  31 analitos x 2 niveis, parametros proprios por lote e analito.

O banco e escrito em bloco (a importacao aceita 200 linhas por vez; 1.600
corridas por ela levariam horas) e as flags pelo mesmo AtualizarFlagsBanco
que a importacao usa. Os parametros entram pela aba Analitos, lote a lote,
pelo mesmo ParametroEditado do usuario.

Mede: abrir, recalculo completo, troca de lote, troca de analito, motor
completo, BI, importacao de UMA corrida nova (a operacao diaria), tamanho do
arquivo. Confere: lote do meio e lote longo com a mesma conferencia dos testes
T1-T9, a janela de 180 pontos, e a reconciliacao Excel x BI.
"""
import datetime as dt
import json
import os
import random
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import suite_etapa1 as S  # noqa: E402

T = 'T10 grande historico'


def gerar(analitos):
    rnd = random.Random(20260919)
    lotes = []
    d = dt.date(2026, 1, 5)
    for i in range(1, 11):
        nome = f'L{i:02d}'
        n_dias = 125 if i < 10 else 250
        por_dia = 1 if i < 10 else 2
        dias = S.dias_uteis(d, n_dias)
        d = dias[-1] + dt.timedelta(days=1)
        lotes.append((nome, dias, por_dia))
    params = {}
    for li, (nome, _, _) in enumerate(lotes):
        for ai, an in enumerate(analitos):
            base = 10.0 * (ai + 1)
            m1 = base * (1 + 0.03 * li)
            m2 = base * 3 * (1 + 0.02 * li)
            params[(nome, an)] = (round(m1, 4), round(m1 * 0.03, 5), round(m2, 4), round(m2 * 0.025, 5))
    linhas = []   # RUN, Data, Nivel, Lote, Analito, Resultado, Status
    run = 0
    for nome, dias, por_dia in lotes:
        for dd in dias:
            for k in range(por_dia):
                run += 1
                quando = dt.datetime(dd.year, dd.month, dd.day, 8 + 5 * k)
                for an in analitos:
                    m1, s1, m2, s2 = params[(nome, an)]
                    z1 = max(-2.6, min(2.6, rnd.gauss(0, 0.8)))
                    z2 = max(-2.6, min(2.6, rnd.gauss(0, 0.8)))
                    linhas.append((run, quando, 1, f'QC-{nome}01', an, round(m1 + z1 * s1, 5), 'Ativo'))
                    linhas.append((run, quando, 2, f'QC-{nome}02', an, round(m2 + z2 * s2, 5), 'Ativo'))
    return lotes, params, linhas


class Med(dict):
    def __setitem__(self, k, v):
        super().__setitem__(k, v)
        print(f'   MEDIDA {k:32s} {v}', flush=True)


def main(base):
    med = Med()
    s = S.Sessao(base, 'T10_monta')
    analitos = [a for a in s.analitos if a]
    lotes, params, linhas = gerar(analitos)
    print(f'{len(lotes)} lotes, {len(analitos)} analitos, {len(linhas)} resultados, '
          f'{linhas[-1][0]} corridas, {linhas[0][1]:%d/%m/%Y} a {linhas[-1][1]:%d/%m/%Y}', flush=True)
    try:
        s.cadastrar(*[l for l, _, _ in lotes])
        s.em_uso(lotes[-1][0])
        # parametros: pela aba Analitos, um lote por vez, a tela inteira de uma vez
        t0 = time.time()
        for nome, _, _ in lotes:
            s.selecionar(nome)
            bloco = [list(params[(nome, an)]) for an in analitos]
            s.ana.Range(f'E4:H{3 + len(analitos)}').Value = bloco
            s.run('mLotes.ParametroEditado', s.ana.Range(f'E4:H{3 + len(analitos)}'))
        med['digitar_parametros_10_lotes_s'] = round(time.time() - t0, 1)
        # banco em bloco
        db = s.sh['DB_Resultados']
        t0 = time.time()
        passo = 10000
        for i in range(0, len(linhas), passo):
            parte = linhas[i:i + passo]
            r0 = 4 + i
            db.Range(f'D{r0}:D{r0 + len(parte) - 1}').NumberFormat = '@'
            db.Range(f'A{r0}:G{r0 + len(parte) - 1}').Value = parte
        s.run('mBanco.AtualizarFlagsBanco', teto=1800)
        med['gravar_banco_s'] = round(time.time() - t0, 1)
        s.run('mDados.AtualizarListasAno', teto=600)
        s.wb.Save()
    finally:
        s.fechar()
    grande = s.caminho
    med['tamanho_arquivo_MB'] = round(os.path.getsize(grande) / 1e6, 2)

    # ---------------- sessao de uso ----------------
    t0 = time.time()
    s = S.Sessao(grande, 'T10_uso')
    med['abrir_s'] = round(time.time() - t0, 1)
    try:
        t0 = time.time(); s.xl.CalculateFull(); s.ex.esperar()
        med['recalculo_completo_s'] = round(time.time() - t0, 1)
        s.analito('Lactato')
        for nome in ('L05', 'L10'):
            t0 = time.time(); s.selecionar(nome)
            med[f'trocar_lote_{nome}_s'] = round(time.time() - t0, 1)
        t0 = time.time(); s.analito('Glicose')
        med['trocar_analito_s'] = round(time.time() - t0, 1)
        s.analito('Lactato')

        # --- lote do meio, sem filtro
        s.selecionar('L05')
        S.conferir(T, s, 'L05', 'Lactato', params[('L05', 'Lactato')], rotulo='[L05] ')
        # --- lote longo (500 corridas): sem filtro -> as 180 mais recentes, aviso de corte
        s.selecionar('L10')
        ok, L = S.conferir(T, s, 'L10', 'Lactato', params[('L10', 'Lactato')], rotulo='[L10 tudo] ')
        S.chk(T, '[L10 tudo] faixa O3 avisa corridas fora do grafico (500 > 180)',
              'fora do gráfico' in str(L['O3']) and int(L['eng1'][12] or 0) == 500 - 180,
              (L['eng1'][12], L['O3']))
        # --- ultimo mes do lote longo: tudo aparece, nada cortado
        fim = lotes[-1][1][-1]
        ini = dt.date(fim.year, fim.month, 1)
        s.filtro(dt.datetime(ini.year, ini.month, ini.day), dt.datetime(fim.year, fim.month, fim.day))
        ok, L = S.conferir(T, s, 'L10', 'Lactato', params[('L10', 'Lactato')], rotulo='[L10 ultimo mes] ')
        S.chk(T, '[L10 ultimo mes] nenhuma corrida do periodo cortada', int(L['eng1'][12] or 0) == 0, L['eng1'][12])
        # --- mes antigo do lote longo: janela termina no fim do periodo, nao no fim do lote
        ini2 = lotes[-1][1][0]
        fim2 = ini2 + dt.timedelta(days=30)
        s.filtro(dt.datetime(ini2.year, ini2.month, ini2.day), dt.datetime(fim2.year, fim2.month, fim2.day))
        S.conferir(T, s, 'L10', 'Lactato', params[('L10', 'Lactato')], rotulo='[L10 primeiro mes] ')
        # --- periodo sem nenhuma corrida do lote escolhido: aviso com os lotes que tem dado
        s.selecionar('L01')
        L = s.ler()
        S.chk(T, '[L01 num periodo de L10] aviso "nenhuma corrida" com os lotes que tem dado',
              'Nenhuma corrida deste lote' in str(L['O3']) and 'L10' in str(L['O3']), L['O3'])
        s.filtro(None, None)

        # --- operacao diaria: importar UMA corrida (31 analitos x 2 niveis)
        dia = fim + dt.timedelta(days=3)
        vals1 = {an: round(params[('L10', an)][0], 4) for an in analitos}
        vals2 = {an: round(params[('L10', an)][2], 4) for an in analitos}
        t0 = time.time()
        s.importar([(dia, 1, 'L10', vals1), (dia, 2, 'L10', vals2)])
        med['importar_1_corrida_s'] = round(time.time() - t0, 1)

        t0 = time.time(); s.run('mEstatistica.AtualizarEstatistica', teto=1800)
        med['motor_completo_s'] = round(time.time() - t0, 1)
        S.conferir_store(T, s, params)
        r = s.run('mBanco.TestarCapacidade', 1)
        S.chk(T, f'capacidade do banco ainda aceita gravacao: {r}', str(r).upper().startswith('PERMIT'), r)
        ult = s.sh['DB_Resultados'].Cells(s.sh['DB_Resultados'].Rows.Count, 1).End(-4162).Row
        med['linhas_banco'] = ult - 3
        med['linhas_livres_ate_teto'] = s.run('mBanco.LinhasLivres')
        # BI por ultimo: com 100 mil linhas pode nao terminar (e o watchdog mata o Excel)
        t0 = time.time()
        try:
            s.run('mBI.AtualizarBIData', teto=1200)
            med['bi_data_s'] = round(time.time() - t0, 1)
            s.selecionar('L10'); s.analito('Lactato')
            r = s.run('mBI.ReconciliarComCalc')
            comp, div = str(r).split('|')[:2]
            S.chk(T, f'[L10] sentinela Excel x BI com 100 mil linhas: {r}', int(comp) > 0 and int(div) == 0, r)
        except Exception as e:
            med['bi_data_s'] = f'> {round(time.time() - t0)} (nao terminou: {e.__class__.__name__})'
            S.chk(T, 'camada BI (mBI.AtualizarBIData) atualiza com 100 mil linhas em < 20 min', False, med['bi_data_s'])
    finally:
        s.fechar()
    print('\nMEDIDAS:')
    for k, v in med.items():
        print(f'   {k:32s} {v}')
    npass = sum(1 for r in S.RESULTADOS if r[2])
    print(f'\nTOTAL T10: {npass}/{len(S.RESULTADOS)} PASS')
    saida = os.path.join(S.AQUI, '..', 'resultados', f'T10_{time.strftime("%Y%m%d_%H%M%S")}.json')
    with open(saida, 'w', encoding='utf-8') as f:
        json.dump({'medidas': med, 'verificacoes': S.RESULTADOS}, f, ensure_ascii=False, indent=1)
    print('resultado:', os.path.abspath(saida))


if __name__ == '__main__':
    main(os.path.abspath(sys.argv[1]))
