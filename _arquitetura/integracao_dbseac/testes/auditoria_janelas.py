# -*- coding: utf-8 -*-
"""ADR-087: auditoria visual das janelas declarativas (sem Excel) -- nenhum texto cortado, nenhum controle sobreposto.

Para cada janela (frmInativar, frmCEQ, frmCarregando): os retangulos dos controles nao se cruzam (fora o que e
desenhado de proposito por cima, como a barra sobre o trilho do indicador), todos cabem na area util da janela, e cada
texto -- os rotulos fixos E os textos que o codigo poe em tempo de execucao (os mais longos) -- cabe no seu controle,
medido com a fonte real (Segoe UI do Windows, a 96 dpi, como o MSForms desenha). Rotulo com quebra de linha: as linhas
necessarias cabem na altura. Sai com 1 se achar problema.

Uso: python auditoria_janelas.py [saida.json]
"""
import json
import os
import sys

from PIL import ImageFont

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..'))
sys.path.insert(0, os.path.join(AQUI, '..', '..', 'etapa1_multilote'))
import instalar_frm_inativar as fi  # noqa: E402
import ceq_form_spec as cq  # noqa: E402
import instalar_frm_carregando as fc  # noqa: E402

FONTES = {False: r'C:\Windows\Fonts\segoeui.ttf', True: r'C:\Windows\Fonts\segoeuib.ttf'}
DPI = 96
BORDA_X, BORDA_Y = 6, 26           # pt aproximados da moldura e da barra de titulo do UserForm
PAD = {'Forms.Label.1': 2, 'Forms.CommandButton.1': 8, 'Forms.OptionButton.1': 18, 'Forms.CheckBox.1': 18,
       'Forms.TextBox.1': 6, 'Forms.ComboBox.1': 22}
_cache = {}

# textos de tempo de execucao (os mais longos que o codigo escreve), por janela e controle
DINAMICOS = {
    'frmInativar': {
        'lblTitulo': ['Resultado ID 1234567 — NÃO CONFORME (X vermelho)'],
        'lblDetalhe': ['ID 1234567 · Hemoglobina corpuscular média · DIMENSION 2 · Nível 3 · Lote 8977\n'
                       'Entre as corridas 120 e 121 (RUN 26102803) · Data 28/10/2026 · Hora 14:32 · Valor 123,45\n'
                       'Gráfico N3 · série "Violação" · Westgard: 1-3s + 2de3-2s + R-4s\n'
                       'Estado atual: NÃO CONFORME (inativado, X vermelho) — fora da média, do DP e do Westgard'],
        'optNC': ['A · MARCAR COMO NÃO CONFORME — continua visível no gráfico como X vermelho',
                  'REATIVAR — o resultado volta a ser ativo (média, DP e Westgard)'],
        'optInat': ['B · INATIVAR — sai do gráfico; o registro original fica preservado',
                    'INATIVAR — tirar o X do gráfico (continua inativado; o registro fica)'],
        'optComent': ['Só registrar um comentário técnico — o resultado continua ativo',
                      'Só atualizar o comentário técnico (continua não conforme)'],
        'lblEfeito': ['Some do gráfico e sai da média, do DP e do Westgard. Nada é apagado: ID, valor, data, motivo e '
                      'trilha ficam. Para reativar depois: aba Inativar.',
                      'O manual sai da média, do DP e do Westgard (não é apagado), mas NÃO traz o automático 1234567 '
                      'de volta: para isso, apague o SUBSTITUI_ID.'],
        'lblMotivo': ['Comentário técnico — texto livre (obrigatório)'],
        'lblAcao': ['O que fazer com este resultado? (escolha uma opção)'],
    },
    'frmCEQ': {
        'lblMsg': ['Já existe este registro (Controllab Rodada 3 2026 · Hemoglobina glicada (HbA1c) · CHM-05) na linha '
                   '2345 da aba EQA.Controllab_Dados. Para corrigir, edite a linha existente.',
                   '✓ Gravado: Hemoglobina glicada (HbA1c) · amostra CHM-05 (linha 2345).'],
        'lblQtd': ['Escolha a rodada para ver o que já foi registrado.'],
    },
}


def fonte(bold, size_pt):
    k = (bold, size_pt)
    if k not in _cache:
        _cache[k] = ImageFont.truetype(FONTES[bold], round(size_pt * DPI / 72))
    return _cache[k]


def largura_pt(texto, bold, size):
    return fonte(bold, size).getlength(texto) * 72 / DPI


def linhas_necessarias(texto, largura, bold, size):
    total = 0
    for par in texto.split('\n'):
        palavras, atual, n = par.split(' '), '', 1
        for p in palavras:
            cand = (atual + ' ' + p).strip()
            if largura_pt(cand, bold, size) <= largura or not atual:
                atual = cand
            else:
                n += 1
                atual = p
        total += n
    return total


def cruzam(a, b):
    return a[0] < b[0] + b[2] - 0.5 and b[0] < a[0] + a[2] - 0.5 and a[1] < b[1] + b[3] - 0.5 and b[1] < a[1] + a[3] - 0.5


def auditar(nome, form, controles, permitidos=()):
    probs = []
    util_w, util_h = form['Width'] - BORDA_X, form['Height'] - BORDA_Y
    rets = {}
    for progid, n, p in controles:
        r = (p['Left'], p['Top'], p['Width'], p['Height'])
        rets[n] = (progid, r, p)
        if r[0] < 0 or r[1] < 0 or r[0] + r[2] > util_w + 0.5 or r[1] + r[3] > util_h + 0.5:
            probs.append(f'{n}: fora da area util da janela ({r} em {util_w}x{util_h})')
    nomes = list(rets)
    for i, a in enumerate(nomes):
        for b in nomes[i + 1:]:
            if (a, b) in permitidos or (b, a) in permitidos:
                continue
            if cruzam(rets[a][1], rets[b][1]):
                probs.append(f'{a} x {b}: sobrepostos {rets[a][1]} / {rets[b][1]}')
    for n, (progid, r, p) in rets.items():
        bold, size = bool(p.get('Font.Bold')), float(p.get('Font.Size', 9))
        textos = [p['Caption']] if p.get('Caption') else []
        textos += DINAMICOS.get(nome, {}).get(n, [])
        livre = r[2] - PAD.get(progid, 4)
        altura_linha = size * 1.35
        for t in textos:
            if p.get('WordWrap') or progid in ('Forms.OptionButton.1',):
                nl = linhas_necessarias(t, livre, bold, size)
                if nl * altura_linha > r[3] + 1:
                    probs.append(f'{n}: texto precisa de {nl} linha(s) ({nl * altura_linha:.0f} pt) e o controle tem '
                                 f'{r[3]} pt: "{t[:60]}..."')
            else:
                w = max(largura_pt(x, bold, size) for x in t.split('\n'))
                if w > livre + 0.5:
                    probs.append(f'{n}: texto de {w:.0f} pt num espaco de {livre:.0f} pt (cortado): "{t[:60]}"')
                if t.count('\n') + 1 > max(1, int((r[3] + 1) // altura_linha)):
                    probs.append(f'{n}: {t.count(chr(10)) + 1} linhas numa altura de {r[3]} pt')
    return probs


def main(saida=None):
    res = {
        'frmInativar': auditar('frmInativar', fi.FORM, fi.CONTROLES),
        'frmCEQ': auditar('frmCEQ', cq.FORM, cq.CONTROLES),
        'frmCarregando': auditar('frmCarregando', fc.FORM, fc.CONTROLES,
                                 permitidos=getattr(fc, 'SOBREPOSTOS_DE_PROPOSITO', ())),
    }
    for k, v in res.items():
        print(f'{k}: {"OK" if not v else str(len(v)) + " problema(s)"}')
        for x in v:
            print('   -', x)
    if saida:
        with open(saida, 'w', encoding='utf-8') as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
    return 1 if any(res.values()) else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else None))
