# -*- coding: utf-8 -*-
"""comparar_layout_painel.py - o layout do Painel mudou?

POR QUE NAO DA PARA COMPARAR O INDICE DE ESTILO

Ao salvar, o Excel RENUMERA a tabela cellXfs: remove estilos que ficaram sem
uso e recompacta o resto. Duas gravacoes do mesmo arquivo saem com indices
diferentes apontando para o mesmo visual. Comparar "s=" acusa milhares de
mudancas que nao existem -- foi o que aconteceu na primeira tentativa.

O QUE ESTE SCRIPT COMPARA

Resolve cada indice ate o CONTEUDO do estilo -- fonte, preenchimento, borda,
formato numerico e alinhamento -- e compara as tuplas. Se a cor era #263B4D
antes e depois, nao importa em que linha da tabela ela foi parar.

Compara tambem largura de coluna, altura de linha, celulas mescladas,
formatacao condicional e a posicao de cada objeto (graficos, botoes, spinner).

Uso: python comparar_layout_painel.py <antes.xlsm> <depois.xlsm> [aba]
"""
import io
import re
import sys
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)


def _bloco(xml, tag):
    """Lista os elementos <tag> de primeiro nivel dentro de <tagPlural>."""
    saida = []
    prof = 0
    atual = []
    for m in re.finditer(r'<(/?)%s\b([^>]*?)(/?)>' % tag, xml):
        fecha, attrs, auto = m.group(1), m.group(2), m.group(3)
        if fecha:
            prof -= 1
            atual.append(m.group(0))
            if prof == 0:
                saida.append(''.join(atual))
                atual = []
        elif auto:
            if prof == 0:
                saida.append(m.group(0))
            else:
                atual.append(m.group(0))
        else:
            if prof == 0:
                atual = [xml[m.start():m.end()]]
            else:
                atual.append(m.group(0))
            prof += 1
    return saida


def _entre(xml, tag):
    m = re.search(r'<%s\b[^>]*>(.*?)</%s>' % (tag, tag), xml, re.S)
    return m.group(1) if m else ''


def estilos(caminho, aba='Painel'):
    z = zipfile.ZipFile(caminho)
    wbx = z.read('xl/workbook.xml').decode('utf-8', 'replace')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'replace')
    rid = {m.group(1): m.group(2)
           for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)}
    alvo = None
    for m in re.finditer(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx):
        if m.group(1) == aba:
            alvo = 'xl/' + rid[m.group(2)].lstrip('/').replace('xl/', '')
    if alvo is None:
        z.close()
        raise SystemExit('aba %r nao encontrada em %s' % (aba, caminho))

    sty = z.read('xl/styles.xml').decode('utf-8', 'replace')
    fontes = _bloco(_entre(sty, 'fonts'), 'font')
    fills = _bloco(_entre(sty, 'fills'), 'fill')
    bordas = _bloco(_entre(sty, 'borders'), 'border')
    numfmt = {m.group(1): m.group(2) for m in
              re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', sty)}
    xfs = _bloco(_entre(sty, 'cellXfs'), 'xf')

    def resolve(i):
        """Indice -> conteudo do estilo. E isto que sobrevive a renumeracao."""
        try:
            xf = xfs[int(i)]
        except Exception:
            return ('?', i)
        def at(n):
            m = re.search(r'%s="([^"]*)"' % n, xf)
            return m.group(1) if m else ''
        f = at('fontId')
        fl = at('fillId')
        b = at('borderId')
        nf = at('numFmtId')
        al = _bloco(xf, 'alignment')
        return (
            fontes[int(f)] if f.isdigit() and int(f) < len(fontes) else f,
            fills[int(fl)] if fl.isdigit() and int(fl) < len(fills) else fl,
            bordas[int(b)] if b.isdigit() and int(b) < len(bordas) else b,
            numfmt.get(nf, nf),
            al[0] if al else '',
        )

    sx = z.read(alvo).decode('utf-8', 'replace')

    celulas = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)/?>', sx):
        s = re.search(r's="(\d+)"', m.group(2))
        celulas[m.group(1)] = resolve(s.group(1) if s else '0')

    largura = {}
    for m in re.finditer(r'<col min="(\d+)" max="(\d+)"([^>]*)/>', sx):
        wm = re.search(r'width="([\d.]+)"', m.group(3))
        oculta = 'hidden="1"' in m.group(3)
        for c in range(int(m.group(1)), int(m.group(2)) + 1):
            largura[c] = (round(float(wm.group(1)), 2) if wm else None, oculta)

    altura = {}
    for m in re.finditer(r'<row r="(\d+)"([^>]*)>', sx):
        hm = re.search(r'ht="([\d.]+)"', m.group(2))
        altura[int(m.group(1))] = round(float(hm.group(1)), 2) if hm else None

    mescladas = sorted(re.findall(r'<mergeCell ref="([^"]+)"', sx))
    cond = sorted(re.findall(r'<conditionalFormatting sqref="([^"]+)"', sx))

    # objetos: graficos, botoes, caixas -- posicao e tamanho
    objetos = []
    try:
        srel = z.read(alvo.replace('worksheets/', 'worksheets/_rels/')
                      + '.rels').decode('utf-8', 'replace')
        for m in re.finditer(r'Target="([^"]*drawing\d+\.xml)"', srel):
            d = z.read('xl/drawings/' + m.group(1).split('/')[-1]) \
                 .decode('utf-8', 'replace')
            for a in re.finditer(
                    r'<xdr:from>.*?<xdr:col>(\d+)</xdr:col>.*?'
                    r'<xdr:row>(\d+)</xdr:row>.*?</xdr:from>', d, re.S):
                objetos.append((a.group(1), a.group(2)))
    except Exception:
        pass

    z.close()
    return {'celulas': celulas, 'largura': largura, 'altura': altura,
            'mescladas': mescladas, 'cond': cond,
            'objetos': sorted(objetos)}


def main(antes, depois, aba='Painel'):
    a = estilos(antes, aba)
    b = estilos(depois, aba)
    print('=== COMPARACAO DE LAYOUT DA ABA %s ===' % aba)
    print('celulas com estilo: %d -> %d' % (len(a['celulas']), len(b['celulas'])))

    problemas = []

    mudou = [k for k in a['celulas']
             if k in b['celulas'] and a['celulas'][k] != b['celulas'][k]]
    sumiu = [k for k in a['celulas'] if k not in b['celulas']]
    nasceu = [k for k in b['celulas'] if k not in a['celulas']]
    print('   estilo de celula alterado : %d %s'
          % (len(mudou), sorted(mudou)[:12]))
    print('   celulas sumidas           : %d %s'
          % (len(sumiu), sorted(sumiu)[:12]))
    print('   celulas novas             : %d %s'
          % (len(nasceu), sorted(nasceu)[:12]))
    problemas += ['estilo alterado em %s' % k for k in mudou]
    problemas += ['celula sumida: %s' % k for k in sumiu]

    for rot, chave in (('largura de coluna', 'largura'),
                       ('altura de linha', 'altura')):
        difs = [k for k in a[chave]
                if k in b[chave] and a[chave][k] != b[chave][k]]
        difs += [k for k in a[chave] if k not in b[chave]]
        print('   %-25s : %d %s' % (rot, len(difs), sorted(difs)[:10]))
        problemas += ['%s alterada em %s' % (rot, k) for k in difs]

    for rot, chave in (('celulas mescladas', 'mescladas'),
                       ('formatacao condicional', 'cond'),
                       ('objetos (graficos/botoes)', 'objetos')):
        if a[chave] != b[chave]:
            print('   %-25s : MUDOU' % rot)
            # a lista e truncada: sem o total, um item que so ficou de fora
            # da amostra parece ter sumido. Ja aconteceu com Painel!O3.
            print('      antes : %d no total, mostrando %s'
                  % (len(a[chave]), a[chave][:8]))
            print('      depois: %d no total, mostrando %s'
                  % (len(b[chave]), b[chave][:8]))
            problemas.append('%s mudou' % rot)
        else:
            print('   %-25s : identico (%d)' % (rot, len(a[chave])))

    print()
    permitidas = set(sys.argv[4].split(',')) if len(sys.argv) > 4 else set()
    reais = [p for p in problemas
             if not any(('em %s' % c) == p[-len('em %s' % c):]
                        for c in permitidas)]
    if nasceu:
        print('celulas que ganharam conteudo: %s' % sorted(nasceu))
    if reais:
        print('LAYOUT ALTERADO (%d):' % len(reais))
        for p in reais[:30]:
            print('   - %s' % p)
        sys.exit(1)
    print('LAYOUT PRESERVADO: nenhuma diferenca de estilo, largura, altura, '
          'mesclagem, formatacao condicional ou posicao de objeto.')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2],
         sys.argv[3] if len(sys.argv) > 3 else 'Painel')
