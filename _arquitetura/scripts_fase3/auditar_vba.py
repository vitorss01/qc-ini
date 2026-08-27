# -*- coding: utf-8 -*-
"""auditar_vba.py - auditoria estatica do projeto VBA inteiro

POR QUE ESTATICA, SE EXISTE O Debug > Compile

Porque o compilador para no PRIMEIRO erro. Numa cadeia com splice, isso vira um
ciclo de "corrige um, recompila, descobre o proximo" que custa uma abertura do
Excel por defeito e ainda mascara os problemas que NAO impedem a compilacao --
duplicata de nome que o VBA resolve escolhendo um dos dois, saida antecipada
que deixa aba destrancada, tratamento de erro que engole excecao. A analise
estatica ve todos de uma vez, e o compilador entra depois como confirmacao.

O QUE PROCURA

  A. procedure chamada e nao definida em lugar nenhum do projeto
  B. nome PUBLICO duplicado entre modulos (o VBA nao compila, e a mensagem
     costuma apontar o modulo errado)
  C. procedure duplicada DENTRO do mesmo modulo
  D. declaracao de modulo (Dim/Const/Type/Declare) depois da primeira
     procedure -- erro que ja apareceu tres vezes neste projeto
  E. chamada ambigua: nome definido em dois modulos, um Public e um Private
  F. On Error GoTo <label> sem o label no mesmo procedimento
  G. procedure que desprotege aba e tem saida sem passar pelo restauro
  H. Public sem uso em nenhum lugar do projeto (candidato a virar Private)
  I. Option Explicit ausente

O QUE **NAO** PROVA

Que compila. Isto e um pente fino, nao um compilador -- VBA tem resolucao de
nome por escopo que este parser nao reproduz inteira. O veredito final continua
sendo o Debug > Compile do proprio Excel.

Uso: python auditar_vba.py <pasta_com_modulos_exportados> [--verboso]
"""
import io
import os
import re
import sys
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                              write_through=True)

RE_PROC = re.compile(
    r'^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?'
    r'(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)',
    re.IGNORECASE)
RE_FIM = re.compile(r'^\s*End\s+(Sub|Function|Property)\b', re.IGNORECASE)
# Declaracao de modulo tem DUAS formas, e a segunda passou batido na primeira
# versao deste analisador:
#
#   Private Const X As String = "a"      <- tem palavra-chave
#   Private gTrace As Object             <- NAO tem: e "escopo nome As tipo"
#
# A segunda e exatamente a forma de gTrace e gNaoAval, que ja quebraram a
# compilacao deste projeto por nascerem no meio do modulo. Exigir Dim/Const
# deixava o analisador cego justamente para o defeito historico -- e um
# analisador cego para o bug que ja aconteceu nao serve para nada.
RE_DECL = re.compile(
    r'^\s*(?:(Public|Private|Global)\s+)?'
    r'(Const|Dim|Type|Enum|Declare|WithEvents)\b', re.IGNORECASE)
RE_DECL_NUA = re.compile(
    r'^\s*(Public|Private|Global)\s+[A-Za-z_]\w*(\s*\([^)]*\))?\s+As\s+',
    re.IGNORECASE)
RE_DECL_PTR = re.compile(
    r'^\s*(?:Public|Private)?\s*Declare\s+PtrSafe\b', re.IGNORECASE)
RE_LABEL = re.compile(r'^([A-Za-z_]\w*):\s*(?:\'.*)?$')
RE_ONERR = re.compile(r'On\s+Error\s+GoTo\s+([A-Za-z_]\w*)', re.IGNORECASE)
RE_GOTO = re.compile(r'(?<!On Error )\bGoTo\s+([A-Za-z_]\w*)', re.IGNORECASE)

# Nomes que o VBA/Excel fornecem: chamar nao significa ter de definir.
INTRINSECOS = set('''
abs array asc atn cbool cbyte ccur cdate cdbl cdec chr chr$ cint clng
cos csng cstr cvar cverr date dateadd datediff datepart dateserial datevalue
day ddb dir doevents environ eof error error$ exp fileattr filedatetime
filelen fix format format$ freefile fv getattr hex hour iif imestatus input
inputb inputbox instr instrrev int ipmt irr isarray isdate isempty iserror
ismissing isnull isnumeric isobject join lbound lcase lcase$ left left$ len
loc lof log ltrim ltrim$ mid mid$ minute mirr month monthname msgbox now nper
npv oct partition pmt ppmt pv qbcolor rate replace rgb right right$ rnd round
rtrim rtrim$ second seek sgn shell sin sln space space$ spc split sqr str
str$ strcomp strconv string string$ strreverse switch syd tab tan time timer
timeserial timevalue trim trim$ typename ubound ucase ucase$ val vartype
weekday weekdayname year cdbl clng
set let get on error resume next exit end if then else elseif select case
for each next do loop while wend with call redim erase stop debug print
'''.split())


FORMULARIOS = set()          # modulos .frm: controles vivem no designer


def carregar(pasta):
    mods = {}
    for nome in sorted(os.listdir(pasta)):
        if not nome.lower().endswith(('.bas', '.cls', '.frm', '.doc', '.txt')):
            continue
        caminho = os.path.join(pasta, nome)
        # newline='' na LEITURA tambem: sem isso o Python aplica universal
        # newlines, um \r solto vira quebra, e o modulo aparece com o dobro de
        # linhas -- com as fronteiras de procedure deslocadas. A auditoria roda
        # inteira sobre lixo e devolve "nada encontrado" com ar de aprovacao.
        texto = io.open(caminho, encoding='cp1252', errors='replace',
                        newline='').read()
        texto = texto.replace('\r\n', '\n').replace('\r', '\n')
        base, ext = os.path.splitext(nome)
        if ext.lower() == '.frm':
            FORMULARIOS.add(base)
        mods[base] = texto.split('\n')
    return mods


def procedures(linhas):
    """Devolve [(nome, escopo, tipo, linha_ini, linha_fim)]."""
    fora = []
    atual = None
    for i, ln in enumerate(linhas):
        if ln.lstrip().startswith("'"):
            continue
        m = RE_PROC.match(ln)
        if m and atual is None:
            escopo = (m.group(1) or 'Public').capitalize()
            atual = [m.group(3), escopo, m.group(2).split()[0].capitalize(), i, i]
            continue
        if atual is not None and RE_FIM.match(ln):
            atual[4] = i
            fora.append(tuple(atual))
            atual = None
    if atual is not None:
        atual[4] = len(linhas) - 1
        fora.append(tuple(atual))
    return fora


def main():
    pasta = os.path.abspath(sys.argv[1])
    verboso = '--verboso' in sys.argv
    mods = carregar(pasta)
    print('=' * 78)
    print('AUDITORIA ESTATICA -- %s' % os.path.basename(pasta))
    print('=' * 78)
    print('%d modulos, %d linhas' % (len(mods), sum(len(v) for v in mods.values())))
    print()

    procs = {}                      # modulo -> lista
    definidos = defaultdict(list)   # nome_lower -> [(modulo, escopo)]
    for mod, linhas in mods.items():
        procs[mod] = procedures(linhas)
        for nome, escopo, tipo, a, b in procs[mod]:
            definidos[nome.lower()].append((mod, escopo, tipo))

    achados = defaultdict(list)

    # ---- J. identificador que colide com nome nativo do VBA ----------------
    #
    # VBA nao distingue caixa, entao um identificador so precisa COINCIDIR com
    # um nome da linguagem para causar estrago. Duas formas, ambas ja vistas
    # neste projeto:
    #
    #   PALAVRA RESERVADA  Dim aM As Double, aS As Double
    #                      'aS' E 'As'. O modulo inteiro para com "Erro de
    #                      sintaxe" -- a mEstatistica da Hematologia nunca
    #                      compilou por causa disso.
    #
    #   FUNCAO NATIVA      Dim vAl As Variant   ...   x = Val(CStr(y))
    #                      'vAl' E 'Val'. A chamada a funcao vira indexacao da
    #                      variavel e devolve "Erro 13: tipos incompativeis",
    #                      longe da declaracao que causou o problema.
    #
    # A auditoria nao analisa gramatica, entao nao encontraria nenhum dos dois
    # pelos outros criterios. Este e barato e pega os dois.
    RESERVADAS = set("""as if then else elseif end sub function property let get set
        dim redim static const public private friend option explicit for next each in
        do loop while wend with select case to step exit goto gosub on error resume
        byval byref optional paramarray new nothing true false empty null me stop
        call type enum declare lib alias implements withevents raiseevent event
        and or not xor eqv imp mod is like typeof preserve""".split())
    NATIVAS = set("""abs asc atn cbool cbyte ccur cdate cdbl cdec chr cint clng
        cos csng cstr cvar cvdate cverr date dateadd datediff datepart dateserial
        datevalue day ddb dir doevents environ eof erl err error exp filelen fix
        format fv hex hour iif imestatus input instr instrrev int ipmt irr isarray
        isdate isempty iserror ismissing isnull isnumeric isobject join lbound lcase
        left len loc lof log ltrim mid minute mirr month monthname msgbox now nper
        npv oct partition pmt ppmt pv qbcolor rate rgb right rnd round rtrim second
        seek sgn sin sln space spc split sqr str strcomp strconv string strreverse
        switch syd tab tan time timer timeserial timevalue trim typename ubound
        ucase val vartype weekday weekdayname year""".split())
    RXDECL = re.compile(
        r'^\s*(?:Dim|ReDim|Static|Private|Public|Global|Const)\s+(.+)$', re.I)
    for mod, linhas in mods.items():
        for i, ln in enumerate(linhas):
            corpo = ln.split("'")[0]
            m = RXDECL.match(corpo)
            if not m:
                continue
            for nome in re.findall(r'(?<![.\w])([A-Za-z_]\w*)\s+As\s',
                                   m.group(1)):
                b = nome.lower()
                if b in RESERVADAS:
                    achados['J. identificador colide com PALAVRA RESERVADA'
                            ].append('%s l.%d: %s' % (mod, i + 1, nome))
                elif b in NATIVAS:
                    # so acusa se a funcao nativa e mesmo usada no modulo --
                    # senao nao ha colisao de fato, so nome infeliz
                    if re.search(r'(?<![.\w])%s\s*\(' % re.escape(b),
                                 '\n'.join(linhas), re.I):
                        achados['J. identificador colide com FUNCAO NATIVA usada'
                                ].append('%s l.%d: %s (sombreia %s)'
                                         % (mod, i + 1, nome, b.capitalize()))

    # ---- I. Option Explicit -------------------------------------------------
    for mod, linhas in mods.items():
        if not any(l.strip().lower().startswith('option explicit')
                   for l in linhas[:15]):
            if any(l.strip() for l in linhas):
                achados['I. Option Explicit ausente'].append(mod)

    # ---- B/C/E. nomes duplicados -------------------------------------------
    for nome, ocor in sorted(definidos.items()):
        if len(ocor) < 2:
            continue
        pormod = defaultdict(list)
        for mod, escopo, tipo in ocor:
            pormod[mod].append(escopo)
        # C: duplicado no MESMO modulo
        for mod, escopos in pormod.items():
            if len(escopos) > 1:
                achados['C. procedure duplicada no mesmo modulo'].append(
                    '%s.%s (%dx)' % (mod, nome, len(escopos)))
        publicos = [m for m, e, t in ocor if e in ('Public', 'Friend')]
        privados = [m for m, e, t in ocor if e == 'Private']
        if len(set(publicos)) > 1:
            achados['B. nome PUBLICO duplicado entre modulos'].append(
                '%s: %s' % (nome, ', '.join(sorted(set(publicos)))))
        elif publicos and privados:
            achados['E. mesmo nome Public num modulo e Private noutro'].append(
                '%s: Public em %s / Private em %s'
                % (nome, ', '.join(sorted(set(publicos))),
                   ', '.join(sorted(set(privados)))))

    # ---- D. declaracao de modulo depois da primeira procedure --------------
    for mod, linhas in mods.items():
        if not procs[mod]:
            continue
        primeira = min(p[3] for p in procs[mod])
        dentro = {(a, b) for _, _, _, a, b in procs[mod]}
        for i, ln in enumerate(linhas):
            if i <= primeira:
                continue
            if any(a <= i <= b for a, b in dentro):
                continue
            if ln.lstrip().startswith("'") or not ln.strip():
                continue
            if (RE_DECL.match(ln) or RE_DECL_PTR.match(ln)
                    or RE_DECL_NUA.match(ln)):
                achados['D. declaracao de modulo apos a 1a procedure'].append(
                    '%s:%d  %s' % (mod, i + 1, ln.strip()[:60]))

    # ---- A. chamada sem definicao ------------------------------------------
    # Conservador de proposito: so conta o que tem cara inequivoca de chamada
    # de procedure do projeto -- "Nome arg" no inicio da linha, "Call Nome",
    # "Modulo.Nome" e "= Nome(".
    RE_CALL = re.compile(r'\bCall\s+([A-Za-z_]\w*)', re.IGNORECASE)
    RE_QUAL = re.compile(r'\b(m[A-Z]\w*)\.([A-Za-z_]\w*)')
    RE_APPRUN = re.compile(r'Application\.Run\s+"([^"]+)"', re.IGNORECASE)
    conhecidos = set(definidos)
    usados = defaultdict(set)
    faltando = defaultdict(set)
    for mod, linhas in mods.items():
        for i, ln in enumerate(linhas):
            s = ln.strip()
            if s.startswith("'") or not s:
                continue
            s = re.sub(r'"[^"]*"', '""', s) if 'Application.Run' not in s else s
            for m in RE_CALL.finditer(s):
                alvo = m.group(1).lower()
                usados[alvo].add(mod)
                if alvo not in conhecidos and alvo not in INTRINSECOS:
                    faltando[m.group(1)].add('%s:%d' % (mod, i + 1))
            for m in RE_QUAL.finditer(s):
                dono, alvo = m.group(1), m.group(2)
                usados[alvo.lower()].add(mod)
                if dono in mods:
                    nomes = {p[0].lower() for p in procs[dono]}
                    consts = {c.lower() for c in re.findall(
                        r'^\s*Public\s+Const\s+(\w+)', '\n'.join(mods[dono]),
                        re.IGNORECASE | re.MULTILINE)}
                    if alvo.lower() not in nomes | consts:
                        faltando['%s.%s' % (dono, alvo)].add(
                            '%s:%d' % (mod, i + 1))
            for m in RE_APPRUN.finditer(ln):
                alvo = m.group(1)
                usados[alvo.lower()].add(mod)
                if alvo.lower() not in conhecidos:
                    faltando['Application.Run("%s")' % alvo].add(
                        '%s:%d' % (mod, i + 1))
    for nome, onde in sorted(faltando.items()):
        achados['A. chamada sem definicao no projeto'].append(
            '%-42s <- %s' % (nome, ', '.join(sorted(onde)[:3])))

    # ---- F. On Error GoTo sem label ----------------------------------------
    for mod, linhas in mods.items():
        for nome, escopo, tipo, a, b in procs[mod]:
            corpo = linhas[a:b + 1]
            labels = {RE_LABEL.match(l).group(1).lower()
                      for l in corpo if RE_LABEL.match(l)}
            for l in corpo:
                if l.lstrip().startswith("'"):
                    continue
                for m in list(RE_ONERR.finditer(l)) + list(RE_GOTO.finditer(l)):
                    alvo = m.group(1)
                    if alvo == '0':
                        continue
                    if alvo.lower() not in labels:
                        achados['F. GoTo/On Error para label inexistente'].append(
                            '%s.%s -> %s' % (mod, nome, alvo))

    # ---- G. desprotege e sai sem restaurar ---------------------------------
    #
    # Reproteger nem sempre e um ".Protect" literal: o projeto usa envolucros
    # (ProtegerAudit, RestaurarProtecao). Procurar so pelo literal acusava
    # mAuditoria.Auditar, que reprotege corretamente -- e um analisador que
    # grita em codigo certo treina quem le a ignora-lo.
    envolucros = set()
    for mod, linhas in mods.items():
        for nome, escopo, tipo, a, b in procs[mod]:
            if re.search(r'\.Protect\b', '\n'.join(linhas[a:b + 1])):
                envolucros.add(nome.lower())
    RE_REPROT = re.compile(
        r'\.Protect\b|\b(%s)\b' % '|'.join(sorted(envolucros) or ['zzz']),
        re.IGNORECASE)

    # Rotinas cuja RAZAO DE EXISTIR e deixar a pasta destrancada: o Modo
    # Desenvolvedor e o UnprotectAll do login. Nao sao defeito, e marca-las
    # como tal esconderia as que sao.
    POR_DESIGN = {'unprotectall', 'desprotegerestrutura', 'mododesenvolvedor',
                  'liberarescrita'}

    for mod, linhas in mods.items():
        for nome, escopo, tipo, a, b in procs[mod]:
            if nome.lower() in POR_DESIGN:
                continue
            corpo = linhas[a:b + 1]
            txt = '\n'.join(corpo)
            libera = re.search(r'\.Unprotect\b|LiberarEscrita\s*\(', txt)
            if not libera:
                continue
            restaura = RE_REPROT.search(txt)
            if not restaura:
                achados['G. desprotege e NAO reprotege'].append(
                    '%s.%s' % (mod, nome))
                continue
            # Reprotege, mas o caminho de ERRO passa por la? Sem On Error, uma
            # excecao no meio deixa a aba destrancada -- e no Audit_Log, que
            # guarda a cadeia de hash, isso e furo de integridade, nao susto.
            if not re.search(r'On\s+Error\s+(GoTo|Resume)', txt, re.IGNORECASE):
                achados['G. janela destrancada sem On Error (erro deixa aberta)'
                        ].append('%s.%s' % (mod, nome))
            i_lib = corpo.index(next(l for l in corpo if re.search(
                r'\.Unprotect\b|LiberarEscrita\s*\(', l)))
            for j in range(i_lib + 1, len(corpo)):
                l = corpo[j]
                if l.lstrip().startswith("'"):
                    continue
                if re.search(r'^\s*(If .*Then\s+)?Exit (Sub|Function)\b', l):
                    # RE_REPROT, e nao o literal ".Protect": a reprotecao pode
                    # ser um envolucro (ProtegerAudit). Usar padroes diferentes
                    # nas duas verificacoes fazia esta acusar uma saida que ja
                    # estava coberta -- inconsistencia do proprio analisador.
                    if not RE_REPROT.search('\n'.join(corpo[i_lib:j + 1])):
                        achados['G. saida antecipada dentro da janela destrancada'
                                ].append('%s.%s linha ~%d: %s'
                                         % (mod, nome, a + j + 1, l.strip()[:44]))

    # ---- J. identificador nao declarado (Option Explicit) -------------------
    #
    # Esta e a classe "Variavel nao definida", que o compilador acusa apontando
    # a LINHA onde parou -- nao a declaracao que falta. Com Option Explicit em
    # todo modulo, um nome que nao esta em lugar nenhum e erro de compilacao.
    #
    # Conservador de proposito: so acusa identificador que nao aparece em NENHUM
    # lugar declaravel do projeto. Nome de tipo (apos "As"), acesso a membro
    # (apos ponto), rotulo e intrinseco do VBA/Excel ficam de fora. Preferimos
    # deixar passar um caso duvidoso a inundar a saida -- lista com ruido nao e
    # lida.
    TIPOS_E_PALAVRAS = set('''
    long string double boolean date variant object integer byte single currency
    worksheet workbook range collection name nothing true false empty null
    application thisworkbook activesheet activeworkbook selection err debug
    intersect union unload load createobject getobject sheets worksheets names
    cells rows columns font value formula address count offset resize end
    entirerow entirecolumn usedrange listobjects shapes chartobjects target
    cancel saveasui cachedvalue
    excel vba me new as byval byref optional paramarray to step in each then
    else elseif end exit resume goto gosub select case is like mod and or not
    xor eqv imp addressof typeof preserve redim erase let set const dim static
    public private friend global sub function property get put close open print
    write input line lock unlock seek width reset kill mkdir rmdir chdir chdrive
    filecopy name setattr randomize option explicit base compare module declare
    lib alias ptrsafe longptr longlong type enum event withevents implements
    on error next loop do while until wend for if
    '''.split())
    # Tem de comecar por LETRA: "_" sozinho e continuacao de linha, e casava
    # com [A-Za-z_]\w* gerando centenas de falsos positivos.
    # "&" no lookbehind: &H428A2F98 e literal hexadecimal, nao identificador.
    # Sem isso a tabela de constantes do SHA-256 em mSeguranca sozinha gerava
    # dezenas de falsos positivos.
    RE_ID = re.compile(r'(?<![.\w!$&])([A-Za-z]\w*)')
    RE_STR = re.compile(r'"[^"]*"')

    declaraveis = set(definidos)
    for mod, linhas in mods.items():
        texto = '\n'.join(linhas)
        for m in re.finditer(
                r'^\s*(?:Public|Private|Global|Dim|Static)\s+'
                r'(?:Const\s+|WithEvents\s+)?([A-Za-z_]\w*)',
                texto, re.IGNORECASE | re.MULTILINE):
            declaraveis.add(m.group(1).lower())
        for m in re.finditer(r'^\s*(?:Public|Private)?\s*(?:Type|Enum)\s+([A-Za-z_]\w*)',
                             texto, re.IGNORECASE | re.MULTILINE):
            declaraveis.add(m.group(1).lower())
        # API do Windows: "Private Declare PtrSafe Sub GetSystemTime Lib ..."
        # O nome vem depois de Sub/Function, nao logo apos o escopo.
        for m in re.finditer(
                r'\bDeclare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+([A-Za-z_]\w*)',
                texto, re.IGNORECASE):
            declaraveis.add(m.group(1).lower())
        # membros de Enum e campos de Type
        for bloco_m in re.finditer(
                r'^\s*(?:Public|Private)?\s*(Type|Enum)\s+\w+(.*?)^\s*End\s+\1',
                texto, re.IGNORECASE | re.MULTILINE | re.DOTALL):
            for l in bloco_m.group(2).split('\n'):
                mm = re.match(r'\s*([A-Za-z_]\w*)', l)
                if mm:
                    declaraveis.add(mm.group(1).lower())
        declaraveis.add(mod.lower())

    for mod, linhas in mods.items():
        # UserForm fica de fora: os controles (txtEx1Ini, cboMes, optAnual) sao
        # declarados na secao de DESIGNER do .frm, que o CodeModule nao contem.
        # Verificar aqui acusaria cada controle como nao declarado -- quase cem
        # falsos positivos que enterrariam qualquer achado real.
        if mod in FORMULARIOS:
            continue
        for nome, escopo, tipo, a, b in procs[mod]:
            corpo = linhas[a:b + 1]
            locais = set()
            # parametros da assinatura (pode ter continuacao de linha)
            assinatura = ''
            for l in corpo:
                assinatura += l.rstrip()
                if not assinatura.endswith('_'):
                    break
                assinatura = assinatura[:-1]
            for m in re.finditer(r'(?:ByVal|ByRef|Optional|ParamArray)?\s*'
                                 r'([A-Za-z_]\w*)\s*(?:\(\s*\))?\s*As\s',
                                 assinatura, re.IGNORECASE):
                locais.add(m.group(1).lower())
            rotulos = {RE_LABEL.match(l).group(1).lower()
                       for l in corpo if RE_LABEL.match(l)}
            for l in corpo:
                for m in re.finditer(
                        r'\b(?:Dim|Static|Const|ReDim)\s+(?:Preserve\s+)?'
                        r'([A-Za-z_]\w*)', l, re.IGNORECASE):
                    locais.add(m.group(1).lower())
                # Dim a As X, b As Y  -> pega os demais
                mm = re.match(r'\s*(?:Dim|Static|Const)\s+(.*)', l, re.IGNORECASE)
                if mm:
                    for parte in mm.group(1).split(','):
                        p = re.match(r'\s*([A-Za-z_]\w*)', parte)
                        if p:
                            locais.add(p.group(1).lower())
                for m in re.finditer(r'\bFor\s+Each\s+([A-Za-z_]\w*)|'
                                     r'\bFor\s+([A-Za-z_]\w*)\s*=', l, re.IGNORECASE):
                    locais.add((m.group(1) or m.group(2)).lower())

            conhecido = declaraveis | locais | rotulos | TIPOS_E_PALAVRAS \
                | INTRINSECOS | {nome.lower()}
            for idx, l in enumerate(corpo):
                s = l.strip()
                if s.startswith("'") or not s:
                    continue
                s = RE_STR.sub('""', s)
                if "'" in s:
                    s = s.split("'")[0]
                # nomes de tipo apos "As" nao sao variaveis
                s = re.sub(r'\bAs\s+[A-Za-z_]\w*', ' As ', s, flags=re.IGNORECASE)
                for m in RE_ID.finditer(s):
                    tok = m.group(1)
                    low = tok.lower()
                    if low in conhecido:
                        continue
                    if low.startswith('xl') or low.startswith('vb') \
                            or low.startswith('mso') or low.startswith('fm'):
                        continue
                    resto = s[m.end():].lstrip()
                    # Seguido de "(" e CHAMADA, nao variavel. Se for procedure
                    # do projeto que nao existe, a categoria A ja acusa com
                    # diagnostico melhor; se for funcao de biblioteca, e ruido.
                    if resto.startswith('('):
                        continue
                    # Seguido de ":=" e ARGUMENTO NOMEADO (Password:=, After:=).
                    # Sem esta regra, toda chamada a .Protect gerava seis
                    # "identificadores nao declarados" -- era a maior parte do
                    # ruido, e ruido treina quem le a ignorar a lista inteira.
                    if resto.startswith(':='):
                        continue
                    achados['J. identificador nao declarado (Option Explicit)'
                            ].append('%s.%s:%d  %s'
                                     % (mod, nome, a + idx + 1, tok))

    # ---- H. Public sem uso --------------------------------------------------
    if verboso:
        for nome, ocor in sorted(definidos.items()):
            if any(e in ('Public', 'Friend') for _, e, _ in ocor):
                if nome not in usados:
                    achados['H. Public sem chamada localizada (informativo)'
                            ].append('%s (%s)' % (nome, ocor[0][0]))

    # ---- saida --------------------------------------------------------------
    graves = ['A. chamada sem definicao no projeto',
              'B. nome PUBLICO duplicado entre modulos',
              'C. procedure duplicada no mesmo modulo',
              'D. declaracao de modulo apos a 1a procedure',
              'F. GoTo/On Error para label inexistente',
              'G. desprotege e NAO reprotege',
              'G. janela destrancada sem On Error (erro deixa aberta)',
              'G. saida antecipada dentro da janela destrancada',
              # Os dois sao GRAVES porque impedem o codigo de rodar: a palavra
              # reservada derruba a COMPILACAO do modulo inteiro, e a funcao
              # sombreada troca uma chamada por indexacao de variavel, sempre
              # com erro longe da declaracao que o causou.
              'J. identificador colide com PALAVRA RESERVADA',
              'J. identificador colide com FUNCAO NATIVA usada']
    nGraves = 0
    for chave in sorted(achados):
        itens = achados[chave]
        marca = 'GRAVE' if chave in graves else 'aviso'
        if chave in graves:
            nGraves += len(itens)
        print('[%s] %s -- %d' % (marca, chave, len(itens)))
        for x in itens[:20]:
            print('        %s' % x)
        if len(itens) > 20:
            print('        ... e mais %d' % (len(itens) - 20))
        print()

    if not achados:
        print('nada encontrado')
    print('-' * 78)
    print('GRAVES: %d' % nGraves)
    return 1 if nGraves else 0


if __name__ == '__main__':
    sys.exit(main())
