# -*- coding: utf-8 -*-
"""instalar_adr060.py -- ADR-060: lotes automaticos com aviso de validade, identificacao, travas da Configuracao.

Uso:  python instalar_adr060.py <Bioquimica|Hematologia> <arquivo.xlsm> [--registrar] [--sem-salvar]

Incremental sobre um arquivo ja com ADR-057/058/059. Faz:
  1. VBA: mLotes, mApp, mIntegracao, mSeguranca de src/<produto>/, mWestgardKnowledge (src/comum: a
     classificacao das regras casa com os codigos do motor) e o modulo da aba Configuracao
     (src/comum/Configuracao.cls); EXIGE compilacao;
  2. Configuracao, cadastro de lotes:
     - nome regValidadeCol = $D$26:$D$125; D com formato de data e VALIDACAO de data (2000..2099);
     - E = "Origem do cadastro" (automatico/manual), AD = sombra da validade (oculta; o "antes" da trilha),
       inicializada com a validade atual;
     - C21 (validade do lote em uso) olhando os 100 lotes (na Bioquimica olhava so 25);
     - textos de instrucao: lote entra sozinho, o usuario registra so a validade;
     - TRAVAS: toda a aba travada, editaveis so C5:C16 (identificacao/constantes), C20 (lote em uso) e
       D26:D125 (validade) -- como ja era a Hematologia; a Bioquimica estava inteira destravada;
  3. identificacao (Inicio C5:C9) passa a ser formula da Configuracao (uma fonte so); na Bioquimica o
     equipamento vira DIMENSION EXL-200 / SIEMENS e serie/controle (copiados da Hematologia) A INFORMAR;
  4. --registrar: roda mLotes.RegistrarLotesRecebidos (lotes ja recebidos entram no cadastro agora);
  5. devolve a protecao de cada aba como estava (instalar_seguranca_usuarios.devolver_protecao) e salva.
Confere tudo lendo de volta. Qualquer falha para ANTES de salvar.
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402
from instalar_modo_historico import normalizar  # noqa: E402
from instalar_seguranca_usuarios import foto_protecao, devolver_protecao  # noqa: E402

MODULOS = ['mLotes', 'mApp', 'mIntegracao', 'mSeguranca']
CFG = 'Configuração'
EDITAVEIS = ['C5:C16', 'C20', 'D26:D125']
IDENT_BIO = {'C9': 'DIMENSION EXL-200 / SIEMENS', 'C10': 'A INFORMAR', 'C11': 'A INFORMAR'}
# Inicio!C5:C9 <- Configuracao (Instituicao, Setor, Equipamento, Controle, Responsavel tecnico)
INICIO_DE = {'C5': 'C5', 'C6': 'C6', 'C7': 'C9', 'C8': 'C11', 'C9': 'C7'}
TXT_B24 = ('Os lotes recebidos do interfaceamento entram SOZINHOS no cadastro (coluna Origem). '
           'Registre aqui só a VALIDADE de cada lote (coluna D, dd/mm/aaaa) — o sistema avisa enquanto faltar. '
           'Lote que ainda não chegou: "+ Novo lote".')
TXT_B30 = ('1.  Lote novo de controle entra sozinho no primeiro resultado recebido; o sistema avisa e leva você '
           'até a VALIDADE.  Depois, cadastre a média e o DP em "Média e DP por lote".')


def log(*a):
    print(*a, flush=True)


def modulo_da_aba(wb, nome):
    for ws in wb.Worksheets:
        if ws.Name == nome:
            return ws.CodeName
    raise SystemExit(f'aba {nome} nao encontrada')


def main(produto, caminho, registrar=False, salvar=True):
    caminho = os.path.abspath(caminho)
    bio = produto.startswith('Bio')
    pasta = os.path.join(AQUI, 'src', 'bio' if bio else 'hema')
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:          # outro Excel segura o arquivo: o Save 'funcionaria' sem gravar nada
            raise SystemExit(f'arquivo aberto SOMENTE LEITURA (outra instancia o segura): {caminho}')
        estrutura = bool(wb.ProtectStructure)
        foto = foto_protecao(wb)
        ii.desproteger_tudo(wb)
        ex.xl.Calculation = -4105

        log('1. VBA (todo o codigo atual das fontes: codigo_atual.py)')
        import codigo_atual
        codigo_atual.instalar(ex, wb, produto, log)

        log('2. Configuração: cadastro de lotes')
        ws = wb.Worksheets(CFG)
        ii.nome(wb, 'regValidadeCol', f"='{CFG}'!$D$26:$D$125")
        ws.Range('C21').Formula = '=IFERROR(INDEX($D$26:$D$125,MATCH($C$20,$C$26:$C$125,0)),"")'
        d = ws.Range('D26:D125')
        d.NumberFormat = 'dd/mm/yyyy'
        d.Validation.Delete()
        # datas como numero de serie: independente do idioma da instalacao
        d.Validation.Add(4, 1, 1, '36526', '73050')        # xlValidateDate, xlValidAlertStop, xlBetween
        d.Validation.IgnoreBlank = True
        d.Validation.InputTitle = 'Validade do lote'
        d.Validation.InputMessage = 'Data de validade do controle (dd/mm/aaaa).'
        d.Validation.ErrorTitle = 'Validade inválida'
        d.Validation.ErrorMessage = 'Digite a validade como data (dd/mm/aaaa), entre 2000 e 2099.'
        ws.Range('C25').Value = 'Lote'
        ws.Range('E25').Value = 'Origem do cadastro'
        ws.Range('E25').Font.Bold = ws.Range('D25').Font.Bold
        ws.Range('E25').Interior.Color = ws.Range('D25').Interior.Color
        ws.Range('E25').Font.Color = ws.Range('D25').Font.Color
        if ws.Columns(5).ColumnWidth < 60:
            ws.Columns(5).ColumnWidth = 62
        ws.Range('AD25').Value = 'Validade registrada (sombra ADR-060)'
        sombra = ws.Range('AD26:AD125')
        sombra.Value = d.Value                                  # o "antes" da proxima edicao
        sombra.NumberFormat = 'dd/mm/yyyy'
        ws.Columns(30).Hidden = True
        ws.Range('B24').Value = TXT_B24
        f2 = ws.Range('F2')
        if str(f2.Value or '').startswith('loteCarregado'):
            f2.Value = 'loteParam→'
        # travas: aba inteira travada, editaveis so as entradas do usuario
        ws.Cells.Locked = True
        for a in EDITAVEIS:
            ws.Range(a).Locked = False

        log('3. identificação')
        cfg_vals = {}
        if bio:
            for cel, v in IDENT_BIO.items():
                antes = ws.Range(cel).Value
                ws.Range(cel).Value = v
                cfg_vals[cel] = (antes, v)
        ini = wb.Worksheets('Início')
        divergentes = []
        for c_ini, c_cfg in INICIO_DE.items():
            atual = str(ini.Range(c_ini).Value or '').strip()
            fonte = str(ws.Range(c_cfg).Value or '').strip()
            if not bio and atual != fonte:
                divergentes.append((c_ini, atual, c_cfg, fonte))
                continue
            ini.Range(c_ini).Formula = f"='{CFG}'!${c_cfg[0]}${c_cfg[1:]}"
            ini.Range(c_ini).Locked = True
        for c in ini.Range('B30').Cells:
            if 'Chegou lote novo' in str(c.Value or ''):
                c.Value = TXT_B30
        if divergentes:
            log(f'   ATENCAO: Inicio diverge da Configuracao (mantido como estava): {divergentes}')

        ex.xl.Calculate()
        # conferencia lida de volta
        falhas = []
        if wb.Names('regValidadeCol').RefersToRange.Address != '$D$26:$D$125':
            falhas.append('regValidadeCol')
        if 'D$26:$D$125' not in ws.Range('C21').Formula:
            falhas.append('C21')
        if ws.Range('D26').Validation.Type != 4:
            falhas.append('validacao D26')
        if not ws.Columns(30).Hidden:
            falhas.append('sombra visivel')
        bloq = [a for a in EDITAVEIS if ws.Range(a).Locked is not False]
        if bloq or ws.Range('B26').Locked is not True or ws.Range('C26').Locked is not True:
            falhas.append(f'travas {bloq}')
        if bio and str(ini.Range('C7').Text) != IDENT_BIO['C9']:
            falhas.append('Inicio!C7 ' + str(ini.Range('C7').Text))
        if falhas:
            raise SystemExit(f'conferencia falhou: {falhas}')
        log('   regValidadeCol, C21, validacao de data, origem, sombra oculta, travas e identificacao conferidos')

        # trilha da identificacao (a propria rotina de auditoria do arquivo)
        for cel, (antes, depois) in cfg_vals.items():
            ex.run("'" + wb.Name + "'!mAuditoria.RegistrarLog", 'IDENTIFICACAO_ALTERADA',
                   f'Configuração!{cel}: {antes} -> {depois} (instalador ADR-060)')

        if registrar:
            log('4. registrar lotes ja recebidos')
            r = str(ex.run("'" + wb.Name + "'!mLotes.RegistrarLotesRecebidos", teto=300))
            log('   ' + r)
            if not r.startswith('OK|'):
                raise SystemExit('RegistrarLotesRecebidos falhou: ' + r)

        devolver_protecao(wb, foto)
        if estrutura:
            wb.Protect(ii.SENHA, True, False)
        if salvar:
            wb.Save()
            log('salvo')
    finally:
        ex.fechar()


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], registrar='--registrar' in a, salvar='--sem-salvar' not in a)
