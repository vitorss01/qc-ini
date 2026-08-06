# -*- coding: utf-8 -*-
"""ajustes_abas.py - itens 1, 2 e 3 do pedido

1. Analitos  : restaura AbrirFormEspecificacoes (o botao existe, o ponto de
               entrada sumiu quando outra sessao reinstalou o mEspecificacoes)
2. DB_Result.: apaga K:AZ -- estrutura morta, conferido que nada a referencia
3. Painel    : filtros de Ano e Mes, sem tocar na largura da coluna A

Python, e nao PowerShell: os rotulos tem acento, e .ps1 lido como ANSI ja
corrompeu esta planilha uma vez.

Uso: python ajustes_abas.py <caminho.xlsm>
"""
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)
import win32com.client as w

SENHA = 'qcini2025'
MESES = ['(todos)', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
         'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

ENTRADA_ESPEC = [
    "",
    "' Ponto de entrada do botao da aba Analitos.",
    "'",
    "' Reinstalar o modulo por script apaga o que foi anexado ao fim dele. O",
    "' botao continua na planilha apontando para um nome que nao existe mais, e",
    "' o Excel responde \"nao e possivel executar a macro\" -- que parece codigo",
    "' orfao e nao e: e o unico caminho para cadastrar as metas de qualidade.",
    "Public Sub AbrirFormEspecificacoes()",
    "    frmEspecificacoes.Show",
    "End Sub",
]


def main(caminho):
    caminho = os.path.abspath(caminho)
    xl = w.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 1
    wb = xl.Workbooks.Open(caminho)
    salvou = False
    try:
        wb.EnableAutoRecover = False
    except Exception:
        pass
    if wb.ReadOnly:
        wb.Close(False)
        xl.Quit()
        raise SystemExit('Somente leitura')
    xl.Calculation = -4135

    try:
        estrutura = wb.ProtectStructure
        if estrutura:
            wb.Unprotect(SENHA)
        for s in wb.Worksheets:
            if s.ProtectContents:
                try:
                    s.Unprotect(SENHA)
                except Exception:
                    pass

        vbp = wb.VBProject
        nomes = [c.Name for c in vbp.VBComponents]

        # ---- ITEM 1 --------------------------------------------------------
        print('=== 1. botao Especificacoes ===')
        if 'frmEspecificacoes' not in nomes:
            print('   frmEspecificacoes AUSENTE -- nao ha o que restaurar')
        else:
            mod = None
            for c in vbp.VBComponents:
                if c.Name == 'mEspecificacoes':
                    mod = c
            if mod is None:
                raise SystemExit('mEspecificacoes ausente')
            cm = mod.CodeModule
            txt = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ''
            if 'AbrirFormEspecificacoes' in txt:
                print('   ponto de entrada ja existia')
            else:
                cm.InsertLines(cm.CountOfLines + 1, '\r\n'.join(ENTRADA_ESPEC))
                print('   AbrirFormEspecificacoes restaurado (%d linhas)'
                      % cm.CountOfLines)
            an = wb.Worksheets('Analitos')
            acao = [sh.OnAction for sh in an.Shapes if sh.Name == 'btnEspec']
            print('   botao btnEspec -> %s' % (acao[0] if acao else 'AUSENTE'))

        # ---- ITEM 2 --------------------------------------------------------
        print('=== 2. DB_Resultados K:AZ ===')
        db = wb.Worksheets('DB_Resultados')
        # Apagar CONTEUDO e formato, nao a coluna: Columns.Delete desloca BA:BD
        # (o bloco de formulas de lote, ate a linha 15.003) para a esquerda e
        # muda toda referencia que aponta para ele.
        db.Range(db.Columns(11), db.Columns(52)).Clear()
        print('   K:AZ limpas (conteudo e formato); BA:BD preservadas no lugar')
        print('   BA3 = %r' % db.Cells(3, 53).Value)

        # ---- ITEM 3 --------------------------------------------------------
        print('=== 3. Painel: filtros de Ano e Mes ===')
        pa = wb.Worksheets('Painel')
        # Destravar explicitamente: ProtectContents pode vir False e a aba ainda
        # recusar Locked/Validation. Tentar as duas formas e seguir mesmo assim
        # -- Locked so importa quando a aba esta protegida.
        for tentativa in (lambda: pa.Unprotect(SENHA), lambda: pa.Unprotect()):
            try:
                tentativa()
                break
            except Exception:
                pass
        larg_antes = pa.Columns(1).ColumnWidth
        anos = set()
        ultdb = db.Cells(db.Rows.Count, 1).End(-4162).Row
        if ultdb >= 4:
            for linha in db.Range(db.Cells(4, 2), db.Cells(ultdb, 2)).Value:
                if hasattr(linha[0], 'year'):
                    anos.add(linha[0].year)
        if not anos:
            import datetime
            anos = {datetime.date.today().year}
        anos = sorted(anos)

        # I3/I4: ao lado do De/Ate que ja existem, sem mexer em A
        pa.Cells(3, 8).Value = 'Ano'
        pa.Cells(4, 8).Value = 'Mês'
        for r in (3, 4):
            pa.Cells(r, 8).Font.Bold = True
            pa.Cells(r, 8).HorizontalAlignment = -4152
            cel = pa.Cells(r, 9)
            cel.Interior.Color = 15921906
            cel.Borders.LineStyle = 1
            cel.HorizontalAlignment = -4108
            try:
                cel.Locked = False
            except Exception:
                pass
        pa.Cells(3, 9).Validation.Delete()
        pa.Cells(3, 9).Validation.Add(3, 1, 1, ','.join(str(a) for a in anos))
        pa.Cells(3, 9).Validation.InCellDropdown = True
        pa.Cells(4, 9).Validation.Delete()
        pa.Cells(4, 9).Validation.Add(3, 1, 1, ','.join(MESES))
        pa.Cells(4, 9).Validation.InCellDropdown = True
        pa.Cells(3, 9).Value = anos[-1]
        pa.Cells(4, 9).Value = MESES[0]

        # NAO PODE HAVER DOIS Worksheet_Change NO MESMO MODULO.
        #
        # O Painel ja tem um, que chama AtualizarEixos quando B3/G3/G4/M/N
        # mudam. Acrescentar outro nao e "mais um gatilho": o VBA recusa o
        # modulo inteiro com "nome ambiguo", e a aba para de responder a
        # qualquer evento. A logica nova vive numa rotina propria e o handler
        # existente ganha UMA linha -- alteracao minima em codigo que ja
        # funciona.
        cod = [
            "",
            "' Ano e Mes sao ATALHO para preencher De/Ate, nao um filtro paralelo.",
            "'",
            "' O Calc filtra por filtroDe/filtroAte. Um segundo caminho de",
            "' filtragem seria uma segunda verdade sobre o mesmo periodo. Quem",
            "' quiser intervalo solto digita em De/Ate e poe o mes em (todos).",
            "'",
            "' Escreve com eventos DESLIGADOS e chama AtualizarEixos na mao: deixar",
            "' o evento reentrar para atualizar o eixo funciona por acidente e",
            "' quebra no dia em que alguem mudar a faixa vigiada.",
            "Public Sub AplicarFiltroAnoMes()",
            "    Dim a As Long, m As Long, s As String, ev As Boolean",
            "    On Error GoTo fim",
            "    If Not IsNumeric(Me.Range(\"I3\").Value) Then Exit Sub",
            "    ev = Application.EnableEvents",
            "    Application.EnableEvents = False",
            "    a = CLng(Me.Range(\"I3\").Value)",
            "    s = Trim$(CStr(Me.Range(\"I4\").Value))",
            "    m = MesDoRotulo(s)",
            "    If m = 0 Then",
            "        Me.Range(\"G3\").Value = DateSerial(a, 1, 1)",
            "        Me.Range(\"G4\").Value = DateSerial(a, 12, 31)",
            "    Else",
            "        Me.Range(\"G3\").Value = DateSerial(a, m, 1)",
            "        Me.Range(\"G4\").Value = DateSerial(a, m + 1, 0)",
            "    End If",
            "    AtualizarEixos",
            "fim:",
            "    Application.EnableEvents = ev",
            "End Sub",
            "",
            "Private Function MesDoRotulo(ByVal s As String) As Long",
            "    Dim v As Variant, i As Long",
            "    v = Array(\"JAN\", \"FEV\", \"MAR\", \"ABR\", \"MAI\", \"JUN\", _",
            "              \"JUL\", \"AGO\", \"SET\", \"OUT\", \"NOV\", \"DEZ\")",
            "    For i = 0 To 11",
            "        If UCase$(Left$(s, 3)) = v(i) Then MesDoRotulo = i + 1: Exit Function",
            "    Next i",
            "End Function",
        ]
        comp = None
        for c in vbp.VBComponents:
            try:
                if c.Type == 100 and c.Properties('Name').Value == 'Painel':
                    comp = c
            except Exception:
                pass
        if comp is None:
            raise SystemExit('modulo de documento do Painel nao encontrado')
        cm = comp.CodeModule
        txt = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ''

        if 'AplicarFiltroAnoMes' in txt:
            print('   rotina ja existia')
        else:
            cm.InsertLines(cm.CountOfLines + 1, '\r\n'.join(cod))
            print('   AplicarFiltroAnoMes + MesDoRotulo instalados')

        # liga no handler que JA existe, sem criar um segundo
        gancho = ('    If Not Intersect(Target, Me.Range("I3:I4")) '
                  'Is Nothing Then AplicarFiltroAnoMes')
        txt = cm.Lines(1, cm.CountOfLines)
        if 'Me.Range("I3:I4")' in txt:
            print('   gancho ja estava no Worksheet_Change')
        else:
            lin = cm.ProcBodyLine('Worksheet_Change', 0)
            cm.InsertLines(lin + 1, gancho)
            print('   gancho inserido no Worksheet_Change existente (L%d)' % (lin + 1))

        print('   anos no dropdown: %s' % ','.join(str(a) for a in anos))

        larg_depois = pa.Columns(1).ColumnWidth
        if abs(larg_antes - larg_depois) > 0.001:
            raise SystemExit('largura da coluna A mudou: %s -> %s'
                             % (larg_antes, larg_depois))
        print('   largura da coluna A intacta: %s' % larg_depois)

        xl.Calculation = -4105
        xl.CalculateFullRebuild()
        if estrutura and not wb.ProtectStructure:
            wb.Protect(SENHA, True, False)
        wb.Save()
        salvou = True
        print('SALVO: %s' % caminho)
    finally:
        try:
            xl.Calculation = -4105
        except Exception:
            pass
        try:
            wb.Close(salvou)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == '__main__':
    main(sys.argv[1])
