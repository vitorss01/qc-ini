Private Sub Workbook_Open()
    On Error Resume Next
    mApp.Reset                      ' calculo/eventos/tela num estado conhecido
    mDados.AtualizarListasAno
    mLotes.SincronizarLotesAoAbrir
    ' AutoSave (arquivo no OneDrive/SharePoint) gravaria a cada poucos segundos -- e cada gravacao
    ' passa pelo estado trancado. Gravar e decisao do usuario (Ctrl+S). Excel sem a propriedade: ignora.
    CallByName ThisWorkbook, "AutoSaveOn", VbLet, False
    On Error GoTo 0
    LockApp
End Sub
' ADR-061: voltar de outra pasta NAO dispara Worksheet_Activate.
Private Sub Workbook_Activate()
    On Error Resume Next
    mUI.AjustarGraficos             ' sai calado se a aba ativa nao for o Painel
    mUI.GrafVigiaLigar
End Sub
' Troca de pasta E fechamento (dispara depois do BeforeClose).
Private Sub Workbook_Deactivate()
    On Error Resume Next
    mUI.GrafVigiaDesligar
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim nome As Variant, erro As String
    On Error Resume Next
    mUI.GrafVigiaDesligar           ' ADR-061: "Salvar como" muda o nome que o tique cita
    Application.EnableEvents = False
    FlushLoteAtual
    ' ADR-065: no disco, sempre TRANCADO. A gravacao e feita aqui, com os eventos desligados: trancar
    ' a vista dentro do BeforeSave e deixar o Excel gravar NAO chegava ao arquivo (medido: o XML saia
    ' com as abas da sessao visiveis e abertas). Cancela a gravacao do Excel e grava a propria.
    Cancel = True
    If SaveAsUI Then
        nome = Application.GetSaveAsFilename(Me.FullName, _
               "Pasta de Trabalho Habilitada para Macro do Excel (*.xlsm),*.xlsm")
        If VarType(nome) = vbBoolean Then GoTo fim          ' "Salvar como" cancelado: nada muda
    End If
    mSeguranca.TrancarParaGravar
    Err.Clear
    If SaveAsUI Then
        Me.SaveAs Filename:=nome, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Else
        Me.Save
    End If
    If Err.Number <> 0 Then erro = Err.Description
    Err.Clear
    If Not mUI.FechandoPasta() Then mSeguranca.RestaurarAposGravar   ' devolve a vista (fechando: nao)
    If Len(erro) > 0 Then
        Me.Saved = False
        ' so com gente no Excel: uma caixa numa instancia de automacao trava quem chamou
        If Application.UserControl And Application.DisplayAlerts Then MsgBox "Nao foi possivel gravar: " & erro, vbExclamation, "QC_INI"
    End If
fim:
    Application.EnableEvents = True
    If Not mUI.FechandoPasta() Then mUI.GrafVigiaLigar
End Sub
' ADR-054/063: janela de outro tamanho = outra largura E outra altura visivel; o Painel
' se encaixa de novo (todos os graficos inteiros na tela). Fora do Painel sai calado.
Private Sub Workbook_WindowResize(ByVal Wn As Window)
    On Error Resume Next
    mUI.PainelEncaixar
End Sub
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error Resume Next
    If mUI.FechandoPasta() Then Exit Sub   ' a pasta vai fechar: nao refaz a vista nem religa o vigia
    Application.EnableEvents = False
    mSeguranca.RestaurarAposGravar  ' devolve a vista da sessao (o disco ficou trancado)
    Application.EnableEvents = True
    mUI.GrafVigiaLigar              ' religa com o nome atual (sai calado fora do Painel)
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim salvo As Boolean
    On Error Resume Next
    ' Auditoria 04/10/2026 (A-16): o "Cancelar" da pergunta "Salvar?" do Excel deixa a aparencia do Excel
    ' devolvida ate o proximo login. Perguntar AQUI foi tentado e revertido: o BeforeClose nao sabe se
    ' quem fecha pediu Close SaveChanges:=False (automacao, outra macro) -- a pergunta aparecia mesmo
    ' assim e travava o fechamento (medido na bateria). Limitacao cosmetica declarada.
    salvo = Me.Saved
    mUI.GrafVigiaEncerrar           ' ADR-061: tique pendente/vencido com a pasta fechada REABRE o arquivo
    SystemLook False
    ' ADR-063: devolver a aparencia do Excel marca a pasta como alterada e o Excel perguntava
    ' "Quer salvar?" mesmo sem nada do usuario mudado. Pasta que estava salva continua salva;
    ' com alteracao de verdade (Saved = False) a pergunta continua.
    If salvo Then Me.Saved = True
End Sub
