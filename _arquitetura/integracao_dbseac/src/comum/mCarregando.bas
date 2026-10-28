Attribute VB_Name = "mCarregando"
Option Explicit
' ============================================================================
'  mCarregando -- indicador de processamento das operacoes demoradas (ADR-084)
'
'  Uma janela so (frmCarregando, sem modal), aberta pela ENTRADA da operacao
'  (botao ATUALIZAR DADOS, inativar/reativar pelo grafico, login) e avancada
'  pelas etapas REAIS do nucleo (mIntegracao.AtualizarDadosCore). Sem timer e sem
'  porcentagem inventada: com o total de etapas conhecido a barra mostra
'  "Etapa k de n"; sem total, so a etapa atual (a barra anda um passo a cada
'  etapa, sem prometer quanto falta).
'
'  Fechamento garantido: a entrada fecha (sucesso ou erro) e, de reserva,
'  mApp.fim (ao zerar o contador de operacao) e mApp.Reset tambem fecham -- uma
'  falha no meio nunca deixa a janela na tela.
'
'  Sem DoEvents: o redesenho e so o da propria janela (Repaint). Um DoEvents
'  aqui deixaria o clique do usuario disparar outra macro no meio da operacao.
'  Automacao (Application.Interactive = False) ou Excel invisivel: a janela NAO
'  abre, mas as etapas sao registradas (Historico) -- o QA confere a sequencia e
'  o fechamento sem interface. A janela e chamada por nome (VBA.UserForms.Add):
'  o projeto compila mesmo sem ela instalada.
' ============================================================================
Private mFrm As Object
Private mAberto As Boolean
Private mTitulo As String
Private mTotal As Long
Private mPasso As Long
Private mHist As String

' Abre (ou reaproveita) o indicador. total = numero de etapas conhecido (0 = indeterminado).
' Ja aberto: nao abre outra janela -- a operacao de dentro so avanca as etapas.
Public Sub Abrir(ByVal titulo As String, Optional ByVal total As Long = 0)
    On Error Resume Next
    If mAberto Then Exit Sub
    mAberto = True
    mTitulo = titulo
    mTotal = total
    mPasso = 0
    mHist = "ABRIR:" & titulo
    If Not Application.Interactive Or Not Application.Visible Then Exit Sub
    Set mFrm = VBA.UserForms.Add("frmCarregando")
    If mFrm Is Nothing Then Exit Sub
    mFrm.Mostrar titulo, "Iniciando...", 0, mTotal
    mFrm.Show vbModeless
    mFrm.Repaint
End Sub

' Proxima etapa (texto curto, em portugues). Sem indicador aberto: nada (o nucleo chamado por automacao).
Public Sub Etapa(ByVal texto As String)
    On Error Resume Next
    If Not mAberto Then Exit Sub
    mPasso = mPasso + 1
    If mTotal > 0 And mPasso > mTotal Then mTotal = mPasso       ' nunca "Etapa 7 de 6"
    mHist = mHist & "|" & texto
    Application.StatusBar = mTitulo & ": " & texto
    If mFrm Is Nothing Then Exit Sub
    mFrm.Mostrar mTitulo, texto, mPasso, mTotal
    mFrm.Repaint
End Sub

' Fecha. Pode ser chamado mais de uma vez (entrada + mApp.fim): a 2a chamada nao faz nada.
Public Sub Fechar()
    On Error Resume Next
    If mAberto Then mHist = mHist & "|FECHAR"
    mAberto = False
    mTotal = 0
    mPasso = 0
    If Not mFrm Is Nothing Then
        mFrm.Hide
        Unload mFrm
    End If
    Set mFrm = Nothing
End Sub

Public Function Aberto() As Boolean
    Aberto = mAberto
End Function

' Para o QA: "ABRIR:<titulo>|etapa 1|etapa 2|...|FECHAR" da ultima operacao.
Public Function Historico() As String
    Historico = mHist
End Function

' Para o QA: a janela esta na tela (so com Interactive = True).
Public Function JanelaVisivel() As Boolean
    On Error Resume Next
    If Not mFrm Is Nothing Then JanelaVisivel = mFrm.Visible
End Function
