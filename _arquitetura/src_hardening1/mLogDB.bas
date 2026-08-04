Attribute VB_Name = "mLogDB"
Option Explicit
' ===== TABELAS DE LOG DENTRO DO DB_RESULTADOS =====
'
' Desenho definido pelo gestor. Duas camadas com papeis distintos:
'
'   Audit_Log        Event Store geral, append-only, ENCADEADO POR HASH.
'                    E a fonte da verdade e a prova de integridade.
'
'   LOG_Resultados   Estrutura fisica desacoplada dentro do banco, por origem.
'   LOG_Registros    Preserva as referencias historicas ja estruturadas e deixa
'                    o auditor filtrar por planilha de origem sem cruzar dados.
'
' As duas tabelas do banco NAO sao fonte independente: cada linha carrega o
' ID_Auditoria do evento que a gerou. Divergiu do Event Store, vale o Event
' Store -- e a divergencia e detectavel justamente por causa do ID.
'
' POSICAO. Blocos deslocados a direita, nunca abaixo dos dados: CarregarDB le o
' bloco A:G inteiro e UltimaLinhaBanco mede a coluna A.
'
' O DB_Resultados ja e denso: A:G e o banco, H tem o botao de lancamento, K:AZ
' guarda os DADOS INTERFACEADOS (com K1:AZ1 mescladas) e BA:BD tem LoteCore,
' 1aOc e RunUnico, referenciados por nome em formulas. A primeira coluna
' realmente livre e BE. Dai:
'   LOG_Resultados  BG..BY
'   LOG_Registros   CB..CT

Public Const LOGDB_CAB As Long = 3          ' linha de cabecalho
Public Const LOGDB_R0 As Long = 4           ' primeira linha de dado
Public Const LOGDB_C_RESULTADOS As Long = 59   ' BG
Public Const LOGDB_C_REGISTROS As Long = 80    ' CB
Public Const LOGDB_NCOL As Long = 19

Public Const ORIG_RESULTADOS As String = "Resultados"
Public Const ORIG_REGISTROS As String = "Registros"

' Coluna inicial do bloco conforme a aba de origem.
Private Function ColunaBase(ByVal origem As String) As Long
    If UCase$(Trim$(origem)) = UCase$(ORIG_REGISTROS) Then
        ColunaBase = LOGDB_C_REGISTROS
    Else
        ColunaBase = LOGDB_C_RESULTADOS
    End If
End Function

' Ultima linha COM DADO no bloco indicado.
'
' Percorre de baixo para cima procurando ID preenchido, em vez de confiar em
' End(xlUp): um ListObject nasce com uma linha de dado em branco que nao e
' vazia para o Excel, e o primeiro registro cairia na linha errada.
Public Function UltimaLinhaLogDB(ByVal origem As String) As Long
    Dim ws As Worksheet, c0 As Long, r As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    c0 = ColunaBase(origem)
    r = ws.Cells(ws.Rows.Count, c0).End(xlUp).Row
    Do While r >= LOGDB_R0
        If Len(Trim$(CStr(ws.Cells(r, c0).Value))) > 0 Then Exit Do
        r = r - 1
    Loop
    If r < LOGDB_R0 Then r = LOGDB_R0 - 1
    UltimaLinhaLogDB = r
End Function

' Acrescenta uma linha ao log da origem indicada.
'
' idAuditoria vem de Auditar: as duas camadas gravam o MESMO identificador, e e
' por ele que se prova que uma nao contradiz a outra.
Public Sub RegistrarLogDB(ByVal origem As String, ByVal idAuditoria As String, _
                          ByVal tipoOperacao As String, _
                          ByVal run As Long, ByVal dtCorrida As Variant, _
                          ByVal nivel As Long, ByVal analito As String, _
                          ByVal lote As String, ByVal resultado As Variant, _
                          ByVal stAnterior As String, ByVal stNovo As String, _
                          ByVal parecer As String)

    Dim ws As Worksheet, c0 As Long, lin As Long, prot As Boolean
    Set ws = ThisWorkbook.Sheets(BANCO)
    c0 = ColunaBase(origem)
    lin = UltimaLinhaLogDB(origem) + 1

    prot = ws.ProtectContents
    If prot Then ws.Unprotect Password:="qcini2025"

    ws.Cells(lin, c0 + 0).NumberFormat = "@"

    ' UMA escrita de bloco em vez de 19 escritas de celula: cada acesso a Range
    ' custa uma travessia COM, e o log grava a cada exclusao.
    Dim linha() As Variant
    ReDim linha(1 To 1, 1 To LOGDB_NCOL)
    linha(1, 1) = idAuditoria
    linha(1, 2) = Now
    linha(1, 3) = tipoOperacao
    linha(1, 4) = origem
    linha(1, 5) = IIf(run = 0, "", run)
    linha(1, 6) = dtCorrida
    linha(1, 7) = IIf(nivel = 0, "", nivel)
    linha(1, 8) = analito
    linha(1, 9) = lote
    linha(1, 10) = resultado
    linha(1, 11) = stAnterior
    linha(1, 12) = stNovo
    linha(1, 13) = parecer
    linha(1, 14) = UsuarioSistema()
    linha(1, 15) = Application.UserName
    linha(1, 16) = Environ$("USERNAME")
    linha(1, 17) = Environ$("COMPUTERNAME")
    linha(1, 18) = ThisWorkbook.Name
    linha(1, 19) = VERSAO_SISTEMA
    ws.Range(ws.Cells(lin, c0), ws.Cells(lin, c0 + LOGDB_NCOL - 1)).Value = linha
    ws.Cells(lin, c0 + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(lin, c0 + 5).NumberFormat = "dd/mm/yyyy"

    ExpandirTabelaLog ws, origem, lin

    If prot Then
        ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True, _
                   AllowFiltering:=True, AllowSorting:=True
    End If
End Sub

' Mantem o ListObject cobrindo a linha nova: sem isso o filtro do auditor nao
' alcanca os registros recentes, que sao os que ele procura primeiro.
Private Sub ExpandirTabelaLog(ByVal ws As Worksheet, ByVal origem As String, ByVal lin As Long)
    On Error Resume Next
    Dim lo As ListObject, nome As String, c0 As Long
    If UCase$(Trim$(origem)) = UCase$(ORIG_REGISTROS) Then
        nome = "tblLogRegistros"
    Else
        nome = "tblLogResultados"
    End If
    Set lo = ws.ListObjects(nome)
    If lo Is Nothing Then Exit Sub
    c0 = ColunaBase(origem)
    If lin > lo.Range.Row + lo.Range.Rows.Count - 1 Then
        lo.Resize ws.Range(ws.Cells(LOGDB_CAB, c0), ws.Cells(lin, c0 + LOGDB_NCOL - 1))
    End If
End Sub

' Quantos registros ha em cada log. Usado pela suite de verificacao.
Public Function ContarLogDB(ByVal origem As String) As Long
    ContarLogDB = UltimaLinhaLogDB(origem) - LOGDB_R0 + 1
    If ContarLogDB < 0 Then ContarLogDB = 0
End Function
