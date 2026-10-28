Attribute VB_Name = "mCadastro"
Option Explicit
' ============================================================================
'  mCadastro -- cadastro central de profissionais do QC_INI (ADR-084)
'  Arquivo: QC_INI_cadastrar usuarios.xlsm, aba REGISTRO_DE_USUARIOS,
'  tabela tb_registro_usuarios (ordem fixa, lida por posicao pelos QCs):
'    A FUNCIONARIOS | B LOGIN | C CARGO | D SENHA_HASH | E ASSINATURA |
'    F CONSELHO OU MAT. | G SITUACAO
'
'  - D guarda so o hash PBKDF2-SHA256 com salt (mSenha). Senha nunca vai para
'    celula, log ou mensagem; os campos de senha do formulario sao limpos.
'  - E comeca VAZIA: o administrador escreve depois o nome do arquivo da imagem
'    (ex.: joao.silva.png) que esta na pasta "assinaturas" ao lado deste
'    arquivo (ou o caminho completo). Os QCs carregam a imagem dali.
'  - G = ATIVO / INATIVO. Inativar nao apaga a linha: as liberacoes passadas
'    continuam identificadas (o QC guardou o texto da identidade na hora).
'  - Toda gravacao exige o login e a senha de um ADMIN ativo com senha. Primeiro
'    acesso (nenhum ADMIN com senha): so a linha de um ADMIN pode receber a
'    primeira senha, informada por ele mesmo (login do administrador = o da linha).
'  - Trilha na aba LOG_CADASTRO (data, usuario do Windows, administrador, acao,
'    login alvo, detalhe) -- sem senha nem hash.
'  Arquivo de Excel nao e fronteira de seguranca: o controle de acesso real e a
'  permissao da pasta onde este arquivo fica (so administradores gravam).
' ============================================================================
Public Const ABA As String = "REGISTRO_DE_USUÁRIOS"
Public Const TABELA As String = "tb_registro_usuários"
Private Const ABA_LOG As String = "LOG_CADASTRO"
Private Const C_NOME As Long = 1
Private Const C_LOGIN As Long = 2
Private Const C_CARGO As Long = 3
Private Const C_HASH As Long = 4
Private Const C_ASSIN As Long = 5
Private Const C_CONS As Long = 6
Private Const C_SIT As Long = 7
Public Const CARGO_ADMIN As String = "ADMIN"
Private Const SENHA_MIN As Long = 8

Public Function Cargos() As Variant
    Cargos = Array(CARGO_ADMIN, "Analista de Laboratório", "Técnico de Laboratório")
End Function

Public Sub AbrirCadastro()
    Dim f As Object
    Set f = VBA.UserForms.Add("frmCadastroUsuario")
    f.Show
    Unload f
End Sub

' Nucleo sem dialogo (o formulario e o QA chamam este). Login novo = inclui; login existente = redefine a senha e
' atualiza nome/cargo/conselho informados. Devolve "OK|CRIADO|<linha>", "OK|ATUALIZADO|<linha>" ou "ERRO|<motivo>".
Public Function CadastrarOuRedefinir(ByVal nome As String, ByVal login As String, ByVal cargo As String, _
                                     ByVal senha As String, ByVal confirmacao As String, ByVal conselho As String, _
                                     ByVal admLogin As String, ByVal admSenha As String) As String
    Dim lo As ListObject, r As Long, novo As Boolean, recusa As String, autor As String, lr As ListRow, sE As String
    Dim cargoAlvo As String, novoHash As String
    On Error GoTo falha
    nome = Trim$(nome): login = Trim$(login): cargo = CargoValido(cargo): conselho = Trim$(conselho)
    Set lo = TabelaUsuarios()
    r = LinhaDoLogin(lo, login)
    novo = (r = 0)
    If novo Then cargoAlvo = cargo Else cargoAlvo = CStr(lo.DataBodyRange.Cells(r, C_CARGO).Value)
    If Len(login) = 0 Then
        recusa = "Informe o login."
    ElseIf InStr(login, " ") > 0 Or InStr(login, "|") > 0 Then
        recusa = "O login não pode ter espaço nem '|'."
    ElseIf novo And Len(nome) = 0 Then
        recusa = "Informe o nome completo."
    ElseIf novo And Len(cargo) = 0 Then
        recusa = "Escolha o cargo na lista."
    ElseIf Len(senha) = 0 Then
        recusa = "Informe a senha."
    ElseIf senha <> confirmacao Then
        recusa = "A senha e a confirmação não são iguais."
    ElseIf Not SenhaForte(senha) Then
        recusa = "A senha precisa de pelo menos " & SENHA_MIN & " caracteres, com letras e números."
    Else
        recusa = Autorizar(lo, admLogin, admSenha, login, cargoAlvo, autor)
    End If
    If Len(recusa) = 0 Then novoHash = mSenha.GerarHashSenha(senha)     ' a senha so existe ate aqui
    senha = "": confirmacao = "": admSenha = ""
    If Len(recusa) > 0 Then
        RegistrarLog autor, "RECUSADO", login, recusa
        CadastrarOuRedefinir = "ERRO|" & recusa
        Exit Function
    End If
    If novo Then
        Set lr = lo.ListRows.Add
        r = lr.Index
        lo.DataBodyRange.Cells(r, C_NOME).Value = nome
        lo.DataBodyRange.Cells(r, C_LOGIN).Value = login
        lo.DataBodyRange.Cells(r, C_CARGO).Value = cargo
        lo.DataBodyRange.Cells(r, C_ASSIN).Value = ""             ' a imagem e cadastrada depois pelo administrador
        lo.DataBodyRange.Cells(r, C_CONS).Value = conselho
        lo.DataBodyRange.Cells(r, C_SIT).Value = "ATIVO"
    Else
        If Len(nome) > 0 Then lo.DataBodyRange.Cells(r, C_NOME).Value = nome
        If Len(cargo) > 0 Then lo.DataBodyRange.Cells(r, C_CARGO).Value = cargo
        If Len(conselho) > 0 Then lo.DataBodyRange.Cells(r, C_CONS).Value = conselho
        If Len(Trim$(CStr(lo.DataBodyRange.Cells(r, C_SIT).Value))) = 0 Then lo.DataBodyRange.Cells(r, C_SIT).Value = "ATIVO"
    End If
    lo.DataBodyRange.Cells(r, C_HASH).NumberFormat = "@"
    lo.DataBodyRange.Cells(r, C_HASH).Value = novoHash
    RegistrarLog autor, IIf(novo, "USUARIO_CRIADO", "SENHA_REDEFINIDA"), login, _
                 "cargo: " & lo.DataBodyRange.Cells(r, C_CARGO).Value & " | linha " & r
    CadastrarOuRedefinir = "OK|" & IIf(novo, "CRIADO", "ATUALIZADO") & "|" & r
    Exit Function
falha:
    sE = Err.Description
    CadastrarOuRedefinir = "ERRO|Não foi possível gravar (detalhe: " & sE & ")"
End Function

' Situacao ATIVO / INATIVO de um login (a linha e o historico ficam). "OK|<situacao>" ou "ERRO|<motivo>".
Public Function DefinirSituacao(ByVal login As String, ByVal ativo As Boolean, ByVal admLogin As String, _
                                ByVal admSenha As String) As String
    Dim lo As ListObject, r As Long, recusa As String, autor As String
    On Error GoTo falha
    Set lo = TabelaUsuarios()
    r = LinhaDoLogin(lo, Trim$(login))
    If r = 0 Then DefinirSituacao = "ERRO|Login não encontrado.": Exit Function
    recusa = Autorizar(lo, admLogin, admSenha, "", "", autor)
    If Len(recusa) = 0 And Not ativo And UCase$(Trim$(login)) = UCase$(Trim$(admLogin)) Then
        recusa = "O administrador não pode inativar a si mesmo."
    End If
    admSenha = ""
    If Len(recusa) > 0 Then
        RegistrarLog autor, "RECUSADO", login, recusa
        DefinirSituacao = "ERRO|" & recusa
        Exit Function
    End If
    lo.DataBodyRange.Cells(r, C_SIT).Value = IIf(ativo, "ATIVO", "INATIVO")
    RegistrarLog autor, IIf(ativo, "USUARIO_REATIVADO", "USUARIO_INATIVADO"), lo.DataBodyRange.Cells(r, C_LOGIN).Value, ""
    DefinirSituacao = "OK|" & IIf(ativo, "ATIVO", "INATIVO")
    Exit Function
falha:
    DefinirSituacao = "ERRO|Não foi possível gravar (detalhe: " & Err.Description & ")"
End Function

' ---------------------------------------------------------------- regras
' "" = autorizado (autor = login do ADMIN que autorizou). Primeiro acesso: nenhum ADMIN com senha -> so a 1a senha de
' uma linha ADMIN, pelo proprio (admLogin = loginAlvo).
Private Function Autorizar(ByVal lo As ListObject, ByVal admLogin As String, ByVal admSenha As String, _
                           ByVal loginAlvo As String, ByVal cargoAlvo As String, ByRef autor As String) As String
    Dim r As Long, h As String
    autor = Trim$(admLogin)
    If Not ExisteAdminComSenha(lo) Then
        If Len(loginAlvo) > 0 And UCase$(cargoAlvo) = CARGO_ADMIN And UCase$(Trim$(admLogin)) = UCase$(loginAlvo) Then
            autor = loginAlvo & " (primeiro acesso)"
            Exit Function
        End If
        Autorizar = "Primeiro acesso: defina antes a senha de um ADMIN (login do administrador = o login dele)."
        Exit Function
    End If
    r = LinhaDoLogin(lo, autor)
    If r = 0 Then Autorizar = "Administrador ou senha inválidos.": Exit Function
    If UCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, C_CARGO).Value))) <> CARGO_ADMIN Then
        Autorizar = "Só um ADMIN pode cadastrar ou alterar usuários.": Exit Function
    End If
    If UCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, C_SIT).Value))) = "INATIVO" Then
        Autorizar = "Administrador inativo.": Exit Function
    End If
    h = CStr(lo.DataBodyRange.Cells(r, C_HASH).Value)
    If Not mSenha.VerificarSenha(admSenha, h) Then Autorizar = "Administrador ou senha inválidos."
End Function

Private Function ExisteAdminComSenha(ByVal lo As ListObject) As Boolean
    Dim i As Long
    If lo.ListRows.Count = 0 Then Exit Function
    For i = 1 To lo.ListRows.Count
        If UCase$(Trim$(CStr(lo.DataBodyRange.Cells(i, C_CARGO).Value))) = CARGO_ADMIN And _
           UCase$(Trim$(CStr(lo.DataBodyRange.Cells(i, C_SIT).Value))) <> "INATIVO" And _
           mSenha.EhHashPbkdf2(CStr(lo.DataBodyRange.Cells(i, C_HASH).Value)) Then
            ExisteAdminComSenha = True
            Exit Function
        End If
    Next i
End Function

Private Function SenhaForte(ByVal s As String) As Boolean
    Dim i As Long, ch As String, letra As Boolean, digito As Boolean
    If Len(s) < SENHA_MIN Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Then digito = True
        If UCase$(ch) <> LCase$(ch) Then letra = True
    Next i
    SenhaForte = letra And digito
End Function

Private Function CargoValido(ByVal c As String) As String
    Dim x As Variant
    For Each x In Cargos()
        If UCase$(Trim$(c)) = UCase$(CStr(x)) Then CargoValido = CStr(x): Exit Function
    Next x
End Function

Private Function TabelaUsuarios() As ListObject
    Set TabelaUsuarios = ThisWorkbook.Worksheets(ABA).ListObjects(TABELA)
End Function

' Linha (1..n) do login na tabela, sem diferenciar maiusculas; 0 = nao ha.
Private Function LinhaDoLogin(ByVal lo As ListObject, ByVal login As String) As Long
    Dim i As Long
    If lo.ListRows.Count = 0 Or Len(Trim$(login)) = 0 Then Exit Function
    For i = 1 To lo.ListRows.Count
        If UCase$(Trim$(CStr(lo.DataBodyRange.Cells(i, C_LOGIN).Value))) = UCase$(Trim$(login)) Then
            LinhaDoLogin = i
            Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------- trilha
Private Sub RegistrarLog(ByVal autor As String, ByVal acao As String, ByVal alvo As String, ByVal detalhe As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_LOG)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 3 Then r = 3
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(r, 2).Value = Environ$("USERNAME")
    ws.Cells(r, 3).Value = autor
    ws.Cells(r, 4).Value = acao
    ws.Cells(r, 5).Value = alvo
    ws.Cells(r, 6).Value = detalhe
End Sub
