# gerar_mEspecificacoes.ps1 - especializa o motor de especificacoes por produto
#
# A Bioquimica continua usando a fonte de producao sem alteracao. A Hematologia
# recebe o mesmo contrato, mas sem fonte/rigor implicitos e com fallback
# operacional para Analitos!Q:R enquanto o banco formal ainda nao foi preenchido.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).

param(
    [Parameter(Mandatory = $true)][string]$Producao,
    [Parameter(Mandatory = $true)][string]$Produto,
    [Parameter(Mandatory = $true)][string]$Saida
)

$ErrorActionPreference = 'Stop'
$enc = [System.Text.Encoding]::Default
$linhas = [System.Collections.ArrayList]@([System.IO.File]::ReadAllLines($Producao, $enc))

function Linhas-Texto([string]$Texto) {
    return @($Texto.Trim("`r", "`n") -split "`r?`n")
}

function Substituir-Rotina {
    param(
        [System.Collections.ArrayList]$Codigo,
        [string]$Nome,
        [string]$NovoCodigo
    )

    $inicio = -1
    $fim = -1
    $tipo = ''
    for ($i = 0; $i -lt $Codigo.Count; $i++) {
        if ($Codigo[$i] -match "^\s*(Public|Private|Friend)?\s*(Sub|Function)\s+$([regex]::Escape($Nome))\b") {
            if ($inicio -ge 0) { throw "Rotina duplicada: $Nome" }
            $inicio = $i
            $tipo = $Matches[2]
        }
    }
    if ($inicio -lt 0) { throw "Rotina nao encontrada: $Nome" }

    for ($i = $inicio + 1; $i -lt $Codigo.Count; $i++) {
        if ($Codigo[$i] -match "^\s*End\s+$tipo\s*$") { $fim = $i; break }
    }
    if ($fim -lt 0) { throw "Fim da rotina nao encontrado: $Nome" }

    $Codigo.RemoveRange($inicio, $fim - $inicio + 1)
    $novas = Linhas-Texto $NovoCodigo
    for ($i = $novas.Count - 1; $i -ge 0; $i--) { $Codigo.Insert($inicio, $novas[$i]) }
}

if ($Produto -eq 'Bioquimica') {
    [System.IO.File]::WriteAllLines($Saida, $linhas.ToArray(), $enc)
    "mEspecificacoes: Bioquimica preservada sem especializacao"
    "saida: $Saida"
    exit 0
}
if ($Produto -ne 'Hematologia') {
    throw "Produto sem perfil de especificacoes: $Produto"
}

$idxConst = -1
for ($i = 0; $i -lt $linhas.Count; $i++) {
    if ($linhas[$i] -match '^Public Const ENGE_R0 As Long') { $idxConst = $i; break }
}
if ($idxConst -lt 0) { throw 'Constante ENGE_R0 nao encontrada.' }
$linhas.Insert($idxConst + 1, 'Private Const SENHA_ESPEC As String = "qcini2025"')

Substituir-Rotina $linhas 'NormRigor' @'
Public Function NormRigor(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, ChrW(211), "O")
    t = Replace(t, ChrW(243), "O")
    If Left$(t, 3) = "OTI" Then NormRigor = "OTI": Exit Function
    If Left$(t, 3) = "DES" Then NormRigor = "DES": Exit Function
    If Left$(t, 3) = "MIN" Then NormRigor = "MIN"
End Function
'@

Substituir-Rotina $linhas 'FatorCV' @'
Public Function FatorCV(ByVal rigor As String) As Double
    Select Case NormRigor(rigor)
        Case "OTI": FatorCV = 0.25
        Case "DES": FatorCV = 0.5
        Case "MIN": FatorCV = 0.75
    End Select
End Function
'@

Substituir-Rotina $linhas 'FatorBias' @'
Public Function FatorBias(ByVal rigor As String) As Double
    Select Case NormRigor(rigor)
        Case "OTI": FatorBias = 0.125
        Case "DES": FatorBias = 0.25
        Case "MIN": FatorBias = 0.375
    End Select
End Function
'@

Substituir-Rotina $linhas 'MetasDaLinha' @'
Public Function MetasDaLinha(ByVal modelo As String, ByVal etp As Variant, _
                             ByVal cvi As Variant, ByVal cvg As Variant, _
                             ByVal rigor As String, ByVal cvtp As Variant, _
                             ByVal biastp As Variant) As String
    Dim vCV As Double, vBias As Double, vETP As Double
    Dim temCV As Boolean, temBias As Boolean, temETP As Boolean

    Select Case UCase$(Trim$(modelo))
        Case MOD_ETP
            If IsNumeric(etp) And Len(Trim$(CStr(etp))) > 0 Then
                vETP = CDbl(etp): temETP = True
                vCV = vETP / 3: temCV = True
            End If
        Case MOD_VB
            If NormRigor(rigor) <> "" Then
                If IsNumeric(cvi) And IsNumeric(cvg) Then
                    If Len(Trim$(CStr(cvi))) > 0 And Len(Trim$(CStr(cvg))) > 0 Then
                        vCV = CDbl(cvi) * FatorCV(rigor): temCV = True
                        vBias = Sqr(CDbl(cvi) ^ 2 + CDbl(cvg) ^ 2) * FatorBias(rigor): temBias = True
                        vETP = vBias + 1.65 * vCV: temETP = True
                    End If
                End If
            End If
        Case MOD_CVBIAS
            If IsNumeric(cvtp) And IsNumeric(biastp) Then
                If Len(Trim$(CStr(cvtp))) > 0 And Len(Trim$(CStr(biastp))) > 0 Then
                    vCV = CDbl(cvtp): temCV = True
                    vBias = CDbl(biastp): temBias = True
                    vETP = vBias + 1.65 * vCV: temETP = True
                End If
            End If
    End Select

    MetasDaLinha = IIf(temCV, Trim$(Str$(vCV)), "") & "|" & _
                   IIf(temBias, Trim$(Str$(vBias)), "") & "|" & _
                   IIf(temETP, Trim$(Str$(vETP)), "")
End Function
'@

Substituir-Rotina $linhas 'FontePadrao' @'
Private Function ModeloSuportado(ByVal modelo As String) As Boolean
    Select Case UCase$(Trim$(modelo))
        Case MOD_ETP, MOD_VB, MOD_CVBIAS
            ModeloSuportado = True
    End Select
End Function

Public Function FontePadrao() As String
    Dim ws As Worksheet, v As String, i As Long
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    v = Trim$(CStr(ws.Range("B2").Value))
    If v = "" Then Exit Function

    For i = 7 To 40
        If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = UCase$(v) Then
            If ModeloSuportado(CStr(ws.Cells(i, 2).Value)) Then FontePadrao = v
            Exit Function
        End If
    Next i
fim:
End Function
'@

Substituir-Rotina $linhas 'RigorPadrao' @'
Public Function RigorPadrao() As String
    Dim ws As Worksheet, v As String
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    v = UCase$(Trim$(CStr(ws.Range("B3").Value)))
    If Left$(v, 3) = "OTI" Then RigorPadrao = "OTI": Exit Function
    If Left$(v, 3) = "DES" Then RigorPadrao = "DES": Exit Function
    If Left$(v, 3) = "MIN" Then RigorPadrao = "MIN"
fim:
End Function
'@

Substituir-Rotina $linhas 'ModeloDaFonte' @'
Public Function ModeloDaFonte(ByVal fonte As String) As String
    Dim ws As Worksheet, i As Long, modelo As String
    If Trim$(fonte) = "" Then Exit Function
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets("Cfg_Especificacoes")
    For i = 7 To 40
        If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = UCase$(Trim$(fonte)) Then
            modelo = UCase$(Trim$(CStr(ws.Cells(i, 2).Value)))
            If ModeloSuportado(modelo) Then ModeloDaFonte = modelo
            Exit Function
        End If
    Next i
fim:
End Function
'@

Substituir-Rotina $linhas 'GravarEspec' @'
Public Function GravarEspec(ByVal ano As Long, ByVal fonte As String, ByVal analito As String, _
                            ByVal etp As Variant, ByVal cvi As Variant, ByVal cvg As Variant, _
                            ByVal rigor As String, ByVal cvtp As Variant, ByVal biastp As Variant) As String
    Dim ws As Worksheet, dados As Variant, i As Long, lin As Long, modelo As String
    Dim metas As String, partes() As String

    fonte = Trim$(fonte)
    analito = Trim$(analito)
    If ano < 1900 Or ano > 2200 Then Err.Raise vbObjectError + 620, "mEspecificacoes.GravarEspec", "Ano invalido."
    If fonte = "" Then Err.Raise vbObjectError + 621, "mEspecificacoes.GravarEspec", "Cadastre e selecione uma fonte."
    If analito = "" Then Err.Raise vbObjectError + 622, "mEspecificacoes.GravarEspec", "Analito obrigatorio."

    modelo = ModeloDaFonte(fonte)
    If modelo = "" Then Err.Raise vbObjectError + 623, "mEspecificacoes.GravarEspec", "Fonte sem modelo suportado."
    If modelo = MOD_VB And NormRigor(rigor) = "" Then
        Err.Raise vbObjectError + 624, "mEspecificacoes.GravarEspec", "Selecione um rigor valido para a fonte VB."
    End If

    metas = MetasDaLinha(modelo, etp, cvi, cvg, rigor, cvtp, biastp)
    partes = Split(metas, "|")
    If Len(partes(0)) = 0 And Len(partes(1)) = 0 And Len(partes(2)) = 0 Then
        Err.Raise vbObjectError + 625, "mEspecificacoes.GravarEspec", "A especificacao nao contem entradas validas para o modelo."
    End If

    Set ws = ThisWorkbook.Sheets(ESPEC)
    dados = CarregarEspec()
    lin = 0
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If CStr(dados(i, ES_ANO)) = CStr(ano) _
               And UCase$(Trim$(CStr(dados(i, ES_FONTE)))) = UCase$(fonte) _
               And UCase$(Trim$(CStr(dados(i, ES_ANALITO)))) = UCase$(analito) Then
                lin = ESPEC_R0 + i - 1
                Exit For
            End If
        Next i
    End If

    Dim novoID As String, prot As Boolean
    If lin = 0 Then
        lin = UltimaLinhaEspec() + 1
        If lin < ESPEC_R0 Then lin = ESPEC_R0
        novoID = "ESP-" & Format$(lin - ESPEC_R0 + 1, "000000")
    Else
        novoID = CStr(ws.Cells(lin, ES_ID).Value)
    End If

    prot = ws.ProtectContents
    On Error GoTo falha
    If prot Then ws.Unprotect Password:=SENHA_ESPEC
    ws.Cells(lin, ES_ID).Value = novoID
    ws.Cells(lin, ES_ANO).Value = ano
    ws.Cells(lin, ES_FONTE).Value = fonte
    ws.Cells(lin, ES_MODELO).Value = modelo
    ws.Cells(lin, ES_ANALITO).Value = analito
    ws.Cells(lin, ES_ETP).Value = etp
    ws.Cells(lin, ES_CVI).Value = cvi
    ws.Cells(lin, ES_CVG).Value = cvg
    ws.Cells(lin, ES_RIGOR).Value = IIf(Trim$(rigor) = "", "", NormRigor(rigor))
    ws.Cells(lin, ES_CVTP).Value = cvtp
    ws.Cells(lin, ES_BIASTP).Value = biastp
    ws.Cells(lin, ES_ATIVO).Value = "Sim"
    ws.Cells(lin, ES_USUARIO).Value = Environ$("USERNAME")
    ws.Cells(lin, ES_DATACAD).Value = Now
    If prot Then ws.Protect Password:=SENHA_ESPEC, UserInterfaceOnly:=True

    InvalidarCacheEspec
    RegistrarLog "ESPEC_GRAVADA", novoID & " " & fonte & " " & ano & " " & analito
    GravarEspec = novoID
    Exit Function
falha:
    Dim numErro As Long, descErro As String
    numErro = Err.Number: descErro = Err.Description
    On Error Resume Next
    If prot Then ws.Protect Password:=SENHA_ESPEC, UserInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise numErro, "mEspecificacoes.GravarEspec", descErro
End Function
'@

Substituir-Rotina $linhas 'AtualizarEngEspec' @'
Private Function EspecLegada(ByVal analito As String, ByRef fonte As String, ByRef etp As Variant) As Boolean
    Dim ws As Worksheet, i As Long, v As Variant
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets("Analitos")
    For i = 4 To 43
        If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = UCase$(Trim$(analito)) Then
            v = ws.Cells(i, 18).Value
            If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
                etp = CDbl(v)
                fonte = Trim$(CStr(ws.Cells(i, 17).Value))
                If fonte = "" Or fonte = "-" Then fonte = "LEGADO_ANALITOS"
                EspecLegada = True
            End If
            Exit Function
        End If
    Next i
fim:
End Function

Public Sub AtualizarEngEspec()
    Dim ws As Worksheet, c As Collection, i As Long
    Dim ano As Long, fonteFormal As String, fonteLegada As String
    Dim r As String, p() As String, etpLegado As Variant
    Dim buf() As Variant, n As Long, prot As Boolean

    Set ws = ThisWorkbook.Sheets(ENGE)
    fonteFormal = FontePadrao()
    ano = AnoDeContexto()
    Set c = ListaAnalitos()
    n = c.Count

    prot = ws.ProtectContents
    On Error GoTo falha
    If prot Then ws.Unprotect Password:=SENHA_ESPEC
    ws.Range(ws.Cells(ENGE_R0, 1), ws.Cells(ENGE_R0 + 39, 9)).ClearContents
    ws.Range("B1").Value = ano
    ws.Range("D1").Value = fonteFormal
    If n = 0 Then GoTo concluir

    ReDim buf(1 To n, 1 To 9)
    For i = 1 To n
        buf(i, 1) = c(i)
        r = ResolverEspec(CStr(c(i)), ano, fonteFormal)
        If fonteFormal <> "" And Left$(r, 2) = "OK" Then
            p = Split(r, "|")
            buf(i, 2) = p(4)
            buf(i, 3) = Val(p(5))
            buf(i, 4) = IIf(Len(p(1)) > 0, Val(p(1)), "")
            buf(i, 5) = IIf(Len(p(2)) > 0, Val(p(2)), "")
            buf(i, 6) = IIf(Len(p(3)) > 0, Val(p(3)), "")
            buf(i, 7) = p(6)
            buf(i, 8) = p(7)
            buf(i, 9) = "CADASTRADA"
        Else
            fonteLegada = "": etpLegado = Empty
            If EspecLegada(CStr(c(i)), fonteLegada, etpLegado) Then
                buf(i, 2) = fonteLegada
                buf(i, 6) = etpLegado
                buf(i, 9) = "FALLBACK LEGADO"
            ElseIf fonteFormal = "" Then
                buf(i, 9) = "FONTE NAO CONFIGURADA"
            ElseIf r = "NAO_CADASTRADA" Then
                buf(i, 2) = fonteFormal
                buf(i, 9) = "NAO CADASTRADA"
            Else
                p = Split(r, "|")
                buf(i, 2) = fonteFormal
                If UBound(p) >= 1 Then buf(i, 8) = p(1)
                buf(i, 9) = "FORMAL INVALIDA; SEM FALLBACK"
            End If
        End If
    Next i
    ws.Range(ws.Cells(ENGE_R0, 1), ws.Cells(ENGE_R0 + n - 1, 9)).Value = buf

concluir:
    If prot Then ws.Protect Password:=SENHA_ESPEC, UserInterfaceOnly:=True
    Exit Sub
falha:
    Dim numErro As Long, descErro As String
    numErro = Err.Number: descErro = Err.Description
    On Error Resume Next
    If prot Then ws.Protect Password:=SENHA_ESPEC, UserInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise numErro, "mEspecificacoes.AtualizarEngEspec", descErro
End Sub
'@

[System.IO.File]::WriteAllLines($Saida, $linhas.ToArray(), $enc)

$texto = [System.IO.File]::ReadAllText($Saida, $enc)
foreach ($obrigatorio in @('FALLBACK LEGADO', 'EspecLegada', 'FontePadrao', 'GravarEspec', 'AtualizarEngEspec')) {
    if ($texto -notmatch [regex]::Escape($obrigatorio)) { throw "Geracao incompleta: $obrigatorio ausente." }
}

"mEspecificacoes: perfil Hematologia gerado com fallback Analitos!Q:R"
"linhas: $($linhas.Count)"
"saida: $Saida"
