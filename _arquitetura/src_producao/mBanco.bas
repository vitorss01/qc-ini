Attribute VB_Name = "mBanco"
Option Explicit
' ===== CAMADA DE CAPACIDADE E FLAGS DO BANCO (ADR-025) =====
'
' POR QUE ESTE MODULO EXISTE
'
' Ate aqui as colunas BA, BB e BC do DB_Resultados eram FORMULA, provisionadas
' ate a linha 15.003. Duas consequencias, ambas ruins:
'
'   1. CAPACIDADE. 1.110 registros/mes contra 15.000 linhas = 13,5 meses. Passado
'      isso o VBA continuava gravando (UltimaLinhaBanco usa End(xlUp), que nao
'      tem teto), mas as linhas novas nasciam SEM as formulas BA:BC e FORA dos
'      intervalos nomeados r*. O dado entrava, ficava salvo, e sumia de todo
'      calculo -- Painel, Calc, graficos, Estatistica. Sem erro, sem aviso.
'      Num sistema de CQ essa e a pior falha possivel: nao o erro visivel, e o
'      numero plausivel e errado, que ninguem confere porque parece certo.
'
'   2. CUSTO. BB e BC usavam COUNTIFS de faixa EXPANSIVA ($E$4:$E4): na linha 4
'      varria 1 celula, na linha 66.603 varria 66.600. Somado, ~2,5n^2. Medido:
'      recalculo completo de 5,7 s com 1 ano de dados e 221 s com 5 anos.
'
' O QUE MUDA
'
' BA, BB e BC deixam de ser formula e passam a ser VALOR gravado por
' AtualizarFlagsBanco, que varre o banco UMA vez com dois dicionarios -- O(n) no
' lugar de O(n^2) -- e escreve tudo num unico bloco.
'
' E os intervalos nomeados r* deixam de ter altura fixa: passam a ser
' redimensionados para a ultima linha REAL a cada gravacao. Isso importa mais do
' que parece. A aba Calc avalia AGGREGATE sobre esses intervalos 180 vezes, e o
' custo e O(180 x altura) SEM curto-circuito em celula vazia. Provisionar 120.000
' linhas de intervalo com 1.110 preenchidas custaria cem vezes mais caro do que
' precisa. Com o nome acompanhando o dado, o custo acompanha o dado.
'
' EQUIVALENCIA -- a regra que nao pode ser quebrada
'
' O valor gravado tem de ser identico ao que a formula produzia:
'
'   BA  =SE($D4="";"";EXT.TEXTO($D4;4;NUM.CARACT($D4)-5))
'       nucleo do lote; a mesma conta de NucleoLote(), que vive em mEntrada.
'
'   BB  =SE(OU($E4="";$G4<>"Ativo");"";
'           SE(CONT.SES($E$4:$E4;$E4;$A$4:$A4;$A4;$G$4:$G4;"Ativo")=1;1;0))
'       1 se esta e a PRIMEIRA linha Ativa do par (analito, RUN); senao 0;
'       vazio quando a propria linha nao esta Ativa.       Nome: rFirst
'
'   BC  =SE(OU($A4="";$G4<>"Ativo");"";
'           SE(CONT.SES($A$4:$A4;$A4;$G$4:$G4;"Ativo")=1;1;0))
'       1 se esta e a PRIMEIRA linha Ativa do RUN; senao 0. Nome: rRunUnico
'
' ATENCAO ao que "primeira" significa: primeira ENTRE AS ATIVAS, na ordem fisica
' das linhas. Por isso a flag NAO e estavel no momento da insercao -- excluir
' logicamente uma linha PROMOVE a proxima duplicata a "primeira". Toda rotina que
' mexe em Status tem de chamar AtualizarFlagsBanco depois. Calcular a flag apenas
' na insercao seria mais rapido e estaria errado.
'
' RASTREABILIDADE
'
' A conta continua explicavel e reproduzivel: e a mesma regra da formula, escrita
' em VBA, com a formula de origem transcrita acima. ConferirFlagsBanco() recalcula
' de forma independente e devolve o numero de divergencias -- serve de prova a
' qualquer momento, inclusive em auditoria.

' Capacidade maxima do banco, em LINHAS DE REGISTRO.
'
' Dimensionamento (medido em janeiro/2026, dado real):
'   observado hoje ....... 1.110 registros/mes (18 corridas x 31 analitos x 2 niveis)
'   pior caso plausivel .. 40 analitos (o cadastro vai ate a linha 43 = 40 vagas)
'                          x 2 niveis x 23 dias uteis = 1.840 registros/mes
'   60 meses no pior caso  1.840 x 60 = 110.400
'   adotado .............. 120.000  (margem de 8,7% sobre o pior caso;
'                                    108 meses = 9 anos no ritmo observado)
'
' Este numero nao custa bytes: nao ha mais formula provisionada no banco. Ele
' existe para que a capacidade seja EXPLICITA e para que o excesso seja
' RECUSADO com mensagem, em vez de aceito e ignorado.
Public Const CAP_LINHAS As Long = 120000

' A mesma senha de mSeguranca.ReprotectAll. Literal aqui, como ja e literal la:
' centraliza-la agora criaria dependencia nova entre modulos por um ganho
' estetico, e o ADR-025 nao e lugar para isso.
Private Const SENHA_ABAS As String = "qcini2025"

' Colunas derivadas do banco. Ficam aqui, e nao em mDados, porque agora sao
' MANTIDAS aqui: quem escreve a coluna e quem define onde ela fica.
Public Const COL_LOTE_NUC As Long = 53        ' BA - nucleo do lote      (rLote)
Public Const COL_FIRST As Long = 54           ' BB - 1a ativa do par     (rFirst)
Public Const COL_RUNUNICO As Long = 55        ' BC - 1a ativa do RUN     (rRunUnico)

Public Function UltimaLinhaCapacidade() As Long
    UltimaLinhaCapacidade = BANCO_R0 + CAP_LINHAS - 1      ' 120.003
End Function

Public Function LinhasLivres() As Long
    LinhasLivres = UltimaLinhaCapacidade() - UltimaLinhaBanco()
    If LinhasLivres < 0 Then LinhasLivres = 0
End Function

' Barreira de capacidade. Chamada ANTES de gravar.
'
' Regra do projeto: e preferivel bloquear com mensagem clara a aceitar o dado e
' deixa-lo invisivel para os calculos.
Public Sub ExigirCapacidade(ByVal quantasNovas As Long)
    Dim livres As Long
    livres = LinhasLivres()
    If quantasNovas > livres Then
        Err.Raise vbObjectError + 513, "mBanco.ExigirCapacidade", _
            "Capacidade do banco esgotada." & vbCrLf & vbCrLf & _
            "Tentativa de gravar " & quantasNovas & " registro(s), mas restam " & _
            livres & " linha(s) livres de " & CAP_LINHAS & "." & vbCrLf & vbCrLf & _
            "Nenhum dado foi gravado. Arquive o historico antigo ou aumente " & _
            "CAP_LINHAS em mBanco antes de continuar."
    End If
End Sub

' Invólucro de teste para a barreira.
'
' ExigirCapacidade levanta erro DE PROPOSITO -- e assim que ela protege o banco.
' Mas um Err.Raise sem tratamento chamado de fora por Application.Run faz o Excel
' abrir um dialogo modal que ninguem pode fechar em automacao: o processo trava
' com CPU zerada, e o sintoma (parece lento) engana quem diagnostica. Foi
' exatamente o que aconteceu na primeira rodada do teste de 60 meses.
'
' Aqui o erro e capturado e devolvido como TEXTO, para que a barreira possa ser
' exercitada por script sem travar o Excel.
Public Function TestarCapacidade(ByVal quantasNovas As Long) As String
    On Error GoTo capturou
    ExigirCapacidade quantasNovas
    TestarCapacidade = "PERMITIU"
    Exit Function
capturou:
    TestarCapacidade = "RECUSOU|" & Err.Description
End Function

' Reescreve BA/BB/BC como VALOR e redimensiona os intervalos nomeados.
'
' Uma varredura, dois dicionarios, uma escrita em bloco. Chamada no fim de
' UpsertResultados e de ExcluirLogico -- toda alteracao de Status muda as flags
' das linhas seguintes, ver a nota sobre "primeira entre as ativas" no topo.
Public Sub AtualizarFlagsBanco()
    Dim ws As Worksheet, ult As Long, dados As Variant
    Dim flags() As Variant, i As Long, n As Long
    Dim vistoAR As Object, vistoR As Object
    Dim analito As String, status As String, lote As String
    Dim chaveAR As String, chaveR As String, run As String
    Dim prot As Boolean

    Set ws = ThisWorkbook.Sheets(BANCO)
    ult = UltimaLinhaBanco()

    RedimensionarNomes ult
    If ult < BANCO_R0 Then Exit Sub

    dados = ws.Range(ws.Cells(BANCO_R0, COL_RUN), ws.Cells(ult, COL_STATUS)).Value
    n = UBound(dados, 1)
    ReDim flags(1 To n, 1 To 3)              ' 1=BA(lote) 2=BB(rFirst) 3=BC(rRunUnico)

    Set vistoAR = CreateObject("Scripting.Dictionary")
    Set vistoR = CreateObject("Scripting.Dictionary")
    vistoAR.CompareMode = 1
    vistoR.CompareMode = 1

    For i = 1 To n
        run = Trim$(CStr(dados(i, COL_RUN)))
        analito = Trim$(CStr(dados(i, COL_ANALITO)))
        status = Trim$(CStr(dados(i, COL_STATUS)))
        lote = Trim$(CStr(dados(i, COL_LOTE)))

        ' --- BA: nucleo do lote. Vazio quando nao ha codigo, como a formula.
        If Len(lote) > 5 Then
            flags(i, 1) = NucleoLote(lote)
        Else
            flags(i, 1) = Empty
        End If

        ' --- BB: primeira linha Ativa do par (analito, RUN)
        If Len(analito) = 0 Or status <> ST_ATIVO Then
            flags(i, 2) = Empty
        Else
            chaveAR = UCase$(analito) & Chr$(1) & run
            If vistoAR.Exists(chaveAR) Then
                flags(i, 2) = 0
            Else
                vistoAR.Add chaveAR, 1
                flags(i, 2) = 1
            End If
        End If

        ' --- BC: primeira linha Ativa do RUN
        If Len(run) = 0 Or status <> ST_ATIVO Then
            flags(i, 3) = Empty
        Else
            chaveR = run
            If vistoR.Exists(chaveR) Then
                flags(i, 3) = 0
            Else
                vistoR.Add chaveR, 1
                flags(i, 3) = 1
            End If
        End If
    Next i

    ' ---- escrita, com a protecao tratada -------------------------------
    '
    ' POR QUE ISTO EXISTE (e por que so apareceu depois)
    '
    ' As abas do produto sao protegidas por ReprotectAll com UserInterfaceOnly,
    ' que libera a escrita por VBA. Só que o Excel NAO PERSISTE o
    ' UserInterfaceOnly ao salvar: no artefato salvo e reaberto a aba volta a
    ' ficar protegida por inteiro ate o login rodar ReprotectAll de novo.
    '
    ' Nesse estado as duas instrucoes abaixo falham com 1004 -- a primeira antes
    ' da segunda, porque formatar celula ainda exige AllowFormattingCells, que e
    ' False. Foi o erro relatado: na producao a aba estava destravada e passava;
    ' no artefato, nao. Diagnostico completo em ANALISE_ESCALABILIDADE.md.
    '
    ' O padrao aqui e o mesmo ja usado em mImportar.MostrarErros e
    ' mImportar.LimparAreaImport: guardar o estado, destravar, escrever,
    ' RESTAURAR o que havia. Restaurar, e nao impor -- se a aba chegou
    ' destravada, ela sai destravada.
    prot = ws.ProtectContents
    On Error GoTo restaura
    If prot Then ws.Unprotect Password:=SENHA_ABAS

    ' Coluna BA como TEXTO: o nucleo pode ter zero a esquerda, e virar numero
    ' faria "0897" <> "897" nas comparacoes de lote.
    ws.Range(ws.Cells(BANCO_R0, COL_LOTE_NUC), ws.Cells(ult, COL_LOTE_NUC)).NumberFormat = "@"
    ws.Range(ws.Cells(BANCO_R0, COL_LOTE_NUC), ws.Cells(ult, COL_RUNUNICO)).Value = flags

restaura:
    ' A protecao volta SEMPRE, inclusive se a escrita levantou erro. Deixar a
    ' aba destravada por causa de uma excecao seria trocar um defeito visivel
    ' por um buraco de seguranca silencioso.
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    On Error Resume Next
    If prot Then
        ws.Protect Password:=SENHA_ABAS, UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
    End If
    On Error GoTo 0
    If nErr <> 0 Then Err.Raise nErr, "mBanco.AtualizarFlagsBanco", sErr
End Sub

' Os intervalos nomeados passam a ter a altura do DADO, nao a de um
' provisionamento fixo. Ver a justificativa de custo no topo do modulo.
'
' RefersToR1C1, e nao RefersTo com texto A1: passar "=DB_Resultados!$E$4:$E$999"
' por .RefersTo fez o Excel gravar "DB_Resultados!L4C5:L999C5" e os nomes
' pararam de resolver -- a aba Calc devolveu vazio e o grafico foi a zero ponto,
' em silencio. R1C1 nao tem essa ambiguidade.
Public Sub RedimensionarNomes(ByVal ult As Long)
    Dim alvo As Long
    alvo = ult
    If alvo < BANCO_R0 Then alvo = BANCO_R0      ' nunca um intervalo vazio: viraria #REF!

    DefinirNome "rRUN", COL_RUN, alvo
    DefinirNome "rData", COL_DATA, alvo
    DefinirNome "rNivel", COL_NIVEL, alvo
    DefinirNome "rAnalito", COL_ANALITO, alvo
    DefinirNome "rValor", COL_RESULT, alvo
    DefinirNome "rStatus", COL_STATUS, alvo
    DefinirNome "rLote", COL_LOTE_NUC, alvo
    DefinirNome "rFirst", COL_FIRST, alvo
    DefinirNome "rRunUnico", COL_RUNUNICO, alvo
End Sub

Private Sub DefinirNome(ByVal nome As String, ByVal col As Long, ByVal ult As Long)
    ThisWorkbook.Names(nome).RefersToR1C1 = _
        "=" & BANCO & "!R" & BANCO_R0 & "C" & col & ":R" & ult & "C" & col
End Sub

' Prova de equivalencia, disponivel a qualquer momento.
'
' Recalcula BB e BC por um caminho INDEPENDENTE do usado em AtualizarFlagsBanco
' (contagem direta, sem dicionario) e compara com o que esta gravado. Devolve
' "linhas|divergencias|primeira divergencia".
'
' Deliberadamente ingenua e O(n^2): e uma conferencia sob demanda, nao um
' caminho de producao. Se ela concordar com a versao rapida, as duas concordam
' com a formula que substituiram.
Public Function ConferirFlagsBanco(Optional ByVal maxLinhas As Long = 5000) As String
    Dim ws As Worksheet, ult As Long, dados As Variant, gravado As Variant
    Dim i As Long, j As Long, n As Long, div As Long, prim As String
    Dim cAR As Long, cR As Long, espBB As Variant, espBC As Variant

    Set ws = ThisWorkbook.Sheets(BANCO)
    ult = UltimaLinhaBanco()
    If ult < BANCO_R0 Then ConferirFlagsBanco = "0|0|": Exit Function
    If ult > BANCO_R0 + maxLinhas - 1 Then ult = BANCO_R0 + maxLinhas - 1

    dados = ws.Range(ws.Cells(BANCO_R0, COL_RUN), ws.Cells(ult, COL_STATUS)).Value
    gravado = ws.Range(ws.Cells(BANCO_R0, COL_LOTE_NUC), ws.Cells(ult, COL_RUNUNICO)).Value
    n = UBound(dados, 1)

    For i = 1 To n
        If Len(Trim$(CStr(dados(i, COL_ANALITO)))) = 0 Or _
           Trim$(CStr(dados(i, COL_STATUS))) <> ST_ATIVO Then
            espBB = Empty
        Else
            cAR = 0
            For j = 1 To i
                If Trim$(CStr(dados(j, COL_STATUS))) = ST_ATIVO Then
                    If UCase$(Trim$(CStr(dados(j, COL_ANALITO)))) = UCase$(Trim$(CStr(dados(i, COL_ANALITO)))) Then
                        If Trim$(CStr(dados(j, COL_RUN))) = Trim$(CStr(dados(i, COL_RUN))) Then cAR = cAR + 1
                    End If
                End If
            Next j
            If cAR = 1 Then espBB = 1 Else espBB = 0
        End If

        If Len(Trim$(CStr(dados(i, COL_RUN)))) = 0 Or _
           Trim$(CStr(dados(i, COL_STATUS))) <> ST_ATIVO Then
            espBC = Empty
        Else
            cR = 0
            For j = 1 To i
                If Trim$(CStr(dados(j, COL_STATUS))) = ST_ATIVO Then
                    If Trim$(CStr(dados(j, COL_RUN))) = Trim$(CStr(dados(i, COL_RUN))) Then cR = cR + 1
                End If
            Next j
            If cR = 1 Then espBC = 1 Else espBC = 0
        End If

        If CStr(gravado(i, 2)) <> CStr(espBB) Or CStr(gravado(i, 3)) <> CStr(espBC) Then
            div = div + 1
            If Len(prim) = 0 Then
                prim = "L" & (BANCO_R0 + i - 1) & " BB=" & CStr(gravado(i, 2)) & _
                       "/esp " & CStr(espBB) & " BC=" & CStr(gravado(i, 3)) & "/esp " & CStr(espBC)
            End If
        End If
    Next i
    ConferirFlagsBanco = CStr(n) & "|" & CStr(div) & "|" & prim
End Function
