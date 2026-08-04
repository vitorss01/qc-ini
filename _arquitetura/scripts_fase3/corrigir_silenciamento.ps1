# corrigir_silenciamento.ps1 - limite atingido passa a falhar ALTO
#
# Achado da inspecao de 04/08/2026: em tres pontos o sistema para de trabalhar
# ao bater num teto e NAO conta a ninguem. Num sistema de CQI isso e pior do
# que quebrar, porque quebrar avisa.
#
#   1. RegistrarEventosWestgard  "If regs <> "" And nEv < 5000"
#      A violacao 5.001 simplesmente nao entra no historico. O analista ve uma
#      lista que parece completa. Auditor nenhum tem como perceber.
#
#   2. MarcarNaoConforme  "If linReg > 0 Then"
#      Com a aba Registros cheia (200 ocorrencias), o resultado SAI dos
#      calculos e os dois logs registram, mas a ocorrencia nao aparece na
#      vitrine. Sucesso parcial silencioso -- o pior tipo.
#
# Depois deste patch os dois registram no Audit_Log e devolvem erro ao chamador.
# Eventos_Westgard e DERIVADA: pode ser reconstruida, entao bloquear e seguro.
# A aba Registros exige acao do gestor (arquivar ocorrencias antigas).
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).
#
# Uso: chamado por gerar_mEstatistica.ps1 e gerar_mRegistros; ver build_all.

param(
    [Parameter(Mandatory = $true)][string]$Arquivo,
    [Parameter(Mandatory = $true)][ValidateSet('mEstatistica', 'mRegistros')][string]$Alvo
)

$ErrorActionPreference = 'Stop'
$enc = [System.Text.Encoding]::Default
$L = [System.IO.File]::ReadAllLines($Arquivo, $enc)
$texto = ($L -join "`r`n")
$mudou = 0

# Idempotencia: o build roda varias vezes e as ancoras desaparecem apos a
# primeira aplicacao. Sem esta guarda, a segunda execucao lancaria "ponto nao
# encontrado" e derrubaria o build inteiro.
$marca = if ($Alvo -eq 'mEstatistica') { 'EVENTOS_WESTGARD_ESTOUROU' } else { 'REGISTROS_CHEIA' }
if ($texto -like "*$marca*") {
    "$Alvo : ja corrigido, nada a fazer"
    exit 0
}

if ($Alvo -eq 'mEstatistica') {
    $antes = 'If regs <> "" And nEv < 5000 Then'
    $depois = @'
                    ' Teto de eventos: descartar em silencio esconderia violacao
                    ' de Westgard do analista e do auditor. Conta e denuncia.
                    If regs <> "" And nEv >= UBound(ev, 1) Then
                        nDescartados = nDescartados + 1
                    End If
                    If regs <> "" And nEv < UBound(ev, 1) Then
'@
    if ($texto -notlike "*$antes*") { throw "Ponto de teto nao encontrado em $Arquivo" }
    $texto = $texto.Replace($antes, $depois.TrimEnd())
    $mudou++

    # declaracao do contador, logo apos o ReDim do buffer
    $antesDecl = 'ReDim ev(1 To 5000, 1 To 8)'
    $depoisDecl = @'
ReDim ev(1 To 5000, 1 To 8)
    Dim nDescartados As Long
    nDescartados = 0
'@
    if ($texto -notlike "*$antesDecl*") { throw "ReDim do buffer de eventos nao encontrado" }
    $texto = $texto.Replace($antesDecl, $depoisDecl.TrimEnd())
    $mudou++

    # denuncia depois de gravar os eventos
    $antesFim = '    ws.Range("J2").Value = nEv'
    $depoisFim = @'
    ws.Range("J2").Value = nEv

    ' Falha ALTA quando o buffer estourou. Eventos_Westgard e DERIVADA -- pode
    ' ser reconstruida --, entao interromper aqui e seguro e forca a decisao.
    If nDescartados > 0 Then
        Auditar CAT_SIS, "EVENTOS_WESTGARD_ESTOUROU", "mEstatistica", _
                0, "", "", "", 0, "", nEv, nDescartados, "", "", _
                "Buffer de " & UBound(ev, 1) & " eventos cheio; " & nDescartados & _
                " violacao(oes) NAO registrada(s)", _
                "Arquivar Eventos_Westgard e reexecutar para restaurar o historico completo"
        Err.Raise vbObjectError + 513, "RegistrarEventosWestgard", _
                  "Historico de Westgard incompleto: " & nDescartados & _
                  " violacao(oes) descartada(s) por buffer cheio (" & UBound(ev, 1) & _
                  "). O evento foi registrado no Audit_Log."
    End If
'@
    if ($texto -notlike "*$antesFim*") { throw 'Linha de total de eventos (J2) nao encontrada' }
    $texto = $texto.Replace($antesFim, $depoisFim.TrimEnd())
    $mudou++
}
else {
    # Ancora montada com -join "`r`n", nao com here-string.
    # O here-string herda a quebra de linha DO ARQUIVO DE SCRIPT (LF, escrito por
    # editor moderno), enquanto $texto vem de -join "`r`n". Anchor com LF nunca
    # casa com texto em CRLF, e o sintoma e "ponto nao encontrado" -- que parece
    # codigo-fonte diferente do esperado, e nao diferenca de quebra de linha.
    $antes = @(
        '    linReg = PrimeiraLinhaLivreReg()',
        '    If linReg > 0 Then'
    ) -join "`r`n"
    $depois = @'
    linReg = PrimeiraLinhaLivreReg()

    ' Vitrine cheia: o resultado JA saiu dos calculos e os dois logs ja
    ' gravaram. Seguir em silencio deixaria a ocorrencia invisivel para quem
    ' confere pela aba -- sucesso parcial, que e o pior tipo de falha.
    ' Diferente de Eventos_Westgard, esta aba NAO e derivada: exige o gestor
    ' arquivar ocorrencias antigas.
    If linReg = 0 Then
        Auditar CAT_SIS, "REGISTROS_CHEIA", "mRegistros", _
                run, dtCorrida, "", lote, nivel, analito, _
                valor, valor, stAntes, tipoNC, _
                "Aba Registros cheia (" & (REG_RN - REG_R0 + 1) & " linhas); " & _
                "ocorrencia nao exibida", _
                "Arquivar ocorrencias antigas da aba Registros"
        Err.Raise vbObjectError + 514, "MarcarNaoConforme", _
                  "Aba Registros cheia. O resultado SAIU dos calculos e a " & _
                  "auditoria registrou, mas a ocorrencia nao pode ser exibida. " & _
                  "Arquive ocorrencias antigas."
    End If

    If linReg > 0 Then
'@
    if ($texto -notlike "*$antes*") { throw "Ponto da vitrine cheia nao encontrado em $Arquivo" }
    $texto = $texto.Replace($antes, $depois.TrimEnd())
    $mudou++
}

# Normaliza para CRLF antes de gravar. Os blocos inseridos vem de here-string
# com LF; gravar misturado deixaria o modulo com quebras inconsistentes, que o
# VBE aceita mas o diff de codigo passa a acusar linha a linha sem motivo.
$texto = ($texto -replace "`r`n", "`n") -replace "`n", "`r`n"

[System.IO.File]::WriteAllText($Arquivo, $texto, $enc)
"$Alvo : $mudou substituicao(oes) aplicada(s)"
