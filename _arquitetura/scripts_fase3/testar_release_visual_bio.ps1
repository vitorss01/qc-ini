# testar_release_visual_bio.ps1 - prova dos campos de identidade da Bioquimica
#
# Confirma em clone que equipamento/serie/controle iniciam vazios, continuam
# desbloqueados sob protecao e alimentam a exibicao da aba Inicio. Nada e salvo.
#
# Nao usar acentos neste arquivo (Windows PowerShell 5.1 le .ps1 como ANSI).

param(
    [Parameter(Mandatory = $true)][string]$Workbook
)

$ErrorActionPreference = 'Stop'
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if ($Workbook -notmatch '(?i)(_EM_DESENVOLVIMENTO|build_hardening)') {
    throw 'Recusado: teste somente em clone _EM_DESENVOLVIMENTO ou build_hardening.'
}

$SENHA = 'qcini2025'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3
$wb = $xl.Workbooks.Open($Workbook)

try {
    if ($wb.ReadOnly) { throw "Somente leitura: $Workbook" }
    # nomes montados por codigo: o PowerShell 5.1 le .ps1 como ANSI e
    # corromperia os literais acentuados (DISP_E_BADINDEX na aba).
    $nomeCfg = 'Configura' + [char]231 + [char]227 + 'o'
    $nomeIni = 'In' + [char]237 + 'cio'
    $cfg = $wb.Worksheets.Item($nomeCfg)
    $ini = $wb.Worksheets.Item($nomeIni)
    $entradas = @('C9','C10','C11')

    foreach ($a in $entradas) {
        if ([string]$cfg.Range($a).Value2 -ne '') { throw "$($cfg.Name)!$a deveria iniciar vazio." }
        if ($cfg.Range($a).Locked) { throw "$($cfg.Name)!$a esta bloqueada." }
    }
    if (-not $cfg.ProtectContents) { throw 'A aba Configuracao precisa estar protegida no artefato final.' }
    if ([string]$ini.Range('C7').Value2 -ne '' -or [string]$ini.Range('C8').Value2 -ne '') {
        throw 'Inicio ainda exibe identidade copiada ou marcador artificial.'
    }

    # Escreve com a planilha PROTEGIDA. Se a celula estiver bloqueada, o COM
    # lanca erro. Os valores sao apagados e a pasta e fechada sem salvar.
    $cfg.Range('C9').Value2 = '__TESTE_EQUIPAMENTO__'
    $cfg.Range('C10').Value2 = '__TESTE_SERIE__'
    $cfg.Range('C11').Value2 = '__TESTE_CONTROLE__'
    $xl.Calculate()
    if ([string]$ini.Range('C7').Value2 -ne '__TESTE_EQUIPAMENTO__') { throw 'Inicio!C7 nao acompanha Configuracao!C9.' }
    if ([string]$ini.Range('C8').Value2 -ne '__TESTE_CONTROLE__') { throw 'Inicio!C8 nao acompanha Configuracao!C11.' }

    $cfg.Range('C9:C11').ClearContents()
    $xl.Calculate()
    'Bio identidade: 3 campos vazios, desbloqueados e editaveis sob protecao'
    'Bio Inicio: equipamento e controle vinculados ao cadastro, sem marcador PENDENTE'
}
finally {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
