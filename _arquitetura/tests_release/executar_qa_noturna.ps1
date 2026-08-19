param(
    [Parameter(Mandatory = $true)][string]$Bioquimica,
    [Parameter(Mandatory = $true)][string]$Hematologia,
    [Parameter(Mandatory = $true)][string]$OutCsv
)

$ErrorActionPreference = 'Stop'
trap {
    Write-Output ("QA_FALHA linha {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
    throw
}
$Bioquimica=(Resolve-Path -LiteralPath $Bioquimica).Path
$Hematologia=(Resolve-Path -LiteralPath $Hematologia).Path
$OutCsv=[IO.Path]::GetFullPath($OutCsv)
$scripts = Split-Path -Parent $PSScriptRoot
$scripts = Join-Path $scripts 'scripts_fase3'
$modQA = Join-Path $PSScriptRoot 'mQA_Noturno.bas'
$qaBase = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'QC_INI_QA\runs')).TrimEnd('\')
$qaRoot = [IO.Path]::GetFullPath((Join-Path $qaBase (Get-Date -Format 'yyyyMMdd_HHmmss')))
if (-not $qaRoot.StartsWith($qaBase + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "Raiz QA fora do diretorio temporario autorizado: $qaRoot"
}
$outFull = [IO.Path]::GetFullPath($OutCsv)
if ($outFull.StartsWith($qaRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutCsv nao pode ficar dentro do clone descartavel.'
}
New-Item -ItemType Directory -Force -Path $qaRoot | Out-Null
$resultados = New-Object System.Collections.ArrayList

function Add-Resultado([string]$produto,[string]$teste,[bool]$passou,[string]$detalhe) {
    [void]$resultados.Add([pscustomobject]@{Produto=$produto;Teste=$teste;Resultado=$(if($passou){'PASS'}else{'FAIL'});Detalhe=$detalhe})
}

function Hash([string]$s) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($s))).Replace('-','').ToLowerInvariant()) }
    finally { $sha.Dispose() }
}

function Nome-Range($wb,[string]$nome) { return $wb.Names.Item($nome).RefersToRange }

try {
foreach ($item in @(
    [pscustomobject]@{Produto='Bioquimica';Fonte=$Bioquimica},
    [pscustomobject]@{Produto='Hematologia';Fonte=$Hematologia}
)) {
    $produto=$item.Produto
    $hashFonteAntes=(Get-FileHash -Algorithm SHA256 -LiteralPath $item.Fonte).Hash
    $t07Run=0
    $t07Finalizou=$false
    Write-Output "[$produto] preparando clone"
    $dir=Join-Path $qaRoot $produto
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $clone=Join-Path $dir ([IO.Path]::GetFileName($item.Fonte))
    Copy-Item -LiteralPath $item.Fonte -Destination $clone -Force

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $scripts 'normalizar_protecao.ps1') -Workbook $clone | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "normalizar_protecao falhou em $produto" }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $scripts 'aplicar_vba.ps1') -Workbook $clone -Modulos $modQA | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "aplicar_vba QA falhou em $produto" }

    $xl=New-Object -ComObject Excel.Application
    # Eventos ficam desabilitados somente durante o Open para nao disparar
    # Workbook_Open na automacao invisivel; macros permanecem disponiveis para
    # as chamadas explicitas de Application.Run.
    $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
    $wb=$xl.Workbooks.Open($clone)
    $xl.EnableEvents=$true
    try {
        Write-Output "[$produto] T01 autenticacao"
        $usr=$wb.Worksheets.Item('Usuarios'); try{$usr.Unprotect('qcini2025')}catch{}
        $linhas=@(48,49,50)
        $contas=@(
            @('qa_analyst','QA Analyst','ANALISTA','qa123','QA'),
            @('qa_tech','QA Tecnico','TECNICO','qa123','QA'),
            @('qa_blank','QA Sem Rubrica','ANALISTA','qa123','')
        )
        for($i=0;$i -lt 3;$i++){
            $r=$linhas[$i]; $u=$contas[$i]
            $usr.Cells.Item($r,1).Value2=$u[0]; $usr.Cells.Item($r,2).Value2=$u[1]
            $usr.Cells.Item($r,3).Value2=$u[2]; $usr.Cells.Item($r,4).Value2=(Hash $u[3]); $usr.Cells.Item($r,5).Value2=$u[4]
        }

        # T01: credencial invalida e valida.
        (Nome-Range $wb 'currentUser').Value2=''; (Nome-Range $wb 'currentPapel').Value2=''
        (Nome-Range $wb 'loginUser').Value2='qa_analyst'; (Nome-Range $wb 'loginPass').Value2='errada'
        $xl.Run("'$($wb.Name)'!DoLogin") | Out-Null
        $msg=[string](Nome-Range $wb 'loginMsg').Value2
        Add-Resultado $produto 'T01_AUTH_NEGATIVA' ($msg -like '*invalid*') $msg
        (Nome-Range $wb 'loginUser').Value2='qa_analyst'; (Nome-Range $wb 'loginPass').Value2='qa123'
        $xl.Run("'$($wb.Name)'!DoLogin") | Out-Null
        $logado=[string](Nome-Range $wb 'currentUser').Value2
        Add-Resultado $produto 'T01_AUTH_POSITIVA' ($logado -eq 'qa_analyst') ("currentUser="+$logado)

        Write-Output "[$produto] T02 papeis"
        # T02: analista nao administra perfis; ADM consegue cadastrar TECNICO.
        try{$usr.Unprotect('qcini2025')}catch{}
        (Nome-Range $wb 'currentPapel').Value2='ANALISTA'
        (Nome-Range $wb 'cadLogin').Value2='qa_escalacao'; (Nome-Range $wb 'cadNome').Value2='QA Escalacao'
        (Nome-Range $wb 'cadSenha').Value2='qa123'; (Nome-Range $wb 'cadPapel').Value2='ADM'
        $xl.Run("'$($wb.Name)'!CadastrarUsuario") | Out-Null
        $cadMsg=[string](Nome-Range $wb 'cadMsg').Value2
        Add-Resultado $produto 'T02_ANALISTA_NAO_CRIA_ADM' ($cadMsg -like 'Apenas ADM*') $cadMsg
        (Nome-Range $wb 'currentPapel').Value2='ADM'
        (Nome-Range $wb 'cadLogin').Value2='qa_novo'; (Nome-Range $wb 'cadNome').Value2='QA Novo'
        (Nome-Range $wb 'cadSenha').Value2='qa123'; (Nome-Range $wb 'cadPapel').Value2='TECNICO'
        $xl.Run("'$($wb.Name)'!CadastrarUsuario") | Out-Null
        $cadMsg=[string](Nome-Range $wb 'cadMsg').Value2
        Add-Resultado $produto 'T02_ADM_CADASTRA' ($cadMsg -like "Usuario 'qa_novo' salvo*") $cadMsg

        Write-Output "[$produto] T03-T05 e T08"
        foreach($t in @('QA_T03_DuplicataAtomica','QA_T03_ReativacaoBloqueada','QA_T04_Capacidade','QA_T08_RunMesmoDia','QA_T05_Reconciliacao')){
            $r=[string]$xl.Run("'$($wb.Name)'!$t")
            Add-Resultado $produto $t ($r -like 'PASS*') $r
        }

        Write-Output "[$produto] T06 auditoria"
        # T06: cadeia de auditoria permanece integra depois das mutacoes QA.
        $r=[string]$xl.Run("'$($wb.Name)'!VerificarIntegridadeLog")
        Add-Resultado $produto 'T06_AUDITORIA_HASH' ($r -like 'OK|*') $r

        Write-Output "[$produto] T07 assinatura"
        # Na Bioquimica, o unico RUN elegivel para T07 e criado pelo modulo QA
        # no clone, com data sentinela e resultado exatamente na media-alvo.
        $temRunOK=$false
        if($produto -eq 'Bioquimica'){
            $prep=[string]$xl.Run("'$($wb.Name)'!QA_T07_PrepararRunPositivo")
            $m=[regex]::Match($prep,'^PASS\|RUN=(\d+);')
            $temRunOK=$m.Success
            if($temRunOK){$t07Run=[int]$m.Groups[1].Value}
            Add-Resultado $produto 'T07_BIO_RUN_QA_ISOLADO' $temRunOK $prep
        } else {
            $bi=$wb.Worksheets.Item('BI_Data'); try{$bi.Unprotect('qcini2025')}catch{}
            $ult=$bi.Cells.Item($bi.Rows.Count,1).End(-4162).Row
            $grupos=@{}
            $mat=$bi.Range("A2:AK$ult").Value2
            for($rr=$mat.GetLowerBound(0);$rr -le $mat.GetUpperBound(0);$rr++){
                if([int]$mat.GetValue($rr,18) -ne 1){continue}
                $run=[int]$mat.GetValue($rr,15); $ver=[string]$mat.GetValue($rr,37)
                if(-not $grupos.ContainsKey($run)){$grupos[$run]=$true}
                if($ver -ne 'OK'){$grupos[$run]=$false}
            }
            $entradaOK=$grupos.GetEnumerator()|Where-Object{$_.Value}|Select-Object -First 1
            $temRunOK=($null -ne $entradaOK)
            if($temRunOK){$t07Run=[int]$entradaOK.Key}
        }
        $lib=$null; foreach($ws in $wb.Worksheets){if($ws.CodeName -eq 'Planilha12'){$lib=$ws;break}}
        try{$lib.Unprotect('qcini2025')}catch{}
        if($temRunOK -and $t07Run -gt 0){
            $lib.Cells.Item(4,1).Value2=[string]$t07Run
            $lib.Cells.Item(4,2).Value2='2098-07-07'
            $lib.Range('C4:F4').ClearContents() | Out-Null
            (Nome-Range $wb 'currentUser').Value2='qa_analyst'; (Nome-Range $wb 'currentPapel').Value2='ANALISTA'
            $r=[string]$xl.Run("'$($wb.Name)'!AssinarCom",$lib.Range('C4'),'qa_blank','qa123','INICIADO')
            Add-Resultado $produto 'T07_RUBRICA_FALHA_FECHADA' ($r -like 'Rubrica*') $r
            $r=[string]$xl.Run("'$($wb.Name)'!AssinarCom",$lib.Range('C4'),'qa_tech','qa123','INICIADO')
            Add-Resultado $produto 'T07_INICIO_VALIDO' ($r -eq 'OK') $r
            $r=[string]$xl.Run("'$($wb.Name)'!AssinarCom",$lib.Range('E4'),'qa_tech','qa123','FINALIZADO')
            Add-Resultado $produto 'T07_TECNICO_NAO_FINALIZA' ($r -like 'Apenas ANALISTA*') $r
            $r=[string]$xl.Run("'$($wb.Name)'!AssinarCom",$lib.Range('E4'),'qa_analyst','qa123','FINALIZADO')
            $t07Finalizou=($r -eq 'OK')
            Add-Resultado $produto 'T07_FINALIZACAO_CQI_OK' $t07Finalizou $r
            if($produto -eq 'Bioquimica'){
                $recon=[string]$xl.Run("'$($wb.Name)'!QA_T05_Reconciliacao")
                Add-Resultado $produto 'T07_RECONCILIACAO_POS_FINALIZACAO' ($recon -like 'PASS*') $recon
            }
        } else {
            Add-Resultado $produto 'T07_RUBRICA_FALHA_FECHADA' $false 'RUN QA valido indisponivel; assinatura nao tentada'
            Add-Resultado $produto 'T07_INICIO_VALIDO' $false 'RUN QA valido indisponivel; assinatura nao tentada'
            Add-Resultado $produto 'T07_TECNICO_NAO_FINALIZA' $false 'RUN QA valido indisponivel; assinatura nao tentada'
            Add-Resultado $produto 'T07_FINALIZACAO_CQI_OK' $false 'RUN QA valido indisponivel; nenhum RUN real foi usado'
        }
        $r=[string]$xl.Run("'$($wb.Name)'!VerificarIntegridadeLog")
        Add-Resultado $produto 'T06_AUDITORIA_APOS_ASSINATURA' ($r -like 'OK|*') $r

        Write-Output "[$produto] T09 save/reopen"
        # T09: o save tecnico ocorre somente no clone. Eventos sao desativados
        # para nao abrir fluxos interativos do Workbook_BeforeSave na automacao.
        $antes=$wb.Worksheets.Item('DB_Resultados').Cells.Item($wb.Worksheets.Item('DB_Resultados').Rows.Count,1).End(-4162).Row
        $xl.EnableEvents=$false
        $salvou=$false
        for($tentativa=1;$tentativa -le 10 -and -not $salvou;$tentativa++){
            try{$wb.Save(); $salvou=$true}catch{if($tentativa -eq 10){throw}; Start-Sleep -Milliseconds 500}
        }
        $fechou=$false
        for($tentativa=1;$tentativa -le 10 -and -not $fechou;$tentativa++){
            try{$wb.Close($false); $fechou=$true}catch{if($tentativa -eq 10){throw}; Start-Sleep -Milliseconds 500}
        }
        $wb=$null
        $xl.AutomationSecurity=3
        $abriu=$false
        for($tentativa=1;$tentativa -le 10 -and -not $abriu;$tentativa++){
            try{$wb=$xl.Workbooks.Open($clone,0,$true); $abriu=$true}catch{if($tentativa -eq 10){throw}; Start-Sleep -Milliseconds 500}
        }
        $depois=$wb.Worksheets.Item('DB_Resultados').Cells.Item($wb.Worksheets.Item('DB_Resultados').Rows.Count,1).End(-4162).Row
        Add-Resultado $produto 'T09_PERSISTENCIA_REABERTURA' ($antes -eq $depois) ("ultimaLinha=$depois")
        if($produto -eq 'Bioquimica' -and $t07Run -gt 0){
            $db=$wb.Worksheets.Item('DB_Resultados'); $ultDB=$db.Cells.Item($db.Rows.Count,1).End(-4162).Row
            $nDB=0
            for($rr=4;$rr -le $ultDB;$rr++){
                if([int]$db.Cells.Item($rr,1).Value2 -eq $t07Run -and [string]$db.Cells.Item($rr,7).Value2 -eq 'Ativo'){$nDB++}
            }
            $bi=$wb.Worksheets.Item('BI_Data'); $ultBI=$bi.Cells.Item($bi.Rows.Count,1).End(-4162).Row
            $nBIOK=0
            for($rr=2;$rr -le $ultBI;$rr++){
                if([int]$bi.Cells.Item($rr,15).Value2 -eq $t07Run -and [int]$bi.Cells.Item($rr,18).Value2 -eq 1 -and [string]$bi.Cells.Item($rr,37).Value2 -eq 'OK'){$nBIOK++}
            }
            $lib2=$null; foreach($ws in $wb.Worksheets){if($ws.CodeName -eq 'Planilha12'){$lib2=$ws;break}}
            $metaIni=[string]$lib2.Cells.Item(4,4).Value2; $metaFim=[string]$lib2.Cells.Item(4,6).Value2
            $persistiu=($t07Finalizou -and $nDB -eq 1 -and $nBIOK -eq 1 -and $metaIni -like 'qa_tech*' -and $metaFim -like 'qa_analyst*')
            Add-Resultado $produto 'T07_PERSISTENCIA_FINALIZACAO' $persistiu ("RUN=$t07Run; DB=$nDB; BI_OK=$nBIOK; inicio=$metaIni; fim=$metaFim")
        }
    }
    finally {
        try{if($null -ne $wb){$wb.Close($false)}}catch{}
        try{$xl.Quit()}catch{}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
    }

    # Estado de distribuicao e contrato sao verificados no candidato imutavel.
    $xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=3
    $w0=$xl.Workbooks.Open($item.Fonte,0,$true)
    try {
        $vis=@($w0.Worksheets|Where-Object{$_.Visible -eq -1}|ForEach-Object{$_.Name})
        $prot=@($w0.Worksheets|Where-Object{-not $_.ProtectContents})
        $lo=$w0.Worksheets.Item('BI_Data').ListObjects.Item('tblBI_Fato')
        Add-Resultado $produto 'T09_ESTRUTURA_PROTEGIDA' ($w0.ProtectStructure -and $prot.Count -eq 0 -and $vis.Count -eq 1 -and $vis[0] -eq 'Login') ("visiveis="+($vis -join ',')+"; desprotegidas="+$prot.Count)
        Add-Resultado $produto 'T09_CONTRATO_60_COLUNAS' ($lo.ListColumns.Count -eq 60) ("linhas="+$lo.ListRows.Count+"; colunas="+$lo.ListColumns.Count)
    }
    finally { $w0.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null }

    $hashFonteDepois=(Get-FileHash -Algorithm SHA256 -LiteralPath $item.Fonte).Hash
    Add-Resultado $produto 'T07_FONTE_IMUTAVEL_SHA256' ($hashFonteAntes -eq $hashFonteDepois) ("antes=$hashFonteAntes; depois=$hashFonteDepois")

    Remove-Item -LiteralPath $dir -Recurse -Force
    Add-Resultado $produto 'T07_CLONE_DESCARTADO' (-not (Test-Path -LiteralPath $dir)) ("removido=$dir")
}

$resultados | Export-Csv -LiteralPath $OutCsv -NoTypeInformation -Encoding UTF8
$falhas=@($resultados|Where-Object{$_.Resultado -ne 'PASS'})
"QA root: $qaRoot"
"Testes: $($resultados.Count); falhas: $($falhas.Count)"
$resultados | Format-Table -AutoSize
if($falhas.Count -gt 0){exit 2}
}
finally {
    if(Test-Path -LiteralPath $qaRoot){
        Remove-Item -LiteralPath $qaRoot -Recurse -Force
    }
    Write-Output ("QA cleanup: raiz temporaria removida=" + (-not (Test-Path -LiteralPath $qaRoot)))
}
