# Variáveis
$hostname = $env:COMPUTERNAME
$FileServerPath = "\\arquivosdti.clickip.local\automacao_dados\papercut_history_users"
$LocalCsv = "C:\PaperCut\logs\csv\papercut-print-log-all-time.csv"
$Username = "CLICKIP\richard.silva"
$Password = "Ri21851619!"  # Em produção, evite plaintext
$CurrentUser = $env:USERNAME

########################### service papercut ##################################

#$msipath="C:\Users\richard.silva\Downloads\papercut-print-logger.msi"

#Start-Process msiexec.exe -ArgumentList "/i `"$msiPath`" RUNNOW=1 /quiet /norestart" -Wait

#Start-Process $msipath -ArgumentList "/SILENT /b" -Wait

$destino="C:\PaperCut"
$svcName="PCPrintLogger"

# Registrar o serviço

if (-not (Get-Service -Name $svcName -ErrorAction SilentlyContinue)) {
    sc.exe create $svcName binPath= "`"$destino\pcpl.exe`" $svcName" start= auto
}

# Iniciar o serviço

Start-Service -Name $svcName

# Confirmar status

$status = Get-Service -Name $svcName

Write-Host "Serviço '$svcName' está: $($status.Status)"


Write-Host "Credenciais AD: $Username\$Password"

###############################################################################

# Converte a senha para SecureString
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)

# Pega nome do usuário atual e data formatada
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[-1]
$CurrentDate = Get-Date -Format "yyyyMMdd-HHmmss"
$RemoteFileName = "papercut_log_$hostname-$CurrentUser.csv"

# Mapeia a unidade de rede (K:)
try {
    New-PSDrive -Name "K" -PSProvider FileSystem -Root $FileServerPath -Credential $Credential -Persist -ErrorAction Stop
    Start-Sleep 5

    $RemoteCsv = "K:\$RemoteFileName"

    if (Test-Path $LocalCsv) {
        Copy-Item -Path $LocalCsv -Destination $RemoteCsv -Force
        Write-Host "Arquivo enviado com sucesso para $RemoteCsv" -ForegroundColor Green
    } else {
        Write-Host "Arquivo local '$LocalCsv' não encontrado!" -ForegroundColor Red
    }
}
catch {
    Write-Host "Erro ao mapear unidade de rede: $_" -ForegroundColor Red
}
finally {
    # Tenta remover o mapeamento mesmo em caso de erro
    if (Get-PSDrive -Name "K" -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name "K"
    }
}

##stop-process -id $PID
