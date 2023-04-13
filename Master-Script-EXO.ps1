Write-Host "

███████╗██╗░░██╗░█████╗░██╗░░██╗░█████╗░███╗░░██╗░██████╗░███████╗  ░█████╗░███╗░░██╗██╗░░░░░██╗███╗░░██╗███████╗
██╔════╝╚██╗██╔╝██╔══██╗██║░░██║██╔══██╗████╗░██║██╔════╝░██╔════╝  ██╔══██╗████╗░██║██║░░░░░██║████╗░██║██╔════╝
█████╗░░░╚███╔╝░██║░░╚═╝███████║███████║██╔██╗██║██║░░██╗░█████╗░░  ██║░░██║██╔██╗██║██║░░░░░██║██╔██╗██║█████╗░░
██╔══╝░░░██╔██╗░██║░░██╗██╔══██║██╔══██║██║╚████║██║░░╚██╗██╔══╝░░  ██║░░██║██║╚████║██║░░░░░██║██║╚████║██╔══╝░░
███████╗██╔╝╚██╗╚█████╔╝██║░░██║██║░░██║██║░╚███║╚██████╔╝███████╗  ╚█████╔╝██║░╚███║███████╗██║██║░╚███║███████╗
╚══════╝╚═╝░░╚═╝░╚════╝░╚═╝░░╚═╝╚═╝░░╚═╝╚═╝░░╚══╝░╚═════╝░╚══════╝  ░╚════╝░╚═╝░░╚══╝╚══════╝╚═╝╚═╝░░╚══╝╚══════╝

#######################################################################################################################
##                                                                                                                   ##
## SCRIPT FOR EXCHANGE ONLINE                                                                                        ##
## DEVELOPED BY: VICTOR MARTINS                                                                                      ##
## https://www.victornanuvem.com/                                                                                    ##
##                                                                                                                   ##
## VERSION 0.5                                                                                                       ##
##                                                                                                                   ##
## Description: This script performs tasks based on CSV files or with all users based on menu selection              ##
## Before using the script, read the definitions of each function, any incorrect action can harm your environment    ##
##                                                                                                                   ##
#######################################################################################################################"-ForegroundColor Yellow


##### SCRIPT OPTIONS #####
# Change it according to your location
# CSV Delimiter configuration
$delimiter=";"
# Path to save log files
$caminholog="C:\temp\logs"
##### END OF SCRIPT OPTIONS ####

#Funcao para logar no Exchange Online
function Login {

#Conectar e logar no ExchangeOnline (MFA)
#Lista o módulo do exchange para verificar se a sessão está ativa
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Host "Exchange Online module is not installed. Do you want to install it now? (Y/N)"

    $response = Read-Host
    if ($response.ToLower() -eq 'y') {
        Install-Module -Name ExchangeOnlineManagement
        Write-Host "The Exchange Online module has been successfully installed." -ForegroundColor Yellow
    }
    else {
        Write-Host "The installation of the Exchange Online module was cancelled.
The script will not proceed."

sair
    }
}
else {

$ExchangeModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable

if ($ExchangeModule.Version -ge '3.0') {
$getsessionsv3 = Get-ConnectionInformation | Select-Object State,Name
$isconnectedv3 = (@($getsessionsv3) -like '@{State=Connected; Name=ExchangeOnline*').Count -gt 0
}
else {
$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
}


If ($isconnected -eq "True" -or $isconnectedv3 -eq "True") {
#Connect-ExchangeOnline ##-ne
#Sessao aberta, valida
Write-Host "
There is an open session for Exchange Online" -ForegroundColor Green

#Usar mesma sessao?
$r = Read-Host "Type [Y] to use the same session or any other key to start a new session"
	if($r.ToLower() -eq 'y')
		{
			Write-Host "Using the same session..." -ForeGroundColor Green
			return;	
		}
#Senao, remover sessao ativa
Write-Host "Disconnecting from Exchange Online active session..." -ForegroundColor Yellow
if ($ExchangeModule.Version -ge '3.0') {
Disconnect-ExchangeOnline -Confirm:$false
}
else {
$getsessions | Remove-PSSession
}
	
	    }

 try
    { 
      Connect-ExchangeOnline -ShowBanner:$false
    }
#Se der erro
  catch
   {
	Write-host "An error occurred while connecting to Exchange Online. Try again." -ForegroundColor Red
    break
   }
  }
}

#Importar CSV
Function Get-FileName($initialDirectory)
{
[System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) |
Out-Null

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.CheckFileExists = $true
#$OpenFileDialog.filter = “All files (*.*)| *.*”
$OpenFileDialog.filter = “CSV Files (*.csv)| *.csv”
#$OpenFileDialog.ShowDialog() | Out-Null


 If ($OpenFileDialog.ShowDialog() -eq "Cancel") 
 {
  [System.Windows.Forms.MessageBox]::Show("No files were selected. Please select a file!", "Error", 0, 
  [System.Windows.Forms.MessageBoxIcon]::Exclamation)
  Break
  }   #$Global:SelectedFile = $OpenFileDialog.SafeFileName
$OpenFileDialog.filename

} #fim da funcao Get-FileName

#funcao para validar o csv se possui header correto
function Validate-CSVHeaders {
#header para validar
    $headerValido = "userprincipalname"
#importa o csv para validacao
#Importar CSV
   
    $content = Import-Csv -Path $file -Delimiter $delimiter
    $headers = ($content | Get-Member -MemberType NoteProperty).Name
#converte o header para minusculo
    $headers = $headers.ToLower()
 if($headerValido -in $headers){
     Write-Host "CSV válido, executando script" -ForegroundColor Green
     Return
 } else {
 #se o csv nao possuir o header, cancela e volta para o menu
 Write-Host "===================== WARNING =====================
Invalid CSV file, returning to main menu
===================================================" -ForegroundColor Red
 Start-Sleep 2
 Menu
    }
   }

#================== Abrir opcoes ==================#
function Menu {
"--------------------------------------------------"
Write-Host -ForegroundColor Yellow " MENU"
" 1- Configure mail forwarding in mailboxes
 2- Delete mail forwarding from mailboxes
 3- Check if mailbox(es) exist(s)
 4- Show or Hide mailbox(es) from GAL
 5- Enable archive and/or litigation hold
 6- Enable autoexpanding archive
 7- Disable 'Email address policies' from Exchange
 8- Check if user migrated to Exchange Online or is on Exchange Server (Only Hybrid)
 9- Add Secondary SMTP Address in Mailbox
 10- Convert Usermailbox to Sharedmailbox
 11- Convert Sharedmailbox to Usermailbox
 12- Add permissions on Sharedmailbox
 13- Create a Sharedmailbox
 14- Remove Inbox Rule from Mailbox
 15- Export mailbox permissions (UserMailbox and/or SharedMailbox)
 16- Check Usermailbox Size, Archive and Litigation Hold settings

  0- Exit
--------------------------------------------------
"

$opcao=Read-Host "Enter the desired option"

#carrega funcao de acordo com a opcao selecionada
if ($opcao -eq "1") { um }
elseif ($opcao -eq "2") { dois }
elseif ($opcao -eq "3") { tres }
elseif ($opcao -eq "4") { quatro }
elseif ($opcao -eq "5") { cinco }
elseif ($opcao -eq "6") { seis }
elseif ($opcao -eq "7") { sete }
elseif ($opcao -eq "8") { oito }
elseif ($opcao -eq "9") { nove }
elseif ($opcao -eq "10") { dez }
elseif ($opcao -eq "11") { onze }
elseif ($opcao -eq "12") { doze }
elseif ($opcao -eq "13") { treze }
elseif ($opcao -eq "14") { quatorze }
elseif ($opcao -eq "15") { quinze }
elseif ($opcao -eq "16") { dezesseis }
elseif ($opcao -eq "0") { sair }
#opcao invalida
else {
Write-Host "==================================================
Invalid option! Please select a valid option.
==================================================" -ForegroundColor Red
#aguarda 2 segundos e carrega o menu novamente
Start-Sleep -Seconds 2
Menu
}
}


#Criando Log
$date = Get-Date -format "ddMMyyyy-HHmmss"
$dateI = Get-Date -Format "dd/MM/yyyy - HH:mm:ss"
$Logs = "master-script-exo_"+"$date"+".log"
Start-Transcript "$caminholog\$logs"

Write-Host "Process started at" $dateI -ForegroundColor Magenta

#sair
function sair {
"Exiting the script..."
break
}

#1- Configurar redirect nas contas
function um {

<# CSV HEADERS REQUIRED 
UserPrincipalName
NewUPN
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#================== Abrir opcoes ==================#
"--------------------------------------------------"
write-host -ForegroundColor Yellow " How do you want to configure the email route?"
" 1- Keep copy of message at source mailbox and forward (DeliverToMailboxAndForward)
 2- Only forward the message without keeping the copy
--------------------------------------------------
"
$fwdop=Read-Host "Enter the desired option"


if ($fwdop -eq "1") {
$DeliverToMailboxAndForward=$true
} else {
$DeliverToMailboxAndForward=$false
}

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

$elapsedTime = [system.diagnostics.stopwatch]::StartNew()
foreach ($item in $usuarios) {
$origem=$item.userprincipalname
$destino=$item.newupn

Write-Progress -Activity "Setting up forwarding in accounts - $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

Write-Host "Creating forwarding from " -NoNewline; Write-Host $origem -ForegroundColor Yellow -NoNewline; Write-Host " to " -NoNewline; Write-Host $destino -ForegroundColor Green;
Set-Mailbox $origem -ForwardingsmtpAddress $destino -DeliverToMailboxAndForward $DeliverToMailboxAndForward

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Write-Progress -Activity "Setting up forwarding in accounts - $origem" -Completed

Menu
}


# 2- Apagar redirect das contas
function dois {

<# CSV HEADERS REQUIRED 
UserPrincipalName
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

$elapsedTime = [system.diagnostics.stopwatch]::StartNew()
foreach ($item in $usuarios) {
$origem=$item.userprincipalname

Write-Progress -Activity "Removing forwarding - $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete


Write-Host "Deleting redirect from " -NoNewline; Write-Host $origem -ForegroundColor Green;
Set-Mailbox $origem -ForwardingSmtpAddress $Null

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Write-Progress -Activity "Removing forwarding - $origem" -Completed

Menu
}

# 3- Checar todas as mailboxes
function tres {

<# CSV HEADERS REQUIRED 
UserPrincipalName
#>

#$correctHeaders = @('userprincipalname')
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter


#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

#cria variaveis para armazenar as caixas
$MbxNao=@()
$MbxSim=@()

#inicio do contador
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {
$origem=$item.userprincipalname


Write-Progress -Activity "Checking $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

if (Get-ExoMailbox $origem -ErrorAction SilentlyContinue) {
    Write-Host $origem
    $MbxSim+=$origem
    
} else {
    Write-Host -ForegroundColor Red $origem
    $MbxNao+=$origem
}

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)


}
$elapsedTime.stop()
Write-Progress -Activity "Checking $origem" -Completed

#================== Abrir opcoes ==================#
"
--------------------------------------------------------------
"
Write-Host "Mailboxes found: " -NoNewline; Write-Host $MbxSim.count -ForegroundColor Green
Write-Host "Mailboxes not found: " -NoNewline; Write-Host $Mbxnao.count -ForegroundColor Red
"
--------------------------------------------------------------"
write-host -ForegroundColor Yellow " Do you want to take some action?"
" 1- Export found and not found mailboxes (.txt) - $caminholog
 2- Copy found mailboxes to clipboard
 3- Copy NOT found mailboxes to clipboard
 4- Go back to main menu
 5- Exit
--------------------------------------------------------------
"
$exportar_op3=Read-Host "Enter the desired option"
if($exportar_op3 -eq "1") {

$MbxSim | Out-File $caminholog\MbxFound-$date.txt
$MbxNao | Out-File $caminholog\MbxNOTFound-$date.txt

} elseif($exportar_op3 -eq "2") {
$MbxSim | clip
} elseif($exportar_op3 -eq "3") {
$MbxNao | clip
} elseif($exportar_op3 -eq "4") {
Menu
} else {
sair
}
}

# 4- Remover hide da GAL
function quatro {

<# CSV HEADERS REQUIRED 
UserPrincipalName
#>


#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

#inicio do contador
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()


#================== Abrir opcoes ==================#
"--------------------------------------------------"
write-host -ForegroundColor Yellow " Do you want to hide or show"
" 1- Hide from GAL
 2- Show in GAL
--------------------------------------------------
"
$ocultarop=Read-Host "Enter the desired option"


if ($ocultarop -eq "1") {
$ocultar=$true
} else {
$ocultar=$false
}
foreach ($item in $usuarios) {
$origem=$item.userprincipalname


Write-Progress -Activity "Changing $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

Write-Host "Changing attributes of " -NoNewline; Write-Host $origem -ForegroundColor Green;
Set-Mailbox -Identity $origem -HiddenFromAddressListsEnabled $ocultar

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Write-Progress -Activity "Changing $origem" -Completed

Menu
}


# 5- Habilitar archive e litigation hold
function cinco {

<# CSV HEADERS REQUIRED 
UserPrincipalName
#>

$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

Write-Host "Selected option: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

#================== Abrir opcoes ==================#
write-host -ForegroundColor Yellow "
--------------------------------------------------
 1- Archive and Litigation Hold
 2- Only Archive
 3- Only Litigation Hold
--------------------------------------------------
"
$archivelitigationop=Read-Host "Enter the desired option"

#Se selecionou opcao com litigation, perguntar o tempo
if($archivelitigationop -eq "1" -or $archivelitigationop -eq "3") {
Write-Host "Type the litigation hold time or leave empty for unlimited"
$litigationtempo=Read-Host "How long for litigation? (empty to unlimited)"
}


foreach ($item in $usuarios) {
$origem=$item.userprincipalname

Write-Progress -Activity "Changing $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

Write-Host "Changing attributes of " -NoNewline; Write-Host $origem -ForegroundColor Green;

#Se selecionado litigation, ler se especificou o tempo e rodar comando correto
if($archivelitigationop -eq "1" -or $archivelitigationop -eq "3") {
if($litigationtempo -ne "") {
Set-Mailbox -Identity $origem -LitigationHoldEnabled $true -LitigationHoldDuration $litigationtempo
} else {
#Ativando litigation por tempo indeterminado
Set-Mailbox -Identity $origem -LitigationHoldEnabled $true
}
}
#Rodar comando se selecionou opcao com archive
if($archivelitigationop -eq "1" -or $archivelitigationop -eq "2") {
Enable-Mailbox -Identity $origem -Archive
}

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Write-Progress -Activity "Changing $origem" -Completed

Menu
}

# 6- Habilitar autoexpand archive
function seis {

<# CSV HEADERS REQUIRED 
UserPrincipalName
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

Write-Host "Selected option: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()


foreach ($item in $usuarios) {
$origem=$item.userprincipalname

Write-Progress -Activity "Changing $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

Write-Host "Changing attributes of " -NoNewline; Write-Host $origem -ForegroundColor Green;
Enable-Mailbox $origem -AutoExpandingArchive

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Write-Progress -Activity "Changing $origem" -Completed

Menu
}


#7- Desabilitar política de email address do Exchange
function sete {

<# CSV HEADERS REQUIRED 
UserPrincipalName
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

Write-Host "Selected option: $opcao" -ForegroundColor Cyan

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

#Variaveis de erro
$MbxMailPolicyError=@()

#Inicio do contador
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname

Write-Progress -Activity "Changing $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

if(Set-Mailbox -Identity $origem -EmailAddressPolicyEnabled $false -ErrorAction SilentlyContinue) {

Write-Host "Changing attributes of " -NoNewline; Write-Host $origem -ForegroundColor Green;

} else {
Write-Host $origem -ForegroundColor Red
$MbxMailPolicyError+=$origem

}

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Write-Progress -Activity "Changing $origem" -Completed

"
--------------------------------------------------------------
"
Write-Host $MbxMailPolicyError.count "5 errors were found when removing the policy" -ForegroundColor Red
Write-Host "Error email list" -ForegroundColor Cyan
Write-Host $MbxMailPolicyError

#================== Abrir opcoes ==================#
"
--------------------------------------------------------------"
write-host -ForegroundColor Yellow " Do you want to take some action?"
" 1- Export error list (.txt) - $caminholog
 2- Display error email list on screen
 3- Copy to clipboard error email List
 4- Go back to main menu
--------------------------------------------------------------
"
$exportar_op7=Read-Host "Choose the desired option"
if($exportar_op7 -eq "1") {
#exportar
$MbxMailPolicyError | Out-File $caminholog\MbxMailPolicyError-$date.txt
} elseif($exportar_op7 -eq "2") {
#exibir na tela
$MbxMailPolicyError
} elseif($exportar_op7 -eq "3") {
#copiar para area de transferencia
$MbxMailPolicyError | clip
} else {

Menu
}
}


# 8 Checar se usuário foi migrado para o Exchange Online ou está no Exchange Server (somente híbrido)
function oito {
<# CSV HEADERS REQUIRED 
UserPrincipalName
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

Write-Host "Option selected: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

#Validando se há conexao ativa com o Office 365 (MsolService)
try
{
    Get-MsolDomain -ErrorAction Stop > $null
}
catch 
{
    Write-Output "Connecting to Office 365..."
    Connect-MsolService
}


foreach ($item in $usuarios) {

$origem=$item.userprincipalname
Write-Progress -Activity "Checking $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

#Inicio da captura via MsolService
Get-MsolUser -UserPrincipalName $origem |
Select-Object -Property DisplayName, UserPrincipalName, isLicensed,
                        @{label='MailboxLocation';expression={
                            switch ($_.MSExchRecipientTypeDetails) {
                                      1 {'Onprem'; break}
                                      2147483648 {'Office365'; break}
                                      default {'Desconhecido'}
                                  }
                        }} #| Export-Csv C:\Stats.csv

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Write-Progress -Activity "Checking $origem" -Completed

Menu
}

# 9 Criar SMTP secundário nas mailboxes
function nove {
<# CSV HEADERS REQUIRED 
UserPrincipalName
SMTPSecundario
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

Write-Host "Selected option: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

$SMTPSecundarioSim=@()
$SMTPSecundarioNao=@()
$SMTPSecundarioNaoUPN=@()
$SMTPSecundarioNaoNew=@()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname
$SMTP = $item.SMTPSecundario

Write-Progress -Activity "$origem -> $SMTP" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

if(Set-Mailbox -Identity "$origem" -EmailAddresses @{add="$SMTP"} -ErrorAction SilentyContinue) {
$comporSmtpOrigem=$origem+$delimiter+$SMTPSecundario
$SMTPSecundarioSim+=$comporSmtpOrigem
Write-Host "Changing mailbox " -NoNewline; Write-Host $origem -ForegroundColor Yellow -NoNewline; Write-Host " adicionando " -NoNewline; Write-Host $SMTP -ForegroundColor Green;

} else {
$comporSmtpOrigem=$origem+$delimiter+$SMTPSecundario
$SMTPSecundarioNao+=$comporSmtpOrigem

Write-Host $origem -ForegroundColor red
}


$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}

Write-Host $SMTPSecundarioSim.count "mailboxes were successfully changed" -ForegroundColor Green
Write-Host $SMTPSecundarioNao.count "mailboxes had problems" -ForegroundColor Red

$elapsedTime.stop()
Write-Progress -Activity "$origem -> $SMTP" -Completed

#================== Abrir opcoes ==================#
"
--------------------------------------------------------------"
write-host -ForegroundColor Yellow " Deseja realizar alguma ação?"
" 1- Export list of errors and successes (.csv) - $caminholog
 2- Display error list on screen
 3- Copy to clipboard error list
 4- Copy to clipboard completed list
 5- Go back to main menu
--------------------------------------------------------------
"
$exportar_op9=Read-Host "Enter the desired option"
if($exportar_op9 -eq "1") {
#exportar
$SMTPSecundarioNao | Export-Csv $caminholog\SMTPSecundarioNao-$date.csv
$SMTPSecundarioSim | Export-Csv $caminholog\SMTPSecundarioSim-$date.csv
} elseif($exportar_op9 -eq "2") {
#exibir na tela
$SMTPSecundarioNao

<#$a = @()
foreach($erro in $SMTPSecundarioNao) {
$a += [pscustomobject]@{a = 1; b = 2}
}
$a | Format-Table#>


} elseif($exportar_op9 -eq "3") {
#copiar para area de transferencia
$SMTPSecundarioNao | clip
}  elseif($exportar_op9 -eq "4") {
#copiar para area de transferencia
$SMTPSecundarioSim | clip
} else {

Menu
}


}


# 10- Converter caixa de Usermailbox para Sharedmailbox
function dez {
<# CSV HEADERS REQUIRED 
UserPrincipalName
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

Write-Host "Selected option: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname

Write-Progress -Activity "Changing $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Converting mailbox " -NoNewline; Write-Host $origem -ForegroundColor Yellow -NoNewline;
Set-Mailbox -Identity $origem -Type Shared

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}
$elapsedTime.stop()
Write-Progress -Activity "Changing $origem" -Completed

Menu
}

# 11- Converter caixa de Sharedmailbox para Usermailbox
function onze {
<# CSV HEADERS REQUIRED 
UserPrincipalName
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

Write-Host "Selected option: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname

Write-Progress -Activity "Changing $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Converting " -NoNewline; Write-Host $origem -ForegroundColor Yellow -NoNewline;
Set-Mailbox -Identity $origem -Type Regular

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}
$elapsedTime.stop()
Write-Progress -Activity "Changing $origem" -Completed

Menu
}


# 12- Adicionar permissoes em sharedmailbox
function doze {
<# CSV HEADERS REQUIRED 
UserPrincipalName (UPN from Sharedmailbox)
UserUPN (User to receive the permission)
Permission (Permission type)
---- FullAccess
---- SendAs
---- Send on Behalf Of
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0


Write-Host "Selected option: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

$errosharedpermission=@()

foreach ($item in $usuarios) {

$permissao=$item.Permission
$usuario=$item.UserUPN
$sharedmbx=$item.UserPrincipalName

Write-Progress -Activity "$sharedmbx - $usuario - $permissao" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Adding permission " -NoNewline; Write-Host $permissao -ForegroundColor Green -NoNewline; Write-Host " to " -NoNewline; Write-Host $usuario -ForegroundColor Yellow -NoNewline; Write-Host " in the mailbox " -NoNewline; Write-Host $sharedmbx -ForegroundColor Yellow;

if($permissao -eq "FullAccess") {
#FullAccess
Add-MailboxPermission -Identity $sharedmbx -User $usuario -AccessRights FullAccess -InheritanceType All
#Write-Host "Full Access $sharedmbx para $usuario"
} elseif($permissao -eq "SendAs") {
#SendAs
Add-RecipientPermission $sharedmbx –Trustee $usuario –AccessRights SendAs –confirm:$false
#Write-Host "Send As $sharedmbx para $usuario"
} elseif($permissao -eq "Send on Behalf Of") {
#SendOnBehalf
Set-Mailbox $sharedmbx -GrantSendOnBehalfTo $usuario
#Write-Host "Send On Behalf $sharedmbx para $usuario"
}

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}
$elapsedTime.stop()
Write-Progress -Activity "$sharedmbx - $usuario - $permissao" -Completed

Menu
}


#13- Criar sharedmailbox
function treze {

<# CSV HEADERS REQUIRED 
UserPrincipalName
Name (Display Name)
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

Write-Host "Selected option: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname
$nome=$item.name
$displayname=$item.displayname

Write-Progress -Activity "Creating $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Creating sharedmailbox " -NoNewline; Write-Host $origem -ForegroundColor Yellow;
New-Mailbox -Shared -Name $nome -PrimarySmtpAddress $origem -DisplayName $displayname

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}
$elapsedTime.stop()
Write-Progress -Activity "Creating $origem" -Completed

Menu


}

# 14- Remover regra da inbox do usuario
function quatorze {
<# CSV HEADERS REQUIRED 
UserPrincipalName
InboxRule (Rule name to be removed)
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

Write-Host "Selected option: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname

Write-Progress -Activity "Changing $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Changing mailbox " -NoNewline; Write-Host $origem -ForegroundColor Yellow;
Remove-InboxRule -Mailbox $origem -Identity $item.InboxRule -AlwaysDeleteOutlookRulesBlob -Confirm:$false


$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}
$elapsedTime.stop()
Write-Progress -Activity "Changing $origem" -Completed


Menu
}


# 15- Exportar permissoes de caixas de correio (Usermbx e/ou Sharedmbx)
function quinze {
<# CSV HEADERS REQUIRED 
UserPrincipalName - opcional
#>

#================== Abrir opcoes ==================#
"
--------------------------------------------------------------"
write-host -ForegroundColor Yellow " What do you want to do"
" 1- Export all Sharedmailbox permissions
 2- Export all Usermailbox permissions
 3- Export all Sharedmailbox AND Usermailbox permissions
 4- Export permissions selecting an .csv file
 5- Go back to main menu
--------------------------------------------------------------
"
$exportar_op15=Read-Host "Choose the desired option"
if($exportar_op15 -eq "1") {
#get sharedmbx
Write-Host "Fetching mailboxes"
$usuarios=Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -PropertySet Delivery -Properties RecipientTypeDetails, DisplayName | Select DisplayName, UserPrincipalName, RecipientTypeDetails, GrantSendOnBehalfTo
} elseif ($exportar_op15 -eq "2") {
#get usermbx
Write-Host "Fetching mailboxes"
$usuarios=Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -PropertySet Delivery -Properties RecipientTypeDetails, DisplayName | Select DisplayName, UserPrincipalName, RecipientTypeDetails, GrantSendOnBehalfTo
} elseif ($exportar_op15 -eq "3") {
Write-Host "Fetching mailboxes"
$usuarios=Get-ExoMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited -PropertySet Delivery -Properties RecipientTypeDetails, DisplayName | Select DisplayName, UserPrincipalName, RecipientTypeDetails, GrantSendOnBehalfTo
} elseif ($exportar_op15 -eq "4") {

$File = Get-FileName
Validate-CSVHeaders
$usuarios = Import-csv -Path $File -Delimiter $delimiter

} elseif ($exportar_op15 -eq "5") {
Menu
}

If ($usuarios.Count -eq 0) { 
    Write-Host "No mailboxes found. Returning to main menu..." -ForegroundColor Red
    Start-Sleep 2
    Menu
} 

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$usuarios.Count
$CurrentItem = 0
$PercentComplete = 0

Write-Host "Option selected: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()
$Report = [System.Collections.Generic.List[Object]]::new() #Create output file 

foreach ($item in $usuarios) {

$origem=$item.userprincipalname
$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

Write-Progress -Activity "Checking $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
#Write-Host "Checando conta " -NoNewline; Write-Host $origem -ForegroundColor Yellow;
$Permissions = Get-ExoRecipientPermission -Identity $origem | ? { $_.Trustee -ne "NT AUTHORITY\SELF" }
    If ($Null -ne $Permissions) {
        # Grab information about SendAs permission and output it into the report
        ForEach ($Permission in $Permissions) {
            $ReportLine = [PSCustomObject] @{
                Mailbox     = $item.DisplayName
                UPN         = $item.UserPrincipalName
                Permission  = $Permission | Select -ExpandProperty AccessRights
                AssignedTo  = $Permission.Trustee
                MailboxType = $item.RecipientTypeDetails 
            } 
            $Report.Add($ReportLine) 
        }
    }

    # Grab information about FullAccess permissions
    $Permissions = Get-ExoMailboxPermission -Identity $origem | ? { $_.User -ne "NT AUTHORITY\SELF" }    
    If ($Null -ne $Permissions) {
        # Grab each permission and output it into the report
        ForEach ($Permission in $Permissions) {
            $ReportLine = [PSCustomObject] @{
                Mailbox     = $item.DisplayName
                UPN         = $item.UserPrincipalName
                Permission  = $Permission | Select -ExpandProperty AccessRights
                AssignedTo  = $Permission.User
                MailboxType = $item.RecipientTypeDetails 
            } 
            $Report.Add($ReportLine) 
        }
    } 

    # Check if this mailbox has granted Send on Behalf of permission to anyone
    If (![string]::IsNullOrEmpty($item.GrantSendOnBehalfTo)) {
        ForEach ($Permission in $item.GrantSendOnBehalfTo) {
            $ReportLine = [PSCustomObject] @{
                Mailbox     = $item.DisplayName
                UPN         = $item.UserPrincipalName
                Permission  = "Send on Behalf Of"
                AssignedTo  = (Get-ExoRecipient -Identity $Permission).PrimarySmtpAddress
                MailboxType = $item.RecipientTypeDetails 
            } 
            $Report.Add($ReportLine) 
        }
    }

}

$elapsedTime.stop()
Write-Progress -Activity "Checking $origem" -Completed

$Report | Sort -Property @{Expression = { $_.MailboxType }; Ascending = $False }, Mailbox | Export-CSV $caminholog\MailboxPermissions-$date.csv -NoTypeInformation -Encoding UTF8 -Delimiter $delimiter
Write-Host $usuarios.Count "mailboxes scanned."
Write-Host "CSV File exported to $caminholog\MailboxPermissions-$date.csv" -ForegroundColor Cyan

Menu
}

# 16- Verificar se archive e/ou litigation hold está habilitado
function dezesseis {
<# CSV HEADERS REQUIRED 
UserPrincipalName - required
#>

Write-Host "Option selected: $opcao" -ForegroundColor Cyan
$Result=@() 
$mailboxes=@()
$MbxNao=@()

#================== Abrir opcoes ==================#
"
--------------------------------------------------------------"
write-host -ForegroundColor Yellow " What do you want to do"
" 1- Check all Usermailbox
 2- Check usermailbox selecting an .csv file
 3- Go back to main menu
--------------------------------------------------------------
"
$exportar_op16=Read-Host "Choose the desired option"
if($exportar_op16 -eq "1") {
#pegar usermbx
Write-Host "Fetching mailboxes..."
#ler todas as mbx
$mailboxes=Get-Mailbox -RecipientTypeDetails Usermailbox -ResultSize Unlimited | Select DisplayName, UserPrincipalName, PrimarySmtpAddress, ArchiveStatus, ArchiveName, ArchiveState, ArchiveWarningQuota, ArchiveQuota, AutoExpandingArchiveEnabled, LitigationHoldEnabled, LitigationHoldDuration
} elseif ($exportar_op16 -eq "2") {
#ler csv
$File = Get-FileName
Validate-CSVHeaders
$importcsv16 = Import-csv -Path $File -Delimiter $delimiter
Write-Host "Fetching mailboxes..."
foreach ($item in $importcsv16) {
#ler mbx do csv e agregar a variavel mailboxes
$mailboxes +=Get-Mailbox $item.userprincipalname | Select DisplayName, UserPrincipalName, PrimarySmtpAddress, ArchiveStatus, ArchiveName, ArchiveState, ArchiveWarningQuota, ArchiveQuota, AutoExpandingArchiveEnabled, LitigationHoldEnabled, LitigationHoldDuration
}

} elseif ($exportar_op15 -eq "3") {
Menu
}

If ($mailboxes.Count -eq 0) { 
    Write-Host "No mailboxes found. Returning to main menu..." -ForegroundColor Red
    Start-Sleep 2
    Menu
} 

#Contador da barra de progresso
#Zerar variaveis
$TotalItems=$mailboxes.Count
$CurrentItem = 0
$PercentComplete = 0


$elapsedTime = [system.diagnostics.stopwatch]::StartNew()
$Report = [System.Collections.Generic.List[Object]]::new() #Create output file 

foreach ($item in $mailboxes) {

$mbx = $item
$upn16=$mbx.userprincipalname
$size_mbx = $null
$size_arc = $null

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

Write-Progress -Activity "Checking $upn16" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% complete: $CurrentItem of $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao ##
#Write-Host "Checando conta " -NoNewline; Write-Host $origem -ForegroundColor Yellow;

#Escreve o progresso
Write-Host "Processando" $mbx.UserPrincipalName

#pegar dados da mailbox
$mbs_mbx = Get-MailboxStatistics $mbx.UserPrincipalName
if ($mbs_mbx.TotalItemSize -ne $null){
$size_mbx = [math]::Round(($mbs_mbx.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1MB),2)
}else{
$size_mbx = 0 }


#se archive estiver ativo
if ($mbx.ArchiveStatus -eq "Active"){
#pegar dados do archive
$mbs_arc = Get-MailboxStatistics $mbx.UserPrincipalName -Archive

#le o tamanho do archive
if ($mbs_arc.TotalItemSize -ne $null){
$size_arch = [math]::Round(($mbs_arc.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1MB),2)
}else{
$size_arch = 0 }
}

#Monta o resultado para exportar no csv
$Result += New-Object -TypeName PSObject -Property $([ordered]@{ 
DisplayName = $mbx.DisplayName
UserPrincipalName = $mbx.UserPrincipalName
PrimarySmtpAddress = $mbx.PrimarySmtpAddress
MailboxSizeInMB = $size_mbx
ArchiveStatus =$mbx.ArchiveStatus
ArchiveName =$mbx.ArchiveName
ArchiveState =$mbx.ArchiveState
ArchiveMailboxSizeInMB = $size_arc
ArchiveWarningQuota=if ($mbx.ArchiveStatus -eq "Active") {$mbx.ArchiveWarningQuota} Else { $null } 
ArchiveQuota = if ($mbx.ArchiveStatus -eq "Active") {$mbx.ArchiveQuota} Else { $null } 
AutoExpandingArchiveEnabled=$mbx.AutoExpandingArchiveEnabled
LitigationHoldEnabled=$mbx.LitigationHoldEnabled
LitigationHoldDuration=if ($mbx.LitigationHoldEnabled -eq "True") {$mbx.LitigationHoldDuration} Else { $null }
})


} 
$elapsedTime.stop()
Write-Progress -Activity "Checking $upn16" -Completed

$Result | Export-CSV $caminholog\MailboxSize-ArchiveLitigationReport-$date.csv -NoTypeInformation -Encoding UTF8 -Delimiter $delimiter
Write-Host $mailboxes.Count "mailboxes scanned."
Write-Host "CSV File exported to $caminholog\MailboxSize-ArchiveLitigationReport-$date.csv" -ForegroundColor Cyan



Menu
}

#Inicia o script
Login
Menu



#Finaliza o log do transcript
$dateF = Get-Date -Format "dd/MM/yyyy - HH:mm:ss"
Write-Host "Processo finalizado em" $dateF -ForegroundColor Magenta
Stop-Transcript
