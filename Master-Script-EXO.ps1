Write-Host "

███████╗██╗░░██╗░█████╗░██╗░░██╗░█████╗░███╗░░██╗░██████╗░███████╗  ░█████╗░███╗░░██╗██╗░░░░░██╗███╗░░██╗███████╗
██╔════╝╚██╗██╔╝██╔══██╗██║░░██║██╔══██╗████╗░██║██╔════╝░██╔════╝  ██╔══██╗████╗░██║██║░░░░░██║████╗░██║██╔════╝
█████╗░░░╚███╔╝░██║░░╚═╝███████║███████║██╔██╗██║██║░░██╗░█████╗░░  ██║░░██║██╔██╗██║██║░░░░░██║██╔██╗██║█████╗░░
██╔══╝░░░██╔██╗░██║░░██╗██╔══██║██╔══██║██║╚████║██║░░╚██╗██╔══╝░░  ██║░░██║██║╚████║██║░░░░░██║██║╚████║██╔══╝░░
███████╗██╔╝╚██╗╚█████╔╝██║░░██║██║░░██║██║░╚███║╚██████╔╝███████╗  ╚█████╔╝██║░╚███║███████╗██║██║░╚███║███████╗
╚══════╝╚═╝░░╚═╝░╚════╝░╚═╝░░╚═╝╚═╝░░╚═╝╚═╝░░╚══╝░╚═════╝░╚══════╝  ░╚════╝░╚═╝░░╚══╝╚══════╝╚═╝╚═╝░░╚══╝╚══════╝

#######################################################################################################################
##                                                                                                                   ##
## SCRIPT DE USO PARA O EXCHANGE ONLINE                                                                              ##
## DESENVOLVIDO POR: VICTOR MARTINS                                                                                  ##
## https://www.victornanuvem.com/                                                                                    ##
##                                                                                                                   ##
## VERSÃO 0.4                                                                                                        ##
##                                                                                                                   ##
## Descrição: Esse script realiza tarefas baseado em arquivos CSV                                                    ##
## Antes de usar o script leia as definições de cada função, qualquer ação incorreta pode prejudicar o seu ambiente  ##
##                                                                                                                   ##
#######################################################################################################################"-ForegroundColor Yellow

#Configuracoes
$delimiter=";"
$caminholog="D:\temp\logs"

#Funcao para logar no Exchange Online
function Login {
#Conectar e logar no ExchangeOnline (MFA)
$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
If ($isconnected -eq "True") {
#Connect-ExchangeOnline ##-ne
#Sessao aberta, valida
Write-Host "
Existe uma sessão aberta para o Exchange Online
Digite S para utilizar a mesma sessão ou qualquer outra tecla para iniciar uma nova sessão" -ForeGroundColor Green

#Usar mesma sessao?
$r = Read-Host "Digite 'S' para usar a mesma conexão"
	if($r.ToLower() -eq 's')
		{
			Write-Host "Utilizando a mesma sessão..." -ForeGroundColor Green
			return;	
		}
#Senao, remover sessao ativa
	$getsessions | Remove-PSSession
	    }

#Contador de erros de conexão
$f = 0
$i = 0
while ($f -eq 0) 
  {
	$i++
	if($i -eq 4)
       {
	     Write-host "Problemas de autenticação, não é possível iniciar o script" -ForegroundColor Red				
         #exit
         #Finaliza o script depois de 3 erros
         break;
        }
#Conectar ao Exchange Online
 try
    { 
      Connect-ExchangeOnline
      $f = 2
    }
#Se der erro
  catch
   {
	Write-host "Erro de senha ou permissão. Tente novamente." -ForegroundColor Red
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
$OpenFileDialog.filter = “Arquivos CSV (*.csv)| *.csv”
#$OpenFileDialog.ShowDialog() | Out-Null


 If ($OpenFileDialog.ShowDialog() -eq "Cancel") 
 {
  [System.Windows.Forms.MessageBox]::Show("Nenhum arquivo foi selecionado. Por favor selecione um arquivo!", "Error", 0, 
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
 Write-Host "===================== ATENÇÃO =====================
CSV Inválido, voltando ao menu inicial
===================================================" -ForegroundColor Red
 Start-Sleep 2
 Menu
    }
   }

#================== Abrir opcoes ==================#
function Menu {
"--------------------------------------------------"
Write-Host -ForegroundColor Yellow " Escolha a opção desejada"
" 1- Configurar redirect nas caixas de correio
 2- Apagar redirect das caixas de correio
 3- Checar se a mailbox existe
 4- Colocar ou removar hide da GAL
 5- Habilitar archive e litigation hold
 6- Habilitar autoexpanding archive
 7- Desabilitar política de email address do Exchange
 8- Checar se usuário foi migrado para o Exchange Online ou está no Exchange Server (Somente Hybrid)
 9- Adicionar SMTP secundário na Mailbox
 10- Converter caixa de Usermailbox para Sharedmailbox
 11- Converter caixa de Sharedmailbox para Usermailbox
 12- Adicionar permissões em Sharedmailbox
 13- Criar sharedmailbox
 14- Remover regra da inbox do usuário
 20- Sair
--------------------------------------------------
"

$opcao=Read-Host "Digite a opção desejada"

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
elseif ($opcao -eq "20") { sair }
#opcao invalida
else {
Write-Host "==================================================
Opção inválida! Selecione uma opção válida.
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

Write-Host "Processo iniciado em" $dateI -ForegroundColor Magenta

#sair
function sair {
"Saindo do script..."
break
}

#1- Configurar redirect das caixas de correio
function um {

<# HEADERS NECESSARIOS 
UserPrincipalName
NewUPN
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
$destino=$item.newupn

Write-Progress -Activity "Configurando redirect nas contas - $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

Write-Host "Criando redirect de " -NoNewline; Write-Host $origem -ForegroundColor Yellow -NoNewline; Write-Host " para " -NoNewline; Write-Host $destino -ForegroundColor Green;
Set-Mailbox $origem -ForwardingsmtpAddress $destino -DeliverToMailboxAndForward $False

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Menu
}


# 2- Apagar redirect das contas
function dois {

<# HEADERS NECESSARIOS 
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

Write-Progress -Activity "Apagando redirect da Executiva - $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete


Write-Host "Apagando redirect de " -NoNewline; Write-Host $origem -ForegroundColor Green;
Set-Mailbox $origem -ForwardingSmtpAddress $Null

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Menu
}

# 3- Checar todas as mailboxes
function tres {

<# HEADERS NECESSARIOS 
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


Write-Progress -Activity "Checando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

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

#================== Abrir opcoes ==================#
"
--------------------------------------------------------------
"
Write-Host "Mailboxes encontradas: " -NoNewline; Write-Host $MbxSim.count -ForegroundColor Green
Write-Host "Mailboxes não encontradas: " -NoNewline; Write-Host $Mbxnao.count -ForegroundColor Red
"
--------------------------------------------------------------"
write-host -ForegroundColor Yellow " Deseja realizar alguma ação?"
" 1- Exportar caixas encontradas e não encontradas (.txt) - $caminholog
 2- Copiar para área de transferência caixas encontradas
 3- Copiar para área de transferência caixas não encontradas
 4- Voltar ao menu inicial
--------------------------------------------------------------
"
$exportar_op3=Read-Host "Escolha a opção desejada"
if($exportar_op3 -eq "1") {

$MbxSim | Out-File $caminholog\MbxEncontrada-$date.txt
$MbxNao | Out-File $caminholog\MbxNaoEncontrada-$date.txt

} elseif($exportar_op3 -eq "2") {
$MbxSim | clip
} elseif($exportar_op3 -eq "3") {
$MbxNao | clip
} else {

Menu
}
}


# 4- Remover hide da GAL
function quatro {

<# HEADERS NECESSARIOS 
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
write-host -ForegroundColor Yellow " Deseja ocultar ou exibir"
" 1- Ocultar na GAL
 2- Exibir na GAL
--------------------------------------------------
"
$ocultarop=Read-Host "Digite a opção desejada"


if ($ocultarop -eq "1") {
$ocultar=$true
} else {
$ocultar=$false
}
foreach ($item in $usuarios) {
$origem=$item.userprincipalname


Write-Progress -Activity "Alterando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

Write-Host "Alterando atributos de " -NoNewline; Write-Host $origem -ForegroundColor Green;
Set-Mailbox -Identity $origem -HiddenFromAddressListsEnabled $ocultar

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Menu
}


# 5- Habilitar archive e litigation hold
function cinco {

<# HEADERS NECESSARIOS 
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

Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

#================== Abrir opcoes ==================#
write-host -ForegroundColor Yellow "
--------------------------------------------------
 1- Archive e Litigation Hold
 2- Somente Archive
 3- Somente Litigation Hold
--------------------------------------------------
"
$archivelitigationop=Read-Host "Digite a opção desejada"

#Se selecionou opcao com litigation, perguntar o tempo
if($archivelitigationop -eq "1" -or $archivelitigationop -eq "3") {
Write-Host "Preencha com o tempo do litigation hold ou deixe vazio para ilimitado"
$litigationtempo=Read-Host "Qual tempo do litigation? (vazio para ilimitado)"
}


foreach ($item in $usuarios) {
$origem=$item.userprincipalname

Write-Progress -Activity "Alterando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

Write-Host "Alterando atributos de " -NoNewline; Write-Host $origem -ForegroundColor Green;

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
Menu
}

# 6- Habilitar autoexpand archive
function seis {

<# HEADERS NECESSARIOS 
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

Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()


foreach ($item in $usuarios) {
$origem=$item.userprincipalname

Write-Progress -Activity "Alterando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

Write-Host "Alterando atributos de " -NoNewline; Write-Host $origem -ForegroundColor Green;
Enable-Mailbox $origem -AutoExpandingArchive

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()
Menu
}


#7- Desabilitar política de email address do Exchange
function sete {

<# HEADERS NECESSARIOS 
UserPrincipalName
#>

#Importar CSV
$File = Get-FileName

Validate-CSVHeaders

$usuarios = Import-csv -Path $File -Delimiter $delimiter

Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan

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

Write-Progress -Activity "Alterando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

if(Set-Mailbox -Identity $origem -EmailAddressPolicyEnabled $false -ErrorAction SilentlyContinue) {

Write-Host "Alterando atributos de " -NoNewline; Write-Host $origem -ForegroundColor Green;

} else {
Write-Host $origem -ForegroundColor Red
$MbxMailPolicyError+=$origem

}

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)

}
$elapsedTime.stop()

"
--------------------------------------------------------------
"
Write-Host "Foram encontrados" $MbxMailPolicyError.count "erros na remoção da policy" -ForegroundColor Red
Write-Host "Lista de e-mails com erro" -ForegroundColor Cyan
Write-Host $MbxMailPolicyError

#================== Abrir opcoes ==================#
"
--------------------------------------------------------------"
write-host -ForegroundColor Yellow " Deseja realizar alguma ação?"
" 1- Exportar lista de erros (.txt) - $caminholog
 2- Exibir lista de erros na tela
 3- Copiar para área de transferência lista de erros
 4- Voltar ao menu inicial
--------------------------------------------------------------
"
$exportar_op7=Read-Host "Escolha a opção desejada"
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
<# HEADERS NECESSARIOS 
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

Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

#Validando se há conexao ativa com o Office 365 (MsolService)
try
{
    Get-MsolDomain -ErrorAction Stop > $null
}
catch 
{
    Write-Output "Conectando ao Office 365..."
    Connect-MsolService
}


foreach ($item in $usuarios) {

$origem=$item.userprincipalname
Write-Progress -Activity "Checando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

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
Menu
}

# 9 Criar SMTP secundário nas mailboxes
function nove {
<# HEADERS NECESSARIOS 
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

Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

$SMTPSecundarioSim=@()
$SMTPSecundarioNao=@()
$SMTPSecundarioNaoUPN=@()
$SMTPSecundarioNaoNew=@()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname
$SMTP = $item.SMTPSecundario

Write-Progress -Activity "$origem -> $SMTP" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

if(Set-Mailbox -Identity "$origem" -EmailAddresses @{add="$SMTP"} -ErrorAction SilentyContinue) {
$comporSmtpOrigem=$origem+$delimiter+$SMTPSecundario
$SMTPSecundarioSim+=$comporSmtpOrigem
Write-Host "Alterando conta " -NoNewline; Write-Host $origem -ForegroundColor Yellow -NoNewline; Write-Host " adicionando " -NoNewline; Write-Host $SMTP -ForegroundColor Green;

} else {
$comporSmtpOrigem=$origem+$delimiter+$SMTPSecundario
$SMTPSecundarioNao+=$comporSmtpOrigem

Write-Host $origem -ForegroundColor red
}


$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}

Write-Host "Foram alterados" $SMTPSecundarioSim.count "e-mails com sucesso" -ForegroundColor Green
Write-Host "O total de" $SMTPSecundarioNao.count "tiveram problemas" -ForegroundColor Red

$elapsedTime.stop()

#================== Abrir opcoes ==================#
"
--------------------------------------------------------------"
write-host -ForegroundColor Yellow " Deseja realizar alguma ação?"
" 1- Exportar lista de erros e sucesso (.csv) - $caminholog
 2- Exibir lista de erros na tela
 3- Copiar para área de transferência lista de erros
 4- Copiar para área de transferência lista concluída
 5- Voltar ao menu inicial
--------------------------------------------------------------
"
$exportar_op9=Read-Host "Escolha a opção desejada"
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
<# HEADERS NECESSARIOS 
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

Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname

Write-Progress -Activity "Alterando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Alterando conta " -NoNewline; Write-Host $origem -ForegroundColor Yellow -NoNewline;
Set-Mailbox -Identity $origem -Type Shared

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}
$elapsedTime.stop()
Menu
}

# 11- Converter caixa de Sharedmailbox para Usermailbox
function onze {
<# HEADERS NECESSARIOS 
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

Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname

Write-Progress -Activity "Alterando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Alterando conta " -NoNewline; Write-Host $origem -ForegroundColor Yellow -NoNewline;
Set-Mailbox -Identity $origem -Type Regular

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}
$elapsedTime.stop()
}


# 12- Adicionar permissoes em sharedmailbox
function doze {
<# HEADERS NECESSARIOS 
UserPrincipalName (UPN da Shared)
UPNUsuario (Usuario a receber a permissao)
Permission (tipo de permissao)
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


Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

$errosharedpermission=@()

foreach ($item in $usuarios) {

$permissao=$item.Permission
$usuario=$item.UPNUsuario
$sharedmbx=$item.UserPrincipalName

Write-Progress -Activity "$sharedmbx - $usuario - $permissao" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Adicionando permissão " -NoNewline; Write-Host $permissao -ForegroundColor Green -NoNewline; Write-Host " para " -NoNewline; Write-Host $usuario -ForegroundColor Yellow -NoNewline; Write-Host " na caixa " -NoNewline; Write-Host $sharedmbx -ForegroundColor Yellow;

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
Menu
}


#13- Criar sharedmailbox
function treze {

<# HEADERS NECESSARIOS 
UserPrincipalName
Nome (Display Name)
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

Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname
$nome=$item.nome

Write-Progress -Activity "Criando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Criando conta " -NoNewline; Write-Host $origem -ForegroundColor Yellow;
New-Mailbox -Shared -Name "Social Club" -PrimarySmtpAddress $origem

$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}
$elapsedTime.stop()
Menu


}

# 14- Remover regra da inbox do usuario
function quatorze {
<# HEADERS NECESSARIOS 
UserPrincipalName
InboxRule (Nome da regra a ser removida)
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

Write-Host "Opcão selecionada: $opcao" -ForegroundColor Cyan
$elapsedTime = [system.diagnostics.stopwatch]::StartNew()

foreach ($item in $usuarios) {

$origem=$item.userprincipalname

Write-Progress -Activity "Alterando $origem" -Status "$([string]::Format("Tempo em execução: {0:d2}:{1:d2}:{2:d2}", $elapsedTime.Elapsed.hours, $elapsedTime.Elapsed.minutes, $elapsedTime.Elapsed.seconds)) | $PercentComplete% completo: $CurrentItem de $TotalItems" -PercentComplete $PercentComplete

## Comando de execucao
Write-Host "Alterando conta " -NoNewline; Write-Host $origem -ForegroundColor Yellow;
Remove-InboxRule -Mailbox $origem -Identity $item.InboxRule -AlwaysDeleteOutlookRulesBlob -Confirm:$false


$CurrentItem++
$PercentComplete = [int](($CurrentItem / $TotalItems) * 100)
}
$elapsedTime.stop()
Menu
}

#Inicia o script
Login
Menu



#Finaliza o log do transcript
$dateF = Get-Date -Format "dd/MM/yyyy - HH:mm:ss"
Write-Host "Processo finalizado em" $dateF -ForegroundColor Magenta
Stop-Transcript
