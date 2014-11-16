<##################################
# Transport Test and notifcation  #
# TransportTest.ps1               #
# 								  #
###################################>

<# Global Variables #>
$Servicename = "spooler"
$Server = "."
$objServ = Get-WmiObject Win32_Service -ComputerName $Server| Where-Object { $_.Name -match "$servicename"}
$passwordsecured = "01000000d08c9ddf0115d1118c7a00c04fc297eb010000005959c511b0e7534591423b55b23377340000000002000000000003660000c00000001000000054ec91e26383b9e908dff13f1347c1fe0000000004800000a000000010000000a89b7f34f896063e57428e99827ee9f6200000006b40ddee37d588a7df7feed772d4345013bec3c42a387a64c51e98eb1453247d140000003f2c66a2747cded4b1fa9500c6bf09f4d290d804"
$password = $passwordsecured|ConvertTo-SecureString
[string]$CertificateThumbprint = "9FDCB2AAC4BE4E77690FB04E422BB6B5E740DBB6"
$SmtpServer = "smtpi-blu.msn.com"
[int]$SMTPPort = 25028
Function SendMail(){
	
	<# Mailclient Config #>
	$SmtpClient = new-object system.net.mail.smtpClient
	$SmtpClient.host = $smtpserver
	$SmtpClient.Port = $SMTPPort
	$netcreds = New-Object system.Net.NetworkCredential -ArgumentList "_GFS-RAS",$password
	$certs = Get-ChildItem cert:\LocalMachine\My|where {$_.Thumbprint -eq $CertificateThumbprint}
	$SmtpClient.ClientCertificates.Add($certs)
	$SmtpClient.UseDefaultCredentials = $false
	$smtpClient.Credentials = $netcreds
	$SmtpClient.EnableSsl = $true
	
	<# Message Config #>
	$MailFrom = "gfs-smtp@microst.com"
	[string]$RCPTTo = "v-jasonb@microsoft.com,v-jmatt@microsoft.com"
	$Subject = "!!!The MSExchange Transport Service has been stopped potential sev1!!!"
	$Body = "The MS exchange transport has stopped on $hostname ! This is a severity one issue please escalate immidately"
	
	try {
		Write-Host "Sending Message using host $SmtpServer"
		$SmtpClient.Send($Mailfrom,$RCPTto,$Subject,$Body)
	}
	catch [System.Exception]{
		$_.Exception.InnerException.Message
		$_.Exception.Message
		$_.Exception.statusCode
	}
}
write-host "The service is $objServ.State"
if($objServ.State -notmatch “Running”){
	Write-Host "Restarting Service"
   	$objserv.StartService()
}
write-host "After the restart attempt the service is $objServ.State"
if($objServ.State -notmatch “Running”){
		Write-Host "The service will not start!"
    	SendMail
	}