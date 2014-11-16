<#############################################
## BaseModule.psm1  						##
## By: Jason Barbier <jabarb@serversave.us> ##
##											##
##											##
#############################################>

Function add-regkey($ComputerName, $key, $Name, $value, $type){
	#Create Registry Key 
	$strMachineName = $ComputerName 
	$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strMachineName) 
	$objRegKey = $objReg.CreateSubKey($key) 
	#Set Reg key Value 
	$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strMachineName)
	$objRegKey = $objReg.openSubKey($key,$true) 
	$objRegkey.setvalue($name,$value,$type)
}
Function set-WindowsInstallPath($ComputerName, $SourcePath, $Domain){
	#Point OS to windows install location on the network 
	if (!$SourcePath){
		Read-Host "Source Path is required:"
	}
	$strMachineName = $ComputerName
	$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strMachineName) 
	$objRegKey = $objReg.openSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Setup",$true) 
	$objRegkey.setvalue('SourcePath',$SourcePath,'String')
	$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strMachineName) 
	$objRegKey = $objReg.openSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Setup",$true) 
	$objRegkey.setvalue('SourcePath',$SourcePath,'String')
}
Function SendMail($smtpserver,$SMTPPort,[string]$MailFrom,[String]$RcptTo,[string]$Subject,[String]$body,[switch]$DefaultCredentials = $false,[switch]$TLS=$true,[string]$CertificateThumbprint){
	
	<# Mailclient Config #>
	$SmtpClient = new-object system.net.mail.smtpClient
	$SmtpClient.host = $smtpserver
	$SmtpClient.Port = $SMTPPort
	$netcreds = New-Object system.Net.NetworkCredential -ArgumentList "_GFS-RAS",$password
	
	if ($TLS -eq $true){
		Write-Host "Retriveing $CertificateThumbprint"
		$certs = Get-ChildItem cert:\LocalMachine\My|where {$_.Thumbprint -eq $CertificateThumbprint}
		Write-Host -NoNewline "Adding $CertificateThumbprint: "
		$SmtpClient.ClientCertificates.Add($certs)
		Write-Host `n
		$smtpClient.Credentials = $netcreds
		$SmtpClient.EnableSsl = $true
	}
	elseif ($TLS -eq $false){
		$SmtpClient.EnableSsl = $false
	}
	
	if ($DefaultCredentials -eq $true){
		$SmtpClient.UseDefaultCredentials = $true
	}
	elseif ($DefaultCredentials -eq $false){
		$SmtpClient.UseDefaultCredentials = $false
	}
	
	<# Message Config #>
	if (!$MailFrom){
		$MailFrom = Read-Host -Prompt "A Mail From: address is required: "
	}
	if (!$RCPTTo){
		$RCPTTo = Read-Host -Prompt "A RCPT To address is required: "
	}
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