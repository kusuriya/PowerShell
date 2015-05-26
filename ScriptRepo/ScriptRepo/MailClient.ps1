<#####
Setting up the SMTP client
#####>
$SmtpClient = new-object system.net.mail.smtpClient
$SmtpServer = "<Selected Mail Server>"
$SmtpClient.host = $SmtpServer
$SmtpClient.Port = <Port for the server>
$hostname = hostname

<#####
Setting up the credentials and certs
#####>

#$netcreds = New-Object system.Net.NetworkCredential -ArgumentList "<Service Account>","<Password>" #you can use this method also to load credential
																									# from a file using variables.
#$netcreds = Get-Credential	#This will prompt for creds
$Password = ConvertTo-SecureString "Converted String or file with string" #To get this password you can do read-host -assecurestring|convertfrom-securestring
$netcreds = New-Object -typename System.Management.Automation.PSCredential -ArgumentList "<username>",$Password
$smtpClient.Credentials = $netcreds
$SmtpClient.UseDefaultCredentials = $false

$CertThumbprint = '<Cert Thumbprint>'
$certs = Get-ChildItem cert:\LocalMachine\My|where {$_.Thumbprint -eq $CertThumbprint}
$SmtpClient.ClientCertificates.Add($certs)
$SmtpClient.EnableSsl = $true


<#####
Set the contents of the Test Message
#####>
$MailFrom = "<Your Source Mail Address>"
$RCPTTo = "<Your Test Message address>"
$Subject = "Testing the mail relay using $SMTPServer."
$Data = @"
Mail relay was successful from $SMTPServer.
If you have recieved this message the relay is working as expected from $hostname .

Thank you,
SMTP Test Service.
"@ 




try { 
$SmtpClient.Send($MailFrom,$RCPTTo,$Subject,$Data) 
} 
catch [System.Exception] 
{ 
$_.Exception.InnerException.Message
#$_.Exception | Get-Member # show the exception's members to see what is available $_.Excepttion.InnerException 
                                      # display the exception's InnerException if it has one
}
