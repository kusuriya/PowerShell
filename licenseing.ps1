$computer = "LocalHost" 
$namespace = "root\CIMV2\TerminalServices" 
$TSlic = Get-WmiObject -class Win32_TerminalServiceSetting -computername $computer -namespace $namespace 
If ($TSlic.TerminalServerMode -eq "1"){Write-Host "TerminalServerMode: Yes"} 
Else {Write-Host "TerminalServerMode: No"} 
If ($TSlic.AllowTSConnections -eq "1"){Write-Host "AllowTSConnections: Yes"} 
Else {Write-Host "AllowTSConnections: No"} 
Write-Host "Licensing Server: " ($TSlic.GetSpecifiedLicenseServerList()).SpecifiedLSList 
Write-Host "RD Licensing Type: " $TSlic.LicensingName
