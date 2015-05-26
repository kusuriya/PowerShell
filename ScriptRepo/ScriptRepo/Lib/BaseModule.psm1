<#############################################
## BaseModule.psm1  						##
## By: Jason Barbier <jabarb@serversave.us> ##
##											##
##											##
#############################################>

#Create Registry Key 
$strMachineName = 'vsaphxtestjm01' 
$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strMachineName) 
$objRegKey = $objReg.CreateSubKey("SOFTWARE\Testing-1-2-3") 

#Set Reg key Value 
$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strMachineName) 
$objRegKey = $objReg.openSubKey("SOFTWARE\Testing-1-2-3",$true) 
$objRegkey.setvalue('foo','0','Dword')

#Point OS to windows install location on the network 
$strMachineName = 'vsaphxtestjm01' 
$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strMachineName) 
$objRegKey = $objReg.openSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Setup",$true) 
$objRegkey.setvalue('SourcePath','\\mgmt.msft.net\services\MSNPLAT\gold\WIN2008R2SP0\winbuilds\7600\SP0\0\AMD64\bits','String')

$objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strMachineName) 
$objRegKey = $objReg.openSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Setup",$true) 
$objRegkey.setvalue('SourcePath','\\mgmt.msft.net\services\MSNPLAT\gold\WIN2008R2SP0\winbuilds\7600\SP0\0\AMD64\bits','String')

