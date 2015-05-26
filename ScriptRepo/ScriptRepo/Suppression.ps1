<###########################################################################
#
# NAME: 	Suppression.ps1
#
# AUTHOR:	Jeremy Matt; Jason Barbier
# EMAIL: 	v-jmatt; v-jasonb
#
# COMMENT: 	Script to suppress server(s) for 60 mins in iAdmin and SCOM 
#			monitoring. Any desired editing of this script is on you 
#			to perform.
#
# You have a royalty-free right to use, modify, reproduce, and
# distribute this script file in any way you find useful, provided that
# you agree that the creator, owner above has no warranty, obligations,
# or liability for such use.
#
# VERSION HISTORY:
# 1.0 10/29/2011 v-jmatt  - Initial release
# 2.0 11/15/2011 v-jasonb - Incorperated Tommy Vu's iadmin post script as a
#							function and streamlined it. Added a switch 
#							to allow changing the location of the bulk 
#							serverlist file. Fixed the if functions to allow 
#							bulk and non bulk suppression and unsuppression to work.
#
############################################################################>


<###############
####Parameters##
###############>

param
(
	[ARRAY]$Server,
	[SWITCH]$help,
	[SWITCH]$Bulk,
	[SWITCH]$Suppress,
	[SWITCH]$Unsuppress,
	[STRING]$ServerList = "C:\scripts\_input\Serverlist.txt"

)

<###############
#####Functions##
###############>


<###########
### Help ###
###########>
function Show-Help() {
$HelpText=@"

To run against a single machine:
	.\Suppression.ps1 <Computer> [-<Action>]
	
	To run against a serverlist:
	.\Suppression.ps1 [-bulk] [-<Action>]
	
	Options:	  				
	-Bulk		--	Run script against a serverlist located at 
				C:\Scripts\_input\serverlist.txt
	
	-Suppress	--	Starts Service
	-Unsuppress	--	Stops Service
"@
Write-Host $HelpText -ForegroundColor Yellow
Write-Host `n
	exit 1
}

<######################
### Supression Post ###
### Function        ###
######################>
Function sup ($action,$MOType,$Name,$Duration,$Reason) {
	$user = whoami
	$usersplit = $user.Split("\")
	[string] $UserLogName = $usersplit[1]
	[string] $FileID="supression.ps1"
	[string] $now=get-date
	[string] $rStr 
	$strURL = "http://XMLInterface/Post/MOSuppress.asp"
	$rStr = '<?xml version="1.0"?><XMLFILE UserLogName="'+$UserLogName+'" FileId="'+$FileID+'">' 
	$rStr=$rStr+'<MonitoredObject MOName="'+$Name+'" MOType="'+$MOType+'" Action="'+$action+'" SuppStartDtim="'+$now+'" ' 
	$rStr=$rStr+'SuppDurHrs="'+$Duration+'" ReasonDesc="'+$Reason+'"/></XMLFILE>'
	$objHTTP = new-object -com msxml2.xmlhttp
	$objHTTP.Open("Post",$strURL,0)
	$objHTTP.SetRequestHeader("Content-Type","application/x-www-form-urlencoded")
	write-host "Please wait, I am posting your data to XMLInterface..."
	write-host "I am currently feeding this XML to XMLInterface:"
	write-host $rStr
	$objHTTP.Send($rStr)
	Write-Host "The server responded: $objHTTP.status"
	switch ($objHTTP.status) {
	200 {"Processing was successful!"; break}
	202 {Write-error "Processing was partially successful. Check the site to ensure supression. Error information: $objHTTP.responsexml.xml"; break}
	602 {Write-error "Processing was unsuccessful. Invalid request. Error information: $objHTTP.responsexml.xml"; break}
	604 {Write-error "Processing was unsuccessful. Database error. Error information: $objHTTP.responsexml.xml"; break}
	610 {Write-error "Processing was unsuccessful. XmlInterface is currently not available. Please try later."; break}
	default {"Processing returned an unexpected HTTP status code - " + $_
		write-host $objHTTP.responsexml.xml
		}
	}
}
<#####################
### Call Supression###
### Function       ###
#####################>
Function call_suppression($Computers) {
$computers
Foreach ($Machine in $Computers){
		If ($Machine -like "*.*"){
			$ComputerSplit = $Machine.Split(".")
			$Computer = $ComputerSplit[0]
			$Computer}
		Else {$Computer = $Machine}
	
		If ($Suppress -match $true) {
					#$ScriptPath
					sup -action 'suppress' -motype 'server' -Name $Computer -Duration 1 -Reason 'This server will be rebooted for mandatory security patching and maintenance'
					\\PHX\services\ssg-mars\Utils\MaintenanceMode\MaintenanceMode.exe /RMS PROD /A $Computer 0 "Mandatory Security Patching - Rebooting $Computer" Now "+60M" /ST "Terminal"
				}
		Elseif ($Unsuppress -match $true) {
					#$ScriptPath
					sup -action 'unsuppress' -motype 'server' -Name $Computer -duration 0 -reason 'Maintenance Complete'
					\\PHX\services\ssg-mars\Utils\MaintenanceMode\MaintenanceMode.exe /RMS PROD /A $Computer /STOP
		}
	}

}

###############
##Script body##
###############
#Script Help#
if($help)
{
	Show-help

}
If ($Bulk -match $true) {
	$Server = Get-Content $ServerList
		If ($Server -like $null){Show-help}
		else {call_suppression($Server)}
}
Elseif ($Bulk -notmatch $true) {
		If ($Server -like $null){Show-help}
		else {call_suppression($Server)}
}