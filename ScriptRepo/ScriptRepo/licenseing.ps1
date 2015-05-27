function Get-TerminalServerLicensingInfo
{
    param
    (
        [Parameter(mandatory=$true)][string]$computer = "LocalHost",
        [Parameter(mandatory=$true)][string]$namespace = "root\CIMV2\TerminalServices"
    )

    $TSlic = Get-WmiObject -class Win32_TerminalServiceSetting -computername $computer -namespace $namespace 
    if ($TSlic.TerminalServerMode -eq "1")
    {
        Write-Output "TerminalServerMode: Yes"
    } 
    else 
    {
        Write-Output "TerminalServerMode: No"
    } 
    if ($TSlic.AllowTSConnections -eq "1")
    {
        Write-Output "AllowTSConnections: Yes"
    } 
    else 
    {
        Write-Output "AllowTSConnections: No"
    } 
    Write-Output "Licensing Server: " ($TSlic.GetSpecifiedLicenseServerList()).SpecifiedLSList 
    Write-Output "RD Licensing Type: " $TSlic.LicensingName
}