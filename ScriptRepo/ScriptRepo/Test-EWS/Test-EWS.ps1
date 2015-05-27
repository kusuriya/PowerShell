<######
Test-EWS.ps1
By Jason Barbier

.Description
    This script is to test EWS and will do so by connecting to a user's mailbox over EWS and
    pull a count of all items in their Inbox and the first 10 subjects.

    Please note this script requires the EWS managed API DLLs to be in the same location as the script.

.Parameter Mailbox
    This is the primary SMTP address of the mailbox you wish to test against

.Parameter Credential
    This is the credentials that have access to read the mailbox provided by Mailbox

.Parameter EWSUrl
    This is the EWS URL you wish to test. This is optional and if you do not supply it
    autodiscover based off the primary smtp address will be used.

.Parameter Authentication
    This switch is used to specify the credential package sent to the server, the valid options are:
        Default: uses the current windows credentials to negotiate auth
        NTLM: generates an NTLM package to submit for auth
        Negotiate: Uses standard Negotiate which will try Kerberos

.Parameter GetFolder
    This runs the test to get messages from a the Inbox to verify that we can get from a folder. This assumes you have
    Permissions to the mailbox you want info from.

.Parameter GetFreeBusy
    This runs the test to get FreeBusy info and a suggestion.

.Parameter GetHeaders
    This can be combined with any of the parameters to show the EWS headers

.Parameter GetDelegates
    This will get a Delegate list assuming you have permission to the mailbox.

.Parameter EWSTracing
    Enables advanced logging including the SOAP traces from EWS.

######>

[CmdletBinding()]

Param 
(
    [Parameter(Position=3)][string]$AuthenticationType = 'default',
    [Parameter(Position=4)]$Credential = $null,
    [Parameter(Position=2)][URI]$EWSUrl = $null,
    [switch]$EWSTracing = $false.
    [switch]$GetDelegates,
    [switch]$GetHeaders = $false,
    [switch]$GetFreeBusy = $false,
    [switch]$GetFolder = $false,
    [switch]$GetCalendar = $false,
    [Parameter(Position=1)][string]$Mailbox = $null
)

Import-Module -Name ".\Microsoft.Exchange.WebServices.dll"

try
{
    $EWS = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)
    $EWS.UserAgent = 'Test-EWS Suite/v0.2'
    if ($EWSTracing -eq $true)
    {
        $EWS.TraceEnabled = $true
        $EWS.TraceEnablePrettyPrinting = $true
        $EWS.TraceFlags = 'All'
    }

    # Configure and connect to EWS
    if (!$mailbox)
    {
        $mailbox = read-host "Please enter a valid Primary SMTP address"
    }

    
    if ($EWSUrl -eq $null)
    {
        Write-Warning "No EWS location Provided, attempting AutoDiscover"
        $EWS.AutoDiscoverURL($mailbox,{$true});
        Write-Verbose "Found EWS at:" $EWS.Url.AbsoluteUri
        $EWS.url.UserInfo
    }
    else
    {
        $EWS.Url = $EWSUrl
    }

    # Setup Credentials
    if (!$Credential)
    {
        $CredentialCache = New-Object System.Net.CredentialCache
        Switch ($AuthenticationType)
        {
            default
            {
                Write-Verbose "Using current Windows credentials."
                $EWS.UseDefaultCredentials = $true
            }
            NTLM
            {
                Write-Verbose "Creating an NTLM credential package."
                $CredentialCache.Add(($EWS.Url.AbsoluteUri),"NTLM",(get-credential))
                $EWS.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($CredentialCache)
            }
            Negotiate
            {
                Write-Verbose "Creating a Negotiate credential package"
                $CredentialCache.Add(($EWS.Url.AbsoluteUri),"Negotiate",(get-credential))
                $EWS.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($CredentialCache)
            }
        }
    }
    else 
    {
        # We Assume if there is somthing in the variable it is valid credentials, because hey you never know
        # the user may actually read the help.
        $EWS.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($Credential)       
    }




    # If they exist show the user where we got the info from.
    if ($EWS.HttpResponseHeaders.ContainsKey('X-FEServer') -eq $true)
    {
        Write-Verbose "FrontEnd Server:" $EWS.HttpResponseHeaders.Item('X-FEServer')
    }
    if ($EWS.HttpResponseHeaders.ContainsKey('X-BEServer') -eq $true)
    {
        Write-Verbose "Backend Server:" $EWS.HttpResponseHeaders.Item('X-BEServer')
    }
    if ($GetHeaders)
    {
        $EWS.HttpHeaders
        $EWS.HttpResponseHeaders
    }
    if ($GetFolder)
    {
    $defaultDisplaySet = 'From','Subject','DateTimeReceived'
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    # Find the Inbox of the user and bind to it.
    $Inboxid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$mailbox) 
    $Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWS,$inboxid)
    $Inboxview = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
    # Poll items from Inbox
    $fiResult = $Inbox.FindItems($Inboxview)
    ""
    write-output $mailbox "'s" $Inbox.DisplayName.tostring() "contains" $Inbox.TotalCount.ToString() "items"
    
    ""
    $fiResult|select -First 5| Sort DateTimeReceived|select Subject,DateTimeReceived
    #Return $fiResult|gm
    }
    if ($GetCalendar)
    {
    # Find the Inbox of the user and bind to it.
    $Calendarid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$mailbox) 
    $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWS,$Calendarid)
    #Lets grab 2 items like LYNC does
    $Calendarview = New-Object Microsoft.Exchange.WebServices.Data.ItemView(2)
    # Poll the Calendar
    $fiResult = $Calendar.FindItems($Calendarview)
    
    $fiResult|select -First 5| Sort DateTimeReceived|select Subject,DateTimeReceived
    ""
    $items = $mailbox+"'s "+$Calendar.DisplayName+" contains "+$Calendar.TotalCount+" items"
    $items
    ""
    }

    if ($GetDelegates -eq $true)
    {
        $EWS.GetDelegates($mailbox,$true)
    }

    if ($GetFreeBusy)
    {
        #Establish the Duration Window
        $StartTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00"))  
        $EndTime = $StartTime.AddDays(1) 
        $Duration = new-object Microsoft.Exchange.WebServices.Data.TimeWindow($StartTime,$EndTime)  
        #Setup grabbing the free busy
        $AvailabilityOptions = new-object Microsoft.Exchange.WebServices.Data.AvailabilityOptions  
        $AvailabilityOptions.RequestedFreeBusyView = [Microsoft.Exchange.WebServices.Data.FreeBusyViewType]::Detailed
        $AvailabilityOptions.MeetingDuration = "30"
        $AvailabilityOptions.MaximumSuggestionsPerDay = "1"
        #Get some FreeBusy 
        $Listtype = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
        $Listtype = $listtype.MakeGenericType("Microsoft.Exchange.WebServices.Data.AttendeeInfo" -as "Type")
        $Attendeesbatch = [Activator]::CreateInstance($listtype) 
        $Attendee = new-object Microsoft.Exchange.WebServices.Data.AttendeeInfo($Mailbox)  
        $Attendeesbatch.add($Attendee)  
        #Display FreeBusy
        $Availresponse = $EWS.GetUserAvailability($Attendeesbatch,$Duration,[Microsoft.Exchange.WebServices.Data.AvailabilityData]::FreeBusyAndSuggestions,$AvailabilityOptions)
        
    
        Write-Output "Result from Querying Free/Busy:"$Availresponse.AttendeesAvailability.Result
        ""
        foreach($Avail in $availresponse.AttendeesAvailability){  
            foreach($CEvent in $Avail.CalendarEvents){
               Write-Output "Subject:"$CEvent.Details.Subject
               Write-Output "Free/Busy Status:"$CEvent.FreeBusyStatus 
               Write-Output "Start:"$CEvent.StartTime
               Write-Output "End:"$CEvent.EndTime
                ""
            }  
        }
        Write-Verbose "A single meeting suggestion:"
        $Availresponse.Suggestions|select Date,Quality
    }
}

#Catch and handle all Exceptions
catch [Microsoft.Exchange.Webservices.Data.AccountIsLockedException]
{
    Write-Error "The account has been locked out. Please unlock your account and try again."
    Write-Error $_.Exception.InnerException
}
catch
{
    Write-Error $_.Exception.InnerException
}
