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

.Parameter FreeBusy
    This runs the test to get FreeBusy info and a suggestion.

.Parameter Headers
    This can be combined with any of the parameters to show the EWS headers

.Parameter GetDelegates
    This will get a Delegate list assuming you have permission to the mailbox.

.Parameter Debug
    Enables advanced logging including the SOAP traces from EWS.

######>

Param 
(
    [string]$Mailbox = $null,
    $Credential = $null,
    [URI]$EWSUrl = $null,
    [string]$AuthenticationType = 'default',
    [switch]$Debug = $false,
    [switch]$GetDelegates,
    [switch]$Headers = $false,
    [switch]$FreeBusy = $false,
    [switch]$GetFolder = $false
)

Import-Module -Name ".\Microsoft.Exchange.WebServices.dll"

try
{
    $EWS = new-object Microsoft.Exchange.WebServices.Data.ExchangeService
    $EWS.UserAgent = 'Basic EWS Test Client.ps1/v0.1'
    if ($Debug -eq $true)
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
        Write-Host -ForegroundColor Yellow "No EWS location Provided, attempting AutoDiscover"
        $EWS.AutoDiscoverURL($mailbox);
        write-host -ForegroundColor Green "Found EWS at:" $EWS.Url.AbsoluteUri
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
                Write-Host -ForegroundColor Yellow "Using current Windows credentials."
                $EWS.UseDefaultCredentials = $true
            }
            NTLM
            {
                Write-Host -ForegroundColor Yellow "Creating an NTLM credential package."
                $CredentialCache.Add(($EWS.Url.AbsoluteUri),"NTLM",(get-credential))
                $EWS.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($CredentialCache)
            }
            Negotiate
            {
                Write-Host -ForegroundColor Yellow "Creating a Negotiate credential package"
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


    # Find the Inbox of the user and bind to it.
    $Inboxid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$mailbox) 
    $Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWS,$inboxid)
    $Inboxview = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)

    # If they exist show the user where we got the info from.
    if ($EWS.HttpResponseHeaders.ContainsKey('X-FEServer') -eq $true)
    {
        Write-Host -ForegroundColor Yellow "FrontEnd Server:" $EWS.HttpResponseHeaders.Item('X-FEServer')
    }
    if ($EWS.HttpResponseHeaders.ContainsKey('X-BEServer') -eq $true)
    {
        Write-Host -ForegroundColor Yellow "Backend Server:" $EWS.HttpResponseHeaders.Item('X-BEServer')
    }
    if ($Headers)
    {
        $EWS.HttpHeaders
        $EWS.HttpResponseHeaders
    }
    if ($GetFolder)
    {
    # Poll items from Inbox
    $fiResult = $Inbox.FindItems($Inboxview)
    
    $fiResult|select -First 5| Sort DateTimeReceived|select Sender,Subject,DateTimeReceived
    ""
    $items = $mailbox+"'s "+$Inbox.DisplayName+" contains "+$Inbox.TotalCount+" items"
    $items
    ""
    }

    if ($GetDelegates -eq $true)
    {
        $EWS.GetDelegates($mailbox,$true)
    }

    if ($FreeBusy)
    {
        $StartTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00"))  
        $EndTime = $StartTime.AddDays(1) 
        $Duration = new-object Microsoft.Exchange.WebServices.Data.TimeWindow($StartTime,$EndTime)  
        $AvailabilityOptions = new-object Microsoft.Exchange.WebServices.Data.AvailabilityOptions  
        $AvailabilityOptions.RequestedFreeBusyView = [Microsoft.Exchange.WebServices.Data.FreeBusyViewType]::DetailedMerged
        $AvailabilityOptions.MeetingDuration = "30"
        $AvailabilityOptions.MaximumSuggestionsPerDay = "1" 
        $Listtype = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
        $Listtype = $listtype.MakeGenericType("Microsoft.Exchange.WebServices.Data.AttendeeInfo" -as "Type")
        $Attendeesbatch = [Activator]::CreateInstance($listtype) 
        $Attendee = new-object Microsoft.Exchange.WebServices.Data.AttendeeInfo($Mailbox)  
        $Attendeesbatch.add($Attendee)  
      
        $Availresponse = $EWS.GetUserAvailability($Attendeesbatch,$Duration,[Microsoft.Exchange.WebServices.Data.AvailabilityData]::FreeBusyAndSuggestions,$AvailabilityOptions)
    
        write-Host "Result from Querying Free/Busy:"$Availresponse.AttendeesAvailability.Result
        ""
        foreach($Avail in $availresponse.AttendeesAvailability){  
            foreach($CEvent in $Avail.CalendarEvents){
               Write-Host "Subject:"$CEvent.Details.Subject
               Write-Host "Free/Busy Status:"$CEvent.FreeBusyStatus 
               Write-Host "Start:"$CEvent.StartTime
               Write-Host "End:"$CEvent.EndTime
                ""
            }  
        }
        Write-Host -ForegroundColor Yellow "A single meeting suggestion:"
        $Availresponse.Suggestions|select Date,Quality
    }
}

#Catch and handle all Exceptions
catch [Microsoft.Exchange.Webservices.Data.AccountIsLockedException]
{
    Write-Error "The account has been locked out. Please unlock your account and try again."
    write-error $_.Exception.InnerException
}
catch
{
    write-error $_.Exception.InnerException
}
