# 
# Add-MailboxDelegate.ps1

<#
.Synopsis
Properly adds a delegate to the mailbox without the need to create an Outlook profile

.Description
Using Outlook is the normal way to add a delegate to a mailbox.  Giving rights to folders and
send-on-behalf rights does not make a user a delegate, it merely gives those rights.  This script
will use the GetDelegate EWS command to properly create a delegate, and Exchange will add the user
to the publicDelegate property of the delegator, give appropriate rights, and create the delegate
properties on the root and inbox folders in the delegator mailbox.

The user running this script must have a mailbox and must have ApplicationImpersonation rights.  Alternatively
ignore the impersonation check if the user themselves are running the script for some reason.

More information:
You experience issues in Outlook when you try to configure
free/busy information or when you try to delegate information
https://support.microsoft.com/en-us/help/958443

.Example
$cred = Get-Credential
C:\PS>Add-Delegate 2016user1 -Credential:$cred -Delegate 2016user4

This will use the credentials from Get-Credential, and add the user 2016user4 as a delegate for 2016user1, using
  the delegate permissions Outlook uses by default.

.Parameter User
The User parameter specifies the target user's alias, email address, user principal name, or Mailbox object
  from Get-Mailbox.
  
.Parameter Delegate
The Delegate parameter specifies the user to be added as a delegate.  Note that permissions settings
  will be shared across all added delegates.

.Parameter ImportDelegateXml
This will import an exported ADUser object with the publicDelegates property pulled.  Note that permissions settings
  will be shared across all added delegates.
  e.g. Get-ADUser 2016user4 -Properties publicDelegates | Export-CliXml .\2016user4.xml

.Parameter CalendarFolderPermissionLevel
Default is Editor.  Options are None, Reviewer, Author, Editor

.Parameter TasksFolderPermissionLevel
Default is Editor.  Options are None, Reviewer, Author, Editor

.Parameter InboxFolderPermissionLevel
Default is None.  Options are None, Reviewer, Author, Editor

.Parameter ContactsFolderPermissionLevel
Default is None.  Options are None, Reviewer, Author, Editor

.Parameter NotesFolderPermissionLevel
Default is None.  Options are None, Reviewer, Author, Editor

.Parameter JournalFolderPermissionLevel
Default is None.  Options are None, Reviewer, Author, Editor

.Parameter ReceiveCopiesOfMeetingMessages
Default is true

.Parameter ViewPrivateItems
Default is false.  Sets whether delegates can view private items.

.Parameter DeliverMeetingRequests
Default is DelegatesOnly.  Options are DelegatesOnly, DelegatesAndMe, DelegatesAndSendInformationToMe, and NoForward

.Parameter Credential
The Credential parameter specifies the credentials to use for the impersonation account.

.Parameter PromptForCreds
The PromptForCreds switch will force a prompt for credentials.  If Credential and PromptForCreds are not set
  the script will use the currently logged-in user.

.Parameter UseDefaultCredentials
The UseDefaultCredentials parameter specifies whether to use the current user for EWS impersonation.  This
  defaults to $true.

.Parameter UseLocalHost
The UseLocalHost parameter specifies whether to use the local machine's hostname for the EWS endpoint.

.Parameter EwsUrl
The EwsUrl specifies the EWS URL to use.

.Parameter IgnoreSsl
The IgnoreSsl parameter specifies whether to ignore SSL validation errors.

.Parameter DomainController
The DomainController parameter specifies the fully qualified domain name (FQDN) of the domain controller that
  retrieves data from Active Directory.

.Parameter IgnoreImpersonationFailure
The IgnoreImpersonationFailure parameter specifies whether the script should continue when the RBAC role
  of ApplicationImpersonation is not found.  This occurs when using an account with full mailbox access but
  no impersonation rights.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [alias('Mailbox')]
    [string]$Identity,

    [Parameter(Mandatory=$false)]
    [PSCredential]$Credential,

    [Parameter(Mandatory=$false)]
    [switch]$UseDefaultCredentials = $false,

    [Parameter(Mandatory=$false)]
    [switch]$UseLocalHost,

    [Parameter(Mandatory=$false)]
    [string]$EwsUrl,

    [Parameter(Mandatory=$false)]
    [switch]$IgnoreSsl,

    [Parameter(Mandatory=$false)]
    [switch]$RemoveOnly,

    [Parameter(Mandatory=$false)]
    [switch]$Force,

    [Parameter(Mandatory=$false)]
    $DomainController,

    [Parameter(Mandatory=$false)]
    [switch]$IgnoreImpersonationFailure,
    
    [Parameter(Mandatory=$false)]
    [string]$Delegate,
    
    [Parameter(Mandatory=$false)]
    [string]$ImportDelegateXml,

    ## Params can't use custom enums unless they're pre-defined, which
    ## isn't possible in PS script like this without using the "using"
    ## statement against an external script that defines the enum.
    ## Without this, something more detailed like these permission levels
    ## are prone to typos, and we can't have that.
    
    ## So, interestingly, we can create the enum in the ArgumentCompleter and
    ## ValidateScript blocks, and they'll be valid for all subsequent parameters.
    
    ## The "None" value throws validation for a loop, so I have to manually
    ## check the value against the enum.  Just doing [FolderPermissionLevel]$_
    ## will not return true if None is passed

    [Parameter(Mandatory=$false)]
    [ArgumentCompleter({
        enum FolderPermissionLevel { None; Reviewer; Author; Editor }
        [FolderPermissionLevel].GetEnumValues()
    })]
    [ValidateScript({
        enum FolderPermissionLevel { None; Reviewer; Author; Editor }
        if ([enum]::GetValues([FolderPermissionLevel]) -contains $_) { $true }
    })]
    $CalendarFolderPermissionLevel = 'Editor',

    [Parameter(Mandatory=$false)]
    [ArgumentCompleter({
        [FolderPermissionLevel].GetEnumValues()
    })]
    [ValidateScript({
        if ([enum]::GetValues([FolderPermissionLevel]) -contains $_) { $true }
    })]
    $TasksFolderPermissionLevel = 'Editor',

    [Parameter(Mandatory=$false)]
    [ArgumentCompleter({
        [FolderPermissionLevel].GetEnumValues()
    })]
    [ValidateScript({
        if ([enum]::GetValues([FolderPermissionLevel]) -contains $_) { $true }
    })]
    $InboxFolderPermissionLevel = 'None',

    [Parameter(Mandatory=$false)]
    [ArgumentCompleter({
        [FolderPermissionLevel].GetEnumValues()
    })]
    [ValidateScript({
        if ([enum]::GetValues([FolderPermissionLevel]) -contains $_) { $true }
    })]
    $ContactsFolderPermissionLevel = 'None',

    [Parameter(Mandatory=$false)]
    [ArgumentCompleter({
        [FolderPermissionLevel].GetEnumValues()
    })]
    [ValidateScript({
        if ([enum]::GetValues([FolderPermissionLevel]) -contains $_) { $true }
    })]
    $NotesFolderPermissionLevel = 'None',

    [Parameter(Mandatory=$false)]
    [ArgumentCompleter({
        [FolderPermissionLevel].GetEnumValues()
    })]
    [ValidateScript({
        if ([enum]::GetValues([FolderPermissionLevel]) -contains $_) { $true }
    })]
    $JournalFolderPermissionLevel = 'None',

    [Parameter(Mandatory=$false)]
    [bool]$ReceiveCopiesOfMeetingMessages = $true,

    [Parameter(Mandatory=$false)]
    [bool]$ViewPrivateItems = $false,

    [Parameter(Mandatory=$false)]
    [ArgumentCompleter({
        enum DeliverMeetingRequestsOption { DelegatesOnly; DelegatesAndMe; DelegatesAndSendInformationToMe; NoForward }
        [DeliverMeetingRequestsOption].GetEnumValues()
    })]
    [ValidateScript({
        enum DeliverMeetingRequestsOption { DelegatesOnly; DelegatesAndMe; DelegatesAndSendInformationToMe; NoForward }
        if ([enum]::GetValues([FolderPermissionLevel]) -contains $_) { $true }
    })]
    $DeliverMeetingRequests = 'DelegatesOnly'
)

Process
{
    ## CalendarFolderPermissionLevel
    ## https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/calendarfolderpermissionlevel
    enum FolderPermissionLevel {
        None
        Reviewer
        Author
        Editor
    }

    ## DeliverMeetingRequests
    ## https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/delivermeetingrequests
    enum DeliverMeetingRequestsOption {
        DelegatesOnly
        DelegatesAndMe
        DelegatesAndSendInformationToMe
        NoForward
    }

    ###################
    ## Session setup ##
    ###################
    
    ## Configure options and credentials

    ## WhatIf and Verbose will automatically be added to things like Set-ADUser
    $verbose = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent
    $whatif = $PSCmdlet.MyInvocation.BoundParameters["WhatIf"].IsPresent

    if ($ImportDelegateXml -eq '' -and $delegate -eq '')
    {
        Write-Warning "ImportDelegateXml and / or Delegate must be passed.  Exiting."
        return
    }

    if ($WhatIf)
    {
        ## This is prepended any time something drastic would normally happen, i.e. removing things from AD
        ##
        $whatIfText = 'What if: '
    }
    else
    {
        $whatIfText = ''
    }
    
    ## Check that we have ADUser from the ActiveDirectory module, and we're in EMS, before spending time doing anything else
    
    if ((Get-Command Get-ADUser -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Count -eq 0)
    {
        do {
            Write-Warning "RSAT PowerShell module needs to be installed.  Add Windows Feature now?"
            Write-Warning "Ref: https://docs.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools"
            Write-Host "[Y] Yes  " -NoNewLine -ForegroundColor Yellow
            Write-Host "[N] No  " -NoNewLine
            $installRSAT = Read-Host -Prompt "(default is `"Y`")"
            $installRSAT = $installRSAT.ToLower().Trim()
        } until ( $installRSAT -eq 'y' -or $installRSAT -eq 'n' -or $installRSAT -eq '' )

        $error.Clear()
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        
        if ($error.Count -gt 0)
        {
            Write-Warning "ActiveDirectory module failed to load.  Error:"
            $error
            return
        }
    }
            
    if ((Get-Command Get-Mailbox -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Count -eq 0)
    {
        Write-Warning 'This script must be executed in the Exchange Management Shell.  Exiting.'
        return
    }

    if ($Credential -eq $null)
    {
        if ($UseDefaultCredentials -eq $true)
        {
            $Credential = @{UserName = [Environment]::UserName}
            Write-Verbose ('Using ' + [Environment]::UserDomainName + '\' + [Environment]::UserName + ' as service account. Use -Credentials to override.')
        }
        else
        {
            ## Wrapped in a try-catch in case the user hits escape or cancel
            try
            {
                $Credential = Get-Credential -ErrorAction SilentlyContinue

                if (@(Get-User $Credential.UserName -ErrorAction SilentlyContinue).Count -eq 0)
                {
                    Write-Warning ('Unable to find user ' + $Credential.UserName + '.  Exiting.')
                    return
                }
            }
            catch
            {
                Write-Warning 'User invalid.  Exiting.'
                return
            }
        }
    }

    ## Verify RBAC rights and set a single DC
    if ($DomainController -ne '' -and $DomainController -ne $null)
    {
        if ((Get-ADDomainController $DomainController -ErrorAction SilentlyContinue).Count -ne 1)
        {
            Write-Error "Unable to find Domain Controller $DomainController.  Exiting."
        }
    }
    else
    {
        $DomainController = $null
    }
    
    Write-Verbose "Using '$DomainController' for RBAC rights lookup.  ViewEntireForest in ADServerSettings is set to $((Get-ADServerSettings).ViewEntireForest)"

    ## DomainController parameter may be RBAC'd out, but if it's not the assumption will be we have rights to it for all relevant commands
    if ((Get-Command Get-ManagementRoleAssignment).Parameters.ContainsKey('DomainController'))
    {
        Write-Verbose "Checking ApplicationImpersonation RBAC role for $($Credential.UserName)"
        $rbac = Get-ManagementRoleAssignment -RoleAssignee ($Credential.UserName) -Role ApplicationImpersonation -DomainController $DomainController #-ErrorAction SilentlyContinue
        
        ## Set DomainController to whatever gave us the RBAC rights, so we're at least consistent
        if ($rbac.OriginatingServer -ne $DomainController)
        {
            Write-Verbose "Setting DomainController to '$($rbac.OriginatingServer)'"
            $DomainController = $rbac.OriginatingServer
        }
    }
    else
    {
        ## No access to -DomainController parameter
        Write-Error "No access to DomainController parameter, but this is required to properly re-set delegates."
        return
    }
    
    if ($rbac -eq $null -or $rbac.Count -eq 0)
    {
        Write-Warning ('ApplicationImpersonation rights not found for ' + $Credential.UserName + '.')
        Write-Warning 'Note that if impersonation was recently added, it may take some time to replicate.  Use the -DomainController option to specify a DC for lookup.'
        Write-Warning ""
        
        ## Article about configuring impersonation
        Write-Warning 'For more information on configuring impersonation in Exchange, go to https://msdn.microsoft.com/en-us/library/office/bb204095(v=exchg.140).aspx'
        Write-Warning "Usage example -- New-ManagementRoleAssignment –Name:impersonationAssignmentName –Role:ApplicationImpersonation –User:serviceAccount"
        Write-Warning ""

        ## Article about management role assignments
        Write-Warning 'For more information on the ApplicationImpersonation role, go to https://technet.microsoft.com/en-us/library/dd776119(v=exchg.150).aspx'

        if ($IgnoreImpersonationFailure)
        {
            Write-Warning '$IgnoreImpersonationFailure is set to $true, continuing.'
        }
        else
        {
            Write-Warning '$IgnoreImpersonationFailure is set to $false, exiting.'
            return
        }
    }
    
    Write-Verbose "Using $($Credential.UserName) as service account."

    #########################################
    ##  Getting Mailbox and Delegate Info  ##
    ##  - Preparing XML for EWS calls      ##
    #########################################

    Write-Host "Gathering mailbox info." -NoNewLine

    ## Get Mailbox Info
    $Mailbox = Get-Mailbox $Identity -DomainController $DomainController
    Write-Host "." -NoNewLine
    $MailboxADUser = Get-ADUser $Mailbox.Alias -Properties publicDelegates -Server $DomainController
    Write-Host "."
    $MailboxRecipient = Get-Recipient $Mailbox.Alias -DomainController $DomainController
    $EmailAddress = $Mailbox.PrimarySmtpAddress
    
    Write-Verbose "Current publicDelegates property:"
    if ( $MailboxADUser.publicDelegates.Count -gt 0)
    {
        $MailboxADUser | Select -ExpandProperty publicDelegates
        $filename = "~\Desktop\$($Mailbox.Alias)_Delegator_$([DateTime]::Now.ToString("yyyyMMdd-HHmmss")).xml"
        Write-Host -ForegroundColor Yellow "Saving current delegates for $($Mailbox.Name) to $filename"
        $MailboxADUser | Export-CliXml $filename
    }
    else
    {
        Write-Verbose "Empty publicDelegates property"
    }
    
    ##############################
    ##  Preparing EWS settings  ##
    ##############################

    ## Get the target EWS url
    ## For POX, I'm not doing AutoD.  Either specify or it's local :P
    if ($EwsUrl -ne '')
    {
        $uri = [System.Uri]$EwsUrl
    }
    else
    {
        $uri = [System.Uri]('https://' + $Env:ComputerName + '/EWS/Exchange.asmx')
        # If we're local, go ahead and ignore SSL
        $IgnoreSsl = $true
    }

    # Ignore cert errors
    if (-not ([System.Management.Automation.PSTypeName]'CertificateUtils').Type)
    {
    add-type @"
        using System.Net;
        using System.Net.Security;
        using System.Security.Cryptography.X509Certificates;
        public static class CertificateUtils {
            public static bool TrustAllCertsCallback(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) {
                return true;
            }
           
            public static void TrustAllCerts() {
                ServicePointManager.ServerCertificateValidationCallback = CertificateUtils.TrustAllCertsCallback;
            }
        }
"@
    }

    ## Ignore SSL warnings due to self-signed certs 
    if ($IgnoreSsl)
    {
        # This is the normal way to do it, but it has runspace issues
        #[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
        [CertificateUtils]::TrustAllCerts()
    }
    
    ## Configure for TLS 1.2
    #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    ###########################################
    ##  EWS calls using POX / Plain Old XML  ##
    ###########################################

    ## Prepare the AddDelegate XML
    
    $addXml = ""
    if ($Delegate -ne '' -and $Delegate -ne $null)
    {
        Write-Verbose "Adding delegate $Delegate"
    
        $newDelegate = Get-Recipient $Delegate -DomainController $DomainController -ErrorAction SilentlyContinue
        if ($newDelegate.Count -ne 1)
        {
            Write-Warning "Recipient $Delegate not found.  Exiting."
            return
        }

$addXml += @"
          <t:DelegateUser>
            <t:UserId>
              <t:PrimarySmtpAddress>$($newDelegate.primarySmtpAddress)</t:PrimarySmtpAddress>
            </t:UserId>
            <t:DelegatePermissions>
              <t:CalendarFolderPermissionLevel>$CalendarFolderPermissionLevel</t:CalendarFolderPermissionLevel>
              <t:TasksFolderPermissionLevel>$TasksFolderPermissionLevel</t:TasksFolderPermissionLevel>
              <t:InboxFolderPermissionLevel>$InboxFolderPermissionLevel</t:InboxFolderPermissionLevel>
              <t:ContactsFolderPermissionLevel>$ContactsFolderPermissionLevel</t:ContactsFolderPermissionLevel>
              <t:NotesFolderPermissionLevel>$NotesFolderPermissionLevel</t:NotesFolderPermissionLevel>
              <t:JournalFolderPermissionLevel>$JournalFolderPermissionLevel</t:JournalFolderPermissionLevel>
            </t:DelegatePermissions>
            <t:ReceiveCopiesOfMeetingMessages>$($ReceiveCopiesOfMeetingMessages.ToString().ToLower())</t:ReceiveCopiesOfMeetingMessages>
            <t:ViewPrivateItems>$($ViewPrivateItems.ToString().ToLower())</t:ViewPrivateItems>
          </t:DelegateUser>
"@
    }

    if ($ImportDelegateXml -ne $null -and $ImportDelegateXml -ne '' -and (Test-Path $ImportDelegateXml) -eq $true)
    {
        Write-Verbose "Adding delegates from $ImportDelegateXml"
        
        $publicDelegates = (Import-CliXml $ImportDelegateXml).publicDelegates
        if ($publicDelegates.Count -eq 0)
        {
            Write-Warning "No delegates found in $ImportDelegateXml.  Exiting."
            return
        }

        foreach ($pubDel in $publicDelegates)
        {
            $newDelegate = Get-Recipient $pubDel -DomainController $DomainController -ErrorAction SilentlyContinue
            if ($newDelegate.Count -ne 1)
            {
                continue
            }
            
$addXml += @"
          <t:DelegateUser>
            <t:UserId>
              <t:PrimarySmtpAddress>$($newDelegate.primarySmtpAddress)</t:PrimarySmtpAddress>
            </t:UserId>
            <t:DelegatePermissions>
              <t:CalendarFolderPermissionLevel>$CalendarFolderPermissionLevel</t:CalendarFolderPermissionLevel>
              <t:TasksFolderPermissionLevel>$TasksFolderPermissionLevel</t:TasksFolderPermissionLevel>
              <t:InboxFolderPermissionLevel>$InboxFolderPermissionLevel</t:InboxFolderPermissionLevel>
              <t:ContactsFolderPermissionLevel>$ContactsFolderPermissionLevel</t:ContactsFolderPermissionLevel>
              <t:NotesFolderPermissionLevel>$NotesFolderPermissionLevel</t:NotesFolderPermissionLevel>
              <t:JournalFolderPermissionLevel>$JournalFolderPermissionLevel</t:JournalFolderPermissionLevel>
            </t:DelegatePermissions>
            <t:ReceiveCopiesOfMeetingMessages>$($ReceiveCopiesOfMeetingMessages.ToString().ToLower())</t:ReceiveCopiesOfMeetingMessages>
            <t:ViewPrivateItems>$($ViewPrivateItems.ToString().ToLower())</t:ViewPrivateItems>
          </t:DelegateUser>
"@
        }
    }

    $addDelegateXml = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
      <t:RequestServerVersion Version="Exchange2016" />
      <t:ExchangeImpersonation>
        <t:ConnectingSID>
          <t:PrimarySmtpAddress>$EmailAddress</t:PrimarySmtpAddress>
        </t:ConnectingSID>
      </t:ExchangeImpersonation>
    </soap:Header>
    <soap:Body>
      <m:AddDelegate>
        <m:Mailbox>
          <t:EmailAddress>$EmailAddress</t:EmailAddress>
        </m:Mailbox>
        <m:DelegateUsers>
$addXml
        </m:DelegateUsers>
        <m:DeliverMeetingRequests>$DeliverMeetingRequests</m:DeliverMeetingRequests>
      </m:AddDelegate>
    </soap:Body>
</soap:Envelope>
"@

    Write-Verbose ($WhatIfText + "Sending AddDelegate XML to $uri`:`r`n$addDelegateXml")

    $retry = 3
    
    do {
        $result = $null
        $error.Clear()

        try {
            if ($UseDefaultCredentials -and !$WhatIf)
            {
                Write-Host -ForegroundColor Yellow "...using default credentials"
                $result = Invoke-WebRequest -Uri $uri -Method Post -Body $addDelegateXml -ContentType "text/xml" -Headers @{'X-AnchorMailbox' = $EmailAddress} -UseDefaultCredentials
            }
            elseif (!$WhatIf)
            {
                Write-Host -ForegroundColor Yellow "...using specified credentials"
                $result = Invoke-WebRequest -Uri $uri -Method Post -Body $addDelegateXml -ContentType "text/xml" -Headers @{'X-AnchorMailbox' = $EmailAddress} -Credential:$Credential
            }

            Write-Host "Result: $($result.StatusCode) $($result.StatusDescription)"

            if ($verbose)
            {
                ## This pretty-prints the headers
                $result.Headers.GetEnumerator() | % {
                    Write-Host "$($_.Key): $($_.Value)"
                }
                Write-Host ""
                ## This pretty-prints the XML
                ([xml]$result.Content).Save([Console]::Out)
                Write-Host ""
            }
        }
        catch
        {
            ## We have to pull the response body a bit differently in this case
            $result = $PSItem.Exception.Response
            Write-Host "Result: $([int]$result.StatusCode) $($result.StatusDescription)"
            
            ## This pretty-prints the headers
            $result.Headers | % { Write-Host "$_`: $($result.GetResponseHeader($_))" };
            Write-Host ""
            
            $stream = $result.GetResponseStream()
            $stream.Position = 0
            $reader = [System.IO.StreamReader]::new($stream)
            $content = $reader.ReadToEnd()
            
            if ($content.Contains("<?xml"))
            {
                ([xml]$content).Save([Console]::Out)
            }
            else
            {
                $content
            }
            Write-Host ""
            Write-Host ""
            
        }
        finally
        {
            $retry--
        }
        
        if ($content -match "ErrorDelegateAlreadyExists")
        {
            Write-Verbose "At least one delegate already exists, not retrying."
            retry = 0;
        }


        if ($result.StatusCode -ne 200)
        {
            Write-Warning "EWS call failed."
            if ($retry -gt 0)
            {
                Write-Warning "Retrying..."
            }
            else
            {
                Write-Warning "Exiting."
            }
        }
    }
    while ( $retry -ne 0 -and $result.StatusCode -ne 200)
    
    ## We should have editor rights on calendar and SOB rights
    sleep 2
    
    Write-Host "Checking publicDelegates attribute for $($Mailbox.Name)"
    $MailboxADUser = Get-ADUser $Mailbox.Alias -Properties publicDelegates -Server $DomainController
    $MailboxADUser | Select -ExpandProperty publicDelegates
    Write-Host ""
    
    Write-Host "Checking calendar rights"
    Get-MailboxFolderPermission "$($Mailbox.Alias):\calendar" -DomainController $DomainController
    Write-Host ""
    
    Write-Host "Checking Send-On-Behalf rights"
    Get-Mailbox $Mailbox -DomainController $DomainController | Select GrantSendOnBehalfTo
    Write-Host ""

    Write-Host "All done with $($Mailbox.Name)"
}