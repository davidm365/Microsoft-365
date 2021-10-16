# 
# Get-Delegate.ps1

<#
.Synopsis
Gets delegate info for a user

.Description
This script attempts to get a list of delegates for a given user

.Example
$cred = Get-Credential
C:\PS>Get-Delegate -Identity 2016user1 -Credential:$cred

This will use the credentials from Get-Credential, and try to pull the current list of delegates from
  the publicDelegates attribute on the user and sending a GetDelegate command via EWS.

.Parameter Identity
The Identity parameter specifies the target user's alias, email address, user principal name, or Mailbox object
  from Get-Mailbox.

.Parameter Credential
The Credential parameter specifies the credentials to use for the impersonation account.

.Parameter UseDefaultCredentials
The UseDefaultCredentials parameter specifies whether to use the current user for EWS impersonation.  This
  defaults to $true.

.Parameter UseLocalHost
The UseLocalHost parameter specifies whether to use the local machine's hostname for the EWS endpoint.

.Parameter EwsUrl
The EwsUrl specifies the EWS URL to use.

.Parameter TraceEnabled
The TraceEnabled parameter enables EWS tracing.

.Parameter IgnoreSsl
The IgnoreSsl parameter specifies whether to ignore SSL validation errors.

.Parameter Debug
Enables XML output for requests / responses

Get-Delegate -Identity 2019user1@exchlab.com -UseDefaultCredentials
    -EwsUrl
    -Credential:$cred
    -User
    -UseLocalHost
    -IgnoreSsl
#>

[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [string]$Identity,

    [Parameter(Mandatory=$false)]
    [PSCredential]$Credential,

    [Parameter(Mandatory=$false)]
    [switch]$UseDefaultCredentials = $true,

    [Parameter(Mandatory=$false)]
    [switch]$UseLocalHost,

    [Parameter(Mandatory=$false)]
    [string]$EwsUrl,

    [Parameter(Mandatory=$false)]
    [switch]$TraceEnabled,

    [Parameter(Mandatory=$false)]
    [switch]$IgnoreSsl,

    [Parameter(Mandatory=$false)]
    [switch]$IgnoreImpersonationFailure
)

Process
{
    ###################
    ## Session setup ##
    ###################
    
    ## Configure options and credentials
    if (![string]::IsNullOrEmpty($Credential)) { $UseDefaultCredentials = $false }

    ## WhatIf and Verbose will automatically be added to things like Set-ADUser
    $debug = $PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent
    $verbose = $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent
    
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
        Write-Warning "This script must be executed in the Exchange Management Shell.  Exiting."
        return
    }

    if ($Credential -eq $null)
    {
        if ($UseDefaultCredentials -eq $true)
        {
            $Credential = New-Object PSCredential("$([Environment]::UserDomainName)\$([Environment]::UserName)", ("password" | ConvertTo-SecureString -asPlainText -Force))
            Write-Verbose "Using '$([Environment]::UserDomainName)\$([Environment]::UserName)' as service account. Use -Credentials to override."
        }
        else
        {
            ## Wrapped in a try-catch in case the user hits escape or cancel
            try
            {
                $Credential = Get-Credential -ErrorAction SilentlyContinue

                if (@(Get-User $Credential.UserName -ErrorAction SilentlyContinue).Count -eq 0)
                {
                    Write-Warning "Unable to find user '$($Credential.UserName)'.  Exiting."
                    return
                }
            }
            catch
            {
                Write-Warning "User invalid.  Exiting."
                return
            }
        }
    }

    ## Verify rights
    Write-Verbose "Checking rights for $($Credential.UserName)"
    
    ## Check if we're SELF
    $self = (Get-User $Identity).SamAccountName -eq (Get-User $Credential.UserName).SamAccountName

    if ($self)
    {
        Write-Verbose "  Should have access (SELF)"
        $UseImpersonation = $false
    }
    else
    {
        $rbac = Get-ManagementRoleAssignment -RoleAssignee ($Credential.UserName) -Role ApplicationImpersonation -ErrorAction SilentlyContinue
        
        if ($rbac -eq $null -or $rbac.Count -eq 0)
        {
            #Before warning, see if we have FullAccess rights
            $FARights = Get-MailboxPermission $Identity -User $Credential.UserName | ? { $_.AccessRights -match "FullAccess" }

            if ($FARights -eq $null -or ($FARights | ? { $_.Deny -eq $true }).Count -gt 0)
            {
                Write-Warning "No FullAccess or ApplicationImpersonation rights found for $($Credential.UserName)."
                Write-Warning "Note that if rights were recently added, it may take some time to replicate."
                Write-Warning ""
                
                ## Article about configuring impersonation
                Write-Warning "For more information on configuring impersonation in Exchange, go to https://msdn.microsoft.com/en-us/library/office/bb204095(v=exchg.140).aspx"
                Write-Warning "Usage example -- New-ManagementRoleAssignment -Name:impersonationAssignmentName -Role:ApplicationImpersonation -User:serviceAccount"
                Write-Warning ""

                ## Article about management role assignments
                Write-Warning "For more information on the ApplicationImpersonation role, go to https://technet.microsoft.com/en-us/library/dd776119(v=exchg.150).aspx"
                
                return
            }
            
            $UseImpersonation = $false
            Write-Verbose "  Should have access (FullAccess)"
        }
        else
        {
            $UseImpersonation = $true
            Write-Verbose "  Should have access (ApplicationImpersonation)"
        }
    }
    
    Write-Verbose "Using $($Credential.UserName) as service account."

    #########################################
    ##  Getting Mailbox and Delegate Info  ##
    ##  - Preparing XML for EWS calls      ##
    #########################################

    Write-Host "Gathering mailbox info." -NoNewLine

    ## Get Mailbox Info
    $Mailbox = Get-Mailbox $Identity
    Write-Host "." -NoNewLine
    $MailboxADUser = Get-ADUser $Mailbox.SamAccountName -Properties publicDelegates
    Write-Host "."
    $EmailAddress = $Mailbox.PrimarySmtpAddress.Address
    
    ##############################
    ##  Preparing EWS settings  ##
    ##############################

    ## Get delegates via EWS
    Write-Host "`r`nGetting delegates from EWS..."
    
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
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    ###########################################
    ##  EWS calls using POX / Plain Old XML  ##
    ###########################################

    if ($UseImpersonation)
    {
        $impersonationXml = @"

        <t:ExchangeImpersonation>
            <t:ConnectingSID>
                <t:PrimarySmtpAddress>$EmailAddress</t:PrimarySmtpAddress>
            </t:ConnectingSID>
        </t:ExchangeImpersonation>
"@
    }
    else
    {
        $impersonationXml = $null
    }
    
    $getDelegateXml = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
        <t:RequestServerVersion Version="Exchange2013_SP1" />$impersonationXml
    </soap:Header>
    <soap:Body>
        <m:GetDelegate IncludePermissions="true">
            <m:Mailbox>
                <t:EmailAddress>$EmailAddress</t:EmailAddress>
            </m:Mailbox>
        </m:GetDelegate>
    </soap:Body>
</soap:Envelope>
"@

    if ($debug)
    {
        Write-Host "`r`nGetDelegate EWS request:"
        Write-Host $getDelegateXml
    }

    Write-Verbose "Sending XML to $uri"

    $retry = 3

    do {
        $result = $null
        $content = $null
        $error.Clear()

        try {
            if ($UseDefaultCredentials)
            {
                Write-Verbose "...using default credentials"
                $result = Invoke-WebRequest -Uri $uri -Method Post -Body $getDelegateXml -ContentType "text/xml" -Headers @{'X-AnchorMailbox' = $EmailAddress} -UseDefaultCredentials
            }
            else
            {
                Write-Verbose "...using specified credentials"
                $PSCred = [PSCredential]$Credential
                $result = Invoke-WebRequest -Uri $uri -Method Post -Body $getDelegateXml -Headers @{'X-AnchorMailbox' = $EmailAddress} -Credential:$PSCred -ContentType "text/xml"
            }

            Write-Verbose "Result: $($result.StatusCode) $($result.StatusDescription)"
            
            ## Makes it easier to do the retry check since the catch won't have $result.Content
            $content = $result.Content

            if ($debug)
            {
                Write-Host "`r`nGetDelegate EWS response:"
                
                ## This pretty-prints the headers
                $result.Headers.GetEnumerator() | % {
                    Write-Host "$($_.Key): $($_.Value)"
                }
                Write-Host ""
                
                ## This pretty-prints the XML
                ([xml]$content).Save([Console]::Out)
                Write-Host ""
                Write-Host ""
            }
            
            $resultXml = [xml]$content
            
            $delegateUsers = $resultXml.Envelope.Body.GetDelegateResponse.ResponseMessages.DelegateUserResponseMessageType.DelegateUser
            
            if ($delegateUsers -ne $null)
            {
                Write-Host "`r`nDelegates listed in EWS GetDelegate:"
                
                $allDelegates = @()
                
                foreach ($delegate in $resultXml.Envelope.Body.GetDelegateResponse.ResponseMessages.DelegateUserResponseMessageType.DelegateUser)
                {
                    $del = '' | SELECT Delegate, Calendar, Tasks, Inbox, Contacts, Notes
                    $del.Delegate = $delegate.UserId.DisplayName
                    $del.Calendar = $delegate.DelegatePermissions.CalendarFolderPermissionLevel
                    $del.Tasks = $delegate.DelegatePermissions.TasksFolderPermissionLevel
                    $del.Inbox = $delegate.DelegatePermissions.InboxFolderPermissionLevel
                    $del.Contacts = $delegate.DelegatePermissions.ContactsFolderPermissionLevel
                    $del.Notes = $delegate.DelegatePermissions.NotesFolderPermissionLevel
                    $allDelegates += $del
                }
                $allDelegates | FT | Out-String | Write-Host
            }
            else
            {
                Write-Host "No delegates listed in XML."
                if (!$debug)
                {
                    Write-Host "Use -Debug to output the raw XML from the EWS response."
                }
            }
        }
        catch
        {
            ## We have to pull the response body a bit differently in this case
            $result = $PSItem.Exception.Response
            Write-Host "Result: $([int]$result.StatusCode) $($result.StatusDescription)"
            
            if ([int]$result.StatusCode -eq 0)
            {
                $result.StatusCode
                $result.StatusDescription
                $PSItem.Exception
                return
            }
            
            ## This pretty-prints the headers
            $result.Headers | % { Write-Host "$_`: $($result.GetResponseHeader($_))" };
            Write-Host ""
            
            $stream = $result.GetResponseStream()
            $stream.Position = 0
            $reader = New-Object System.IO.StreamReader -ArgumentList $stream
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

        if ($result.StatusCode -ne 200 -or $error.Count -gt 0)
        {
            if ($content -match "ErrorImpersonateUserDenied")
            {
                ## Known permanent failure
                Write-Warning "Permanent failure.  Impersonation rights not found, but we have the RBAC right.  Likely using an admin user with explicit deny.  Exiting."
                return
            }

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
    while ( $retry -gt 0 -and $result.StatusCode -ne 200 -and $error.count -ne 0 )

    Write-Host "`r`nChecking publicDelegates attribute for $($Mailbox.Name)..."
    $MailboxADUser = Get-ADUser $Mailbox.SamAccountName -Properties publicDelegates | Select -ExpandProperty publicDelegates
    
    if ($MailboxADUser.Count -gt 0)
    {
        $MailboxADUser | Out-String | Write-Host
    }
    else
    {
        Write-Host "None (Empty property in AD)"
    }
    
    Write-Host "`r`nChecking calendar rights..."
    Get-MailboxFolderPermission "$($Mailbox.Alias):\calendar"
    
    Write-Host "`r`nChecking Send-On-Behalf rights..."
    $grantSOB = (Get-Mailbox $Mailbox).GrantSendOnBehalfTo
    if ($grantSOB.Count -gt 0)
    {
        $grantSOB | FT DistinguishedName | Out-String | Write-Host
    }
    else
    {
        Write-Host "No Send-On-Behalf rights assigned."
    }

    Write-Host "`r`nAll done with $($Mailbox.Name)"
}