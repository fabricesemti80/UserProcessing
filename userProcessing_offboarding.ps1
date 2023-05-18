<#
.SYNOPSIS
    Short description
.DESCRIPTION
    Long description
.EXAMPLE
    Example of how to use this cmdlet
.EXAMPLE
    Another example of how to use this cmdlet
#>
[cmdletbinding(SupportsShouldProcess = $True)]
[OutputType([int])]
param(
    [Parameter(Mandatory = $false)][switch] $test
)
BEGIN {
    <# * THE FIRST PART OF THE SCRIPTBLOCK IS TO PREPARE THE SCRIPT ENVIRONMENT - IMPORTING MODULES, DEFINING THE CURRENT LOCATION, ETC. - AND DEFINING THE VARIOUS LOG FILES USED TROUGHOUT THE SCRIPT#>
    #region ENVIRONMENT
    $Error.clear()
    $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
    Import-Module PSWriteColor -Verbose:$false
    Import-Module ActiveDirectory -Verbose:$false
    Import-Module ExchangeOnlineManagement -Verbose:$false
    Import-Module DFSN -Verbose:$false
    Import-Module RemoteDesktop -Verbose:$false
    $CurrentPath = Split-Path -Parent $PSCommandPath
    $sDate = Get-Date -Format yyyyMMdd
    Import-Module $CurrentPath\userProcessing.psm1 -Force -DisableNameChecking -WarningAction SilentlyContinue
    #endregion
    #region LOGS
    if (-not (Test-Path "$CurrentPath\out-off\$sDate")) {
        [void](New-Item -ItemType Directory -Path "$CurrentPath\out-off" -Name $sDate)
    }
    if (-not (Test-Path "$CurrentPath\logs\$sDate")) {
        [void](New-Item -ItemType Directory -Path "$CurrentPath\logs" -Name $sDate)
    }
    $doneFolder = "$CurrentPath\in-off\done" 
    if (-not (Test-Path $doneFolder )) {
        [void](New-Item -ItemType Directory -Path "$CurrentPath\in-off" -Name 'done')
    }        
    # logs
    $logTime = (Get-Date -Format yyy-MM-dd-HH-mm)
    $transcriptFile = $CurrentPath + '\logs\' + $sDate + '\' + $logTime + '_offboarding_transcript.log'
    $transcriptFile = $CurrentPath + '\logs\' + $sDate + '\' + $logTime + '_offboarding_transcript.log'
    $backupFile = $CurrentPath + '\out-off\' + $sDate + '\' + $logTime + '_offboarding_backup.csv'
    $errorFile = $CurrentPath + '\logs\' + $sDate + '\' + $logTime + '_offboarding_error.log'
    $reportFile = $CurrentPath + '\out-off\' + $sDate + '\' + $logTime + '_offboarding_report.csv'
    $actionLog = $CurrentPath + '\logs\' + $sDate + '\' + $logTime + '_offboarding_action.log'
    Start-Transcript -Path $transcriptFile -Force
    #endregion
}
PROCESS {
    <#  * THE PROCESS BLOCK CONTAINS THE MAJORITY OF THE SCRIPT'S WORK CODES. THIS STARTS WITH LISTING THE ACCEPTABLE DOMAINS (FROM THE CONFIGURATION FILE)#>
    $domainList = (Import-PowerShellDataFile $CurrentPath\userProcessing.psd1).DomainList
    <# * FOR EACH DOMAIN WE WILL IMPORT THE SPECIFIC VARIABLES - IE. EXCHANGE SERVER NAME OR ADMIN ACCOUNT NAME FOR AZURE ACTIVE DIRECTORY - AND ALSO IMPORT THE VARIABLES SHARED BETWEEN THE DOMAINS.  THE OTHER VERY IMPORTANT ACTION IN THIS BLOCK IS CACHING THE ACTIVE DIRECTORY USER BASE. THIS ADS A FEW MINUTE DELAY AT THE BEGINING OF THE SCRIPT, BUT SAVES TIME LATER ON AND ALLOWS THE SCRIPT'S FUNCTIONS TO USE THIS, ELIMINIATING THE NEED TO PASS ON CREDENTIAL PARAMETER TO QUERY ACTIVE DIRETORY #> 
    foreach ($domain in $domainList) {
        $summaryReport = @()
        $domain = $domain.domain
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer] [i]  Working on domain  [$domain] " -C Black -B White -LinesBefore 1
        $config = $((Import-PowerShellDataFile $CurrentPath\userProcessing.psd1).$domain)
        $config += $((Import-PowerShellDataFile $CurrentPath\userProcessing.psd1).common)
        $AD_Credential = New-Credential -outFile $($CurrentPath + '\creds\' + $config.DomainNetBios + '_AD_Cred.xml' ) -userName $config.AD_Admin
        $AAD_Credential = New-Credential -outFile $($CurrentPath + '\creds\' + $config.DomainNetBios + '_AAD_Cred.xml' ) -userName $config.AAD_Admin
        $Exch_Credential = New-Credential -outFile $($CurrentPath + '\creds\' + $config.DomainNetBios + '_EXC_Cred.xml' ) -userName $config.Exchange_Admin
        $DC = (Get-ADForest $config.SystemDomain -Credential $AD_Credential | Select-Object -ExpandProperty RootDomain | Get-ADDomain | Select-Object -Property PDCEmulator).PDCEmulator   
        $csvImport = $null
        #* CLEANUP EARLIER RUNS
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Moving previously processed files to  ', "[$doneFolder]" -Color White, Green, White, Green
        Cleanup-EarlierRuns -inputFolder $($config.OffBoardingInputFolder) -doneFolder $doneFolder
        if ($test.IsPresent) {
            $file = Get-FileName -initialDirectory $($config.OffBoardingInputFolder)
            if ($file) {
                $csvImport = Import-Csv $file | Where-Object { $_.Domain -match $domain } 
            }            
        }
        else {
            $csvImport = Import-WorkFiles -inputFolder $($config.OffBoardingInputFolder) | Where-Object { $_.Domain -match $domain }
            # * STORE FILE LIST
            $fileList = Get-ChildItem -Path $($config.OffBoardingInputFolder) | Where-Object { $_.Name -like '*.csv' }
        }
        $csvImport | Format-Table -AutoSize -Wrap        
        # import matching lines
        if ($csvImport) {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Caching all AD users of the domain ', "[$domain]", ' (this can take some time ...) ' -Color White, Green, White, Magenta, White
            #$currentADUsers = Get-ADUser -Filter * -Properties * -ResultSetSize $null -ResultPageSize 9999 -Server $DC -Credential $AD_Credential
            $currentADUsers = Get-ADUser_custom -credential $AD_Credential -systemDomain $($config.SystemDomain) # -server $DC 
        }
        else {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Skipping AD user cacheing ', "[$domain]", ' (no matching lines in the input files ...) ' -Color White, Green, White, Magenta, White
        }
        <# * THE SCRIPT HERE PROCESSES EACH LINE OF ALL THE CSV FILES FOUND IN THE INPUT FOLDER, THAT IS MATCHING THE CURRENTLY PROCESSED DOMAIN. EACH LINE WILL BE PROCESSED ONE WAY OR ANOTHER #>  
        foreach ($line in $csvImport) {
            $output = $null
            $output = [Ordered]@{
                'LeaverName'          = $null
                'LeaverEID'           = $line.EmployeeID
                'Domain'              = $domain                
                'IsProcessed'         = $null
                'ADAcDisabled'        = $null
                'ADPwdReset'          = $null
                'ADGrpsRemoved'       = $null
                'ADDescSet'           = $null
                'ADOrgDetailsCleared' = $null
                'AccMoveedToLeaverOU' = $null
                'MBXProcessed'        = $null
                'MSOUnlicensed'       = $null
                'DATAarchived'        = $null
                'VDI_DD_removed'      = $null
            }
            $output.LeaverName = $leaverName = ($line.FirstName + ' ' + $line.LastName)            
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog -Color White -BackGroundColor Black "[$timer] [i]  Processing user   [ $leaverName ]" -LinesBefore 1
            $leaverUser = $currentADUsers | Where-Object { $_.GivenName -match $line.FirstName -and $_.SurName -match $line.LastName -and $_.EmployeeID -match $line.EmployeeID -and $_.DistinguishedName -notmatch $config.LeaverOU } # only process, if there is a match on Surname, LastName and EmployeeID fields too
            <# *FIRST WE CHECK IF THE REQUESTED LEAVER IS FOUND IN THE DOMAIN AT ALL #>
            #region EMAIL ALERT LEAVER NOT FOUND
            if (-not $leaverUser) {
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog -Color White -BackGroundColor Red "[$timer] [w]  This user account  [$($line.FirstName) $($line.LastName)] and EmployeeID [$($line.EmployeeID)] did not result any matches or the account is already in the [$($config.LeaverOU)] (OU for leavers) - please check if these details are accurate and re-submit once corrected!"
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Leaver can not be identified / already processed. Alerting Service Desk and cancelling work! ' -Color White, Yellow, White
                $output.IsProcessed = 'No (reason: user not found)'
                $props = @{
                    'mode'                = 'leaverNotFound'
                    'domain'              = $domain
                    'smtpServer'          = $config.SMTPServer
                    'testRecipients'      = Get-Content $config.TestRecipientList
                    'recipients'          = Get-Content $config.RecipientList
                    'departmentSignature' = $config.SenderSignature
                    'leaverName'          = $leaverName
                    'leaverSender'        = ($config.LeaverSender + '@' + $config.SystemDomain)
                }
                if ($Test.IsPresent) {
                    Process-Emailing @props -test
                }
                else {
                    Process-Emailing @props
                }
                #endregion
            }
            <# * IF THE LEAVER IS FOUND, THE SCRIPT IS READY TO PROCESS#>
            else {
                # reimport the leaver user  
                $leaverUser = Get-ADUser -Identity $($leaveruser.SAMAccountName) -Properties * -Server $DC -Credential $AD_Credential
                $stopWatch = [system.diagnostics.stopwatch]::StartNew()
                <# *THE SCRIPT FIRST STORE THE GROUPS AND MANAGER OF THE OUTGOING USER #>
                #region CREATE BACKUP OBJECT
                $output.IsProcessed = 'yes'          
                # build object
                $backupObject = @{
                    'Manager' = $null
                    'Groups'  = $null
                }
                if ($leaverUser.manager) {
                    $backupObject.Manager = $leaverUser.manager
                }
                if ($leaverUser.memberof) {
                    $backupObject.Groups = $leaverUser.memberof
                }
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Storeing manager and group details ' -Color White, green, White
                #endregion
                <# * THE ACTIVE DIRECTORY OBJECT WILL BE FIRST PROCESSED AS A LEAVER#>
                #region ACTIVE DIRECTORY
                try {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Disableing the AD account ', " [ $leaverName ] " -Color White, green, White, Green
                    Disable-ADAccount -Identity $leaverUser.DistinguishedName -Server $DC -Credential $AD_Credential
                    $output.ADAcDisabled = 'yes'
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to disable the AD account ', " [ $leaverName ] " -Color White, Red, White, Red
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    $output.ADAcDisabled = 'error'
                    # failing to configure is non-fatal
                    Continue
                }
                # reset password
                Add-Type -AssemblyName System.Web
                $randomPW = [System.Web.Security.Membership]::GeneratePassword(12, 4) 
                try {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Reseting password for the AD account ', " [ $leaverName ] " -Color White, green, White, Green
                    $output.ADPwdReset = 'yes'
                    Set-ADAccountPassword -Identity $leaverUser.DistinguishedName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $randomPW -Force) -Server $DC -Credential $AD_Credential
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to change password of the AD account ', " [ $leaverName ] " -Color White, Red, White, Red
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    $output.ADPwdReset = 'error'
                    # failing to configure is non-fatal
                    Continue
                }
                # remove from all groups
                try {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Removing account from all security groups ', " [ $leaverName ] " -Color White, green, White, Green
                    $output.ADGrpsRemoved = 'yes'
                    $securityGroups = $leaverUser.memberof
                    foreach ($sg in $securityGroups) {
                        Remove-ADGroupMember -Identity $sg -Members $leaverUser.DistinguishedName -Server $DC -Credential $AD_Credential -Confirm:$false #-WhatIf
                    }
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to remove account  ', " [ $leaverName ] ", ' from all security groups ' -Color White, Red, White, Red, White
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    $output.ADGrpsRemoved = 'error'
                    # failing to configure is non-fatal
                    Continue
                }
                # update description
                try {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Updating description of the account ', " [ $leaverName ] " -Color White, green, White, Green
                    $Description = 'LEFT - PWreset by ' + ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) + ' at ' + (Get-Date -Format G)
                    Set-ADUser -Identity $leaverUser.DistinguishedName -Description $Description -Server $DC -Credential $AD_Credential
                    $output.ADDescSet = 'yes'
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to update the description of the account ', " [ $leaverName ] " -Color White, Red, White, Red
                    $output.ADDescSet = 'error'
                    # failing to configure is non-fatal
                    Continue
                } 
                # clear company details
                try {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Clearing organisational details ', " [ $leaverName ] " -Color White, green, White, Green
                    $fields = 'physicalDeliveryOfficeName', 'Telephonenumber', 'Title', 'Department', 'Company', 'extensionAttribute10'
                    $fields = 'physicalDeliveryOfficeName', 'Telephonenumber', 'Title', 'Department', 'Company', 'manager', 'extensionAttribute10'
                    foreach ($f in $fields) {
                        $leaverUser | Set-ADObject -Clear $f -Server $DC -Credential $AD_Credential
                    }
                    $output.ADOrgDetailsCleared = 'yes'
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to clear organisational details ', " [ $leaverName ] " -Color White, Red, White, Red
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    $output.ADOrgDetailsCleared = 'error'
                    # failing to configure is non-fatal
                    Continue
                }   
                # move to the leaver OU  
                try {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Moving account ', " [ $leaverName ] ", ' to the leaver OU ', " [ $($config.LeaverOU) ] " -Color White, green, White, Green , White, Green
                    Move-ADObject -Identity $leaverUser.DistinguishedName -TargetPath $config.LeaverOU -Server $DC -Credential $AD_Credential
                    $output.AccMoveedToLeaverOU = 'yes'
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to move the account ', " [ $leaverName ] ", ' to the leaver OU ', " [ $($config.LeaverOU) ] " -Color White, Red, White, Red, White, Red
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    $output.AccMoveedToLeaverOU = 'error'
                    # failing to configure is non-fatal
                    Continue
                }
                #endregion
                <# * NEXT THE MAILBOX OF THE USER WILL BE PROCESSED #>
                #region EXCHANGE
                # connect to exchange onprem and online
                try {
                    ## remove previous sessions
                    Get-PSSession | Remove-PSSession
                    # connect to Exchange ON-PREM
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Connecting to on-prem exchange ' -Color White, Green, White
                    Connect-OnPremExchange -credential $Exch_Credential -exchangeServer ($config.ExchangeServer + '.' + $config.SystemDomain)   
                    # connect to Exchange ONLINE (prefix: "o")
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Connecting to online exchange ' -Color White, Green, White
                    Connect-OnlineExchange -credential $AAD_Credential  
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to connect to the exchange environments.' -Color White, Red, White
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    [void] (Stop-Transcript -ErrorAction Ignore)
                    Break   
                }
                # process mailbox
                try {
                    # if the mailbox is on-prem --> migrate online
                    If (!(($leaverUser).TargetAddress -match 'onmicrosoft.com')) {
                        # after ensureing the target address for o365 is on the mailbox.
                        $onlineProxyAddress = 'smtp:' + $leaveruser.SAMAccountName + '@' + $config.EOTargetDomain
                        if (((Get-ADUser -Identity $leaveruser.SAMAccountName -Properties * -Server $DC -Credential $AD_Credential).ProxyAddresses) -notcontains $onlineProxyAddress) {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding online target address ', "[$($onlineProxyAddress)]", ' to the mailbox of ', "[$($leaveruser.UserPrincipalName)]" -Color White, Green, White, Green, White, Green
                            Set-ADUser -Identity $($leaveruser.SAMAccountName) -Add @{ProxyAddresses = $onlineProxyAddress } -Server $DC -Credential $AD_Credential
                            # sync changes cross-domain
                            try {
                                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Synchronising domain controllers of domain ', " [$domain] " -Color White, Green, White, Magenta
                                Sync-ActiveDirectory -server $DC -credential $AD_Credential -Verbose
                            }
                            catch {
                                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Syncronisation of domain controllers failed for domain ', "[$domain]", ' As this prevents the rest of the process, we proceed to the next user!' -Color White, Red, White, Magenta, White
                                "[$timer] [ERROR] Processing of the account [$($leaveruser.Name)] failed! Error details: " | Out-File $ErrorFile -Append
                                Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append  
                                [void] (Stop-Transcript -ErrorAction Ignore)
                                Break  
                            }                            
                            # sync changes to AZURE AD
                            try {                                
                                Sync-AzureActiveDirectory -server $($config.AADSyncServer + '.' + $config.SystemDomain) -credential $AD_Credential
                            }
                            catch {
                                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Syncronisation to AZURE AD failed for domain ', "[$domain]", ' As this prevents the rest of the process, we proceed to the next user!' -Color White, Red, White, Magenta, White
                                "[$timer] [ERROR] Processing of the account [$($leaveruser.Name)] failed! Error details: " | Out-File $ErrorFile -Append
                                Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append  
                                [void] (Stop-Transcript -ErrorAction Ignore)
                                Break  
                            }
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) 
                            Write-Warning "[$timer]  [e] THIS SHOULD NOT HAVE HAPPENED! The $onlineProxyAddress is not yet syncronised on $($leaveruser.SAMAccountName) ! To allow this to happen, we suspend the script for 40 minutes, to allow complete Azure syncronisation. Sorry for the inconvenience! "
                            Start-Sleep -Seconds 2700 # long sleep
                        }
                        else {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Online target addresss ', "[$($onlineProxyAddress)]", ' is present on mailbox ', "[$($leaveruser.UserPrincipalName)]" -Color White, Green, White, Green, White, Green
                        }
                        # Do the mailbox move                
                        $props = @{
                            'UPN'                  = $leaverUser.UserPrincipalName 
                            'remoteHostName'       = ($config.OnpremisesMRSProxyURL + '.' + $config.SystemDomain)
                            'targetDeliveryDomain' = $config.EOTargetDomain
                            'user'                 = $leaverUser
                        }
                        Move-MailboxOnline @props -remoteCredential $AD_Credential
                    }
                    elseif (($leaverUser).TargetAddress) {
                        <# reconfigure mailbox:
                    - add forwarding address, if defined
                    - set litigation hold to 7 years (for  primary group domain only)
                    - add out of office message
                    #TODO: take note 
                    - remove additional proxy addresses (this is to prevent clash, if the leaver returns to the company) 
                    - hide from GLAddress list
                    - 
                    #>
                        $props = @{
                            'UPN'                = $leaverUser.UserPrincipalName
                            'litigationHoldTime' = 2555 
                            'systemDomain'       = $config.systemDomain
                            # 'forwardee'          = $forwardingAddress
                        }
                        #TODO: Check forwarding
                        If ($Line.ForwardingAddress -like "*$Domain*") {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Forwarding address ', "[$($Line.ForwardingAddress)]", ' matches with ', "[$domain]" -Color White, Green, White, Green, White, Green
                            # add forwarding address only if given and matching to the domain
                            $forwardingAddress = $oomRecipient = $Line.ForwardingAddress
                            $props['forwardee'] = $forwardingAddress                        
                        }
                        elseif ($Line.ForwardingAddress) {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Forwarding address ', "[$($Line.ForwardingAddress)]", ' does not match ', "[$domain]" -Color White, Green, White, Green, White, Green
                        }
                        else {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Fowarding address not specified ' -Color White, Green, White
                        }
                        # else {
                        #     $forwardingAddress = $oomRecipient = $null
                        # }
                        # reconfigure mailbox
                        $x = 0
                        $success = $false
                        do {
                            $x++
                            Write-Host "Attempt $x"
                            try {
                                if ($config.ProcessSpecial -match 'Yes') {
                                    Reconfigure-Onlinemailbox @props -withLitigation
                                }
                                elseif ($config.ProcessSpecial -match 'No') {
                                    Reconfigure-Onlinemailbox @props 
                                }                    
                                $props = @{
                                    'company'      = $leaveruser.Company 
                                    'oomRecipient' = $oomRecipient
                                    'UPN'          = $leaverUser.UserPrincipalName
                                    'name'         = $leaverName
                                }
                                # set OOM
                                Set-OOOMessage @props
                                $success = $true
                            }
                            catch {
                                Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                                $success = $false
                                Continue
                            }
                        } until ($x -eq 5 -or $success -eq $true)
                        if ($success -eq $True) {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Succesfully processed the mailbox', " [ $($leaverUser.UserPrincipalName) ] " -Color White, Green, White, Green
                        }
                        else {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Could not locate leaver mailbox', " [ $($leaverUser.UserPrincipalName) ] ", " (max attempt [$x] reached)" -Color White, Red, White, Red, white
                        }
                        # remove proxy addresses
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Removing proxy addresses ', " [ $leaverName ] " -Color White, green, White, Green
                        Set-ADUser -Identity $leaverUser.SAMAccountName -Clear ProxyAddresses -Server $DC -Credential $AD_Credential
                        # hide from Address Book
                        try {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Hiding account from the Global Address Book (replace) ', " [ $leaverName ] " -Color White, green, White, Green
                            Set-ADUser -Identity $leaverUser.SAMAccountName -Replace @{ msExchHideFromAddressLists = $true } -Server $DC -Credential $AD_Credential
                        }
                        catch {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Hiding account from the Global Address Book (add) ', " [ $leaverName ] " -Color White, green, White, Green
                            Set-ADUser -Identity $leaverUser.SAMAccountName -Add @{ msExchHideFromAddressLists = $true } -Server $DC -Credential $AD_Credential
                            Continue
                        }
                        $output.MBXProcessed = 'yes'
                    }
                    else {
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Could not locate leaver mailbox', " [ $($leaverUser.UserPrincipalName) ] "-Color White, Red, White, Red 
                        $output.MBXProcessed = 'error - mailbox not found'
                    }
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to process the mailbox', " [ $($leaverUser.UserPrincipalName) ] "-Color White, Red, White, Red
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    $output.MBXProcessed = 'error'
                    # failing to configure is non-fatal
                    Continue
                }
                #endregion
                <# * THE LICENSES OF THE USER WILL BE REMOVED (REALLY JUST THE DIRECT ASSIGNED LICENSES HERE, AS THE REMOVAL FROM ALL AD GROUPS - PART OF THE AD PROCESSING - WOULD RESULT IN REMOVE GROUP ASSIGNMENTS) #>
                #region LICENSING
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Connecting to MSOL Service ' -Color White, green, White
                # connect to MS online
                [void] (Connect-MsolService -Credential $AAD_Credential)
                # remove all directly assigned licenses
                try {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Removing licenses from user:', " [ $leaverName ] " -Color White, green, White, Green
                    Remove-DirectLicenses -UPN $leaverUser.UserPrincipalName -Verbose
                    $output.MSOUnlicensed = 'yes'
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to remove directly assigned licenses from user ', " [ $($leaverUser.UserPrincipalName) ] " -Color White, Red, White, Red
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    $output.MSOUnlicensed = 'error'
                    # failing to configure is non-fatal
                    Continue
                }
                #endregion
                <# * NEXT THE MANAGER/GROUPS OF THE USER IS BEING BACKED UP#>
                #region DATA BACKUP
                try {
                    $archiveDrive = $config.ArchiveServer + '.' + $config.SystemDomain + '\' + $config.ArchiveDisk + $config.ArchivePathSuffix
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Backing up user details ' -Color White, green, White
                    if ($config.ProcessSpecial -match 'Yes') {
                        $archiveFolder = '\\' + $archiveDrive + '\ARCHIVE_LEAVERS'
                        # backup manager and groups
                        $props = @{
                            'archiveFolder'  = $archiveFolder
                            'manager'        = $backupObject.Manager
                            'groups'         = $backupObject.Groups
                            'samAccountName' = $leaverUser.SAMAccountName
                            'backupServer'   = $config.ArchiveServer + '.' + $config.SystemDomain
                        }
                        Backup-LeaverDetails @props -credential $AD_Credential
                        # backup DFS
                        $props = @{
                            'archiveFolder'  = $archiveFolder
                            'samAccountName' = $leaverUser.SAMAccountName
                            'backupServer'   = ($config.ArchiveServer + '.' + $config.SystemDomain)
                            'peopleDFS'      = ('\\' + $config.SystemDomain + '\PEOPLE\' + $leaverUser.SAMAccountName)
                            'profileDFS'     = ('\\' + $config.SystemDomain + '\PROFILES\' + $leaverUser.SAMAccountName)
                            'dfsnServer'     = $DC
                        }
                        Backup-DFS @props -credential $AD_Credential             
                    }
                    elseif ($config.ProcessSpecial -match 'No') {
                        $archiveFolder = '\\' + $archiveDrive + '\Leavers'
                        # backup manager and groups
                        $props = @{
                            'archiveFolder'  = $archiveFolder
                            'manager'        = $backupObject.Manager
                            'groups'         = $backupObject.Groups
                            'samAccountName' = $leaverUser.SAMAccountName
                            'backupServer'   = ($config.ArchiveServer + '.' + $config.SystemDomain)
                        }
                        Backup-LeaverDetails @props -credential $AD_Credential
                    }
                    $output.DATAarchived = 'yes'
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to complete data backups for user ', " [ $($leaverUser.SAMAccountName) ] " -Color White, Red, White, Red
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    $output.DATAarchived = 'error'
                    # failing to configure is non-fatal
                    Continue
                }
                #endregion
                <# * IF THE DOMAIN USING IT, THE VDI OF THE USER AND THE RDS PROFILE DISK IS BEING REMOVED#>
                #region VDI / profile disk removal
                try {
                    if ($config.ProcessSpecial -match 'Yes') {
                        # remove VDI
                        $props = @{
                            'cbServer'       = ($config.ConnectionBrokerServer + '.' + $config.SystemDomain)
                            'domainNetBIOS'  = $config.DomainNetBios
                            'samAccountName' = $leaverUser.SAMAccountName
                            'SCCMServer'     = ($config.SCCMServer + '.' + $config.SystemDomain)
                            'VMMServer'      = ($config.VMMServer + '.' + $config.SystemDomain)
                            'SCCMSiteCode'   = $config.SCCMSiteCode
                            'server'         = $DC
                            #'systemDomain' = $config.SystemDomain
                        }
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Removing the VDI of the leaver (if exists) ' -Color White, Green, White
                        $VDIRemovalResult = Remove-VDI @props -credential $AD_Credential
                        # # remove profile disk (no data kept)                
                        # $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Removing profiledisk for  ', " [ $($leaverUser.SAMAccountName) ] ", ' from ', " [ $($config.RDSDiskFileServer + '.' + $config.SystemDomain)]" -Color White, Green, White, Green , White, Green
                        # Remove-ProfileDisk -profileDiskLocation $('\\' + $config.RDSDiskFileServer + '.' + $config.SystemDomain + '\' + $config.ProfileDiskFolder) -samAccountName $leaverUser.SAMAccountName                        
                        # Write-Host 'Output at the end of VDI processing' -ForegroundColor Cyan
                        $output.VDI_DD_removed = $VDIRemovalResult
                        # remove profile disk (no data kept)                
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Removing profiledisk for  ', " [ $($leaverUser.SAMAccountName) ] ", ' from ', " [ $($config.RDSDiskFileServer + '.' + $config.SystemDomain)]" -Color White, Green, White, Green , White, Green
                        Remove-ProfileDisk -profileDiskLocation $('\\' + $config.RDSDiskFileServer + '.' + $config.SystemDomain + '\' + $config.ProfileDiskFolder) -samAccountName $leaverUser.SAMAccountName                          
                    }
                    else {
                        $output.VDI_DD_removed = 'n/a'
                    }
                }
                catch {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to remove the VDI or profiledisk of the user:', " [ $($leaverUser.SamAccountName) ] " -Color White, Red, White, Red
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    $output.VDI_DD_removed = 'error'                    
                    # failing to configure is non-fatal
                    Continue
                }
                #region EMAIL HR TO RETRIEVE PHONE
                $props = @{
                    'mode'                = 'returnPhone'
                    'domain'              = $domain
                    'smtpServer'          = $config.SMTPServer
                    'testRecipients'      = Get-Content $config.TestRecipientList
                    'recipients'          = Get-Content $config.RecipientList
                    'HRRecipients'        = Get-Content $config.HR_Recipients
                    'departmentSignature' = $config.SenderSignature
                    'leaverName'          = $leaverName
                    'leaverEID'           = $leaveruser.EmployeeID
                    'leaverSender'        = ($config.LeaverSender + '@' + $config.SystemDomain)
                }
                if ($Test.IsPresent) {
                    Process-Emailing @props -test
                }
                else {
                    Process-Emailing @props
                }
                #endregion
                #region UPDATE USERBASE
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Updating userbase post-processing ' -Color White, Green, White
                try {
                    #$currentADUsers = Get-ADUser -Filter * -Properties * -ResultSetSize $null -ResultPageSize 9999 -Server $DC -Credential $AD_Credential
                    $currentADUsers = Get-ADUser_custom -credential $AD_Credential -systemDomain $($config.SystemDomain) # -server $DC 
                }
                catch {
                    Write-Warning "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                    Continue
                }
                #endregion
                #endregion
                $stopWatch.stop()
                Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Removing user ', " [$($leaveruser.DisplayName)] ", ' took ', " [$([math]::Round(($stopWatch.Elapsed.TotalMinutes), 2))] ", ' minutes ' -Color White, Magenta, White, Magenta, White, Magenta, White
                # Only add succesfull sending to the report
                $summaryReport += [pscustomobject]$output 
            }
        }
        <# * NEXT WE START PRODUCING REPORTS OF THE WORK DONE. THE SUPPORT TEAM RECEIVES AN EMAIL WITH LINKS TO THE PROCESSING LOGS, SUCCESFULLY PROCESSED USER LIST, ERROR LOGS AND THE BACKUP FILES #>
        if ($csvImport) {
            $csvImport | Export-Csv $backupFile -Append -NoTypeInformation
        }
        Write-Host # line break
        if ($summaryReport) {         
            $summaryReport | Export-Csv $reportFile -Append -NoTypeInformation
        }
        <# * SENDING REPORT EMAIL ONLY, IF THERE WAS ANY WORK DONE #>
        if ($summaryReport -and $csvImport) {
            #region EMAIL SD REMOVAL SUMMARY
            $props = @{
                'mode'                = 'toSupportRemoval'
                'domain'              = $domain
                'smtpServer'          = $config.SMTPServer
                'testRecipients'      = Get-Content $config.TestRecipientList
                'recipients'          = Get-Content $config.RecipientList
                'HRRecipients'        = Get-Content $config.HR_Recipients
                'departmentSignature' = $config.SenderSignature
                'leaverName'          = $leaverName
                'leaverEID'           = $leaveruser.EmployeeID
                'leaverSender'        = ($config.LeaverSender + '@' + $config.SystemDomain)
                'report'              = $summaryReport
                # 'failureReport'       = $failureSummaryReport
                'link1'               = ($config.offboardingOutputFolder + '\' + $sDate + '\' + $logTime + '_offboarding_report.csv') # $reportFile #
                'link2'               = ($config.LogsFolder + '\' + $sDate + '\' + $logTime + '_offboarding_error.log') # $errorFile # 
                'link3'               = ($config.offboardingOutputFolder + '\' + $sDate + '\' + $logTime + '_offboarding_backup.csv') # $backupFile # 
                'link4'               = ($config.LogsFolder + '\' + $sDate + '\' + $logTime + '_offboarding_transcript.log') # $transcriptFile #
                'link5'               = ($config.LogsFolder + '\' + $sDate + '\' + $logTime + '_offboarding_action.log') # $actionLog #         
                'Attachments'         = @(
                    $reportFile
                    #$($config.OffBoardingOutputFolder + '\' + $sDate + '\' + $logTime + '_report.csv')
                )
            }
            if ($Test.IsPresent) {
                Process-Emailing @props -test
            }
            else {
                Process-Emailing @props
            }
            #endregion
        }
        elseif ($csvImport) {
            #region EMAIL ALERT
            $props = @{
                'leaverName'          = $leaverName -replace '\.', ' '
                'mode'                = 'genericOffboardingFailure'
                'domain'              = $domain
                'smtpServer'          = $config.SMTPServer
                'testRecipients'      = Get-Content $config.TestRecipientList
                'recipients'          = Get-Content $config.RecipientList
                'departmentSignature' = $config.SenderSignature
                #'templateName'        = $line.CopyAccount
                'starterSender'       = ''
                'leaverSender'        = ($config.LeaverSender + '@' + $config.SystemDomain)
            }   
            if ($Test.IsPresent) {
                Process-Emailing @props -test
            }
            else {
                Process-Emailing @props
            }
            #endregion 
        }
    }
}
END {
    <# * FINALLY THE SCRIPT DOES A CLEANUP, REMOVING THE STORED USER CACHE, RENAMING / REMOVING THE INPUT FILES, ETC. ALSO IT ENDS THE TRANSCRIPT AND SENDS ERRORS TO THE ERROR REPORT FILE #>
    # drop user cache
    $currentADUsers = $null
    # # drop input files
    # Get-ChildItem -Path $config.OffBoardingInputFolder | Where-Object { $_.Name -like '*.csv' } | Remove-Item -WhatIf
    # rename input files
    if (-not $test.ispresnt) {
        $CSVs = $fileList # Get-ChildItem -Path $config.OffBoardingInputFolder | Where-Object { $_.Name -like '*.csv' }
        foreach ($CSV in $CSVs) {
            $newName = $CSV.Name -replace '.csv', '.processed'
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Renaming file  ', " [$($CSV.Name) --> $newName] " -Color White, Green, White, Cyan
            Rename-Item -Path $CSV.FullName -NewName $newName -ErrorAction SilentlyContinue
        }
    }
    # Finalise logs
    if ($Error) {
        "[WARN] ERRORS DURING SCRIPT RUN [$sDate]" | Out-File $ErrorFile
        $Error | Out-File $ErrorFile 
    }
    else {
        "[INFO] NO ERRORS DURING SCRIPT RUN [$sDate]" | Out-File $ErrorFile
    }
    #TODO: Drop input files
    [void] (Stop-Transcript -ErrorAction Ignore)
    # Close all open sessions
    Get-PSSession | Remove-PSSession
    Get-CimSession | Remove-CimSession
}
# SIG # Begin signature block
# MIIOWAYJKoZIhvcNAQcCoIIOSTCCDkUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUpu9Plca64MMvrAplXB+z/ovH
# mBGgggueMIIEnjCCA4agAwIBAgITTwAAAAb2JFytK6ojaAABAAAABjANBgkqhkiG
# 9w0BAQsFADBiMQswCQYDVQQGEwJHQjEQMA4GA1UEBxMHUmVhZGluZzElMCMGA1UE
# ChMcV2VzdGNvYXN0IChIb2xkaW5ncykgTGltaXRlZDEaMBgGA1UEAxMRV2VzdGNv
# YXN0IFJvb3QgQ0EwHhcNMTgxMjA0MTIxNzAwWhcNMzgxMjA0MTE0NzA2WjBrMRIw
# EAYKCZImiZPyLGQBGRYCdWsxEjAQBgoJkiaJk/IsZAEZFgJjbzEZMBcGCgmSJomT
# 8ixkARkWCXdlc3Rjb2FzdDEmMCQGA1UEAxMdV2VzdGNvYXN0IEludHJhbmV0IElz
# c3VpbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC7nBk9j3wR
# GgkxrPuXjIXlptisoOhKZp7KCB+BhxaxlTGW5lxhEaNirirM4jaM04kXojFZxhHV
# lTl2W3TPOfeIEXxcZYigPgh9d6wgTTb2cSRq1872YjMytxSps14LAbY8CEu+fQmC
# AbL6V8EgtnAmzMBBqOOi6x7bMHoGkJPwDOSUM01LHPoT8cg9KVIFioJHpex/Xeko
# FiRwgW7uS+dh57iCGRWVCZaDrFIXWKj4dOHJigsEPkbmJUPSYILF8SYglFiJpM7b
# xl3RPuy2GvJRq5Ikyn0SvnpAG72Ge664PV5sFdtzdNkIE7RsE6zUEqK1v2pt7CcC
# qh4en3v54ouZAgMBAAGjggFCMIIBPjASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsG
# AQQBgjcVAgQWBBSBYkDZbTpVK0nuvapWivWUf0tBKDAdBgNVHQ4EFgQUU3PVQuhx
# ickSLEsfPyKpNozqrT8wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0P
# BAQDAgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAUuxfhV4noKzmJ
# eDD6ejIRp0cSBu8wPQYDVR0fBDYwNDAyoDCgLoYsaHR0cDovL3BraS53ZXN0Y29h
# c3QuY28udWsvcGtpL3Jvb3RjYSgxKS5jcmwwSAYIKwYBBQUHAQEEPDA6MDgGCCsG
# AQUFBzAChixodHRwOi8vcGtpLndlc3Rjb2FzdC5jby51ay9wa2kvcm9vdGNhKDEp
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEAaYMr/xfHuo3qezz8rtbzGkfUwqNFjd0s
# 7d02B07aO5q0i7LMtZTMxph9DbeJRvm+d8Sr4DSiWgtJdb0eYsx4xj5lDrsXDuO2
# 2Mb4hKjtqzDVW5PEJzC72BPOSfkgfW6PZmscMPtJnn0TPM24DzkYmjhnsA97Ltjv
# 1wuvUi2G0nPIbzfBZWnnuCx5PhSovssQU5E3ZlVLew6a8WME0lPOmR9c38TARqWh
# tvS/wqmUaCEUF6rmUDY0MgY/Wrg2TIbtlYFWe9PksI4jmTE4Ndy5BW8smx+8YOoF
# fCOldshHHgFJVG7Bat6vrT8AaUSs6crPBRMpbeouD0iujXts+LdV2TCCBvgwggXg
# oAMCAQICEzQAA+ZyHBAttK7qIqcAAQAD5nIwDQYJKoZIhvcNAQELBQAwazESMBAG
# CgmSJomT8ixkARkWAnVrMRIwEAYKCZImiZPyLGQBGRYCY28xGTAXBgoJkiaJk/Is
# ZAEZFgl3ZXN0Y29hc3QxJjAkBgNVBAMTHVdlc3Rjb2FzdCBJbnRyYW5ldCBJc3N1
# aW5nIENBMB4XDTIwMDUxODA4MTk1MloXDTI2MDUxODA4Mjk1MlowgacxEjAQBgoJ
# kiaJk/IsZAEZFgJ1azESMBAGCgmSJomT8ixkARkWAmNvMRkwFwYKCZImiZPyLGQB
# GRYJd2VzdGNvYXN0MRIwEAYDVQQLEwlXRVNUQ09BU1QxDTALBgNVBAsTBExJVkUx
# DjAMBgNVBAsTBVVTRVJTMQ8wDQYDVQQLEwZBZG1pbnMxHjAcBgNVBAMTFUZhYnJp
# Y2UgU2VtdGkgKEFETUlOKTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# APVwqF2TGtzPlxftCjtb23neDu2cWyovIpo1TgU0ptNYrJM8tAY6W8Yt5Vw+8xzU
# 45sxmbMzU2JpJaqEPFe3+gXWJtL99/ZusyXCDbubzYmNu06WE6XqMqG/KRfZ3BpN
# Gw5s3KlxWVj/H12i7JPbMvfyAl8lgz/YBO0XVdoozcAglEck7c8DBaRTb4J7vX/O
# IS7dYu+gmkZJCv2+O6vTNTlK7bIHAQPWzSPibzU9dRPlHiPOTcHoYB+YNpmbgNxn
# fdaFMB+xY1GcYoKwVRl6UEF/od8TKehzUp/hHFlXiH+miz692ptXhi3dOp6R4Stn
# Ku0IoBfBi/CQcgl5Uko6kckCAwEAAaOCA1YwggNSMD4GCSsGAQQBgjcVBwQxMC8G
# JysGAQQBgjcVCIb24huEi+UUg4mdM4f4p0GE8aVDgSaGkPwogZ23PAIBZAIBAjAT
# BgNVHSUEDDAKBggrBgEFBQcDAzALBgNVHQ8EBAMCB4AwGwYJKwYBBAGCNxUKBA4w
# DDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQU7eheFlEriypJznAoYQVEx7IAmBkwHwYD
# VR0jBBgwFoAUU3PVQuhxickSLEsfPyKpNozqrT8wggEuBgNVHR8EggElMIIBITCC
# AR2gggEZoIIBFYY6aHR0cDovL3BraS53ZXN0Y29hc3QuY28udWsvcGtpLzAxX2lu
# dHJhbmV0aXNzdWluZ2NhKDEpLmNybIaB1mxkYXA6Ly8vQ049V2VzdGNvYXN0JTIw
# SW50cmFuZXQlMjBJc3N1aW5nJTIwQ0EoMSksQ049Qk5XQURDUzAxLENOPUNEUCxD
# Tj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1
# cmF0aW9uLERDPXdlc3Rjb2FzdCxEQz1jbyxEQz11az9jZXJ0aWZpY2F0ZVJldm9j
# YXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9pbnQw
# ggEmBggrBgEFBQcBAQSCARgwggEUMEYGCCsGAQUFBzAChjpodHRwOi8vcGtpLndl
# c3Rjb2FzdC5jby51ay9wa2kvMDFfaW50cmFuZXRpc3N1aW5nY2EoMSkuY3J0MIHJ
# BggrBgEFBQcwAoaBvGxkYXA6Ly8vQ049V2VzdGNvYXN0JTIwSW50cmFuZXQlMjBJ
# c3N1aW5nJTIwQ0EsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENO
# PVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9d2VzdGNvYXN0LERDPWNvLERD
# PXVrP2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9u
# QXV0aG9yaXR5MDUGA1UdEQQuMCygKgYKKwYBBAGCNxQCA6AcDBp3Y2FkbWluLmZz
# QHdlc3Rjb2FzdC5jby51azANBgkqhkiG9w0BAQsFAAOCAQEAeM0HkiWDX+fmhIsv
# WxZb+D/tLDztccfYND16zFAoReu0VmTUz570CEMhLyHGh1jk3y/pb26UmjqHFeVh
# /EVu/EQNCuT5gQPKh64FQsBVinugNHWMhDySywykKwkdnqEpY++UNxQyyj6xpTM0
# tg+h8Wd1IlDN98SwLBy4x16SwgGTdwKvU9CyBuMRQjPlSJKjCL+14T0C8d2SBGW3
# 9uLCqjyMd288Q3QgrbDoHSg/x+vsnrDzOHMThM/2aMPbcO0wqafK9G5qdoIc0dqe
# So/vU6rsNLwQ1sniJQxerKZnWJjEfl8M5OcUxws5n7D3fqpHZ2VxLCIYp6yuPkHY
# R5daezGCAiQwggIgAgEBMIGCMGsxEjAQBgoJkiaJk/IsZAEZFgJ1azESMBAGCgmS
# JomT8ixkARkWAmNvMRkwFwYKCZImiZPyLGQBGRYJd2VzdGNvYXN0MSYwJAYDVQQD
# Ex1XZXN0Y29hc3QgSW50cmFuZXQgSXNzdWluZyBDQQITNAAD5nIcEC20ruoipwAB
# AAPmcjAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUg2Q5Vlh+qhnHu/uxAlMk/P1JEBwwDQYJKoZI
# hvcNAQEBBQAEggEA4QqqphVSEIvfxj9R/Sm09C62VIWZd9BjIziT7p7zXEC4KybF
# Urhz21+Zlf2F+WZF490oJpLWHT1rCCEll8tXGtlSPaesU3QnVCxaKvsoiCXMvOL1
# WY21FE136Eayj1o8kXda1Ai+SRKCnIOvf5dDelR9x6NnG/HwbgSy9SjsCLUCOUcg
# jSK07He8D3Po7uikZPJ4pquESa5lSgY+6fpUZX4AxZ0KjL6vJuRuTufxKWIeN1fu
# UXCQcyn5dDwjMuC2BiqwaXkoubhjwPAPq8+R7yQSwLTveWDajK0Bj7CHYJqKr6hk
# I0vpKNEFkNXlUtZeMtW27kdOJ2lrygL4PfDVsQ==
# SIG # End signature block
