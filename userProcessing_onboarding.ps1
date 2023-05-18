<#
.SYNOPSIS
    The script is designed to carry out the onboarding process for the company employees.
.DESCRIPTION
    Please see the readme.md !
.EXAMPLE
    Example of how to use this cmdlet
.EXAMPLE
    Another example of how to use this cmdlet
#>
[cmdletbinding(SupportsShouldProcess = $True)]
param(
    [Parameter(Mandatory = $false)][switch] $test
)
BEGIN {
    <# * THE FIRST PART OF THE SCRIPTBLOCK IS TO PREPARE THE SCRIPT ENVIRONMENT - IMPORTING MODULES, DEFINING THE CURRENT LOCATION, ETC. - AND DEFINING THE VARIOUS LOG FILES USED TROUGHOUT THE SCRIPT #>
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
    if (-not (Test-Path "$CurrentPath\out-on\$sDate")) {
        [void](New-Item -ItemType Directory -Path "$CurrentPath\out-on" -Name $sDate)
    }
    if (-not (Test-Path "$CurrentPath\logs\$sDate")) {
        [void](New-Item -ItemType Directory -Path "$CurrentPath\logs" -Name $sDate)
    }
    $doneFolder = "$CurrentPath\in-on\done" 
    if (-not (Test-Path $doneFolder )) {
        [void](New-Item -ItemType Directory -Path "$CurrentPath\in-on" -Name 'done')
    }     
    #endregion
    #region LOGS
    $logTime = (Get-Date -Format yyy-MM-dd-HH-mm)
    $transcriptFile = $CurrentPath + '\logs\' + $sDate + '\' + $logTime + '_onboarding_transcript.log'
    #$transcriptFile = $CurrentPath + '\logs\' + $sDate + '\' + $logTime + '_onboarding_transcript.log'
    $backupFile = $CurrentPath + '\out-on\' + $sDate + '\' + $logTime + '_onboarding_backup.csv'
    $errorFile = $CurrentPath + '\logs\' + $sDate + '\' + $logTime + '_onboarding_error.log'
    $reportFile = $CurrentPath + '\out-on\' + $sDate + '\' + $logTime + '_report.csv'
    $actionLog = $CurrentPath + '\logs\' + $sDate + '\' + $logTime + '_onboarding_action.log'
    [void] (Remove-Item -Path $backupFile -ErrorAction Ignore -Force)
    Start-Transcript -Path $transcriptFile -Force
    #endregion
}
PROCESS {
    <# *  THE PROCESS BLOCK CONTAINS THE MAJORITY OF THE SCRIPT'S WORK CODES. THIS STARTS WITH LISTING THE ACCEPTABLE DOMAINS (FROM THE CONFIGURATION FILE) #>
    $domainList = (Import-PowerShellDataFile $CurrentPath\userProcessing.psd1).DomainList
    <# * FOR EACH DOMAIN WE WILL IMPORT THE SPECIFIC VARIABLES - IE. EXCHANGE SERVER NAME OR ADMIN ACCOUNT NAME FOR AZURE ACTIVE DIRECTORY - AND ALSO IMPORT THE VARIABLES SHARED BETWEEN THE DOMAINS. THE OTHER VERY IMPORTANT ACTION IN THIS BLOCK IS CACHING THE ACTIVE DIRECTORY USER BASE. THIS ADS A FEW MINUTE DELAY AT THE BEGINING OF THE SCRIPT, BUT SAVES TIME LATER ON AND ALLOWS THE SCRIPT'S  FUNCTIONS TO USE THIS, ELIMINIATING THE NEED TO PASS ON CREDENTIAL PARAMETER TO QUERY ACTIVE DIRETORY#>    
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
        Cleanup-EarlierRuns -inputFolder $($config.OnBoardingInputFolder) -doneFolder $doneFolder
        if ($test.IsPresent) {
            $file = Get-FileName -initialDirectory $($config.OnBoardingInputFolder)
            if ($file) {
                $csvImport = Import-Csv $file | Where-Object { $_.Domain -match $domain } 
            }   
        }
        else {
            $csvImport = Import-WorkFiles -inputFolder $($config.OnBoardingInputFolder) | Where-Object { $_.Domain -match $domain }
            # * STORE FILE LIST
            $fileList = Get-ChildItem -Path $($config.OnBoardingInputFolder) | Where-Object { $_.Name -like '*.csv' }
        }
        $csvImport | Format-Table -AutoSize -Wrap
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Caching all AD users of the domain ', "[$domain]", ' (this can take some time ...) ' -Color White, Green, White, Magenta, White
        if ($csvImport) {
            $currentADUsers = Get-ADUser -Filter * -Properties * -Server $DC -Credential $AD_Credential
        }  
        <# * THE SCRIPT HERE PROCESSES EACH LINE OF ALL THE CSV FILES FOUND IN THE INPUT FOLDER, THAT IS MATCHING THE CURRENTLY PROCESSED DOMAIN. EACH LINE WILL BE PROCESSED ONE WAY OR ANOTHER #>      
        foreach ($line in $csvImport) {            
            $checkName = $line.FirstName + '.' + $line.LastName
            $checkEID = $line.EmployeeID
            <# * FIRST THE SCRIPT CHECKS IF THE USER EXISTS, BY COMPARING THE INPUT TO THE USERBASE. IF IT FINDS A USER THAT HAS MATCHING EMPLOYEEID OR MATCHING NAMES, THEN IT ASSUMES THE USER IS ALREADY CREATED, AND DOES NOT CREATE A NEW ACCOUNT #>
            #region USER EXISTS CHECKS
            if ( $currentADUsers | Where-Object { ($_.SAMAccountName -match $checkName) -and ($_.UserPrincipalName -match $checkName) -and ($_.EmployeeID -match $checkEID) } ) {
                $offendingUser = $currentADUsers | Where-Object { ($_.SAMAccountName -match $checkName) -and ($_.UserPrincipalName -match $checkName) -and ($_.EmployeeID -match $checkEID) }
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog -Color White -BackGroundColor Red "[$timer] [w]  This user account  [$($line.FirstName) $($line.LastName)] and EmployeeID [$($line.EmployeeID)] already matches to a user [$($offendingUser.SAMAccountName)  / $($offendingUser.EmployeeID)], skipping user creation - please contact HR and re-submit new starter with a unique user details!" -LinesAfter 1
                Add-Member -InputObject $line -NotePropertyName 'IsProcessed' -NotePropertyValue 'No' -Force   
                #region EMAIL ALERT
                $props = @{
                    'newStarterFullName'  = $checkName -replace '\.', ' '
                    'mode'                = 'genericOnboardingFailure'
                    'domain'              = $domain
                    'smtpServer'          = $config.SMTPServer
                    'testRecipients'      = Get-Content $config.TestRecipientList
                    'recipients'          = Get-Content $config.RecipientList
                    'departmentSignature' = $config.SenderSignature
                    'templateName'        = $line.CopyAccount
                    'starterSender'       = ($config.StarterSender + '@' + $config.SystemDomain)
                    'leaverSender'        = ''
                }   
                if ($Test.IsPresent) {
                    Process-Emailing @props -test
                }
                else {
                    Process-Emailing @props
                }
                #endregion             
            }
            elseif ( $currentADUsers | Where-Object { $checkEID -eq $_.EmployeeID } ) {
                <# Do not creat new user also, if the employee ID is taken #>
                $offendingUser = $currentADUsers | Where-Object { $_.EmployeeID -eq $checkEID } 
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog -Color White -BackGroundColor Red "[$timer] [w]  EmployeeID [$checkEID]  already in use by [$($offendingUser.SAMAccountName) / $($offendingUser.EmployeeID)] skipping user creation - please contact HR and re-submit new starter with a unique Employee ID!"
                Add-Member -InputObject $line -NotePropertyName 'IsProcessed' -NotePropertyValue 'No' -Force
                #region EMAIL ALERT
                $props = @{
                    'newStarterFullName'  = $checkName -replace '\.', ' '
                    'mode'                = 'genericOnboardingFailure'
                    'domain'              = $domain
                    'smtpServer'          = $config.SMTPServer
                    'testRecipients'      = Get-Content $config.TestRecipientList
                    'recipients'          = Get-Content $config.RecipientList
                    'departmentSignature' = $config.SenderSignature
                    'templateName'        = $line.CopyAccount
                    'starterSender'       = ($config.StarterSender + '@' + $config.SystemDomain)
                    'leaverSender'        = ''
                }   
                if ($Test.IsPresent) {
                    Process-Emailing @props -test
                }
                else {
                    Process-Emailing @props
                }
                #endregion                  
            }
            #endregion
            <# * IF THE INPUT SEEMS TO BE GENUINE - NOT MATCHING TO AN EXISTING ACCOUNT - THE NEXT CHECK COMMENCES. HERE THE SCRIPT CHECKS, IF THE SUPPLIED TEMPLATE ACCOUNT IS NOT A PROCESSED LEAVER. THIS IS BECAUSE LEAVERS - PROCESSED BY THE USER PROCESSING - ARE REMOVED FROM ALL AD GROUPS AND THEIR ORGANIZATIONAL DETAILS ARE PURGED, HENCE CLONING THEM WOULD NOT CREATE A PROPER  USER ACCOUNT#>
            else {    
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog -Color White -BackGroundColor Black "[$timer] [i]  Processing user   [$($line.FirstName) $($line.LastName)]  (name is subject to change) " -LinesBefore 1
                $templateName = $($line.CopyAccount -replace ' ', '')
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' We will use this user as a template:  ', "[$templateName]" -Color White, Green, White, Green
                $templateUser = $currentADUsers | Where-Object { $_.SAMAccountName -like "$($line.CopyAccount)" } # this account will be copyed to create the new user 
                if ($templateUser) {
                    $userDomain = ($templateUser.UserPrincipalName).Split('\@')[1]
                    [string]$templateOU = Get-ParentOU $templateUser
                    Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' New users account will be created in OU:', " [$($templateOU)] " -Color White, Green, White, Green
                }
                <# * THIS IS A FAILURE, IF THE TEMPLATE IS IN THE LEAVER OU #>
                #region TEMPLATE IS A LEAVER
                if (    ($templateOU -match $config.LeaverOU) ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog -Color White -BackGroundColor Red "[$timer] [w]  Defined template user [$($line.CopyAccount)] is a leaver. Execution cancelled, user [ $checkName ] was not created - please review and provide a non-leaver template!"   
                    Add-Member -InputObject $line -NotePropertyName 'IsProcessed' -NotePropertyValue 'No' -Force
                    #$timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Defined template user is a leaver. Execution cancelled. ', " [$($line.CopyAccount)] " -Color White, Yellow, White, Yellow
                    #region EMAIL ALERT
                    $props = @{
                        'newStarterFullName'  = $checkName -replace '\.', ' '
                        'mode'                = 'templateOrManagerError'
                        'domain'              = $domain
                        'smtpServer'          = $config.SMTPServer
                        'testRecipients'      = Get-Content $config.TestRecipientList
                        'recipients'          = Get-Content $config.RecipientList
                        'departmentSignature' = $config.SenderSignature
                        'templateName'        = $line.CopyAccount
                        'starterSender'       = ($config.StarterSender + '@' + $config.SystemDomain)
                        'leaverSender'        = ''
                        'manager'             = $line.manager
                    }   
                    if ($Test.IsPresent) {
                        Process-Emailing @props -test
                    }
                    else {
                        Process-Emailing @props
                    }
                    #endregion
                } 
                #endregion
                <# * THIS IS A FAILURE IF THE TEMPLATE CAN NOT BE FOUND IN THE DOMAIN #>
                #region TEMPLATE NOT FOUND
                elseif ( !($templateOU) -or !($currentADUsers | Where-Object { $_.SamAccountName -eq $line.CopyAccount }) -or !($currentADUsers | Where-Object { $_.SamAccountName -match $line.manager }) ) {
                    # catch nonexistent account
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog -Color White -BackGroundColor Red "[$timer] [w]  Defined template user [$($line.CopyAccount)] or manager [$($line.Manager)] not found in the user database. Execution cancelled - user [ $checkName ] was not created  - please review and provide an existing user account for template!"  
                    #$timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' We could not find the template user', " [$($line.CopyAccount)] ", ' in the domain ' -Color White, Yellow, White, Yellow
                    Add-Member -InputObject $line -NotePropertyName 'IsProcessed' -NotePropertyValue 'No' -Force
                    #region EMAIL ALERT
                    $props = @{
                        'newStarterFullName'  = $checkName -replace '\.', ' '
                        'mode'                = 'templateOrManagerError'
                        'domain'              = $domain
                        'smtpServer'          = $config.SMTPServer
                        'testRecipients'      = Get-Content $config.TestRecipientList
                        'recipients'          = Get-Content $config.RecipientList
                        'departmentSignature' = $config.SenderSignature
                        'templateName'        = $line.CopyAccount
                        'starterSender'       = ($config.StarterSender + '@' + $config.SystemDomain)
                        'leaverSender'        = ''
                        'manager'             = $line.manager
                    }
                    if ($Test.IsPresent) {
                        Process-Emailing @props -test
                    }
                    else {
                        Process-Emailing @props
                    }
                    #endregion
                }
                #endregion
                <# * IF THE INPUT LINE PASSED ALL CHECKS, THE ACTUAL PROCESSING BEGINS#>  
                #region  CREATION       
                else {
                    $stopWatch = [system.diagnostics.stopwatch]::StartNew()
                    Add-Member -InputObject $line -NotePropertyName 'IsProcessed' -NotePropertyValue 'Yes' -Force
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Template account ', " [$($line.CopyAccount)] ", ' is usable ' -Color White, Green, White, Green, White
                    # parameters to be sent to the new user generator function
                    #region DEFINE NEW AD USER PROPERTIES
                    $params = @{ 
                        firstName          = $($line.FirstName -replace ' ', '')
                        lastName           = $($line.LastName -replace ' ', '')
                        allADUsers         = $currentADUsers
                        userDomain         = $userDomain
                        eID                = $($line.EmployeeID -replace ' ', '')
                        EOTargetDomain     = $($config.EOTargetDomain -replace ' ', '')
                        systemDomain       = $config.SystemDomain
                        templateUser       = $templateUser
                        templateOU         = $templateOU
                        contractType       = $line.ContractType
                        endDate            = $line.EndDate
                        startDate          = $line.StartDate
                        holdiayEntitlement = $line.HolidayEntitlement
                        manager            = $($line.Manager -replace ' ', '')
                        jobTitle           = $line.JobTitle
                    }  
                    # new user generator function              
                    $userConfiguration = Get-UserDetails @params 
                    # returned object split to categories
                    $basicProps = $userConfiguration.basicdetails  
                    $extensionProps = $userConfiguration.extensionDetails
                    $settingsProps = $userConfiguration.settingDetails      
                    $emailProps = $userConfiguration.emailDetails
                    #endregion
                    <# * THE FIRST THING WE DO IS CREATING AND CONFIGURING THE NEW AD OBJECT USING THE PROVIDED TEMPLATE ACCOUNT AND THE DETAILS SUPPLIED WITHIN THE INPUT FILE #>
                    #region ACTIVE DIRECTORY
                    try {
                        # create new AD user object
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Creating new user ', " [$($basicProps.Name)] ", ' based on the provided parameters ' -Color White, green, White, Green, White
                        New-ADUser @basicProps -Server $DC -Credential $AD_Credential 
                        # wait here to ensure AD user is present
                        do {
                            $delay = 30
                            $checkUser = $null
                            $checkUser = Get-ADUser $($basicProps.SAMAccountName) -Server $DC -Credential $AD_Credential
                            if (-not $checkUser) {
                                Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Waiting for  new user ', " [$($basicProps.Name)] ", ' to appear. Next check is in ', " [$delay] " , ' seconds' -Color White, Yellow, White, Yellow, White, Yellow, White
                                Start-Sleep -Seconds $delay
                            }
                        } until ($checkUser)
                    }
                    catch {
                        # error break out
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Creating the new user ', " [$($basicProps.Name)] ", ' failed. User NOT created, proceeding to the next account! ' -Color White, Red, White, Red, Red
                        "[$timer] [ERROR] Creation of account [$($basicProps.Name)] failed! Error details: " | Out-File $ErrorFile -Append
                        Add-Member -InputObject $line -NotePropertyName 'IsProcessed' -NotePropertyValue 'No' -Force
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                        [void] (Stop-Transcript -ErrorAction Ignore)
                        Break              
                    }
                    try {
                        # set up the AD user with some properties
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Configuring user extension atributes and Manager for user ', " [$($basicProps.Name)] ", ' based on the provided details ' -Color White, green, White, Green, White
                        # add manager
                        Set-ADUser -Identity $($basicProps.SAMAccountName) -Add $extensionProps -Manager $settingsProps.manager -Server $DC -Credential $AD_Credential #-Verbose
                        # add expiry if defined
                        if ($settingsProps.adAccountExpiration) {
                            Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding expration date to user', " [$($basicProps.Name)] ", ' account expires: ', "[$($settingsProps.adAccountExpiration)]" -Color White, green, White, Green, White, Green
                            Set-ADAccountExpiration -Identity $($basicProps.SAMAccountName) -DateTime $($settingsProps.adAccountExpiration) -Server $DC -Credential $AD_Credential
                        }
                        # update job title
                        if ($settingsProps.jobTitle) {
                            Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding job title to user', " [$($basicProps.Name)] ", ' job title: ', "[$($settingsProps.jobTitle)]" -Color White, green, White, Green, White, Green
                            Set-ADUser -Identity $($basicProps.SAMAccountName) -Title $($settingsProps.jobTitle) -Server $DC -Credential $AD_Credential
                        }                        
                    }
                    catch {
                        # error non-fatal
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Configuring new user ', "[$($basicProps.Name)]", ' failed, ', ' manual adjustment will be needed, exceution continues!' -Color White, Red, White, Red, Red, White
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append                    
                        Continue
                    }
                    try {
                        # add AD user to the groups of the template (except the licensing groups as per dicision)
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding new user ', " [$($basicProps.Name)] ", ' to the groups of the template account. ' -Color White, green, White, Green, White
                        $templateUser.Memberof | Where-Object { $_ -notmatch 'CN=LICENSE' } | ForEach-Object { Add-ADGroupMember $_ $($basicProps.SAMAccountName) -Server $DC -Credential $AD_Credential }
                    }
                    catch {
                        # error non-fatal
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Adding new user ', " [$($basicProps.Name)] ", ' to the groups of the template account ', 'failed,', ' manual adjustment will be needed, exceution continues!' -Color White, Red, White, Red, White, Red, White
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append    
                        Continue
                    }
                    if ($config.ProcessSpecial -notmatch 'Yes') {
                        try {
                            $personalHomeDrive = $config.HomeDriveLetter
                            $personalHomeDriveFolder = '\\' + $config.systemDomain + '\' + $config.HomeDrviePath + '\' + $basicProps.SamAccountName
                            # add Home drive to the user's profile 
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ' , ' Adding home drive ' , "[$personalHomeDrive]" , ' to new user ' , "[$($basicProps.SamAccountName)]" , ' path will be: ' , " [$personalHomeDrive / $personalHomeDriveFolder] " -Color White, green, White, Green, White, Green , White, Green
                            # $templateUser.Memberof | Where-Object { $_ -notmatch 'CN=LICENSE' } | ForEach-Object { Add-ADGroupMember $_ $($basicProps.SAMAccountName) -Server $DC -Credential $AD_Credential }
                            Set-ADUser -Identity $basicProps.SamAccountName -HomeDirectory $personalHomeDriveFolder -HomeDrive $personalHomeDrive -Server $DC -Credential $AD_Credential
                        }
                        catch {
                            # error non-fatal
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Setting home drive  ', " [ $personalHomeDrive / $personalHomeDriveFolder] ", 'failed,', ' manual adjustment will be needed, exceution continues!' -Color White, Red, White, Red, White, Red
                            Write-Warning "[$timer] $($_.Exception.Message)" 
                            "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append    
                            Continue
                        }
                    }
                    try {
                        # sync changes cross-domain
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Synchronising domain controllers of domain ', " [$domain] " -Color White, Green, White, Magenta
                        Sync-ActiveDirectory -server $DC -credential $AD_Credential -Verbose
                    }
                    catch {
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Syncronisation of domain controllers failed for domain ', "[$domain]", ' As this prevents the rest of the process, we proceed to the next user!' -Color White, Red, White, Magenta, White
                        "[$timer] [ERROR] Creation of account [$($basicProps.Name)] failed! Error details: " | Out-File $ErrorFile -Append
                        Add-Member -InputObject $line -NotePropertyName 'IsProcessed' -NotePropertyValue 'No' -Force
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                        [void] (Stop-Transcript -ErrorAction Ignore)
                        Break  
                    }
                    #endregion
                    <# * THE NEXT STEP IN THE PROCESS IS TO LICENSE THE NEWLY CREATED - BY SYNCING THE AD OBJECT TO AAD - O365 USER. 
                    * - IF THE TEMPLATE HAD E3 LICENSE OR HIGHER, THE NEW USER IS BEING ADDED TO THE E3 OFFICE 365 GROUP (SERVICE DESK CAN MANUALY UPGRADE THIS LATER TO E5 OR HIGHER IF NECCESARY BY MOVING THE USER TO THE E5 GROUP INSTEAD)
                    * - OTHERWISE THE NEW USER WILL GET AN F1 LICENSE #>
                    #region LICENSING
                    try {
                        #connect  MSOL service
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Connecting to Microsoft Online ' -Color White, Green, White
                        Connect-MsolService -Credential $AAD_Credential 
                        # licensing
                        # if the template was licensed...
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Collecting template users ', "[$($templateUser.UserPrincipalName)]" , ' licenses ' -Color White, Green, White, Green, White
                        $LicenseSKUs = ((Get-MsolUser -UserPrincipalName $templateUser.UserPrincipalName -ErrorAction SilentlyContinue).Licenses).AccountSkuid 
                        # ...license with F1 if the template had F1 license
                        if (($LicenseSKUs) -and ($LicenseSKUs -match 'DESKLESSPACK') ) {
                            Add-ADGroupMember 'LICENSE-Office_365_F3_F1' -Members $($basicProps.SamAccountName) -Server $DC -Credential $AD_Credential
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding license ', ' Office F3 ', ' to user ', " [$($basicProps.UserPrincipalName)] " -Color White, Green, White, Green, White, Green
                        }
                        # ... or give E3 with any other template that was licenses
                        elseif ($LicenseSKUs) {
                            Add-ADGroupMember 'LICENSE-Office_365_E3' -Members $($basicProps.SamAccountName) -Server $DC -Credential $AD_Credential
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding license ', ' Office E3 ', ' to user ', " [$($basicProps.UserPrincipalName)] " -Color White, Green, White, Green, White, Green     
                        }
                        else {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Template unlicensed, skipping licensing of user', " [$($basicProps.UserPrincipalName)] " -Color White, Yellow, White, Yellow
                        }      
                        # wait for the new account to sync to the cloud
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Ensureing the useraccount synced to Azure ' -Color White, Green, White
                        $props = @{
                            'UPN'          = $basicProps.UserPrincipalName
                            'systemDomain' = $config.SystemDomain
                        }
                        if ($config.ProcessSpecial -match 'Yes') {
                            Wait-ForMSOLAccount @props -delay 120
                        }
                        elseif ($config.ProcessSpecial -match 'No') {
                            Wait-ForMSOLAccount @props -delay 180
                        }                    
                        # set usage location
                        $usageLocation = (Get-MsolUser -UserPrincipalName $templateUser.UserPrincipalName).UsageLocation 
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Setting usage location for user ', " [$($basicProps.UserPrincipalName)] ", ' to ', " [$usageLocation] " -Color White, Green, White, Green, White, Green
                        [void] (Set-MsolUser -UserPrincipalName $($basicProps.UserPrincipalName) -UsageLocation $usageLocation -ErrorAction Continue)
                    }
                    catch {
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Licensing for failed for user ', " [$($basicProps.Name)] ", ' please review manually and assign licenses using licensing groups! ' -Color White, Red, White, Red, White
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                        # failing to configure is non-fatal
                        Continue
                    }                
                    #endregion
                    <# * ONCE THE AD OBJECT IS CREATED, WE ARE CONFIGURING A MAILBOX FOR THIS USER. THE ENVIRONMENT - EXCHANGE ONLINE OR ON-PREM EXHANGE - IS DETERMINED BY THE LOCATION OF THE TEMPLATE'S MAILBOX #>
                    #region MAILBOX
                    Try { 
                        #region connect
                        # (remove pottential stalled sessions)
                        Get-PSSession | Remove-PSSession
                        # if the domain is slow to replicate, wait here while that happens
                        if ($config.ProcessSpecial -match 'No') {
                            Start-Sleep -Seconds 1200 # allow account to sync
                            # * WAIT FOR THE AD OBJECT
                        }
                        do {
                            $adObjectFound = $false
                            $SAM = $basicProps.SAMAccountName
                            $adObjectFound = Get-ADUser -Filter { SamAccountName -eq $SAM } -ErrorAction silentlycontinue -Server $DC -Credential $AD_Credential
                            if (!$adObjectFound) {
                                Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Waiting for  new user ', " [$($basicProps.Name)] ", ' to appear. Next check is in ', " [$delay] " , ' seconds' -Color White, Yellow, White, Yellow, White, Yellow, White
                                Start-Sleep -Seconds $delay
                            }
                        } until ($adObjectFound)
                        Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Found user object', " [$($basicProps.Name)] ", ' in Active Directory, continuing ' -Color White, Green, White, Green, White
                        #}                        
                        # connect to exchange                    
                        # connect to Exchange ON-PREM
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Connecting to on-prem exchange ' -Color White, Green, White
                        Connect-OnPremExchange -credential $Exch_Credential -exchangeServer ($config.ExchangeServer + '.' + $config.SystemDomain)   
                        # connect to Exchange ONLINE (default prefix: "o")
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Connecting to online exchange (commands will be prefixed)' -Color White, Green, White
                        Connect-OnlineExchange -credential $AAD_Credential 
                        #endregion
                        ## create online mailbox, if the template user's mailbox target points to o365
                        # * Updated based on the instructions of Martin 15/04/2021 
                        # if ( $templateUser.TargetAddress -match 'onmicrosoft.com' ) { 
                        # enable remote mailbox on-prem
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Provisioning online mailbox for user ', "[$($basicProps.Name)] " -Color White, Green, White, Green
                        Create-ONLINEMailbox -remoteRoutingAddress $($emailProps.remoteRoutingAddress) -samAccountName $($basicProps.SAMAccountName) 
                        # sync changes to AZURE AD
                        Sync-AzureActiveDirectory -server $($config.AADSyncServer + '.' + $config.SystemDomain) -credential $AD_Credential
                        # GET (store) exchange GUID from the online exchange (we only connect to o365 for this part in the new user creation process)                   
                        $exchangeGUID = Sync-ExchangeGUIDs -UPN $($basicProps.UserPrincipalName) -mode 'GET'
                        # -->
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Syncing Exchange GUID ', " [$exchangeGUID] ", ' to user account ', "[$($basicProps.Name)] " -Color White, Green, White, Green, White, Green
                        # <--
                        # PUT (apply) the exchange GUID onto the mailbox via on-prem exchange
                        Sync-ExchangeGUIDs -UPN $($basicProps.UserPrincipalName) -mode 'PUT' -GUID $exchangeGUID
                        #  }
                        # * Updated based on the instructions of Martin 15/04/2021
                        # ## create on-prem mailbox
                        # else {
                        #     # mail-enable the user on-prem
                        #     $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Provisioning on-prem mailbox for user ', "[$($basicProps.Name)] " -Color White, Green, White, Green
                        #     Create-ONPremMailbox -targetDomain $config.EOTargetDomain -samAccountName $($basicProps.SAMAccountName) 
                        # }
                        ## finally drop all exchange sessions
                        Get-PSSession | Remove-PSSession
                    }
                    catch {
                        # error breakout
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Provisioning the mailbox for user ', " [$($basicProps.Name)] ", 'failed. As this would prevent the user to work, we are stopping this creation and move to the next.' -Color White, Red, White, Red, Red
                        "[$timer] [ERROR] Creation of account [$($basicProps.Name)] failed! Error details: " | Out-File $ErrorFile -Append
                        Add-Member -InputObject $line -NotePropertyName 'IsProcessed' -NotePropertyValue 'No' -Force
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append 
                        [void] (Stop-Transcript -ErrorAction Ignore)
                        Break   
                    }
                    # add SMTP addresses to the mailbox
                    try {
                        # primary SMTP
                        if ($userDomain -notmatch ($config.SystemDomain)) {
                            # detect non-UK user
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Non-UK user detected. Updating SMTP address allocation: ', "[$($emailProps.oldSMTPAddress) --> $($emailProps.primarySMTPAddress)]" -Color White, Yellow, White, Yellow
                            Set-ADUser -Identity $($basicProps.SAMAccountName) -Remove @{ProxyAddresses = $emailProps.oldSMTPAddress } -Server $DC -Credential $AD_Credential
                            Set-ADUser -Identity $($basicProps.SAMAccountName) -Add @{ProxyAddresses = $emailProps.primarySMTPAddress } -Server $DC -Credential $AD_Credential
                        }
                        else {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' UK user, leaving mailbox with the default primary SMTP address ', "[$($emailProps.primarySMTPAddress)]" -Color White, Green, White, Green
                            # fix, if the new account does not get the proxy address
                            if ((Get-ADUser -Identity $($basicProps.SAMAccountName) -Properties ProxyAddresses -Server $DC -Credential $AD_Credential) -notmatch $emailProps.primarySMTPAddress ) {
                                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Primary SMTP was missing, adding this: ', "[$($emailProps.primarySMTPAddress)]" -Color White, Green, White, Green
                                Set-ADUser -Identity $($basicProps.SAMAccountName) -Add @{ProxyAddresses = $emailProps.primarySMTPAddress } -Server $DC -Credential $AD_Credential
                            }
                        }                    
                        # secondary SMTP
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding secondary SMTP address: ', "[$($emailProps.secondarySMTPAddress)]" -Color White, Green, White, Green
                        Set-ADUser -Identity $($basicProps.SAMAccountName) -Add @{ProxyAddresses = $emailProps.secondarySMTPAddress } -Server $DC -Credential $AD_Credential
                        # terciary SMTP
                        if ($emailProps.terciarySMTPAddress) {
                            # only for non-UK user
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Non-UK user detected. Adding original smtp (.co.uk) as an alternative third address: ', "[$($emailProps.terciarySMTPAddress) (should match: $($emailProps.oldSMTPAddress))]" -Color White, Yellow, White, Yellow
                            Set-ADUser -Identity $($basicProps.SAMAccountName) -Add @{ProxyAddresses = $emailProps.terciarySMTPAddress } -Server $DC -Credential $AD_Credential
                        }                       
                    }
                    catch {
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Email details update failed for user ', " [$($basicProps.Name)] ", ' please review manually ' -Color White, Red, White, Red, White
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                        # failing to configure is non-fatal
                        Continue
                    }
                    try {
                        # sync changes cross-domain
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Synchronising domain controllers of domain ', " [$domain] " -Color White, Green, White, Magenta
                        Sync-ActiveDirectory -server $DC -credential $AD_Credential -Verbose
                        $delay = 60
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Break for ', " [$delay] ", ' seconds to allow sync to finish ' -Color White, Green, White, green, White
                        Start-Sleep -Seconds $delay
                    }
                    catch {
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Syncronisation of domain controllers failed for domain ', "[$domain]", ' As this prevents the rest of the process, we proceed to the next user!' -Color White, Red, White, Magenta, White
                        "[$timer] [ERROR] Creation of account [$($basicProps.Name)] failed! Error details: " | Out-File $ErrorFile -Append
                        Add-Member -InputObject $line -NotePropertyName 'IsProcessed' -NotePropertyValue 'No' -Force
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append  
                        [void] (Stop-Transcript -ErrorAction Ignore)
                        Break  
                    }
                    #endregion
                    <# * IF THE DOMAIN OF THE USER IS SET TO USING THIS, THE SCRIPT CREATES DFS TARGET PATHS FOR THE USER ACCOUNT#>
                    #region DFS
                    try {
                        if ($config.ProcessSpecial -match 'Yes') {
                            # create DFS targets for the new account
                            $dfsProps = @{                  
                                'PeopleDFS'         = ('\\' + $config.SystemDomain + '\PEOPLE\' + $basicProps.SAMAccountName )
                                'PeopleTargetPath'  = ('\\' + $config.PeopleFileServer + '.' + $config.SystemDomain + '\PEOPLE$\' + $basicProps.SAMAccountName)
                                'ProfileDFS'        = ('\\' + $config.SystemDomain + '\PROFILES\' + $basicProps.SAMAccountName )
                                'ProfileTargetPath' = ('\\' + $config.ProfileFileServer + '.' + $config.SystemDomain + '\PROFILES$\' + $basicProps.SAMAccountName)
                            }
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Setting up DFS for the user ', " [$($basicProps.Name)] ", ' in domain ' , " [$domain] " -Color White, Green, White, Green, White, Magenta
                            Create-DFS @dfsProps
                        }
                        else {
                            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Not setting up DFS, as it is not implemented in domain ', " [$domain] " -Color White, Yellow, White, Magenta
                        }
                    }
                    catch {
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' DFS setup failed for user ', " [$($basicProps.Name)] ", ' please review manually ' -Color White, Red, White, Red, White
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                        # failing to configure is non-fatal
                        Continue
                    } 
                    #endregion
                    <# * ONCE THE NEW USER IS CREATED, THE SCRIPT ADDS THIS NEW ACCOUNT TO THE CACHED DOMAIN USERS.THIS IS IMPORTANT, AS IF THIS WOULD NOT HAPPEN, HUMAN ERRORS  - SUCH AS ENTERING A USER ACCOUNT TWO TIMES TO THE IMPUT FILE - WOULD RESULT DOUBLE USER CREATION #>
                    #region UPDATE USERBASE
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Updating userbase with ', " [$($basicProps.SamAccountName)] ", ' user ' -Color White, Green, White, Green, White
                    $currentADUsers += Get-ADUser -Identity $basicProps.SamAccountName -Properties * -Server $DC -Credential $AD_Credential
                    #endregion
                    <# * EACH USER'S CREATION TRIGGERS AN EMAIL CONTAINING THE INITIAL PASSWORD OF THE NEW USER ALONG SOME IMPORTANT DETAILS. THESE ARE BEING SENT TO THE MANAGER OF THE USER SUPPLIED VIA THE INPUT FILE#>
                    #region EMAIL TO MANAGER
                    try {
                        #TODO: Is manager having password and username in the same email accpetable?
                        # saving report of individual user creation
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Exporting user details of', " [$($basicProps.Name)] ", ' into report file ' , " [$reportFile] " -Color White, Green, White, Green, White, Cyan
                        $report = [pscustomobject] $basicProps | Select-Object -Property * -ExcludeProperty 'AccountPassword', 'Description' 
                        $report | Export-Csv $reportFile -Append -Force -NoTypeInformation  
                        #region EMAIL TO LINE MANAGER
                        # French version
                        if ($userDomain -match '.fr') {
                            $props = @{
                                'mode'                = 'toFrenchManager'
                                'domain'              = $domain
                                'smtpServer'          = $config.SMTPServer
                                'testRecipients'      = Get-Content $config.TestRecipientList
                                'recipients'          = ($currentADUsers | Where-Object { $_.UserPrincipalName -match $settingsProps.manager }).UserPrincipalName
                                'departmentSignature' = $config.SenderSignature
                                'templateName'        = $line.CopyAccount
                                'starterSender'       = ($config.StarterSender + '@' + $config.SystemDomain)
                                'leaverSender'        = ''
                                'computerUsagePolicy' = $config.ComputerUsagePolicy
                                'report'              = $report # report per user
                                'serviceDeskLink'     = $config.ServiceDeskLink
                                'userPWD'             = $([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($basicProps.AccountPassword )))
                                'userName'            = $basicProps.Name
                            }
                        }
                        # standard (English) version
                        else {
                            $props = @{
                                'mode'                = 'toManager'
                                'domain'              = $domain
                                'smtpServer'          = $config.SMTPServer
                                'testRecipients'      = Get-Content $config.TestRecipientList
                                'recipients'          = ($currentADUsers | Where-Object { $_.UserPrincipalName -match $settingsProps.manager }).UserPrincipalName
                                'departmentSignature' = $config.SenderSignature
                                'templateName'        = $line.CopyAccount
                                'starterSender'       = ($config.StarterSender + '@' + $config.SystemDomain)
                                'leaverSender'        = ''
                                'computerUsagePolicy' = $config.ComputerUsagePolicy
                                'report'              = $report # report per user
                                'serviceDeskLink'     = $config.ServiceDeskLink
                                'userPWD'             = $([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($basicProps.AccountPassword )))
                                'userName'            = $basicProps.Name
                            }
                        }
                        if ($Test.IsPresent) {
                            if ($config.ProcessSpecial -match 'Yes') {
                                Process-Emailing @props -test -addServiceDeskLink
                            }
                            else {
                                Process-Emailing @props -test
                            }
                        }
                        else {
                            if ($config.ProcessSpecial -match 'Yes') {
                                Process-Emailing @props -addServiceDeskLink
                            }
                            else {
                                Process-Emailing @props
                            }
                        }
                        #endregion
                    }
                    catch {
                        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to send report to line manager ' -Color White, Red, White
                        Write-Warning "[$timer] $($_.Exception.Message)" 
                        "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
                        # failing to configure is non-fatal
                        Continue
                    }        
                    #endregion
                    $stopWatch.stop()
                    Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Creating the user ', " [$($basicprops.Name)] ", ' took ', " [$([math]::Round(($stopWatch.Elapsed.TotalMinutes), 2))] ", ' minutes ' -Color White, Magenta, White, Magenta, White, Magenta, White
                    Add-Member -InputObject $line -NotePropertyName 'ProcessingTime(min)' -NotePropertyValue $([math]::Round(($stopWatch.Elapsed.TotalMinutes), 2)) -Force
                    # ADD SUMMARY TO EMAIL ONLY IF CREATION SUCCEEDED
                    $summaryReport += $report
                }
                #endregion
            }
        }
        <# * NEXT WE START PRODUCING REPORTS OF THE WORK DONE. THE SUPPORT TEAM RECEIVES AN EMAIL WITH LINKS TO THE PROCESSING LOGS, SUCCESFULLY PROCESSED USER LIST, ERROR LOGS AND THE BACKUP FILES #>
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Saving processed input into report file ' , " [$backupFile] " -Color White, Green, White, Cyan
        if ($csvImport) {
            $csvImport | Export-Csv $backupFile -Append -NoTypeInformation -Force
        }
        #region EMAIL TO SUPPORT
        try {
            if ($csvImport -and $report) {
                #region EMAIL SUMMARY TO SERVICEDESK
                $props = @{
                    'mode'                = 'toSupportCreation'
                    'domain'              = $domain
                    'smtpServer'          = $config.SMTPServer
                    'testRecipients'      = Get-Content $config.TestRecipientList
                    'recipients'          = Get-Content $config.RecipientList
                    'departmentSignature' = $config.SenderSignature
                    'templateName'        = $line.CopyAccount
                    'starterSender'       = ($config.StarterSender + '@' + $config.SystemDomain)
                    'leaverSender'        = ''
                    'computerUsagePolicy' = $config.ComputerUsagePolicy
                    'report'              = $summaryReport
                    'serviceDeskLink'     = $config.ServiceDeskLink
                    'userPWD'             = $([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($basicProps.AccountPassword )))
                    'link1'               = ($config.OnBoardingOutputFolder + '\' + $sDate + '\' + $logTime + '_report.csv') # $reportFile #
                    'link2'               = ($config.LogsFolder + '\' + $sDate + '\' + $logTime + '_onboarding_error.log') # $errorFile # 
                    'link3'               = ($config.OnBoardingOutputFolder + '\' + $sDate + '\' + $logTime + '_onboarding_backup.csv') # $backupFile # 
                    'link4'               = ($config.LogsFolder + '\' + $sDate + '\' + $logTime + '_onboarding_transcript.log') # $transcriptFile #
                    'link5'               = ($config.LogsFolder + '\' + $sDate + '\' + $logTime + '_onboarding_action.log') # $actionLog # 
                    'Attachments'         = @(
                        $($config.OnBoardingOutputFolder + '\' + $sDate + '\' + $logTime + '_report.csv')
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
        }
        catch {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Failed to send report to  ', " [$r] " -Color White, Red, White, Red
            Write-Warning "[$timer] $($_.Exception.Message)" 
            "[$timer] $($_.Exception.Message)" | Out-File $ErrorFile -Append
            # failing to configure is non-fatal
            Continue
        }   
        #endregion
    }
}
END {
    <# * FINALLY THE SCRIPT DOES A CLEANUP, REMOVING THE STORED USER CACHE, RENAMING / REMOVING THE INPUT FILES, ETC. ALSO IT ENDS THE TRANSCRIPT AND SENDS ERRORS TO THE ERROR REPORT FILE #>
    # drop user cache
    $currentADUsers = $null
    # # drop input files
    # Get-ChildItem -Path $config.OnBoardingInputFolder | Where-Object { $_.Name -like '*.csv' } | Remove-Item -WhatIf
    # rename input files
    if (-not $test.IsPresent) {
        $CSVs = $fileList # Get-ChildItem -Path $config.OnBoardingInputFolder | Where-Object { $_.Name -like '*.csv' }
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
}
# SIG # Begin signature block
# MIIOWAYJKoZIhvcNAQcCoIIOSTCCDkUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU4H9fXeYTUZmCvaitVM+HclrV
# ufagggueMIIEnjCCA4agAwIBAgITTwAAAAb2JFytK6ojaAABAAAABjANBgkqhkiG
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUnbi/vBG5aXla5M/3QFD3vCmtoCwwDQYJKoZI
# hvcNAQEBBQAEggEA6xBU+lvadf6rgLc4XNnFi01+mP5G2uhTPmc1fZ7xc1DHTLgw
# xhlayo53TLw/+kiyolR1D3cziGSPDQ2xKWcQ5g7a8E0qzimycXFsuESxru6Gqky/
# RhDTUxQnh2ksnvjKNr8DIYVjSbX6xSX5B2rGBdr0MZWMR5JHNyKLPNOezfdBuQcU
# iZYkdimGzJeHM8rscpu5/O7MOadZ2dlM0G9+Rv5Z0G97liqAT6Cc9o/u4Ay9NtgU
# QtDMggoGjqlKO9zXNjrkxs/k36YC5epYZpzIAnhP09dycvNFPY+New4NiCoKLM6r
# IhfWgRKANVccrp9oQl9XoMs8iF8NmMZo/d6tiQ==
# SIG # End signature block
