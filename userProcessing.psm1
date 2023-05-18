function New-Credential {
    [CmdletBinding()]
    param (
        [string] $outFile = $(Throw '-outFile is required'),
        [string] $userName = $(Throw '-userName is required')
    )
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)
    if (-not (Test-Path $outFile)) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)  ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Creating new credential for account  ', " [$userName] " -Color White, Green, White, Green
        Get-Credential -UserName $userName -Message "Please enter password for the [$userName] account" | Export-Clixml -Path $outFile -Force
    }
    $cred = Import-Clixml -Path $outFile
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Imported credentials for user  ', " [$userName] ", ' from file ', "[$outFile]" -Color White, Green, White, Green, White, Green
    return $cred
}
function Import-Workfiles {
    [CmdletBinding()]
    param (
        [string] $inputFolder = $(Throw '-inputFolder is required')
    )
    $import = @()
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Importing contents of workfile from the input folder ', " [$inputFolder] ", ' (*.csv-s only) ' -Color White, Green, White, Green, White
    $inputfiles = (Get-ChildItem $inputFolder | Where-Object { $_.Name -like '*.csv' }).FullName
    foreach ($i in $inputfiles) {
        $import += Import-Csv $i
    }
    return $import
}
Function Get-FileName($initialDirectory) {  
    <#
https://devblogs.microsoft.com/scripting/hey-scripting-guy-can-i-open-a-file-dialog-box-with-windows-powershell/
#>
    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') |
        Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = 'CSV (*.csv)| *.csv'
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
} #end function Get-FileName
function Get-UNIQUEValue {
    [cmdletbinding()]
    param(
        [Parameter (Mandatory)] [ValidateSet('UPN', 'SAM', 'eID', 'RRA', 'SSMTP')] [string] $mode,
        [string] $value = $(Throw '-value is required'),
        [object] $allADUsers = $(Throw '-allADUsers is required')
    ) 
    $x = 1
    do {
        $adCheck = $newValue = $null
        if ($mode -match 'UPN') {
            $insert = [string]$x + '@'
            $newValue = $value -replace '@', $insert
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Cross-checking value of UPN ' , " [$newValue] ", ' with the current user base ' -Color White, Green, White, Yellow
            $adCheck = $allADUsers | Where-Object { $_.UserPrincipalName -match $newValue }
        }
        elseif ($mode -match 'SAM') {
            $insert = [string]$x
            if ($value.Length -gt 19) {
                $newValue = $value.substring(0, 19) + $insert #length of the new sAM shoulc be under 21 characters
            }
            else {
                $newValue = $value + $insert
            } 
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Cross-checking value of SAM ' , " [$newValue] ", ' with the current user base ' -Color White, Green, White, Yellow         
            $adCheck = $allADUsers | Where-Object { $_.SAMAccountName -match $newValue }
        }
        # elseif ($mode -match 'eID') {
        #     $insert = [string]$x
        #     $newValue = $value + $insert
        #     $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Cross-checking value of EmployeeID ' , " [$newValue] ", ' with the current user base ' -Color White, Green, White, Yellow
        #     $adCheck = $allADUsers | Where-Object { $_.employeeid -match $newValue }
        # }
        elseif ($mode -match 'RRA') {
            $insert = [string]$x + '@'
            $newValue = $value -replace '@', $insert
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Cross-checking value of the remote routing address' , " [$newValue] ", ' with the current user base ' -Color White, Green, White, Yellow
            $adCheck = $allADUsers | Where-Object { $_.proxyAddresses -match $newValue }
        }
        elseif ($mode -match 'SSMTP') {
            $insert = [string]$x + '@'
            $newValue = $value -replace '@', $insert
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Cross-checking value of the secondary smtp address ' , " [$newValue] ", ' with the current user base ' -Color White, Green, White, Yellow
            $adCheck = $allADUsers | Where-Object { $_.proxyAddresses -match $newValue }
        }
        $x++
    } while ($adCheck)
    #return $newValue
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' New value obtained: ' , " [$newValue] " -Color White, Green, White, Green
    return $newValue
}
function Get-UserDetails {
    [CmdletBinding()]
    param (
        [string] $firstName
        = $(Throw '-firstName  is required'),
        [string] $lastName 
        = $(Throw '-lastName  is required'),
        [string] $eID 
        = $(Throw '-eID  is required'),
        [string] $templateOU 
        = $(Throw '-templateOU  is required'),
        [object] $allADUsers 
        = $(Throw '-allADUsers  is required'),
        [object] $templateUser 
        = $(Throw '-templateUser  is required'),
        [string] $userDomain
        = $(Throw '-userDomain  is required'),
        [string] $systemDomain
        = $(Throw '-systemDomain  is required'),
        [string] $EOTargetDomain
        = $(Throw '-EOTargetDomain  is required'),
        [string] $contractType
        = $(Throw '-contractType  is required'),
        [string] $endDate
        = $(Throw '-endDate is required'),
        [string] $startDate
        = $(Throw '-startDate is required'),
        [string] $holdiayEntitlement
        = $(Throw '-holdiayEntitlement  is required'),
        [string] $manager
        = $(Throw '-manager  is required'),
        $jobTitle
    )
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)
    # declare work hash tables
    $userObject = @{
    }    
    $basicDetails = @{
    }
    $emailDetails = @{        
    }
    $extensionDetails = @{        
    }
    $settingDetails = @{        
    }
    #region COLLECT USER BASIC DETAILS
    ## format name
    $TextInfo = (Get-Culture).TextInfo
    $FN = $firstName.ToLower() # convert to lowercase
    $LN = $lastName.ToLower()
    $FN = $TextInfo.ToTitleCase($FN); $display_FN = $FN # capitalise (store this for display as-is)
    $LN = $TextInfo.ToTitleCase($LN); $display_LN = $LN
    $FN = $FN -replace "'", '' # for AD, have to replace special characters
    $LN = $LN -replace "'", ''
    $FN = $FN -replace ' ', '' # for AD, have to replace space character
    $LN = $LN -replace ' ', ''
    $FN = $FN.Trim()
    $LN = $LN.Trim()  
    ## create user password
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Generating and storing new password for the user account ' -Color White, Green, White
    $newPassword = Generate-Password
    ## new user principal name
    $newUPN = $($FN + '.' + $LN + '@' + $userDomain)
    if ($allADUsers.UserPrincipalName -match $newUPN) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' User Principal Name ' , " [$newUPN] ", ' is not unique. Running replacing it. ' -Color White, Yellow, White, Yellow, White
        $params = @{
            mode       = 'UPN'
            allADUsers = $allADUsers
            value      = $newUPN
        }
        $newUPN = Get-UNIQUEValue @params
    }
    # new sam account name
    $newSAM = ($FN + '.' + $LN)
    if ($newSAM.Length -gt 20) { 
        $newSAM = $newSAM.substring(0, 20) # the lenght of the SAM should be under 21 characters
    } 
    if ($allADUsers.SAMAccountName -match $newSAM) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' SAM Account Name ' , " [$newSAM] ", ' is not unique. Running replacing it. ' -Color White, Yellow, White, Yellow, White
        $params = @{
            mode       = 'SAM'
            allADUsers = $allADUsers
            value      = $newSAM
        }
        $newSAM = Get-UNIQUEValue @params
    }
    ## new employee id
    $newEID = $eID
    # if ($allADUsers.employeeid - $newEID) {
    #     $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' EmployeeID ' , " [$newEID] ", ' is not unique. Running replacing it. ' -Color White, Yellow, White, Yellow, White
    #     $params = @{
    #         mode       = 'eID'
    #         allADUsers = $allADUsers
    #         value      = $newEID
    #     }
    #     $newEID = Get-UNIQUEValue @params
    # }
    ## collect all basic details
    #finally:
    $basicDetails += @{ 'Name' = [string]$($display_FN + ' ' + $display_LN) }
    $basicDetails += @{ 'SamAccountName' = [string]$($newSAM) }
    $basicDetails += @{ 'Instance' = [string]$($templateUser.DistinguishedName) }
    $basicDetails += @{ 'Title' = [string]$($templateUser.title) }
    $basicDetails += @{ 'Department' = [string]$($templateUser.Department) }
    $basicDetails += @{ 'Company' = [string]$($templateUser.Company ) }
    $basicDetails += @{ 'Office' = [string]$($templateUser.Office) }
    $basicDetails += @{ 'DisplayName' = [string]$(($display_FN + ' ' + $display_LN)) }
    $basicDetails += @{ 'GivenName' = [string]$($FN) }
    $basicDetails += @{ 'SurName' = [string]$($LN) }
    $basicDetails += @{ 'Description' = [string]$('NEW STARTER - Created by ' + ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) + ' at ' + (Get-Date -Format G) ) } # description entry to help identify, who set the account up
    $basicDetails += @{ 'Enabled' = $true }
    $basicDetails += @{ 'Path' = [string]$($templateOU) }
    $basicDetails += @{ 'UserPrincipalName' = [string]$($newUPN) }
    $basicDetails += @{ 'AccountPassword' = (ConvertTo-SecureString -AsPlainText $newPassword -Force) }
    $basicDetails += @{ 'ChangePasswordAtLogon' = $true }
    $basicDetails += @{ 'EmployeeID' = [string]$($newEID) }
    #endregion
    #region COLLECT EMAIL ADDRESS DETAILS
    ## remote routing address
    $newRemoteRoutingAddress = $($FN + '.' + $LN) + '@' + $EOTargetDomain
    if ($allADUsers.proxyAddresses -match $newRemoteRoutingAddress) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Remote Routing Address ' , " [$newRemoteRoutingAddress] ", ' is not unique. Running replacing it. ' -Color White, Yellow, White, Yellow, White
        $params = @{
            mode       = 'RRA'
            allADUsers = $allADUsers
            value      = $newRemoteRoutingAddress
        }
        $newRemoteRoutingAddress = Get-UNIQUEValue @params
    }
    $emailDetails += @{'remoteRoutingAddress' = $newRemoteRoutingAddress }
    ## secondary smtp
    $newSecSMTP = 'smtp:' + $FN + $LN.substring(0, 1) + '@' + $UserDomain
    if ($allADUsers.proxyAddresses -match $newSecSMTP) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Secondary SMTP Address ' , " [$newSecSMTP] ", ' is not unique. Running replacing it. ' -Color White, Yellow, White, Yellow, White        
        $params = @{
            mode       = 'SSMTP'
            allADUsers = $allADUsers
            value      = $newSecSMTP
        }
        $newSecSMTP = Get-UNIQUEValue @params
    }
    # primary SMTP 
    if ($userDomain -match $systemDomain) {
        $newPrimarySMTP = ('SMTP:' + $newUPN) 
    }
    else {
        $oldPrimarySMTP = ('SMTP:' + $newUPN) -replace $userDomain, $systemDomain 
        $newPrimarySMTP = ('SMTP:' + $newUPN) 
        $newTerciarySMTP = ('smtp:' + $newUPN) -replace $userDomain, $systemDomain 
        $emailDetails += @{'terciarySMTPAddress' = $newTerciarySMTP }   # --> non-UK users will have a third SMTP
        $emailDetails += @{'oldSMTPAddress' = $oldPrimarySMTP }         # --> non-UK users will need this replaced with the new
    }
    $emailDetails += @{'primarySMTPAddress' = $newPrimarySMTP }         # --> all usrs will have a primary SMTP
    # secondary SMTP
    $emailDetails += @{'secondarySMTPAddress' = $newSecSMTP }           # --> all usres will have a secondary smtp
    #endregion
    #region COLLECT EXTENSION ATTRIBUTE DETAILS
    ## contract type
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) 
    if ($contractType -match 'FullTime') {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Setting user contract type to ', "[$($contractType)]" -Color White, Green, White, Green
        $extensionDetails += @{'extensionAttribute11' = 0 }
    }
    elseif ($contractType -match 'PartTime') {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Setting user contract type to ', "[$($contractType)]" -Color White, Green, White, Green
        $extensionDetails += @{'extensionAttribute11' = 2 }
    }
    elseif ($contractType -match 'Temp') {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Setting user contract type to ', "[$($contractType)]" -Color White, Green, White, Green
        $extensionDetails += @{'extensionAttribute11' = 1 }
    }
    elseif ($contractType -match 'External') {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Setting user contract type to ', "[$($contractType)]" -Color White, Green, White, Green
        $extensionDetails += @{'extensionAttribute11' = 3 }
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Contract type is incorrect - not setting this' -Color White, Yellow, White
    }
    ## start date
    if ($startDate -match '^(19|20)\d\d[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])$') {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Setting user start date to ', "[$($startDate)]" -Color White, Green, White, Green
        $extensionDetails += @{'extensionAttribute13' = $startDate }
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Start date is incorrect - not setting this' -Color White, Yellow, White
    }
    ## company specific details
    if ($templateUser.extensionAttribute5) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding company code from template ', "[$($templateUser.extensionAttribute5)]" -Color White, Green, White, Green
        $extensionDetails += @{'extensionAttribute5' = $templateUser.extensionAttribute5 }
    }
    if ($templateUser.extensionAttribute6) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding country code from template ', "[$($templateUser.extensionAttribute6)]" -Color White, Green, White, Green
        $extensionDetails += @{'extensionAttribute6' = $templateUser.extensionAttribute6 }
    }
    if ($templateUser.extensionAttribute7) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding (short)company code from template ', "[$($templateUser.extensionAttribute7)]" -Color White, Green, White, Green
        $extensionDetails += @{'extensionAttribute7' = $templateUser.extensionAttribute7 }
    }
    ## holiday entitlement
    if ($holidayEntitlement) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding holiday entitlement ', "[$($holidayEntitlement)]" -Color White, Green, White, Green
        $extensionDetails += @{'extensionAttribute15' = $holidayEntitlement }
    }   
    ## JBA access (adding by default)
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding JBA access ', ' disable this manually from AD, if this is not required' -Color White, Green, White, Green
    $extensionDetails += @{'extensionAttribute10' = 1 }
    #endregion
    #region COLLECT OTHER SETTINGS
    ## expiration date
    if ($endDate -match '^(19|20)\d\d[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])$') {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Setting user end date to ', "[$($endDate)]" -Color White, Green, White, Green
        # $valueSplat += @{
        #     'adAccountExpiration' = $endDate 
        # }
        $settingDetails += @{'adAccountExpiration' = $endDate } 
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' End date is incorrect - not setting this' -Color White, Yellow, White
    }
    ## manager
    try {
        #$allADUsers | Where-Object {$_.SamAccountName -match $manager}
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Setting line manager for the user: ', "[$($manager)]" -Color White, Green, White, Green
        # $valueSplat += @{
        #     'manager' = $manager
        # }
        $settingDetails += @{'manager' = $manager } 
    }
    catch {
        Write-Warning $_.Exception.Message
        Continue
    }   
    ## job title
    try {
        if ($null -notlike $jobTitle) {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Updating job title of the user to: ', "[$($jobTitle)]" -Color White, Green, White, Green
            $settingDetails += @{'jobTitle' = $jobTitle } 
        }
    }
    catch {
        Write-Warning $_.Exception.Message
        Continue
    }  
    #endregion
    # merge hash tables and return final output
    $userObject += @{basicdetails = $basicDetails } 
    $userObject += @{emailDetails = $emailDetails }  
    $userObject += @{extensionDetails = $extensionDetails } 
    $userObject += @{settingDetails = $settingDetails } 
    return $userObject
}
function Get-ParentOU {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [object] $user
    )
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)
    $parentOU = ($user | Select-Object @{ n = 'Path'; e = {
                $_.DistinguishedName -replace "CN=$($_.cn),", '' 
            } 
        }).path
    Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Found parent OU of user: ', " [$($user.DisplayName)] ", ' OU: ', " [$parentOU] " -Color White, Green, White, Green, White, Green
    return $parentOU 
}
function Generate-Password {
    function Get-RandomCharacters ($length, $characters) {
        $random = 1..$length | ForEach-Object {
            Get-Random -Maximum $characters.Length 
        }
        $private:ofs = ''
        return [string]$characters[$random]
    }
    function Scramble-String ([string]$inputString) {
        $characterArray = $inputString.ToCharArray()
        $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length
        $outputString = -join $scrambledStringArray
        return $outputString
    }
    $password = Get-RandomCharacters -length 2 -characters 'ABCDEFGHJKLMNOPQRSTUVWXYZ'
    $password += Get-RandomCharacters -length 2 -characters 'abcdefghjiklmnopqrstuvwxyz'
    $password += Get-RandomCharacters -length 3 -characters '1234567890'
    $password += Get-RandomCharacters -length 2 -characters 'abcdefghjiklmnopqrstuvwxyz'
    $password += Get-RandomCharacters -length 2 -characters '!@#$%-&?/' #'!$@-%' 
    return $password
}
function Sync-ActiveDirectory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $server,
        [pscredential]
        $credential
    )
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Syncronising ', ' [AD-DCs <--> AD-DCs]' -Color White, Green, White, Magenta
    $pso = New-PSSessionOption -ProxyAccessType NoProxyServer
    [void] (Invoke-Command -ComputerName $server -Credential $credential -SessionOption $pso -ScriptBlock {
            & 'C:\Windows\System32\repadmin.exe' /syncall
        })
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Sync complete - ', ' [OK] ' -Color White, Green, White, Green
}
function Sync-AzureActiveDirectory {
    [CmdletBinding()]
    param (
        [PSCredential] $credential = $(Throw '-credential is required'),
        [string] $server = $(Throw '-server is required')
    )
    <# The function of the#>
    # START
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Syncronising ', ' [AD --> AAD]' -Color White, Green, White, Magenta
    #Reusable sync function
    # (only changes)
    function DeltaSync {
        $pso = New-PSSessionOption -ProxyAccessType NoProxyServer
        [void](Invoke-Command -ComputerName $server -Credential $credential -SessionOption $pso -ScriptBlock {
                Import-Module 'C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1' #Import the AAD Sync module
                Start-ADSyncSyncCycle -PolicyType Delta #Start Delta - chagnes only - sync
            } -ErrorAction Stop)
    }   
    # Attempts.
    $count = 0
    do {
        try {
            DeltaSync
            $success = $true
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Sync complete - ', ' [OK] ' -Color White, Green, White, Green
        }
        catch {
            $delay = 30
            # Each failure (usually fail is the result of an ongoing sync) increases the coutner.
            $attempt = ($count + 1)
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Sync attempt failed  - ', " [$attempt] ", ' retry in ', " [$delay] ", ' seconds ' -Color White, Yellow, White, Yellow, White, Yellow, White
            $success = $false
            Start-Sleep -Seconds $delay
        }
        $count++
        # The sync will finish if either of these conditions met:
        # - succesfull sync
        # - 5 failures (no point trying at that time, as that means there are major sync issues)
    }until($count -eq 5 -or $success)
    # After 5 failed sync attempt the script quits. This should not happen!
    if (-not($success)) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Sync error (maximum number of tries reached) - ', ' [ERR] ' -Color White, Red, White, Red
    }
}
function Connect-OnPremExchange {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCredential] $credential,
        [Parameter(Mandatory = $true)] [string]
        $exchangeServer
    )
    #Get-PSSession | Remove-PSSession
    $pso = New-PSSessionOption -ProxyAccessType NoProxyServer -IdleTimeout 36000000
    [void] (Import-PSSession (New-PSSession -SessionOption $pso -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeServer/PowerShell/ -Authentication Kerberos -Credential $credential -Verbose:$false ) -DisableNameChecking -AllowClobber -Verbose:$false)
}
# function Connect-OnlineExchange {
#     [CmdletBinding()]
#     param (
#         [Parameter(Mandatory = $true)]
#         [PSCredential]
#         $credential,
#         [Parameter(Mandatory = $false)]
#         [string]
#         $prefix = 'o'
#     )
#     # Get-PSSession | Remove-PSSession
#     # #[void] (Import-Module ExchangeOnlineManagement -Force)
#     # [void] (Connect-ExchangeOnline -ShowBanner:$false -Credential $credential -ShowProgress $false -ShowBanner:$false)
#     $pso = New-PSSessionOption -ProxyAccessType NoProxyServer 
#     $Session = New-PSSession -SessionOption $pso -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection 
#     #$Session = New-PSSession -SessionOption $pso -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -AllowRedirection -Authentication Basic -Credential $credential
#     [void] (Import-PSSession $Session -Prefix $prefix -AllowClobber -WarningAction Ignore -DisableNameChecking)
#     #Get-oMailbox | Select-Object -First 5
#     # [void] (Import-PSSession $Session -Prefix $prefix -AllowClobber -WarningAction Ignore -DisableNameChecking)
# }
function Connect-OnlineExchange {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCredential]
        $credential,
        [Parameter(Mandatory = $false)]
        [string]
        $prefix = 'o'
    )
    #$pso = New-PSSessionOption -ProxyAccessType NoProxyServer 
    #$Session = New-PSSession -SessionOption $pso -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection 
    #[void] (Import-PSSession $Session -Prefix $prefix -AllowClobber -WarningAction Ignore -DisableNameChecking)
    Connect-ExchangeOnline -ShowBanner:$false -Prefix $prefix -Credential $credential
}
function Create-ONPremMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string]
        $targetDomain,
        $samAccountName
    )
    [void](Enable-Mailbox -Identity $samAccountName ) # create the mailbox
    [void](Enable-Mailbox -Identity $samAccountName -RemoteArchive -ArchiveDomain $targetDomain) # places the archive in the cloud
}
function Create-ONLINEMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string]
        $remoteRoutingAddress,
        $samAccountName
    )
    [void](Enable-RemoteMailbox -Identity $samAccountName -RemoteRoutingAddress $remoteRoutingAddress)
}
function Sync-ExchangeGUIDs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True)] [string] $UPN,
        [Parameter (Mandatory)] [ValidateSet('GET', 'PUT')] [string] $mode,
        [Parameter(Mandatory = $false)] [string] $GUID
    )
    if ($mode -match 'GET') {
        do {
            $delay = 120
            $GUID = $null
            $GUID = (  (Get-oMailbox $UPN -ErrorAction Ignore).ExchangeGUID  ).Guid
            if (-not $GUID) {
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Waiting for mailbox GUID for ', " [$UPN] ", ' to appear. Next check is in ', " [$delay] " , ' seconds' -Color White, Yellow, White, Yellow, White, Yellow, White
                Start-Sleep -Seconds $delay
            }        
        } until ($GUID)
        return $GUID
    }
    elseif ($mode -match 'PUT' -and $GUID) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding GUID ', " [$GUID] ", ' to mailbox ', " [$UPN] " -Color White, Green, White, Green, White, Green
        Get-RemoteMailbox -Identity $UPN | Set-RemoteMailbox -ExchangeGuid $GUID
        #Get-RemoteMailbox -Identity $UPN | Select-Object Name, RecipientTypeDetails, ExchangeGUID
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [e] ', ' Mode not recognized or no GUID' -Color White, Red, White
    }
}
function Wait-ForMSOLAccount {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [String] $UPN,
        [Parameter(Mandatory = $true)] [String] $systemDomain,
        [Parameter(Mandatory = $true)] [String] $delay
    )
    # (delay between each check. Increased the domain with slower replication)
    do {
        $MSOLACCPresent = $null
        $MSOLACCPresent = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Ignore
        if (!($MSOLACCPresent)) {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Waiting for Microsoft Online sync of user principal name ', " [$UPN] ", ' reattempt in ' , " [$delay]" -Color White, Yellow, White, Yellow, White , Yellow
            Start-Sleep -Seconds $delay
        }
    }
    until($MSOLACCPresent)    
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' User principal name ', " [$UPN] ", ' is now synced to Microsoft Online' -Color White, Green, White, Yellow, White
}
function Set-FolderOwnership {
    param (
        [Parameter(Mandatory = $true)] [String] $folder
    )
    $folderProperties = Get-ItemProperty $folder
    $Path = $folderProperties.FullName
    $Acl = (Get-Item $Path).GetAccessControl('Access')
    $Username = $folderProperties.Name
    $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule($Username, 'Modify', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $Acl.SetAccessRule($Ar)
    Set-Acl -Path $Path -AclObject $Acl
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Updated ownership of folder:  ', " [$folder] ", ' it is now modifiable by the user: ', " [$Username] " -Color White, Green, White, Green, White, Green
}
function Create-DFS {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [String] $PeopleDFS,
        [Parameter(Mandatory = $true)] [String] $PeopleTargetPath,
        [Parameter(Mandatory = $true)] [String] $ProfileDFS,
        [Parameter(Mandatory = $true)] [String] $ProfileTargetPath
    )
    # Create PEOPLE DFS folder
    if (Get-DfsnFolderTarget -Path $PeopleDFS -ErrorAction Ignore) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' DFS path  ', " [$PeopleDFS] ", ' already exists' -Color White, Yellow, White, Yellow
    }
    else {
        [void] (New-Item -Path $PeopleTargetPath -ItemType Directory -Force)
        [void] ( New-DfsnFolder -Path $PeopleDFS -State Online -TargetPath $PeopleTargetPath -TargetState Online -ReferralPriorityClass globalhigh )
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Creating DFS path  ', " [$PeopleDFS] ", ' with initial target ', " [$PeopleTargetPath] " -Color White, Green, White, Green, White, Green
    }
    # Create PROFILE DFS folder
    if (Get-DfsnFolderTarget -Path $ProfileDFS -ErrorAction Ignore) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' DFS path  ', " [$ProfileDFS] ", ' already exists' -Color White, Yellow, White, Yellow
    }
    else {
        # otherwise it is not existing yet
        [void] (New-Item -Path $ProfileTargetPath -ItemType Directory -Force)
        [void] ( New-DfsnFolder -Path $ProfileDFS -State Online -TargetPath $ProfileTargetPath -TargetState Online -ReferralPriorityClass globalhigh )
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Creating DFS path  ', " [$ProfileDFS] ", ' with initial target ', " [$ProfileTargetPath] " -Color White, Green, White, Green, White, Green
    }
    # Take ownership of PEOPLE & PROFILE folders
    Set-FolderOwnership -folder $PeopleDFS
    Set-FolderOwnership -folder $ProfileDFS
}
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
function Process-Emailing {
    [cmdletbinding()]
    [OutputType([int])]
    param(
        [Parameter (Mandatory)] [ValidateSet(`
                'toManager', 'toFrenchManager', 'toSupportCreation', 'toSupportRemoval', 'templateOrManagerError', 'leaverNotFound', 'returnPhone', 'genericOnboardingFailure', 'genericOffboardingFailure' `
        )] [string] $mode,
        [Parameter(Mandatory)] $domain,
        [Parameter(Mandatory)] $smtpServer,
        [Parameter(Mandatory = $false)] [switch] $test,
        [Parameter(Mandatory = $false)] $testRecipients,
        [Parameter(Mandatory = $false)] $recipients,
        [Parameter(Mandatory = $false)] $HRRecipients,
        [Parameter(Mandatory = $false)] $departmentSignature,
        [Parameter(Mandatory = $false)] $templateName,
        [Parameter(Mandatory = $false)] $manager,
        [Parameter(Mandatory = $false)] $starterSender,
        [Parameter(Mandatory = $false)] $leaverSender,
        [Parameter(Mandatory = $false)] $computerUsagePolicy,
        [Parameter(Mandatory = $false)] $report,
        # [Parameter(Mandatory = $false)] $failureReport,
        [Parameter(Mandatory = $false)] $serviceDeskLink,
        [Parameter(Mandatory = $false)] $userPWD,
        [Parameter(Mandatory = $false)] $userName,
        [Parameter(Mandatory = $false)] [string] $link1,
        [Parameter(Mandatory = $false)] [string] $link2,
        [Parameter(Mandatory = $false)] [string] $link3,
        [Parameter(Mandatory = $false)] [string] $link4,
        [Parameter(Mandatory = $false)] [string] $link5,
        [Parameter(Mandatory = $false)] $attachments,
        [Parameter(Mandatory = $false)] [string] $leaverName,
        [Parameter(Mandatory = $false)] [string] $leaverEID,
        [Parameter(Mandatory = $false)] [string] $newStarterFullName,
        [Parameter(Mandatory = $false)] [switch] $addServiceDeskLink
    )
    begin {
        $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
        $head = @'
<style>
BODY{background-color:lightgrey;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}
TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}
</style>
'@
        $TextEncoding = [System.Text.Encoding]::UTF8
    }
    process {
        switch ($mode) {
            # option 1
            'toManager' {
                # from
                $from = $starterSender
                # to (recipients) # subject
                $subject = "New employee alert - $userName [$domain]"
                if ($test.IsPresent) {
                    $recipients = $testRecipients
                    $subject = '[TEST MODE] - ' + $subject
                }
                else {
                    $recipients = $recipients 
                }                  
                # email body
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Preparing email body ' -Color White, Green, White
                $EmailBody = "
                    <font face= ""Century Gothic"">
                    Hello,
                    <p> We have created the following user: <span style=`"color:blue`">" + " $($report.DisplayName) " + "</span> <br>
                    <p> Please provide the following details for this user for his/her first login: <br>
                    <ul style=""list-style-type:disc"">
                    <li> <p> SAM Account (login) name: <span style=`"color:blue`">" + " $($report.SAMAccountName) " + "</span>  </li>
                    <li> <p> Initial password: <span style=`"color:red`">" + " $userPWD " + '</span>  </li>
                    <li> <p> (please ensure to keep this information confidential; <br>
                        the user will also need to change this password during the first login to company systems; <br>
                        the new password should be known only by the user and should not be shared!) </li>
                    </li>
                    </ul>
                    '
                if ($addServiceDeskLink.ispresent) {
                    $EmailBody += "
                <p> As this information is not stored by IT, please ensure to keep this email! In case of any issues please contact the <a href='$serviceDeskLink'>Service Desk</a> ! <br>
                "
                }
                $EmailBody += "
                    <p> Please also find the computer usage policy <a href='$ComputerUsagePolicy'>here</a> ! <br>
                    " + `
                ($report | ConvertTo-Html -Head $head | Out-String) + `
                    "
                    <p> Thank you. <br>
                    <p>Regards, <br>
                    $departmentSignature
                    </P>
                    </font>
                    "
                # send email
                foreach ($to in $recipients ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Sending report to line manager ', " [$to] " -Color White, Green, White, Cyan
                    $props = @{
                        'SmtpServer' = $smtpServer
                        'To'         = $to
                        'From'       = $from
                        'Subject'    = $subject
                        'Body'       = $EmailBody
                        'Priority'   = 'High'
                        'Encoding'   = $TextEncoding
                    }
                    Send-MailMessage @props -BodyAsHtml 
                }   
                Break 
            }
            # option 2
            'toFrenchManager' {
                # from
                $from = $starterSender
                # to (recipients) # Subject
                $subject = "Alerte aux nouveaux employes - $userName [$domain]"
                if ($test.IsPresent) {
                    $recipients = $testRecipients
                    $subject = '[TEST MODE] - ' + $subject
                }
                else {
                    $recipients = $recipients 
                }    
                # email body
                $EmailBody =
                "
                    <font face= ""Century Gothic"">
                    Bonjour,
                    <p> Nous venons de creer l'utilisateur suivant: <span style=`"color:blue`">" + " $($report.DisplayName) " + "</span> <br>
                    <p> Merci de transmettre l'identifiant et mot de passe : <br>
                    <ul style=""list-style-type:disc"">
                    <li> <p> Identifiant: <span style=`"color:blue`">" + " $($report.SAMAccountName) " + "</span>  </li>
                    <li> <p> Mot de passe provisoire: <span style=`"color:red`">" + " $userPWD " + "</span>  </li>
                    <li> <p> Le mot de passe doit rester confidentiel <br>
                        Le mot de passe provisoire doit etre modifie des la premiere connexion par l'utilisateur; <br>
                        Le mot de passe ne doit pas etre partage!) </li>
                    </li>
                    </ul>
                    <p> Ces informations ne sont pas conservees par le service IT! En cas de probleme, merci de contacter <a href='$serviceDeskLink'>Service Desk</a> ! <br>
                    <p> Pour tous renseignements complementaires, n'hesitez pas a consulter notre computer usage policy <a href='$ComputerUsagePolicy'>here</a> ! <br>
                    " + `
                ($report | ConvertTo-Html -Head $head | Out-String) + `
                    "<p> Merci. <br>
                    <p>Cordialement, <br>
                    $departmentSignature
                    </P>
                    </font>
                    " 
                # send email
                foreach ($to in $recipients ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Sending report to line manager ', " [$to] " -Color White, Green, White, Cyan
                    $props = @{
                        'SmtpServer' = $smtpServer
                        'To'         = $to
                        'From'       = $from
                        'Subject'    = $subject
                        'Body'       = $EmailBody
                        'Priority'   = 'High'
                        'Encoding'   = $TextEncoding
                    }
                    Send-MailMessage @props -BodyAsHtml 
                }                   
                Break 
            }
            # option 3
            'toSupportCreation' {
                # from
                $from = $starterSender
                # to (recipients) # subject
                $subject = "Summary report - users created [$domain]"
                if ($test.IsPresent) {
                    $recipients = $testRecipients
                    $subject = '[TEST MODE] - ' + $subject
                }
                else {
                    $recipients = $recipients 
                }                  
                # email body
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Preparing email body ' -Color White, Green, White
                $EmailBody = "
                    <font face= 'Century Gothic'>
                    Hello,
                    <p> Please find below the details of the user creation on these links: <br>
                        <ul style=""list-style-type:disc"">
                        <li><a href='file:///$link1'>User creation summary</a></li>
                        <li><a href='file:///$link2'>Errors during script run</a></li>
                        <li><a href='file:///$link3'>Processed Input</a></li>
                        <li><a href='file:///$link5'>Script actions</a></li>  
                        <li><a href='file:///$link4'>Full script run details</a></li>   
                        </ul>
                    <p> Please see user onboarding summary at the end of this email. <br>
                    " + `
                ($report | ConvertTo-Html -Head $head | Out-String) + `
                    "
                    <p> Thank you. <br>
                    <p>Regards, <br>
                    $departmentSignature
                    </P>
                    </font> 
                    "
                # send email
                foreach ($to in $recipients ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Sending SUMMARY report to ', " [$to] " -Color White, Green, White, Cyan
                    $props = @{
                        'SmtpServer'  = $smtpServer
                        'To'          = $to
                        'From'        = $from
                        'Subject'     = $subject
                        'Body'        = $EmailBody
                        'Priority'    = 'High'
                        'Encoding'    = $TextEncoding
                        'Attachments' = $attachments # this one actually has attachments too
                    }
                    Send-MailMessage @props -BodyAsHtml 
                }   
                Break 
            }
            # option 4
            'toSupportRemoval' {
                # from
                $from = $leaverSender
                # to (recipients) # subject                
                $subject = "Summary report - users removed [$domain]"
                if ($test.IsPresent) {
                    $recipients = $testRecipients
                    $subject = '[TEST MODE] - ' + $subject
                }
                else {
                    $recipients = $recipients 
                } 
                # email body
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Preparing email body ' -Color White, Green, White
                $EmailBody = "
                    <font face= 'Century Gothic'>
                    Hello,
                    <p> Please find below the details of the user removal on these links: <br>
                        <ul style=""list-style-type:disc"">
                        <li><a href='file:///$link1'>User removal summary</a></li>
                        <li><a href='file:///$link2'>Errors during script run</a></li>
                        <li><a href='file:///$link3'>Processed Input</a></li>
                        <li><a href='file:///$link5'>Script actions</a></li>  
                        <li><a href='file:///$link4'>Full script run details</a></li>          
                        </ul>
                    <p> Succesfull processing: <br>  " + ($report | ConvertTo-Html -Head $head | Out-String) + `
                    # ' <p> Failed processing: <br> ' + ($failureReport | ConvertTo-Html -Head $head | Out-String) + `
                    "     
                    <p> Thank you. <br>
                    <p>Regards, <br>
                    $departmentSignature
                    </P>
                    </font> 
                    "
                # send email
                foreach ($to in $recipients ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Sending SUMMARY report to ', " [$to] " -Color White, Green, White, Cyan
                    $props = @{
                        'SmtpServer'  = $smtpServer
                        'To'          = $to
                        'From'        = $from
                        'Subject'     = $subject
                        'Body'        = $EmailBody
                        'Priority'    = 'High'
                        'Encoding'    = $TextEncoding
                        'Attachments' = $attachments # this one actually has attachments too
                    }
                    Send-MailMessage @props -BodyAsHtml 
                }   
                Break 
            }
            # option 5
            'templateOrManagerError' {
                # from
                $from = $starterSender
                # to (recipients) # subject                
                $subject = "Failed user creation [$newStarterFullName] - template or manager is leaver or does not exists [$domain]" 
                if ($test.IsPresent) {
                    $recipients = $testRecipients
                    $subject = '[TEST MODE] - ' + $subject
                }
                else {
                    $recipients = $recipients 
                }                
                # email body
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Preparing email body ' -Color White, Green, White
                $EmailBody = "
                    <font face= ""Century Gothic"">
                    Hello,
                    <p> The submitted template   <span style=`"color:red`">" + "[ $templateName ] " + "</span> or the provided manager <span style=`"color:red`">" + "[ $manager ] " + "</span>  was not found in the $domain Active Directory or is in the leaver OU. <br>        
                    <p> Please check the spelling / provide an account outside of the leaver OU, then re-submit this job!<br>
                    <p> Thank you. <br>
                    <p>Regards, <br>
                    $departmentSignature
                    </P>
                    </font> 
                    "
                # send email
                foreach ($to in $recipients ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Emailing failure (template is leaver or not found) to', " [$to] " -Color White, Yellow, White, Cyan
                    $props = @{
                        'SmtpServer' = $smtpServer
                        'To'         = $to
                        'From'       = $from
                        'Subject'    = $subject
                        'Body'       = $EmailBody
                        'Priority'   = 'High'
                        'Encoding'   = $TextEncoding
                    }
                    Send-MailMessage @props -BodyAsHtml 
                }                
                Break 
            }
            # option 6
            'leaverNotFound' {
                # from
                $from = $leaverSender
                # to (recipients) # subject                
                $subject = "Failed user removal - user not found or already processed [$domain]" 
                if ($test.IsPresent) {
                    $recipients = $testRecipients
                    $subject = '[TEST MODE] - ' + $subject
                }
                else {
                    $recipients = $recipients 
                } 
                # email body
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Preparing email body ' -Color White, Green, White
                $EmailBody = "
                    <font face= ""Century Gothic"">
                    Hello,
                    <p> The submitted leaver  <span style=`"color:red`">" + `
                    " $leaverName " + "</span> is not found in AD... <br>    
                    <p> Please verify, if this account belongs to an active user !<br>
                    <p> Thank you. <br>
                    <p>Regards, <br>
                    $departmentSignature
                    </P>
                    </font> 
                    "
                # send email
                foreach ($to in $recipients ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Emailing failure (leaver not found in AD or already processed) to', " [$to] " -Color White, Yellow, White, Cyan 
                    $props = @{
                        'SmtpServer' = $smtpServer
                        'To'         = $to
                        'From'       = $from
                        'Subject'    = $subject
                        'Body'       = $EmailBody
                        'Priority'   = 'High'
                        'Encoding'   = $TextEncoding
                    }
                    Send-MailMessage @props -BodyAsHtml 
                }   
                Break 
            }
            # option 7
            'returnPhone' {
                # from
                $from = $leaverSender
                # to (recipients) # subject
                $subject = "Leaver employee alert - return company mobile phone - $userName [$domain]"
                if ($test.IsPresent) {
                    $recipients = $testRecipients
                    $subject = '[TEST MODE] - ' + $subject
                }
                else {
                    $recipients = $HRRecipients
                }                 
                # email body
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Preparing email body ' -Color White, Green, White
                $EmailBody = "
                    <font face= ""Century Gothic"">
                    Hello,
                    <p> The following user  <span style=`"color:red`">" + " $leaverName  / $leaverEID " + "</span> has now left the company. <br>    
                    <p> Please check, if this user had company mobile phone  and if yes, it is returned. <br>
                    <p> Thank you. <br>
                    <p>Regards, <br>
                    $departmentSignature
                    </P>
                    </font> 
                    "
                # send email
                foreach ($to in $recipients ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Emailing HR to retrieve company mobile from  ', $leaverName, ' to ', " [$to] " -Color White, Green, White, Green, White, Cyan 
                    $props = @{
                        'SmtpServer' = $smtpServer
                        'To'         = $to
                        'From'       = $from
                        'Subject'    = $subject
                        'Body'       = $EmailBody
                        'Priority'   = 'High'
                        'Encoding'   = $TextEncoding
                    }
                    Send-MailMessage @props -BodyAsHtml 
                }  
                Break 
            }
            # option 8
            'genericOnboardingFailure' {
                # from
                $from = $starterSender
                # to (recipients) # subject                
                $subject = "Failed user creation [$newStarterFullName] -  [$domain]" 
                if ($test.IsPresent) {
                    $recipients = $testRecipients
                    $subject = '[TEST MODE] - ' + $subject
                }
                else {
                    $recipients = $recipients 
                }                
                # email body
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Preparing email body ' -Color White, Green, White
                $EmailBody = "
                    <font face= ""Century Gothic"">
                    Hello,
                    <p> Creation of this user   <span style=`"color:red`">" + " $newStarterFullName " + "</span> failed <br>   
                    <p> This usually happens, if the account or the employee ID already exist, creating a conflict.      
                    <p> Please re-submit this account after making neccesary corrections!<br>
                    <p> Thank you. <br>
                    <p>Regards, <br>
                    $departmentSignature
                    </P>
                    </font> 
                    "
                # send email
                foreach ($to in $recipients ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Emailing failure (template is leaver or not found) to', " [$to] " -Color White, Yellow, White, Cyan
                    $props = @{
                        'SmtpServer' = $smtpServer
                        'To'         = $to
                        'From'       = $from
                        'Subject'    = $subject
                        'Body'       = $EmailBody
                        'Priority'   = 'High'
                        'Encoding'   = $TextEncoding
                    }
                    Send-MailMessage @props -BodyAsHtml 
                }                
                Break 
            }  
            # option 9  
            'genericOffboardingFailure' {
                # from
                $from = $leaverSender
                # to (recipients) # subject                
                $subject = "Failed user removal [$leaverName] -  [$domain]" 
                if ($test.IsPresent) {
                    $recipients = $testRecipients
                    $subject = '[TEST MODE] - ' + $subject
                }
                else {
                    $recipients = $recipients 
                }                
                # email body
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Preparing email body ' -Color White, Green, White
                $EmailBody = "
                    <font face= ""Century Gothic"">
                    Hello,
                    <p> Offboarding of this user   <span style=`"color:red`">" + " $leaverName " + "</span> failed <br>        
                    <p> Please re-submit this account after making neccesary corrections!<br>
                    <p> Thank you. <br>
                    <p>Regards, <br>
                    $departmentSignature
                    </P>
                    </font> 
                    "
                # send email
                foreach ($to in $recipients ) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Emailing failure (template is leaver or not found) to', " [$to] " -Color White, Yellow, White, Cyan
                    $props = @{
                        'SmtpServer' = $smtpServer
                        'To'         = $to
                        'From'       = $from
                        'Subject'    = $subject
                        'Body'       = $EmailBody
                        'Priority'   = 'High'
                        'Encoding'   = $TextEncoding
                    }
                    Send-MailMessage @props -BodyAsHtml 
                }                
                Break 
            }                      
            Default {
                Write-Warning 'Invalid mode selected!'
            }
        }
        # finally send the message
    }
    end {
    }
}
# function Create-EmailBody {
#     [CmdletBinding()]
#     param (
#         [Parameter(Mandatory = $false)] [string] $computerUsagePolicy,
#         [Parameter (Mandatory)] [ValidateSet(`
#                 'toManager', 'toFrenchManager', 'toSupportCreation', 'toSupportRemoval', 'templateOrManagerError', 'leaverNotFound', 'returnPhone' `
#         )] [string] $mode,
#         # [Parameter (Mandatory = $false )] [switch] $toManager,
#         # [Parameter (Mandatory = $false )] [switch] $toFrenchManager,
#         # [Parameter (Mandatory = $false )] [switch] $toSupportCreation,
#         # [Parameter (Mandatory = $false )] [switch] $toSupportRemoval,
#         # [Parameter (Mandatory = $false )] [switch] $templateOrManagerError,
#         # [Parameter (Mandatory = $false )] [switch] $leaverNotFound,
#         # [Parameter (Mandatory = $false )] [switch] $returnPhone,
#         [Parameter(Mandatory = $false)] [string] $templateName,
#         [Parameter(Mandatory = $false)] [string] $leaverName,
#         [Parameter(Mandatory = $false)] [string] $leaverEID,
#         [Parameter(Mandatory = $false)] [pscustomobject] $report,
#         [Parameter(Mandatory = $false)] [string] $userPassword,
#         [Parameter(Mandatory = $false)] [string] $link1,
#         [Parameter(Mandatory = $false)] [string] $link2,
#         [Parameter(Mandatory = $false)] [string] $link3,
#         [Parameter(Mandatory = $false)] [string] $link4,
#         [Parameter(Mandatory = $false)] [string] $serviceDeskLink,
#         [Parameter(Mandatory = $false)] [string] $departmentSignature
#     ) 
#     $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Preparing email body ' -Color White, Green, White
#     # Email body formating
#     $head = @'
# <style>
# BODY{background-color:lightgrey;}
# TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
# TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}
# TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}
# </style>
# '@
#     $EmailBody = switch ($mode) {
#         'toManager' { 
#             "
#             <font face= ""Century Gothic"">
#             Hello,
#             <p> We have created the following user: <span style=`"color:blue`">" + " $($report.DisplayName) " + "</span> <br>
#             <p> Please provide the following details for this user for his/her first login: <br>
#             <ul style=""list-style-type:disc"">
#             <li> <p> SAM Account (login) name: <span style=`"color:blue`">" + " $($report.SAMAccountName) " + "</span>  </li>
#             <li> <p> Initial password: <span style=`"color:red`">" + " $userPassword " + "</span>  </li>
#             <li> <p> (please ensure to keep this information confidential; <br>
#                 the user will also need to change this password during the first login to company systems; <br>
#                 the new password should be known only by the user and should not be shared!) </li>
#             </li>
#             </ul>
#             <p> As this information is not stored by IT, please ensure to keep this email! In case of any issues please contact the <a href='$serviceDeskLink'>Service Desk</a> ! <br>
#             <p> Please also find the computer usage policy <a href='$ComputerUsagePolicy'>here</a> ! <br>
#             <p> Thank you. <br>
#             <p>Regards, <br>
#             $departmentSignature
#             </P>
#             </font>
#             " + ($report | ConvertTo-Html -Head $head | Out-String)
#             ; Break
#         }
#         'toFrenchManager' {
#             "
#             <font face= ""Century Gothic"">
#             Bonjour,
#             <p> Nous venons de creer l'utilisateur suivant: <span style=`"color:blue`">" + " $($report.DisplayName) " + "</span> <br>
#             <p> Merci de transmettre l'identifiant et mot de passe : <br>
#             <ul style=""list-style-type:disc"">
#             <li> <p> Identifiant: <span style=`"color:blue`">" + " $($report.SAMAccountName) " + "</span>  </li>
#             <li> <p> Mot de passe provisoire: <span style=`"color:red`">" + " $userPassword " + "</span>  </li>
#             <li> <p> Le mot de passe doit rester confidentiel <br>
#                 Le mot de passe provisoire doit etre modifie des la premiere connexion par l'utilisateur; <br>
#                 Le mot de passe ne doit pas etre partage!) </li>
#             </li>
#             </ul>
#             <p> Ces informations ne sont pas conservees par le service IT! En cas de probleme, merci de contacter <a href='$serviceDeskLink'>Service Desk</a> ! <br>
#             <p> Pour tous renseignements complementaires, n'hesitez pas a consulter notre computer usage policy <a href='$ComputerUsagePolicy'>here</a> ! <br>
#             <p> Merci. <br>
#             <p>Cordialement, <br>
#             $departmentSignature
#             </P>
#             </font>
#             " + ($report | ConvertTo-Html -Head $head | Out-String)
#             ; Break
#         }
#         'toSupportCreation' {
#             "
#             <font face= 'Century Gothic'>
#             Hello,
#             <p> Please find below the details of the user creation. <br>
#                 <ul style=""list-style-type:disc"">
#                 <li><a href='file:///$link1'>User creation summary</a></li>
#                 <li><a href='file:///$link2'>Errors during script run</a></li>
#                 <li><a href='file:///$link3'>Processed Input</a></li>
#                 <li><a href='file:///$link4'>Script run details</a></li>          
#                 </ul>
#             <p> Please see user onboarding summary at the end of this email. <br>
#             <p> Thank you. <br>
#             <p>Regards, <br>
#             $departmentSignature
#             </P>
#             </font> 
#             " + ($report | ConvertTo-Html -Head $head | Out-String)
#             ; Break
#         }
#         'templateOrManagerError' {
#             "
#             <font face= ""Century Gothic"">
#             Hello,
#             <p> The submitted template   <span style=`"color:red`">" + " $templateName " + "</span> or the provided manager is a leaver or not found in AD... <br>        
#             <p> Please submit a new template user !<br>
#             <p> Thank you. <br>
#             <p>Regards, <br>
#             $departmentSignature
#             </P>
#             </font> 
#             "
#             ; Break
#         }
#         'returnPhone' {
#             "
#             <font face= ""Century Gothic"">
#             Hello,
#             <p> The following user  <span style=`"color:red`">" + " $leaverName [$leaverEID] " + "</span> has now left the company. <br>    
#             <p> Please check, if the user had company phone and ensure it is returned. <br>
#             <p> Thank you. <br>
#             <p>Regards, <br>
#             $departmentSignature
#             </P>
#             </font> 
#             "
#             ; Break
#         }
#         'toSupportRemoval' {
#             "
#             <font face= 'Century Gothic'>
#             Hello,
#             <p> Please find below the details of the user removal on these links: <br>
#                 <ul style=""list-style-type:disc"">
#                 <li><a href='file:///$link1'>User removal summary</a></li>
#                 <li><a href='file:///$link2'>Errors during script run</a></li>
#                 <li><a href='file:///$link3'>Processed Input</a></li>
#                 <li><a href='file:///$link4'>Script run details</a></li>          
#                 </ul>
#             <p> Please see user offboarding summary at the end of this email. <br>
#             <p> Thank you. <br>
#             <p>Regards, <br>
#             $departmentSignature
#             </P>
#             </font> 
#             " + ($report | ConvertTo-Html -Head $head | Out-String)
#             ; Break
#         }
#         'leaverNotFound' {
#             "
#             <font face= ""Century Gothic"">
#             Hello,
#             <p> The submitted leaver  <span style=`"color:red`">" + " $leaverName " + '</span> is not found in AD... <br>    
#             <p> Please verify, if this account belongs to an active user !<br>
#             <p> Thank you. <br>
#             <p>Regards, <br>
#             $departmentSignature
#             </P>
#             </font> 
#             '
#             ; Break
#         }
#         Default {
#             Write-Warning 'You need to select a valid mode to create the email body!'
#             Break
#         }
#     }
#     #     $EmailBody =
#     #     '' 
#     #     if ($toManager.ispresent) {
#     #         elseif ($toFrenchManager.ispresent) {
#     #             $EmailBody +=
#     #             # $EmailBody -replace 'e', 'e'
#     #             # $EmailBody -replace 'e', 'e'
#     #             $EmailBody += ($report | ConvertTo-Html -Head $head | Out-String)
#     #         }    
#     #         elseif ($toSupportCreation.IsPresent) { 
#     #             <# ref:
#     #         https://foswiki.org/Support/Faq72
#     #         #>
#     #             $EmailBody +=
#     #             $EmailBody += 
#     #         }  
#     #         elseif ($templateOrManagerError.IsPresent) {
#     #             $EmailBody +=
#     #         }
#     #         elseif ($returnPhone.IsPresent) {
#     #             $EmailBody +=
#     #         }
#     #         elseif ($toSupportRemoval.IsPresent) {
#     #             # Email body formating
#     #             $head = @'
#     # <style>
#     # BODY{background-color:lightgrey;}
#     # TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
#     # TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}
#     # TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}
#     # </style>
#     # '@        
#     #             <# ref:
#     #         https://foswiki.org/Support/Faq72
#     #         #>
#     #             $EmailBody +=
#     #             $EmailBody += 
#     #         }
#     #         elseif ($leaverNotFound.IsPresent) {
#     #             $EmailBody +=
#     #         }
#     return $EmailBody
# }
function Move-MailboxOnline {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [String] $UPN,
        [Parameter(Mandatory = $true)] [String] $remoteHostName,
        [Parameter(Mandatory = $true)] [String] $targetDeliveryDomain,
        [Parameter(Mandatory = $true)] [pscredential] $remoteCredential,
        [Parameter(Mandatory = $true)] [pscustomobject] $user
    )
    $x = 0
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Moving mailbox to o365 ', " [$UPN] " -Color White, Yellow, White, Yellow
    $BatchName = "$UPN - leaver mailbox processing [$x]"
    function doTheMove {
        # sub script in order to make this easy to repeat, in case of Exchange online disk issues
        if (($user).msExchArchiveAddress -match 'onmicrosoft.com') {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Performing ', ' primary mailbox only ', ' move using remote ', " [$remoteHostName]" -Color White, Green, White, Green, White, Green
            [void] (New-oMoveRequest -PrimaryOnly -Identity $UPN -RemoteCredential $remoteCredential -Remote -RemoteHostName $remoteHostName -BatchName $Batchname -TargetDeliveryDomain $targetDeliveryDomain )
        }
        else {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Performing ', ' full mailbox ', ' move ' -Color White, Green, White, Green, White
            [void] ( New-oMoveRequest -Identity $UPN -RemoteCredential $remoteCredential -Remote -RemoteHostName $remoteHostName -BatchName $Batchname -TargetDeliveryDomain $targetDeliveryDomain)
        }
    }
    doTheMove 
    # Now wait for the move to complete
    do {
        $x++
        #check, if there is move request
        if (Get-oMoveRequest -BatchName $BatchName) {
            $delay = '120'
            $BatchStatus = (Get-oMoveRequest -BatchName $BatchName).status
            if ($BatchStatus -match 'StalledDueToTarget_DiskLatency') {
                # remove previous move and start a new one, in case of Exchange onlin issues
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Problem on the receiving end ', " [$BatchStatus] ", ' re-attempting ... ' -Color White, Yellow, White, Yellow, White
                $BatchName = "$UPN - leaver mailbox processing [$x]"
                # remove current run
                Get-oMoveRequest -BatchName $BatchName | Remove-oMoveRequest -Force -confirm:$false
                Start-Sleep -Seconds 5
                # re-create run
                doTheMove 
            }
            $BatchPercentage = (Get-oMoveRequest -BatchName $BatchName | Get-oMoveRequestStatistics).Percentcomplete
            #$BatchDetails = Get-MoveRequest | Where-Object { $_.BatchName -match $UserPrincipalName } | Select-Object displayname,BatchName,status
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Waiting for batch ', " [$BatchName /  $BatchPercentage %] " , ' to complete. Next check in ', " [$delay] ", ' seconds ' -Color White, Yellow, White, Yellow, White, Yellow, White
            #$BatchDetails
            Start-Sleep $delay
        }
        # do this until the move request completes or if it does not exist anymore
    } until (($BatchStatus -eq 'Completed') -or (!(Get-oMoveRequest -BatchName $BatchName)))
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm); Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Batch ', " [$BatchName] " , ' completed ' -Color White, Yellow, White, Yellow, White         
}
function Reconfigure-Onlinemailbox {
    [CmdletBinding()]
    param (       
        [Parameter(Mandatory = $true)] [String] $UPN,
        [Parameter(Mandatory = $false)] [int] $litigationHoldTime,
        [Parameter(Mandatory = $true)] [String] $systemDomain,
        [Parameter(Mandatory = $false)] [String] $forwardee,
        [Parameter(Mandatory = $false)] [switch] $withLitigation
    )
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Reconfiguring mailbox ', " [$UPN] " -Color White, Green, White, Green
    # litigation hold
    if ($withLitigation.IsPresent) {
        # w litigation
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Enableing litigation hold on mailbox ', " [$UPN] " -Color White, Green, White, Green
        Set-oMailbox -Identity $UPN -LitigationHoldEnabled $true -LitigationHoldDuration $litigationHoldTime -WarningAction Ignore
    }
    else {
        # w/o litigation
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Litigation hold is not possible in domain', " [$systemDomain] " -Color White, Yellow, White, Magenta
    }
    # convert to shared
    $delay = 30
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Converting mailbox ', " [$UPN] ", ' into shared  ' -Color White, Green, White, Green, White
    If (((Get-oMailbox $UPN).recipienttypedetails) -notlike '*SharedMailbox*') {
        Get-oMailbox $UPN | Set-oMailbox -Type Shared
        # Wait for the conversion to complete
        do {
            $mbxType = (Get-oMailbox $UPN).recipienttypedetails; # shared
            if ($mbxType -like '*UserMailbox*') {
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Conversion is in progress, next check in  ', "[$delay]", ' seconds ' -Color White, Yellow, White, Yellow, White
                Start-Sleep -Seconds $delay
            }
        } until ($mbxType -like '*SharedMailbox*')
    }
    # add forwarding if required
    if ($forwardee) {        
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding forwarding address ', " [$forwardee] ", ' on the mailbox  ' -Color White, Green, White, Green, White
        Set-oMailbox -Identity $UPN -ForwardingAddress $forwardee -DeliverToMailboxAndForward $true
    }
}
function Set-OOOMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)] [string]$company,        
        [Parameter(Mandatory = $false)] [string]$oomRecipient,
        [Parameter(Mandatory = $true)] [string]$UPN,
        [Parameter(Mandatory = $true)] [string]$name
    )
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm) ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Adding out-of-office messge to ', " [$UPN] ", ' mailbox  ' -Color White, Green, White, Green, White
    $date = (Get-Date -Format yyy-MM-dd)
    $Msg = 
    "
        <html> <body>
        <p> Dear Sender, </p>
        <p> As of today [$date] I am no longer working for $company.<br>
        "
    if ($oomRecipient) {
        $Msg += 
        "
        <p> Please re-send your email to, $oomRecipient ! </p>
        "
    }
    $Msg += 
    "
    <p>
    Thank you, <br>
    $name    
    </p>
    </body>
    </html>
    "
    Set-oMailboxAutoReplyConfiguration -Identity $UPN -AutoReplyState Disabled -ExternalMessage $null -InternalMessage $null -Confirm:$false
    Set-oMailboxAutoReplyConfiguration -Identity $UPN -ExternalAudience 'All' -AutoReplyState Enabled -InternalMessage $Msg -ExternalMessage $Msg -Confirm:$false
}
# Function to check the OOOM
function Check-OOMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $mbx = (Read-Host -Prompt 'please enter the full mailbox address to check for oom')
    )
    function Show-HTML ([string]$HTML) {
        [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
        [xml]$XAML = @'
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PowerShell HTML GUI" WindowStartupLocation="CenterScreen">
            <WebBrowser Name="WebBrowser"></WebBrowser>
    </Window>
'@
        #Read XAML
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $Form = [Windows.Markup.XamlReader]::Load($reader)
        #===========================================================================
        # Store Form Objects In PowerShell
        #===========================================================================
        $WebBrowser = $Form.FindName('WebBrowser')
        $WebBrowser.NavigateToString($HTML)
        $Form.ShowDialog()
    }
    # Check OOOM (update address)
    Show-HTML -HTML ((Get-MailboxAutoReplyConfiguration -Identity $mbx).externalmessage)
}
function Remove-DirectLicenses {
    #Requires -Module MSOnline
    [CmdletBinding(SupportsShouldProcess)]
    param (
        # Saves the report to the script location
        [Parameter()]
        [string]
        $UPN
    )
    <# 
    Source: https://github.com/nicolonsky/Techblog/blob/master/CleanupAzureADLicensing/Invoke-CleanupAADDirectLicenseAssignments.ps1
    #>
    $user = Get-MsolUser -UserPrincipalName $UPN
    # processing all licenses per user
    foreach ($license in $user.Licenses) {
        if ($license.GroupsAssigningLicense -contains $user.ObjectId -or $license.GroupsAssigningLicense.Count -lt 1) {        
            #$directLicenseAssignmentCount++
            Write-Verbose "User $($user.UserPrincipalName) ($($user.ObjectId)) has direct license assignment for sku '$($license.AccountSkuId)')"
            if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Remove license assignment for sku '$($license.AccountSkuId)'")) {
                #Write-Warning "Removing license assignment for sku '$($license.AccountSkuId) on target '$($user.UserPrincipalName)'"
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Removing license assignment ', " [$($license.AccountSkuId) ] " , ' from target user ' , " [$($user.UserPrincipalName)] " -Color White, Yellow, White, Yellow, White, Yellow
                #Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $license.AccountSkuId
                Set-MsolUserLicense -ObjectId $user.ObjectId -RemoveLicenses $license.AccountSkuId
            }
        }
        else {
            Write-Verbose "User $($user.UserPrincipalName) ($($user.ObjectId)) has group license assignment for sku '$($license.AccountSkuId)')"
        }
    }
}
function Backup-LeaverDetails {
    [CmdletBinding()]
    param (
        # Saves the report to the script location
        [Parameter(Mandatory)] [string] $archiveFolder,
        [Parameter(Mandatory = $false)] [string] $manager,
        [Parameter(Mandatory)] [string] $backupServer,
        [Parameter(Mandatory = $false)] [pscustomobject] $groups,
        [Parameter(Mandatory)] [string] $samAccountName,
        [Parameter(Mandatory)] [pscredential] $credential
    )
    # create personal backup folder
    $personalBackupFolder = ($archiveFolder + '\' + $samAccountName)        
    Invoke-Command -ComputerName $backupServer -ScriptBlock {
        if (-not (Test-Path -Path $using:personalBackupFolder)) {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Host -ForegroundColor White "[$timer] [i]  Creating archive folder "  # must be Write-Host, as PSWriteColor module not installed on the file server
            [void] ( New-Item -ItemType Directory -Path $using:personalBackupFolder -Force)
        }
        else {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Host -ForegroundColor Yellow "[$timer] [w]  Archive folder exists"  # must be Write-Host, as PSWriteColor module not installed on the file server
        }
    } -Credential $credential
    #  define backup files
    $grpBackupFile = $personalBackupFolder + '\' + $samAccountName + '_SecurityGroups_' + (Get-Date -Format d_M_yyyy) + '.csv'
    $mgrBackupFile = $personalBackupFolder + '\' + $samAccountName + '_Manager_' + (Get-Date -Format d_M_yyyy) + '.txt'
    # backup details
    if ($null -ne $manager) {
        # $manager | Out-File $mgrBackupFile -Force -ErrorAction Continue
        Invoke-Command -ComputerName $backupServer -ScriptBlock {
            $using:manager | Out-File $using:mgrBackupFile -Force #-ErrorAction Continue
        } -Credential $credential
    }
    else {
        'Manager N/A' | Out-File $mgrBackupFile -Force
    }
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Manager details saved to file ', " [$mgrBackupFile] " -Color White, green, White, Green
    Invoke-Command -ComputerName $backupServer -ScriptBlock {
        if ($null -ne $using:groups) {
            # convert group data to CSV
            $data = @()
            foreach ($group in $using:groups) {
                $row = New-Object PSObject
                $row | Add-Member -MemberType NoteProperty -Name 'Group' -Value $group        
                $data += $row
            }
            $data | Export-Csv $using:grpBackupFile -NoTypeInformation -Append -Force 
        }
        else {
            'Groups N/A' | Export-Csv $using:grpBackupFile -NoTypeInformation -Append -Force
        }
    } -Credential $credential
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Previous group membership details saved to file ', " [$grpBackupFile] " -Color White, green, White, Green
}
function Backup-DFS {
    [CmdletBinding()]
    param (
        # Saves the report to the script location
        [Parameter(Mandatory)] [string] $archiveFolder,
        [Parameter(Mandatory)] [string] $backupServer,
        [Parameter(Mandatory)] [pscredential] $credential,
        [Parameter(Mandatory)] [string] $samAccountName,
        [Parameter(Mandatory)] [string] $peopleDFS,
        [Parameter(Mandatory)] [string] $profileDFS,
        [Parameter(Mandatory)] [string] $dfsnServer
    )
    # create personal backup folder
    $personalBackupFolder = ($archiveFolder + '\' + $samAccountName)
    if (-not (Test-Path $personalBackupFolder ) ) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Archive folder does not exist, creating ', " [$personalBackupFolder] " -Color White, Yellow, White, Yellow
        Invoke-Command -ComputerName $backupServer -ScriptBlock {
            [void] ( New-Item -ItemType Directory -Path $using:personalBackupFolder -Force ) 
        } -Credential $credential
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Archive folder ', " [$personalBackupFolder] ", ' exists ' -Color White, Yellow, White, Yellow, White
    }
    # copy the contents of the user's PEOPLE folder for backup purposes
    if (Test-Path $peopleDFS) {
        $DFSTargetPath = Get-DfsnFolderTarget -Path $peopleDFS | Select-Object -ExpandProperty TargetPath    
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Backing up the users DFSN folder ', " [$DFSTargetPath] " , ' into archive folder ' , " [$personalBackupFolder] " -Color White, Green, White, Green, White, Green
        $rCopyLog = $personalBackupFolder + '\' + 'robocopy.log'
        [void] (ROBOCOPY $DfsTargetPath $personalBackupFolder /MOVE /E /MT:10 /W:2 /R:2 /LOG:$rCopyLog)
        #        [void] (ROBOCOPY $DfsTargetPath $personalBackupFolder /MOVE /E /ZB /MT:10 /W:2 /R:2 /NJH /NDL /NC /LOG:$rCopyLog)
    }
    # remove DFS folders
    $CimSession = New-CimSession -ComputerName $dfsnServer -Credential $credential
    if (Test-Path $peopleDFS) {
        Remove-DfsnFolder -ErrorAction SilentlyContinue -Path $peopleDFS -CimSession $CimSession -Force
    }
    if (Test-Path $profileDFS) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Removing profile DFS entry ', " [$profileDFS] " -Color White, Yellow, White, Yellow
        Remove-DfsnFolder -ErrorAction SilentlyContinue -Path $profileDFS -CimSession $CimSession -Force
    }
}
function Remove-ProfileDisk {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string] $samAccountName,
        [Parameter(Mandatory)] [string] $profileDiskLocation
    )
    $SID = $null
    # Backing up the EAP (should be "stop")
    $EAP = $ErrorActionPreference
    # Setting EAP to ignore errors
    $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Ignore  
    # Try to get the user object from the SAM
    $userObject = New-Object System.Security.Principal.NTAccount($samAccountName)
    # If there is user object, get the SID from this
    if ($userObject) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Host -ForegroundColor White "[$timer] [i]  User object found [$($userObject.Value)] " 
        $SID = $userObject.Translate([System.Security.Principal.SecurityIdentifier]).value
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Host -ForegroundColor White "[$timer] [i]  User object not found ." 
    }
    # If there is a SID process  it
    if ($SID) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Host -ForegroundColor White "[$timer] [i]  SID found [$SID] " 
        If (Get-ChildItem -Path $profileDiskLocation -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$SID*" } ) {
            Get-ChildItem -Path $profileDiskLocation | Where-Object { $_.Name -like "*$SID*" } | Remove-Item -ErrorAction Stop -Force
        }
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Host -ForegroundColor White "[$timer] [i]  SID not found " 
    }
    # Setting back EAP
    $ErrorActionPreference = [System.Management.Automation.ActionPreference]::$EAP
}
function Remove-VDI {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string]
        #$systemDomain,
        $cbServer,
        $domainNetBIOS,
        $samAccountName,
        $SCCMServer,
        $VMMServer,
        $SCCMSiteCode,
        $server,
        [Parameter(Mandatory = $true)][pscredential]
        $credential
    )
    ## FIND THE VDI
    # Find the active Connection Broker
    $activeCBServer = (Get-RDConnectionBrokerHighAvailability -ConnectionBroker $cbServer).ActiveManagementServer
    # Find all the VDI collections
    $collections = (Get-RDVirtualDesktopCollection -ConnectionBroker $activeCBServer).CollectionName
    $collections_count = $collections.count
    $collections = @($collections)
    <# Check each collection until we find the VDI of the user we are currently processing. End the search, if:
    - we found the VDI (leads to premature ending of the loop)
    - we checked each collection without reasult
    #>
    $x = 0
    $is_removed = $null
    do {
        foreach ($c in $collections) {
            # Identify, if the user HAS A VDI
            $x++
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Checking collection', " [$c] ", " [$x/$collections_count] " -Color White, DarkYellow, White, DarkYellow, White
            $VDIHostname = $null
            $VDIHostname = Get-RDPersonalVirtualDesktopAssignment -CollectionName (Get-RDVirtualDesktopCollection $c -ConnectionBroker $activeCBServer).CollectionName `
                -User ($domainNetBIOS + '\' + "$samAccountName") -ConnectionBroker $activeCBServer | Select-Object -ExpandProperty VirtualDesktopName -Verbose
            if ($VDIHostname) {
                $is_removed = $VDIHostname
            }
            # If the user HAS A VDI ...
            if ($null -ne $VDIHostname) {
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' User  ', " [$samAccountName] " , ' has been assigned VDI ', " [$VDIHostname] " , ' in collection ', " [$c] " -Color White, Green, White, Green, White, Green, White, Green
                # Collect the VDI-s details
                $null = $VDIObject                
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Collecting ', " [$VDIHostname] " , ' VM-s details for processing ' -Color White, Green, White, Green, White
                $VDIObject = Get-SCVirtualMachine -VMMServer $VMMServer -Name $VDIHostname
                # Remove the VDI from the collection to remove assignment
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Removing VDI ', " [$VDIHostname] " , ' from the collection ', "[$c]" -Color White, Green, White, Green, White, Green
                Remove-RDVirtualDesktopFromCollection -ConnectionBroker $activeCBServer -CollectionName (Get-RDVirtualDesktopCollection $c -ConnectionBroker $activeCBServer).CollectionName -VirtualDesktopName @($VDIHostname) -Verbose -Force
                # ...verify that the VDI has an SCCM/AD object and remove it
                if ($VDIObject) {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' Removing VDI object  ', " [$($VDIObject.Name)] " -Color White, Green, White, Green
                    Try {
                        $removeVDIObjectprops = @{
                            'VDIObject'    = $VDIObject
                            'SCCMServer'   = $SCCMServer
                            'SCCMSiteCode' = $SCCMSiteCode
                            'server'       = $server
                        }
                        Remove-VDIObject @removeVDIObjectprops -credential $credential
                    }
                    catch {
                        Continue
                        # $is_removed = ('NO -' + $($VDIObject.name) + '-PROCESSING_FAILED')
                    }
                }
                # ...if there is no SCCM/AD object found, report this to the output
                else {
                    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [i] ', ' VDI object is not available ' -Color White, Green, White
                }               
            }
            # if the user HAS NO VDI, report this to the console
            else {
                $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' User  ', " [$samAccountName] " , ' has no VDI assigned in collection ', " [$c] " -Color White, Yellow, White, Yellow, White, Yellow
                # $is_removed = 'N/A-N/A-NA'
            }
        }
    } until (($x -ge $collections_count) -or ($is_removed -notlike $null)) 
    # Continue # move to the next user and start a new loop
    # output depending on of the VDI was found
    if ($is_removed -notlike $null) {
        $return = $is_removed
    }
    else {
        $return = 'NO VDI'
    }
    return  $return
}
function Remove-VDIObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][object] $VDIObject,
        [Parameter(Mandatory = $true)][string] $SCCMServer,
        [Parameter(Mandatory = $true)][string] $SCCMSiteCode,
        [Parameter(Mandatory = $true)][string] $server,
        [Parameter(Mandatory = $true)][pscredential] $credential
    )
    # Shut down the VDI first
    if ($VDIObject.status -eq 'Running') {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Shutting down VDI VM', " [$($VDIObject.Name)] " -Color White, Yellow, White, Yellow
        $VDIObject | Stop-SCVirtualMachine -Force 
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' VDI VM ', " [$($VDIObject.Name)] ", ' is powered off ' -Color White, Yellow, White, Yellow, White
    }    
    # Than remove the VDI
    $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Removing VM', " [$($VDIObject.Name)] " -Color White, Yellow, White, Yellow                 
    $VDIObject | Remove-SCVirtualMachine -Force 
    # Next remove the AD object
    if (Get-ADComputer -Identity $($VDIObject.Name) -Server $server -Credential $credential) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Removing computer object ', " [$VDIHostname] " -Color White, Yellow, White, Yellow                    
        Get-ADComputer -Identity $($VDIObject.Name) -Server $server -Credential $credential | Remove-ADObject -Recursive -Confirm:$False 
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' Computer object ', " [$VDIHostname] ", ' not found in AD ' -Color White, Yellow, White, Yellow, White                          
    }                    
    # Next remove the VDI object from SSCM
    $SCCMVDIDevice = Get-WmiObject -Class 'SMS_R_System' -Namespace "root/SMS/site_$($SCCMSiteCode)" -ComputerName $SCCMServer -Filter "Name='$VDIHostname'" # Save the VDI System Center Configuration Manager Device Object to $SCCMVDIDevice
    if ($SCCMVDIDevice) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)   ; Write-Color -LogFile $actionLog "[$timer]" , ' [w] ', ' System Center Configuration Manager device ', " [$VDIHostname  / (resourceID: $($SCCMVDIDevice.ResourceID) )] ", ' is deleted ' -Color White, Yellow, White, Yellow, White                     
        $SCCMVDIDevice.Delete()  
    }
}
function Cleanup-EarlierRuns {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][string] $inputFolder
        ,
        [Parameter(Mandatory = $true)][string] $doneFolder
    )
    $earlierRunFiles = Get-ChildItem $inputFolder | Where-Object { $_.FullName -match '.Processed' } 
    foreach ($file in $earlierRunFiles) {
        Write-Host "Moving $($file.FullName) -->TO--> $doneFolder "
        Move-Item $($file.FullName) -Destination $doneFolder -Force
    }  
}
function Get-ADUser_custom {
    # https://social.technet.microsoft.com/Forums/en-US/15666ba3-ba83-4ceb-9af6-77194072b413/returning-employeeid-property-with-systemdirectoryservicesdirectorysearcher?forum=winserverpowershell
    # https://social.technet.microsoft.com/Forums/en-US/59565819-6464-45fa-92c6-e3999d4c1cea/powershell-script-for-ad-user-information?forum=ITCG
    # https://lazywinadmin.com/2013/10/powershell-using-adsi-with-alternate.html
    # https://stackoverflow.com/questions/90652/can-i-get-more-than-1000-records-from-a-directorysearcher
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][pscredential] $Credential
        ,
        # [Parameter(Mandatory = $true)][string] $server
        # ,
        [Parameter(Mandatory = $true)][string] $systemDomain
    )
    $splits = $SystemDomain -split '\.'
    $searchroot = 'LDAP://DC=' + $splits[0] + ',DC=' + $splits[1] + ',DC=' + $splits[2]
    Write-Host "We will use $searchroot to search for users" -ForegroundColor Magenta
    # Create an ADSI Search
    $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
    # Get only the Group objects
    $Searcher.Filter = '(&(objectCategory=User)(objectclass=person))'
    # Limit the output to 50 objects
    $Searcher.SizeLimit = '9999'
    $searcher.PageSize = '9999'
    # Get the current domain
    #$DomainDN = $(([adsisearcher]"$searchroot").Searchroot.path)    
    #    $DomainDN = $(([adsisearcher]"").Searchroot.path)
    # Create an object "DirectoryEntry" and specify the domain, username and password
    $Domain = New-Object `
        -TypeName System.DirectoryServices.DirectoryEntry `
        -ArgumentList $searchroot, #$DomainDN,
    $($Credential.UserName),
    $($Credential.GetNetworkCredential().password)
    # Add the Domain to the search
    $Searcher.SearchRoot = $Domain
    # Execute the Search
    $currentADUsers = $Searcher.FindAll() | 
        ForEach-Object {
            [pscustomobject]@{
                Name              = $_.properties['name'][0]
                EmployeeID        = $($_.Properties.employeeid) # $_.properties['employeeid'][0]
                GivenName         = $($_.Properties.givenname)
                SurName           = $($_.Properties.sn)
                DistinguishedName = $($_.Properties.distinguishedname)
                SAMAccountName    = $($_.Properties.samaccountname)
            }
        }
    return  $currentADUsers
}
# SIG # Begin signature block
# MIIOWAYJKoZIhvcNAQcCoIIOSTCCDkUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUFAmePUHcCE6z9YYXBfjexl/r
# /0igggueMIIEnjCCA4agAwIBAgITTwAAAAb2JFytK6ojaAABAAAABjANBgkqhkiG
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUm4NHZyS/UukxDmoNQ95FDXBiUCAwDQYJKoZI
# hvcNAQEBBQAEggEAsxKEEdokVvNaENNWb2kluTWY6TOTW6Wj9oYZGzeJA6BtWMpL
# 46M6rYYIiKW8f04VGZUKsMz9k6BZxqkc9jDCQbb6KZUE7zEOfYsAbbfrNhhUKkr0
# PeahxiduExVx+dIegplXfDe9PMZQzSk9ygX5nH6i/vjPWn5u+0O28TMl3wH4jDnb
# qJ6GB9oFUjtGRX8gYfKLhZZBM/C4LxgvHtzJlKqYxwFCB3K1a3lyEXlEqVqk/2v8
# XVNCbFcBs6A5DDFz3xhBtdL9xQkIp8Qq0RPThugdkw8fcg5ub1JwHYRgFO2G5vzO
# JxlTOydR/Cgk2+g81S4B3iRojD4b1KDE0+TCRQ==
# SIG # End signature block
