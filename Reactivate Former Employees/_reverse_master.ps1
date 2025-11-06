<#
.SYNOPSIS
Reactivates former employee accounts and mailboxes:
 - Enable account
 - Unblock sign-in (Entra ID)
 - Unhide from GAL
 - Clear message delivery restrictions
 - Disable auto-reply
 - Convert to regular mailbox
 - Clear custom attributes
 - Add to distribution list
 - Remove all delegations
 - Disable archive if exists
 - Add to security groups based on type
 - Set regional configuration
 - Force sign-out from all sessions (Reset MFA)
 - Logs all actions
#>

# ========================= LOAD CONFIG =========================
$ConfigFilePath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "config.json"

if (Test-Path $ConfigFilePath) {
    $Config = Get-Content -Path $ConfigFilePath | ConvertFrom-Json
    Write-Host "Config loaded successfully." -ForegroundColor Green
} else {
    Write-Host "Config file not found: $ConfigFilePath" -ForegroundColor Red
    exit
}

# ========================= ACCESS CONFIG VARIABLES =========================
$AdminUPN         = $Config.AdminUPN
$DistributionList = $Config.DistributionList
$Groups           = $Config.Groups
$LogFile          = Join-Path $ScriptPath "_reverse_log.txt"
$CSVFilePath      = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) $Config.CSVFilePath

# Ensure CSV path is valid before proceeding
if (-not (Test-Path $CSVFilePath)) {
    Write-Host "[X] CSV file not found: $CSVFilePath" -ForegroundColor Red
    exit
}

# ========================= HELPER FUNCTIONS =========================
function Write-Log {
    param ([string]$Message, [string]$Level = "INFO")
    $Time = (Get-Date).ToString("HH:mm:ss")
    $LogEntry = "[$Time] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogEntry
    switch ($Level) {
        "INFO"    { Write-Host "  [*] $Message" -ForegroundColor Cyan }
        "SUCCESS" { Write-Host "  [+] $Message" -ForegroundColor Green }
        "WARN"    { Write-Host "  [!] $Message" -ForegroundColor Yellow }
        "ERROR"   { Write-Host "  [X] $Message" -ForegroundColor Red }
        default   { Write-Host "  [ ] $Message" }
    }
}

function Ensure-Module {
    param ([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        $resp = Read-Host "[?] Module $Name is missing. Install it? (Y/n) [Enter = Yes]"
        if ($resp -eq "" -or $resp -match "^[Yy]") {
            Install-Module $Name -Force -AllowClobber -Scope CurrentUser
        } else {
            Write-Log "Required module $Name is not installed. Exiting..." "ERROR"
            exit
        }
    }
}

# === Reset MFA / Force Sign-out ===
function Reset-MFA {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )

    Write-Host "`n[INFO] Resetting MFA / forcing sign-out for $UserPrincipalName ..." -ForegroundColor Cyan

    try {
        # Ensure AzureAD module is available
        if (-not (Get-Module -ListAvailable -Name AzureAD)) {
            Write-Host "[INFO] Installing AzureAD module..." -ForegroundColor Yellow
            Install-Module AzureAD -Force -AllowClobber -Scope CurrentUser
        }

        # Import module and connect if not connected
        Import-Module AzureAD -ErrorAction Stop
        try {
            # Try a simple call to see if connection exists
            Get-AzureADTenantDetail -ErrorAction Stop | Out-Null
        } catch {
            Write-Host "[INFO] Connecting to Azure AD..." -ForegroundColor Yellow
            Connect-AzureAD -ErrorAction Stop
        }

        # Get user object
        $User = Get-AzureADUser -ObjectId $UserPrincipalName -ErrorAction Stop
        if (-not $User) {
            Write-Host "[!] User not found: $UserPrincipalName" -ForegroundColor Red
            return
        }

        # Revoke tokens
        Revoke-AzureADUserAllRefreshToken -ObjectId $User.ObjectId -ErrorAction Stop
        Write-Host "[+] Tokens revoked successfully for $UserPrincipalName" -ForegroundColor Green

        # Verify refresh token timestamp
        $updatedUser = Get-AzureADUser -ObjectId $UserPrincipalName
        $lastTokenTime = $updatedUser.RefreshTokensValidFromDateTime
        if ($lastTokenTime) {
            Write-Host "[i] Last Refresh Token valid from: $lastTokenTime (UTC)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "[!] Failed to revoke tokens for ${UserPrincipalName}: $($_.Exception.Message)" -ForegroundColor Red
    }
}


# ========================= MODULE CHECK =========================
$RequiredModules = @(
    "ExchangeOnlineManagement",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups"
    "AzureAD"
)
foreach ($mod in $RequiredModules) {
    Ensure-Module -Name $mod
    Import-Module $mod
}

# ========================= INITIALIZE =========================
$SessionStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $LogFile -Value "`n=== Session started at $SessionStart ==="
Write-Host "`n==============================================" -ForegroundColor Gray
Write-Host "   REVERSE_ZZ - Account Reactivation Utility" -ForegroundColor Magenta
Write-Host "==============================================" -ForegroundColor Gray

# ========================= CONNECT TO EXO =========================
try {
    Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ErrorAction Stop
    Write-Log "Connected to Exchange Online as $AdminUPN" "SUCCESS"
} catch {
    Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" "ERROR"
    exit
}

# ========================= CONNECT TO GRAPH =========================
$GraphConnected = $false
for ($i = 1; $i -le 3 -and -not $GraphConnected; $i++) {
    try {
        Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All" -ErrorAction Stop
        Write-Log "Connected to Microsoft Graph (Attempt $i)" "SUCCESS"
        $GraphConnected = $true
    } catch {
        Write-Log ("Failed to connect to Microsoft Graph (Attempt {0}): {1}" -f $i, $_.Exception.Message) "WARN"
        Start-Sleep -Seconds 5
    }
}
if (-not $GraphConnected) {
    Write-Log "Could not connect to Microsoft Graph after 3 attempts. Exiting..." "ERROR"
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# ========================= IMPORT CSV =========================
if (-not (Test-Path $CSVFilePath)) {
    Write-Log "CSV file not found: $CSVFilePath" "ERROR"
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
    exit
}

$Members = Import-Csv -Path $CSVFilePath
Write-Log ("Imported {0} member(s) from CSV." -f $Members.Count) "INFO"

# ========================= COUNTERS =========================
$Total = 0; $SuccessCount = 0; $WarningCount = 0; $ErrorCount = 0

# ========================= PROCESS EACH USER =========================
foreach ($index in 0..($Members.Count - 1)) {
    $Total++
    $Member = $Members[$index]
    $Email = $Member.email.Trim()
    $Type  = $Member.type.Trim()
    # Display progress with current index and total
    Write-Host "`n--------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ("Processing: {0} ({1}/{2})" -f $Email, ($index + 1), $Members.Count) -ForegroundColor White
    Write-Host "--------------------------------------------------" -ForegroundColor DarkGray

    # 1️⃣ Enable account
    try {
        $User = Get-MgUser -UserId $Email -ErrorAction Stop
        if ($User.AccountEnabled -ne $true) {
            Update-MgUser -UserId $User.Id -AccountEnabled:$true
            Write-Log ("Enabled account for {0}" -f $Email) "SUCCESS"
            $SuccessCount++
        } else {
            Write-Log ("{0} account already enabled" -f $Email) "INFO"
        }
    } catch {
        Write-Log ("Failed to enable account for {0}: {1}" -f $Email, $_.Exception.Message) "ERROR"
        $ErrorCount++
    }

    # 2️⃣ Unhide from GAL
    try {
        Set-Mailbox -Identity $Email -HiddenFromAddressListsEnabled:$false -ErrorAction Stop
        Write-Log ("Unhidden from GAL for {0}" -f $Email) "SUCCESS"
        $SuccessCount++
    } catch {
        Write-Log ("Failed to unhide from GAL: {0}" -f $_.Exception.Message) "WARN"
        $WarningCount++
    }

    # 3️⃣ Message delivery restrictions
    try {
        Set-Mailbox -Identity $Email -AcceptMessagesOnlyFromSendersOrMembers @() -RejectMessagesFromSendersOrMembers @() -ErrorAction Stop
        Write-Log ("Message delivery restrictions reset for {0}" -f $Email) "SUCCESS"
        $SuccessCount++
    } catch {
        Write-Log ("Failed to reset message delivery restrictions for {0}: {1}" -f $Email, $_.Exception.Message) "WARN"
        $WarningCount++
    }

    # 4️⃣ Remove all delegations
    try {
        Remove-MailboxPermission -Identity $Email -User "*" -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue
        Remove-RecipientPermission -Identity $Email -Trustee "*" -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue
        Set-Mailbox -Identity $Email -GrantSendOnBehalfTo @() -ErrorAction SilentlyContinue
        Write-Log ("Removed all delegations (FullAccess, SendAs, SendOnBehalf) for {0}" -f $Email) "SUCCESS"
        $SuccessCount++
    } catch {
        Write-Log ("Error removing delegations for {0}: {1}" -f $Email, $_.Exception.Message) "WARN"
        $WarningCount++
    }

    # 5️⃣ Convert to regular mailbox
    try {
        $Mailbox = Get-Mailbox -Identity $Email
        if ($Mailbox.RecipientTypeDetails -ne "RegularMailbox") {
            Set-Mailbox -Identity $Email -Type Regular -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Log ("{0} converted to regular mailbox" -f $Email) "SUCCESS"
            $SuccessCount++
        } else {
            Write-Log ("{0} is already a regular mailbox" -f $Email) "INFO"
        }
    } catch {
        if ($_ -match "The mailbox .* is already of the type.*Regular") {
            Write-Log ("{0} is already a regular mailbox (skipping conversion)" -f $Email) "INFO"
        } else {
            Write-Log ("Error converting {0}: {1}" -f $Email, $_.Exception.Message) "ERROR"
            $ErrorCount++
        }
    }

    # 6️⃣ Disable auto-reply
    try {
        Set-MailboxAutoReplyConfiguration -Identity $Email -AutoReplyState Disabled -ErrorAction Stop
        Write-Log ("Auto-reply disabled for {0}" -f $Email) "SUCCESS"
        $SuccessCount++
    } catch {
        Write-Log ("Error disabling auto-reply for {0}: {1}" -f $Email, $_.Exception.Message) "WARN"
        $WarningCount++
    }

    # 7️⃣ Clear custom attributes
    try {
        Set-Mailbox -Identity $Email -ErrorAction Stop `
            -CustomAttribute1 $null -CustomAttribute2 $null -CustomAttribute3 $null `
            -CustomAttribute4 $null -CustomAttribute5 $null -CustomAttribute6 $null `
            -CustomAttribute7 $null -CustomAttribute8 $null -CustomAttribute9 $null `
            -CustomAttribute10 $null -CustomAttribute11 $null -CustomAttribute12 $null `
            -CustomAttribute13 $null -CustomAttribute14 $null -CustomAttribute15 $null
        Write-Log ("Cleared custom attributes for {0}" -f $Email) "SUCCESS"
        $SuccessCount++
    } catch {
        Write-Log ("Error clearing attributes for {0}: {1}" -f $Email, $_.Exception.Message) "WARN"
        $WarningCount++
    }

    # 8️⃣ Add to distribution list
    try {
        Add-DistributionGroupMember -Identity $DistributionList -Member $Email -ErrorAction Stop
        Write-Log ("{0} added to {1}" -f $Email, $DistributionList) "SUCCESS"
        $SuccessCount++
    } catch {
        if ($_.Exception.Message -match "already a member") {
            Write-Log ("{0} already member of {1}" -f $Email, $DistributionList) "WARN"
            $WarningCount++
        } else {
            Write-Log ("Failed to add {0} to DL: {1}" -f $Email, $_.Exception.Message) "ERROR"
            $ErrorCount++
        }
    }

    # 9️⃣ Disable mailbox archive
    try {
        $Mailbox = Get-Mailbox -Identity $Email -ErrorAction Stop
        if ($Mailbox.ArchiveStatus -eq "Active") {
            Disable-Mailbox -Identity $Email -Archive -Confirm:$false
            Write-Log ("Disabled archive for {0}" -f $Email) "SUCCESS"
            $SuccessCount++
        } else {
            Write-Log ("No archive enabled for {0}" -f $Email) "INFO"
        }
    } catch {
        Write-Log ("Error disabling archive for {0}: {1}" -f $Email, $_.Exception.Message) "WARN"
        $WarningCount++
    }

    # 🔟 Add to Security Groups based on type
    try {
        if ($null -eq $Groups) {
            Write-Host "[X] No groups found in config." -ForegroundColor Red
            exit
        }

        # Always add to MetaCompliance Users
        if ($Groups.MetaComplianceUsers) {
            New-MgGroupMember -GroupId $Groups.MetaComplianceUsers -DirectoryObjectId $User.Id -ErrorAction Stop
        } else {
            Write-Host "[X] MetaComplianceUsers group ID not found." -ForegroundColor Red
        }

        if ($Type -eq "E1" -and $Groups.SyriaE1) {
            New-MgGroupMember -GroupId $Groups.SyriaE1 -DirectoryObjectId $User.Id -ErrorAction Stop
        }

        if ($Type -eq "E3") {
            if ($Groups.SY_AMP_EPM) {
                New-MgGroupMember -GroupId $Groups.SY_AMP_EPM -DirectoryObjectId $User.Id -ErrorAction Stop
            }
            if ($Groups.MicrosoftE3) {
                New-MgGroupMember -GroupId $Groups.MicrosoftE3 -DirectoryObjectId $User.Id -ErrorAction Stop
            }
        }

        Write-Log ("Added {0} to security groups based on type {1}" -f $Email, $Type) "SUCCESS"
        $SuccessCount++
    } catch {
        Write-Log ("Failed to add {0} to security groups: {1}" -f $Email, $_.Exception.Message) "WARN"
        $WarningCount++
    }
    # 11️⃣ Set regional configuration
    try {
        Set-MailboxRegionalConfiguration -Identity $Email -Language en-GB -TimeZone "Turkey Standard Time" -ErrorAction Stop
        Write-Log ("Set regional configuration (Language=en-GB, TimeZone=Turkey) for {0}" -f $Email) "SUCCESS"
        $SuccessCount++
    } catch {
        Write-Log ("Failed to set regional configuration for {0}: {1}" -f $Email, $_.Exception.Message) "WARN"
        $WarningCount++
    }

    # 1️⃣2️⃣ Reset MFA / Force sign-out
    Reset-MFA -UserPrincipalName $Email
}

# ========================= CLEANUP =========================
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
$SessionEnd = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $LogFile -Value "=== Session ended at $SessionEnd ===`n"

Write-Host "`n==============================================" -ForegroundColor Gray
Write-Host ("   Process completed at {0}" -f $SessionEnd) -ForegroundColor Magenta
Write-Host "==============================================" -ForegroundColor Gray
Write-Host ("   Total processed : {0}" -f $Total) -ForegroundColor White
Write-Host ("   Success actions : {0}" -f $SuccessCount) -ForegroundColor Green
Write-Host ("   Warnings        : {0}" -f $WarningCount) -ForegroundColor Yellow
Write-Host ("   Errors          : {0}" -f $ErrorCount) -ForegroundColor Red
Write-Host "   Log file saved to:" -ForegroundColor White
Write-Host ("   {0}" -f $LogFile) -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Gray
