<#
.SYNOPSIS
Temporary password reset script for a single user
#>

# ========================= CONFIG =========================
$TargetUser = "hyalim@sy.goal.ie"

# ========================= HELPER FUNCTIONS =========================
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Time = (Get-Date).ToString("HH:mm:ss")
    $LogEntry = "[$Time] [$Level] $Message"
    Write-Host $LogEntry
}

# ========================= CONNECT TO GRAPH =========================
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop
    Write-Log "Connected to Microsoft Graph" "SUCCESS"
} catch {
    Write-Log ("Failed to connect to Microsoft Graph: {0}" -f $_.Exception.Message) "ERROR"
    exit
}

# ========================= ENTER NEW PASSWORD =========================
$NewPassword = Read-Host "Enter new password for $TargetUser" -AsSecureString

# ========================= PASSWORD RESET =========================
try {
    $User = Get-MgUser -UserId $TargetUser -ErrorAction Stop

    # Password profile object
    $PasswordProfile = @{
        ForceChangePasswordNextSignIn = $true
        Password                       = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                                            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewPassword))
    }

    # Attempt to update password
    Update-MgUser -UserId $User.Id -PasswordProfile $PasswordProfile -ErrorAction Stop
    Write-Log ("Password for {0} has been reset successfully" -f $TargetUser) "SUCCESS"
} catch {
    Write-Log ("[X] Failed to reset password for {0}: {1}" -f $TargetUser, $_.Exception.Message) "WARN"
    Write-Log "Insufficient privileges may be the cause. Ensure your account has a suitable role and scope." "INFO"
}

# ========================= CLEANUP =========================
try {
    Disconnect-MgGraph -Confirm:$false
    Write-Log "Disconnected from Microsoft Graph" "INFO"
} catch {
    Write-Log "No active Microsoft Graph session to disconnect" "INFO"
}

Write-Host "[*] Finished."
