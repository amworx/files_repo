<#
.SYNOPSIS
    Forces user to sign out of all sessions (Azure AD / Microsoft 365)
    and verifies token issue time.
.DESCRIPTION
    Works with Authentication Administrator or higher.
    Does not delete authentication methods, but ensures user must sign in again.
#>

# --- Step 1: Import and connect ---
Write-Host "`nConnecting to Azure AD..." -ForegroundColor Cyan

try {
    Import-Module AzureAD -ErrorAction Stop
    Connect-AzureAD -ErrorAction Stop
    Write-Host "Connected successfully." -ForegroundColor Green
}
catch {
    Write-Host "[x] Failed to connect to Azure AD:" $_.Exception.Message -ForegroundColor Red
    exit
}

# --- Step 2: Define user ---
$userUpn = "sytest1@sy.goal.ie"

try {
    $user = Get-AzureADUser -ObjectId $userUpn -ErrorAction Stop
    Write-Host "Target user found: $($user.DisplayName) [$userUpn]" -ForegroundColor Yellow
}
catch {
    Write-Host "[x] Could not find user $userUpn" -ForegroundColor Red
    exit
}

# --- Step 3: Revoke sessions ---
Write-Host "`nRevoking all sign-in sessions..." -ForegroundColor Cyan
try {
    Revoke-AzureADUserAllRefreshToken -ObjectId $user.ObjectId -ErrorAction Stop
    Write-Host "✅ All refresh tokens revoked successfully for $($user.DisplayName)." -ForegroundColor Green
}
catch {
    Write-Host "[x] Error revoking tokens:" $_.Exception.Message -ForegroundColor Red
    exit
}

# --- Step 4: Verify new token timestamp ---
Write-Host "`nVerifying token status..." -ForegroundColor Cyan
try {
    $updatedUser = Get-AzureADUser -ObjectId $userUpn
    $lastTokenTime = $updatedUser.RefreshTokensValidFromDateTime

    if ($lastTokenTime) {
        Write-Host "Last refresh token issued: $lastTokenTime (UTC)" -ForegroundColor Yellow
        Write-Host "✅ If this time is newer than before, revocation succeeded." -ForegroundColor Green
    }
    else {
        Write-Host "[i] No RefreshTokensValidFromDateTime property found (legacy user type)." -ForegroundColor DarkYellow
    }
}
catch {
    Write-Host "[x] Could not verify token time:" $_.Exception.Message -ForegroundColor Red
}

Write-Host "`n--- Done ---" -ForegroundColor Cyan
