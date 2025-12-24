param (
    [Parameter(Mandatory)]
    [string]$Mailbox,

    [Parameter(Mandatory)]
    [string]$User,

    [switch]$DryRun,

    [switch]$Apply
)

$LogDir = ".\logs"
$LogFile = "$LogDir\mailbox-access.log"

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

function Write-Log($message) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp | $message" | Out-File -FilePath $LogFile -Append
}


function Fail($msg) {
    Write-Error "[ERROR] $msg"
    exit 1
}

# --- SAFETY CHECK ---

if (-not $DryRun -and -not $Apply) {
    Fail "You must run the script with -DryRun first, then with -Apply"
}

if ($Apply -and -not $DryRun) {
    Write-Host "[WARNING] Make sure you already ran -DryRun before applying changes"
}

# --- VALIDATION ---

Write-Host "[INFO] Checking mailbox..."
if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
    Fail "Not connected to Exchange Online. Run Connect-ExchangeOnline first."
}
$mailboxObj = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
if (-not $mailboxObj) {
    Fail "Shared mailbox '$Mailbox' not found"
}
Write-Log "Checked mailbox '$Mailbox' and user '$User'"


Write-Host "[INFO] Checking user..."
$userObj = Get-Recipient -Identity $User -ErrorAction SilentlyContinue
if (-not $userObj) {
    Fail "User '$User' not found"
}

# --- PERMISSION CHECKS ---

$hasFullAccess = Get-MailboxPermission `
    -Identity $Mailbox `
    -User $User `
    -ErrorAction SilentlyContinue |
    Where-Object { $_.AccessRights -contains "FullAccess" }

$hasSendAs = Get-RecipientPermission `
    -Identity $Mailbox `
    -Trustee $User `
    -ErrorAction SilentlyContinue |
    Where-Object { $_.AccessRights -contains "SendAs" }

Write-Host ""
Write-Host "[STATUS] Current permissions:"

if ($hasFullAccess) {
    Write-Host "FullAccess (already granted)"
} else {
    Write-Host "FullAccess (missing)"
}

if ($hasSendAs) {
    Write-Host "SendAs (already granted)"
} else {
    Write-Host "SendAs (missing)"
}

# --- DRY-RUN MODE ---

if ($DryRun) {
    Write-Host ""
    Write-Host "[DRY-RUN] No changes were made."
    exit 0
}
Write-Log "DRY-RUN for mailbox '$Mailbox' user '$User'"

# --- APPLY MODE WITH CONFIRMATION ---

$missing = @()
if (-not $hasFullAccess) { $missing += "FullAccess" }
if (-not $hasSendAs) { $missing += "SendAs" }

if ($missing.Count -eq 0) {
    Write-Host "[INFO] No changes required"
    exit 0
}

Write-Host ""
Write-Host "User '$User' does NOT have:"
$missing | ForEach-Object { Write-Host " - $_" }

$confirmation = Read-Host "Do you want to grant missing permissions? (yes/no)"

if ($confirmation -ne "yes") {
    Write-Host "[INFO] Operation cancelled by user"
    exit 0
}

# --- APPLY CHANGES ---
Write-Log "APPLY permissions for mailbox '$Mailbox' user '$User'"
if (-not $hasFullAccess) {
    Write-Host "[INFO] Granting FullAccess..."
    Add-MailboxPermission `
        -Identity $Mailbox `
        -User $User `
        -AccessRights FullAccess `
        -InheritanceType All `
        -AutoMapping $false `
        -ErrorAction Stop
}

if (-not $hasSendAs) {
    Write-Host "[INFO] Granting SendAs..."
    Add-RecipientPermission `
        -Identity $Mailbox `
        -Trustee $User `
        -AccessRights SendAs `
        -Confirm:$false `
        -ErrorAction Stop
}

Write-Host "[SUCCESS] Missing permissions granted successfully"
Write-Log "Completed successfully"