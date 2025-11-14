<#
.SYNOPSIS
    Generate comprehensive M365 Guest Accounts Activity Report

.DESCRIPTION
    This script generates a detailed report of guest account activity in Microsoft 365, including:
    - Guest user list with last sign-in activity
    - Guest access permissions and group memberships
    - Inactive guest accounts
    - Recent guest account activity
    - Guest invitation status

.PARAMETER OutputPath
    Path where the report files will be saved (default: current directory)

.PARAMETER InactiveDays
    Number of days to consider a guest account inactive (default: 90)

.PARAMETER ExportFormat
    Export format: HTML, CSV, or Both (default: Both)

.EXAMPLE
    .\Get-GuestAccountsReport.ps1 -OutputPath "C:\Reports" -InactiveDays 90 -ExportFormat Both

.NOTES
    Requires Microsoft Graph PowerShell SDK
    Required Permissions: User.Read.All, AuditLog.Read.All, Directory.Read.All, Group.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".",

    [Parameter(Mandatory = $false)]
    [int]$InactiveDays = 90,

    [Parameter(Mandatory = $false)]
    [ValidateSet("HTML", "CSV", "Both")]
    [string]$ExportFormat = "Both"
)

# Function to check and install required modules
function Install-RequiredModules {
    Write-Host "Checking required modules..." -ForegroundColor Cyan

    $requiredModules = @(
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Groups",
        "Microsoft.Graph.Identity.DirectoryManagement",
        "Microsoft.Graph.Reports"
    )

    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing module: $module" -ForegroundColor Yellow
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module $module -ErrorAction Stop
    }

    Write-Host "All required modules are installed." -ForegroundColor Green
}

# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

    $scopes = @(
        "User.Read.All",
        "AuditLog.Read.All",
        "Directory.Read.All",
        "Group.Read.All"
    )

    try {
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit 1
    }
}

# Function to get all guest users
function Get-GuestUsers {
    Write-Host "Retrieving guest users..." -ForegroundColor Cyan

    try {
        $guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All `
            -Property Id, DisplayName, UserPrincipalName, Mail, CreatedDateTime, `
                      AccountEnabled, ExternalUserState, ExternalUserStateChangeDateTime, `
                      SignInActivity -ErrorAction Stop

        Write-Host "Found $($guestUsers.Count) guest users." -ForegroundColor Green
        return $guestUsers
    }
    catch {
        Write-Error "Failed to retrieve guest users: $_"
        return @()
    }
}

# Function to get group memberships for a user
function Get-UserGroupMemberships {
    param([string]$UserId)

    try {
        $groups = Get-MgUserMemberOf -UserId $UserId -All -ErrorAction SilentlyContinue
        $groupNames = $groups | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } |
                      ForEach-Object { $_.AdditionalProperties.displayName }
        return ($groupNames -join "; ")
    }
    catch {
        return "Error retrieving groups"
    }
}

# Function to process guest user data
function Get-GuestAccountDetails {
    param([array]$GuestUsers)

    Write-Host "Processing guest user details..." -ForegroundColor Cyan
    $reportData = @()
    $currentDate = Get-Date
    $inactiveThreshold = $currentDate.AddDays(-$InactiveDays)
    $counter = 0

    foreach ($guest in $GuestUsers) {
        $counter++
        Write-Progress -Activity "Processing guest accounts" -Status "Processing $counter of $($GuestUsers.Count)" `
                       -PercentComplete (($counter / $GuestUsers.Count) * 100)

        # Get last sign-in date
        $lastSignIn = $null
        $lastSignInDate = "Never"
        $daysSinceLastSignIn = "N/A"

        if ($guest.SignInActivity) {
            $lastSignIn = $guest.SignInActivity.LastSignInDateTime
            if ($lastSignIn) {
                $lastSignInDate = $lastSignIn.ToString("yyyy-MM-dd HH:mm:ss")
                $daysSinceLastSignIn = [math]::Round(($currentDate - $lastSignIn).TotalDays)
            }
        }

        # Determine if account is inactive
        $isInactive = $false
        if ($lastSignIn) {
            $isInactive = $lastSignIn -lt $inactiveThreshold
        }
        else {
            $isInactive = $true
        }

        # Get group memberships
        Write-Host "  Processing: $($guest.DisplayName)" -ForegroundColor Gray
        $groupMemberships = Get-UserGroupMemberships -UserId $guest.Id

        # Determine invitation status
        $invitationStatus = switch ($guest.ExternalUserState) {
            "Accepted" { "Accepted" }
            "PendingAcceptance" { "Pending" }
            default { "Unknown" }
        }

        # Create report object
        $reportData += [PSCustomObject]@{
            DisplayName = $guest.DisplayName
            UserPrincipalName = $guest.UserPrincipalName
            Email = $guest.Mail
            AccountEnabled = $guest.AccountEnabled
            InvitationStatus = $invitationStatus
            CreatedDate = if ($guest.CreatedDateTime) { $guest.CreatedDateTime.ToString("yyyy-MM-dd") } else { "N/A" }
            InvitationAcceptedDate = if ($guest.ExternalUserStateChangeDateTime) {
                $guest.ExternalUserStateChangeDateTime.ToString("yyyy-MM-dd")
            } else { "N/A" }
            LastSignInDate = $lastSignInDate
            DaysSinceLastSignIn = $daysSinceLastSignIn
            IsInactive = $isInactive
            GroupMemberships = $groupMemberships
        }
    }

    Write-Progress -Activity "Processing guest accounts" -Completed
    Write-Host "Processed $($reportData.Count) guest accounts." -ForegroundColor Green
    return $reportData
}

# Function to generate HTML report
function Export-HTMLReport {
    param(
        [array]$Data,
        [string]$FilePath
    )

    Write-Host "Generating HTML report..." -ForegroundColor Cyan

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>M365 Guest Accounts Activity Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        h1 {
            color: #0078d4;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 10px;
        }
        .summary {
            background-color: white;
            padding: 20px;
            margin: 20px 0;
            border-radius: 5px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .summary-item {
            display: inline-block;
            margin: 10px 20px 10px 0;
            padding: 15px;
            background-color: #f0f0f0;
            border-radius: 5px;
            min-width: 200px;
        }
        .summary-label {
            font-weight: bold;
            color: #666;
            font-size: 12px;
            text-transform: uppercase;
        }
        .summary-value {
            font-size: 24px;
            color: #0078d4;
            font-weight: bold;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            background-color: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        th {
            background-color: #0078d4;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
        }
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #ddd;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .inactive {
            background-color: #fff3cd;
        }
        .disabled {
            background-color: #f8d7da;
        }
        .status-badge {
            padding: 4px 8px;
            border-radius: 3px;
            font-size: 12px;
            font-weight: bold;
        }
        .status-accepted {
            background-color: #d4edda;
            color: #155724;
        }
        .status-pending {
            background-color: #fff3cd;
            color: #856404;
        }
        .status-enabled {
            background-color: #d4edda;
            color: #155724;
        }
        .status-disabled {
            background-color: #f8d7da;
            color: #721c24;
        }
        .timestamp {
            color: #666;
            font-size: 11px;
            margin-top: 20px;
        }
    </style>
</head>
<body>
    <h1>M365 Guest Accounts Activity Report</h1>
    <div class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</div>

    <div class="summary">
        <h2>Summary</h2>
        <div class="summary-item">
            <div class="summary-label">Total Guest Accounts</div>
            <div class="summary-value">$($Data.Count)</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Active Accounts</div>
            <div class="summary-value">$(($Data | Where-Object { !$_.IsInactive }).Count)</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Inactive ($InactiveDays+ days)</div>
            <div class="summary-value">$(($Data | Where-Object { $_.IsInactive }).Count)</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Disabled Accounts</div>
            <div class="summary-value">$(($Data | Where-Object { !$_.AccountEnabled }).Count)</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Pending Invitations</div>
            <div class="summary-value">$(($Data | Where-Object { $_.InvitationStatus -eq 'Pending' }).Count)</div>
        </div>
    </div>

    <h2>Guest Account Details</h2>
    <table>
        <thead>
            <tr>
                <th>Display Name</th>
                <th>Email</th>
                <th>Account Status</th>
                <th>Invitation Status</th>
                <th>Created Date</th>
                <th>Last Sign-In</th>
                <th>Days Since Sign-In</th>
                <th>Group Memberships</th>
            </tr>
        </thead>
        <tbody>
"@

    foreach ($item in $Data | Sort-Object IsInactive -Descending) {
        $rowClass = ""
        if ($item.IsInactive) { $rowClass = "inactive" }
        if (!$item.AccountEnabled) { $rowClass = "disabled" }

        $accountStatus = if ($item.AccountEnabled) {
            "<span class='status-badge status-enabled'>Enabled</span>"
        } else {
            "<span class='status-badge status-disabled'>Disabled</span>"
        }

        $invitationStatus = if ($item.InvitationStatus -eq "Accepted") {
            "<span class='status-badge status-accepted'>Accepted</span>"
        } else {
            "<span class='status-badge status-pending'>Pending</span>"
        }

        $html += @"
            <tr class='$rowClass'>
                <td>$($item.DisplayName)</td>
                <td>$($item.Email)</td>
                <td>$accountStatus</td>
                <td>$invitationStatus</td>
                <td>$($item.CreatedDate)</td>
                <td>$($item.LastSignInDate)</td>
                <td>$($item.DaysSinceLastSignIn)</td>
                <td>$($item.GroupMemberships)</td>
            </tr>
"@
    }

    $html += @"
        </tbody>
    </table>
</body>
</html>
"@

    $html | Out-File -FilePath $FilePath -Encoding UTF8
    Write-Host "HTML report saved to: $FilePath" -ForegroundColor Green
}

# Function to export CSV report
function Export-CSVReport {
    param(
        [array]$Data,
        [string]$FilePath
    )

    Write-Host "Generating CSV report..." -ForegroundColor Cyan
    $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV report saved to: $FilePath" -ForegroundColor Green
}

# Main execution
try {
    Write-Host "`n=== M365 Guest Accounts Activity Report ===" -ForegroundColor Cyan
    Write-Host "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" -ForegroundColor Gray

    # Ensure output directory exists
    if (!(Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    # Install required modules
    Install-RequiredModules

    # Connect to Microsoft Graph
    Connect-ToMicrosoftGraph

    # Get guest users
    $guestUsers = Get-GuestUsers

    if ($guestUsers.Count -eq 0) {
        Write-Host "No guest users found in the tenant." -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
        exit 0
    }

    # Process guest user data
    $reportData = Get-GuestAccountDetails -GuestUsers $guestUsers

    # Generate reports
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

    if ($ExportFormat -eq "HTML" -or $ExportFormat -eq "Both") {
        $htmlPath = Join-Path $OutputPath "GuestAccountsReport_$timestamp.html"
        Export-HTMLReport -Data $reportData -FilePath $htmlPath
    }

    if ($ExportFormat -eq "CSV" -or $ExportFormat -eq "Both") {
        $csvPath = Join-Path $OutputPath "GuestAccountsReport_$timestamp.csv"
        Export-CSVReport -Data $reportData -FilePath $csvPath
    }

    # Display summary
    Write-Host "`n=== Report Summary ===" -ForegroundColor Cyan
    Write-Host "Total Guest Accounts: $($reportData.Count)" -ForegroundColor White
    Write-Host "Active Accounts: $(($reportData | Where-Object { !$_.IsInactive }).Count)" -ForegroundColor Green
    Write-Host "Inactive Accounts ($InactiveDays+ days): $(($reportData | Where-Object { $_.IsInactive }).Count)" -ForegroundColor Yellow
    Write-Host "Disabled Accounts: $(($reportData | Where-Object { !$_.AccountEnabled }).Count)" -ForegroundColor Red
    Write-Host "Pending Invitations: $(($reportData | Where-Object { $_.InvitationStatus -eq 'Pending' }).Count)" -ForegroundColor Yellow

    Write-Host "`nEnd Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    Write-Host "`n=== Report Generation Completed ===" -ForegroundColor Green

    # Disconnect from Microsoft Graph
    Disconnect-MgGraph | Out-Null
}
catch {
    Write-Error "An error occurred: $_"
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    exit 1
}
