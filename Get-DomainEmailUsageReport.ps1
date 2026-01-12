<#
.SYNOPSIS
    Analyzes email usage for a specific domain in Exchange Online to assess impact before domain removal.

.DESCRIPTION
    This script connects to Exchange Online and analyzes message trace data for the last 30 days
    to identify all users sending or receiving emails using a specific domain. This helps administrators
    understand the impact before removing an accepted domain from their M365 tenant.

    The script does NOT use Graph API or require app registration - it uses Exchange Online PowerShell only.

.PARAMETER Domain
    The domain to analyze (e.g., "xyz.com"). Required parameter.

.PARAMETER OutputPath
    Directory path where the reports will be saved. Defaults to current directory.

.PARAMETER Days
    Number of days to analyze (1-30). Defaults to 30 days. Note: Message trace limited to 10 days for detailed data.

.PARAMETER ExportFormat
    Format for the export: "HTML", "CSV", or "Both". Defaults to "Both".

.EXAMPLE
    .\Get-DomainEmailUsageReport.ps1 -Domain "xyz.com"

.EXAMPLE
    .\Get-DomainEmailUsageReport.ps1 -Domain "xyz.com" -OutputPath "C:\Reports" -Days 10

.NOTES
    Author: M365 Analysis Tool
    Version: 1.2
    Requires: Exchange Online Management PowerShell Module
    Permissions: Exchange Administrator or Global Administrator
    Note: Uses Get-MessageTraceV2 with automatic pagination (no PageSize/Page params)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Domain to analyze (e.g., xyz.com)")]
    [ValidateNotNullOrEmpty()]
    [string]$Domain,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$Days = 10,

    [Parameter(Mandatory = $false)]
    [ValidateSet("HTML", "CSV", "Both")]
    [string]$ExportFormat = "Both"
)

#Requires -Version 5.1

# Global variables
$script:TotalMessagesFound = 0
$script:UniqueUsersCount = 0
$script:StartDate = (Get-Date).AddDays(-$Days)
$script:EndDate = Get-Date

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "Info"    { "Cyan" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
        "Success" { "Green" }
    }

    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Install-RequiredModules {
    Write-Log "Checking required PowerShell modules..." -Level "Info"

    $requiredModule = "ExchangeOnlineManagement"

    try {
        if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
            Write-Log "Installing $requiredModule module..." -Level "Warning"
            Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
            Write-Log "$requiredModule module installed successfully." -Level "Success"
        } else {
            Write-Log "$requiredModule module is already installed." -Level "Success"
        }

        Import-Module $requiredModule -ErrorAction Stop
        Write-Log "Module imported successfully." -Level "Success"
        return $true
    }
    catch {
        Write-Log "Failed to install or import required modules: $_" -Level "Error"
        return $false
    }
}

function Connect-ToExchangeOnline {
    Write-Log "Connecting to Exchange Online..." -Level "Info"

    try {
        # Check if already connected
        $existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue

        if ($existingConnection) {
            Write-Log "Already connected to Exchange Online." -Level "Success"
            return $true
        }

        # Connect to Exchange Online
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Successfully connected to Exchange Online." -Level "Success"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $_" -Level "Error"
        return $false
    }
}

function Get-DomainEmailActivity {
    param(
        [string]$DomainToAnalyze,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    Write-Log "Analyzing email activity for domain: $DomainToAnalyze" -Level "Info"
    Write-Log "Date range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))" -Level "Info"
    Write-Log "Note: This may take several minutes depending on email volume..." -Level "Warning"

    $activityData = @{}

    try {
        # Search for messages where sender OR recipient contains the domain
        Write-Log "Querying message trace data (this may take a while)..." -Level "Info"

        # Due to message trace limitations, we'll query in smaller chunks
        $currentStart = $StartDate
        $allMessages = @()

        while ($currentStart -lt $EndDate) {
            $currentEnd = $currentStart.AddDays(1)
            if ($currentEnd -gt $EndDate) { $currentEnd = $EndDate }

            Write-Log "Fetching messages from $($currentStart.ToString('yyyy-MM-dd HH:mm')) to $($currentEnd.ToString('yyyy-MM-dd HH:mm'))..." -Level "Info"

            # Get messages sent FROM the domain using Get-MessageTraceV2
            # Note: Get-MessageTraceV2 handles pagination automatically
            $sentMessages = Get-MessageTraceV2 -StartDate $currentStart -EndDate $currentEnd |
                Where-Object { $_.SenderAddress -like "*@$DomainToAnalyze" }

            # Get messages sent TO the domain using Get-MessageTraceV2
            $receivedMessages = Get-MessageTraceV2 -StartDate $currentStart -EndDate $currentEnd |
                Where-Object { $_.RecipientAddress -like "*@$DomainToAnalyze" }

            $allMessages += $sentMessages
            $allMessages += $receivedMessages

            $currentStart = $currentEnd
            Start-Sleep -Milliseconds 500  # Rate limiting
        }

        Write-Log "Found $($allMessages.Count) total messages involving domain $DomainToAnalyze" -Level "Success"
        $script:TotalMessagesFound = $allMessages.Count

        # Process messages and aggregate by user
        foreach ($message in $allMessages) {
            # Process sender
            if ($message.SenderAddress -like "*@$DomainToAnalyze") {
                $email = $message.SenderAddress.ToLower()

                if (-not $activityData.ContainsKey($email)) {
                    $activityData[$email] = [PSCustomObject]@{
                        EmailAddress = $email
                        SentCount = 0
                        ReceivedCount = 0
                        TotalMessages = 0
                        FirstSeen = $message.Received
                        LastSeen = $message.Received
                        Status = $message.Status
                        MessageIds = @()
                    }
                }

                $activityData[$email].SentCount++
                $activityData[$email].TotalMessages++

                if ($message.Received -lt $activityData[$email].FirstSeen) {
                    $activityData[$email].FirstSeen = $message.Received
                }
                if ($message.Received -gt $activityData[$email].LastSeen) {
                    $activityData[$email].LastSeen = $message.Received
                }
            }

            # Process recipient
            if ($message.RecipientAddress -like "*@$DomainToAnalyze") {
                $email = $message.RecipientAddress.ToLower()

                if (-not $activityData.ContainsKey($email)) {
                    $activityData[$email] = [PSCustomObject]@{
                        EmailAddress = $email
                        SentCount = 0
                        ReceivedCount = 0
                        TotalMessages = 0
                        FirstSeen = $message.Received
                        LastSeen = $message.Received
                        Status = $message.Status
                        MessageIds = @()
                    }
                }

                $activityData[$email].ReceivedCount++
                $activityData[$email].TotalMessages++

                if ($message.Received -lt $activityData[$email].FirstSeen) {
                    $activityData[$email].FirstSeen = $message.Received
                }
                if ($message.Received -gt $activityData[$email].LastSeen) {
                    $activityData[$email].LastSeen = $message.Received
                }
            }
        }

        $script:UniqueUsersCount = $activityData.Count
        Write-Log "Found $($script:UniqueUsersCount) unique email addresses using domain $DomainToAnalyze" -Level "Success"

        # Convert to array and sort by total messages (descending)
        $sortedData = $activityData.Values | Sort-Object -Property TotalMessages -Descending

        return $sortedData
    }
    catch {
        Write-Log "Error analyzing email activity: $_" -Level "Error"
        return $null
    }
}

function Get-MailboxDetails {
    param(
        [array]$EmailAddresses
    )

    Write-Log "Enriching data with mailbox details..." -Level "Info"

    $enrichedData = @()
    $counter = 0

    foreach ($address in $EmailAddresses) {
        $counter++
        $percentComplete = [math]::Round(($counter / $EmailAddresses.Count) * 100)
        Write-Progress -Activity "Enriching mailbox data" -Status "Processing $($address.EmailAddress)" -PercentComplete $percentComplete

        try {
            # Try to get mailbox information
            $mailbox = Get-Mailbox -Identity $address.EmailAddress -ErrorAction SilentlyContinue

            $enriched = [PSCustomObject]@{
                EmailAddress = $address.EmailAddress
                DisplayName = if ($mailbox) { $mailbox.DisplayName } else { "N/A" }
                MailboxType = if ($mailbox) { $mailbox.RecipientTypeDetails } else { "Unknown/External" }
                SentCount = $address.SentCount
                ReceivedCount = $address.ReceivedCount
                TotalMessages = $address.TotalMessages
                FirstActivity = $address.FirstSeen.ToString("yyyy-MM-dd HH:mm")
                LastActivity = $address.LastSeen.ToString("yyyy-MM-dd HH:mm")
                IsActive = $true
                PrimarySmtpAddress = if ($mailbox) { $mailbox.PrimarySmtpAddress } else { "N/A" }
            }

            $enrichedData += $enriched
        }
        catch {
            # If mailbox doesn't exist, still add the data
            $enriched = [PSCustomObject]@{
                EmailAddress = $address.EmailAddress
                DisplayName = "N/A"
                MailboxType = "Unknown/External"
                SentCount = $address.SentCount
                ReceivedCount = $address.ReceivedCount
                TotalMessages = $address.TotalMessages
                FirstActivity = $address.FirstSeen.ToString("yyyy-MM-dd HH:mm")
                LastActivity = $address.LastSeen.ToString("yyyy-MM-dd HH:mm")
                IsActive = $true
                PrimarySmtpAddress = "N/A"
            }

            $enrichedData += $enriched
        }
    }

    Write-Progress -Activity "Enriching mailbox data" -Completed
    Write-Log "Mailbox enrichment completed for $($enrichedData.Count) addresses." -Level "Success"

    return $enrichedData
}

function Export-HTMLReport {
    param(
        [array]$Data,
        [string]$OutputPath,
        [string]$Domain,
        [int]$Days
    )

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $fileName = "DomainEmailUsageReport_${Domain}_${timestamp}.html"
    $filePath = Join-Path -Path $OutputPath -ChildPath $fileName

    Write-Log "Generating HTML report..." -Level "Info"

    # Calculate summary statistics
    $totalUsers = $Data.Count
    $totalMessages = ($Data | Measure-Object -Property TotalMessages -Sum).Sum
    $totalSent = ($Data | Measure-Object -Property SentCount -Sum).Sum
    $totalReceived = ($Data | Measure-Object -Property ReceivedCount -Sum).Sum
    $activeMailboxes = ($Data | Where-Object { $_.MailboxType -ne "Unknown/External" }).Count

    # Build HTML content
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Domain Email Usage Report - $Domain</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f5f5;
            padding: 20px;
            color: #333;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #d32f2f 0%, #c62828 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            font-size: 28px;
            margin-bottom: 10px;
        }
        .header .subtitle {
            font-size: 14px;
            opacity: 0.9;
        }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #fafafa;
            border-bottom: 2px solid #e0e0e0;
        }
        .summary-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid #d32f2f;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .summary-card .label {
            font-size: 12px;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }
        .summary-card .value {
            font-size: 32px;
            font-weight: bold;
            color: #d32f2f;
        }
        .impact-notice {
            background: #fff3cd;
            border: 2px solid #ffc107;
            border-radius: 8px;
            padding: 20px;
            margin: 20px 30px;
        }
        .impact-notice h3 {
            color: #856404;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
        }
        .impact-notice h3::before {
            content: "⚠️";
            margin-right: 10px;
            font-size: 24px;
        }
        .impact-notice p {
            color: #856404;
            line-height: 1.6;
        }
        .content {
            padding: 30px;
        }
        h2 {
            color: #d32f2f;
            margin-bottom: 20px;
            font-size: 22px;
            border-bottom: 2px solid #d32f2f;
            padding-bottom: 10px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 14px;
        }
        thead {
            background: #d32f2f;
            color: white;
            position: sticky;
            top: 0;
        }
        th {
            padding: 12px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.5px;
        }
        td {
            padding: 12px;
            border-bottom: 1px solid #e0e0e0;
        }
        tbody tr:hover {
            background: #f5f5f5;
        }
        tbody tr:nth-child(odd) {
            background: #fafafa;
        }
        .high-usage {
            background: #ffebee !important;
        }
        .badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: bold;
            text-transform: uppercase;
        }
        .badge-mailbox {
            background: #e3f2fd;
            color: #1976d2;
        }
        .badge-external {
            background: #fce4ec;
            color: #c2185b;
        }
        .footer {
            text-align: center;
            padding: 20px;
            background: #fafafa;
            color: #666;
            font-size: 12px;
            border-top: 1px solid #e0e0e0;
        }
        .metric-inline {
            font-weight: bold;
            color: #d32f2f;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔍 Domain Email Usage Analysis Report</h1>
            <div class="subtitle">Domain: <strong>$Domain</strong> | Analysis Period: Last $Days Days | Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</div>
        </div>

        <div class="summary">
            <div class="summary-card">
                <div class="label">Total Affected Users</div>
                <div class="value">$totalUsers</div>
            </div>
            <div class="summary-card">
                <div class="label">Total Messages</div>
                <div class="value">$totalMessages</div>
            </div>
            <div class="summary-card">
                <div class="label">Messages Sent</div>
                <div class="value">$totalSent</div>
            </div>
            <div class="summary-card">
                <div class="label">Messages Received</div>
                <div class="value">$totalReceived</div>
            </div>
            <div class="summary-card">
                <div class="label">Active Mailboxes</div>
                <div class="value">$activeMailboxes</div>
            </div>
        </div>

        <div class="impact-notice">
            <h3>Impact Assessment</h3>
            <p><strong>CRITICAL:</strong> Removing domain <strong>$Domain</strong> will affect <span class="metric-inline">$totalUsers unique email addresses</span> that have sent or received <span class="metric-inline">$totalMessages messages</span> in the last $Days days.</p>
            <p style="margin-top: 10px;"><strong>Recommended Actions:</strong></p>
            <ul style="margin-left: 20px; margin-top: 5px; line-height: 1.8;">
                <li>Review all affected users below and migrate them to a different domain</li>
                <li>Update email addresses in all systems and applications</li>
                <li>Notify affected users about the upcoming change</li>
                <li>Set up email forwarding rules before domain removal</li>
                <li>Verify that no active mailboxes will be orphaned</li>
            </ul>
        </div>

        <div class="content">
            <h2>📊 Affected Users (Sorted by Activity)</h2>
            <table>
                <thead>
                    <tr>
                        <th>Rank</th>
                        <th>Email Address</th>
                        <th>Display Name</th>
                        <th>Mailbox Type</th>
                        <th>Total Messages</th>
                        <th>Sent</th>
                        <th>Received</th>
                        <th>First Activity</th>
                        <th>Last Activity</th>
                        <th>Primary SMTP</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Add table rows
    $rank = 1
    foreach ($item in $Data) {
        $rowClass = if ($item.TotalMessages -gt 100) { ' class="high-usage"' } else { '' }
        $badgeClass = if ($item.MailboxType -ne "Unknown/External") { "badge-mailbox" } else { "badge-external" }

        $htmlContent += @"
                    <tr$rowClass>
                        <td><strong>$rank</strong></td>
                        <td><strong>$($item.EmailAddress)</strong></td>
                        <td>$($item.DisplayName)</td>
                        <td><span class="badge $badgeClass">$($item.MailboxType)</span></td>
                        <td><strong>$($item.TotalMessages)</strong></td>
                        <td>$($item.SentCount)</td>
                        <td>$($item.ReceivedCount)</td>
                        <td>$($item.FirstActivity)</td>
                        <td>$($item.LastActivity)</td>
                        <td>$($item.PrimarySmtpAddress)</td>
                    </tr>
"@
        $rank++
    }

    $htmlContent += @"
                </tbody>
            </table>
        </div>

        <div class="footer">
            <p><strong>M365 Domain Email Usage Report</strong> | Generated with Exchange Online PowerShell</p>
            <p>Report Path: $filePath</p>
        </div>
    </div>
</body>
</html>
"@

    try {
        $htmlContent | Out-File -FilePath $filePath -Encoding UTF8
        Write-Log "HTML report saved to: $filePath" -Level "Success"
        return $filePath
    }
    catch {
        Write-Log "Failed to save HTML report: $_" -Level "Error"
        return $null
    }
}

function Export-CSVReport {
    param(
        [array]$Data,
        [string]$OutputPath,
        [string]$Domain
    )

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $fileName = "DomainEmailUsageReport_${Domain}_${timestamp}.csv"
    $filePath = Join-Path -Path $OutputPath -ChildPath $fileName

    Write-Log "Generating CSV report..." -Level "Info"

    try {
        $Data | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
        Write-Log "CSV report saved to: $filePath" -Level "Success"
        return $filePath
    }
    catch {
        Write-Log "Failed to save CSV report: $_" -Level "Error"
        return $null
    }
}

#endregion

#region Main Execution

function Main {
    Write-Log "========================================" -Level "Info"
    Write-Log "Domain Email Usage Analysis Tool" -Level "Info"
    Write-Log "========================================" -Level "Info"
    Write-Log "Domain: $Domain" -Level "Info"
    Write-Log "Analysis Period: Last $Days days" -Level "Info"
    Write-Log "Output Path: $OutputPath" -Level "Info"
    Write-Log "========================================" -Level "Info"

    try {
        # Normalize domain (remove @ if present)
        $script:Domain = $Domain.TrimStart('@')

        # Validate output path
        if (-not (Test-Path -Path $OutputPath)) {
            Write-Log "Output path does not exist. Creating directory: $OutputPath" -Level "Warning"
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }

        # Step 1: Install required modules
        if (-not (Install-RequiredModules)) {
            throw "Failed to install required modules. Exiting."
        }

        # Step 2: Connect to Exchange Online
        if (-not (Connect-ToExchangeOnline)) {
            throw "Failed to connect to Exchange Online. Exiting."
        }

        # Step 3: Analyze email activity
        $emailActivity = Get-DomainEmailActivity -DomainToAnalyze $script:Domain -StartDate $script:StartDate -EndDate $script:EndDate

        if ($null -eq $emailActivity -or $emailActivity.Count -eq 0) {
            Write-Log "No email activity found for domain $script:Domain in the specified period." -Level "Warning"
            Write-Log "This could mean:" -Level "Info"
            Write-Log "  - The domain is not being used for email" -Level "Info"
            Write-Log "  - No messages were sent/received in the last $Days days" -Level "Info"
            Write-Log "  - The domain might already be removed or not configured" -Level "Info"

            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            return 0
        }

        # Step 4: Enrich with mailbox details
        $enrichedData = Get-MailboxDetails -EmailAddresses $emailActivity

        # Step 5: Generate reports
        if ($ExportFormat -eq "HTML" -or $ExportFormat -eq "Both") {
            $htmlPath = Export-HTMLReport -Data $enrichedData -OutputPath $OutputPath -Domain $script:Domain -Days $Days
        }

        if ($ExportFormat -eq "CSV" -or $ExportFormat -eq "Both") {
            $csvPath = Export-CSVReport -Data $enrichedData -OutputPath $OutputPath -Domain $script:Domain
        }

        # Display summary
        Write-Log "========================================" -Level "Success"
        Write-Log "Analysis Complete!" -Level "Success"
        Write-Log "========================================" -Level "Success"
        Write-Log "Total messages found: $script:TotalMessagesFound" -Level "Info"
        Write-Log "Unique email addresses: $script:UniqueUsersCount" -Level "Info"
        Write-Log "Top 5 most active users:" -Level "Info"

        $top5 = $enrichedData | Select-Object -First 5
        foreach ($user in $top5) {
            Write-Log "  - $($user.EmailAddress): $($user.TotalMessages) messages (Sent: $($user.SentCount), Received: $($user.ReceivedCount))" -Level "Info"
        }

        Write-Log "========================================" -Level "Success"

        # Disconnect from Exchange Online
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Disconnected from Exchange Online." -Level "Info"

        return 0
    }
    catch {
        Write-Log "An error occurred during execution: $_" -Level "Error"
        Write-Log $_.ScriptStackTrace -Level "Error"

        # Ensure we disconnect even on error
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

        return 1
    }
}

# Execute main function
$exitCode = Main
exit $exitCode

#endregion
