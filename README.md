# M365 Reporting and Analysis Tools

A collection of PowerShell scripts for Microsoft 365 reporting and analysis tasks.

## Available Tools

1. **Guest Accounts Activity Report** - Analyze and monitor guest account activity
2. **Domain Email Usage Report** - Assess email usage impact before domain removal

## Features

This script provides detailed information about guest accounts in your M365 tenant:

- **Guest User List**: Complete inventory of all guest accounts
- **Sign-In Activity**: Last sign-in date and days since last activity
- **Access Permissions**: Group memberships and access rights
- **Inactive Accounts**: Identify guest accounts that haven't signed in for a specified period
- **Invitation Status**: Track pending and accepted invitations
- **Account Status**: Enabled/disabled account information
- **Multiple Export Formats**: HTML (formatted report) and CSV (data analysis)

## Prerequisites

### Required PowerShell Modules

The script will automatically install the following modules if not present:

- Microsoft.Graph.Users
- Microsoft.Graph.Groups
- Microsoft.Graph.Identity.DirectoryManagement
- Microsoft.Graph.Reports

### Required Permissions

You need to have one of the following admin roles in M365:

- Global Administrator
- Global Reader
- Reports Reader
- User Administrator

The script requires the following Microsoft Graph API permissions:

- `User.Read.All` - Read all users' full profiles
- `AuditLog.Read.All` - Read audit log data
- `Directory.Read.All` - Read directory data
- `Group.Read.All` - Read all groups

## Installation

1. Download the script to your local machine
2. Ensure you have PowerShell 5.1 or higher
3. Run PowerShell as Administrator (for module installation)

## Usage

### Basic Usage

Run the script with default settings (90 days inactivity threshold, both HTML and CSV output):

```powershell
.\Get-GuestAccountsReport.ps1
```

### Advanced Usage

#### Specify Output Path

```powershell
.\Get-GuestAccountsReport.ps1 -OutputPath "C:\Reports"
```

#### Custom Inactivity Threshold

Set inactive threshold to 60 days:

```powershell
.\Get-GuestAccountsReport.ps1 -InactiveDays 60
```

#### Choose Export Format

Export only HTML:

```powershell
.\Get-GuestAccountsReport.ps1 -ExportFormat HTML
```

Export only CSV:

```powershell
.\Get-GuestAccountsReport.ps1 -ExportFormat CSV
```

#### Combined Parameters

```powershell
.\Get-GuestAccountsReport.ps1 -OutputPath "C:\Reports" -InactiveDays 30 -ExportFormat Both
```

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `OutputPath` | String | Current directory | Path where report files will be saved |
| `InactiveDays` | Integer | 90 | Number of days to consider an account inactive |
| `ExportFormat` | String | Both | Export format: HTML, CSV, or Both |

## Output

### HTML Report

The HTML report includes:

- **Summary Dashboard**: Quick statistics including total guests, active/inactive counts, disabled accounts, and pending invitations
- **Detailed Table**: Sortable table with all guest account information
- **Visual Indicators**: Color-coded rows for inactive and disabled accounts
- **Professional Styling**: Clean, modern design suitable for executive reporting

### CSV Report

The CSV export includes all data fields for further analysis in Excel or other tools:

- DisplayName
- UserPrincipalName
- Email
- AccountEnabled
- InvitationStatus
- CreatedDate
- InvitationAcceptedDate
- LastSignInDate
- DaysSinceLastSignIn
- IsInactive
- GroupMemberships

## Report Data Fields

### Guest Account Information

- **Display Name**: Full name of the guest user
- **User Principal Name**: UPN of the guest account
- **Email**: Email address of the guest user
- **Account Status**: Whether the account is enabled or disabled
- **Invitation Status**: Accepted or Pending
- **Created Date**: When the guest account was created
- **Invitation Accepted Date**: When the guest accepted the invitation
- **Last Sign-In Date**: Most recent sign-in timestamp
- **Days Since Last Sign-In**: Number of days since last activity
- **Is Inactive**: Boolean flag based on the InactiveDays parameter
- **Group Memberships**: List of all groups the guest is a member of

## Examples

### Example 1: Monthly Review

Generate a monthly review of guest accounts with 30-day inactivity threshold:

```powershell
.\Get-GuestAccountsReport.ps1 -OutputPath "C:\MonthlyReports" -InactiveDays 30
```

### Example 2: Quarterly Audit

Generate a quarterly audit with 90-day threshold, CSV only for compliance:

```powershell
.\Get-GuestAccountsReport.ps1 -OutputPath "C:\AuditReports\Q1" -InactiveDays 90 -ExportFormat CSV
```

### Example 3: Executive Summary

Create an HTML report for executive review:

```powershell
.\Get-GuestAccountsReport.ps1 -OutputPath "C:\ExecutiveReports" -ExportFormat HTML
```

## Troubleshooting

### Module Installation Issues

If you encounter module installation errors:

1. Run PowerShell as Administrator
2. Ensure you have internet connectivity
3. Check execution policy: `Get-ExecutionPolicy`
4. If needed, set execution policy: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

### Authentication Issues

If you cannot connect to Microsoft Graph:

1. Ensure you have appropriate admin permissions
2. Check that your account has the required roles
3. Verify MFA if enabled
4. Try disconnecting first: `Disconnect-MgGraph`

### Permission Errors

If you get permission-related errors:

1. Ensure you consented to all required permissions during first login
2. Check with your Global Administrator if permissions are restricted
3. Review the consent screen carefully when authenticating

## Security Considerations

- The script requires read-only permissions and does not modify any data
- Always review the permissions requested during authentication
- Store reports securely as they contain sensitive user information
- Consider scheduling regular reports to monitor guest account hygiene
- Use the report to identify and remove inactive guest accounts

## Best Practices

1. **Regular Monitoring**: Run the report monthly or quarterly
2. **Cleanup Inactive Accounts**: Use the report to identify guests who should be removed
3. **Audit Group Memberships**: Review what resources guests have access to
4. **Track Pending Invitations**: Follow up on invitations that haven't been accepted
5. **Archive Reports**: Keep historical reports for compliance and trend analysis

## Scheduling the Report

### Using Task Scheduler

Create a scheduled task to run the report automatically:

```powershell
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\Scripts\Get-GuestAccountsReport.ps1 -OutputPath C:\Reports"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 6am
Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "M365 Guest Report" -Description "Weekly guest account activity report"
```

## Support

For issues, questions, or contributions, please refer to the repository documentation.

## License

This script is provided as-is for use in your M365 environment.

## Version History

- **1.0.0** (2025-11-14): Initial release
  - Guest account enumeration
  - Sign-in activity tracking
  - Group membership reporting
  - Inactive account detection
  - HTML and CSV export formats
  - Professional report formatting

---

# Domain Email Usage Report

## Overview

The **Get-DomainEmailUsageReport.ps1** script helps you analyze email usage for a specific domain before removing it from your M365 tenant. This tool is essential for understanding the impact of domain removal by identifying all users who are actively sending or receiving emails using that domain.

## Use Case

Before removing an accepted domain from Exchange Online mail flow, you need to:

- **Identify Impact**: Understand who will be affected by the domain removal
- **Measure Activity**: See how many emails are being sent/received on that domain
- **Plan Migration**: Determine which users need to be migrated to a different domain
- **Avoid Service Disruption**: Ensure no active email addresses will be orphaned

## Features

- ✅ **No Graph API Required**: Uses Exchange Online PowerShell only (no app registration needed)
- ✅ **Message Trace Analysis**: Analyzes actual email traffic from the last 1-10 days
- ✅ **Comprehensive Coverage**: Identifies both senders and recipients using the domain
- ✅ **Mailbox Enrichment**: Retrieves mailbox details for each affected user
- ✅ **Activity Ranking**: Sorts users by total message count (top users first)
- ✅ **Impact Assessment**: Provides visual summary of potential impact
- ✅ **Professional Reports**: Generates HTML and CSV reports with detailed statistics

## Prerequisites

### Required PowerShell Module

The script will automatically install:

- **ExchangeOnlineManagement** - Exchange Online PowerShell module

### Required Permissions

You need one of the following admin roles:

- Exchange Administrator
- Global Administrator
- Global Reader (for read-only analysis)

### Important Notes

- **Message Trace Limitation**: Exchange Online message trace is limited to 10 days of detailed data
- **Processing Time**: Large email volumes may take several minutes to process
- **Rate Limiting**: The script includes delays to avoid throttling

## Installation

1. Download `Get-DomainEmailUsageReport.ps1` to your local machine
2. Ensure you have PowerShell 5.1 or higher
3. Run PowerShell as Administrator (for module installation)

## Usage

### Basic Usage

Analyze domain usage for the last 10 days (default):

```powershell
.\Get-DomainEmailUsageReport.ps1 -Domain "xyz.com"
```

### Specify Output Path

```powershell
.\Get-DomainEmailUsageReport.ps1 -Domain "xyz.com" -OutputPath "C:\Reports"
```

### Custom Analysis Period

Analyze the last 7 days:

```powershell
.\Get-DomainEmailUsageReport.ps1 -Domain "xyz.com" -Days 7
```

### Export Format Options

Export only HTML:

```powershell
.\Get-DomainEmailUsageReport.ps1 -Domain "xyz.com" -ExportFormat HTML
```

Export only CSV:

```powershell
.\Get-DomainEmailUsageReport.ps1 -Domain "xyz.com" -ExportFormat CSV
```

### Combined Parameters

```powershell
.\Get-DomainEmailUsageReport.ps1 -Domain "xyz.com" -OutputPath "C:\Reports" -Days 10 -ExportFormat Both
```

## Parameters

| Parameter | Type | Default | Required | Description |
|-----------|------|---------|----------|-------------|
| `Domain` | String | - | **Yes** | Domain to analyze (e.g., "xyz.com") |
| `OutputPath` | String | Current directory | No | Path where reports will be saved |
| `Days` | Integer | 10 | No | Number of days to analyze (1-10) |
| `ExportFormat` | String | Both | No | Export format: HTML, CSV, or Both |

## Output

### HTML Report

The HTML report includes:

- **Executive Summary Dashboard**:
  - Total affected users
  - Total messages (sent + received)
  - Messages sent from the domain
  - Messages received to the domain
  - Active mailboxes count

- **Impact Assessment Notice**:
  - Critical warning about removal impact
  - Number of affected users and messages
  - Recommended actions before domain removal

- **Detailed User Table** (sorted by activity):
  - Rank (by message volume)
  - Email address using the domain
  - Display name
  - Mailbox type (UserMailbox, SharedMailbox, External, etc.)
  - Total messages
  - Sent count
  - Received count
  - First activity date/time
  - Last activity date/time
  - Primary SMTP address

- **Visual Indicators**:
  - High-usage rows highlighted in red (>100 messages)
  - Mailbox type badges (internal vs external)
  - Color-coded sections for easy navigation

### CSV Report

The CSV export includes all data fields for analysis:

- EmailAddress
- DisplayName
- MailboxType
- SentCount
- ReceivedCount
- TotalMessages
- FirstActivity
- LastActivity
- IsActive
- PrimarySmtpAddress

## Report Data Fields

### Email Activity Information

- **Email Address**: The email address using the domain (e.g., user@xyz.com)
- **Display Name**: Full name of the user/mailbox
- **Mailbox Type**: Type of mailbox (UserMailbox, SharedMailbox, External, etc.)
- **Total Messages**: Combined sent + received message count
- **Sent Count**: Number of messages sent FROM this address
- **Received Count**: Number of messages received TO this address
- **First Activity**: Earliest message date/time in the analysis period
- **Last Activity**: Most recent message date/time in the analysis period
- **Primary SMTP Address**: The primary SMTP address of the mailbox (if different)

## Example Scenarios

### Scenario 1: Pre-Removal Assessment

Before removing domain "oldcompany.com":

```powershell
.\Get-DomainEmailUsageReport.ps1 -Domain "oldcompany.com" -OutputPath "C:\DomainRemoval" -Days 10
```

**Result**: Comprehensive report showing all active users on that domain

### Scenario 2: Quick Check

Quick 3-day check to see if domain is still in use:

```powershell
.\Get-DomainEmailUsageReport.ps1 -Domain "test.com" -Days 3 -ExportFormat CSV
```

**Result**: Fast CSV export for spreadsheet analysis

### Scenario 3: Executive Briefing

Generate a professional HTML report for stakeholders:

```powershell
.\Get-DomainEmailUsageReport.ps1 -Domain "legacy.com" -OutputPath "C:\ExecutiveReports" -ExportFormat HTML
```

**Result**: Polished HTML report with impact assessment

## Understanding the Results

### If Users Are Found

The report will show:
- Total number of affected users
- Message volume per user (sorted by most active)
- Mailbox details for migration planning

**Action Required**:
- Migrate users to a different domain
- Update email addresses in all systems
- Set up forwarding rules
- Notify affected users

### If No Users Are Found

The script will report:
```
No email activity found for domain xyz.com in the specified period.
```

**This could mean**:
- Domain is not being used for email
- No messages in the last X days
- Domain already removed or not configured
- Safe to proceed with domain removal

## Best Practices

1. **Run Before Domain Removal**: Always analyze usage before making changes
2. **Use Maximum Days**: Use 10 days for comprehensive analysis
3. **Review Top Users**: Focus on high-volume users first for migration
4. **Check Mailbox Types**: Identify shared mailboxes and distribution lists
5. **Archive Reports**: Keep reports for audit trail and compliance
6. **Verify Primary Addresses**: Ensure users have alternative addresses
7. **Set Up Forwarding**: Configure email forwarding before removal
8. **Notify Users**: Give affected users advance notice
9. **Test Migration**: Migrate test users first
10. **Monitor Post-Removal**: Check for bounce-backs after removal

## Migration Workflow

1. **Analyze**: Run this script to identify affected users
2. **Plan**: Review report and plan migration strategy
3. **Communicate**: Notify affected users about upcoming changes
4. **Migrate**: Update email addresses to new domain
5. **Forward**: Set up email forwarding rules
6. **Verify**: Confirm all users have working email
7. **Remove**: Remove the old domain from tenant
8. **Monitor**: Watch for any issues post-removal

## Troubleshooting

### Module Installation Issues

If module installation fails:

```powershell
# Run as Administrator
Install-Module -Name ExchangeOnlineManagement -Force
```

### Authentication Issues

If you cannot connect to Exchange Online:

```powershell
# Try manual connection
Connect-ExchangeOnline -UserPrincipalName admin@yourtenant.com
```

### No Data Returned

If no messages are found:

1. Verify the domain is spelled correctly
2. Confirm the domain is configured in Exchange
3. Check if domain has been used recently
4. Try reducing the -Days parameter
5. Verify your permissions

### Slow Performance

If the script is slow:

1. Reduce the -Days parameter (try 3-5 days)
2. Run during off-peak hours
3. Check Exchange Online service health
4. Large message volumes take longer to process

## Security Considerations

- Script uses **read-only operations** (no data modification)
- Requires Exchange Administrator permissions
- Reports contain **sensitive email addresses** - store securely
- Review permissions before granting access
- Use least-privilege principle
- Audit trail maintained by Exchange Online

## Technical Details

### How It Works

1. **Connection**: Connects to Exchange Online via ExchangeOnline module
2. **Message Trace**: Queries Get-MessageTraceV2 for specified date range
3. **Filtering**: Filters messages where sender OR recipient matches domain
4. **Aggregation**: Groups messages by email address
5. **Enrichment**: Retrieves mailbox details for each address
6. **Sorting**: Sorts users by total message count (descending)
7. **Reporting**: Generates formatted HTML and/or CSV reports

### Message Trace Limitations

- **10-day maximum**: Detailed message trace limited to 10 days
- **Rate limiting**: API calls are throttled
- **Processing time**: Large volumes may take 5-10 minutes
- **Paging**: Results returned in pages of 5000

### Performance Optimization

- Script processes data day-by-day to avoid timeouts
- Includes 500ms delays to prevent throttling
- Uses pagination for large result sets
- Minimal memory footprint

## Scheduling

### Using Task Scheduler

Run weekly analysis of critical domains:

```powershell
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\Scripts\Get-DomainEmailUsageReport.ps1 -Domain 'oldcompany.com' -OutputPath 'C:\Reports'"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 7am
Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "Domain Usage Check" -Description "Weekly domain email usage analysis"
```

## Support

For issues or questions about this script, please refer to the repository documentation or open an issue.

## License

This script is provided as-is for use in your M365 environment.

## Version History

- **1.1** (2026-01-12): Updated to use Get-MessageTraceV2
  - Replaced deprecated Get-MessageTrace with Get-MessageTraceV2
  - Get-MessageTrace was deprecated as of September 1st, 2025

- **1.0** (2026-01-12): Initial release
  - Exchange Online message trace analysis
  - Domain email usage reporting
  - HTML and CSV export formats

---

## Comparison: Guest Report vs Domain Report

| Feature | Guest Accounts Report | Domain Email Usage Report |
|---------|----------------------|---------------------------|
| **Purpose** | Monitor guest access | Assess domain removal impact |
| **API Used** | Microsoft Graph API | Exchange Online PowerShell |
| **Data Source** | Azure AD / Entra ID | Exchange message trace |
| **Time Range** | All-time + sign-in activity | Last 1-10 days |
| **Authentication** | Connect-MgGraph | Connect-ExchangeOnline |
| **Focus** | User accounts & permissions | Email traffic & usage |
| **Use Case** | Security & compliance | Domain migration planning |
