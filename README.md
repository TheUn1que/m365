# M365 Guest Accounts Activity Report

A comprehensive PowerShell script to generate detailed reports on guest account activity in Microsoft 365.

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
