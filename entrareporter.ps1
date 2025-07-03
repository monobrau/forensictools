<#
.SYNOPSIS
    Generates comprehensive forensic reports for specified user accounts by gathering data
    from both Microsoft Graph and Exchange Online with enhanced license detection and multi-user support.

.DESCRIPTION
    This script is a command-line tool designed for forensic investigations. Upon execution, it
    will prompt for an interactive login. It then fetches a list of all users and presents them
    in a selectable grid with multi-select capability. After users are selected, you can choose
    to generate individual reports or a consolidated report. The script then connects to Microsoft 365 
    to collect detailed information about the selected users and saves it to the chosen location.

    Enhanced features:
    - Improved license detection with comprehensive SKU mapping
    - Support for multiple user selection
    - Individual or consolidated report options
    - Per-user license analysis
    - Better error handling and progress tracking

    All collected data is compiled into Excel spreadsheets, with each category of
    information placed in a separate, clearly labeled worksheet.

    This script requires the 'Microsoft.Graph' and 'ImportExcel' modules.

.PARAMETER ConsolidatedReport
    When specified, generates a single consolidated report for all selected users.
    Otherwise, generates individual reports for each user.

.EXAMPLE
    .\Get-UserForensicReport.ps1

    This command will prompt for authentication, display a multi-user selection window,
    ask for report type preference, and generate the requested reports.

.EXAMPLE
    .\Get-UserForensicReport.ps1 -ConsolidatedReport

    Forces consolidated report generation for all selected users.

.NOTES
    Author: Enhanced Version
    Version: 4.0
    - Enhanced license detection with comprehensive SKU mapping and user-level license analysis
    - Added multi-user selection capability
    - Added option for consolidated vs individual reports
    - Improved progress tracking and error handling
    - Added per-user license assignment analysis
    - Better organization of output files

    Prerequisites:
    1. PowerShell 5.1 or later.
    2. The Microsoft.Graph PowerShell module. Install with: Install-Module Microsoft.Graph -Force
    3. The ImportExcel PowerShell module. Install with: Install-Module ImportExcel -Force

    Required Permissions:
    The account running this script will need sufficient permissions in Microsoft 365.
    Roles like 'Global Reader', 'Security Reader', or 'Global Administrator' are recommended.
    The script will prompt for consent for the required API permissions upon first run.
#>
[CmdletBinding()]
param (
    [switch]$ConsolidatedReport
)

# --- ENHANCED LICENSE DETECTION FUNCTIONS ---

function Get-TenantLicenseDetails {
    <#
    .SYNOPSIS
        Provides comprehensive tenant license analysis with detailed SKU mapping
    #>
    
    $licenseMapping = @{
        # Entra ID (Azure AD) Licenses
        'AAD_BASIC'                    = @{ Name = 'Azure Active Directory Basic'; Tier = 'Basic'; LogRetention = 7 }
        'AAD_PREMIUM'                  = @{ Name = 'Azure Active Directory Premium P1'; Tier = 'P1'; LogRetention = 30 }
        'AAD_PREMIUM_P2'               = @{ Name = 'Azure Active Directory Premium P2'; Tier = 'P2'; LogRetention = 30 }
        
        # Microsoft 365 Enterprise
        'ENTERPRISEPACK'               = @{ Name = 'Microsoft 365 E3'; Tier = 'P1'; LogRetention = 30 }
        'ENTERPRISEPREMIUM'            = @{ Name = 'Microsoft 365 E5'; Tier = 'P2'; LogRetention = 30 }
        'ENTERPRISEPREMIUM_NOPSTNCONF' = @{ Name = 'Microsoft 365 E5 (without Audio Conferencing)'; Tier = 'P2'; LogRetention = 30 }
        
        # Education Licenses
        'STANDARDPACK_FACULTY'         = @{ Name = 'Microsoft 365 A1 for Faculty'; Tier = 'Basic'; LogRetention = 7 }
        'ENTERPRISEPACK_FACULTY'       = @{ Name = 'Microsoft 365 A3 for Faculty'; Tier = 'P1'; LogRetention = 30 }
        'ENTERPRISEPREMIUM_FACULTY'    = @{ Name = 'Microsoft 365 A5 for Faculty'; Tier = 'P2'; LogRetention = 30 }
        'STANDARDPACK_STUDENT'         = @{ Name = 'Microsoft 365 A1 for Students'; Tier = 'Basic'; LogRetention = 7 }
        'ENTERPRISEPACK_STUDENT'       = @{ Name = 'Microsoft 365 A3 for Students'; Tier = 'P1'; LogRetention = 30 }
        'ENTERPRISEPREMIUM_STUDENT'    = @{ Name = 'Microsoft 365 A5 for Students'; Tier = 'P2'; LogRetention = 30 }
        
        # Business Licenses
        'O365_BUSINESS_ESSENTIALS'     = @{ Name = 'Microsoft 365 Business Basic'; Tier = 'Basic'; LogRetention = 7 }
        'O365_BUSINESS_PREMIUM'        = @{ Name = 'Microsoft 365 Business Standard'; Tier = 'Basic'; LogRetention = 7 }
        'SPB'                          = @{ Name = 'Microsoft 365 Business Premium'; Tier = 'P1'; LogRetention = 30 }
        
        # Government Licenses
        'ENTERPRISEPACK_GOV'           = @{ Name = 'Microsoft 365 E3 (Government)'; Tier = 'P1'; LogRetention = 30 }
        'ENTERPRISEPREMIUM_GOV'        = @{ Name = 'Microsoft 365 E5 (Government)'; Tier = 'P2'; LogRetention = 30 }
        
        # Standalone Licenses
        'EMS'                          = @{ Name = 'Enterprise Mobility + Security E3'; Tier = 'P1'; LogRetention = 30 }
        'EMSPREMIUM'                   = @{ Name = 'Enterprise Mobility + Security E5'; Tier = 'P2'; LogRetention = 30 }
        'WINDOWS_STORE'                = @{ Name = 'Windows Store for Business'; Tier = 'Basic'; LogRetention = 7 }
        
        # Developer and Trial
        'DEVELOPERPACK'                = @{ Name = 'Microsoft 365 E3 Developer'; Tier = 'P1'; LogRetention = 30 }
        'ENTERPRISEPACK_USGOV_DOD'     = @{ Name = 'Microsoft 365 E3 (DOD)'; Tier = 'P1'; LogRetention = 30 }
        'ENTERPRISEPACK_USGOV_GCCHIGH' = @{ Name = 'Microsoft 365 E3 (GCC High)'; Tier = 'P1'; LogRetention = 30 }
    }

    try {
        Write-Host "Performing comprehensive tenant license analysis..." -ForegroundColor Cyan
        $skus = Get-MgSubscribedSku -All -ErrorAction Stop
        
        $tenantLicenses = @()
        $highestTier = 'Basic'
        $maxLogRetention = 7
        
        foreach ($sku in $skus) {
            $skuInfo = $licenseMapping[$sku.SkuPartNumber]
            if (-not $skuInfo) {
                $skuInfo = @{ 
                    Name = $sku.SkuPartNumber; 
                    Tier = 'Unknown'; 
                    LogRetention = 7 
                }
            }
            
            $licenseDetail = [PSCustomObject]@{
                SkuPartNumber = $sku.SkuPartNumber
                ProductName = $skuInfo.Name
                Tier = $skuInfo.Tier
                LogRetention = $skuInfo.LogRetention
                ConsumedUnits = $sku.ConsumedUnits
                PrepaidUnits = $sku.PrepaidUnits.Enabled
                CapabilityStatus = $sku.CapabilityStatus
            }
            $tenantLicenses += $licenseDetail
            
            # Determine highest tier
            if ($skuInfo.Tier -eq 'P2' -and $highestTier -ne 'P2') {
                $highestTier = 'P2'
                $maxLogRetention = 30
            }
            elseif ($skuInfo.Tier -eq 'P1' -and $highestTier -eq 'Basic') {
                $highestTier = 'P1'
                $maxLogRetention = 30
            }
        }
        
        $summary = [PSCustomObject]@{
            HighestTier = $highestTier
            MaxLogRetention = $maxLogRetention
            TotalLicenses = $tenantLicenses.Count
            P2Licenses = ($tenantLicenses | Where-Object { $_.Tier -eq 'P2' }).Count
            P1Licenses = ($tenantLicenses | Where-Object { $_.Tier -eq 'P1' }).Count
            BasicLicenses = ($tenantLicenses | Where-Object { $_.Tier -eq 'Basic' }).Count
            Details = $tenantLicenses
        }
        
        Write-Host "License Analysis Complete:" -ForegroundColor Green
        Write-Host "  Highest Tier: $($summary.HighestTier)" -ForegroundColor Yellow
        Write-Host "  Max Log Retention: $($summary.MaxLogRetention) days" -ForegroundColor Yellow
        Write-Host "  P2 Licenses: $($summary.P2Licenses)" -ForegroundColor Yellow
        Write-Host "  P1 Licenses: $($summary.P1Licenses)" -ForegroundColor Yellow
        Write-Host "  Basic Licenses: $($summary.BasicLicenses)" -ForegroundColor Yellow
        
        return $summary
    }
    catch {
        Write-Warning "Could not perform comprehensive license analysis. Error: $_"
        return @{
            HighestTier = 'Unknown'
            MaxLogRetention = 7
            TotalLicenses = 0
            P2Licenses = 0
            P1Licenses = 0
            BasicLicenses = 0
            Details = @()
        }
    }
}

function Get-UserLicenseDetails {
    <#
    .SYNOPSIS
        Gets detailed license information for a specific user
    #>
    param(
        [string]$UserId,
        [hashtable]$LicenseMapping
    )
    
    try {
        $userLicenses = Get-MgUserLicenseDetail -UserId $UserId -ErrorAction Stop
        $userLicenseDetails = @()
        
        foreach ($license in $userLicenses) {
            $skuInfo = $LicenseMapping[$license.SkuPartNumber]
            if (-not $skuInfo) {
                $skuInfo = @{ 
                    Name = $license.SkuPartNumber; 
                    Tier = 'Unknown'; 
                    LogRetention = 7 
                }
            }
            
            $userLicenseDetails += [PSCustomObject]@{
                SkuPartNumber = $license.SkuPartNumber
                ProductName = $skuInfo.Name
                Tier = $skuInfo.Tier
                LogRetention = $skuInfo.LogRetention
                ServicePlans = ($license.ServicePlans | Where-Object { $_.ProvisioningStatus -eq 'Success' } | Select-Object -ExpandProperty ServicePlanName) -join ', '
            }
        }
        
        return $userLicenseDetails
    }
    catch {
        Write-Warning "Could not retrieve license details for user $UserId. Error: $_"
        return @()
    }
}

# --- SCRIPT INITIALIZATION AND PREREQUISITE CHECK ---

Write-Host "=== Enhanced User Forensic Report Generator v4.0 ===" -ForegroundColor Magenta
Write-Host "Checking for required PowerShell modules..." -ForegroundColor Cyan

$requiredModules = @("Microsoft.Graph", "ImportExcel", "Microsoft.Graph.Users", "Microsoft.Graph.Identity.SignIns", "Microsoft.Graph.Identity.DirectoryManagement")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Error "Module '$module' is not installed. Please run 'Install-Module $module -Force' and try again."
        return
    }
}
Write-Host "All required modules are present." -ForegroundColor Green

# --- CONNECTION TO SERVICES ---

$graphScopes = @(
    "User.Read.All",
    "AuditLog.Read.All",
    "Directory.Read.All",
    "Reports.Read.All",
    "Policy.Read.All",
    "UserAuthenticationMethod.Read.All"
)

try {
    Write-Host "Connecting to Microsoft Graph (browser authentication will open)..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $graphScopes
    Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green

    $graphContext = Get-MgContext
    $loggedInUser = $graphContext.Account

    Write-Host "Connecting to Exchange Online using the existing session..." -ForegroundColor Cyan
    Connect-ExchangeOnline -UserPrincipalName $loggedInUser -ShowBanner:$false
    Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft 365 services. Please check your credentials and permissions. Error: $_"
    return
}

# --- ENHANCED LICENSE DETECTION ---

$tenantLicenseInfo = Get-TenantLicenseDetails

# --- ENHANCED USER SELECTION (MULTI-SELECT) ---

Write-Host "Fetching user list for selection..." -ForegroundColor Cyan
$allUsers = Get-MgUser -All -Property DisplayName, UserPrincipalName, Id, AccountEnabled, CreatedDateTime | 
    Select-Object DisplayName, UserPrincipalName, Id, AccountEnabled, 
    @{Name='CreatedDate'; Expression={$_.CreatedDateTime.ToString('yyyy-MM-dd')}} |
    Sort-Object DisplayName

if (-not $allUsers) {
    Write-Error "Could not retrieve any users from the tenant. Disconnecting."
    Disconnect-MgGraph
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

Write-Host "Select one or more users to investigate (use Ctrl+Click for multiple selection):" -ForegroundColor Yellow
$selectedUsers = $allUsers | Out-GridView -PassThru -Title "Select Users to Investigate (Ctrl+Click for multiple selection)"

if (-not $selectedUsers) {
    Write-Warning "No users were selected. Exiting script."
    Disconnect-MgGraph
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

$userCount = ($selectedUsers | Measure-Object).Count
Write-Host "Selected $userCount user(s) for investigation:" -ForegroundColor Green
$selectedUsers | ForEach-Object { Write-Host "  - $($_.DisplayName) ($($_.UserPrincipalName))" -ForegroundColor White }

# --- REPORT TYPE SELECTION ---

if (-not $ConsolidatedReport -and $userCount -gt 1) {
    $reportChoice = Read-Host "Generate [I]ndividual reports or [C]onsolidated report? (I/C, default: I)"
    $ConsolidatedReport = ($reportChoice -eq 'C' -or $reportChoice -eq 'c')
}

# --- OUTPUT PATH SELECTION ---

Add-Type -AssemblyName System.Windows.Forms

if ($ConsolidatedReport) {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Title = "Save Consolidated Forensic Report As"
    $saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    $saveFileDialog.FileName = "ConsolidatedForensicReport_$($userCount)Users_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    
    if ($saveFileDialog.ShowDialog() -ne 'OK') {
        Write-Warning "No output file was selected. Exiting script."
        Disconnect-MgGraph
        Disconnect-ExchangeOnline -Confirm:$false
        return
    }
    $OutputPath = $saveFileDialog.FileName
    Write-Host "Consolidated report will be saved to: $OutputPath" -ForegroundColor Green
}
else {
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select folder to save individual forensic reports"
    $folderDialog.ShowNewFolderButton = $true
    
    if ($folderDialog.ShowDialog() -ne 'OK') {
        Write-Warning "No output folder was selected. Exiting script."
        Disconnect-MgGraph
        Disconnect-ExchangeOnline -Confirm:$false
        return
    }
    $OutputFolder = $folderDialog.SelectedPath
    Write-Host "Individual reports will be saved to: $OutputFolder" -ForegroundColor Green
}

# --- DATA COLLECTION FUNCTION ---

function Get-UserForensicData {
    param(
        [PSCustomObject]$User,
        [int]$UserIndex,
        [int]$TotalUsers,
        [PSCustomObject]$TenantLicenseInfo
    )
    
    $UserPrincipalName = $User.UserPrincipalName
    $UserId = $User.Id
    $reportData = @{}
    $ErrorActionPreference = 'SilentlyContinue'
    
    try {
        Write-Host "[$UserIndex/$TotalUsers] Processing: $UserPrincipalName" -ForegroundColor Cyan

        # 1. Get General User Information with License Details
        Write-Host "  [1/11] Getting general user information and licenses..."
        $userInfo = Get-MgUser -UserId $UserPrincipalName -Property * | 
            Select-Object DisplayName, UserPrincipalName, Id, UserType, AccountEnabled, CreatedDateTime, 
            SignInActivity, Mail, JobTitle, Department, OfficeLocation, UsageLocation
        
        # Get user-specific licenses
        $licenseMapping = @{
            'AAD_BASIC'                    = @{ Name = 'Azure Active Directory Basic'; Tier = 'Basic'; LogRetention = 7 }
            'AAD_PREMIUM'                  = @{ Name = 'Azure Active Directory Premium P1'; Tier = 'P1'; LogRetention = 30 }
            'AAD_PREMIUM_P2'               = @{ Name = 'Azure Active Directory Premium P2'; Tier = 'P2'; LogRetention = 30 }
            'ENTERPRISEPACK'               = @{ Name = 'Microsoft 365 E3'; Tier = 'P1'; LogRetention = 30 }
            'ENTERPRISEPREMIUM'            = @{ Name = 'Microsoft 365 E5'; Tier = 'P2'; LogRetention = 30 }
            'ENTERPRISEPREMIUM_NOPSTNCONF' = @{ Name = 'Microsoft 365 E5 (without Audio Conferencing)'; Tier = 'P2'; LogRetention = 30 }
            'STANDARDPACK_FACULTY'         = @{ Name = 'Microsoft 365 A1 for Faculty'; Tier = 'Basic'; LogRetention = 7 }
            'ENTERPRISEPACK_FACULTY'       = @{ Name = 'Microsoft 365 A3 for Faculty'; Tier = 'P1'; LogRetention = 30 }
            'ENTERPRISEPREMIUM_FACULTY'    = @{ Name = 'Microsoft 365 A5 for Faculty'; Tier = 'P2'; LogRetention = 30 }
            'O365_BUSINESS_ESSENTIALS'     = @{ Name = 'Microsoft 365 Business Basic'; Tier = 'Basic'; LogRetention = 7 }
            'O365_BUSINESS_PREMIUM'        = @{ Name = 'Microsoft 365 Business Standard'; Tier = 'Basic'; LogRetention = 7 }
            'SPB'                          = @{ Name = 'Microsoft 365 Business Premium'; Tier = 'P1'; LogRetention = 30 }
            'EMS'                          = @{ Name = 'Enterprise Mobility + Security E3'; Tier = 'P1'; LogRetention = 30 }
            'EMSPREMIUM'                   = @{ Name = 'Enterprise Mobility + Security E5'; Tier = 'P2'; LogRetention = 30 }
        }
        $userLicenses = Get-UserLicenseDetails -UserId $UserId -LicenseMapping $licenseMapping
        
        $reportData.Add("User Info", $userInfo)
        $reportData.Add("User Licenses", $userLicenses)

        # 2. Get Assigned Admin Roles
        Write-Host "  [2/11] Getting assigned admin roles..."
        $adminRoles = Get-MgUserMemberOf -UserId $UserId -All | 
            Where-Object { $_.'@odata.type' -eq '#microsoft.graph.directoryRole' } | 
            Select-Object -ExpandProperty AdditionalProperties
        $reportData.Add("Admin Roles", $adminRoles)

        # 3. Get Group Memberships
        Write-Host "  [3/11] Getting user group memberships..."
        $groupMemberships = Get-MgUserMemberOf -UserId $UserId -All | 
            Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' } | 
            Select-Object -ExpandProperty AdditionalProperties
        $reportData.Add("Group Memberships", $groupMemberships)

        # 4. Enhanced Sign-in Logs with User-Specific License Check
        $userLogDays = $TenantLicenseInfo.MaxLogRetention
        if ($userLicenses) {
            $userMaxRetention = ($userLicenses | Measure-Object LogRetention -Maximum).Maximum
            if ($userMaxRetention -gt $userLogDays) {
                $userLogDays = $userMaxRetention
            }
        }

        Write-Host "  [4/11] Getting user sign-in logs (last $userLogDays days based on license analysis)..."
        try {
            $startDate = (Get-Date).AddDays(-$userLogDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            $filter = "userPrincipalName eq '$UserPrincipalName' and createdDateTime ge $startDate"
            $signInLogsRaw = Get-MgAuditLogSignIn -Filter $filter -All -ErrorAction Stop
            
            if ($null -ne $signInLogsRaw) {
                $signInLogsFormatted = $signInLogsRaw | Select-Object CreatedDateTime, UserPrincipalName, AppDisplayName, 
                    IpAddress, @{N = 'City'; E = { $_.Location.City } }, @{N = 'State'; E = { $_.Location.State } }, 
                    @{N = 'CountryOrRegion'; E = { $_.Location.CountryOrRegion } }, @{N = 'Status'; E = { $_.Status.ErrorCode } }, 
                    @{N = 'FailureReason'; E = { $_.Status.FailureReason } }, 
                    @{N = 'DeviceDetail'; E = { ($_.DeviceDetail | ConvertTo-Json -Depth 1 -Compress) } }
                $reportData.Add("Sign-in Logs", $signInLogsFormatted)
            } else {
                $reportData.Add("Sign-in Logs", [PSCustomObject]@{ 
                    Status = "No sign-in logs found for this user in the last $userLogDays days."
                })
            }
        }
        catch {
            $errorMessage = $_.ToString()
            if ($errorMessage -like "*Authentication_RequestFromPremiumTenantWithoutPremiumLicense*") {
                Write-Warning "    Could not retrieve sign-in logs. This tenant does not have the required license."
                $reportData.Add("Sign-in Logs", [PSCustomObject]@{ 
                    Status = "Data Not Available"; 
                    Reason = "Tenant does not have the required Entra ID P1 or P2 license to access this API." 
                })
            }
            else {
                Write-Error "    An unexpected error occurred while fetching sign-in logs: $errorMessage"
                $reportData.Add("Sign-in Logs", [PSCustomObject]@{ 
                    Status = "Error"; 
                    Reason = $errorMessage 
                })
            }
        }

        # 5. Get Directory Audit Logs
        Write-Host "  [5/11] Getting directory audit logs (last 90 days)..."
        $auditLogsRaw = Get-MgAuditLogDirectoryAudit -Filter "initiatedBy/user/id eq '$UserId'" -All
        $auditLogsFormatted = $auditLogsRaw | Select-Object ActivityDateTime, ActivityDisplayName, Category, Result, 
            @{N = 'TargetResource'; E = { ($_.TargetResources | Select-Object -First 1).DisplayName } }
        $reportData.Add("Directory Audits", $auditLogsFormatted)

        # --- EXCHANGE ONLINE DATA COLLECTION ---
        Write-Host "  [6/11] Checking for Exchange Online mailbox..."
        $hasMailbox = $false
        try {
            $mailboxCheck = Get-EXOMailbox -Identity $UserPrincipalName -ErrorAction Stop
            if ($mailboxCheck) { $hasMailbox = $true }
        }
        catch {
            # User has no mailbox
        }

        if ($hasMailbox) {
            Write-Host "    Mailbox found. Proceeding with Exchange data collection." -ForegroundColor Green
            
            # 7. Get Mailbox Statistics
            Write-Host "  [7/11] Getting mailbox statistics..."
            $mailboxStats = Get-EXOMailboxStatistics -Identity $UserPrincipalName | 
                Select-Object DisplayName, ItemCount, TotalItemSize, LastLogonTime, LastUserActionTime
            $reportData.Add("Mailbox Stats", $mailboxStats)

            # 8. Get Detailed Inbox Rules
            Write-Host "  [8/11] Getting and processing inbox rules..."
            $rawInboxRules = Get-InboxRule -Mailbox $UserPrincipalName
            $detailedInboxRules = foreach ($rule in $rawInboxRules) {
                [PSCustomObject]@{
                    RuleName                      = $rule.Name
                    Enabled                       = $rule.Enabled
                    Priority                      = $rule.Priority
                    Description                   = $rule.Description
                    DeleteMessage                 = $rule.DeleteMessage
                    StopProcessingRules           = $rule.StopProcessingRules
                    MoveToFolderName              = if ($rule.MoveToFolder) { 
                        (Get-EXOMailboxFolderStatistics -Identity $UserPrincipalName -FolderScope Inbox | 
                         Where-Object { $_.FolderId -eq $rule.MoveToFolder.ToString() }).FolderPath 
                    } else { $null }
                    ForwardTo                     = ($rule.ForwardTo | ForEach-Object { $_.DisplayName }) -join ", "
                    RedirectTo                    = ($rule.RedirectTo | ForEach-Object { $_.DisplayName }) -join ", "
                    ForwardingSmtpAddress         = $rule.ForwardingSmtpAddress
                    From                          = ($rule.From | ForEach-Object { $_.DisplayName }) -join ", "
                    SentTo                        = ($rule.SentTo | ForEach-Object { $_.DisplayName }) -join ", "
                    SubjectContainsWords          = ($rule.SubjectContainsWords) -join ", "
                    BodyContainsWords             = ($rule.BodyContainsWords) -join ", "
                    RecipientAddressContainsWords = ($rule.RecipientAddressContainsWords) -join ", "
                }
            }
            $reportData.Add("Inbox Rules (Detailed)", $detailedInboxRules)

            # 9. Get Mailbox Folder Permissions
            Write-Host "  [9/11] Getting mailbox folder permissions..."
            $folderPermissions = Get-EXOMailboxFolderPermission -Identity $UserPrincipalName
            $reportData.Add("Folder Permissions", $folderPermissions)

            # 10. Search Mailbox Audit Log
            Write-Host "  [10/11] Searching mailbox audit log (last 90 days)..."
            $mailboxAuditLog = Search-MailboxAuditLog -Identity $UserPrincipalName -ShowDetails -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date)
            $reportData.Add("Mailbox Audit Log", $mailboxAuditLog)
        }
        else {
            Write-Warning "    No Exchange Online mailbox found for '${UserPrincipalName}'. Skipping Exchange-related data."
            $reportData.Add("Exchange Data", [PSCustomObject]@{ 
                Status = "Data Not Available"; 
                Reason = "No Exchange Online mailbox was found for this user." 
            })
        }

        # 11. Enhanced MFA Analysis
        Write-Host "  [11/11] Analyzing MFA status..."
        $mfaSummary = [PSCustomObject]@{
            OverallStatus           = "Unknown"
            Summary                 = ""
            PerUserMfaStatus        = "Not configured"
            SecurityDefaultsEnabled = "Unknown"
            CAPoliciesRequiringMfa  = "None"
            UserLicenseTier        = if ($userLicenses) { ($userLicenses | Measure-Object Tier -Maximum).Maximum } else { "None" }
        }
        
        $authMethods = Get-MgUserAuthenticationMethod -UserId $UserId
        $mfaMethods = $authMethods | Where-Object { $_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' }
        if ($mfaMethods) { 
            $mfaSummary.PerUserMfaStatus = "Enabled. Methods: " + (($mfaMethods).'@odata.type' -replace '#microsoft.graph.' -join ', ') 
        }
        
        $securityDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy
        $mfaSummary.SecurityDefaultsEnabled = if ($securityDefaults.IsEnabled) { "Enabled" } else { "Disabled" }
        
        $allCaPolicies = Get-MgIdentityConditionalAccessPolicy -All
        $applicableMfaPolicies = @()
        foreach ($policy in $allCaPolicies) {
            if ($policy.State -eq "enabled" -and $policy.GrantControls.BuiltInControls -contains "mfa") {
                $isIncluded = ($policy.Conditions.Users.IncludeUsers -contains "All") -or ($policy.Conditions.Users.IncludeUsers -contains $UserId)
                $isExcluded = $policy.Conditions.Users.ExcludeUsers -contains $UserId
                if ($isIncluded -and -not $isExcluded) { $applicableMfaPolicies += $policy.DisplayName }
            }
        }
        if ($applicableMfaPolicies.Count -gt 0) { $mfaSummary.CAPoliciesRequiringMfa = $applicableMfaPolicies -join "; " }

        # Enhanced MFA status determination
        if ($mfaSummary.SecurityDefaultsEnabled -eq "Enabled") {
            $mfaSummary.OverallStatus = "Protected (Security Defaults)"
            $mfaSummary.Summary = "MFA is enforced for this user via tenant-wide Security Defaults."
        }
        elseif ($applicableMfaPolicies.Count -gt 0) {
            $mfaSummary.OverallStatus = "Protected (Conditional Access)"
            $mfaSummary.Summary = "MFA is enforced by one or more Conditional Access policies."
        }
        elseif ($mfaMethods) {
            $mfaSummary.OverallStatus = "Potentially Protected (Per-User MFA)"
            $mfaSummary.Summary = "User has MFA methods registered, but enforcement via policy was not detected."
        }
        else {
            $mfaSummary.OverallStatus = "NOT PROTECTED"
            $mfaSummary.Summary = "No MFA enforcement method detected for this user."
        }
        
        $reportData.Add("MFA Status", $mfaSummary)
        $reportData.Add("All MFA CA Policies", ($allCaPolicies | Where-Object { $_.GrantControls.BuiltInControls -contains "mfa" } | Select-Object DisplayName, State, Id))

        Write-Host "  Data collection complete for $UserPrincipalName" -ForegroundColor Green
        return $reportData
    }
    catch {
        Write-Error "  An error occurred during data collection for ${UserPrincipalName}. Error: $_"
        return $null
    }
    finally {
        $ErrorActionPreference = 'Continue'
    }
}

# --- EXCEL REPORT GENERATION FUNCTION ---

function Export-ForensicReport {
    param(
        [hashtable]$ReportData,
        [string]$OutputPath,
        [string]$UserName = ""
    )
    
    if ($ReportData.Count -gt 0) {
        Write-Host "Generating Excel report at: $OutputPath" -ForegroundColor Cyan
        try {
            $excelParams = @{
                Path          = $OutputPath
                InputObject   = $null
                AutoFilter    = $true
                AutoSize      = $true
                TableName     = ''
                WorksheetName = ''
            }

            $isFirstSheet = $true

            foreach ($sheetName in $ReportData.Keys) {
                $data = $ReportData[$sheetName]

                if ($null -eq $data) {
                    Write-Warning "No data found for '${sheetName}'. Skipping this worksheet."
                    continue
                }

                $excelParams.WorksheetName = if ($UserName) { "$UserName - $sheetName" } else { $sheetName }
                $tableName = ($excelParams.WorksheetName -replace '[^a-zA-Z0-9]', '')
                if ($tableName.Length -eq 0) { $tableName = "Data" }
                $excelParams.TableName = $tableName.Substring(0, [Math]::Min(255, $tableName.Length))
                $excelParams.InputObject = $data

                if ($isFirstSheet) {
                    Export-Excel @excelParams
                    $isFirstSheet = $false
                }
                else {
                    Export-Excel @excelParams -Append
                }
            }

            Write-Host "Excel report successfully generated!" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Error "Failed to generate the Excel report. Error: $_"
            return $false
        }
    }
    else {
        Write-Warning "No data was collected. The Excel report will not be generated."
        return $false
    }
}

# --- MAIN PROCESSING LOOP ---

Write-Host "`n=== Starting Data Collection ===" -ForegroundColor Magenta

if ($ConsolidatedReport) {
    # Consolidated Report Generation
    Write-Host "Generating consolidated report for $userCount users..." -ForegroundColor Cyan
    $consolidatedData = @{}
    
    # Add tenant license information to consolidated report
    try {
        if ($tenantLicenseInfo.Details -and $tenantLicenseInfo.Details.Count -gt 0) {
            $consolidatedData.Add("Tenant License Summary", $tenantLicenseInfo.Details)
        }
    }
    catch {
        Write-Warning "Could not add tenant license summary to report: $_"
    }
    
    $userIndex = 1
    foreach ($user in $selectedUsers) {
        $userData = Get-UserForensicData -User $user -UserIndex $userIndex -TotalUsers $userCount -TenantLicenseInfo $tenantLicenseInfo
        
        if ($userData) {
            foreach ($dataType in $userData.Keys) {
                $sheetName = "$($user.DisplayName) - $dataType"
                try {
                    $consolidatedData.Add($sheetName, $userData[$dataType])
                }
                catch {
                    Write-Warning "Could not add data for ${sheetName}: $_"
                }
            }
        }
        $userIndex++
    }
    
    $success = Export-ForensicReport -ReportData $consolidatedData -OutputPath $OutputPath
    if ($success) {
        Write-Host "`nConsolidated report generated successfully!" -ForegroundColor Green
        Write-Host "Report saved to: $OutputPath" -ForegroundColor Yellow
    }
}
else {
    # Individual Report Generation
    Write-Host "Generating individual reports for $userCount users..." -ForegroundColor Cyan
    $successCount = 0
    $userIndex = 1
    
    foreach ($user in $selectedUsers) {
        $userData = Get-UserForensicData -User $user -UserIndex $userIndex -TotalUsers $userCount -TenantLicenseInfo $tenantLicenseInfo
        
        if ($userData) {
            # Add tenant license info to each individual report
            try {
                if ($tenantLicenseInfo.Details -and $tenantLicenseInfo.Details.Count -gt 0) {
                    $userData.Add("Tenant License Summary", $tenantLicenseInfo.Details)
                }
            }
            catch {
                Write-Warning "Could not add tenant license summary to individual report: $_"
            }
            
            $safeFileName = ($user.UserPrincipalName -replace '[\\/:"*?<>|]', '_')
            $individualPath = Join-Path $OutputFolder "ForensicReport_$($safeFileName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            $success = Export-ForensicReport -ReportData $userData -OutputPath $individualPath
            if ($success) {
                $successCount++
                Write-Host "Individual report generated for $($user.DisplayName)" -ForegroundColor Green
            }
        }
        $userIndex++
    }
    
    Write-Host "`nIndividual report generation complete!" -ForegroundColor Green
    Write-Host "Successfully generated $successCount out of $userCount reports" -ForegroundColor Yellow
    Write-Host "Reports saved to: $OutputFolder" -ForegroundColor Yellow
}

# --- DISCONNECT FROM SERVICES ---

Write-Host "`n=== Cleanup ===" -ForegroundColor Magenta
Write-Host "Disconnecting from all services..." -ForegroundColor Cyan
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Script finished successfully!" -ForegroundColor Green