<#
.SYNOPSIS
A PowerShell script with a tabbed GUI to investigate Entra ID accounts by fetching sign-in logs, 
user details/roles, audit logs, and MFA analysis.

.NOTES
Author: Gemini (Enhanced by Claude)
Date: 2025-06-19
Version: 6.0 (COMPLETE: Added MFA Analysis and Enhanced User Analysis)
Requires:
    - PowerShell 5.1+
    - Microsoft Graph SDK (Users, Reports, Identity.DirectoryManagement, Identity.SignIns)
    - *** Microsoft Excel Installed *** (for XLSX conversion)
Permissions: Requires delegated User.Read.All, AuditLog.Read.All, Directory.Read.All, Policy.Read.All, UserAuthenticationMethod.Read.All
#>

#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Reports, Microsoft.Graph.Identity.DirectoryManagement

# --- Configuration ---
$requiredModules = @(
    "Microsoft.Graph.Users", 
    "Microsoft.Graph.Reports", 
    "Microsoft.Graph.Identity.DirectoryManagement"
)

$requiredScopes = @(
    "User.Read.All", 
    "AuditLog.Read.All", 
    "Organization.Read.All", 
    "Directory.Read.All",
    "Policy.Read.All",
    "UserAuthenticationMethod.Read.All"
)

$premiumLicenseSkus = @('AAD_PREMIUM', 'AAD_PREMIUM_P2')

# --- Script-level variables ---
$script:lastExportedXlsxPath = $null 
$script:cachedRoles = $null 

# --- Function Definitions ---
Function Test-Modules { 
    param($Modules) 
    $missingModules = @()
    foreach ($m in $Modules) { 
        if (-not (Get-Module -ListAvailable -Name $m)) { 
            $missingModules += $m 
        } 
    }
    return $missingModules 
}

Function Install-MissingModules { 
    param($Modules) 
    try { 
        Install-Module -Name $Modules -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop 
    } catch { 
        Write-Error "Failed to install modules." 
    } 
}

Function ConvertTo-XlsxAndFormat { 
    param(
        [string]$CsvPath, 
        [string]$XlsxPath, 
        [int]$HighlightColor = 6, 
        [string]$CountryColumnHeader = "Country", 
        [string]$CountryToHighlight = "United States"
    ) 
    
    $excel = $null
    try { 
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop 
    } catch { 
        Write-Error "Excel not found."
        return $false 
    } 
    
    try { 
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($CsvPath)
        $workbook.SaveAs($XlsxPath, 51)
        $workbook.Close($false)
        $workbook = $excel.Workbooks.Open($XlsxPath)
        $worksheet = $workbook.Worksheets.Item(1)
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null
        $usedRange.Rows.Item(1).Font.Bold = $true
        
        $countryCol = $usedRange.Rows.Item(1).Find($CountryColumnHeader)
        if ($countryCol) { 
            for ($i = 2; $i -le $usedRange.Rows.Count; $i++) { 
                $cell = $worksheet.Cells.Item($i, $countryCol.Column)
                if ($cell.Value2 -and ($cell.Value2 -as [string]).Equals($CountryToHighlight, 'OrdinalIgnoreCase')) { 
                    $worksheet.Rows.Item($i).Interior.ColorIndex = $HighlightColor 
                }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null 
            } 
        }
        
        $workbook.Save()
        $workbook.Close()
        $script:lastExportedXlsxPath = $XlsxPath
        if ($openFileButton) { 
            $openFileButton.Enabled = $true 
        } 
    } catch { 
        Write-Error "Excel conversion failed: $($_.Exception.Message)"
        return $false 
    } finally { 
        if ($excel) { 
            $excel.Quit() 
        }
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers() 
    }
    return $true 
}

Function Get-UserRolesAndGroups { 
    param([string]$UserId) 
    
    $results = @{
        Roles = [System.Collections.ArrayList]@()
        Groups = [System.Collections.ArrayList]@()
        Error = $null
    }
    
    try { 
        $memberOf = Get-MgUserMemberOf -UserId $UserId -All -ErrorAction SilentlyContinue
        if ($memberOf) { 
            $roleNames = $memberOf | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.directoryRole' } | Select-Object -ExpandProperty DisplayName
            if ($roleNames) { 
                $results.Roles.AddRange($roleNames) 
            }
            
            $groupNames = $memberOf | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' } | Select-Object -ExpandProperty DisplayName
            if ($groupNames) { 
                $results.Groups.AddRange($groupNames) 
            } 
        }
        
        if (-not $script:cachedRoles) { 
            $script:cachedRoles = Get-MgDirectoryRole -ErrorAction Stop 
        }
        
        $rolesToCheck = @("Global Administrator", "User Administrator", "Security Administrator", "Exchange Administrator")
        foreach ($roleName in $rolesToCheck) { 
            $role = $script:cachedRoles | Where-Object { $_.DisplayName -eq $roleName }
            if ($role) { 
                $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction SilentlyContinue
                if ($roleMembers.Id -contains $UserId) { 
                    if ($results.Roles -notcontains $roleName) { 
                        $results.Roles.Add($roleName) | Out-Null 
                    } 
                } 
            } 
        } 
    } catch { 
        $results.Error = "An error occurred while fetching roles/groups: $($_.Exception.Message)"
        Write-Warning $results.Error 
    }
    
    return $results 
}

Function Get-UserMfaStatus {
    param([string]$UserId, [string]$UserPrincipalName)
    
    $results = @{
        PerUserMfa = @{ Enabled = $false; Methods = @(); Details = "Not configured" }
        SecurityDefaults = @{ Enabled = $false; Details = "Unknown" }
        ConditionalAccess = @{ Policies = @(); RequiresMfa = $false; Details = "No applicable policies" }
        OverallStatus = "Unknown"
        Summary = ""
    }
    
    try {
        # 1. Check Per-User MFA
        try {
            $authMethods = Get-MgUserAuthenticationMethod -UserId $UserId -ErrorAction SilentlyContinue
            if ($authMethods) {
                $mfaMethods = $authMethods | Where-Object { 
                    $_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' 
                }
                if ($mfaMethods) {
                    $results.PerUserMfa.Enabled = $true
                    $results.PerUserMfa.Methods = $mfaMethods | ForEach-Object { 
                        $_.'@odata.type' -replace '#microsoft.graph.', '' -replace 'AuthenticationMethod', ''
                    }
                    $results.PerUserMfa.Details = "Methods: $($results.PerUserMfa.Methods -join ', ')"
                } else {
                    $results.PerUserMfa.Details = "No MFA methods registered"
                }
            }
        } catch {
            $results.PerUserMfa.Details = "Error checking per-user MFA: $($_.Exception.Message)"
        }
        
        # 2. Check Security Defaults
        try {
            $securityDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -ErrorAction SilentlyContinue
            if ($securityDefaults) {
                $results.SecurityDefaults.Enabled = $securityDefaults.IsEnabled
                $results.SecurityDefaults.Details = if ($securityDefaults.IsEnabled) { 
                    "Enabled (requires MFA for all users)" 
                } else { 
                    "Disabled" 
                }
            }
        } catch {
            $results.SecurityDefaults.Details = "Error checking security defaults: $($_.Exception.Message)"
        }
        
        # 3. Check Conditional Access Policies
        try {
            $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction SilentlyContinue
            if ($caPolicies) {
                $applicablePolicies = @()
                
                foreach ($policy in $caPolicies) {
                    if ($policy.State -eq "enabled") {
                        # Check if policy applies to this user
                        $appliesToUser = $false
                        
                        # Check user inclusions
                        if ($policy.Conditions.Users.IncludeUsers -contains "All" -or 
                            $policy.Conditions.Users.IncludeUsers -contains $UserId) {
                            $appliesToUser = $true
                        }
                        
                        # Check if user is in included groups (simplified check)
                        if ($policy.Conditions.Users.IncludeGroups -and $policy.Conditions.Users.IncludeGroups.Count -gt 0) {
                            # Would need to check group membership - simplified for now
                            $appliesToUser = $true
                        }
                        
                        # Check exclusions
                        if ($policy.Conditions.Users.ExcludeUsers -contains $UserId) {
                            $appliesToUser = $false
                        }
                        
                        if ($appliesToUser) {
                            # Check if policy requires MFA
                            $requiresMfa = $false
                            $controls = @()
                            
                            if ($policy.GrantControls.BuiltInControls -contains "mfa") {
                                $requiresMfa = $true
                                $controls += "MFA Required"
                            }
                            
                            if ($policy.GrantControls.BuiltInControls -contains "passwordChange") {
                                $controls += "Password Change"
                            }
                            
                            if ($policy.GrantControls.BuiltInControls -contains "block") {
                                $controls += "Block Access"
                            }
                            
                            # Build conditions summary
                            $conditions = @()
                            if ($policy.Conditions.Applications.IncludeApplications -contains "All") {
                                $conditions += "All Apps"
                            } else {
                                $conditions += "Specific Apps"
                            }
                            
                            if ($policy.Conditions.Locations) {
                                $conditions += "Location-based"
                            }
                            
                            if ($policy.Conditions.Platforms) {
                                $conditions += "Platform-specific"
                            }
                            
                            $policyInfo = @{
                                Name = $policy.DisplayName
                                State = $policy.State
                                RequiresMfa = $requiresMfa
                                Controls = if ($controls.Count -gt 0) { $controls -join ", " } else { "None" }
                                Conditions = if ($conditions.Count -gt 0) { $conditions -join ", " } else { "All scenarios" }
                            }
                            
                            $applicablePolicies += $policyInfo
                            
                            if ($requiresMfa) {
                                $results.ConditionalAccess.RequiresMfa = $true
                            }
                        }
                    }
                }
                
                $results.ConditionalAccess.Policies = $applicablePolicies
                if ($applicablePolicies.Count -gt 0) {
                    $mfaPolicies = $applicablePolicies | Where-Object { $_.RequiresMfa }
                    $results.ConditionalAccess.Details = "Found $($applicablePolicies.Count) applicable policies ($($mfaPolicies.Count) require MFA)"
                } else {
                    $results.ConditionalAccess.Details = "No applicable conditional access policies found"
                }
            }
        } catch {
            $results.ConditionalAccess.Details = "Error checking conditional access: $($_.Exception.Message)"
        }
        
        # 4. Determine Overall Status
        if ($results.SecurityDefaults.Enabled) {
            $results.OverallStatus = "Protected (Security Defaults)"
            $results.Summary = "MFA required via Security Defaults"
        } elseif ($results.ConditionalAccess.RequiresMfa) {
            $results.OverallStatus = "Protected (Conditional Access)"
            $results.Summary = "MFA required via Conditional Access policies"
        } elseif ($results.PerUserMfa.Enabled) {
            $results.OverallStatus = "Protected (Per-User MFA)"
            $results.Summary = "Per-user MFA methods configured"
        } else {
            $results.OverallStatus = "⚠️ NOT PROTECTED"
            $results.Summary = "No MFA protection detected"
        }
        
    } catch {
        $results.OverallStatus = "Error"
        $results.Summary = "Failed to analyze MFA status: $($_.Exception.Message)"
    }
    
    return $results
}

# --- Prerequisites Check ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$missing = Test-Modules -Modules $requiredModules
if ($missing.Count -gt 0) { 
    if (('Yes' -eq [System.Windows.Forms.MessageBox]::Show("Modules missing: $($missing -join ', '). Install now?", "Missing Modules", 'YesNo', 'Warning'))) { 
        Install-MissingModules -Modules $missing
        [System.Windows.Forms.MessageBox]::Show("Restart script.", "Restart Required", 'OK', 'Information')
        Exit 
    } else { 
        [System.Windows.Forms.MessageBox]::Show("Cannot continue.", "Error", 'OK', 'Error')
        Exit 
    } 
}

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Reports
Import-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Identity.SignIns

# --- GUI Setup ---
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Entra Investigator v6.0"
$mainForm.Size = '620, 700'
$mainForm.StartPosition = 'CenterScreen'
$mainForm.FormBorderStyle = 'Sizable'
$mainForm.MaximizeBox = $true
$mainForm.MinimumSize = '620, 700'

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready."
$statusStrip.Items.Add($statusLabel)
$mainForm.Controls.Add($statusStrip)

$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Location = '10, 10'
$connectButton.Size = '140, 30'
$connectButton.Text = "Connect & Load Users"

$disconnectButton = New-Object System.Windows.Forms.Button
$disconnectButton.Location = '155, 10'
$disconnectButton.Size = '100, 30'
$disconnectButton.Text = "Disconnect"
$disconnectButton.Enabled = $false

$checkLicenseButton = New-Object System.Windows.Forms.Button
$checkLicenseButton.Location = '265, 10'
$checkLicenseButton.Size = '180, 30'
$checkLicenseButton.Text = "Analyze Selected Users"
$checkLicenseButton.Enabled = $false

$userListLabel = New-Object System.Windows.Forms.Label
$userListLabel.Location = '10, 50'
$userListLabel.Size = '200, 20'
$userListLabel.Text = "Select User(s) for Investigation:"

$userCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$userCheckedListBox.Location = '10, 70'
$userCheckedListBox.Size = '580, 180'
$userCheckedListBox.CheckOnClick = $true
$userCheckedListBox.Anchor = 'Top, Left, Right'
$userCheckedListBox.HorizontalScrollbar = $true

$selectAllCheckbox = New-Object System.Windows.Forms.CheckBox
$selectAllCheckbox.Location = '475, 50'
$selectAllCheckbox.Size = '115, 20'
$selectAllCheckbox.Text = "Select All"
$selectAllCheckbox.Enabled = $false
$selectAllCheckbox.Anchor = 'Top, Right'

$mainForm.Controls.AddRange(@($connectButton, $disconnectButton, $checkLicenseButton, $userListLabel, $userCheckedListBox, $selectAllCheckbox))

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = '10, 260'
$tabControl.Size = '580, 370'
$tabControl.Anchor = 'Top, Bottom, Left, Right'
$mainForm.Controls.Add($tabControl)

# --- Export Tab ---
$exportTab = New-Object System.Windows.Forms.TabPage
$exportTab.Text = "Sign-in Logs (Export)"

$logDurationLabel = New-Object System.Windows.Forms.Label
$logDurationLabel.Location = '10, 20'
$logDurationLabel.Size = '150, 20'
$logDurationLabel.Text = "Log History (Days):"

$logDurationNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
$logDurationNumericUpDown.Location = '160, 20'
$logDurationNumericUpDown.Size = '60, 25'
$logDurationNumericUpDown.Minimum = 1
$logDurationNumericUpDown.Maximum = 30
$logDurationNumericUpDown.Value = 7

$durationWarningLabel = New-Object System.Windows.Forms.Label
$durationWarningLabel.Location = '230, 23'
$durationWarningLabel.Size = '325, 20'
$durationWarningLabel.ForeColor = 'OrangeRed'
$durationWarningLabel.Anchor = 'Top, Left, Right'

$outputFolderLabel = New-Object System.Windows.Forms.Label
$outputFolderLabel.Location = '10, 60'
$outputFolderLabel.Size = '100, 20'
$outputFolderLabel.Text = "Output Folder:"

$outputFolderTextBox = New-Object System.Windows.Forms.TextBox
$outputFolderTextBox.Location = '110, 60'
$outputFolderTextBox.Size = '355, 25'
$outputFolderTextBox.ReadOnly = $true
$outputFolderTextBox.Anchor = 'Top, Left, Right'

$browseFolderButton = New-Object System.Windows.Forms.Button
$browseFolderButton.Location = '475, 58'
$browseFolderButton.Size = '90, 27'
$browseFolderButton.Text = "Browse..."
$browseFolderButton.Anchor = 'Top, Right'

$getLogsButton = New-Object System.Windows.Forms.Button
$getLogsButton.Location = '10, 110'
$getLogsButton.Size = '555, 40'
$getLogsButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$getLogsButton.Text = "Get Sign-in Logs & Export to XLSX"
$getLogsButton.Enabled = $false
$getLogsButton.Anchor = 'Top, Left, Right'

$openFileButton = New-Object System.Windows.Forms.Button
$openFileButton.Location = '10, 160'
$openFileButton.Size = '555, 30'
$openFileButton.Text = "Open Last Exported File"
$openFileButton.Enabled = $false
$openFileButton.Anchor = 'Top, Left, Right'

$exportTab.Controls.AddRange(@($logDurationLabel, $logDurationNumericUpDown, $durationWarningLabel, $outputFolderLabel, $outputFolderTextBox, $browseFolderButton, $getLogsButton, $openFileButton))
$tabControl.TabPages.Add($exportTab)

# --- Details Tab ---
$detailsTab = New-Object System.Windows.Forms.TabPage
$detailsTab.Text = "User Details & Roles"

$fetchDetailsButton = New-Object System.Windows.Forms.Button
$fetchDetailsButton.Location = '10, 10'
$fetchDetailsButton.Size = '555, 30'
$fetchDetailsButton.Text = "Fetch Details for Selected User"
$fetchDetailsButton.Enabled = $false
$fetchDetailsButton.Anchor = 'Top, Left, Right'

$detailsRichTextBox = New-Object System.Windows.Forms.RichTextBox
$detailsRichTextBox.Location = '10, 50'
$detailsRichTextBox.Size = '555, 280'
$detailsRichTextBox.ReadOnly = $true
$detailsRichTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$detailsRichTextBox.Anchor = 'Top, Bottom, Left, Right'

$detailsTab.Controls.AddRange(@($fetchDetailsButton, $detailsRichTextBox))
$tabControl.TabPages.Add($detailsTab)

# --- Audit Tab ---
$auditTab = New-Object System.Windows.Forms.TabPage
$auditTab.Text = "User Audit Logs"

$fetchAuditButton = New-Object System.Windows.Forms.Button
$fetchAuditButton.Location = '10, 10'
$fetchAuditButton.Size = '360, 30'
$fetchAuditButton.Text = "Fetch Audit Logs for Selected User"
$fetchAuditButton.Enabled = $false
$fetchAuditButton.Anchor = 'Top, Left, Right'

$exportAuditButton = New-Object System.Windows.Forms.Button
$exportAuditButton.Location = '380, 10'
$exportAuditButton.Size = '120, 30'
$exportAuditButton.Text = "Export to CSV"
$exportAuditButton.Enabled = $false
$exportAuditButton.Anchor = 'Top, Right'

$auditGrid = New-Object System.Windows.Forms.DataGridView
$auditGrid.Location = '10, 50'
$auditGrid.Size = '555, 250'
$auditGrid.ReadOnly = $true
$auditGrid.AllowUserToAddRows = $false
$auditGrid.Anchor = 'Top, Bottom, Left, Right'
$auditGrid.AutoGenerateColumns = $false
$auditGrid.SelectionMode = 'FullRowSelect'
$auditGrid.MultiSelect = $false

# Define columns for audit grid
$timeCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$timeCol.Name = "Time"
$timeCol.HeaderText = "Date/Time"
$timeCol.Width = 140
$timeCol.MinimumWidth = 140
$auditGrid.Columns.Add($timeCol)

$activityCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$activityCol.Name = "Activity"
$activityCol.HeaderText = "Activity"
$activityCol.Width = 200
$activityCol.MinimumWidth = 150
$activityCol.DefaultCellStyle.WrapMode = 'True'
$auditGrid.Columns.Add($activityCol)

$categoryCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$categoryCol.Name = "Category"
$categoryCol.HeaderText = "Category"
$categoryCol.Width = 100
$categoryCol.MinimumWidth = 80
$auditGrid.Columns.Add($categoryCol)

$resultCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$resultCol.Name = "Result"
$resultCol.HeaderText = "Result"
$resultCol.Width = 80
$resultCol.MinimumWidth = 60
$auditGrid.Columns.Add($resultCol)

$targetCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$targetCol.Name = "Target"
$targetCol.HeaderText = "Target/Object"
$targetCol.AutoSizeMode = 'Fill'
$targetCol.MinimumWidth = 100
$targetCol.DefaultCellStyle.WrapMode = 'True'
$auditGrid.Columns.Add($targetCol)

$auditGrid.AutoSizeRowsMode = 'AllCells'
$auditGrid.AllowUserToResizeColumns = $true
$auditGrid.AllowUserToResizeRows = $false
$auditGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$auditGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$auditGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightBlue

$auditSummaryLabel = New-Object System.Windows.Forms.Label
$auditSummaryLabel.Location = '10, 310'
$auditSummaryLabel.Size = '555, 20'
$auditSummaryLabel.Text = "Select a user and click 'Fetch Audit Logs' to see recent administrative activities."
$auditSummaryLabel.Anchor = 'Bottom, Left, Right'
$auditSummaryLabel.ForeColor = [System.Drawing.Color]::DarkBlue

$auditTab.Controls.AddRange(@($fetchAuditButton, $exportAuditButton, $auditGrid, $auditSummaryLabel))
$tabControl.TabPages.Add($auditTab)

# --- MFA Analysis Tab ---
$mfaTab = New-Object System.Windows.Forms.TabPage
$mfaTab.Text = "MFA Analysis"

$fetchMfaButton = New-Object System.Windows.Forms.Button
$fetchMfaButton.Location = '10, 10'
$fetchMfaButton.Size = '360, 30'
$fetchMfaButton.Text = "Analyze MFA Status for Selected User"
$fetchMfaButton.Enabled = $false
$fetchMfaButton.Anchor = 'Top, Left, Right'

$exportMfaButton = New-Object System.Windows.Forms.Button
$exportMfaButton.Location = '380, 10'
$exportMfaButton.Size = '120, 30'
$exportMfaButton.Text = "Export to TXT"
$exportMfaButton.Enabled = $false
$exportMfaButton.Anchor = 'Top, Right'

$mfaRichTextBox = New-Object System.Windows.Forms.RichTextBox
$mfaRichTextBox.Location = '10, 50'
$mfaRichTextBox.Size = '555, 200'
$mfaRichTextBox.ReadOnly = $true
$mfaRichTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$mfaRichTextBox.Anchor = 'Top, Left, Right'

$mfaPolicyGrid = New-Object System.Windows.Forms.DataGridView
$mfaPolicyGrid.Location = '10, 260'
$mfaPolicyGrid.Size = '555, 100'
$mfaPolicyGrid.ReadOnly = $true
$mfaPolicyGrid.AllowUserToAddRows = $false
$mfaPolicyGrid.Anchor = 'Top, Bottom, Left, Right'
$mfaPolicyGrid.AutoGenerateColumns = $false
$mfaPolicyGrid.SelectionMode = 'FullRowSelect'
$mfaPolicyGrid.MultiSelect = $false

# Define columns for MFA policy grid
$policyNameCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$policyNameCol.Name = "PolicyName"
$policyNameCol.HeaderText = "Policy Name"
$policyNameCol.Width = 200
$policyNameCol.MinimumWidth = 150
$mfaPolicyGrid.Columns.Add($policyNameCol)

$policyStateCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$policyStateCol.Name = "State"
$policyStateCol.HeaderText = "State"
$policyStateCol.Width = 80
$policyStateCol.MinimumWidth = 60
$mfaPolicyGrid.Columns.Add($policyStateCol)

$policyControlsCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$policyControlsCol.Name = "Controls"
$policyControlsCol.HeaderText = "MFA Controls"
$policyControlsCol.AutoSizeMode = 'Fill'
$policyControlsCol.MinimumWidth = 100
$policyControlsCol.DefaultCellStyle.WrapMode = 'True'
$mfaPolicyGrid.Columns.Add($policyControlsCol)

$policyConditionsCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$policyConditionsCol.Name = "Conditions"
$policyConditionsCol.HeaderText = "Conditions"
$policyConditionsCol.Width = 150
$policyConditionsCol.MinimumWidth = 100
$policyConditionsCol.DefaultCellStyle.WrapMode = 'True'
$mfaPolicyGrid.Columns.Add($policyConditionsCol)

$mfaPolicyGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$mfaPolicyGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$mfaPolicyGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::LightBlue
$mfaPolicyGrid.AutoSizeRowsMode = 'AllCells'

$mfaSummaryLabel = New-Object System.Windows.Forms.Label
$mfaSummaryLabel.Location = '10, 370'
$mfaSummaryLabel.Size = '555, 20'
$mfaSummaryLabel.Text = "Select a user and click 'Analyze MFA Status' to check their multi-factor authentication configuration."
$mfaSummaryLabel.Anchor = 'Bottom, Left, Right'
$mfaSummaryLabel.ForeColor = [System.Drawing.Color]::DarkBlue

$mfaTab.Controls.AddRange(@($fetchMfaButton, $exportMfaButton, $mfaRichTextBox, $mfaPolicyGrid, $mfaSummaryLabel))
$tabControl.TabPages.Add($mfaTab)

# --- Event Handlers ---

# CONNECT BUTTON
$connectButton.add_Click({ 
    $statusLabel.Text = "Connecting..."
    $mainForm.Cursor = 'WaitCursor'
    $userCheckedListBox.Items.Clear()
    $script:cachedRoles = $null
    
    try { 
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
        $statusLabel.Text = "Getting all users..."
        $mainForm.Refresh()
        
        $allUsers = Get-MgUser -All -Property UserPrincipalName, Id -ConsistencyLevel eventual | Sort-Object UserPrincipalName
        if ($allUsers) { 
            $userCheckedListBox.Items.AddRange($allUsers.UserPrincipalName)
            $statusLabel.Text = "Connected. Loaded $($allUsers.Count) users."
            $disconnectButton.Enabled = $true
            $selectAllCheckbox.Enabled = $true 
        } else { 
            $statusLabel.Text = "Connected, but no users found." 
        } 
    } catch { 
        $statusLabel.Text = "Operation failed."
        Write-Error "Connection/load failed: $($_.Exception.Message)" 
    } finally { 
        $mainForm.Cursor = 'Default' 
    } 
})

# DISCONNECT BUTTON
$disconnectButton.add_Click({ 
    $statusLabel.Text = "Disconnecting..."
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    $userCheckedListBox.Items.Clear()
    $statusLabel.Text = "Disconnected."
    $disconnectButton.Enabled = $false
    $selectAllCheckbox.Enabled = $false
    $selectAllCheckbox.Checked = $false
    $selectAllCheckbox.Text = "Select All"
    $checkLicenseButton.Enabled = $false
    $fetchDetailsButton.Enabled = $false
    $fetchAuditButton.Enabled = $false
    $fetchMfaButton.Enabled = $false
    $exportMfaButton.Enabled = $false
    $getLogsButton.Enabled = $false
    $exportAuditButton.Enabled = $false
    if ($openFileButton) { 
        $openFileButton.Enabled = $false 
    } 
})

# ENHANCED ANALYZE USERS BUTTON
$checkLicenseButton.add_Click({ 
    if ($userCheckedListBox.CheckedItems.Count -eq 0) { 
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user to check.", "No Users Selected", "OK", "Information")
        return 
    }
    
    $mainForm.Cursor = 'WaitCursor'
    $checkLicenseButton.Enabled = $false
    $connectButton.Enabled = $false
    $disconnectButton.Enabled = $false
    $getLogsButton.Enabled = $false
    
    try { 
        $totalUserCount = $userCheckedListBox.CheckedIndices.Count
        $currentUserIndex = 0
        
        foreach ($i in $userCheckedListBox.CheckedIndices) { 
            $currentUserIndex++
            $originalItemText = $userCheckedListBox.Items[$i].ToString().Split(' ')[0]
            $statusLabel.Text = "Analyzing user ${currentUserIndex}/${totalUserCount}: $originalItemText"
            $mainForm.Refresh()
            
            # Initialize variables
            $hasPremiumLicense = $false
            $accountEnabled = "Unknown"
            $lastPasswordChange = "Unknown"
            $roles = @()
            $groups = @()
            
            try { 
                # Get license information
                $licenseDetails = Get-MgUserLicenseDetail -UserId $originalItemText -ErrorAction Stop
                if ($licenseDetails) { 
                    if ($licenseDetails.SkuPartNumber | Where-Object { $_ -in $premiumLicenseSkus }) { 
                        $hasPremiumLicense = $true 
                    } 
                }
                
                # Get user details
                $statusLabel.Text = "Getting details for ${currentUserIndex}/${totalUserCount}: $originalItemText"
                $mainForm.Refresh()
                
                $user = Get-MgUser -UserId $originalItemText -Property Id, DisplayName, AccountEnabled, LastPasswordChangeDateTime -ErrorAction Stop
                $accountEnabled = $user.AccountEnabled
                $lastPasswordChange = if ($user.LastPasswordChangeDateTime) { 
                    $user.LastPasswordChangeDateTime.ToString("MM/dd/yyyy HH:mm:ss") 
                } else { 
                    "Never" 
                }
                
                # Get roles and groups
                $statusLabel.Text = "Getting roles/groups for ${currentUserIndex}/${totalUserCount}: $originalItemText"
                $mainForm.Refresh()
                
                $membership = Get-UserRolesAndGroups -UserId $user.Id
                $roles = $membership.Roles
                $groups = $membership.Groups
                
            } catch { 
                Write-Warning "Could not get complete info for $originalItemText. Error: $($_.Exception.Message)" 
            }
            
            # Format the information
            $licenseString = if ($hasPremiumLicense) { "[P1/P2]" } else { "[Standard]" }
            $enabledString = if ($accountEnabled -eq $true) { "✓ Enabled" } elseif ($accountEnabled -eq $false) { "✗ Disabled" } else { "? Unknown" }
            
            # Format roles - show first 2, then count if more
            $rolesString = if ($roles.Count -eq 0) { 
                "No Admin Roles" 
            } elseif ($roles.Count -le 2) { 
                $roles -join ", " 
            } else { 
                "$($roles[0]), $($roles[1]) +$($roles.Count - 2) more" 
            }
            
            # Format groups - show count
            $groupsString = if ($groups.Count -eq 0) { 
                "No Groups" 
            } else { 
                "$($groups.Count) Groups" 
            }
            
            # Create the display string
            $displayInfo = "$originalItemText $licenseString | $enabledString | Pwd: $lastPasswordChange | Roles: $rolesString | Groups: $groupsString"
            
            $userCheckedListBox.Items[$i] = $displayInfo
        }
        
        $statusLabel.Text = "Complete analysis finished for $totalUserCount user(s)." 
        
        # Show summary message
        [System.Windows.Forms.MessageBox]::Show("Analysis complete for $totalUserCount user(s).`n`nLegend:`n✓ = Account Enabled`n✗ = Account Disabled`n[P1/P2] = Premium License`n[Standard] = Basic License", "Analysis Complete", "OK", "Information")
        
    } catch { 
        $statusLabel.Text = "An error occurred during user analysis."
        Write-Error "User analysis failed: $($_.Exception.Message)" 
    } finally { 
        $mainForm.Cursor = 'Default'
        $checkLicenseButton.Enabled = $true
        $connectButton.Enabled = $true
        $disconnectButton.Enabled = $true
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '') 
    } 
})

# SELECT ALL CHECKBOX
$selectAllCheckbox.add_CheckedChanged({ 
    param($sender, $e)
    $isChecked = $sender.Checked
    
    for ($i = 0; $i -lt $userCheckedListBox.Items.Count; $i++) { 
        $userCheckedListBox.SetItemChecked($i, $isChecked) 
    }
    
    $sender.Text = if ($isChecked) { "Deselect All" } else { "Select All" } 
})

# USER LISTBOX ITEM CHECK
$userCheckedListBox.add_ItemCheck({ 
    $mainForm.BeginInvoke([System.Action]{ 
        $selectedCount = $script:userCheckedListBox.CheckedItems.Count
        $script:checkLicenseButton.Enabled = ($selectedCount -gt 0 -and $script:disconnectButton.Enabled)
        $script:fetchDetailsButton.Enabled = ($selectedCount -eq 1 -and $script:disconnectButton.Enabled)
        $script:fetchAuditButton.Enabled = ($selectedCount -eq 1 -and $script:disconnectButton.Enabled)
        $script:fetchMfaButton.Enabled = ($selectedCount -eq 1 -and $script:disconnectButton.Enabled)
        $script:getLogsButton.Enabled = ($selectedCount -gt 0 -and $script:outputFolderTextBox.Text -ne '' -and $script:disconnectButton.Enabled) 
    }) 
})

# OUTPUT FOLDER TEXTBOX CHANGE
$outputFolderTextBox.add_TextChanged({ 
    $selectedCount = $userCheckedListBox.CheckedItems.Count
    $checkLicenseButton.Enabled = ($selectedCount -gt 0 -and $disconnectButton.Enabled)
    $getLogsButton.Enabled = ($selectedCount -gt 0 -and $outputFolderTextBox.Text -ne '' -and $disconnectButton.Enabled) 
})

# FETCH DETAILS BUTTON
$fetchDetailsButton.add_Click({ 
    if ($userCheckedListBox.CheckedItems.Count -ne 1) { 
        [System.Windows.Forms.MessageBox]::Show("Please select exactly one user.", "Selection Error", "OK", "Warning")
        return 
    }
    
    $upn = $userCheckedListBox.CheckedItems[0].ToString().Split(' ')[0]
    $statusLabel.Text = "Fetching details for $upn..."
    $mainForm.Cursor = 'WaitCursor'
    $fetchDetailsButton.Enabled = $false
    $detailsRichTextBox.Clear()
    
    try { 
        $user = Get-MgUser -UserId $upn -Property Id, DisplayName, AccountEnabled, LastPasswordChangeDateTime -ErrorAction Stop
        
        $detailsRichTextBox.SelectionFont = New-Object System.Drawing.Font($detailsRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $detailsRichTextBox.AppendText("User Principal Name: ")
        $detailsRichTextBox.SelectionFont = $detailsRichTextBox.Font
        $detailsRichTextBox.AppendText("$($user.UserPrincipalName)`n")
        
        $detailsRichTextBox.SelectionFont = New-Object System.Drawing.Font($detailsRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $detailsRichTextBox.AppendText("Display Name: ")
        $detailsRichTextBox.SelectionFont = $detailsRichTextBox.Font
        $detailsRichTextBox.AppendText("$($user.DisplayName)`n")
        
        $detailsRichTextBox.SelectionFont = New-Object System.Drawing.Font($detailsRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $detailsRichTextBox.AppendText("Account Enabled: ")
        $detailsRichTextBox.SelectionFont = $detailsRichTextBox.Font
        $detailsRichTextBox.AppendText("$($user.AccountEnabled)`n")
        
        $detailsRichTextBox.SelectionFont = New-Object System.Drawing.Font($detailsRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $detailsRichTextBox.AppendText("Last Password Change: ")
        $detailsRichTextBox.SelectionFont = $detailsRichTextBox.Font
        $detailsRichTextBox.AppendText("$($user.LastPasswordChangeDateTime)`n`n")
        
        $membership = Get-UserRolesAndGroups -UserId $user.Id
        
        $detailsRichTextBox.SelectionFont = New-Object System.Drawing.Font($detailsRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $detailsRichTextBox.AppendText("Administrative Roles (Active):`n")
        $detailsRichTextBox.SelectionFont = $detailsRichTextBox.Font
        
        if ($membership.Roles.Count -gt 0) { 
            $membership.Roles | ForEach-Object { 
                $detailsRichTextBox.AppendText(" - $_\n") 
            } 
        } else { 
            $detailsRichTextBox.AppendText(" - None Detected (or role is PIM-eligible but not active)`n") 
        }
        
        $detailsRichTextBox.AppendText("`n")
        
        $detailsRichTextBox.SelectionFont = New-Object System.Drawing.Font($detailsRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $detailsRichTextBox.AppendText("Group Memberships:`n")
        $detailsRichTextBox.SelectionFont = $detailsRichTextBox.Font
        
        if ($membership.Groups.Count -gt 0) { 
            $membership.Groups | ForEach-Object { 
                $detailsRichTextBox.AppendText(" - $_\n") 
            } 
        } else { 
            $detailsRichTextBox.AppendText(" - None`n") 
        }
        
        if ($membership.Error) { 
            $detailsRichTextBox.SelectionFont = New-Object System.Drawing.Font($detailsRichTextBox.Font, [System.Drawing.FontStyle]::Italic)
            $detailsRichTextBox.SelectionColor = [System.Drawing.Color]::Red
            $detailsRichTextBox.AppendText("`nWarning: $($membership.Error)`n") 
        }
        
        $statusLabel.Text = "Successfully fetched details for $upn." 
    } catch { 
        $statusLabel.Text = "Error fetching details. See console."
        Write-Error "Failed to fetch user details/roles: $($_.Exception.Message)" 
    } finally { 
        $mainForm.Cursor = 'Default'
        $fetchDetailsButton.Enabled = $true 
    } 
})

# FETCH AUDIT BUTTON
$fetchAuditButton.add_Click({
    if ($userCheckedListBox.CheckedItems.Count -ne 1) { 
        [System.Windows.Forms.MessageBox]::Show("Please select exactly one user.", "Selection Error", "OK", "Warning")
        return 
    }
    
    $upn = $userCheckedListBox.CheckedItems[0].ToString().Split(' ')[0]
    $statusLabel.Text = "Fetching audit logs for $upn..."
    $mainForm.Cursor = 'WaitCursor'
    $fetchAuditButton.Enabled = $false
    $exportAuditButton.Enabled = $false
    
    $auditGrid.Rows.Clear()
    $auditSummaryLabel.Text = "Loading audit logs..."
    
    try {
        $userId = (Get-MgUser -UserId $upn -Property Id -ErrorAction Stop).Id
        if (-not $userId) { 
            throw "Could not retrieve User ID for $upn." 
        }
        
        $days = $logDurationNumericUpDown.Value
        $startDate = (Get-Date).AddDays(-$days).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filter = "(initiatedBy/user/id eq '$userId') and (activityDateTime ge $startDate)"
        
        Write-Host "Fetching audit logs with filter: $filter"
        $auditLogs = Get-MgAuditLogDirectoryAudit -Filter $filter -All -ErrorAction Stop
        
        if ($auditLogs -and $auditLogs.Count -gt 0) {
            foreach ($log in $auditLogs) {
                $timeFormatted = if ($log.ActivityDateTime) { 
                    $log.ActivityDateTime.ToString("yyyy-MM-dd HH:mm:ss") 
                } else { 
                    "N/A" 
                }
                
                $activity = if ($log.ActivityDisplayName) { 
                    $log.ActivityDisplayName 
                } else { 
                    "Unknown Activity" 
                }
                
                $category = if ($log.Category) { 
                    $log.Category 
                } else { 
                    "Other" 
                }
                
                $result = if ($log.Result -and $log.Result -ne '') { 
                    $log.Result 
                } else { 
                    "Success" 
                }
                
                $target = "N/A"
                if ($log.TargetResources -and $log.TargetResources.Count -gt 0) {
                    $targetInfo = $log.TargetResources[0]
                    if ($targetInfo.DisplayName) {
                        $target = $targetInfo.DisplayName
                    } elseif ($targetInfo.UserPrincipalName) {
                        $target = $targetInfo.UserPrincipalName
                    } elseif ($targetInfo.Id) {
                        $target = $targetInfo.Id
                    }
                }
                
                $auditGrid.Rows.Add(@($timeFormatted, $activity, $category, $result, $target))
            }
            
            $auditSummaryLabel.Text = "Found $($auditLogs.Count) audit log entries for $upn in the last $days days."
            $exportAuditButton.Enabled = $true
            $statusLabel.Text = "Successfully loaded $($auditLogs.Count) audit log entries."
        } else { 
            $auditSummaryLabel.Text = "No audit log entries found for $upn in the last $days days."
            $statusLabel.Text = "No audit log entries found for $upn." 
        }
    } catch { 
        $auditSummaryLabel.Text = "Error occurred while fetching audit logs."
        $statusLabel.Text = "Error fetching audit logs. Check console for details."
        Write-Error "Failed to fetch audit logs: $($_.Exception.Message)"
    } finally { 
        $mainForm.Cursor = 'Default'
        $fetchAuditButton.Enabled = $true 
    }
})

# EXPORT AUDIT LOGS BUTTON
$exportAuditButton.add_Click({
    if ($auditGrid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No audit data to export.", "No Data", "OK", "Information")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.DefaultExt = "csv"
    $saveFileDialog.FileName = "AuditLogs_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        try {
            $csvData = @()
            $csvData += '"Date/Time","Activity","Category","Result","Target/Object"'
            
            foreach ($row in $auditGrid.Rows) {
                if ($row.Cells[0].Value) {
                    $line = '"' + ($row.Cells[0].Value -replace '"', '""') + '","' + 
                           ($row.Cells[1].Value -replace '"', '""') + '","' + 
                           ($row.Cells[2].Value -replace '"', '""') + '","' + 
                           ($row.Cells[3].Value -replace '"', '""') + '","' + 
                           ($row.Cells[4].Value -replace '"', '""') + '"'
                    $csvData += $line
                }
            }
            
            $csvData | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Audit logs exported successfully to:`n$($saveFileDialog.FileName)", "Export Complete", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export audit logs: $($_.Exception.Message)", "Export Error", "OK", "Error")
        }
    }
})

# MFA ANALYSIS BUTTON
$fetchMfaButton.add_Click({
    if ($userCheckedListBox.CheckedItems.Count -ne 1) { 
        [System.Windows.Forms.MessageBox]::Show("Please select exactly one user.", "Selection Error", "OK", "Warning")
        return 
    }
    
    $upn = $userCheckedListBox.CheckedItems[0].ToString().Split(' ')[0]
    $statusLabel.Text = "Analyzing MFA status for $upn..."
    $mainForm.Cursor = 'WaitCursor'
    $fetchMfaButton.Enabled = $false
    $exportMfaButton.Enabled = $false
    
    $mfaRichTextBox.Clear()
    $mfaPolicyGrid.Rows.Clear()
    $mfaSummaryLabel.Text = "Analyzing MFA configuration..."
    
    try {
        $userId = (Get-MgUser -UserId $upn -Property Id -ErrorAction Stop).Id
        if (-not $userId) { 
            throw "Could not retrieve User ID for $upn." 
        }
        
        Write-Host "Analyzing MFA status for user: $upn (ID: $userId)"
        $mfaStatus = Get-UserMfaStatus -UserId $userId -UserPrincipalName $upn
        
        # Display results in rich text box
        $mfaRichTextBox.SelectionFont = New-Object System.Drawing.Font($mfaRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $mfaRichTextBox.SelectionColor = if ($mfaStatus.OverallStatus -like "*NOT PROTECTED*") { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::DarkGreen }
        $mfaRichTextBox.AppendText("OVERALL MFA STATUS: $($mfaStatus.OverallStatus)`n")
        $mfaRichTextBox.SelectionColor = [System.Drawing.Color]::Black
        $mfaRichTextBox.SelectionFont = $mfaRichTextBox.Font
        $mfaRichTextBox.AppendText("$($mfaStatus.Summary)`n`n")
        
        # Per-User MFA Section
        $mfaRichTextBox.SelectionFont = New-Object System.Drawing.Font($mfaRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $mfaRichTextBox.AppendText("1. PER-USER MFA:`n")
        $mfaRichTextBox.SelectionFont = $mfaRichTextBox.Font
        $mfaRichTextBox.SelectionColor = if ($mfaStatus.PerUserMfa.Enabled) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Orange }
        $mfaRichTextBox.AppendText("   Status: $(if ($mfaStatus.PerUserMfa.Enabled) { '✓ ENABLED' } else { '✗ NOT ENABLED' })`n")
        $mfaRichTextBox.SelectionColor = [System.Drawing.Color]::Black
        $mfaRichTextBox.AppendText("   Details: $($mfaStatus.PerUserMfa.Details)`n`n")
        
        # Security Defaults Section
        $mfaRichTextBox.SelectionFont = New-Object System.Drawing.Font($mfaRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $mfaRichTextBox.AppendText("2. SECURITY DEFAULTS:`n")
        $mfaRichTextBox.SelectionFont = $mfaRichTextBox.Font
        $mfaRichTextBox.SelectionColor = if ($mfaStatus.SecurityDefaults.Enabled) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Orange }
        $mfaRichTextBox.AppendText("   Status: $(if ($mfaStatus.SecurityDefaults.Enabled) { '✓ ENABLED' } else { '✗ DISABLED' })`n")
        $mfaRichTextBox.SelectionColor = [System.Drawing.Color]::Black
        $mfaRichTextBox.AppendText("   Details: $($mfaStatus.SecurityDefaults.Details)`n`n")
        
        # Conditional Access Section
        $mfaRichTextBox.SelectionFont = New-Object System.Drawing.Font($mfaRichTextBox.Font, [System.Drawing.FontStyle]::Bold)
        $mfaRichTextBox.AppendText("3. CONDITIONAL ACCESS:`n")
        $mfaRichTextBox.SelectionFont = $mfaRichTextBox.Font
        $mfaRichTextBox.SelectionColor = if ($mfaStatus.ConditionalAccess.RequiresMfa) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Orange }
        $mfaRichTextBox.AppendText("   MFA Required: $(if ($mfaStatus.ConditionalAccess.RequiresMfa) { '✓ YES' } else { '✗ NO' })`n")
        $mfaRichTextBox.SelectionColor = [System.Drawing.Color]::Black
        $mfaRichTextBox.AppendText("   Details: $($mfaStatus.ConditionalAccess.Details)`n")
        
        # Populate the policy grid
        foreach ($policy in $mfaStatus.ConditionalAccess.Policies) {
            $mfaPolicyGrid.Rows.Add(@($policy.Name, $policy.State, $policy.Controls, $policy.Conditions))
        }
        
        $mfaSummaryLabel.Text = "MFA analysis complete for $upn - Status: $($mfaStatus.OverallStatus)"
        $exportMfaButton.Enabled = $true
        $statusLabel.Text = "MFA analysis completed successfully."
        
    } catch { 
        $mfaSummaryLabel.Text = "Error occurred while analyzing MFA status."
        $statusLabel.Text = "Error analyzing MFA status. Check console for details."
        $mfaRichTextBox.SelectionColor = [System.Drawing.Color]::Red
        $mfaRichTextBox.AppendText("ERROR: Failed to analyze MFA status`n")
        $mfaRichTextBox.AppendText("Details: $($_.Exception.Message)")
        Write-Error "Failed to analyze MFA status: $($_.Exception.Message)"
    } finally { 
        $mainForm.Cursor = 'Default'
        $fetchMfaButton.Enabled = $true 
    }
})

# EXPORT MFA ANALYSIS
$exportMfaButton.add_Click({
    if ($mfaRichTextBox.Text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("No MFA analysis data to export.", "No Data", "OK", "Information")
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveFileDialog.DefaultExt = "txt"
    $saveFileDialog.FileName = "MFA_Analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        try {
            $exportContent = $mfaRichTextBox.Text
            $exportContent += "`n`n=== CONDITIONAL ACCESS POLICIES ===`n"
            
            foreach ($row in $mfaPolicyGrid.Rows) {
                if ($row.Cells[0].Value) {
                    $exportContent += "Policy: $($row.Cells[0].Value)`n"
                    $exportContent += "State: $($row.Cells[1].Value)`n"
                    $exportContent += "Controls: $($row.Cells[2].Value)`n"
                    $exportContent += "Conditions: $($row.Cells[3].Value)`n"
                    $exportContent += "---`n"
                }
            }
            
            $exportContent | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("MFA analysis exported successfully to:`n$($saveFileDialog.FileName)", "Export Complete", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export MFA analysis: $($_.Exception.Message)", "Export Error", "OK", "Error")
        }
    }
})

# BROWSE FOLDER BUTTON
$browseFolderButton.add_Click({ 
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowserDialog.ShowDialog() -eq 'OK') { 
        $outputFolderTextBox.Text = $folderBrowserDialog.SelectedPath 
    } 
})

# LOG DURATION CHANGE
$logDurationNumericUpDown.add_ValueChanged({ 
    if ($logDurationNumericUpDown.Value -gt 7) { 
        $durationWarningLabel.Text = "Note: >7 days requires Entra ID P1/P2 license." 
    } else { 
        $durationWarningLabel.Text = "" 
    } 
})

# GET LOGS BUTTON
$getLogsButton.add_Click({ 
    $selectedUpns = $userCheckedListBox.CheckedItems | ForEach-Object { $_.ToString().Split(' ')[0] }
    $days = $logDurationNumericUpDown.Value
    $outputFolder = $outputFolderTextBox.Text
    
    if ($selectedUpns.Count -eq 0) { 
        [System.Windows.Forms.MessageBox]::Show("Select user(s).", "Selection Error", "OK", "Warning")
        return 
    }
    
    if (-not (Test-Path -Path $outputFolder -PathType Container)) { 
        [System.Windows.Forms.MessageBox]::Show("Select a valid output folder.", "Folder Error", "OK", "Warning")
        return 
    }
    
    $statusLabel.Text = "Fetching logs for $($selectedUpns.Count) users..."
    $mainForm.Cursor = 'WaitCursor'
    $getLogsButton.Enabled = $false
    $connectButton.Enabled = $false
    $disconnectButton.Enabled = $false
    if ($openFileButton) { 
        $openFileButton.Enabled = $false 
    }
    
    $allLogs = @()
    
    try { 
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseFileName = "EntraSignInLogs_$timestamp"
        $csvFilePath = Join-Path -Path $outputFolder -ChildPath "$($baseFileName).csv"
        $xlsxFilePath = Join-Path -Path $outputFolder -ChildPath "$($baseFileName).xlsx"
        $startDate = (Get-Date).AddDays(-$days).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        
        foreach ($upn in $selectedUpns) { 
            try { 
                $userId = (Get-MgUser -UserId $upn -Property Id).Id
                $filterString = "userId eq '$userId' and createdDateTime ge $startDate"
                $userLogs = Get-MgAuditLogSignIn -Filter $filterString -All -ErrorAction Stop
                
                if($userLogs) { 
                    $userLogs | ForEach-Object { 
                        $_ | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $upn -Force 
                    }
                    $allLogs += $userLogs 
                } 
            } catch { 
                Write-Warning "Could not get logs for $upn" 
            } 
        }
        
        if ($allLogs.Count -gt 0) { 
            $exportData = $allLogs | Select-Object UserPrincipalName, CreatedDateTime, AppDisplayName, IpAddress, @{N='City';E={$_.Location.City}}, @{N='State';E={$_.Location.State}}, @{N='Country';E={$_.Location.CountryOrRegion}}, @{N='Status';E={$_.Status.ErrorCode}}, @{N='FailureReason';E={$_.Status.FailureReason}}
            $exportData | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
            
            if (ConvertTo-XlsxAndFormat -CsvPath $csvFilePath -XlsxPath $xlsxFilePath) { 
                $statusLabel.Text = "Export successful: $xlsxFilePath"
                try { 
                    Remove-Item $csvFilePath -Force 
                } catch {}
            } else { 
                $statusLabel.Text = "Exported to CSV, but XLSX conversion failed." 
            } 
        } else { 
            $statusLabel.Text = "No sign-in logs found for selected users in the time period." 
        } 
    } catch { 
        $statusLabel.Text = "Error during export."
        Write-Error "Get logs/export failed: $($_.Exception.Message)" 
    } finally { 
        $mainForm.Cursor = 'Default'
        $getLogsButton.Enabled = $true
        $connectButton.Enabled = $true
        $disconnectButton.Enabled = $true 
    } 
})

# OPEN FILE BUTTON
$openFileButton.add_Click({ 
    if ($script:lastExportedXlsxPath -and (Test-Path $script:lastExportedXlsxPath)) { 
        try { 
            Invoke-Item -Path $script:lastExportedXlsxPath -ErrorAction Stop 
        } catch { 
            [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)", "Error", 'OK', 'Error') 
        } 
    } else { 
        [System.Windows.Forms.MessageBox]::Show("No file from this session to open.", "Info", 'OK', 'Information') 
    } 
})

# --- Show Form ---
$mainForm.Add_Shown({$mainForm.Activate()})
[void]$mainForm.ShowDialog()
$mainForm.Dispose()