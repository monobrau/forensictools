<#
.SYNOPSIS
A PowerShell script with a GUI to check and enable Entra ID auditing, and to download audit logs for all or specific users.

.NOTES
Author: Gemini
Date: 2025-07-15
Version: 1.1 (Added open file button)
Requires:
    - PowerShell 5.1+
    - Microsoft Graph SDK (Authentication, Reports)
Permissions: Requires delegated AuditLog.Read.All, Directory.Read.All, Policy.Read.All
#>

# --- Configuration ---
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Reports", "Microsoft.Graph.Users")
$requiredScopes = @("AuditLog.Read.All", "Directory.Read.All", "Policy.Read.All")

# --- Script-level variables ---
$script:lastExportedCsvPath = $null

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
        Write-Host "Installing required PowerShell modules: $($Modules -join ', ')"
        Install-Module -Name $Modules -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
    } catch {
        Write-Error "Failed to install one or more required modules. Please install them manually and restart the script."
        Exit
    }
}

# --- Prerequisites Check ---
$missing = Test-Modules -Modules $requiredModules
if ($missing.Count -gt 0) {
    if ('Yes' -eq [System.Windows.Forms.MessageBox]::Show("The following required modules are missing: $($missing -join ', '). Do you want to install them now?", "Missing Modules", 'YesNo', 'Warning')) {
        Install-MissingModules -Modules $missing
        [System.Windows.Forms.MessageBox]::Show("Modules have been installed. Please restart the PowerShell session and run the script again.", "Restart Required", 'OK', 'Information')
        Exit
    } else {
        [System.Windows.Forms.MessageBox]::Show("The script cannot continue without the required modules.", "Error", 'OK', 'Error')
        Exit
    }
}

# --- GUI Setup ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Entra ID Audit Log Manager v1.1"
$mainForm.Size = '800, 650' # Increased height for new button
$mainForm.StartPosition = 'CenterScreen'
$mainForm.FormBorderStyle = 'FixedDialog'
$mainForm.MaximizeBox = $false

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready. Please connect to Entra ID."
$statusStrip.Items.Add($statusLabel)
$mainForm.Controls.Add($statusStrip)

# --- Connection Controls ---
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Location = '10, 10'
$connectButton.Size = '120, 30'
$connectButton.Text = "Connect"
$mainForm.Controls.Add($connectButton)

# --- Auditing Status Controls ---
$auditGroupBox = New-Object System.Windows.Forms.GroupBox
$auditGroupBox.Location = '140, 10'
$auditGroupBox.Size = '630, 80'
$auditGroupBox.Text = "Auditing Status"
$mainForm.Controls.Add($auditGroupBox)

$auditStatusLabel = New-Object System.Windows.Forms.Label
$auditStatusLabel.Location = '20, 30'
$auditStatusLabel.Size = '400, 20'
$auditStatusLabel.Text = "Status: Unknown (Connect to check)"
$auditStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$auditGroupBox.Controls.Add($auditStatusLabel)

$enableAuditButton = New-Object System.Windows.Forms.Button
$enableAuditButton.Location = '450, 25'
$enableAuditButton.Size = '160, 30'
$enableAuditButton.Text = "Enable Auditing"
$enableAuditButton.Enabled = $false
$auditGroupBox.Controls.Add($enableAuditButton)

# --- User Selection Controls ---
$userListLabel = New-Object System.Windows.Forms.Label
$userListLabel.Location = '10, 100'
$userListLabel.Size = '200, 20'
$userListLabel.Text = "Available Users:"
$mainForm.Controls.Add($userListLabel)

$userCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$userCheckedListBox.Location = '10, 125'
$userCheckedListBox.Size = '760, 200'
$userCheckedListBox.CheckOnClick = $true
$userCheckedListBox.HorizontalScrollbar = $true
$userCheckedListBox.Enabled = $false
$mainForm.Controls.Add($userCheckedListBox)

$selectAllCheckbox = New-Object System.Windows.Forms.CheckBox
$selectAllCheckbox.Location = '10, 330'
$selectAllCheckbox.Size = '120, 20'
$selectAllCheckbox.Text = "Select All"
$selectAllCheckbox.Enabled = $false
$mainForm.Controls.Add($selectAllCheckbox)

# --- Log Download Controls ---
$downloadGroupBox = New-Object System.Windows.Forms.GroupBox
$downloadGroupBox.Location = '10, 360'
$downloadGroupBox.Size = '760, 230' # Increased height for new button
$downloadGroupBox.Text = "Download Audit Logs"
$mainForm.Controls.Add($downloadGroupBox)

$logDurationLabel = New-Object System.Windows.Forms.Label
$logDurationLabel.Location = '20, 30'
$logDurationLabel.Size = '150, 20'
$logDurationLabel.Text = "Log History (Days):"
$downloadGroupBox.Controls.Add($logDurationLabel)

$logDurationNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
$logDurationNumericUpDown.Location = '170, 28'
$logDurationNumericUpDown.Size = '60, 25'
$logDurationNumericUpDown.Minimum = 1
$logDurationNumericUpDown.Maximum = 90
$logDurationNumericUpDown.Value = 7
$downloadGroupBox.Controls.Add($logDurationNumericUpDown)

$downloadAllButton = New-Object System.Windows.Forms.Button
$downloadAllButton.Location = '20, 70'
$downloadAllButton.Size = '720, 40'
$downloadAllButton.Text = "Download ALL Audit Logs"
$downloadAllButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$downloadAllButton.Enabled = $false
$downloadGroupBox.Controls.Add($downloadAllButton)

$downloadSelectedButton = New-Object System.Windows.Forms.Button
$downloadSelectedButton.Location = '20, 120'
$downloadSelectedButton.Size = '720, 40'
$downloadSelectedButton.Text = "Download Logs for SELECTED Users"
$downloadSelectedButton.Enabled = $false
$downloadGroupBox.Controls.Add($downloadSelectedButton)

# *** NEW BUTTON ***
$openCsvButton = New-Object System.Windows.Forms.Button
$openCsvButton.Location = '20, 170'
$openCsvButton.Size = '720, 40'
$openCsvButton.Text = "Open Last Exported CSV"
$openCsvButton.Enabled = $false
$downloadGroupBox.Controls.Add($openCsvButton)

# --- Functions for Core Logic ---
Function Check-AuditStatus {
    $mainForm.Cursor = 'WaitCursor'
    $statusLabel.Text = "Checking audit status..."
    try {
        Get-MgAuditLogDirectoryAudit -Top 1 | Out-Null
        $auditStatusLabel.Text = "Status: Auditing is ENABLED"
        $auditStatusLabel.ForeColor = [System.Drawing.Color]::Green
        $enableAuditButton.Enabled = $false
        $downloadAllButton.Enabled = $true
        return $true
    } catch {
        if ($_.Exception.Message -like "*could not be found*") {
            $auditStatusLabel.Text = "Status: Auditing is DISABLED or no logs exist"
            $auditStatusLabel.ForeColor = [System.Drawing.Color]::Red
            $enableAuditButton.Enabled = $true
            $downloadAllButton.Enabled = $false
        } else {
             $auditStatusLabel.Text = "Status: Error checking status."
             $auditStatusLabel.ForeColor = [System.Drawing.Color]::OrangeRed
             Write-Warning "An unexpected error occurred while checking audit status: $($_.Exception.Message)"
        }
        return $false
    } finally {
        $statusLabel.Text = "Audit status check complete."
        $mainForm.Cursor = 'Default'
    }
}

Function Enable-Auditing {
    $mainForm.Cursor = 'WaitCursor'
    $statusLabel.Text = "Attempting to enable auditing..."
    try {
        [System.Windows.Forms.MessageBox]::Show("To enable Entra ID auditing, your tenant must have a valid license (e.g., Entra ID P1 or P2).`n`nAuditing will begin automatically once a valid license is applied. There is no specific 'enable' button via the Graph API for general auditing; it is license-activated.", "How to Enable Auditing", 'OK', 'Information')
        Check-AuditStatus
    } catch {
        $statusLabel.Text = "Error."
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", 'OK', 'Error')
    } finally {
        $mainForm.Cursor = 'Default'
    }
}

Function Load-Users {
    $mainForm.Cursor = 'WaitCursor'
    $statusLabel.Text = "Fetching all users..."
    try {
        $allUsers = Get-MgUser -All -Property UserPrincipalName, Id -ConsistencyLevel eventual | Sort-Object UserPrincipalName
        if ($allUsers) {
            $userCheckedListBox.Items.Clear()
            $userCheckedListBox.Items.AddRange($allUsers.UserPrincipalName)
            $statusLabel.Text = "Loaded $($allUsers.Count) users."
            $userCheckedListBox.Enabled = $true
            $selectAllCheckbox.Enabled = $true
        } else {
            $statusLabel.Text = "Connected, but no users were found."
        }
    } catch {
        $statusLabel.Text = "Failed to load users."
        Write-Error "User loading failed: $($_.Exception.Message)"
    } finally {
        $mainForm.Cursor = 'Default'
    }
}

Function Download-Logs {
    param(
        [System.Collections.ArrayList]$UserPrincipalNames
    )

    $days = $logDurationNumericUpDown.Value
    $startDate = (Get-Date).AddDays(-$days).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.FileName = "EntraAuditLogs_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($saveFileDialog.ShowDialog() -ne 'OK') {
        return
    }
    $filePath = $saveFileDialog.FileName

    $mainForm.Cursor = 'WaitCursor'
    $downloadAllButton.Enabled = $false
    $downloadSelectedButton.Enabled = $false
    $openCsvButton.Enabled = $false
    
    try {
        $filterStrings = @()
        if ($null -ne $UserPrincipalNames -and $UserPrincipalNames.Count -gt 0) {
            $statusLabel.Text = "Fetching user IDs..."
            foreach ($upn in $UserPrincipalNames) {
                try {
                    $userId = (Get-MgUser -UserId $upn -Property Id).Id
                    $filterStrings.Add("(initiatedBy/user/id eq '$userId')")
                } catch {
                    Write-Warning "Could not find user '$upn'. Skipping."
                }
            }
        }
        
        $finalFilter = "(activityDateTime ge $startDate)"
        if ($filterStrings.Count -gt 0) {
            $userFilters = $filterStrings -join " or "
            $finalFilter += " and ($userFilters)"
        }
        
        $statusLabel.Text = "Downloading logs. This may take some time..."
        $allLogs = Get-MgAuditLogDirectoryAudit -Filter $finalFilter -All
        
        if ($allLogs.Count -gt 0) {
            $statusLabel.Text = "Processing $($allLogs.Count) log entries..."
            $exportData = $allLogs | Select-Object `
                @{N='DateTime';E={$_.ActivityDateTime}}, `
                @{N='User';E={$_.InitiatedBy.User.UserPrincipalName}}, `
                @{N='Activity';E={$_.ActivityDisplayName}}, `
                Category, `
                Result, `
                @{N='TargetResource';E={$_.TargetResources[0].DisplayName}}, `
                @{N='TargetResourceType';E={$_.TargetResources[0].Type}}
            
            $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
            
            # *** MODIFICATION: Store path and enable button ***
            $script:lastExportedCsvPath = $filePath
            $openCsvButton.Enabled = $true
            
            $statusLabel.Text = "Export complete!"
            [System.Windows.Forms.MessageBox]::Show("$($allLogs.Count) log entries exported to:`n$filePath", "Export Complete", 'OK', 'Information')
        } else {
            $statusLabel.Text = "No audit logs found for the specified criteria."
            [System.Windows.Forms.MessageBox]::Show("No audit logs found within the last $days days for the selected scope.", "No Logs Found", 'OK', 'Information')
        }
    } catch {
        $statusLabel.Text = "An error occurred during download."
        Write-Error "Failed to download audit logs: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to download audit logs: $($_.Exception.Message)", "Export Error", 'OK', 'Error')
    } finally {
        $mainForm.Cursor = 'Default'
        $downloadAllButton.Enabled = $true
        $downloadSelectedButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0)
    }
}


# --- Event Handlers ---
$connectButton.add_Click({
    $statusLabel.Text = "Connecting to Microsoft Graph..."
    $mainForm.Cursor = 'WaitCursor'
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
        
        $statusLabel.Text = "Connected successfully."
        $connectButton.Text = "Reconnect"
        
        # *** MODIFICATION: Reset state on new connection ***
        $script:lastExportedCsvPath = $null
        $openCsvButton.Enabled = $false
        
        if (Check-AuditStatus) {
            Load-Users
        }
    } catch {
        $statusLabel.Text = "Connection failed."
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to Microsoft Graph. Error: $($_.Exception.Message)", "Connection Error", 'OK', 'Error')
    } finally {
        $mainForm.Cursor = 'Default'
    }
})

$enableAuditButton.add_Click({
    Enable-Auditing
})

$selectAllCheckbox.add_CheckedChanged({
    param($sender, $e)
    $isChecked = $sender.Checked
    for ($i = 0; $i -lt $userCheckedListBox.Items.Count; $i++) {
        $userCheckedListBox.SetItemChecked($i, $isChecked)
    }
})

$userCheckedListBox.add_ItemCheck({
    $mainForm.BeginInvoke([System.Action]{
        $downloadSelectedButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0)
    })
})

$downloadAllButton.add_Click({
    Download-Logs -UserPrincipalNames $null
})

$downloadSelectedButton.add_Click({
    $selectedUsers = [System.Collections.ArrayList]@($userCheckedListBox.CheckedItems)
    if ($selectedUsers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user from the list.", "No Users Selected", "OK", "Warning")
        return
    }
    Download-Logs -UserPrincipalNames $selectedUsers
})

# *** NEW EVENT HANDLER ***
$openCsvButton.add_Click({
    if ($script:lastExportedCsvPath -and (Test-Path $script:lastExportedCsvPath)) {
        try {
            # Use Invoke-Item to open the file with its default application
            Invoke-Item -Path $script:lastExportedCsvPath -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not open the file. Ensure you have an application (like Excel) associated with .csv files.`nError: $($_.Exception.Message)", "Error Opening File", 'OK', 'Error')
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No file has been exported in this session, or the previously exported file has been moved or deleted.", "File Not Found", 'OK', 'Information')
    }
})

# --- Show Form ---
$mainForm.Add_Shown({$mainForm.Activate()})
[void]$mainForm.ShowDialog()
$mainForm.Dispose()