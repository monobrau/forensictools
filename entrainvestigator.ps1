<#
.SYNOPSIS
A PowerShell script with a GUI to fetch Entra ID sign-in logs for selected users,
and export to CSV, expanding complex properties like Status and DeviceDetail. Filters logs by User ID for better reliability.
Users are loaded automatically upon successful connection to Microsoft Graph.

.DESCRIPTION
This script provides a Windows Forms interface to:
- Connect to Microsoft Graph (and automatically load Entra ID users).
- Select users for investigation.
- Select the duration (1-30 days) for sign-in log history, with license warnings.
- Select an output folder.
- Fetch sign-in logs for the selected users (using User ID filter) and duration.
- Export logs directly to CSV format, with expanded Status and DeviceDetail properties for readability.

.NOTES
Author: Gemini
Date: 2025-05-06
Version: 2.0 (Users now load automatically after connecting to Graph; removed separate 'Load Users' button)
Requires: PowerShell 5.1+, Microsoft Graph SDK (Users, Reports).
Permissions: Requires delegated User.Read.All and AuditLog.Read.All permissions in Entra ID.

.LINK
Install Modules: Install-Module Microsoft.Graph.Users, Microsoft.Graph.Reports -Scope CurrentUser -Force

.EXAMPLE
.\EntraID_Forensic_Log_Fetcher.ps1
#>

#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Reports

# --- Configuration ---
$requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Reports")
$requiredScopes = @("User.Read.All", "AuditLog.Read.All")

# --- Function Definitions ---

Function Test-Modules {
    param($Modules)
    $missingModules = @()
    foreach ($moduleName in $Modules) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            $missingModules += $moduleName
        }
    }
    return $missingModules
}

Function Install-MissingModules {
    param($Modules)
    Write-Host "Attempting to install missing modules: $($Modules -join ', ')" -ForegroundColor Yellow
    try {
        Install-Module -Name $Modules -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        Write-Host "Modules installed successfully. Please restart the script." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install modules. Please install them manually: Install-Module -Name $($Modules -join ', ') -Scope CurrentUser"
        if ($statusLabel) {
            $statusLabel.Text = "Error installing modules. See console."
        }
    }
}

# --- Check Prerequisites ---
$missing = Test-Modules -Modules $requiredModules
if ($missing.Count -gt 0) {
    $choice = [System.Windows.Forms.MessageBox]::Show("Required PowerShell modules are missing: $($missing -join ', ').`n`nDo you want to attempt installation now? (Requires internet connection and administrator privileges if installing for AllUsers)", "Missing Modules", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        Install-MissingModules -Modules $missing
    } else {
        [System.Windows.Forms.MessageBox]::Show("Script cannot continue without required modules. Please install them manually and restart.", "Prerequisites Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Exit
    }
    [System.Windows.Forms.MessageBox]::Show("Please restart the script after module installation.", "Restart Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Exit
}

# Import necessary modules after check/install
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Reports

# --- GUI Setup ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Main Form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Entra ID Forensic Log Fetcher"
$mainForm.Size = New-Object System.Drawing.Size(600, 580)
$mainForm.MinimumSize = New-Object System.Drawing.Size(550, 530)
$mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$mainForm.MaximizeBox = $false
$mainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

# Status Strip and Label
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Name = "statusLabel"
$statusLabel.Text = "Ready. Please connect to Microsoft Graph."
$statusStrip.Items.Add($statusLabel)
$mainForm.Controls.Add($statusStrip)

# Connect Button
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Location = New-Object System.Drawing.Point(20, 20)
$connectButton.Size = New-Object System.Drawing.Size(200, 30) # Made button wider
$connectButton.Text = "Connect & Load Users"       # Updated Text
$connectButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$connectButton.add_Click({
    param($sender, $e)
    $statusLabel.Text = "Connecting to Microsoft Graph..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $userCheckedListBox.Items.Clear() # Clear previous user list
    $getLogsButton.Enabled = $false   # Disable get logs button initially

    try {
        # --- Connect to Graph ---
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
        $context = Get-MgContext
        $statusLabel.Text = "Connected as $($context.Account). Tenant: $($context.TenantId). Loading users..."
        $mainForm.Refresh() # Ensure status update is visible
        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green

        # --- Load Users (Moved here) ---
        Write-Host "Loading users..."
        # Fetch users - Use -All for complete list, but be mindful of performance
        $users = Get-MgUser -All -ErrorAction Stop -Select UserPrincipalName, Id, DisplayName -ConsistencyLevel eventual | Sort-Object UserPrincipalName
        
        if ($users) {
            foreach ($user in $users) {
                $userCheckedListBox.Items.Add($user.UserPrincipalName, $false)
            }
            $statusLabel.Text = "Connected. Loaded $($users.Count) users. Select users to investigate."
            Write-Host "Loaded $($users.Count) users." -ForegroundColor Green
            # $getLogsButton can be enabled if an output folder is already selected, handled by its own logic
        } else {
            $statusLabel.Text = "Connected. No users found or error loading users."
            [System.Windows.Forms.MessageBox]::Show("Connected to Graph, but no users found in the tenant or an error occurred during user loading.", "No Users Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }

    } catch {
        $statusLabel.Text = "Operation failed. Check console for errors."
        Write-Error "Microsoft Graph connection or user loading failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to Microsoft Graph or load users. Ensure you have internet connectivity and the necessary permissions. `n`nError: $($_.Exception.Message)", "Connection/Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        # Check if getLogsButton should be enabled based on current state
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '')
    }
})
$mainForm.Controls.Add($connectButton)

# User List Label
$userListLabel = New-Object System.Windows.Forms.Label
$userListLabel.Location = New-Object System.Drawing.Point(20, 65)
$userListLabel.Size = New-Object System.Drawing.Size(200, 20)
$userListLabel.Text = "Select User(s) to Investigate:"
$userListLabel.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$mainForm.Controls.Add($userListLabel)

# User CheckedListBox
$userCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$userCheckedListBox.Location = New-Object System.Drawing.Point(20, 90)
$userCheckedListBox.Size = New-Object System.Drawing.Size(545, 200)
$userCheckedListBox.CheckOnClick = $true
$userCheckedListBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$userCheckedListBox.add_ItemCheck({
    $mainForm.BeginInvoke([System.Action]{
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '')
    })
})
$mainForm.Controls.Add($userCheckedListBox)

# Log Duration Label
$logDurationLabel = New-Object System.Windows.Forms.Label
$logDurationLabel.Location = New-Object System.Drawing.Point(20, 305)
$logDurationLabel.Size = New-Object System.Drawing.Size(150, 20)
$logDurationLabel.Text = "Log History (Days):"
$logDurationLabel.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$mainForm.Controls.Add($logDurationLabel)

# Log Duration NumericUpDown Control
$logDurationNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
$logDurationNumericUpDown.Location = New-Object System.Drawing.Point(170, 305)
$logDurationNumericUpDown.Size = New-Object System.Drawing.Size(60, 25)
$logDurationNumericUpDown.Minimum = 1
$logDurationNumericUpDown.Maximum = 30
$logDurationNumericUpDown.Value = 7
$logDurationNumericUpDown.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$mainForm.Controls.Add($logDurationNumericUpDown)

# Warning Label for Duration
$durationWarningLabel = New-Object System.Windows.Forms.Label
$durationWarningLabel.Location = New-Object System.Drawing.Point(240, 308)
$durationWarningLabel.Size = New-Object System.Drawing.Size(325, 20)
$durationWarningLabel.Text = ""
$durationWarningLabel.ForeColor = [System.Drawing.Color]::OrangeRed
$durationWarningLabel.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$mainForm.Controls.Add($durationWarningLabel)

$logDurationNumericUpDown.add_ValueChanged({
    if ($logDurationNumericUpDown.Value -gt 7) {
        $durationWarningLabel.Text = "Note: >7 days requires Entra ID P1/P2 license."
    } else {
        $durationWarningLabel.Text = ""
    }
})
if ($logDurationNumericUpDown.Value -gt 7) { $durationWarningLabel.Text = "Note: >7 days requires Entra ID P1/P2 license." }

# Output Folder Label
$outputFolderLabel = New-Object System.Windows.Forms.Label
$outputFolderLabel.Location = New-Object System.Drawing.Point(20, 345)
$outputFolderLabel.Size = New-Object System.Drawing.Size(100, 20)
$outputFolderLabel.Text = "Output Folder:"
$outputFolderLabel.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$mainForm.Controls.Add($outputFolderLabel)

# Output Folder TextBox
$outputFolderTextBox = New-Object System.Windows.Forms.TextBox
$outputFolderTextBox.Location = New-Object System.Drawing.Point(120, 345)
$outputFolderTextBox.Size = New-Object System.Drawing.Size(345, 25)
$outputFolderTextBox.ReadOnly = $true
$outputFolderTextBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$outputFolderTextBox.add_TextChanged({
    $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '')
})
$mainForm.Controls.Add($outputFolderTextBox)

# Browse Button (for Output Folder)
$browseFolderButton = New-Object System.Windows.Forms.Button
$browseFolderButton.Location = New-Object System.Drawing.Point(475, 343)
$browseFolderButton.Size = New-Object System.Drawing.Size(90, 27)
$browseFolderButton.Text = "Browse..."
$browseFolderButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$browseFolderButton.add_Click({
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "Select the folder to save the log files"
    $folderBrowserDialog.ShowNewFolderButton = $true
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputFolderTextBox.Text = $folderBrowserDialog.SelectedPath
        $statusLabel.Text = "Output folder selected: $($outputFolderTextBox.Text)"
        # Enable Get Logs button if users are also selected
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '')
    }
})
$mainForm.Controls.Add($browseFolderButton)

# Get Logs Button
$getLogsButton = New-Object System.Windows.Forms.Button
$getLogsButton.Location = New-Object System.Drawing.Point(20, 390)
$getLogsButton.Size = New-Object System.Drawing.Size(545, 40)
$getLogsButton.Text = "Get Sign-in Logs for Selected Users"
$getLogsButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$getLogsButton.Enabled = $false
$getLogsButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$getLogsButton.add_Click({
    param($sender, $e)

    $selectedUpns = $userCheckedListBox.CheckedItems | ForEach-Object { $_ }
    $days = $logDurationNumericUpDown.Value
    $outputFolder = $outputFolderTextBox.Text

    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "No User Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
     if (-not ($outputFolder) -or (-not (Test-Path -Path $outputFolder -PathType Container))) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid output folder.", "Invalid Output Folder", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    if ($days -lt 1 -or $days -gt 30) {
         [System.Windows.Forms.MessageBox]::Show("Invalid duration selected. Please enter a value between 1 and 30.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $startDate = (Get-Date).AddDays(-$days).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    $statusLabel.Text = "Fetching logs for $($selectedUpns.Count) users (Last $days days)..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $getLogsButton.Enabled = $false
    # $loadUsersButton.Enabled = $false # Button removed
    $connectButton.Enabled = $false # Disable connect button during log fetch
    $logDurationNumericUpDown.Enabled = $false

    $allLogs = @()
    $errorOccurred = $false

    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseFileName = "EntraSignInLogs_$timestamp"
        $csvFilePath = Join-Path -Path $outputFolder -ChildPath "$($baseFileName).csv"

        Write-Host "Fetching logs starting from $startDate for users: $($selectedUpns -join ', ')"

        $totalUsers = $selectedUpns.Count
        $currentUserIndex = 0

        foreach ($userPrincipalName in $selectedUpns) {
            $currentUserIndex++
            $statusLabel.Text = "Processing user $currentUserIndex/$totalUsers ($userPrincipalName)..."
            $mainForm.Refresh()

            Write-Host "Processing user: $userPrincipalName"
            $userId = $null
            try {
                $statusLabel.Text = "Getting User ID for $userPrincipalName..."
                $mainForm.Refresh()
                $userObject = Get-MgUser -UserId $userPrincipalName -Property Id -ErrorAction Stop
                $userId = $userObject.Id
                Write-Host " Found User ID: $userId for $userPrincipalName"

                if (-not $userId) {
                    Write-Warning "Could not retrieve User ID for '$userPrincipalName'. Skipping user."
                    continue
                }

                $filterString = "userId eq '$userId' and createdDateTime ge $startDate"
                $statusLabel.Text = "Fetching logs for User ID $userId ($userPrincipalName)..."
                $mainForm.Refresh()

                $userLogs = Get-MgAuditLogSignIn -Filter $filterString -All -ErrorAction Stop `
                    -Property UserPrincipalName, CreatedDateTime, AppDisplayName, IpAddress, Location, DeviceDetail, Status, ConditionalAccessStatus, RiskDetail, RiskLevelAggregated, RiskLevelDuringSignIn, RiskState, RiskEventTypes_v2, IsInteractive, ResourceDisplayName

                if ($userLogs) {
                    Write-Host " Found $($userLogs.Count) log entries for $userPrincipalName (ID: $userId)." -ForegroundColor Cyan
                    $userLogs | ForEach-Object { if (-not $_.UserPrincipalName) { $_.UserPrincipalName = $userPrincipalName } }
                    $allLogs += $userLogs
                } else {
                    Write-Host " No sign-in logs found for $userPrincipalName (ID: $userId) in the specified period." -ForegroundColor Yellow
                }
            } catch {
                 Write-Warning "Could not retrieve logs for user '$userPrincipalName' (ID: $userId). Error: $($_.Exception.Message). Skipping user."
                 $errorOccurred = $true
            }
        }

        if ($allLogs.Count -eq 0) {
            if (-not $errorOccurred) {
                 $statusLabel.Text = "No sign-in logs found for the selected users in the specified period."
                 [System.Windows.Forms.MessageBox]::Show("No sign-in logs found for any of the selected users within the last $days days.", "No Logs Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                 $statusLabel.Text = "Finished processing users, but errors occurred and no logs were found/exported."
                 [System.Windows.Forms.MessageBox]::Show("Finished processing users, but errors occurred during log fetching (see console) and no logs were ultimately found/exported.", "Processing Complete with Errors", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        } else {
            $statusLabel.Text = "Exporting $($allLogs.Count) log entries to CSV..."
            Write-Host "Exporting $($allLogs.Count) total log entries to $csvFilePath"

            $exportData = $allLogs | Select-Object UserPrincipalName, CreatedDateTime, AppDisplayName, IpAddress, `
                @{Name='City';Expression={$_.Location.City}}, `
                @{Name='State';Expression={$_.Location.State}}, `
                @{Name='Country';Expression={$_.Location.CountryOrRegion}}, `
                @{Name='DeviceOperatingSystem';Expression={$_.DeviceDetail.OperatingSystem}}, `
                @{Name='DeviceBrowser';Expression={$_.DeviceDetail.Browser}}, `
                @{Name='DeviceIsCompliant';Expression={$_.DeviceDetail.IsCompliant}}, `
                @{Name='DeviceIsManaged';Expression={$_.DeviceDetail.IsManaged}}, `
                @{Name='DeviceTrustType';Expression={$_.DeviceDetail.TrustType}}, `
                @{Name='StatusErrorCode';Expression={$_.Status.ErrorCode}}, `
                @{Name='StatusFailureReason';Expression={$_.Status.FailureReason}}, `
                @{Name='StatusAdditionalDetails';Expression={$_.Status.AdditionalDetails}}, `
                ConditionalAccessStatus, RiskDetail, RiskLevelAggregated, RiskLevelDuringSignIn, RiskState, RiskEventTypes_v2, IsInteractive, ResourceDisplayName

            try {
                 $exportData | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                 $statusLabel.Text = "Successfully exported $($allLogs.Count) logs to $csvFilePath"
                 [System.Windows.Forms.MessageBox]::Show("Successfully exported $($allLogs.Count) sign-in log entries to CSV:`n$csvFilePath", "CSV Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                 Write-Error "Failed to export data to CSV '$csvFilePath'. Error: $($_.Exception.Message)"
                 $statusLabel.Text = "Error exporting data to CSV. Check console."
                 [System.Windows.Forms.MessageBox]::Show("Failed to export data to CSV. Please check file permissions and console for errors.`n`nError: $($_.Exception.Message)", "CSV Export Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                 $errorOccurred = $true
            }
        }
    } catch {
        $statusLabel.Text = "An unexpected error occurred. Check console."
        Write-Error "An unexpected error occurred: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show("An unexpected error occurred during the process. Check the PowerShell console for details.`n`nError: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $errorOccurred = $true
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '')
        # $loadUsersButton.Enabled = $true # Button removed
        $connectButton.Enabled = $true   # Re-enable connect button
        $logDurationNumericUpDown.Enabled = $true
        if ($errorOccurred) {
             $statusLabel.Text = "Operation finished with errors. Check console/messages."
        } elseif ($allLogs.Count -gt 0) {
             # Status label already shows CSV success message
        } else {
             # Status label already shows 'No logs found' or error message
        }
    }
})
$mainForm.Controls.Add($getLogsButton)

# --- Show Form ---
$mainForm.Add_Shown({$mainForm.Activate()}) # Bring form to front
[void]$mainForm.ShowDialog()

# --- Script End ---
Write-Host "Script finished."
# Optional: Disconnect Graph session on exit
# Disconnect-MgGraph -ErrorAction SilentlyContinue
