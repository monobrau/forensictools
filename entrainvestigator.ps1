<#
.SYNOPSIS
A PowerShell script with a GUI to fetch Entra ID sign-in logs for selected users,
and export to CSV. Filters logs by User ID for better reliability.

.DESCRIPTION
This script provides a Windows Forms interface to:
- Connect to Microsoft Graph.
- Load Entra ID users.
- Select users for investigation.
- Select the duration (1-30 days) for sign-in log history, with license warnings.
- Select an output folder.
- Fetch sign-in logs for the selected users (using User ID filter) and duration.
- Export logs directly to CSV format. (XLSX conversion and formatting skipped).

.NOTES
Author: Gemini
Date: 2025-05-05
Version: 1.8 (Skipped XLSX conversion and formatting to avoid Excel COM errors)
Requires: PowerShell 5.1+, Microsoft Graph SDK (Users, Reports). Excel is NOT required for this version.
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
# $highlightColorIndex = 6 # No longer needed

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
        # Add message to GUI status bar as well
        if ($statusLabel) {
            $statusLabel.Text = "Error installing modules. See console."
        }
    }
}

# --- Function ConvertTo-Xlsx is no longer called, but kept for reference ---
# Function ConvertTo-Xlsx {
#     param(
#         [Parameter(Mandatory=$true)]
#         [string]$CsvPath,
#         [Parameter(Mandatory=$true)]
#         [string]$XlsxPath,
#         [Parameter(Mandatory=$false)]
#         [int]$HighlightColor = 6, # Default Yellow
#         [Parameter(Mandatory=$false)]
#         [string]$CountryColumnHeader = "Country",
#         [Parameter(Mandatory=$false)]
#         [string]$CountryToExclude = "United States"
#     )

#     $excel = $null
#     $workbook = $null
#     $worksheet = $null
#     $usedRange = $null
#     $columns = $null
#     $rows = $null
#     $headerRange = $null
#     $countryColumnObject = $null
#     $countryColumnIndex = $null

#     # Excel Constants
#     $xlFormulas = -4123 # LookIn constant
#     $xlWhole = 1       # LookAt constant

#     try {
#         $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
#         $excel.Visible = $false
#         $excel.DisplayAlerts = $false

#         Write-Host "Converting '$CsvPath' to '$XlsxPath'..."
#         $workbook = $excel.Workbooks.Open($CsvPath)
#         $workbook.SaveAs($XlsxPath, 51) # 51 = xlOpenXMLWorkbook
#         $workbook.Close($false)
#         Write-Host "Initial conversion successful. Now formatting..." -ForegroundColor Green

#         $workbook = $excel.Workbooks.Open($XlsxPath)
#         $worksheet = $workbook.Worksheets.Item(1)
#         $usedRange = $worksheet.UsedRange
#         $columns = $usedRange.Columns
#         $rows = $usedRange.Rows

#         Write-Host "Auto-fitting columns..."
#         $columns.AutoFit() | Out-Null

#         Write-Host "Searching for '$CountryColumnHeader' column..."
#         $headerRange = $worksheet.Range("1:1")
#         $countryColumnObject = $headerRange.Find($CountryColumnHeader, $null, $xlFormulas, $xlWhole)

#         if ($countryColumnObject) {
#             $countryColumnIndex = $countryColumnObject.Column
#             Write-Host "'$CountryColumnHeader' column found at index $countryColumnIndex. Highlighting non-'$CountryToExclude' rows..."
#             for ($i = 2; $i -le $rows.Count; $i++) {
#                 $cell = $null; $rowRange = $null
#                 try {
#                     $cell = $worksheet.Cells.Item($i, $countryColumnIndex)
#                     $countryValue = $cell.Value2
#                     if ($countryValue -and ($countryValue -as [string]).Trim() -ne '' -and -not ($countryValue -as [string]).Equals($CountryToExclude, [System.StringComparison]::OrdinalIgnoreCase)) {
#                         $rowRange = $worksheet.Rows.Item($i)
#                         $rowRange.Interior.ColorIndex = $HighlightColor
#                     }
#                 } finally {
#                     if ($cell) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null }
#                     if ($rowRange) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rowRange) | Out-Null }
#                 }
#             }
#              Write-Host "Highlighting complete." -ForegroundColor Green
#         } else {
#             Write-Warning "Could not find the '$CountryColumnHeader' column header. Skipping row highlighting."
#         }

#         Write-Host "Saving formatted XLSX file..."
#         $workbook.Save()
#         $workbook.Close()
#         Write-Host "XLSX formatting complete." -ForegroundColor Green

#     } catch {
#         Write-Error "Failed during Excel formatting or conversion. Error: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
#         if ($statusLabel) { $statusLabel.Text = "Error: Failed during XLSX conversion/formatting." }
#         return $false
#     } finally {
#         # Clean up COM objects
#         if ($countryColumnObject) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($countryColumnObject) | Out-Null }
#         if ($headerRange) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($headerRange) | Out-Null }
#         if ($columns) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($columns) | Out-Null }
#         if ($rows) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rows) | Out-Null }
#         if ($usedRange) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null }
#         if ($worksheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
#         if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
#         if ($excel) { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
#         [gc]::Collect(); [gc]::WaitForPendingFinalizers()
#     }
#     return $true
# }


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
    # Exit after attempting install, user needs to restart
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
$mainForm.Size = New-Object System.Drawing.Size(600, 580) # Increased height slightly for warning label
$mainForm.MinimumSize = New-Object System.Drawing.Size(550, 530)
$mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog # Or Sizable
$mainForm.MaximizeBox = $false
$mainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi # Improve scaling

# Status Strip and Label (Replaces StatusBar)
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Name = "statusLabel"
$statusLabel.Text = "Ready. Please connect to Microsoft Graph."
$statusStrip.Items.Add($statusLabel)
$mainForm.Controls.Add($statusStrip)


# Connect Button
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Location = New-Object System.Drawing.Point(20, 20)
$connectButton.Size = New-Object System.Drawing.Size(150, 30)
$connectButton.Text = "Connect to Graph"
$connectButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$connectButton.add_Click({
    param($sender, $e)
    $statusLabel.Text = "Connecting to Microsoft Graph..." # Use statusLabel
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Disconnect if already connected to ensure fresh login/scopes
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
        $context = Get-MgContext
        $statusLabel.Text = "Connected as $($context.Account) Tenant: $($context.TenantId)" # Use statusLabel
        $loadUsersButton.Enabled = $true
        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
    } catch {
        $statusLabel.Text = "Connection failed. Check console for errors." # Use statusLabel
        Write-Error "Microsoft Graph connection failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to Microsoft Graph. Ensure you have internet connectivity and the necessary permissions. `n`nError: $($_.Exception.Message)", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $loadUsersButton.Enabled = $false
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
$mainForm.Controls.Add($connectButton)

# Load Users Button
$loadUsersButton = New-Object System.Windows.Forms.Button
$loadUsersButton.Location = New-Object System.Drawing.Point(180, 20)
$loadUsersButton.Size = New-Object System.Drawing.Size(150, 30)
$loadUsersButton.Text = "Load Entra Users"
$loadUsersButton.Enabled = $false # Disabled until connected
$loadUsersButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$loadUsersButton.add_Click({
    param($sender, $e)
    $statusLabel.Text = "Loading users... This may take a while for large tenants." # Use statusLabel
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $userCheckedListBox.Items.Clear()
    $getLogsButton.Enabled = $false # Disable until users are loaded and selected
    try {
        # Fetch users - Use -All for complete list, but be mindful of performance
        # Added -ConsistencyLevel eventual to handle potential advanced query issues on some tenants
        $users = Get-MgUser -All -ErrorAction Stop -Select UserPrincipalName, Id, DisplayName -ConsistencyLevel eventual | Sort-Object UserPrincipalName
        if ($users) {
            foreach ($user in $users) {
                # Add UserPrincipalName to the listbox
                $userCheckedListBox.Items.Add($user.UserPrincipalName, $false) # Add unchecked
            }
            $statusLabel.Text = "Loaded $($users.Count) users. Select users to investigate." # Use statusLabel
            Write-Host "Loaded $($users.Count) users." -ForegroundColor Green
        } else {
            $statusLabel.Text = "No users found or error loading users." # Use statusLabel
             [System.Windows.Forms.MessageBox]::Show("No users found in the tenant or an error occurred.", "No Users", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } catch {
        $statusLabel.Text = "Error loading users. Check console." # Use statusLabel
        Write-Error "Error fetching users from Microsoft Graph: $($_.Exception.Message)"
         [System.Windows.Forms.MessageBox]::Show("Error fetching users from Microsoft Graph.`n`nError: $($_.Exception.Message)", "Error Loading Users", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
$mainForm.Controls.Add($loadUsersButton)

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
    # Enable Get Logs button only if at least one user is checked
    # Need a slight delay because the CheckedItems count doesn't update immediately
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
$logDurationNumericUpDown.Maximum = 30 # Set max duration
$logDurationNumericUpDown.Value = 7    # Default to 7 days
$logDurationNumericUpDown.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$mainForm.Controls.Add($logDurationNumericUpDown)

# Warning Label for Duration
$durationWarningLabel = New-Object System.Windows.Forms.Label
$durationWarningLabel.Location = New-Object System.Drawing.Point(240, 308) # Positioned next to NumericUpDown
$durationWarningLabel.Size = New-Object System.Drawing.Size(325, 20)
$durationWarningLabel.Text = "" # Initially empty
$durationWarningLabel.ForeColor = [System.Drawing.Color]::OrangeRed
$durationWarningLabel.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$mainForm.Controls.Add($durationWarningLabel)

# Event handler for NumericUpDown value change
$logDurationNumericUpDown.add_ValueChanged({
    param($sender, $e)
    if ($logDurationNumericUpDown.Value -gt 7) {
        $durationWarningLabel.Text = "Note: >7 days requires Entra ID P1/P2 license."
    } else {
        $durationWarningLabel.Text = "" # Clear warning if 7 or less
    }
})

# Trigger initial check for default value
if ($logDurationNumericUpDown.Value -gt 7) {
    $durationWarningLabel.Text = "Note: >7 days requires Entra ID P1/P2 license."
}

# Output Folder Label
$outputFolderLabel = New-Object System.Windows.Forms.Label
$outputFolderLabel.Location = New-Object System.Drawing.Point(20, 345) # Adjusted Y position
$outputFolderLabel.Size = New-Object System.Drawing.Size(100, 20)
$outputFolderLabel.Text = "Output Folder:"
$outputFolderLabel.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$mainForm.Controls.Add($outputFolderLabel)

# Output Folder TextBox
$outputFolderTextBox = New-Object System.Windows.Forms.TextBox
$outputFolderTextBox.Location = New-Object System.Drawing.Point(120, 345) # Adjusted Y position
$outputFolderTextBox.Size = New-Object System.Drawing.Size(345, 25)
$outputFolderTextBox.ReadOnly = $true # Make it read-only, set via Browse button
$outputFolderTextBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$outputFolderTextBox.add_TextChanged({
    # Enable Get Logs button only if a folder is selected and at least one user is checked
    $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '')
})
$mainForm.Controls.Add($outputFolderTextBox)

# Browse Button (for Output Folder)
$browseFolderButton = New-Object System.Windows.Forms.Button
$browseFolderButton.Location = New-Object System.Drawing.Point(475, 343) # Adjusted Y position
$browseFolderButton.Size = New-Object System.Drawing.Size(90, 27)
$browseFolderButton.Text = "Browse..."
$browseFolderButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$browseFolderButton.add_Click({
    param($sender, $e)
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "Select the folder to save the log files"
    $folderBrowserDialog.ShowNewFolderButton = $true
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputFolderTextBox.Text = $folderBrowserDialog.SelectedPath
        $statusLabel.Text = "Output folder selected: $($outputFolderTextBox.Text)" # Use statusLabel
    }
})
$mainForm.Controls.Add($browseFolderButton)

# Get Logs Button
$getLogsButton = New-Object System.Windows.Forms.Button
$getLogsButton.Location = New-Object System.Drawing.Point(20, 390) # Adjusted Y position
$getLogsButton.Size = New-Object System.Drawing.Size(545, 40)
$getLogsButton.Text = "Get Sign-in Logs for Selected Users"
$getLogsButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$getLogsButton.Enabled = $false # Disabled initially
$getLogsButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$getLogsButton.add_Click({
    param($sender, $e)

    $selectedUpns = $userCheckedListBox.CheckedItems | ForEach-Object { $_ } # Get the checked UPNs
    $days = $logDurationNumericUpDown.Value # Get value directly from NumericUpDown
    $outputFolder = $outputFolderTextBox.Text

    if ($selectedUpns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user.", "No User Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
     if (-not ($outputFolder) -or (-not (Test-Path -Path $outputFolder -PathType Container))) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid output folder.", "Invalid Output Folder", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Validate days (already constrained by NumericUpDown, but good practice)
    if ($days -lt 1 -or $days -gt 30) {
         [System.Windows.Forms.MessageBox]::Show("Invalid duration selected. Please enter a value between 1 and 30.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $startDate = (Get-Date).AddDays(-$days).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    $statusLabel.Text = "Fetching logs for $($selectedUpns.Count) users (Last $days days)..." # Use statusLabel
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $getLogsButton.Enabled = $false # Disable during processing
    $loadUsersButton.Enabled = $false
    $connectButton.Enabled = $false
    $logDurationNumericUpDown.Enabled = $false # Disable duration change during fetch

    $allLogs = @()
    $errorOccurred = $false

    try {
        # Create a timestamp for unique filenames
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseFileName = "EntraSignInLogs_$timestamp"
        $csvFilePath = Join-Path -Path $outputFolder -ChildPath "$($baseFileName).csv"
        # $xlsxFilePath = Join-Path -Path $outputFolder -ChildPath "$($baseFileName).xlsx" # No longer needed

        Write-Host "Fetching logs starting from $startDate for users: $($selectedUpns -join ', ')"

        $totalUsers = $selectedUpns.Count
        $currentUserIndex = 0

        foreach ($userPrincipalName in $selectedUpns) {
            $currentUserIndex++
            $statusLabel.Text = "Fetching logs for user $currentUserIndex/$totalUsers ($userPrincipalName)..." # Use statusLabel
            $mainForm.Refresh() # Update UI

            Write-Host "Processing user: $userPrincipalName"
            $userId = $null
            try {
                # --- Get User ID from UPN ---
                $statusLabel.Text = "Getting User ID for $userPrincipalName..."
                $mainForm.Refresh()
                $userObject = Get-MgUser -UserId $userPrincipalName -Property Id -ErrorAction Stop
                $userId = $userObject.Id
                Write-Host " Found User ID: $userId for $userPrincipalName"

                if (-not $userId) {
                    Write-Warning "Could not retrieve User ID for '$userPrincipalName'. Skipping user."
                    continue # Move to the next user in the loop
                }

                # --- Construct Filter using User ID ---
                $filterString = "userId eq '$userId' and createdDateTime ge $startDate"
                $statusLabel.Text = "Fetching logs for User ID $userId ($userPrincipalName)..." # Update status
                $mainForm.Refresh()

                # Use -All to handle pagination automatically
                # Select specific properties including location details
                $userLogs = Get-MgAuditLogSignIn -Filter $filterString -All -ErrorAction Stop `
                    -Property UserPrincipalName, CreatedDateTime, AppDisplayName, IpAddress, Location, DeviceDetail, Status, ConditionalAccessStatus, RiskDetail, RiskLevelAggregated, RiskLevelDuringSignIn, RiskState, RiskEventTypes_v2, IsInteractive, ResourceDisplayName

                if ($userLogs) {
                    Write-Host " Found $($userLogs.Count) log entries for $userPrincipalName (ID: $userId)." -ForegroundColor Cyan
                    # Add UPN back to the log object if it's missing (sometimes happens when filtering by ID)
                    $userLogs | ForEach-Object { if (-not $_.UserPrincipalName) { $_.UserPrincipalName = $userPrincipalName } }
                    $allLogs += $userLogs # Add logs to the main collection
                } else {
                    Write-Host " No sign-in logs found for $userPrincipalName (ID: $userId) in the specified period." -ForegroundColor Yellow
                }

            } catch {
                 Write-Warning "Could not retrieve logs for user '$userPrincipalName' (ID: $userId). Error: $($_.Exception.Message). Skipping user."
                 # Optionally add this info to a separate error log
                 $errorOccurred = $true # Mark error occurred if log fetching fails for a user
            }
        } # End foreach user

        if ($allLogs.Count -eq 0) {
            # Only show this message if no errors occurred during fetch attempts for any user
            if (-not $errorOccurred) {
                 $statusLabel.Text = "No sign-in logs found for the selected users in the specified period." # Use statusLabel
                 [System.Windows.Forms.MessageBox]::Show("No sign-in logs found for any of the selected users within the last $days days.", "No Logs Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                 $statusLabel.Text = "Finished processing users, but errors occurred and no logs were found/exported."
                 [System.Windows.Forms.MessageBox]::Show("Finished processing users, but errors occurred during log fetching (see console) and no logs were ultimately found/exported.", "Processing Complete with Errors", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }

        } else {
            $statusLabel.Text = "Exporting $($allLogs.Count) log entries to CSV..." # Use statusLabel
            Write-Host "Exporting $($allLogs.Count) total log entries to $csvFilePath"

            # Select relevant properties for export, including calculated location properties
            $exportData = $allLogs | Select-Object UserPrincipalName, CreatedDateTime, AppDisplayName, IpAddress, `
                @{Name='City';Expression={$_.Location.City}}, `
                @{Name='State';Expression={$_.Location.State}}, `
                @{Name='Country';Expression={$_.Location.CountryOrRegion}}, `
                DeviceDetail, Status, ConditionalAccessStatus, RiskDetail, RiskLevelAggregated, RiskLevelDuringSignIn, RiskState, RiskEventTypes_v2, IsInteractive, ResourceDisplayName

            # --- Export directly to CSV ---
            try {
                 $exportData | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                 $statusLabel.Text = "Successfully exported $($allLogs.Count) logs to $csvFilePath" # Use statusLabel
                 [System.Windows.Forms.MessageBox]::Show("Successfully exported $($allLogs.Count) sign-in log entries to CSV:`n$csvFilePath", "CSV Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                 # Optional: Open the folder
                 # Invoke-Item $outputFolder
            } catch {
                 Write-Error "Failed to export data to CSV '$csvFilePath'. Error: $($_.Exception.Message)"
                 $statusLabel.Text = "Error exporting data to CSV. Check console."
                 [System.Windows.Forms.MessageBox]::Show("Failed to export data to CSV. Please check file permissions and console for errors.`n`nError: $($_.Exception.Message)", "CSV Export Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                 $errorOccurred = $true
            }

            # --- Skip XLSX Conversion ---
            # Write-Host "Skipping XLSX conversion and formatting."
            # if (ConvertTo-Xlsx -CsvPath $csvFilePath -XlsxPath $xlsxFilePath -CountryColumnHeader "Country" -CountryToExclude "United States") {
            #      $statusLabel.Text = "Successfully exported and formatted $($allLogs.Count) logs to $xlsxFilePath" # Use statusLabel
            #      [System.Windows.Forms.MessageBox]::Show("Successfully exported $($allLogs.Count) sign-in log entries and formatted the file:`n$xlsxFilePath", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            # } else {
            #      $statusLabel.Text = "Exported to CSV ($csvFilePath), but XLSX conversion/formatting failed." # Use statusLabel
            #      [System.Windows.Forms.MessageBox]::Show("Log data exported successfully to CSV:`n$csvFilePath`n`nHowever, conversion to XLSX or formatting failed. Please ensure Excel is installed or check console for errors.", "CSV Exported, XLSX Failed/Formatting Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #      $errorOccurred = $true
            # }
        }

    } catch {
        $statusLabel.Text = "An unexpected error occurred. Check console." # Use statusLabel
        Write-Error "An unexpected error occurred: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show("An unexpected error occurred during the process. Check the PowerShell console for details.`n`nError: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $errorOccurred = $true
    } finally {
        # Re-enable controls
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '') # Re-evaluate based on state
        $loadUsersButton.Enabled = $true # Assuming still connected
        $connectButton.Enabled = $true
        $logDurationNumericUpDown.Enabled = $true # Re-enable duration change
        if ($errorOccurred) {
             $statusLabel.Text = "Operation finished with errors. Check console/messages." # Use statusLabel
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
