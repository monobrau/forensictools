<#
.SYNOPSIS
A PowerShell script with a GUI to fetch Entra ID sign-in logs for selected users,
and export to a formatted XLSX file. Filters logs by User ID for better reliability.
Users are loaded automatically upon successful connection. Includes a disconnect button.

.DESCRIPTION
This script provides a Windows Forms interface to:
- Connect to Microsoft Graph (and automatically load Entra ID users and attempt to determine tenant domain).
- Disconnect the Microsoft Graph session.
- Select users for investigation.
- Select the duration (1-30 days) for sign-in log history, with license warnings.
- Select an output folder.
- Fetch sign-in logs for the selected users (using User ID filter) and duration.
- Export logs directly to CSV format, then convert to XLSX and apply formatting:
    - Auto-fit columns.
    - Bold header row.
    - Highlight rows yellow where Country *is* 'United States'. (Highlighting logic reversed)
- The XLSX filename will include the tenant's primary domain name.

.NOTES
Author: Gemini
Date: 2025-05-06
Version: 3.3 (Reversed highlighting logic to highlight US rows instead of non-US rows)
Requires:
    - PowerShell 5.1+
    - Microsoft Graph SDK (Users, Reports, Identity.DirectoryManagement)
    - *** Microsoft Excel Installed *** (for XLSX conversion and formatting)
Permissions: Requires delegated User.Read.All, AuditLog.Read.All, and Organization.Read.All permissions in Entra ID.

.LINK
Install Modules: Install-Module Microsoft.Graph.Users, Microsoft.Graph.Reports, Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force

.EXAMPLE
.\EntraID_Forensic_Log_Fetcher.ps1
#>

#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Reports, Microsoft.Graph.Identity.DirectoryManagement

# --- Configuration ---
$requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Reports", "Microsoft.Graph.Identity.DirectoryManagement")
$requiredScopes = @("User.Read.All", "AuditLog.Read.All", "Organization.Read.All")
$highlightColorIndexYellow = 6 # Excel ColorIndex for Yellow

# Script-level variable to store tenant domain
$script:tenantDomainNameForFile = $null

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

Function ConvertTo-XlsxAndFormat {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CsvPath,
        [Parameter(Mandatory=$true)]
        [string]$XlsxPath,
        [Parameter(Mandatory=$false)]
        [int]$HighlightColor = $highlightColorIndexYellow,
        [Parameter(Mandatory=$false)]
        [string]$CountryColumnHeader = "Country",
        # Changed parameter name slightly for clarity
        [Parameter(Mandatory=$false)]
        [string]$CountryToHighlight = "United States" 
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null
    $columns = $null
    $rows = $null
    $headerRange = $null
    $countryColumnObject = $null
    $countryColumnIndex = $null

    # Excel Constants
    $xlOpenXMLWorkbook = 51 # FileFormat for .xlsx
    $xlFormulas = -4123    # LookIn constant for Find
    $xlWhole = 1           # LookAt constant for Find
    $xlByRows = 1          # SearchOrder constant
    $xlNext = 1            # SearchDirection constant
    $missing = [System.Reflection.Missing]::Value # For optional COM parameters

    # Check if Excel is installed by trying to create the object
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    } catch {
        Write-Error "Failed to create Excel COM object. Ensure Microsoft Excel is installed and accessible. Error: $($_.Exception.Message)"
        if ($statusLabel) { $statusLabel.Text = "Error: Excel not found or accessible." }
        return $false # Indicate failure
    }

    try {
        $excel.Visible = $false         # Keep Excel hidden
        $excel.DisplayAlerts = $false   # Don't show Excel prompts

        Write-Host "Converting '$CsvPath' to '$XlsxPath'..."
        # Open CSV and save as XLSX
        $workbook = $excel.Workbooks.Open($CsvPath)
        $workbook.SaveAs($XlsxPath, $xlOpenXMLWorkbook)
        $workbook.Close($false) # Close the CSV representation
        Write-Host "Initial conversion successful. Now formatting..." -ForegroundColor Green

        # Re-open the XLSX for formatting
        $workbook = $excel.Workbooks.Open($XlsxPath)
        $worksheet = $workbook.Worksheets.Item(1) # Get the first sheet
        $usedRange = $worksheet.UsedRange
        $columns = $usedRange.Columns
        $rows = $usedRange.Rows

        if ($usedRange.Rows.Count -gt 0) {
            # --- AutoFit Columns ---
            Write-Host " - Autofitting columns..."
            $columns.AutoFit() | Out-Null

             # --- Bold Header Row ---
            Write-Host " - Bolding header row..."
            $headerRange = $worksheet.Rows.Item(1)
            $headerRange.Font.Bold = $true

            # --- Highlight Rows where Country IS $CountryToHighlight (if more than header exists) ---
            if ($usedRange.Rows.Count -gt 1) {
                Write-Host " - Searching for '$CountryColumnHeader' column..."
                # Using [System.Reflection.Missing]::Value for optional parameters
                $countryColumnObject = $headerRange.Find(
                    $CountryColumnHeader, # What
                    $missing,             # After
                    $xlFormulas,          # LookIn
                    $xlWhole,             # LookAt
                    $xlByRows,            # SearchOrder
                    $xlNext,              # SearchDirection
                    $false,               # MatchCase
                    $missing,             # MatchByte
                    $missing              # SearchFormat
                )

                if ($countryColumnObject) {
                    $countryColumnIndex = $countryColumnObject.Column
                    Write-Host "   - '$CountryColumnHeader' column found at index $countryColumnIndex. Highlighting rows where country is '$CountryToHighlight'..."
                    # Iterate through data rows
                    for ($i = 2; $i -le $rows.Count; $i++) {
                        $cell = $null; $rowRange = $null
                        try {
                            $cell = $worksheet.Cells.Item($i, $countryColumnIndex)
                            $countryValue = $cell.Value2
                            # --- MODIFIED CONDITION: Highlight if country IS the specified value ---
                            if ($countryValue -and ($countryValue -as [string]).Trim() -ne '' -and ($countryValue -as [string]).Equals($CountryToHighlight, [System.StringComparison]::OrdinalIgnoreCase)) {
                                # Highlight the entire row
                                $rowRange = $worksheet.Rows.Item($i)
                                $rowRange.Interior.ColorIndex = $HighlightColor
                            }
                        } finally {
                            # Release cell and row COM objects within the loop
                            if ($cell) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null }
                            if ($rowRange) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rowRange) | Out-Null }
                        }
                    }
                    Write-Host "   - Highlighting complete." -ForegroundColor Green
                } else {
                    Write-Warning "   - Could not find the '$CountryColumnHeader' column header. Skipping row highlighting."
                }
             } else {
                 Write-Host " - Only header row found, skipping row highlighting."
             }
        } else {
             Write-Host " - Worksheet appears empty, skipping formatting."
        }


        # Save the changes to the XLSX file
        Write-Host "Saving formatted XLSX file..."
        $workbook.Save()
        $workbook.Close()
        Write-Host "XLSX formatting complete." -ForegroundColor Green

    } catch {
        Write-Error "Failed during Excel formatting or conversion. Error: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
        if ($statusLabel) { $statusLabel.Text = "Error: Failed during XLSX conversion/formatting." }
        # Attempt to close workbook even if error occurred during formatting
        try { if ($workbook -ne $null) { $workbook.Close($false) } } catch {}
        return $false # Indicate failure
    } finally {
        # Clean up COM objects meticulously
        if ($countryColumnObject) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($countryColumnObject) | Out-Null }
        if ($headerRange) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($headerRange) | Out-Null }
        if ($columns) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($columns) | Out-Null }
        if ($rows) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rows) | Out-Null }
        if ($usedRange) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null }
        if ($worksheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
        if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } # Already closed in try/catch, just release
        if ($excel) { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
        # Force garbage collection
        [gc]::Collect(); [gc]::WaitForPendingFinalizers()
        Write-Host "COM cleanup finished."
    }
    return $true # Indicate success
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
Import-Module Microsoft.Graph.Identity.DirectoryManagement # For Get-MgOrganization

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
$connectButton.Size = New-Object System.Drawing.Size(160, 30)
$connectButton.Text = "Connect & Load Users"
$connectButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$connectButton.add_Click({
    param($sender, $e)
    $statusLabel.Text = "Connecting to Microsoft Graph..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $userCheckedListBox.Items.Clear()
    $getLogsButton.Enabled = $false
    $script:tenantDomainNameForFile = $null # Reset tenant domain
    $localTenantId = $null # To store tenant ID for fallback

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
        $context = Get-MgContext
        $localTenantId = $context.TenantId # Store for later fallback if needed
        $statusLabel.Text = "Connected as $($context.Account). Tenant: $localTenantId. Fetching org details..."
        $mainForm.Refresh()
        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green

        # --- Attempt 1: Fetch Organization Details to get Tenant Domain ---
        try {
            Write-Host "Attempt 1: Fetching organization details..."
            $orgDetails = Get-MgOrganization -Property Id, DisplayName, VerifiedDomains -ErrorAction SilentlyContinue # Continue if this fails
            
            if ($orgDetails -and $orgDetails.Count -gt 0) {
                $currentOrg = $orgDetails[0]
                Write-Host "Organization DisplayName: $($currentOrg.DisplayName)"

                if ($currentOrg.VerifiedDomains) {
                    $defaultDomain = $currentOrg.VerifiedDomains | Where-Object {$_.IsDefault -eq $true} | Select-Object -ExpandProperty Name -First 1
                    if ($defaultDomain) {
                        $script:tenantDomainNameForFile = $defaultDomain
                        Write-Host "Default tenant domain found via Get-MgOrganization: $($script:tenantDomainNameForFile)" -ForegroundColor Green
                    } else {
                        Write-Warning "No default domain found (IsDefault -eq `$true) via Get-MgOrganization."
                        $firstDomainName = $currentOrg.VerifiedDomains | Select-Object -ExpandProperty Name -First 1
                        if ($firstDomainName) {
                            $script:tenantDomainNameForFile = $firstDomainName
                            Write-Host "Using first verified tenant domain via Get-MgOrganization: $($script:tenantDomainNameForFile)" -ForegroundColor Yellow
                        } else {
                            Write-Warning "No verified domains found via Get-MgOrganization."
                        }
                    }
                } else {
                    Write-Warning "VerifiedDomains property is null or empty from Get-MgOrganization."
                }
            } else {
                Write-Warning "Could not retrieve organization details via Get-MgOrganization."
            }
        } catch {
            Write-Warning "Error during Get-MgOrganization: $($_.Exception.Message)."
        }
        
        $statusLabel.Text = "Org details processed. Loading users..."
        $mainForm.Refresh()

        # --- Load Users ---
        Write-Host "Loading users..."
        $users = Get-MgUser -All -ErrorAction Stop -Select UserPrincipalName, Id, DisplayName -ConsistencyLevel eventual | Sort-Object UserPrincipalName
        
        if ($users) {
            foreach ($user in $users) {
                $userCheckedListBox.Items.Add($user.UserPrincipalName, $false)
            }
            Write-Host "Loaded $($users.Count) users." -ForegroundColor Green

            # --- Attempt 2: Parse domain from first UPN if Get-MgOrganization failed ---
            if (-not $script:tenantDomainNameForFile -and $users.Count -gt 0) {
                Write-Host "Attempt 2: Parsing domain from first user's UPN..."
                $firstUserUpn = $users[0].UserPrincipalName
                if ($firstUserUpn -like "*@*") {
                    $domainFromUpn = $firstUserUpn.Split('@')[1]
                    if (-not [string]::IsNullOrWhiteSpace($domainFromUpn)) {
                        $script:tenantDomainNameForFile = $domainFromUpn
                        Write-Host "Tenant domain determined from UPN: $($script:tenantDomainNameForFile)" -ForegroundColor Green
                    } else {
                        Write-Warning "Could not parse a valid domain from UPN '$firstUserUpn'."
                    }
                } else {
                    Write-Warning "First user UPN '$firstUserUpn' does not contain '@'."
                }
            }
        } else {
            Write-Warning "No users loaded. Cannot parse domain from UPN."
        }

        # --- Attempt 3: Final Fallback to Tenant ID ---
        if (-not $script:tenantDomainNameForFile) {
            Write-Warning "All attempts to find domain name failed. Using Tenant ID for filename."
            $script:tenantDomainNameForFile = $localTenantId 
        }
        
        $statusLabel.Text = "Connected. Loaded $($users.Count) users. Tenant for filename: $($script:tenantDomainNameForFile)"
        $disconnectButton.Enabled = $true

    } catch {
        $statusLabel.Text = "Operation failed. Check console for errors."
        Write-Error "Microsoft Graph connection or user loading failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to Microsoft Graph or load users. Ensure you have internet connectivity and the necessary permissions. `n`nError: $($_.Exception.Message)", "Connection/Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $disconnectButton.Enabled = $false
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '')
    }
})
$mainForm.Controls.Add($connectButton)

# Disconnect Button
$disconnectButton = New-Object System.Windows.Forms.Button
$disconnectButton.Location = New-Object System.Drawing.Point(190, 20)
$disconnectButton.Size = New-Object System.Drawing.Size(160, 30)
$disconnectButton.Text = "Disconnect from Graph"
$disconnectButton.Enabled = $false
$disconnectButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left)
$disconnectButton.add_Click({
    param($sender, $e)
    $statusLabel.Text = "Disconnecting from Microsoft Graph..."
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
        $userCheckedListBox.Items.Clear()
        $getLogsButton.Enabled = $false
        $statusLabel.Text = "Disconnected. Ready to connect."
        $disconnectButton.Enabled = $false
        $connectButton.Enabled = $true
        $script:tenantDomainNameForFile = $null # Clear stored tenant domain
    } catch {
        $statusLabel.Text = "Error during disconnection. Check console."
        Write-Error "Error disconnecting from Microsoft Graph: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("An error occurred while trying to disconnect.`n`nError: $($_.Exception.Message)", "Disconnection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $disconnectButton.Enabled = $true
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
$mainForm.Controls.Add($disconnectButton)


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
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '' -and $disconnectButton.Enabled)
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
    $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '' -and $disconnectButton.Enabled)
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
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '' -and $disconnectButton.Enabled)
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
    $connectButton.Enabled = $false 
    $disconnectButton.Enabled = $false
    $logDurationNumericUpDown.Enabled = $false

    $allLogs = @()
    $errorOccurred = $false
    $csvExported = $false # Flag to track if CSV was created

    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        # --- Use Tenant Domain for Filename ---
        $safeTenantDomain = "UnknownTenant"
        if ($script:tenantDomainNameForFile) {
            if ($script:tenantDomainNameForFile -match "^\w{8}-\w{4}-\w{4}-\w{4}-\w{12}$") {
                $safeTenantDomain = $script:tenantDomainNameForFile
            } else {
                $safeTenantDomain = $script:tenantDomainNameForFile -replace "[^a-zA-Z0-9_.-]", "" -replace "\.", "_" 
            }
        }
        $baseFileName = "EntraSignInLogs_$($safeTenantDomain)_$timestamp"
        $csvFilePath = Join-Path -Path $outputFolder -ChildPath "$($baseFileName).csv"
        $xlsxFilePath = Join-Path -Path $outputFolder -ChildPath "$($baseFileName).xlsx" # Define XLSX path

        Write-Host "Fetching logs starting from $startDate for users: $($selectedUpns -join ', ')"
        Write-Host "Output file will be: $xlsxFilePath (via $csvFilePath)"

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
                 Write-Host "Successfully exported data to temporary CSV: $csvFilePath" -ForegroundColor Green
                 $csvExported = $true
            } catch {
                 Write-Error "Failed to export data to CSV '$csvFilePath'. Error: $($_.Exception.Message)"
                 $statusLabel.Text = "Error exporting data to CSV. Check console."
                 [System.Windows.Forms.MessageBox]::Show("Failed to export data to CSV. Please check file permissions and console for errors.`n`nError: $($_.Exception.Message)", "CSV Export Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                 $errorOccurred = $true
            }

            # --- Convert and Format if CSV export was successful ---
            if ($csvExported) {
                $statusLabel.Text = "Converting CSV to XLSX and formatting..."
                Write-Host "Attempting conversion to XLSX format and applying formatting..."
                # Pass the correct Country value to highlight
                if (ConvertTo-XlsxAndFormat -CsvPath $csvFilePath -XlsxPath $xlsxFilePath -CountryColumnHeader "Country" -CountryToHighlight "United States") {
                    $statusLabel.Text = "Successfully exported and formatted $($allLogs.Count) logs to $xlsxFilePath"
                    [System.Windows.Forms.MessageBox]::Show("Successfully exported $($allLogs.Count) sign-in log entries and formatted the file:`n$xlsxFilePath", "XLSX Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    # Optional: Clean up temporary CSV
                    try { Remove-Item -Path $csvFilePath -Force -ErrorAction SilentlyContinue } catch {}
                } else {
                    # Conversion/Formatting failed, message already shown by ConvertTo-XlsxAndFormat
                    $statusLabel.Text = "Exported to CSV ($csvFilePath), but XLSX conversion/formatting failed."
                    [System.Windows.Forms.MessageBox]::Show("Log data exported successfully to CSV:`n$csvFilePath`n`nHowever, conversion to XLSX or formatting failed. Please ensure Excel is installed or check console for errors.", "CSV Exported, XLSX Failed/Formatting Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    $errorOccurred = $true
                }
            }
        }
    } catch {
        $statusLabel.Text = "An unexpected error occurred. Check console."
        Write-Error "An unexpected error occurred: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
        [System.Windows.Forms.MessageBox]::Show("An unexpected error occurred during the process. Check the PowerShell console for details.`n`nError: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $errorOccurred = $true
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        $getLogsButton.Enabled = ($userCheckedListBox.CheckedItems.Count -gt 0 -and $outputFolderTextBox.Text -ne '' -and $disconnectButton.Enabled)
        $connectButton.Enabled = $true   
        $disconnectButton.Enabled = $true 
        $logDurationNumericUpDown.Enabled = $true
        if ($errorOccurred) {
             $statusLabel.Text = "Operation finished with errors. Check console/messages."
        } elseif ($allLogs.Count -gt 0) {
             # Status label already shows success/failure message from export/convert step
        } else {
             # Status label already shows 'No logs found' or error message
        }
        # Clean up temporary CSV if it still exists and XLSX failed
        if ($csvExported -and $errorOccurred -and (Test-Path $csvFilePath)) {
             Write-Host "Cleaning up temporary CSV file ($csvFilePath) as XLSX conversion failed."
             # Keep the CSV in case of failure for manual processing
             # try { Remove-Item -Path $csvFilePath -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
})
$mainForm.Controls.Add($getLogsButton)

# --- Show Form ---
$mainForm.Add_Shown({$mainForm.Activate()}) # Corrected: Single underscore
[void]$mainForm.ShowDialog()

# --- Script End ---
Write-Host "Script finished."
# Optional: Disconnect Graph session on exit
# Disconnect-MgGraph -ErrorAction SilentlyContinue
