#Requires -Modules Microsoft.Graph, ImportExcel
<#
.SYNOPSIS
Provides a GUI for basic forensic analysis of Entra ID user sign-in logs.
.DESCRIPTION
Connects to Microsoft Graph, lists Entra ID users, allows selection,
fetches sign-in logs for selected users, exports data to a temporary CSV,
imports from CSV, then exports to XLSX with basic formatting.
Conditional Formatting and Column Hiding features disabled due to environment compatibility.
.NOTES
Author: AI Assistant (Gemini)
Date:   2025-05-05
Version: 1.4 (Removes ConditionalFormatting and HideColumn features)

Requires necessary Entra ID permissions (e.g., AuditLog.Read.All, User.Read.All).
Log retention is subject to your Entra ID license and configuration (P1/P2 needed for >7 days).
This script attempts to fetch logs up to 90 days back.
Requires ImportExcel module.
.LINK
https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.reports/get-mgauditlogsignin
https://github.com/dfinke/ImportExcel
#>

# --- Configuration ---
$global:requiredModules = @('Microsoft.Graph', 'ImportExcel')
# NOTE: Column hiding feature is disabled below, but list kept for reference
$global:columnsToHide = @(
    'Id', 'UserId', 'AppId', 'ConditionalAccessPolicies', 'DeviceDetail',
    'ConditionalAccessStatus', 'CorrelationId', 'RiskDetail', 'RiskLevelAggregated',
    'RiskLevelDuringSignIn', 'RiskState', 'RiskEventTypes', 'ResourceDisplayName',
    'ResourceTenantId', 'HomeTenantId', 'ServicePrincipalId', 'MfaDetail',
    'NetworkLocationDetails'
)
# Maximum days back to attempt fetching logs
$maxDaysBack = 90
# --- End Configuration ---

# --- Module Check and Import ---
Function Ensure-Modules {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$ModuleNames
    )
    Write-Verbose "Checking required modules: $($ModuleNames -join ', ')"
    $missingModules = @()
    foreach ($moduleName in $ModuleNames) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            $missingModules += $moduleName
            Write-Warning "Required module '$moduleName' not found."
        }
    }

    if ($missingModules.Count -gt 0) {
        $message = "The following required modules are missing:`n$($missingModules -join "`n")`n`nWould you like to attempt installation from the PowerShell Gallery (requires internet and administrator rights if installing for all users)?"
        $result = [System.Windows.Forms.MessageBox]::Show($message, "Missing Modules", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            foreach ($moduleName in $missingModules) {
                try {
                    Write-Host "Attempting to install module '$moduleName'..."
                    Install-Module -Name $moduleName -Repository PSGallery -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                    Write-Host "Successfully installed '$moduleName' for the current user."
                } catch {
                    Write-Warning "Failed to install '$moduleName' for the current user. Error: $($_.Exception.Message)"
                    Write-Host "Attempting to install for all users (requires elevation)..."
                    try {
                        Install-Module -Name $moduleName -Repository PSGallery -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
                        Write-Host "Successfully installed '$moduleName' for all users."
                    } catch {
                        Write-Error "Failed to install module '$moduleName'. Please install it manually and restart the script. Error: $($_.Exception.Message)"
                        Exit 1
                    }
                }
            }
             # Re-check after install attempt
             foreach ($moduleName in $ModuleNames) {
                if (-not (Get-Module -ListAvailable -Name $moduleName)) {
                     Write-Error "Module '$moduleName' still not found after install attempt. Please check PowerShell Gallery access and permissions."
                     Exit 1
                }
             }
        } else {
            Write-Error "Required modules are missing. Please install them manually (`Install-Module $($missingModules -join ', ')`) and restart the script."
            Exit 1
        }
    }

    Write-Verbose "Importing required modules..."
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
        Import-Module Microsoft.Graph.Reports -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
    } catch {
        Write-Error "Failed to import required modules. Please ensure they are installed correctly. Error: $($_.Exception.Message)"
        Exit 1
    }
    Write-Verbose "Modules loaded successfully."
}

# --- GUI Setup ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Entra ID Sign-in Log Investigator'
$form.Size = New-Object System.Drawing.Size(600, 550)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $true

$lblInstructions = New-Object System.Windows.Forms.Label
$lblInstructions.Location = New-Object System.Drawing.Point(10, 10)
$lblInstructions.Size = New-Object System.Drawing.Size(560, 30)
$lblInstructions.Text = "1. Connect to Entra ID. 2. Select user(s) from the list below. 3. Click 'Investigate Selected Users'."
$form.Controls.Add($lblInstructions)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(10, 45)
$btnConnect.Size = New-Object System.Drawing.Size(120, 30)
$btnConnect.Text = 'Connect & Get Users'
$form.Controls.Add($btnConnect)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(140, 45)
$lblStatus.Size = New-Object System.Drawing.Size(430, 30)
$lblStatus.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblStatus)

$lstUsers = New-Object System.Windows.Forms.ListBox
$lstUsers.Location = New-Object System.Drawing.Point(10, 85)
$lstUsers.Size = New-Object System.Drawing.Size(560, 350)
$lstUsers.SelectionMode = 'MultiExtended'
$lstUsers.Enabled = $false
$form.Controls.Add($lstUsers)

$btnInvestigate = New-Object System.Windows.Forms.Button
$btnInvestigate.Location = New-Object System.Drawing.Point(10, 445)
$btnInvestigate.Size = New-Object System.Drawing.Size(180, 30)
$btnInvestigate.Text = 'Investigate Selected Users'
$btnInvestigate.Enabled = $false
$form.Controls.Add($btnInvestigate)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(490, 445)
$btnClose.Size = New-Object System.Drawing.Size(80, 30)
$btnClose.Text = 'Close'
$form.Controls.Add($btnClose)

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(10, 485)
$lblProgress.Size = New-Object System.Drawing.Size(560, 20)
$lblProgress.Text = ''
$form.Controls.Add($lblProgress)

# --- Global Variables ---
$global:EntraUsers = @()
$global:IsConnected = $false

# --- Functions ---
Function Update-Status ($Message, [string]$ColorName = 'Black', $ProgressText = $null) {
    $ActualColor = switch ($ColorName.ToLower()) {
        'black'  { [System.Drawing.Color]::Black }
        'red'    { [System.Drawing.Color]::Red }
        'green'  { [System.Drawing.Color]::Green }
        'orange' { [System.Drawing.Color]::Orange }
        'blue'   { [System.Drawing.Color]::Blue }
        default  { [System.Drawing.Color]::Black }
    }
    try {
        $script:lblStatus.ForeColor = $ActualColor
        $script:lblStatus.Text = "Status: $Message"
    } catch {
        Write-Warning "Error setting status label properties: $($_.Exception.Message)"
        try { $script:lblStatus.Text = "Status: Error displaying status." } catch {}
    }
    if ($ProgressText) { $script:lblProgress.Text = $ProgressText } else { $script:lblProgress.Text = "" }
    $script:form.Refresh()
}

Function Connect-And-LoadUsers {
    Update-Status "Checking modules..." 'Black'
    Ensure-Modules -ModuleNames $requiredModules
    $script:form.Refresh()

    Update-Status "Connecting to Microsoft Graph..." 'Black'
    $scopes = @("User.Read.All", "AuditLog.Read.All")

    try {
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -eq $context) {
            Write-Host "No existing connection found. Attempting to connect..."
            Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        } else {
            Write-Host "Already connected as $($context.Account). Scopes: $($context.Scopes -join ', ')"
            $missingScopes = $scopes | Where-Object { $context.Scopes -notcontains $_ }
            if ($missingScopes.Count -gt 0) {
                Write-Warning "Current connection is missing required scopes: $($missingScopes -join ', '). Reconnecting..."
                Disconnect-MgGraph
                Connect-MgGraph -Scopes $scopes -ErrorAction Stop
            }
        }
        $global:IsConnected = $true
        Update-Status "Connected. Fetching users..." 'Green'

        $script:lstUsers.Items.Clear()
        $script:lstUsers.Enabled = $false
        $script:btnInvestigate.Enabled = $false
        $script:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        $global:EntraUsers = Get-MgUser -All -Select Id, UserPrincipalName, DisplayName -ErrorAction Stop | Sort-Object UserPrincipalName

        if ($global:EntraUsers) {
            foreach ($user in $global:EntraUsers) {
                $script:lstUsers.Items.Add($user.UserPrincipalName) | Out-Null
            }
            $script:lstUsers.Enabled = $true
            $script:btnInvestigate.Enabled = $true
            Update-Status "Connected. Users loaded ($($global:EntraUsers.Count) found)." 'Green'
        } else {
            Update-Status "Connected. No users found or error fetching users." 'Orange'
        }

    } catch {
        $global:IsConnected = $false
        Update-Status "Connection Failed: $($_.Exception.Message)" 'Red'
        [System.Windows.Forms.MessageBox]::Show("Failed to connect or get users.`nError: $($_.Exception.Message)`nPlease ensure you have the correct permissions and network connectivity.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $script:form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

Function Investigate-Users {
    if ($script:lstUsers.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one user from the list.", "No User Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $script:btnInvestigate.Enabled = $false
    $script:btnConnect.Enabled = $false
    $script:lstUsers.Enabled = $false
    $script:btnClose.Enabled = $false
    $script:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Update-Status "Processing..." 'Blue' ""

    $selectedUserPrincipalNames = $script:lstUsers.SelectedItems
    $allSignInLogs = @()
    $startDate = (Get-Date).AddDays(-$maxDaysBack).ToString("yyyy-MM-ddTHH:mm:ssZ")

    Update-Status "Fetching sign-in logs..." 'Blue' "Fetching logs for 0 of $($selectedUserPrincipalNames.Count) users..."
    $userCount = 0
    foreach ($upn in $selectedUserPrincipalNames) {
        $userCount++
        $user = $global:EntraUsers | Where-Object { $_.UserPrincipalName -eq $upn }
        if ($user) {
            Update-Status "Fetching sign-in logs for '$upn'..." 'Blue' "Fetching logs for $userCount of $($selectedUserPrincipalNames.Count) users..."
            try {
                $filterString = "(createdDateTime ge $startDate) and (userId eq '$($user.Id)')"
                Write-Verbose "Fetching logs for User ID: $($user.Id) with filter: $filterString"
                $signInLogs = Get-MgAuditLogSignIn -Filter $filterString -All -ErrorAction Stop
                $signInLogs | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $upn -Force }
                $allSignInLogs += $signInLogs
                Write-Host "Found $($signInLogs.Count) sign-in logs for $upn since $startDate"
            } catch {
                Write-Warning "Could not retrieve sign-in logs for '$upn'. Error: $($_.Exception.Message)"
            }
        } else {
            Write-Warning "Could not find user object for '$upn'. Skipping."
        }
    }

    $tempCsvPath = $null
    if ($allSignInLogs.Count -eq 0) {
        Update-Status "Processing complete. No sign-in logs found for selected users." 'Orange'
        [System.Windows.Forms.MessageBox]::Show("No sign-in logs were found for the selected users within the retention period (max $maxDaysBack days attempted).", "No Logs Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        Update-Status "Processing logs and preparing report..." 'Blue' "Found $($allSignInLogs.Count) total log entries."

        $processedLogs = $allSignInLogs | Select-Object UserPrincipalName, CreatedDateTime, IpAddress, @{N='Country';E={$_.Location.countryOrRegion}}, @{N='State';E={$_.Location.state}}, @{N='City';E={$_.Location.city}}, AppDisplayName, ResourceDisplayName, IsInteractive, Status, ConditionalAccessStatus, DeviceDetail, CorrelationId, Id, UserId, AppId, RiskDetail, RiskLevelAggregated, RiskLevelDuringSignIn, RiskState, RiskEventTypes, ServicePrincipalId, HomeTenantId, ResourceTenantId, MfaDetail, NetworkLocationDetails, ConditionalAccessPolicies

        $tempCsvPath = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName() + ".csv")
        Write-Verbose "Using temporary CSV path: $tempCsvPath"

        try {
            Update-Status "Exporting data to temporary CSV..." 'Blue'
            $processedLogs | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop

            Update-Status "Importing data from temporary CSV..." 'Blue'
            $logsFromCsv = Import-Csv -Path $tempCsvPath -Encoding UTF8 -ErrorAction Stop

            $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
            $saveFileDialog.Title = "Save Sign-in Log Report"
            $saveFileDialog.FileName = "EntraSignInLogs_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"

            if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $outputPath = $saveFileDialog.FileName
                try {
                    # --- Create Excel File (BASIC Formatting Only) ---
                    # Conditional formatting rule definition REMOVED
                    # Parameters -ConditionalFormatting and -HideColumn REMOVED from Export-Excel call

                    Update-Status "Creating Excel file..." 'Blue' "File: $outputPath"
                    Write-Verbose "Exporting to $outputPath from CSV data (Basic Formatting)"

                    # Basic Export-Excel call without the problematic parameters
                    $logsFromCsv | Export-Excel -Path $outputPath -WorksheetName 'SignIns' -AutoSize -FreezeTopRow -BoldTopRow -ErrorAction Stop

                    Update-Status "Report saved successfully!" 'Green' "File: $outputPath"
                    [System.Windows.Forms.MessageBox]::Show("Sign-in log report created successfully at:`n$outputPath", "Report Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                } catch {
                    Update-Status "Error creating Excel file: $($_.Exception.Message)" 'Red'
                    [System.Windows.Forms.MessageBox]::Show("Failed to create the Excel report.`nError: $($_.Exception.Message)", "Excel Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } else {
                Update-Status "Excel export cancelled by user." 'Orange'
            }

        } catch {
            Update-Status "Error during CSV processing: $($_.Exception.Message)" 'Red'
            [System.Windows.Forms.MessageBox]::Show("An error occurred during the temporary CSV processing.`nError: $($_.Exception.Message)", "CSV Processing Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            if ($null -ne $tempCsvPath -and (Test-Path $tempCsvPath)) {
                Write-Verbose "Removing temporary file: $tempCsvPath"
                Remove-Item -Path $tempCsvPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $script:btnInvestigate.Enabled = $global:IsConnected
    $script:btnConnect.Enabled = $true
    $script:lstUsers.Enabled = $global:IsConnected
    $script:btnClose.Enabled = $true
    $script:form.Cursor = [System.Windows.Forms.Cursors]::Default
}

# --- Event Handlers ---
$btnConnect.Add_Click({ Connect-And-LoadUsers })
$btnInvestigate.Add_Click({ Investigate-Users })
$btnClose.Add_Click({
    if ($global:IsConnected) { try { Disconnect-MgGraph } catch {} }
    $form.Close()
})

# --- Show Form ---
$lblStatus.ForeColor = [System.Drawing.Color]::Red
$lblStatus.Text = 'Status: Not Connected'
$form.ShowDialog() | Out-Null
# --- End Script ---