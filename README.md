# Entra ID Forensic Log Fetcher GUI

A PowerShell script with a graphical user interface (GUI) designed to assist with forensic analysis by fetching sign-in logs for specified Entra ID (formerly Azure AD) user accounts.

## Description

This tool provides a simple Windows Forms interface to connect to Microsoft Graph, select Entra ID users, specify a time duration, and export their sign-in logs to a CSV file. It aims to streamline the initial log collection process for forensic investigations involving user accounts.

## Features

* **Graphical User Interface:** Easy-to-use GUI built with Windows Forms.
* **Microsoft Graph Integration:** Connects securely to Microsoft Graph using the official PowerShell SDK.
* **User Loading:** Loads a list of all users from your Entra ID tenant.
* **User Selection:** Allows selecting one or multiple users for log retrieval via a checklist.
* **Custom Duration:** Specify the number of days (1-30) of sign-in history to retrieve.
    * Includes a warning that retrieving logs beyond 7 days typically requires an Entra ID P1 or P2 license.
* **Folder Selection:** Choose a destination folder for the exported logs.
* **Robust Filtering:** Fetches logs using the user's immutable Object ID for better reliability.
* **CSV Export:** Exports the collected sign-in logs, including detailed location information (City, State, Country), to a timestamped CSV file.

## Prerequisites

* **PowerShell:** Version 5.1 or later.
* **Microsoft Graph PowerShell SDK:** The script requires the `Microsoft.Graph.Users` and `Microsoft.Graph.Reports` modules.
* **Entra ID Permissions:** The account running the script needs the following delegated permissions granted in Microsoft Graph/Entra ID:
    * `User.Read.All` (to list users)
    * `AuditLog.Read.All` (to read sign-in logs)
* **Operating System:** Windows (due to Windows Forms GUI).

## Installation

1.  **Save the Script:** Download or save the script file (e.g., `EntraID_Forensic_Log_Fetcher.ps1`) to your local machine.
2.  **Install Modules:** If you don't have the required Microsoft Graph modules installed, run the following command in PowerShell (as the user who will run the script):
    ```powershell
    Install-Module Microsoft.Graph.Users, Microsoft.Graph.Reports -Scope CurrentUser -Repository PSGallery -Force
    ```
    The script also includes a check and will prompt for installation if modules are missing.

## Usage

1.  **Run the Script:** Open PowerShell, navigate to the directory where you saved the script, and run it:
    ```powershell
    .\EntraID_Forensic_Log_Fetcher.ps1
    ```
2.  **Connect:** Click the "Connect to Graph" button. A Microsoft login window will appear. Authenticate using an account with the required permissions.
3.  **Load Users:** Click "Load Entra Users". The list box will populate with users from your tenant.
4.  **Select Users:** Check the boxes next to the users you want to investigate.
5.  **Set Duration:** Enter the number of past days (1-30) for which you want to retrieve logs. Note the license warning if selecting more than 7 days.
6.  **Choose Output Folder:** Click "Browse..." and select a destination folder for the CSV file.
7.  **Get Logs:** Click the "Get Sign-in Logs for Selected Users" button.
8.  **Monitor Progress:** The status bar at the bottom will show the script's progress. The PowerShell console window will also display detailed messages, including any warnings or errors encountered during the process (e.g., if logs couldn't be retrieved for a specific user).
9.  **Review Output:** Upon completion, a message box will confirm the successful export of the CSV file to the specified location.

## Notes

* The actual availability of sign-in logs depends on your Entra ID tenant's log retention settings and license (Free = 7 days, P1/P2 = 30 days by default, potentially longer if configured with Log Analytics).
* This version exports directly to CSV. Previous versions included XLSX conversion and formatting, but this was removed due to potential issues with Excel COM object interaction across different environments. You can easily open the CSV in Excel or other spreadsheet software for further analysis.
* Error handling is included, but ensure the account running the script has sufficient permissions and network connectivity to Microsoft Graph endpoints. Check the console output for detailed error messages if issues occur.

## License

Consider adding a license file (e.g., MIT License) if you plan to share this script publicly.

