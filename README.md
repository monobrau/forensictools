# Entra ID Forensic Log Fetcher GUI (v3.4)

A PowerShell script with a graphical user interface (GUI) designed to assist with forensic analysis by fetching sign-in logs for specified Entra ID (formerly Azure AD) user accounts and exporting them to a formatted Excel (XLSX) file.

## Description

This tool provides a simple Windows Forms interface to:
* Connect to Microsoft Graph.
* Automatically load all Entra ID users.
* Allow selection of one or more users for investigation.
* Specify a time duration (1-30 days) for log retrieval.
* Export sign-in logs to an XLSX file, which includes:
    * The tenant's primary domain name (or a fallback identifier) in the filename.
    * Auto-fitted column widths.
    * A bolded header row.
    * Rows highlighted in yellow where the sign-in country is "United States".
* Provide a button to disconnect the Graph session.
* Provide a button to open the last successfully exported XLSX file.

## Features

* **Graphical User Interface:** Easy-to-use GUI built with Windows Forms.
* **Microsoft Graph Integration:** Connects securely to Microsoft Graph using the official PowerShell SDK.
* **Automatic User Loading:** Loads a list of all users from your Entra ID tenant immediately after successful connection.
* **User Selection:** Allows selecting one or more users for log retrieval via a checklist.
* **Custom Duration:** Specify the number of past days (1-30) of sign-in history to retrieve.
    * Includes a warning that retrieving logs beyond 7 days typically requires an Entra ID P1 or P2 license.
* **Folder Selection:** Choose a destination folder for the exported log files.
* **Robust Filtering:** Fetches logs using the user's immutable Object ID for better reliability.
* **Intelligent Filenaming:** Includes the tenant's primary domain name (or UPN-derived domain, or Tenant ID as fallback) in the exported XLSX filename for easy identification.
* **Formatted XLSX Export:**
    * Exports logs first to a temporary CSV, then converts to XLSX.
    * **Auto-fits column widths** in the XLSX file.
    * Makes the **header row bold**.
    * **Highlights rows in yellow** where the `Country` is "United States".
    * Expands complex properties like `DeviceDetail` and `Status` into readable columns, including MFA-related details from `Status.AdditionalDetails`.
* **Session Management:** Includes a "Disconnect from Graph" button to terminate the session and clear user data.
* **Open Last File:** A button to quickly open the most recently generated XLSX report.

## Prerequisites

* **PowerShell:** Version 5.1 or later.
* **Microsoft Graph PowerShell SDK:** The script requires the following modules:
    * `Microsoft.Graph.Users`
    * `Microsoft.Graph.Reports`
    * `Microsoft.Graph.Identity.DirectoryManagement`
* **Microsoft Excel:** **Must be installed** on the machine running the script for the XLSX conversion and formatting features to work.
* **Entra ID Permissions:** The account running the script needs the following delegated permissions granted in Microsoft Graph/Entra ID:
    * `User.Read.All` (to list users)
    * `AuditLog.Read.All` (to read sign-in logs)
    * `Organization.Read.All` (to determine tenant domain for filename)
* **Operating System:** Windows (due to Windows Forms GUI and Excel COM automation).

## Installation

1.  **Save the Script:** Download or save the script file (e.g., `EntraID_Forensic_Log_Fetcher.ps1`) to your local machine.
2.  **Install Modules:** If you don't have the required Microsoft Graph modules installed, run the following command in PowerShell (as the user who will run the script):
    ```powershell
    Install-Module Microsoft.Graph.Users, Microsoft.Graph.Reports, Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Repository PSGallery -Force
    ```
    The script also includes a check and will prompt for installation if modules are missing.

## Usage

1.  **Run the Script:** Open PowerShell, navigate to the directory where you saved the script, and run it:
    ```powershell
    .\EntraID_Forensic_Log_Fetcher.ps1
    ```
2.  **Connect & Load Users:** Click the "Connect & Load Users" button. A Microsoft login window will appear. Authenticate using an account with the required permissions. Users will be loaded automatically. The script will also attempt to determine your tenant's domain for the output filename.
3.  **Select Users:** Check the boxes next to the users you want to investigate.
4.  **Set Duration:** Enter the number of past days (1-30) for which you want to retrieve logs. Note the license warning if selecting more than 7 days.
5.  **Choose Output Folder:** Click "Browse..." and select a destination folder for the XLSX file.
6.  **Get Logs:** Click the "Get Sign-in Logs for Selected Users" button.
7.  **Monitor Progress:** The status bar at the bottom will show the script's progress. The PowerShell console window will also display detailed messages, including any warnings or errors encountered.
8.  **Review Output:** Upon completion, a message box will confirm the successful export and formatting of the XLSX file.
9.  **Open File (Optional):** Click the "Open Last Exported File" button to launch the generated report in Excel.
10. **Disconnect (Optional):** Click "Disconnect from Graph" to end the session. This will clear the user list and disable log fetching until you reconnect.

## Notes

* The actual availability of sign-in logs depends on your Entra ID tenant's log retention settings and license (Free = 7 days, P1/P2 = 30 days by default, potentially longer if configured with Log Analytics).
* Error handling is included, but ensure the account running the script has sufficient permissions and network connectivity to Microsoft Graph endpoints. Check the console output for detailed error messages if issues occur, especially related to Excel COM automation if Excel is not properly installed or accessible.
* If Excel is not installed or accessible, the script will export a CSV file, but the XLSX conversion and formatting will fail. The CSV file will be retained in this case.

## License

Consider adding a license file (e.g., MIT License) if you plan to share this script publicly.
