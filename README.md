# Entra ID Sign-in Log Investigator GUI

**Version:** 1.4 (Last Updated: 2025-05-05)

A PowerShell script that provides a simple Graphical User Interface (GUI) for basic forensic analysis of Microsoft Entra ID (formerly Azure AD) user sign-in logs. It allows selecting users and exporting their recent sign-in activity to an Excel file.

## Features

* **Graphical User Interface:** Uses Windows Forms for easier interaction.
* **Entra ID Connection:** Connects to Microsoft Graph using delegated permissions.
* **User Listing:** Fetches and displays a list of all users in the tenant.
* **User Selection:** Allows selection of one or multiple users for investigation.
* **Sign-in Log Retrieval:** Fetches sign-in logs for the selected user(s). By default, it attempts to retrieve logs from the last 90 days, but this is limited by your tenant's actual log retention policies (see Limitations).
* **Intermediate CSV:** Exports data to a temporary CSV file behind the scenes before creating the Excel report.
* **Excel Export:** Exports the retrieved sign-in data to an `.xlsx` file with basic formatting (Auto-sized columns, frozen top row, bold top row).

## Prerequisites

1.  **Operating System:** Windows (required for Windows Forms GUI).
2.  **PowerShell:** Version 5.1 or later.
3.  **PowerShell Modules:**
    * `Microsoft.Graph` (specifically Authentication, Users, Reports submodules)
    * `ImportExcel`
    * The script includes a check and will prompt to install these modules from the PowerShell Gallery if they are missing (requires internet connection and potentially administrator rights).
4.  **Entra ID Permissions:** The user running the script needs delegated permissions to read user information and audit logs in Microsoft Graph. Required permissions typically include:
    * `User.Read.All`
    * `AuditLog.Read.All`
    * Alternatively, Entra ID roles like `Global Reader` or `Security Reader` usually suffice.
5.  **Internet Connection:** Required to connect to Microsoft Graph and potentially download modules.

## Installation

1.  Download the script file (e.g., `EntraLogInvestigator.ps1`) to your local machine.
2.  Ensure all prerequisites are met, especially PowerShell modules and Entra ID permissions.

## Usage

1.  **Open PowerShell:** Navigate to the directory where you saved the script file.
2.  **Run the script:**
    ```powershell
    .\EntraLogInvestigator.ps1
    ```
    * If prompted to install missing modules, type `Y` and press Enter (you might need to run PowerShell as Administrator for module installation).
3.  **Connect:** Click the "Connect & Get Users" button in the GUI. A Microsoft sign-in window will appear. Authenticate using an account with the necessary Entra ID permissions.
4.  **Wait for User List:** The script will fetch and display the list of users. The status bar will update.
5.  **Select Users:** Select one or more users from the list using your mouse (Ctrl+Click for multiple individual users, Shift+Click for a range).
6.  **Investigate:** Click the "Investigate Selected Users" button.
7.  **Wait for Processing:** The script will fetch sign-in logs for the selected users. This may take some time depending on the number of users and logs. The status bar provides updates.
8.  **Save Report:** A "Save As" dialog box will appear. Choose a location and filename for the Excel report (`.xlsx`) and click "Save".
9.  **Review Report:** A confirmation message will appear. Open the saved Excel file to review the sign-in log data.
10. **Close:** Click the "Close" button to exit the application.

## Important Notes & Limitations

* **Log Retention:** Microsoft Entra ID sign-in log retention depends on your tenant's license (e.g., Free/Basic: 7 days, P1/P2: 30 days by default). This script can only retrieve logs available within your tenant's configured retention period, even though it attempts up to 90 days back.
* **Feature Removals (v1.4):** Due to compatibility issues encountered with the `ImportExcel` module in some environments, the following features have been **removed** from this version:
    * **Conditional Formatting:** The script no longer highlights sign-in rows from non-US locations in yellow.
    * **Column Hiding:** The script no longer automatically hides columns deemed less human-readable (e.g., IDs, GUIDs). All columns selected by the script will be present in the output file.
* **Administrator Rights:** Running PowerShell as Administrator might be required for the initial installation of modules if installing for "All Users".
* **Throttling:** Fetching large amounts of data from Microsoft Graph may be subject to API throttling. The script has basic error handling but no sophisticated retry logic for throttling.

## Troubleshooting

* **"Parameter cannot be found..." errors during Excel export:** This script version (v1.4) specifically removes the `-ConditionalFormatting` and `-HideColumn` parameters to avoid these errors encountered with older or problematic `ImportExcel` module installations. If you encounter *other* parameter errors related to `Export-Excel`, ensure the `ImportExcel` module is installed correctly.
* **Connection Failed:** Verify you have the correct Entra ID permissions (`User.Read.All`, `AuditLog.Read.All`), network connectivity to Microsoft Graph endpoints, and that the account used for login is valid.
