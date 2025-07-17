# 🔍 Entra Investigator

**A comprehensive PowerShell GUI tool for investigating Microsoft Entra ID (Azure AD) accounts during security incidents, compliance audits, and user analysis.**

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/en-us/powershell/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph%20SDK-Required-green.svg)](https://docs.microsoft.com/en-us/graph/powershell/installation)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-6.0-brightgreen.svg)](CHANGELOG.md)

---

## 📋 Table of Contents

- [Features](#-features)
- [Screenshots](#-screenshots)
- [Prerequisites](#-prerequisites)
- [Installation](#-installation)
- [Quick Start](#-quick-start)
- [Detailed Usage](#-detailed-usage)
- [Required Permissions](#-required-permissions)
- [Use Cases](#-use-cases)
- [Troubleshooting](#-troubleshooting)
- [Contributing](#-contributing)
- [License](#-license)
- [Changelog](#-changelog)

---

## 🚀 Features

### **Core Investigation Capabilities**

- **🔐 Comprehensive MFA Analysis**
  - Per-user MFA method detection
  - Security Defaults status checking
  - Conditional Access policy evaluation
  - Overall protection status with risk indicators

- **👤 Enhanced User Analysis**
  - License type detection (Premium vs Standard)
  - Account status monitoring (Enabled/Disabled)
  - Password age analysis
  - Administrative role identification
  - Group membership counting

- **📊 Sign-in Log Export**
  - Bulk sign-in log extraction for multiple users
  - Excel export with automatic formatting
  - Geographic location tracking
  - Failed authentication analysis
  - Customizable time ranges (1-30 days)

- **📝 Audit Log Investigation**
  - Administrative activity tracking
  - Real-time audit log analysis
  - Filtered by specific users
  - Export capabilities for compliance

- **🎯 User Details & Roles**
  - Complete user profile information
  - Active directory role assignments
  - Group membership details
  - Account status and settings

### **Export & Reporting**

- **📄 Multiple Export Formats**: CSV, XLSX, TXT
- **🎨 Formatted Excel Output**: Automatic column sizing, highlighting
- **📋 Compliance Ready**: Export formats suitable for audits
- **🔄 Bulk Operations**: Process multiple users simultaneously

---

## 📸 Screenshots

### Main Interface
*User selection and analysis tools*

### MFA Analysis Tab
*Comprehensive multi-factor authentication status checking*

### Sign-in Logs Export
*Bulk log extraction with Excel formatting*

---

## 🗂️ Script Overview

This toolkit includes three PowerShell scripts for Microsoft Entra ID (Azure AD) investigation and reporting:

### 1. `entrainvestigator.ps1`
A GUI-based tool for interactive investigation of Entra ID accounts. Features tabbed navigation for sign-in logs, user details, audit logs, and MFA analysis. Best for hands-on, multi-user analysis and export.

**How to use:**
```powershell
# Run the GUI tool
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
./entrainvestigator.ps1
```
- Connect to Microsoft Graph, select users, and use the tabs for investigation and export.

---

### 2. `entrareporter.ps1`
A command-line tool for generating comprehensive forensic reports for selected users. Supports multi-user selection, consolidated or individual Excel reports, and detailed license/MFA/audit analysis.

**How to use:**
```powershell
# Run the forensic report generator
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
./entrareporter.ps1
```
- Authenticate, select users from a grid, choose report type, and save Excel and summary reports.

---

### 3. `auditexplorer.ps1`
A focused tool for exporting Entra ID audit logs for selected users. Provides a Windows Forms interface for user selection and CSV export of audit activity.

**How to use:**
```powershell
# Run the audit log explorer
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
./auditexplorer.ps1
```
- Connect, select users, set time range, and export audit logs to CSV.

---

## 🔧 Prerequisites

### **System Requirements**
- **Windows 10/11** or **Windows Server 2016+**
- **PowerShell 5.1** or later
- **Microsoft Excel** (for XLSX export functionality)
- **Internet Connection** (for Microsoft Graph API access)

### **Required PowerShell Modules**
The tool will automatically prompt to install missing modules:

```powershell
Microsoft.Graph.Users
Microsoft.Graph.Reports  
Microsoft.Graph.Identity.DirectoryManagement
Microsoft.Graph.Identity.SignIns
```

### **Microsoft Graph Permissions**
- `User.Read.All`
- `AuditLog.Read.All`
- `Organization.Read.All`
- `Directory.Read.All`
- `Policy.Read.All`
- `UserAuthenticationMethod.Read.All`

---

## 📦 Installation

Download all three scripts (`entrainvestigator.ps1`, `entrareporter.ps1`, `auditexplorer.ps1`) to your preferred directory. See usage above for how to launch each tool.

### **Option 1: Direct Download**

1. Download the latest `Entra-Investigator.ps1` file
2. Save to your preferred directory
3. Run with PowerShell (see Quick Start below)

### **Option 2: Git Clone**

```bash
git clone https://github.com/yourusername/entra-investigator.git
cd entra-investigator
```

### **Option 3: PowerShell Gallery** *(Coming Soon)*

```powershell
Install-Script -Name Entra-Investigator
```

---

## 🚀 Quick Start

### **Step 1: Launch the Tool**

```powershell
# Navigate to the script directory
cd C:\Path\To\Script

# Execute with appropriate execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\Entra-Investigator.ps1
```

### **Step 2: Connect to Microsoft Graph**

1. Click **"Connect & Load Users"**
2. Sign in with appropriate administrative credentials
3. Consent to required permissions
4. Wait for user list to populate

### **Step 3: Select Users for Investigation**

- ✅ Check individual users or use **"Select All"**
- 🔍 Click **"Analyze Selected Users"** for enhanced information display

### **Step 4: Investigate Using Tabs**

- **Sign-in Logs**: Export authentication data
- **User Details**: View comprehensive user information  
- **Audit Logs**: Check administrative activities
- **MFA Analysis**: Evaluate security posture

---

## 📖 Detailed Usage

### **Enhanced User Analysis**

The **"Analyze Selected Users"** feature provides at-a-glance security information:

```
john.doe@company.com [P1/P2] | ✓ Enabled | Pwd: 07/24/2019 21:53:30 | Roles: Global Administrator | Groups: 5 Groups
jane.smith@company.com [Standard] | ✗ Disabled | Pwd: Never | Roles: No Admin Roles | Groups: No Groups
```

**Legend:**
- `[P1/P2]` = Premium License | `[Standard]` = Basic License
- `✓ Enabled` = Active Account | `✗ Disabled` = Inactive Account
- `Pwd:` = Last password change date
- `Roles:` = Administrative role assignments
- `Groups:` = Total group memberships

### **MFA Analysis Deep Dive**

The MFA Analysis tab provides comprehensive multi-factor authentication assessment:

#### **Overall Status Examples:**

✅ **Protected (Security Defaults)**
```
OVERALL MFA STATUS: Protected (Security Defaults)
MFA required via Security Defaults

1. PER-USER MFA: ✗ NOT ENABLED
2. SECURITY DEFAULTS: ✓ ENABLED  
3. CONDITIONAL ACCESS: ✗ NO
```

⚠️ **Not Protected**
```
OVERALL MFA STATUS: ⚠️ NOT PROTECTED
No MFA protection detected

1. PER-USER MFA: ✗ NOT ENABLED
2. SECURITY DEFAULTS: ✗ DISABLED
3. CONDITIONAL ACCESS: ✗ NO
```

### **Sign-in Log Export**

1. **Select Users**: Choose one or more users to investigate
2. **Set Time Range**: Configure days of history (1-30 days)
3. **Choose Output Folder**: Select destination for export files
4. **Export**: Click "Get Sign-in Logs & Export to XLSX"

**Output Includes:**
- User Principal Name
- Timestamp
- Application Used
- IP Address & Geographic Location
- Authentication Status
- Failure Reasons (if applicable)

### **Audit Log Investigation**

Monitor administrative activities performed by specific users:

- **Date/Time**: When the activity occurred
- **Activity**: What action was performed
- **Category**: Type of administrative action
- **Result**: Success/Failure status
- **Target/Object**: What was modified

---

## 🔐 Required Permissions

### **Microsoft Graph API Permissions**

| Permission | Type | Justification |
|------------|------|---------------|
| `User.Read.All` | Application/Delegated | Read user profiles and basic information |
| `AuditLog.Read.All` | Application/Delegated | Access sign-in and audit logs |
| `Organization.Read.All` | Application/Delegated | Read tenant-level settings |
| `Directory.Read.All` | Application/Delegated | Read directory objects and relationships |
| `Policy.Read.All` | Application/Delegated | Read conditional access and security policies |
| `UserAuthenticationMethod.Read.All` | Application/Delegated | Read user MFA methods and settings |

### **Administrative Roles Required**

**Minimum Required:**
- **Security Reader** - For audit log access
- **Global Reader** - For comprehensive tenant information

**Recommended:**
- **Security Administrator** - For full investigation capabilities
- **Global Administrator** - For complete access (incident response scenarios)

---

## 💼 Use Cases

### **🚨 Security Incident Response**

- **Compromised Account Investigation**: Analyze sign-in patterns, MFA status, and recent activities
- **Privilege Escalation Detection**: Check for unexpected role assignments
- **Lateral Movement Tracking**: Export sign-in logs for timeline analysis

### **✅ Compliance Auditing**

- **MFA Coverage Assessment**: Verify multi-factor authentication deployment
- **Administrative Activity Review**: Export audit logs for compliance reporting
- **User Access Reviews**: Bulk analysis of user permissions and status

### **🔍 Routine Security Operations**

- **User Onboarding Verification**: Ensure proper MFA setup for new users
- **Periodic Access Reviews**: Bulk user analysis for role and group cleanup
- **Password Policy Compliance**: Identify accounts with old passwords

### **📊 Security Metrics & Reporting**

- **MFA Adoption Rates**: Measure tenant-wide MFA deployment
- **Administrative Activity Trends**: Track admin actions over time
- **Geographic Access Patterns**: Identify unusual sign-in locations

---

## 🛠️ Troubleshooting

### **Common Issues**

#### **"Connect & Load Users" Button Not Working**
```powershell
# Verify Graph modules are installed
Get-Module -ListAvailable Microsoft.Graph*

# Manual module installation
Install-Module Microsoft.Graph -Scope CurrentUser
```

#### **Permission Errors**
- Ensure you have appropriate administrative roles
- Verify all required Graph permissions are consented
- Check tenant conditional access policies aren't blocking access

#### **Excel Export Failures**
- Verify Microsoft Excel is installed locally
- Check file path permissions for output directory
- Ensure Excel isn't currently running/locked

#### **Empty MFA Policy Grid**
This is **expected behavior** when:
- Security Defaults are enabled (no CA policies needed)
- No Conditional Access policies exist
- User is excluded from existing policies

### **Debug Mode**

Enable verbose logging for troubleshooting:

```powershell
$VerbosePreference = "Continue"
.\Entra-Investigator.ps1
```

---

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details.

### **Development Setup**

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### **Bug Reports**

Please use the [Issue Tracker](../../issues) and include:

- **Environment Details**: PowerShell version, Windows version, Graph SDK version
- **Error Messages**: Full error text and stack traces
- **Steps to Reproduce**: Detailed reproduction steps
- **Expected vs Actual Behavior**: What should happen vs what does happen

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 📅 Changelog

### **Version 6.0** *(Latest)*
- ✨ **NEW**: Comprehensive MFA Analysis tab
- ✨ **NEW**: Enhanced user analysis with license, status, and role information
- 🔧 **IMPROVED**: Better audit log display with 5-column layout
- 🔧 **IMPROVED**: Export capabilities for all analysis types
- 🐛 **FIXED**: All event handlers now working correctly

### **Version 5.5**
- 🔧 **FIXED**: Connect & Load Users functionality restored
- 🔧 **IMPROVED**: Audit log column definitions and display

### **Version 5.4**
- 🔧 **FIXED**: Layout stability for audit grid
- 🔧 **IMPROVED**: Excel conversion and formatting

[View Full Changelog](CHANGELOG.md)

---

## 🙏 Acknowledgments

- **Microsoft Graph Team** - For excellent API documentation
- **PowerShell Community** - For Windows Forms examples and best practices
- **Security Community** - For feedback and feature requests

---

## 📞 Support

- **Documentation**: [Wiki](../../wiki)
- **Issues**: [GitHub Issues](../../issues)
- **Discussions**: [GitHub Discussions](../../discussions)
- **Security Issues**: Please email security@yourorg.com

---

**⭐ If this tool helps your security investigations, please star the repository!**

---

*Entra Investigator - Making Microsoft Entra ID security investigations efficient and comprehensive.*