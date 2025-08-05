# Microsoft 365 Management Tools Collection

A comprehensive suite of web-based tools for Microsoft 365 administrators to manage, monitor, and report on various aspects of their environment, including Intune policies, application deployment, compliance monitoring, and Windows Update reporting.

---

## 🚀 Tools Included

### 1. Intune App Repository
**Files**:  
`intune-app-repository.html`, `intune-app-repository.css`, `intune-app-repository.js`, `intune-app-repository-data.js`  
A searchable catalog of common enterprise applications with direct download links for MSI/EXE packages, optimized for Intune deployment.

**Features**:
- 100+ pre-configured applications with download URLs
- ARM64 support indicators
- Filter by category, publisher, or search
- Direct MSI download links for enterprise deployment
- Copy URL functionality for easy sharing

---

### 2. Intune Compliance Manager
**File**: `IntuneComplianceManager.html`  
Real-time monitoring and management of device compliance status across your Intune-managed environment.

**Features**:
- Live device compliance monitoring
- Policy assignment tracking
- Compliance trend analysis
- Export capabilities (CSV/JSON)
- Interactive dashboards with Chart.js visualizations

---

### 3. Intune Policy Manager
**File**: `IntunePolicyManager.html`  
Upload, manage, and bulk-modify Intune configuration policies with advanced assignment capabilities.

**Features**:
- Bulk policy upload from JSON
- Policy duplication and modification
- Group-based assignment management
- Policy deletion with safety confirmations
- Export existing policies for backup

---

### 4. Microsoft 365 Policy Downloader
**File**: `Microsoft365PolicyDownloader.html`  
Export and backup policies from all Microsoft 365 admin portals including Intune, Exchange, SharePoint, Teams, Azure AD, and Compliance centers.

**Features**:
- Multi-service policy extraction
- Comprehensive backup capabilities
- Individual or bulk export options
- Support for all major M365 services

---

### 5. Windows Update Report Generator
**Files**:  
`WindowsUpdateReport.html`, `windows-update-report.css`, `windows-update-report.js`  
Generate professional PDF reports from Windows Update CSV data with charts and compliance analytics.

**Features**:
- CSV data import and parsing
- Interactive Chart.js visualizations
- Professional PDF report generation
- Compliance rate calculations
- Security severity analysis

---

## 🔒 Security Features
All tools include:
- Password protection with SHA-256 hashing
- Session timeout after 30 minutes of inactivity
- 2-hour session expiry
- No data storage (all processing is client-side)
- OAuth 2.0 authentication via Microsoft Entra ID

---

## 🛠️ Setup Instructions

### Prerequisites
- Modern web browser (Chrome, Edge, Firefox, Safari)
- Microsoft 365 Global Administrator or equivalent permissions
- Entra ID App Registration (for tools requiring authentication)

### Quick Start
1. Download the required files from this repository
2. Host the files on a web server or run them locally
3. Register an Entra ID application (for authenticated tools)
4. Configure permissions as specified in each tool
5. Access the tool and enter your Client ID

---

## 🔑 Entra ID App Registration

For tools requiring Microsoft Graph access:

1. Go to [Microsoft Entra Portal](https://entra.microsoft.com)
2. Navigate to: **Applications → App registrations**
3. Click **New registration**
4. Configure the following:
   - **Name**: Your tool name
   - **Supported account types**: Single or multi-tenant
   - **Redirect URI**: Single-page application (SPA) — your hosted URL
5. Add the required API permissions (see below)
6. Grant admin consent for the permissions

---

## ✅ Required Permissions by Tool

### Intune Compliance Manager
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementApps.Read.All`
- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`
- `User.Read`

### Intune Policy Manager
- `DeviceManagementConfiguration.ReadWrite.All`
- `DeviceManagementApps.ReadWrite.All`
- `DeviceManagementServiceConfig.ReadWrite.All`
- `Group.Read.All`
- `User.Read`

### Microsoft 365 Policy Downloader
- `DeviceManagementConfiguration.Read.All`
- `Policy.Read.All`
- `Directory.Read.All`
- `Group.Read.All`
- `Sites.Read.All`
- `User.Read`

---

## 📊 Browser Compatibility

| Browser  | Minimum Version | Notes         |
|----------|------------------|----------------|
| Chrome   | 80+              | Recommended    |
| Edge     | 80+              | Recommended    |
| Firefox  | 75+              | Fully supported |
| Safari   | 13.1+            | Fully supported |

---

## 🎯 Use Cases

- **MSPs**: Manage multiple client environments efficiently  
- **IT Departments**: Streamline policy management and compliance monitoring  
- **Security Teams**: Generate compliance reports and track security updates  
- **Consultants**: Quickly assess and document client environments  

---

## ⚠️ Important Notes

- **Authentication**: Each tool requires a valid Entra ID app registration with proper permissions  
- **Data Privacy**: All processing occurs client-side — no data is sent to external servers  
- **Browser Storage**: Tools use `sessionStorage` for temporary data (cleared on browser close)  
- **Rate Limiting**: Microsoft Graph API limits apply; retry logic is included  
- **Password**: Default password should be changed in production deployments  

---

## 🤝 Contributing

We welcome issues, feature requests, and pull requests. When contributing:
- Maintain the existing code style
- Test thoroughly across browsers
- Update documentation as needed
- Follow security best practices

---

## 📝 License

These tools are provided *as-is* for educational and administrative purposes. Ensure compliance with your organization’s policies before use.

---

## 🆘 Support

For issues or questions:
- Use the built-in help section in each tool
- Review browser console logs for error messages
- Confirm all prerequisites are met
- Verify that API permissions are correctly configured

---

## 🔄 Updates

- Regular updates for new Intune features
- Application repository updates quarterly
- Security patches as needed
- Feature requests reviewed monthly

> **Note**: These tools require Microsoft 365 administrative privileges. Always test in a non-production environment first.
```
