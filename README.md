# the HAM

Hosted Administration & Management: A unified approach to managing Microsoft 365 services including Exchange Online, SharePoint, OneDrive, and Intune through web-based tools.

## Tools

### Intune Policy Manager

Deploy CIS compliance workflows and security baselines to Microsoft Intune with bulk upload capabilities.

**Access:** https://luiserodz.github.io/theHAM/IntunePolicyManager.html

### Microsoft 365 Policy Downloader

Export and backup Intune configurations, compliance policies, and device configurations for documentation or migration.

**Access:** https://luiserodz.github.io/theHAM/Microsoft365PolicyDownloader.html

### Intune App Repository

Browse and download common application installers for deployment through Intune.

**Access:** https://luiserodz.github.io/theHAM/IntuneAppRepository.html

## Usage

### Intune Policy Manager

1. Visit https://luiserodz.github.io/theHAM/IntunePolicyManager.html
2. Enter your Azure AD Client ID (or use the pre-configured one)
3. Sign in with your Microsoft work account
4. Select JSON workflow files to upload
5. Configure assignment options (All Users, All Devices, or specific groups)
6. Click Upload to deploy policies to Intune

### Microsoft 365 Policy Downloader

1. Visit https://luiserodz.github.io/theHAM/Microsoft365PolicyDownloader.html
2. Enter your Azure AD Client ID (or use the pre-configured one)
3. Sign in with your Microsoft work account
4. Select policy types to export (Configuration, Compliance, Apps)
5. Choose individual policies or bulk download
6. Download policies as JSON files

### Intune App Repository

1. Visit https://luiserodz.github.io/theHAM/IntuneAppRepository.html
2. Use the search box or filters to find an application
3. Click the **x86/x64** or **ARM64** button to download the installer
4. Use the copy icon to copy download links to the clipboard
5. Check **ARM64 Only** to filter for applications with ARM packages

## Prerequisites

- Microsoft 365 Global Administrator or Intune Administrator role
- Valid Intune licenses
- Modern web browser (Chrome, Edge, Firefox)
- JSON workflow files (for Policy Manager)
- Azure AD app registration with appropriate permissions

## Authentication Setup

1. Register an application in Azure AD
2. Add required API permissions:

- `DeviceManagementConfiguration.ReadWrite.All`
- `DeviceManagementApps.ReadWrite.All`
- `Group.Read.All`

3. Configure redirect URI: `https://luiserodz.github.io/theHAM/[ToolName].html`
4. Grant admin consent for permissions
5. Copy the Application (Client) ID

## Features

### Intune Policy Manager

- Bulk upload multiple policies simultaneously
- Automatic policy naming with custom prefixes
- Group-based assignment targeting
- Real-time upload progress tracking
- Detailed results with CSV export
- Support for all Intune configuration policy types

### Microsoft 365 Policy Downloader

- Export all policy types (Configuration, Compliance, Application)
- Individual or bulk download options
- JSON format for easy reimport
- Policy metadata preservation
- Filtered search and selection
- Backup scheduling recommendations

## Troubleshooting

### Authentication Issues

- Enable popups for GitHub Pages domain
- Verify admin consent is granted in Azure AD
- Check redirect URI matches exactly
- Clear browser cache and cookies
- Try incognito/private browsing mode

### API Errors

- **403 Forbidden**: Missing permissions - grant admin consent
- **401 Unauthorized**: Token expired - sign in again
- **429 Too Many Requests**: Rate limiting - reduce batch size
- **400 Bad Request**: Invalid JSON format - verify workflow files

### Browser Issues

- Disable ad blockers and privacy extensions
- Update to latest browser version
- Check browser console (F12) for detailed errors
- Ensure JavaScript is enabled

## Security

- Authentication handled directly with Microsoft
- No credentials stored locally or transmitted to third parties
- All operations performed client-side in browser
- Review JSON content before production deployment
- Use separate app registrations for test/production

## Support

Report issues or request features through GitHub Issues. Include:

- Browser and version
- Error messages from console
- Steps to reproduce
- Expected vs actual behavior
