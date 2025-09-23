# SharePoint Files Integration for Salesforce

A secure and comprehensive Lightning Web Component (LWC) that enables seamless integration between Salesforce and SharePoint Online, allowing users to browse, view, and access SharePoint files with complete version history directly within Salesforce.

## 🚀 Features

### Core Functionality
- **Browse SharePoint Sites**: Automatically discover and list available SharePoint sites
- **File Management**: View files from SharePoint Document Libraries
- **Version History**: Access complete version history for any file
- **File Viewing**: Open and view specific versions of files in browser/download
- **Real-time Data**: Fresh data with no caching for dynamic content

### User Experience
- **Responsive Design**: Mobile-friendly Lightning Design System (SLDS) UI
- **Loading States**: Clear visual feedback during API operations
- **Error Handling**: Comprehensive error messages and troubleshooting
- **Intuitive Navigation**: Easy-to-use dropdowns and buttons

## 🏗️ Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────────┐
│                 │    │                  │    │                     │
│   Salesforce    │    │   Salesforce     │    │   Microsoft Graph   │
│   Lightning     │◄──►│   Apex           │◄──►│   API               │
│   Web Component │    │   Controller     │    │   (SharePoint)      │
│                 │    │                  │    │                     │
└─────────────────┘    └──────────────────┘    └─────────────────────┘
         │                       │                         │
         │                       │                         │
         ▼                       ▼                         ▼
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────────┐
│                 │    │                  │    │                     │
│   Browser       │    │   Named          │    │   SharePoint        │
│   (Frontend)    │    │   Credentials    │    │   Online            │
│                 │    │   (OAuth)        │    │                     │
└─────────────────┘    └──────────────────┘    └─────────────────────┘
```

### Component Architecture

```
sharePointFileViewer (LWC)
├── sharePointFileViewer.html (UI Template)
├── sharePointFileViewer.js (Controller Logic)
├── sharePointFileViewer.js-meta.xml (Metadata)
└── SharePointIntegrationController.cls (Apex Backend)
```

## 🔧 Technical Stack

### Frontend
- **Lightning Web Components (LWC)**: Modern web standards-based framework
- **Lightning Design System (SLDS)**: Salesforce's design system
- **JavaScript ES6+**: Modern JavaScript features

### Backend
- **Apex Classes**: Salesforce server-side logic
- **Named Credentials**: Secure OAuth 2.0 authentication
- **Microsoft Graph API**: RESTful API for SharePoint integration

### Authentication
- **OAuth 2.0**: Industry-standard authentication protocol
- **Azure App Registration**: Microsoft identity platform integration
- **Salesforce Auth Provider**: Managed authentication flow

## 📦 Components

### 1. Custom Metadata Type (`Graph_SharePoint_Config__mdt`)

**Purpose:** Centralized configuration management for SharePoint integration

**Fields:**
- `DriveId__c`: SharePoint Drive ID for document library
- `SiteId__c`: SharePoint Site ID  
- `TenantId__c`: Microsoft 365 Tenant ID
- `IsActive__c`: Indicates if configuration is active
- `Description__c`: Configuration description

### 2. Apex Controller (`SharePointIntegrationController.cls`)

**Configuration Methods:**
- `getActiveConfig()`: Get active SharePoint configuration
- `getSharePointConfig()`: Return configuration for LWC
- `getAvailableConfigurations()`: Get all active configurations

**API Methods:**
- `testConnection()`: Verify Microsoft Graph API connectivity
- `getSharePointSites()`: Retrieve available SharePoint sites
- `getFollowedSites()`: Get user's followed SharePoint sites
- `getSharePointFiles()`: Fetch files from SharePoint document library
- `getFileVersionHistory()`: Get version history for specific files
- `getVersionDownloadUrl()`: Generate secure download URLs for file versions

**Classes:**
- `SharePointFile`: Data structure for file information
- `SharePointSite`: Data structure for site information
- `SharePointFileVersion`: Data structure for version information
- `ConfigOption`: Data structure for configuration options

### 3. Lightning Web Component (`sharePointFileViewer`)

**Features:**
- **Configuration dropdown** - Select from available configurations
- **Site selection dropdown** - Choose SharePoint sites
- **Drive ID input field** - Override drive ID if needed
- **File data table** with sorting and filtering
- **Version history modal** dialog
- **Action buttons** for file operations
- **Loading states** and error handling
- **Automatic configuration loading** from Custom Metadata

##  Security Analysis

###  **SECURE IMPLEMENTATION**

**No Secrets Exposed:**
- All authentication tokens managed server-side
- No client secrets or sensitive data in frontend code
- OAuth flows handled by Salesforce Named Credentials

**Access Control:**
- Server-side authorization through Apex
- User permissions respected via SharePoint API
- Secure API endpoint calls with proper headers

**Data Protection:**
- No sensitive data cached in browser
- Secure HTTPS communication
- Input validation and sanitization

**Best Practices Followed:**
- `@AuraEnabled` methods with appropriate caching
- Exception handling without exposing internal details
- Proper use of Salesforce sharing rules (`with sharing`)
- No hardcoded credentials or endpoints

## 🛠️ Setup Instructions

### Prerequisites
- Salesforce org with Lightning Experience enabled
- Azure Active Directory tenant
- SharePoint Online access
- System Administrator permissions
- Custom Metadata Type management permissions
- SharePoint Drive ID and Site ID (see steps below to obtain)

### Step 1: Obtain SharePoint Configuration Details

Before setting up the integration, you'll need to collect these SharePoint details for your Custom Metadata configuration:

1. **Get SharePoint Site ID**:
   - Navigate to your SharePoint site in browser
   - Use Microsoft Graph Explorer: `https://developer.microsoft.com/en-us/graph/graph-explorer`
   - Query: `GET https://graph.microsoft.com/v1.0/sites/{hostname}:{path}`
   - Example: `GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/teamsite`
   - Note the `id` field from the response

2. **Get SharePoint Drive ID**:
   - Using Graph Explorer, query: `GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives`
   - Find your document library in the results
   - Note the `id` field for your target document library

3. **Get Tenant ID**:
   - Azure Portal → Azure Active Directory → Properties
   - Copy the "Tenant ID" value

📋 **Keep these values ready** - you'll need them for Custom Metadata configuration in Step 4.

### Step 2: Azure App Registration

1. **Create App Registration**:
   ```
   Name: SharePoint-Salesforce-Integration
   Supported account types: Single tenant
   Redirect URI: https://your-salesforce-domain.my.salesforce.com/services/authcallback/MicrosoftApp
   ```

2. **Configure API Permissions**:
   ```
   Microsoft Graph:
   - Files.Read.All (Application)
   - Sites.Read.All (Application)
   - User.Read (Delegated)
   ```

3. **Generate Client Secret**:
   - Note down: Application ID and Client Secret

📖 **Official Salesforce Guide**: [Create a Microsoft Azure App](https://help.salesforce.com/s/articleView?id=ind.sf_contracts_create_a_microsoft_azure_app.htm&language=en_US&type=5)

### Step 3: Salesforce Configuration

1. **Create Auth Provider**:
   ```
   Provider Type: Open ID Connect
   Name: MicrosoftApp
   Consumer Key: [Your Azure App ID]
   Consumer Secret: [Your Azure Client Secret]
   Authorize Endpoint URL: https://login.microsoftonline.com/[tenant-id]/oauth2/v2.0/authorize
   Token Endpoint URL: https://login.microsoftonline.com/[tenant-id]/oauth2/v2.0/token
   Default Scopes: https://graph.microsoft.com/Files.Read.All https://graph.microsoft.com/Sites.Read.All
   ```

   📖 **Official Salesforce Guide**: [Configure an Auth Provider for Microsoft App](https://help.salesforce.com/s/articleView?id=ind.sf_contracts_configure_an_auth_provider_for_microsoft_app.htm&type=5)

2. **Create Named Credential**:
   ```
   Label: MicrosoftApp
   Name: MicrosoftApp
   URL: https://graph.microsoft.com/v1.0
   Identity Type: Named Principal
   Authentication Protocol: OAuth 2.0
   Authentication Provider: MicrosoftApp
   Generate Authorization Header: ✓
   ```

   📖 **Official Salesforce Guide**: [Create a Named Credential for Microsoft App](https://help.salesforce.com/s/articleView?id=ind.sf_contracts_create_a_named_credential_for_sf_connector_for_word_add-in_app.htm&type=5)

### Step 4: Configure Custom Metadata

1. **Deploy Custom Metadata Type**:
   ```bash
   sf project deploy start --source-dir force-app/main/default/objects
   ```

2. **Deploy Custom Metadata Records**:
   ```bash
   sf project deploy start --source-dir force-app/main/default/customMetadata
   ```

3. **Update Configuration** (via Setup > Custom Metadata Types):
   - Navigate to Setup > Custom Metadata Types
   - Click "Manage Records" next to "Graph SharePoint Config"
   - Edit the "Default" record with your values **collected in Step 1**:
     ```
     Drive Id: [Your SharePoint Drive ID from Step 1]
     Site Id: [Your SharePoint Site ID from Step 1] 
     Tenant Id: [Your Microsoft 365 Tenant ID from Step 1]
     Is Active: ✓ (checked)
     Description: [Your description, e.g., "Production SharePoint Library"]
     ```

   💡 **Pro Tip**: You can create multiple configuration records for different environments:
   - "Production" (IsActive: ✓)
   - "Sandbox" (IsActive: ✗) 
   - "Testing" (IsActive: ✗)

### Step 5: Deploy Components

1. **Deploy Apex Class**:
   ```bash
   sf project deploy start --source-dir force-app/main/default/classes
   ```

2. **Deploy LWC**:
   ```bash
   sf project deploy start --source-dir force-app/main/default/lwc
   ```

### Step 6: Configure Component

1. **Add to Lightning Page**:
   - Edit Lightning App Page
   - Add `sharePointFileViewer` component
   - Save and activate

## 🎯 Usage

### Basic Operations

1. **Browse Files**:
   - Component loads and tests connection automatically
   - Sites populated in dropdown
   - Click "Load Files" to view documents

2. **View Version History**:
   - Click "View Versions" button in any file row
   - Modal opens with complete version history
   - See who modified, when, and file sizes

3. **Access File Versions**:
   - In version history modal, click "View/Download" for any version
   - File opens in new browser tab or downloads
   - Works with all file types (Office docs, PDFs, images, etc.)

### Advanced Features

- **Configuration Management**: Switch between multiple SharePoint environments
- **Site Selection**: Switch between different SharePoint sites
- **Drive ID Override**: Manually specify different document libraries
- **Multi-Environment Support**: Different configurations for Dev/Test/Prod
- **Error Recovery**: Automatic fallback between different API endpoints
- **Real-time Updates**: Fresh data on each load operation

##  Troubleshooting

### Common Issues

**No Sites Visible:**
```
Check: Auth Provider configuration
Solution: Verify OAuth scopes and permissions
```

**Files Not Loading:**
```
Check: Drive ID accuracy
Solution: Use SharePoint admin to find correct drive ID
```

**Version History Empty:**
```
Check: File versioning enabled in SharePoint
Solution: Enable versioning in document library settings
```

**Connection Test Failed:**
```
Check: Named Credentials authentication status
Solution: Re-authenticate through Named Credentials
```

**Configuration Error:**
```
Check: Custom Metadata Type deployment
Solution: Deploy Graph_SharePoint_Config__mdt and create Default record
```

**No Configurations Available:**
```
Check: Custom Metadata records with IsActive = true
Solution: Create/activate configuration records in Setup
```

**Invalid SharePoint IDs:**
```
Check: Drive ID and Site ID format and validity
Solution: Re-obtain IDs using Microsoft Graph Explorer (Step 1)
```

**Component Shows "Configuration not found":**
```
Check: 
1. Custom Metadata Type deployed successfully
2. At least one record with IsActive = true exists
3. Record has valid DriveId__c value
Solution: Follow Step 4 deployment and configuration steps
```

### Debug Console Logs

Enable browser developer tools to see detailed logs:
```javascript
SharePointFileViewer - Component connected
SharePointFileViewer - Configuration loaded: {DriveId__c: "b!...", MasterLabel: "Default"}
SharePointFileViewer - Using DriveId from config: b!FGWYM5Wq60WB-ZdllAPDhj_JgW4OtCxMsIwvWJ_ZCeAt20_iedKPTpFNEUHjPjMr
SharePointFileViewer - Configuration test result: SUCCESS: Using configuration "Default"
SharePointFileViewer - Connection test result: Status: 200
SharePointFileViewer - Sites loaded successfully
SharePointFileViewer - Files loaded successfully
```

**Configuration Success Indicators:**
- ✅ "Configuration loaded" with actual config object
- ✅ "Using DriveId from config" with your specific Drive ID
- ✅ "Configuration test result: SUCCESS"

## API Endpoints Used

### Microsoft Graph API

| Endpoint | Purpose | Method |
|----------|---------|---------|
| `/me` | Test connection | GET |
| `/sites/root` | Get root SharePoint site | GET |
| `/me/followedSites` | Get followed sites | GET |
| `/sites/{id}/drives/{id}/root/children` | List files | GET |
| `/sites/{id}/drives/{id}/items/{id}/versions` | Version history | GET |
| `/sites/{id}/drives/{id}/items/{id}/versions/{id}/content` | File download | GET |

## Performance Considerations

- **Caching Strategy**: Sites cached, files always fresh
- **Pagination**: Large file lists handled efficiently
- **Error Recovery**: Multiple fallback mechanisms
- **Network Optimization**: Minimal API calls, batch operations where possible

## Contributing

### Development Setup

1. Clone repository
2. Configure SFDX project
3. Set up Azure app registration
4. Configure Named Credentials
5. Deploy and test

### Code Standards

- Follow Salesforce coding conventions
- Use meaningful variable names
- Include comprehensive error handling
- Add debug logging for troubleshooting
- Document all public methods

##  License


## Support

For support and questions:
- Check troubleshooting section
- Review Salesforce debug logs
- Verify Azure app registration configuration
- Test Named Credentials authentication

## Version History

### v1.0.0
- Initial release with basic file browsing
- SharePoint site discovery
- File listing with metadata

### v1.1.0
- Added version history functionality
- File viewing/downloading capabilities
- Enhanced error handling and user feedback

## 🎯 Future Enhancements

- File upload capabilities
- Advanced search and filtering
- Bulk operations
- Integration with Salesforce records
- Mobile app compatibility
- Performance optimizations

## 📚 Official References

### Salesforce Documentation
- [Create a Microsoft Azure App](https://help.salesforce.com/s/articleView?id=ind.sf_contracts_create_a_microsoft_azure_app.htm&language=en_US&type=5) - Complete guide for Azure App Registration setup
- [Configure an Auth Provider for Microsoft App](https://help.salesforce.com/s/articleView?id=ind.sf_contracts_configure_an_auth_provider_for_microsoft_app.htm&type=5) - Step-by-step Auth Provider configuration
- [Create a Named Credential for Microsoft App](https://help.salesforce.com/s/articleView?id=ind.sf_contracts_create_a_named_credential_for_sf_connector_for_word_add-in_app.htm&type=5) - Named Credentials setup guide

### Microsoft Documentation
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/) - Complete Microsoft Graph API documentation
- [SharePoint REST API](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest) - SharePoint file operations
- [Azure App Registrations](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app) - Azure AD app setup

### Salesforce Developer Resources
- [Lightning Web Components Developer Guide](https://developer.salesforce.com/docs/component-library/documentation/en/lwc) - LWC development best practices
- [Apex Developer Guide](https://developer.salesforce.com/docs/atlas.en-us.apexcode.meta/apexcode/) - Apex development guidelines
- [Named Credentials](https://help.salesforce.com/s/articleView?id=sf.named_credentials_about.htm&type=5) - Understanding Named Credentials

---

**Built with ❤️ for secure SharePoint-Salesforce integration**