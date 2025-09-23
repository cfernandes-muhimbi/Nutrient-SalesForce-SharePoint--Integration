# SharePoint Integration Architecture

## 🏗️ High-Level Architecture

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                                 SALESFORCE ORG                                  │
│                                                                                 │
│  ┌─────────────────┐    ┌──────────────────┐    ┌─────────────────────────────┐ │
│  │                 │    │                  │    │                             │ │
│  │   Lightning     │    │   Apex           │    │   Named Credentials         │ │
│  │   Web Component │◄──►│   Controller     │◄──►│   (OAuth Management)        │ │
│  │   (Frontend)    │    │   (Server-side)  │    │                             │ │
│  │                 │    │                  │    │                             │ │
│  └─────────────────┘    └──────────────────┘    └─────────────────────────────┘ │
│           │                       │                         │                   │
└───────────┼───────────────────────┼─────────────────────────┼───────────────────┘
            │                       │                         │
            │                       │                         │
            ▼                       ▼                         ▼
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────────────────────┐
│                 │    │                  │    │                                 │
│   Browser       │    │   Microsoft      │    │   SharePoint Online             │
│   (User Interface)   │   Graph API      │◄──►│   (Document Libraries)          │
│                 │    │   (gateway.ms.com) │   │                                 │
└─────────────────┘    └──────────────────┘    └─────────────────────────────────┘
```

## 🔐 Security Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                           SECURE DATA FLOW                          │
│                                                                     │
│  ┌──────────────┐  1. UI Actions   ┌──────────────────────────────┐ │
│  │              │─────────────────►│                              │ │
│  │   Browser    │                  │   Lightning Web Component    │ │
│  │   (Client)   │◄─────────────────│   (Frontend - No Secrets)   │ │
│  │              │  2. Visual Data  │                              │ │
│  └──────────────┘                  └──────────────────────────────┘ │
│                                                 │                   │
│                                      3. Method │                   │
│                                         Calls  │                   │
│                                                 ▼                   │
│  ┌──────────────┐  4. API Calls    ┌──────────────────────────────┐ │
│  │              │◄─────────────────│                              │ │
│  │  Microsoft   │                  │   Apex Controller            │ │
│  │  Graph API   │─────────────────►│   (Server-side Security)    │ │
│  │              │  5. JSON Data    │                              │ │
│  └──────────────┘                  └──────────────────────────────┘ │
│         ▲                                       │                   │
│         │                          6. OAuth     │                   │
│         │                             Token     │                   │
│         │                             Request   │                   │
│         │                                       ▼                   │
│  ┌──────────────┐                  ┌──────────────────────────────┐ │
│  │              │                  │                              │ │
│  │  SharePoint  │                  │   Named Credentials          │ │
│  │  Online      │                  │   (Secure Token Storage)     │ │
│  │              │                  │                              │ │
│  └──────────────┘                  └──────────────────────────────┘ │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

## 🧩 Component Architecture

### Lightning Web Component Structure

```
sharePointFileViewer/
├── sharePointFileViewer.html
│   ├── Site Selection Dropdown
│   ├── Drive ID Input
│   ├── Load Files Button
│   ├── Files Data Table
│   │   ├── Name Column (with links)
│   │   ├── Size Column
│   │   ├── Created By Column
│   │   ├── Created Date Column
│   │   ├── Modified Date Column
│   │   └── Version History Button Column
│   └── Version History Modal
│       ├── Version Data Table
│       ├── View/Download Buttons
│       └── Loading States
│
├── sharePointFileViewer.js
│   ├── Component Properties
│   │   ├── @track files = []
│   │   ├── @track sites = []
│   │   ├── @track selectedSiteId
│   │   ├── @track selectedDriveId
│   │   └── @track selectedFileVersions = []
│   ├── Data Table Columns
│   │   ├── columns[] (main files table)
│   │   └── versionColumns[] (version history table)
│   ├── Lifecycle Methods
│   │   ├── connectedCallback()
│   │   └── testConnectionFirst()
│   ├── Data Loading Methods
│   │   ├── loadSites()
│   │   ├── loadFiles()
│   │   └── viewVersionHistory()
│   ├── Event Handlers
│   │   ├── handleRowAction()
│   │   ├── handleVersionRowAction()
│   │   ├── handleSiteChange()
│   │   └── handleRefresh()
│   └── Utility Methods
│       ├── showToast()
│       ├── processSitesResult()
│       └── closeVersionHistory()
│
└── sharePointFileViewer.js-meta.xml
    ├── API Version: 58.0
    ├── isExposed: true
    └── Targets: AppPage, HomePage, RecordPage
```

### Apex Controller Structure

```
SharePointIntegrationController.cls
├── Connection Testing
│   └── testConnection()
│       ├── Endpoint: /me
│       ├── Purpose: Verify Graph API connectivity
│       └── Returns: Connection status
│
├── Site Management
│   ├── getSharePointSites()
│   │   ├── Endpoint: /sites/root
│   │   ├── Purpose: Get root SharePoint site
│   │   └── Returns: SharePointSite[]
│   └── getFollowedSites()
│       ├── Endpoint: /me/followedSites
│       ├── Purpose: Get user's followed sites
│       └── Returns: SharePointSite[]
│
├── File Operations
│   ├── getSharePointFiles()
│   │   ├── Endpoint: /sites/{id}/drives/{id}/root/children
│   │   ├── Purpose: List files in document library
│   │   └── Returns: SharePointFile[]
│   ├── getFileVersionHistory()
│   │   ├── Endpoint: /sites/{id}/drives/{id}/items/{id}/versions
│   │   ├── Purpose: Get file version history
│   │   └── Returns: SharePointFileVersion[]
│   └── getVersionDownloadUrl()
│       ├── Endpoint: /sites/{id}/drives/{id}/items/{id}/versions/{id}/content
│       ├── Purpose: Generate download URL for specific version
│       └── Returns: String (URL)
│
└── Data Classes
    ├── SharePointFile
    │   ├── String id
    │   ├── String name
    │   ├── String webUrl
    │   ├── Integer size
    │   ├── String createdBy
    │   ├── String createdDateTime
    │   └── String lastModifiedDateTime
    ├── SharePointSite
    │   ├── String id
    │   ├── String displayName
    │   └── String webUrl
    └── SharePointFileVersion
        ├── String id
        ├── String lastModifiedDateTime
        ├── Integer size
        └── String lastModifiedBy
```

## 🌊 Data Flow Sequence

### 1. Component Initialization
```
User loads page
    ↓
connectedCallback() executed
    ↓
testConnectionFirst() called
    ↓
testConnection() Apex method
    ↓
HTTP GET to /me endpoint
    ↓
Named Credentials provide OAuth token
    ↓
Microsoft Graph API responds
    ↓
Connection test result displayed
    ↓
loadSites() automatically called
```

### 2. Site Loading Process
```
loadSites() called
    ↓
getSharePointSites() Apex method
    ↓
HTTP GET to /sites/root endpoint
    ↓
If successful: process site data
    ↓
If failed: try getFollowedSites()
    ↓
HTTP GET to /me/followedSites endpoint
    ↓
Sites populated in dropdown
    ↓
First site auto-selected
    ↓
loadFiles() automatically called
```

### 3. File Loading Process
```
"Load Files" button clicked
    ↓
handleRefresh() event handler
    ↓
Files array cleared
    ↓
loadFiles() method called
    ↓
getSharePointFiles() Apex method
    ↓
HTTP GET to /sites/{id}/drives/{id}/root/children
    ↓
JSON response processed
    ↓
File objects created with metadata
    ↓
Data table populated with files
    ↓
"View Versions" buttons enabled
```

### 4. Version History Workflow
```
"View Versions" button clicked
    ↓
handleRowAction() event handler
    ↓
viewVersionHistory() method called
    ↓
getFileVersionHistory() Apex method
    ↓
HTTP GET to /sites/{id}/drives/{id}/items/{id}/versions
    ↓
Version data processed
    ↓
Modal dialog opened
    ↓
Version history table populated
    ↓
"View/Download" buttons available
```

### 5. File Version Access
```
"View/Download" button clicked
    ↓
handleVersionRowAction() event handler
    ↓
viewVersionFile() method called
    ↓
getVersionDownloadUrl() Apex method
    ↓
HTTP GET to /sites/{id}/drives/{id}/items/{id}/versions/{id}/content
    ↓
Microsoft returns download URL or redirect
    ↓
URL processed in Apex
    ↓
window.open() opens file in new tab
    ↓
File displayed or downloaded
```

## 🔧 API Integration Details

### Microsoft Graph API Endpoints

| Method | Endpoint | Purpose | Response |
|--------|----------|---------|----------|
| GET | `/me` | Test connectivity | User profile |
| GET | `/sites/root` | Get root site | Site object |
| GET | `/me/followedSites` | Get followed sites | Site collection |
| GET | `/sites/{id}/drives/{id}/root/children` | List files | File collection |
| GET | `/sites/{id}/drives/{id}/items/{id}/versions` | Version history | Version collection |
| GET | `/sites/{id}/drives/{id}/items/{id}/versions/{id}/content` | Download link | Redirect/URL |

### Authentication Flow

```
┌─────────────────┐
│ User Action     │
│ (Load Component)│
└─────┬───────────┘
      │
      ▼
┌─────────────────┐
│ Apex Method     │
│ Called          │
└─────┬───────────┘
      │
      ▼
┌─────────────────┐
│ Named Credential│
│ Provides Token  │
└─────┬───────────┘
      │
      ▼
┌─────────────────┐
│ HTTP Request    │
│ to Graph API    │
└─────┬───────────┘
      │
      ▼
┌─────────────────┐
│ Microsoft       │
│ Validates Token │
└─────┬───────────┘
      │
      ▼
┌─────────────────┐
│ SharePoint      │
│ Data Returned   │
└─────────────────┘
```

## 🔒 Security Implementation

### 1. No Secrets in Frontend
- All OAuth tokens stored in Salesforce Named Credentials
- No client secrets, API keys, or tokens in JavaScript
- Frontend only handles display data and user interactions

### 2. Server-Side Security
- All API calls made from Apex (server-side)
- Input validation and sanitization in Apex methods
- Proper exception handling without exposing internals
- Salesforce sharing rules enforced (`with sharing`)

### 3. Authentication Security
- OAuth 2.0 with Microsoft identity platform
- Tokens automatically refreshed by Named Credentials
- Secure token storage in Salesforce infrastructure
- No token management required by developers

### 4. Network Security
- All communications over HTTPS
- API endpoints validated before requests
- Response data sanitized before return
- No direct browser-to-SharePoint connections

## 📊 Performance Optimization

### Caching Strategy
```
Sites Data: Cacheable (rarely changes)
Files Data: Non-cacheable (frequently updated)
Version History: Non-cacheable (specific to user action)
Download URLs: Non-cacheable (temporary, secure URLs)
```

### Error Recovery
```
Connection Test → Sites Loading → File Loading
      ↓                ↓             ↓
   Retry Logic    Fallback to    Clear Error
   + User Alert   Followed Sites  + User Alert
```

### UI Performance
- Loading states prevent multiple API calls
- Table pagination for large file lists
- Modal dialogs for focused interactions
- Reactive data updates with proper change detection

This architecture ensures a secure, scalable, and maintainable SharePoint integration that follows Salesforce and security best practices.