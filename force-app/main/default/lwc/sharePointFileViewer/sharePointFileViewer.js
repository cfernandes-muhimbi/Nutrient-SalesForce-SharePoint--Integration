import { LightningElement, track } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import getSharePointFiles from '@salesforce/apex/SharePointIntegrationController.getSharePointFiles';
import getSharePointSites from '@salesforce/apex/SharePointIntegrationController.getSharePointSites';
import getFollowedSites from '@salesforce/apex/SharePointIntegrationController.getFollowedSites';
import testConnection from '@salesforce/apex/SharePointIntegrationController.testConnection';
import testConfigurationUsage from '@salesforce/apex/SharePointIntegrationController.testConfigurationUsage';
import getFileVersionHistory from '@salesforce/apex/SharePointIntegrationController.getFileVersionHistory';
import getVersionDownloadUrl from '@salesforce/apex/SharePointIntegrationController.getVersionDownloadUrl';
import getSharePointConfig from '@salesforce/apex/SharePointIntegrationController.getSharePointConfig';
import getAvailableConfigurations from '@salesforce/apex/SharePointIntegrationController.getAvailableConfigurations';

export default class SharePointFileViewer extends LightningElement {
    @track files = [];
    @track sites = [];
    @track selectedSiteId = '';
    @track selectedDriveId = '';
    @track selectedConfig = '';
    @track availableConfigs = [];
    @track currentConfig = null;
    @track isLoading = false;
    @track error;
    @track tableKey = 0;
    @track selectedFileVersions = [];
    @track showVersionHistory = false;
    @track selectedFileName = '';
    @track selectedFileId = '';
    @track isDownloading = false;

    columns = [
        { 
            label: 'Name', 
            fieldName: 'webUrl', 
            type: 'url',
            typeAttributes: {
                label: { fieldName: 'name' },
                target: '_blank'
            }
        },
        { 
            label: 'Size', 
            fieldName: 'size', 
            type: 'number',
            cellAttributes: { alignment: 'left' }
        },
        { 
            label: 'Created By', 
            fieldName: 'createdBy', 
            type: 'text' 
        },
        { 
            label: 'Created Date', 
            fieldName: 'createdDateTime', 
            type: 'date',
            typeAttributes: {
                year: 'numeric',
                month: 'short',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit'
            }
        },
        { 
            label: 'Modified Date', 
            fieldName: 'lastModifiedDateTime', 
            type: 'date',
            typeAttributes: {
                year: 'numeric',
                month: 'short',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit'
            }
        },
        {
            label: 'Version History',
            type: 'button',
            typeAttributes: {
                label: 'View Versions',
                name: 'view_versions',
                variant: 'brand'
            },
            cellAttributes: { alignment: 'center' }
        },
        {
            label: 'Edit PDF',
            type: 'button',
            typeAttributes: {
                label: 'Edit in SharePoint',
                name: 'edit_pdf',
                variant: 'success',
                disabled: { fieldName: 'isNotPdf' }
            },
            cellAttributes: { alignment: 'center' }
        }
    ];

    versionColumns = [
        { 
            label: 'Version ID', 
            fieldName: 'id', 
            type: 'text' 
        },
        { 
            label: 'Size', 
            fieldName: 'size', 
            type: 'number',
            cellAttributes: { alignment: 'left' }
        },
        { 
            label: 'Modified By', 
            fieldName: 'lastModifiedBy', 
            type: 'text' 
        },
        { 
            label: 'Modified Date', 
            fieldName: 'lastModifiedDateTime', 
            type: 'date',
            typeAttributes: {
                year: 'numeric',
                month: 'short',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit'
            }
        },
        {
            label: 'Actions',
            type: 'button',
            typeAttributes: {
                label: 'View/Download',
                name: 'view_version',
                variant: 'brand',
                disabled: { fieldName: 'isLatestVersion' }
            },
            cellAttributes: { alignment: 'center' }
        }
    ];

    connectedCallback() {
        console.log('SharePointFileViewer - Component connected');
        this.loadConfiguration();
    }

    loadConfiguration() {
        console.log('SharePointFileViewer - Loading configuration');
        this.isLoading = true;
        
        Promise.all([
            getSharePointConfig(),
            getAvailableConfigurations()
        ])
        .then(results => {
            const [config, availableConfigs] = results;
            console.log('SharePointFileViewer - Configuration loaded:', config);
            console.log('SharePointFileViewer - Available configurations:', availableConfigs);
            
            // Validate configuration
            if (!config || !config.DriveId__c) {
                throw new Error('Invalid configuration: Missing DriveId');
            }
            
            this.currentConfig = config;
            this.availableConfigs = availableConfigs;
            this.selectedDriveId = config.DriveId__c;
            this.selectedSiteId = config.SiteId__c || '';  // SiteId is optional
            this.selectedConfig = config.MasterLabel;
            
            console.log('SharePointFileViewer - Using DriveId from config:', this.selectedDriveId);
            console.log('SharePointFileViewer - Using SiteId from config:', this.selectedSiteId);
            
            // Now test connection and load sites
            this.testConnectionFirst();
        })
        .catch(error => {
            console.error('SharePointFileViewer - Configuration load failed:', error);
            console.error('SharePointFileViewer - Error details:', error.body);
            this.error = 'Configuration load failed: ' + (error.body ? error.body.message : error.message);
            this.showToast('Configuration Error', this.error, 'error');
            
            // STOP HERE - Don't continue if configuration fails
            this.isLoading = false;
            return;
        })
        .finally(() => {
            this.isLoading = false;
        });
    }

    testConnectionFirst() {
        console.log('SharePointFileViewer - Testing configuration and connection');
        
        // First test if configuration is actually being used
        testConfigurationUsage()
            .then(configResult => {
                console.log('SharePointFileViewer - Configuration test result:', configResult);
                
                // Now test connection
                return testConnection();
            })
            .then(connectionResult => {
                console.log('SharePointFileViewer - Connection test result:', connectionResult);
                this.loadSites();
            })
            .catch(error => {
                console.error('SharePointFileViewer - Test failed:', error);
                this.error = 'Test failed: ' + (error.body ? error.body.message : error.message);
                this.showToast('Test Error', this.error, 'error');
            });
    }

    loadSites() {
        console.log('SharePointFileViewer - loadSites called - trying root site first');
        this.isLoading = true;
        
        // Try root site first
        getSharePointSites()
            .then(result => {
                console.log('SharePointFileViewer - Root site loaded successfully:', result);
                this.processSitesResult(result);
            })
            .catch(error => {
                console.error('SharePointFileViewer - Root site failed, trying followed sites:', error);
                // If root site fails, try followed sites
                return getFollowedSites();
            })
            .then(result => {
                if (result) {
                    console.log('SharePointFileViewer - Followed sites loaded successfully:', result);
                    this.processSitesResult(result);
                }
            })
            .catch(error => {
                console.error('SharePointFileViewer - Both site methods failed:', error);
                console.error('SharePointFileViewer - Error body:', error.body);
                console.error('SharePointFileViewer - Error message:', error.message);
                this.error = error.body ? error.body.message : error.message;
                this.showToast('Error', 'Failed to load SharePoint sites: ' + this.error, 'error');
            })
            .finally(() => {
                console.log('SharePointFileViewer - loadSites finally block');
                this.isLoading = false;
            });
    }

    processSitesResult(result) {
        if (result && result.length > 0) {
            this.sites = result.map(site => ({
                label: site.displayName || site.webUrl,
                value: site.id,
                webUrl: site.webUrl
            }));
            console.log('SharePointFileViewer - Mapped sites:', this.sites);
            if (this.sites.length > 0) {
                this.selectedSiteId = this.sites[0].value;
                console.log('SharePointFileViewer - Selected site ID:', this.selectedSiteId);
                this.loadFiles();
            }
            this.error = undefined;
        } else {
            console.log('SharePointFileViewer - No sites found in result');
            this.sites = [];
        }
    }

    loadFiles() {
        console.log('SharePointFileViewer - loadFiles called');
        console.log('SharePointFileViewer - selectedSiteId:', this.selectedSiteId);
        console.log('SharePointFileViewer - selectedDriveId:', this.selectedDriveId);
        
        if (!this.selectedSiteId || !this.selectedDriveId) {
            console.log('SharePointFileViewer - Missing siteId or driveId, returning');
            return;
        }

        this.isLoading = true;
        console.log('SharePointFileViewer - Calling getSharePointFiles with params:', { 
            siteId: this.selectedSiteId, 
            driveId: this.selectedDriveId 
        });
        
        getSharePointFiles({ siteId: this.selectedSiteId, driveId: this.selectedDriveId })
            .then(result => {
                console.log('SharePointFileViewer - Files loaded successfully:', result);
                
                if (result && result.length > 0) {
                    // Process the files
                    const processedFiles = result.map(file => ({
                        ...file,
                        createdDateTime: file.createdDateTime ? new Date(file.createdDateTime) : null,
                        lastModifiedDateTime: file.lastModifiedDateTime ? new Date(file.lastModifiedDateTime) : null,
                        isNotPdf: !file.name.toLowerCase().endsWith('.pdf')
                    }));
                    
                    // Force reactive update
                    this.files = [...processedFiles];
                    this.tableKey = this.tableKey + 1; // Force table re-render
                    console.log('SharePointFileViewer - Processed files:', this.files);
                    console.log('SharePointFileViewer - Files array length:', this.files.length);
                    console.log('SharePointFileViewer - Table columns:', this.columns);
                    this.error = undefined;
                } else {
                    console.log('SharePointFileViewer - No files returned from API');
                    this.files = [];
                }
            })
            .catch(error => {
                console.error('SharePointFileViewer - Error loading files:', error);
                this.error = error.body ? error.body.message : error.message;
                this.files = [];
                this.showToast('Error', 'Failed to load SharePoint files: ' + this.error, 'error');
            })
            .finally(() => {
                console.log('SharePointFileViewer - loadFiles finally block');
                this.isLoading = false;
            });
    }

    handleSiteChange(event) {
        console.log('SharePointFileViewer - handleSiteChange called with:', event.detail.value);
        this.selectedSiteId = event.detail.value;
        this.loadFiles();
    }

    handleDriveIdChange(event) {
        console.log('SharePointFileViewer - handleDriveIdChange called with:', event.detail.value);
        this.selectedDriveId = event.detail.value;
    }

    handleConfigChange(event) {
        console.log('SharePointFileViewer - handleConfigChange called with:', event.detail.value);
        const selectedConfigName = event.detail.value;
        this.selectedConfig = selectedConfigName;
        
        // Find the selected configuration
        const selectedConfig = this.availableConfigs.find(config => config.value === selectedConfigName);
        if (selectedConfig) {
            console.log('SharePointFileViewer - Switching to configuration:', selectedConfig);
            this.selectedDriveId = selectedConfig.driveId;
            this.selectedSiteId = selectedConfig.siteId;
            
            // Clear existing data and reload
            this.files = [];
            this.sites = [];
            this.error = undefined;
            this.tableKey = this.tableKey + 1;
            
            // Reload with new configuration
            this.loadSites();
        }
    }

    handleRefresh() {
        console.log('SharePointFileViewer - handleRefresh called (Load Files button clicked)');
        console.log('SharePointFileViewer - Current selectedSiteId:', this.selectedSiteId);
        console.log('SharePointFileViewer - Current selectedDriveId:', this.selectedDriveId);
        
        // Clear existing files first
        this.files = [];
        this.error = undefined;
        this.tableKey = this.tableKey + 1;
        
        // Force reload
        this.loadFiles();
    }

    showToast(title, message, variant) {
        const evt = new ShowToastEvent({
            title: title,
            message: message,
            variant: variant
        });
        this.dispatchEvent(evt);
    }

    // Handle row actions (version history button)
    handleRowAction(event) {
        console.log('SharePointFileViewer - handleRowAction called with event:', event);
        const actionName = event.detail.action.name;
        const row = event.detail.row;
        
        console.log('SharePointFileViewer - Row action:', actionName, 'for file:', row.name);
        console.log('SharePointFileViewer - Full event detail:', event.detail);
        
        switch (actionName) {
            case 'view_versions':
                this.viewVersionHistory(row);
                break;
            case 'edit_pdf':
                this.editPdfInSharePoint(row);
                break;
            default:
                break;
        }
    }

    // View version history for a file
    viewVersionHistory(file) {
        console.log('SharePointFileViewer - Loading version history for file:', file.name);
        
        this.selectedFileName = file.name;
        this.selectedFileId = file.id;
        this.isLoading = true;
        this.showVersionHistory = false;
        
        getFileVersionHistory({ 
            siteId: this.selectedSiteId, 
            driveId: this.selectedDriveId, 
            fileId: file.id 
        })
        .then(result => {
            console.log('SharePointFileViewer - Version history loaded:', result);
            
            if (result && result.length > 0) {
                this.selectedFileVersions = result.map((version, index) => ({
                    ...version,
                    lastModifiedDateTime: version.lastModifiedDateTime ? new Date(version.lastModifiedDateTime) : null,
                    isLatestVersion: index === 0  // First version is the latest
                }));
            } else {
                this.selectedFileVersions = [];
            }
            
            this.showVersionHistory = true;
            console.log('SharePointFileViewer - Processed versions:', this.selectedFileVersions);
        })
        .catch(error => {
            console.error('SharePointFileViewer - Error loading version history:', error);
            this.error = error.body ? error.body.message : error.message;
            this.showToast('Error', 'Failed to load version history: ' + this.error, 'error');
            this.selectedFileVersions = [];
        })
        .finally(() => {
            this.isLoading = false;
        });
    }

    // Close version history modal
    closeVersionHistory() {
        this.showVersionHistory = false;
        this.selectedFileVersions = [];
        this.selectedFileName = '';
        this.selectedFileId = '';
    }

    // Handle version row actions (view/download specific versions)
    handleVersionRowAction(event) {
        const actionName = event.detail.action.name;
        const versionRow = event.detail.row;
        
        console.log('SharePointFileViewer - Version action:', actionName, 'for version:', versionRow.id);
        
        switch (actionName) {
            case 'view_version':
                this.viewVersionFile(versionRow);
                break;
            default:
                break;
        }
    }

    // View/Download specific version of a file
    viewVersionFile(version) {
        console.log('SharePointFileViewer - Viewing version:', version.id, 'of file:', this.selectedFileName);
        
        this.isDownloading = true;
        
        getVersionDownloadUrl({ 
            siteId: this.selectedSiteId, 
            driveId: this.selectedDriveId, 
            fileId: this.selectedFileId,
            versionId: version.id
        })
        .then(downloadUrl => {
            console.log('SharePointFileViewer - Download URL received:', downloadUrl);
            
            if (downloadUrl) {
                // Open the file in a new tab/window
                window.open(downloadUrl, '_blank');
                this.showToast('Success', 'Opening version ' + version.id + ' of ' + this.selectedFileName, 'success');
            } else {
                throw new Error('No download URL received');
            }
        })
        .catch(error => {
            console.error('SharePointFileViewer - Error getting download URL:', error);
            this.error = error.body ? error.body.message : error.message;
            this.showToast('Error', 'Failed to open version: ' + this.error, 'error');
        })
        .finally(() => {
            this.isDownloading = false;
        });
    }

    // Edit PDF in SharePoint using Muhimbi PDF Editor
    editPdfInSharePoint(file) {
        console.log('SharePointFileViewer - Edit PDF in SharePoint:', file.name);
        console.log('SharePointFileViewer - File object:', JSON.stringify(file));
        
        try {
            // Use the file's webUrl directly to construct parameters
            let fileWebUrl = file.webUrl;
            console.log('SharePointFileViewer - Original webUrl:', fileWebUrl);
            
            // Parse the URL to extract components
            const fileUrl = new URL(fileWebUrl);
            
            // Site collection URL: protocol + hostname (e.g., https://tenant.sharepoint.com)
            const siteCollectionUrl = `${fileUrl.protocol}//${fileUrl.hostname}`;
            
            // Relative path: the pathname (e.g., /sites/team/Documents/file.pdf)  
            const relativePath = decodeURIComponent(fileUrl.pathname);
            
            console.log('SharePointFileViewer - Site Collection URL:', siteCollectionUrl);
            console.log('SharePointFileViewer - Relative Path:', relativePath);
            
            // Base64 encode both parameters (remove padding as per Muhimbi requirements)
            const encodedSiteUrl = btoa(siteCollectionUrl).replace(/=+$/, '');
            const encodedRelativePath = btoa(relativePath).replace(/=+$/, '');
            
            console.log('SharePointFileViewer - Encoded Site URL (s):', encodedSiteUrl);
            console.log('SharePointFileViewer - Encoded Relative Path (f):', encodedRelativePath);
            
            // Construct the direct Muhimbi PDF Editor URL matching the expected format
            const muhimbiUrl = `https://pdf-editor.muhimbi.com/pdf/edit?s=${encodedSiteUrl}&f=${encodedRelativePath}`;
            
            console.log('SharePointFileViewer - Final Muhimbi URL:', muhimbiUrl);
            console.log('SharePointFileViewer - Expected format: https://pdf-editor.muhimbi.com/pdf/edit?s=aHR0cHM6Ly90ZW5hbnQuc2hhcmVwb2ludC5jb20&f=L3NpdGVzL3RlYW0vRG9jdW1lbnRzL3JlcG9ydC5wZGY');
            
            // Test the Base64 encoding by decoding back
            console.log('SharePointFileViewer - Decoded site URL:', atob(encodedSiteUrl));
            console.log('SharePointFileViewer - Decoded relative path:', atob(encodedRelativePath));
            
            // Open the Muhimbi URL directly in a new tab
            try {
                window.open(muhimbiUrl, '_blank');
                this.showToast('Success', 'Opening PDF editor for ' + file.name, 'success');
            } catch (popupError) {
                console.error('SharePointFileViewer - Popup blocked, trying direct navigation:', popupError);
                // Fallback: direct navigation if popup is blocked
                window.location.href = muhimbiUrl;
            }
            
        } catch (error) {
            console.error('SharePointFileViewer - Error constructing PDF editor URL:', error);
            this.showToast('Error', 'Failed to open PDF editor: ' + error.message, 'error');
        }
    }

}