public with sharing class SharePointIntegrationController {
    
    // Configuration management
    private static Graph_SharePoint_Config__mdt getActiveConfig() {
        try {
            Graph_SharePoint_Config__mdt config = [
                SELECT Id, MasterLabel, DriveId__c, SiteId__c, TenantId__c, IsActive__c, Description__c
                FROM Graph_SharePoint_Config__mdt 
                WHERE IsActive__c = true 
                LIMIT 1
            ];
            return config;
        } catch (Exception e) {
            System.debug('No active SharePoint configuration found: ' + e.getMessage());
            throw new AuraHandledException('SharePoint configuration not found. Please configure Graph SharePoint Config custom metadata.');
        }
    }
    
    @AuraEnabled(cacheable=false)
    public static Graph_SharePoint_Config__mdt getSharePointConfig() {
        try {
            return getActiveConfig();
        } catch (Exception e) {
            throw new AuraHandledException('Error loading SharePoint configuration: ' + e.getMessage());
        }
    }
    
    @AuraEnabled(cacheable=false)
    public static List<ConfigOption> getAvailableConfigurations() {
        try {
            List<ConfigOption> options = new List<ConfigOption>();
            List<Graph_SharePoint_Config__mdt> configs = [
                SELECT Id, MasterLabel, DriveId__c, SiteId__c, TenantId__c, IsActive__c, Description__c
                FROM Graph_SharePoint_Config__mdt 
                WHERE IsActive__c = true
                ORDER BY MasterLabel
            ];
            
            for (Graph_SharePoint_Config__mdt config : configs) {
                ConfigOption option = new ConfigOption();
                option.label = config.MasterLabel;
                option.value = config.MasterLabel;
                option.description = config.Description__c;
                option.driveId = config.DriveId__c;
                option.siteId = config.SiteId__c;
                options.add(option);
            }
            
            return options;
        } catch (Exception e) {
            throw new AuraHandledException('Error loading configurations: ' + e.getMessage());
        }
    }
    
    // Test method to verify configuration is working
    @AuraEnabled(cacheable=false)
    public static String testConfigurationUsage() {
        try {
            Graph_SharePoint_Config__mdt config = getActiveConfig();
            return 'SUCCESS: Using configuration "' + config.MasterLabel + '" with DriveId: ' + config.DriveId__c;
        } catch (Exception e) {
            return 'FAILED: ' + e.getMessage();
        }
    }

    // Test method to verify connection
    @AuraEnabled(cacheable=false)
    public static String testConnection() {
        try {
            String endpoint = 'callout:MicrosoftApp/me';
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Content-Type', 'application/json');
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            System.debug('Test Connection Status: ' + res.getStatusCode());
            System.debug('Test Connection Response: ' + res.getBody());
            
            return 'Status: ' + res.getStatusCode() + ' - Response: ' + res.getBody();
        } catch (Exception e) {
            System.debug('Test Connection Exception: ' + e.getMessage());
            return 'Error: ' + e.getMessage();
        }
    }
    
    @AuraEnabled(cacheable=false)
    public static List<SharePointFile> getSharePointFiles(String siteId, String driveId) {
        List<SharePointFile> files = new List<SharePointFile>();
        
        try {
            // Microsoft Graph API endpoint to get files from SharePoint drive
            String endpoint = 'callout:MicrosoftApp/sites/' + siteId + '/drives/' + driveId + '/root/children';
            
            System.debug('SharePoint Files Endpoint: ' + endpoint);
            System.debug('Site ID: ' + siteId + ', Drive ID: ' + driveId);
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Content-Type', 'application/json');
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            System.debug('SharePoint Files Response Status: ' + res.getStatusCode());
            System.debug('SharePoint Files Response Body: ' + res.getBody());
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> responseMap = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
                List<Object> value = (List<Object>) responseMap.get('value');
                
                if (value != null && !value.isEmpty()) {
                    for (Object fileObj : value) {
                        Map<String, Object> fileMap = (Map<String, Object>) fileObj;
                        SharePointFile file = new SharePointFile();
                        file.id = (String) fileMap.get('id');
                        file.name = (String) fileMap.get('name');
                        file.webUrl = (String) fileMap.get('webUrl');
                        file.size = (Integer) fileMap.get('size');
                        
                        Map<String, Object> createdBy = (Map<String, Object>) fileMap.get('createdBy');
                        if (createdBy != null) {
                            Map<String, Object> user = (Map<String, Object>) createdBy.get('user');
                            if (user != null) {
                                file.createdBy = (String) user.get('displayName');
                            }
                        }
                        
                        file.createdDateTime = (String) fileMap.get('createdDateTime');
                        file.lastModifiedDateTime = (String) fileMap.get('lastModifiedDateTime');
                        
                        files.add(file);
                    }
                } else {
                    System.debug('No files found in drive');
                }
            } else {
                throw new AuraHandledException('HTTP Error ' + res.getStatusCode() + ': ' + res.getBody());
            }
        } catch (Exception e) {
            System.debug('Exception in getSharePointFiles: ' + e.getMessage());
            throw new AuraHandledException('Error fetching SharePoint files: ' + e.getMessage());
        }
        
        return files;
    }
    
    @AuraEnabled(cacheable=true)
    public static List<SharePointSite> getSharePointSites() {
        List<SharePointSite> sites = new List<SharePointSite>();
        
        try {
            // Microsoft Graph API endpoint to get all sites - try root site first
            String endpoint = 'callout:MicrosoftApp/sites/root';
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Content-Type', 'application/json');
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            System.debug('SharePoint Sites Response Status: ' + res.getStatusCode());
            System.debug('SharePoint Sites Response Body: ' + res.getBody());
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> responseMap = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
                
                // For /sites/root endpoint, we get a single site object, not an array
                SharePointSite site = new SharePointSite();
                site.id = (String) responseMap.get('id');
                site.displayName = (String) responseMap.get('displayName');
                site.webUrl = (String) responseMap.get('webUrl');
                sites.add(site);
                
                System.debug('Root site found: ' + site.displayName);
            } else {
                throw new AuraHandledException('HTTP Error ' + res.getStatusCode() + ': ' + res.getBody());
            }
        } catch (Exception e) {
            System.debug('Exception in getSharePointSites: ' + e.getMessage());
            throw new AuraHandledException('Error fetching SharePoint sites: ' + e.getMessage());
        }
        
        return sites;
    }
    
    // Alternative method to get sites using followed sites
    @AuraEnabled(cacheable=true)
    public static List<SharePointSite> getFollowedSites() {
        List<SharePointSite> sites = new List<SharePointSite>();
        
        try {
            // Microsoft Graph API endpoint to get followed sites
            String endpoint = 'callout:MicrosoftApp/me/followedSites';
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Content-Type', 'application/json');
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            System.debug('Followed Sites Response Status: ' + res.getStatusCode());
            System.debug('Followed Sites Response Body: ' + res.getBody());
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> responseMap = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
                List<Object> value = (List<Object>) responseMap.get('value');
                
                if (value != null && !value.isEmpty()) {
                    for (Object siteObj : value) {
                        Map<String, Object> siteMap = (Map<String, Object>) siteObj;
                        SharePointSite site = new SharePointSite();
                        site.id = (String) siteMap.get('id');
                        site.displayName = (String) siteMap.get('displayName');
                        site.webUrl = (String) siteMap.get('webUrl');
                        sites.add(site);
                    }
                } else {
                    System.debug('No followed sites found');
                }
            } else {
                throw new AuraHandledException('HTTP Error ' + res.getStatusCode() + ': ' + res.getBody());
            }
        } catch (Exception e) {
            System.debug('Exception in getFollowedSites: ' + e.getMessage());
            throw new AuraHandledException('Error fetching followed sites: ' + e.getMessage());
        }
        
        return sites;
    }
    
    // Get version history for a specific file
    @AuraEnabled(cacheable=false)
    public static List<SharePointFileVersion> getFileVersionHistory(String siteId, String driveId, String fileId) {
        List<SharePointFileVersion> versions = new List<SharePointFileVersion>();
        
        try {
            // Microsoft Graph API endpoint to get file versions
            String endpoint = 'callout:MicrosoftApp/sites/' + siteId + '/drives/' + driveId + '/items/' + fileId + '/versions';
            
            System.debug('Version History Endpoint: ' + endpoint);
            System.debug('Site ID: ' + siteId + ', Drive ID: ' + driveId + ', File ID: ' + fileId);
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Content-Type', 'application/json');
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            System.debug('Version History Response Status: ' + res.getStatusCode());
            System.debug('Version History Response Body: ' + res.getBody());
            
            if (res.getStatusCode() == 200) {
                Map<String, Object> responseMap = (Map<String, Object>) JSON.deserializeUntyped(res.getBody());
                List<Object> value = (List<Object>) responseMap.get('value');
                
                if (value != null && !value.isEmpty()) {
                    for (Object versionObj : value) {
                        Map<String, Object> versionMap = (Map<String, Object>) versionObj;
                        SharePointFileVersion version = new SharePointFileVersion();
                        version.id = (String) versionMap.get('id');
                        version.lastModifiedDateTime = (String) versionMap.get('lastModifiedDateTime');
                        version.size = (Integer) versionMap.get('size');
                        
                        Map<String, Object> lastModifiedBy = (Map<String, Object>) versionMap.get('lastModifiedBy');
                        if (lastModifiedBy != null) {
                            Map<String, Object> user = (Map<String, Object>) lastModifiedBy.get('user');
                            if (user != null) {
                                version.lastModifiedBy = (String) user.get('displayName');
                            }
                        }
                        
                        versions.add(version);
                    }
                } else {
                    System.debug('No versions found for file');
                }
            } else {
                throw new AuraHandledException('HTTP Error ' + res.getStatusCode() + ': ' + res.getBody());
            }
        } catch (Exception e) {
            System.debug('Exception in getFileVersionHistory: ' + e.getMessage());
            throw new AuraHandledException('Error fetching file version history: ' + e.getMessage());
        }
        
        return versions;
    }
    
    // Get download URL for a specific version of a file
    @AuraEnabled(cacheable=false)
    public static String getVersionDownloadUrl(String siteId, String driveId, String fileId, String versionId) {
        try {
            // Microsoft Graph API endpoint to get version content/download URL
            String endpoint = 'callout:MicrosoftApp/sites/' + siteId + '/drives/' + driveId + '/items/' + fileId + '/versions/' + versionId + '/content';
            
            System.debug('Version Download URL Endpoint: ' + endpoint);
            System.debug('Site ID: ' + siteId + ', Drive ID: ' + driveId + ', File ID: ' + fileId + ', Version ID: ' + versionId);
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Content-Type', 'application/json');
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            System.debug('Version Download Response Status: ' + res.getStatusCode());
            System.debug('Version Download Response Headers: ' + res.getHeaderKeys());
            
            if (res.getStatusCode() == 302) {
                // Redirect response contains the download URL in Location header
                String downloadUrl = res.getHeader('Location');
                System.debug('Download URL: ' + downloadUrl);
                return downloadUrl;
            } else if (res.getStatusCode() == 200) {
                // Direct content - return the endpoint URL for download
                return endpoint;
            } else {
                throw new AuraHandledException('HTTP Error ' + res.getStatusCode() + ': ' + res.getBody());
            }
        } catch (Exception e) {
            System.debug('Exception in getVersionDownloadUrl: ' + e.getMessage());
            throw new AuraHandledException('Error getting version download URL: ' + e.getMessage());
        }
    }
    
    public class SharePointFile {
        @AuraEnabled public String id;
        @AuraEnabled public String name;
        @AuraEnabled public String webUrl;
        @AuraEnabled public Integer size;
        @AuraEnabled public String createdBy;
        @AuraEnabled public String createdDateTime;
        @AuraEnabled public String lastModifiedDateTime;
    }
    
    public class SharePointSite {
        @AuraEnabled public String id;
        @AuraEnabled public String displayName;
        @AuraEnabled public String webUrl;
    }
    
    public class SharePointFileVersion {
        @AuraEnabled public String id;
        @AuraEnabled public String lastModifiedDateTime;
        @AuraEnabled public Integer size;
        @AuraEnabled public String lastModifiedBy;
    }
    
    public class ConfigOption {
        @AuraEnabled public String label;
        @AuraEnabled public String value;
        @AuraEnabled public String description;
        @AuraEnabled public String driveId;
        @AuraEnabled public String siteId;
    }
}