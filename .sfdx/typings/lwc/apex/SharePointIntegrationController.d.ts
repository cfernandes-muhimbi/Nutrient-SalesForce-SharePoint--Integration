declare module "@salesforce/apex/SharePointIntegrationController.getSharePointConfig" {
  export default function getSharePointConfig(): Promise<any>;
}
declare module "@salesforce/apex/SharePointIntegrationController.getAvailableConfigurations" {
  export default function getAvailableConfigurations(): Promise<any>;
}
declare module "@salesforce/apex/SharePointIntegrationController.testConfigurationUsage" {
  export default function testConfigurationUsage(): Promise<any>;
}
declare module "@salesforce/apex/SharePointIntegrationController.testConnection" {
  export default function testConnection(): Promise<any>;
}
declare module "@salesforce/apex/SharePointIntegrationController.getSharePointFiles" {
  export default function getSharePointFiles(param: {siteId: any, driveId: any}): Promise<any>;
}
declare module "@salesforce/apex/SharePointIntegrationController.getSharePointSites" {
  export default function getSharePointSites(): Promise<any>;
}
declare module "@salesforce/apex/SharePointIntegrationController.getFollowedSites" {
  export default function getFollowedSites(): Promise<any>;
}
declare module "@salesforce/apex/SharePointIntegrationController.getFileVersionHistory" {
  export default function getFileVersionHistory(param: {siteId: any, driveId: any, fileId: any}): Promise<any>;
}
declare module "@salesforce/apex/SharePointIntegrationController.getVersionDownloadUrl" {
  export default function getVersionDownloadUrl(param: {siteId: any, driveId: any, fileId: any, versionId: any}): Promise<any>;
}
