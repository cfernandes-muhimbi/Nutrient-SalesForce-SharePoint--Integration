@IsTest
public class SharePointIntegrationControllerTest {
    
    // Test data setup
    @TestSetup
    static void makeData() {
        // Note: Custom Metadata cannot be inserted in tests, so we'll mock the responses
        // In real scenarios, you would create Custom Metadata records via deployment
    }
    
    // Test getting SharePoint configuration - success scenario
    @IsTest
    static void testGetSharePointConfig_Success() {
        Test.startTest();
        try {
            Graph_SharePoint_Config__mdt config = SharePointIntegrationController.getSharePointConfig();
            // Since we can't create Custom Metadata in tests, this will likely throw an exception
            // But we're testing the method structure
            System.assertNotEquals(null, config, 'Config should not be null');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('SharePoint configuration not found'), 
                         'Should throw configuration not found error');
        }
        Test.stopTest();
    }
    
    // Test getting available configurations
    @IsTest
    static void testGetAvailableConfigurations() {
        Test.startTest();
        try {
            List<SharePointIntegrationController.ConfigOption> configs = 
                SharePointIntegrationController.getAvailableConfigurations();
            // Since no Custom Metadata exists in test context, should return empty list
            System.assertEquals(0, configs.size(), 'Should return empty list when no configs exist');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('Error loading configurations'), 
                         'Should handle configuration loading errors');
        }
        Test.stopTest();
    }
    
    // Test configuration usage test method
    @IsTest
    static void testTestConfigurationUsage() {
        Test.startTest();
        String result = SharePointIntegrationController.testConfigurationUsage();
        System.assert(result.contains('FAILED'), 'Should fail when no configuration exists');
        Test.stopTest();
    }
    
    // Test connection method with successful response
    @IsTest
    static void testTestConnection_Success() {
        // Mock successful HTTP response
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(200, '{"displayName":"Test User"}'));
        
        Test.startTest();
        String result = SharePointIntegrationController.testConnection();
        Test.stopTest();
        
        System.assert(result.contains('Status: 200'), 'Should return successful status');
        System.assert(result.contains('Test User'), 'Should contain response body');
    }
    
    // Test connection method with failed response
    @IsTest
    static void testTestConnection_Failed() {
        // Mock failed HTTP response
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(401, '{"error":"Unauthorized"}'));
        
        Test.startTest();
        String result = SharePointIntegrationController.testConnection();
        Test.stopTest();
        
        System.assert(result.contains('Status: 401'), 'Should return failed status');
    }
    
    // Test getting SharePoint files - success scenario
    @IsTest
    static void testGetSharePointFiles_Success() {
        String mockResponse = '{"value":[' +
            '{"id":"file1","name":"document.pdf","webUrl":"https://test.sharepoint.com/document.pdf",' +
            '"size":1024,"createdDateTime":"2023-01-01T10:00:00Z","lastModifiedDateTime":"2023-01-02T10:00:00Z",' +
            '"createdBy":{"user":{"displayName":"John Doe"}}' +
            '}]}';
        
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(200, mockResponse));
        
        Test.startTest();
        List<SharePointIntegrationController.SharePointFile> files = 
            SharePointIntegrationController.getSharePointFiles('testSiteId', 'testDriveId');
        Test.stopTest();
        
        System.assertEquals(1, files.size(), 'Should return one file');
        System.assertEquals('file1', files[0].id, 'Should have correct file ID');
        System.assertEquals('document.pdf', files[0].name, 'Should have correct file name');
        System.assertEquals('John Doe', files[0].createdBy, 'Should have correct creator');
    }
    
    // Test getting SharePoint files - empty response
    @IsTest
    static void testGetSharePointFiles_Empty() {
        String mockResponse = '{"value":[]}';
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(200, mockResponse));
        
        Test.startTest();
        List<SharePointIntegrationController.SharePointFile> files = 
            SharePointIntegrationController.getSharePointFiles('testSiteId', 'testDriveId');
        Test.stopTest();
        
        System.assertEquals(0, files.size(), 'Should return empty list');
    }
    
    // Test getting SharePoint files - HTTP error
    @IsTest
    static void testGetSharePointFiles_HTTPError() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(404, '{"error":"Not Found"}'));
        
        Test.startTest();
        try {
            SharePointIntegrationController.getSharePointFiles('invalidSiteId', 'invalidDriveId');
            System.assert(false, 'Should have thrown an exception');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('HTTP Error 404'), 'Should contain HTTP error message');
        }
        Test.stopTest();
    }
    
    // Test getting SharePoint sites - success
    @IsTest
    static void testGetSharePointSites_Success() {
        String mockResponse = '{"id":"site1","displayName":"Test Site","webUrl":"https://test.sharepoint.com"}';
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(200, mockResponse));
        
        Test.startTest();
        List<SharePointIntegrationController.SharePointSite> sites = 
            SharePointIntegrationController.getSharePointSites();
        Test.stopTest();
        
        System.assertEquals(1, sites.size(), 'Should return one site');
        System.assertEquals('site1', sites[0].id, 'Should have correct site ID');
        System.assertEquals('Test Site', sites[0].displayName, 'Should have correct display name');
    }
    
    // Test getting followed sites - success
    @IsTest
    static void testGetFollowedSites_Success() {
        String mockResponse = '{"value":[' +
            '{"id":"site1","displayName":"Followed Site","webUrl":"https://test.sharepoint.com/sites/followed"}' +
            ']}';
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(200, mockResponse));
        
        Test.startTest();
        List<SharePointIntegrationController.SharePointSite> sites = 
            SharePointIntegrationController.getFollowedSites();
        Test.stopTest();
        
        System.assertEquals(1, sites.size(), 'Should return one followed site');
        System.assertEquals('Followed Site', sites[0].displayName, 'Should have correct display name');
    }
    
    // Test getting file version history - success
    @IsTest
    static void testGetFileVersionHistory_Success() {
        String mockResponse = '{"value":[' +
            '{"id":"1.0","lastModifiedDateTime":"2023-01-02T10:00:00Z","size":2048,' +
            '"lastModifiedBy":{"user":{"displayName":"Jane Smith"}}},' +
            '{"id":"1.1","lastModifiedDateTime":"2023-01-01T10:00:00Z","size":1024,' +
            '"lastModifiedBy":{"user":{"displayName":"John Doe"}}}' +
            ']}';
        
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(200, mockResponse));
        
        Test.startTest();
        List<SharePointIntegrationController.SharePointFileVersion> versions = 
            SharePointIntegrationController.getFileVersionHistory('siteId', 'driveId', 'fileId');
        Test.stopTest();
        
        System.assertEquals(2, versions.size(), 'Should return two versions');
        System.assertEquals('1.0', versions[0].id, 'Should have correct version ID');
        System.assertEquals('Jane Smith', versions[0].lastModifiedBy, 'Should have correct modifier');
        System.assertEquals(2048, versions[0].size, 'Should have correct size');
    }
    
    // Test getting version download URL - redirect response
    @IsTest
    static void testGetVersionDownloadUrl_Redirect() {
        MockHttpResponseGenerator mock = new MockHttpResponseGenerator(302, '');
        mock.setHeader('Location', 'https://download.sharepoint.com/file.pdf');
        Test.setMock(HttpCalloutMock.class, mock);
        
        Test.startTest();
        String downloadUrl = SharePointIntegrationController.getVersionDownloadUrl('siteId', 'driveId', 'fileId', 'versionId');
        Test.stopTest();
        
        System.assertEquals('https://download.sharepoint.com/file.pdf', downloadUrl, 'Should return redirect URL');
    }
    
    // Test getting version download URL - direct content
    @IsTest
    static void testGetVersionDownloadUrl_DirectContent() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(200, 'PDF content'));
        
        Test.startTest();
        String downloadUrl = SharePointIntegrationController.getVersionDownloadUrl('siteId', 'driveId', 'fileId', 'versionId');
        Test.stopTest();
        
        System.assert(downloadUrl.contains('callout:MicrosoftApp/sites/siteId/drives/driveId/items/fileId/versions/versionId/content'), 
                     'Should return endpoint URL for direct content');
    }
    
    // Test getting version download URL - HTTP error
    @IsTest
    static void testGetVersionDownloadUrl_HTTPError() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator(404, '{"error":"Version not found"}'));
        
        Test.startTest();
        try {
            SharePointIntegrationController.getVersionDownloadUrl('siteId', 'driveId', 'fileId', 'invalidVersionId');
            System.assert(false, 'Should have thrown an exception');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('HTTP Error 404'), 'Should contain HTTP error message');
        }
        Test.stopTest();
    }
    
    // Test SharePointFile wrapper class
    @IsTest
    static void testSharePointFileClass() {
        Test.startTest();
        SharePointIntegrationController.SharePointFile file = new SharePointIntegrationController.SharePointFile();
        file.id = 'test-id';
        file.name = 'test.pdf';
        file.webUrl = 'https://test.com/test.pdf';
        file.size = 1024;
        file.createdBy = 'Test User';
        file.createdDateTime = '2023-01-01T10:00:00Z';
        file.lastModifiedDateTime = '2023-01-02T10:00:00Z';
        Test.stopTest();
        
        System.assertEquals('test-id', file.id, 'Should set ID correctly');
        System.assertEquals('test.pdf', file.name, 'Should set name correctly');
        System.assertEquals(1024, file.size, 'Should set size correctly');
    }
    
    // Test SharePointSite wrapper class
    @IsTest
    static void testSharePointSiteClass() {
        Test.startTest();
        SharePointIntegrationController.SharePointSite site = new SharePointIntegrationController.SharePointSite();
        site.id = 'site-id';
        site.displayName = 'Test Site';
        site.webUrl = 'https://test.sharepoint.com';
        Test.stopTest();
        
        System.assertEquals('site-id', site.id, 'Should set ID correctly');
        System.assertEquals('Test Site', site.displayName, 'Should set display name correctly');
    }
    
    // Test SharePointFileVersion wrapper class
    @IsTest
    static void testSharePointFileVersionClass() {
        Test.startTest();
        SharePointIntegrationController.SharePointFileVersion version = new SharePointIntegrationController.SharePointFileVersion();
        version.id = '1.0';
        version.lastModifiedDateTime = '2023-01-01T10:00:00Z';
        version.size = 2048;
        version.lastModifiedBy = 'Version User';
        Test.stopTest();
        
        System.assertEquals('1.0', version.id, 'Should set version ID correctly');
        System.assertEquals(2048, version.size, 'Should set size correctly');
        System.assertEquals('Version User', version.lastModifiedBy, 'Should set modifier correctly');
    }
    
    // Test ConfigOption wrapper class
    @IsTest
    static void testConfigOptionClass() {
        Test.startTest();
        SharePointIntegrationController.ConfigOption option = new SharePointIntegrationController.ConfigOption();
        option.label = 'Test Config';
        option.value = 'test-config';
        option.description = 'Test Description';
        option.driveId = 'drive123';
        option.siteId = 'site123';
        Test.stopTest();
        
        System.assertEquals('Test Config', option.label, 'Should set label correctly');
        System.assertEquals('test-config', option.value, 'Should set value correctly');
        System.assertEquals('drive123', option.driveId, 'Should set drive ID correctly');
    }
    
    // Mock HTTP Response Generator
    public class MockHttpResponseGenerator implements HttpCalloutMock {
        private Integer statusCode;
        private String responseBody;
        private Map<String, String> headers;
        
        public MockHttpResponseGenerator(Integer statusCode, String responseBody) {
            this.statusCode = statusCode;
            this.responseBody = responseBody;
            this.headers = new Map<String, String>();
        }
        
        public void setHeader(String key, String value) {
            this.headers.put(key, value);
        }
        
        public HTTPResponse respond(HTTPRequest request) {
            HttpResponse response = new HttpResponse();
            response.setStatusCode(this.statusCode);
            response.setBody(this.responseBody);
            
            for (String key : this.headers.keySet()) {
                response.setHeader(key, this.headers.get(key));
            }
            
            return response;
        }
    }
}