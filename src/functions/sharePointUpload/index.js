const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
const path = require('path');
const { FileUploadHandler } = require('../fileUpload/index.js');

class SharePointHandler {
    constructor() {
        // Check if SharePoint credentials are available
        if (!process.env.SHAREPOINT_CLIENT_ID || !process.env.SHAREPOINT_CLIENT_SECRET || !process.env.SHAREPOINT_TENANT_ID) {
            return;
        }

        const credential = new ClientSecretCredential(
            process.env.SHAREPOINT_TENANT_ID,
            process.env.SHAREPOINT_CLIENT_ID,
            process.env.SHAREPOINT_CLIENT_SECRET
        );

        this.client = Client.initWithMiddleware({
            authProvider: {
                getAccessToken: async () => {
                    const token = await credential.getToken('https://graph.microsoft.com/.default');
                    return token.token;
                }
            }
        });
    }

    async getFileFromSharePoint(siteId, driveId, itemId) {
        if (!this.client) {
            throw new Error('SharePoint client not initialized - missing credentials');
        }

        try {
            const response = await this.client
                .api(`/sites/${siteId}/drives/${driveId}/items/${itemId}/content`)
                .get();
            return response;
        } catch (error) {
            throw new Error(`Failed to fetch file from SharePoint: ${error.message}`);
        }
    }
}

module.exports = async function (context, req) {
    try {
        // Check if SharePoint credentials are available
        if (!process.env.SHAREPOINT_CLIENT_ID || !process.env.SHAREPOINT_CLIENT_SECRET) {
            context.log.warn('SharePoint credentials not configured, using mock data');
            context.res = {
                status: 200,
                body: {
                    message: 'SharePoint integration not configured. Using mock data.',
                    mockMode: true,
                    results: [{
                        fileName: 'mock-file.txt',
                        success: true,
                        storage: 'local',
                        localPath: path.join(process.env.LocalUploadPath || 'uploads', 'mock-file.txt'),
                        mockData: true
                    }]
                }
            };
            return;
        }

        const files = req.body?.files;
        
        if (!files?.length) {
            context.res = {
                status: 400,
                body: {
                    error: 'No files provided',
                    details: 'Request must include an array of SharePoint files'
                }
            };
            return;
        }

        const sharePointHandler = new SharePointHandler();
        const fileHandler = new FileUploadHandler();
        const results = [];

        for (const file of files) {
            try {
                const { siteId, driveId, itemId, fileName } = file;
                
                if (!siteId || !driveId || !itemId || !fileName) {
                    results.push({
                        fileName: fileName || 'unknown',
                        success: false,
                        error: 'Missing required SharePoint file information'
                    });
                    continue;
                }

                const fileContent = await sharePointHandler.getFileFromSharePoint(
                    siteId,
                    driveId,
                    itemId
                );

                const result = await fileHandler.uploadFile(fileContent, fileName);
                results.push({
                    fileName,
                    ...result
                });
            } catch (error) {
                results.push({
                    fileName: file.fileName || 'unknown',
                    success: false,
                    error: error.message
                });
            }
        }

        context.res = {
            status: 200,
            body: {
                message: 'SharePoint files processed',
                results
            }
        };
    } catch (error) {
        context.log.error('SharePoint upload error:', error);
        context.res = {
            status: 500,
            body: {
                error: 'SharePoint upload failed',
                details: error.message
            }
        };
    }
};
