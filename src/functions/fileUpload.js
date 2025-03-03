const { app } = require('@azure/functions');
const { BlobServiceClient } = require('@azure/storage-blob');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
const fs = require('fs').promises;
const path = require('path');

// Utility class for handling file uploads
class FileUploadHandler {
    constructor() {
        this.connectionString = process.env.AzureStorageConnectionString;
        this.localUploadPath = process.env.LocalUploadPath || 'uploads';
        this.ensureLocalUploadPath();
    }

    async ensureLocalUploadPath() {
        try {
            await fs.mkdir(this.localUploadPath, { recursive: true });
        } catch (error) {
            console.error('Error creating upload directory:', error);
        }
    }

    async uploadToBlob(fileContent, fileName, containerName = 'uploads') {
        try {
            const blobServiceClient = BlobServiceClient.fromConnectionString(this.connectionString);
            const containerClient = blobServiceClient.getContainerClient(containerName);
            
            // Create container if it doesn't exist
            await containerClient.createIfNotExists();

            const blockBlobClient = containerClient.getBlockBlobClient(fileName);
            await blockBlobClient.upload(fileContent, fileContent.length);
            
            return blockBlobClient.url;
        } catch (error) {
            console.error('Failed to upload to blob storage:', error);
            return this.saveToLocal(fileContent, fileName);
        }
    }

    async saveToLocal(fileContent, fileName) {
        try {
            const filePath = path.join(this.localUploadPath, fileName);
            await fs.writeFile(filePath, fileContent);
            return filePath;
        } catch (error) {
            console.error('Failed to save file locally:', error);
            throw error;
        }
    }
}

// SharePoint handler class
class SharePointHandler {
    constructor() {
        const credential = new ClientSecretCredential(
            process.env.SharePointTenantId,
            process.env.SharePointClientId,
            process.env.SharePointClientSecret
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

    async getFileFromSharePoint(driveId, itemId) {
        try {
            const response = await this.client
                .api(`/drives/${driveId}/items/${itemId}/content`)
                .get();
            return response;
        } catch (error) {
            console.error('Error fetching file from SharePoint:', error);
            throw error;
        }
    }
}

// Function to handle direct file uploads
app.http('uploadFiles', {
    methods: ['POST'],
    authLevel: 'function',
    handler: async (request, context) => {
        try {
            const fileHandler = new FileUploadHandler();
            const files = await request.parseFormData();
            const filesData = [];

            for (const [_, file] of files.entries()) {
                if (file.length) {
                    const fileContent = file.buffer;
                    const fileName = file.filename;

                    try {
                        const fileUrl = await fileHandler.uploadToBlob(fileContent, fileName);
                        filesData.push({ fileName, url: fileUrl });
                    } catch (error) {
                        const localPath = await fileHandler.saveToLocal(fileContent, fileName);
                        filesData.push({ fileName, localPath });
                    }
                }
            }

            return { 
                status: 200,
                jsonBody: { 
                    message: 'Files uploaded successfully',
                    files: filesData 
                }
            };
        } catch (error) {
            context.error('Error processing file upload:', error);
            return {
                status: 500,
                jsonBody: { error: error.message }
            };
        }
    }
});

// Function to handle SharePoint file uploads
app.http('sharePointUpload', {
    methods: ['POST'],
    authLevel: 'function',
    handler: async (request, context) => {
        try {
            const { sharePointFiles } = await request.json();
            
            if (!sharePointFiles || !sharePointFiles.length) {
                return {
                    status: 400,
                    jsonBody: { error: 'No SharePoint files provided' }
                };
            }

            const sharePointHandler = new SharePointHandler();
            const fileHandler = new FileUploadHandler();
            const filesData = [];

            for (const file of sharePointFiles) {
                try {
                    const fileContent = await sharePointHandler.getFileFromSharePoint(
                        file.driveId,
                        file.itemId
                    );
                    
                    try {
                        const fileUrl = await fileHandler.uploadToBlob(fileContent, file.name);
                        filesData.push({ fileName: file.name, url: fileUrl });
                    } catch (error) {
                        const localPath = await fileHandler.saveToLocal(fileContent, file.name);
                        filesData.push({ fileName: file.name, localPath });
                    }
                } catch (error) {
                    context.error(`Failed to process SharePoint file ${file.name}:`, error);
                    filesData.push({ fileName: file.name, error: error.message });
                }
            }

            return {
                status: 200,
                jsonBody: {
                    message: 'SharePoint files processed',
                    files: filesData
                }
            };
        } catch (error) {
            context.error('Error processing SharePoint upload:', error);
            return {
                status: 500,
                jsonBody: { error: error.message }
            };
        }
    }
});
