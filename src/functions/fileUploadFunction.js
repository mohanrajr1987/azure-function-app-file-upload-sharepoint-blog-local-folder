import { app } from '@azure/functions';
import { BlobServiceClient } from '@azure/storage-blob';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { v4 as uuidv4 } from 'uuid';
import { promises as fs } from 'fs';
import { join } from 'path';
import { fileURLToPath } from 'url';

const __dirname = fileURLToPath(new URL('.', import.meta.url));

// File Upload Handler Class
class FileUploadHandler {
    #connectionString;
    #localUploadPath;

    constructor() {
        this.#connectionString = process.env.AzureStorageConnectionString;
        this.#localUploadPath = process.env.LocalUploadPath || join(__dirname, '..', '..', 'uploads');
        this.#initializeStorage().catch(console.error);
    }

    async #initializeStorage() {
        try {
            await fs.mkdir(this.#localUploadPath, { recursive: true });
        } catch (error) {
            console.error('Error creating upload directory:', error);
        }
    }

    async uploadFile(fileContent, fileName) {
        // Try blob storage first if connection string is available
        if (this.#connectionString) {
            try {
                return await this.#uploadToBlob(fileContent, fileName);
            } catch (error) {
                console.warn('Blob storage upload failed, falling back to local storage:', error);
                return await this.#saveToLocal(fileContent, fileName);
            }
        } else {
            // If no connection string, go straight to local storage
            console.info('No Azure Storage connection string found, using local storage');
            return await this.#saveToLocal(fileContent, fileName);
        }
    }

    async #uploadToBlob(fileContent, fileName, containerName = 'uploads') {
        const blobServiceClient = BlobServiceClient.fromConnectionString(this.#connectionString);
        const containerClient = blobServiceClient.getContainerClient(containerName);
        
        // Ensure container exists
        await containerClient.createIfNotExists({
            access: 'blob'
        });

        // Generate unique blob name
        const blobName = `${uuidv4()}-${fileName}`;
        const blockBlobClient = containerClient.getBlockBlobClient(blobName);
        
        await blockBlobClient.upload(fileContent, Buffer.byteLength(fileContent));
        
        return {
            success: true,
            storage: 'blob',
            url: blockBlobClient.url,
            blobName
        };
    }

    async #saveToLocal(fileContent, fileName) {
        const uniqueFileName = `${uuidv4()}-${fileName}`;
        const filePath = join(this.#localUploadPath, uniqueFileName);
        await fs.writeFile(filePath, fileContent);
        return {
            success: true,
            storage: 'local',
            localPath: filePath,
            fileName: uniqueFileName
        };
    }
}

// SharePoint Handler Class
class SharePointHandler {
    constructor() {
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
        try {
            const response = await this.client
                .api(`/sites/${siteId}/drives/${driveId}/items/${itemId}/content`)
                .get();
            return response;
        } catch (error) {
            console.error('SharePoint file fetch failed:', error);
            throw error;
        }
    }
}

// File Upload Function
app.http('fileUpload', {
    methods: ['POST'],
    authLevel: 'function',
    handler: async (request, context) => {
        try {
            const fileHandler = new FileUploadHandler();
            const files = await request.parseFormData();
            const results = [];

            for (const [_, file] of files.entries()) {
                if (file.length) {
                    const result = await fileHandler.uploadToBlob(
                        file.buffer,
                        file.filename
                    );
                    results.push({
                        fileName: file.filename,
                        ...result
                    });
                }
            }

            return {
                status: 200,
                jsonBody: {
                    message: 'Files processed successfully',
                    results
                }
            };
        } catch (error) {
            context.error('File upload error:', error);
            return {
                status: 500,
                jsonBody: {
                    error: 'File upload failed',
                    details: error.message
                }
            };
        }
    }
});

// SharePoint Upload Function
app.http('sharePointUpload', {
    methods: ['POST'],
    authLevel: 'function',
    handler: async (request, context) => {
        try {
            const { files } = await request.json();
            
            if (!files || !Array.isArray(files) || files.length === 0) {
                return {
                    status: 400,
                    jsonBody: {
                        error: 'Invalid request. Expected array of SharePoint files'
                    }
                };
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

                    const uploadResult = await fileHandler.uploadToBlob(
                        fileContent,
                        fileName
                    );

                    results.push({
                        fileName,
                        ...uploadResult
                    });
                } catch (error) {
                    results.push({
                        fileName: file.fileName || 'unknown',
                        success: false,
                        error: error.message
                    });
                }
            }

            return {
                status: 200,
                jsonBody: {
                    message: 'SharePoint files processed',
                    results
                }
            };
        } catch (error) {
            context.error('SharePoint upload error:', error);
            return {
                status: 500,
                jsonBody: {
                    error: 'SharePoint upload failed',
                    details: error.message
                }
            };
        }
    }
});
