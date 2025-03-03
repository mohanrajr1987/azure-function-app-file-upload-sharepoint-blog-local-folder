const { BlobServiceClient } = require('@azure/storage-blob');
const { v4: uuidv4 } = require('uuid');
const fs = require('fs').promises;
const path = require('path');

class FileUploadHandler {
    constructor() {
        this.connectionString = process.env.AzureStorageConnectionString;
        this.localUploadPath = process.env.LocalUploadPath || path.join(__dirname, '..', '..', '..', 'uploads');
    }

    async ensureLocalUploadPath() {
        try {
            await fs.mkdir(this.localUploadPath, { recursive: true });
        } catch (error) {
            console.error('Error creating upload directory:', error);
        }
    }

    async uploadFile(fileContent, fileName) {
        await this.ensureLocalUploadPath();
        
        // Try blob storage first if connection string is available
        if (this.connectionString) {
            try {
                return await this.uploadToBlob(fileContent, fileName);
            } catch (error) {
                console.warn('Blob storage upload failed, falling back to local storage:', error);
                return await this.saveToLocal(fileContent, fileName);
            }
        } else {
            // If no connection string, go straight to local storage
            console.info('No Azure Storage connection string found, using local storage');
            return await this.saveToLocal(fileContent, fileName);
        }
    }

    async uploadToBlob(fileContent, fileName, containerName = 'uploads') {
        const blobServiceClient = BlobServiceClient.fromConnectionString(this.connectionString);
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

    async saveToLocal(fileContent, fileName) {
        const uniqueFileName = `${uuidv4()}-${fileName}`;
        const filePath = path.join(this.localUploadPath, uniqueFileName);
        await fs.writeFile(filePath, fileContent);
        return {
            success: true,
            storage: 'local',
            localPath: filePath,
            fileName: uniqueFileName
        };
    }
}

module.exports = async function (context, req) {
    try {
        const fileHandler = new FileUploadHandler();
        
        // Check if we have any files
        if (!req.files || req.files.length === 0) {
            context.res = {
                status: 400,
                body: {
                    error: 'No files were uploaded',
                    details: 'Request must include at least one file'
                }
            };
            return;
        }

        const results = [];

        // Process each file
        for (const file of req.files) {
            try {
                const result = await fileHandler.uploadFile(
                    file.buffer,
                    file.originalname || 'unnamed-file'
                );
                results.push({
                    fileName: file.originalname,
                    ...result
                });
            } catch (error) {
                results.push({
                    fileName: file.originalname || 'unnamed-file',
                    success: false,
                    error: error.message
                });
            }
        }

        context.res = {
            status: 200,
            body: {
                message: 'Files processed successfully',
                results
            }
        };
    } catch (error) {
        context.log.error('Error processing file upload:', error);
        context.res = {
            status: 500,
            body: {
                error: 'File upload failed',
                details: error.message
            }
        };
    }
};
