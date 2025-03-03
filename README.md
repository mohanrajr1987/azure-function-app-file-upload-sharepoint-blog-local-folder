# Azure Function App - File Upload with SharePoint Integration

This Azure Function App provides endpoints for handling file uploads with the following features:
- Single and multiple file uploads
- Integration with SharePoint for fetching files
- Storage in Azure Blob Storage (primary) or local folder (fallback)
- Built with Node.js 20

## Prerequisites

1. **Node.js 20**
   ```bash
   # Check Node.js version
   node --version  # Should be v20.x.x
   ```

2. **Azure Functions Core Tools v4**
   ```bash
   npm install -g azure-functions-core-tools@4
   ```

3. **Azure Requirements**
   - Azure Subscription
   - Azure Storage Account
   - Azure Function App (if deploying to Azure)

4. **SharePoint Requirements**
   - SharePoint Online Subscription
   - App Registration in Azure AD with appropriate permissions
   - Site Collection with necessary permissions

## Configuration

1. **Environment Setup**
   ```bash
   # Copy the example environment file
   cp .env.example .env
   ```

2. **Configure Environment Variables**
   Edit the `.env` file with your values:
   ```env
   # Required Configuration
   AZURE_STORAGE_CONNECTION_STRING=your_storage_connection_string
   SHAREPOINT_TENANT_ID=your_tenant_id
   SHAREPOINT_CLIENT_ID=your_client_id
   SHAREPOINT_CLIENT_SECRET=your_client_secret
   
   # Optional Configuration
   LOCAL_UPLOAD_PATH=uploads            # Default: uploads
   MAX_FILE_SIZE=10485760              # Default: 10MB
   FUNCTION_TIMEOUT=300000             # Default: 5 minutes
   ```

3. **SharePoint App Registration Setup**
   - Navigate to Azure Portal > Azure Active Directory
   - Create a new App Registration
   - Required API Permissions:
     - Microsoft Graph API:
       - Files.Read.All
       - Sites.Read.All
   - Generate a Client Secret

## Installation

1. **Install Dependencies**
   ```bash
   npm install
   ```

2. **Start the Function App**
   ```bash
   # Development
   npm start
   
   # With specific port
   func start --port 7073
   ```

## API Endpoints

### 1. File Upload
**Endpoint**: `/api/fileUpload`
```bash
# Example using curl
curl -X POST \
  -F "file=@/path/to/your/file.txt" \
  http://localhost:7073/api/fileUpload
```

### 2. SharePoint Upload
**Endpoint**: `/api/sharePointUpload`
```bash
# Example using curl
curl -X POST \
  -H "Content-Type: application/json" \
  -d '{
    "files": [{
      "siteId": "your-site-id",
      "driveId": "your-drive-id",
      "itemId": "your-item-id",
      "fileName": "example.txt"
    }]
  }' \
  http://localhost:7073/api/sharePointUpload
```

## Features

### Storage Options
1. **Azure Blob Storage (Primary)**
   - Automatically creates container if not exists
   - Generates unique file names using UUID
   - Handles concurrent uploads

2. **Local Storage (Fallback)**
   - Automatically used if Azure Storage is unavailable
   - Maintains original file structure
   - Generates unique file names

### SharePoint Integration
- Fetches files from SharePoint using Microsoft Graph API
- Supports multiple file downloads
- Handles SharePoint authentication automatically
- Falls back to mock data when SharePoint is not configured

## Error Handling

- **Automatic Fallback**: If Azure Blob Storage fails, automatically falls back to local storage
- **Detailed Error Messages**: All errors include specific error codes and messages
- **Independent Processing**: Each file in a batch is processed independently

## Security Best Practices

1. **Authentication**
   - Function-level authentication enabled
   - SharePoint uses OAuth 2.0 with client credentials

2. **Data Protection**
   - All credentials stored in environment variables
   - No sensitive data logged
   - File names sanitized before storage

3. **Input Validation**
   - File size limits enforced
   - File type validation
   - Request payload validation

## Troubleshooting

1. **Common Issues**
   - Port conflicts: Use `--port` to specify different port
   - Storage errors: Check connection string
   - SharePoint errors: Verify credentials and permissions

2. **Logging**
   - Check function logs: `func logs`
   - Application Insights (if configured)
   - Local storage logs in `uploads` directory

## Development

1. **Local Development**
   ```bash
   # Install dependencies
   npm install
   
   # Start with debugging
   func start --debug
   ```

2. **Testing**
   ```bash
   # Run tests
   npm test
   ```

3. **Deployment**
   ```bash
   # Deploy to Azure
   func azure functionapp publish YOUR_FUNCTION_APP_NAME
   ```

