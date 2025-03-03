const fs = require('fs').promises;
const path = require('path');
const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
});

const question = (query) => new Promise((resolve) => readline.question(query, resolve));

async function setup() {
    try {
        console.log('Azure Function App - Setup Script');
        console.log('================================\n');

        // Check if .env exists
        const envPath = path.join(__dirname, '.env');
        const envExamplePath = path.join(__dirname, '.env.example');
        
        if (!await exists(envPath) && await exists(envExamplePath)) {
            await fs.copyFile(envExamplePath, envPath);
            console.log('Created .env file from .env.example');
        }

        // Create required directories
        const dirs = ['uploads', 'scripts'];
        for (const dir of dirs) {
            const dirPath = path.join(__dirname, dir);
            if (!await exists(dirPath)) {
                await fs.mkdir(dirPath);
                console.log(`Created ${dir} directory`);
            }
        }

        // Get configuration values
        const config = {
            AZURE_STORAGE_CONNECTION_STRING: await question('Enter Azure Storage Connection String (leave empty for local storage only): '),
            AZURE_STORAGE_CONTAINER_NAME: await question('Enter Azure Storage Container Name (default: uploads): ') || 'uploads',
            SHAREPOINT_TENANT_ID: await question('Enter SharePoint Tenant ID (leave empty to skip SharePoint integration): '),
            SHAREPOINT_CLIENT_ID: await question('Enter SharePoint Client ID: '),
            SHAREPOINT_CLIENT_SECRET: await question('Enter SharePoint Client Secret: '),
            SHAREPOINT_SITE_URL: await question('Enter SharePoint Site URL: '),
            LOCAL_UPLOAD_PATH: 'uploads',
            MAX_FILE_SIZE: '10485760', // 10MB
            FUNCTION_TIMEOUT: '300000', // 5 minutes
            APPINSIGHTS_INSTRUMENTATIONKEY: await question('Enter Application Insights Instrumentation Key (optional): ')
        };

        // Update .env file
        let envContent = '';
        for (const [key, value] of Object.entries(config)) {
            if (value) {
                envContent += `${key}=${value}\n`;
            }
        }

        await fs.writeFile(envPath, envContent);
        console.log('\nConfiguration saved to .env file');

        // Run config script to update local.settings.json
        const { updateLocalSettings } = require('./scripts/config');
        await updateLocalSettings();
        console.log('Updated local.settings.json with environment variables');

        console.log('\nSetup complete! You can now run:');
        console.log('1. npm install');
        console.log('2. npm start\n');

    } catch (error) {
        console.error('Error during setup:', error);
        if (error.code === 'MODULE_NOT_FOUND') {
            console.error('Please run npm install first to install required dependencies');
        }
    } finally {
        readline.close();
    }
}

async function exists(path) {
    try {
        await fs.access(path);
        return true;
    } catch {
        return false;
    }
}

setup();
