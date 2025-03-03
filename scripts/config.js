const fs = require('fs').promises;
const path = require('path');
const dotenv = require('dotenv');

async function updateLocalSettings() {
    try {
        // Load .env file
        const envPath = path.join(__dirname, '..', '.env');
        const envConfig = dotenv.parse(await fs.readFile(envPath));

        // Load local.settings.json
        const settingsPath = path.join(__dirname, '..', 'local.settings.json');
        const settings = JSON.parse(await fs.readFile(settingsPath, 'utf8'));

        // Update Values with environment variables
        for (const [key, value] of Object.entries(settings.Values)) {
            if (value.startsWith('%') && value.endsWith('%')) {
                const envKey = value.slice(1, -1);
                if (envConfig[envKey]) {
                    settings.Values[key] = envConfig[envKey];
                }
            }
        }

        // Write updated settings
        await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2));
        console.log('Successfully updated local.settings.json with environment variables');

    } catch (error) {
        if (error.code === 'ENOENT') {
            console.error('Error: Required configuration files not found. Please run "npm run setup" first.');
        } else {
            console.error('Error updating configuration:', error);
        }
        process.exit(1);
    }
}

// Run if called directly
if (require.main === module) {
    updateLocalSettings().catch(console.error);
}

module.exports = { updateLocalSettings };
