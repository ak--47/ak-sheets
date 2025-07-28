// Test setup file for integration tests
import { config } from 'dotenv';
import { existsSync } from 'fs';
import { resolve } from 'path';
import { initSheets } from '../src/index.js';

// Load environment variables from .env file
config();

// Global test configuration
global.testConfig = {
  credentialsFile: process.env.GOOGLE_SHEETS_TEST_CREDENTIALS_FILE || 'credentials.json',
  skipIntegrationTests: false
};

// Check if credentials file exists
const credentialsPath = resolve(global.testConfig.credentialsFile);
if (!existsSync(credentialsPath)) {
  console.warn(`Warning: Credentials file not found at ${credentialsPath}`);
  console.warn('Integration tests will be skipped.');
  global.testConfig.skipIntegrationTests = true;
}

// Initialize ak-sheets if credentials are available
(async () => {
  if (!global.testConfig.skipIntegrationTests) {
    try {
      await initSheets({
        credentials: credentialsPath,
        environment: 'test',
        maxRetries: 3,        // Reduced for faster tests
        maxBackoffMs: 10000   // Reduced for faster tests
      });
      console.log('ak-sheets initialized for integration testing');
    } catch (error) {
      console.warn(`Warning: Failed to initialize ak-sheets: ${error.message}`);
      console.warn('Integration tests will be skipped.');
      global.testConfig.skipIntegrationTests = true;
    }
  }
})();