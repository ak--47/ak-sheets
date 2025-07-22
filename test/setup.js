// Test setup file for integration tests
import { config } from 'dotenv';
import { readFileSync, existsSync } from 'fs';
import { resolve } from 'path';
import { init } from '../src/index.js';

// Load environment variables from .env file
config();

// Check if we're in CI or have required test environment variables
const requiredEnvVars = [
  'GOOGLE_SHEETS_TEST_CREDENTIALS_FILE',
  'GOOGLE_SHEETS_TEST_SPREADSHEET_ID',
  'GOOGLE_SHEETS_TEST_SPREADSHEET_NAME',
  'LOCAL_COPY_OF_TEST_SPREADSHEET'
];

const missingEnvVars = requiredEnvVars.filter(envVar => !process.env[envVar]);

if (missingEnvVars.length > 0) {
  console.warn(`Warning: Missing test environment variables: ${missingEnvVars.join(', ')}`);
  console.warn('Integration tests may fail. Please check your .env file.');
}

// Global test configuration
global.testConfig = {
  credentialsFile: process.env.GOOGLE_SHEETS_TEST_CREDENTIALS_FILE || 'credentials.json',
  testSpreadsheetId: process.env.GOOGLE_SHEETS_TEST_SPREADSHEET_ID,
  testSpreadsheetName: process.env.GOOGLE_SHEETS_TEST_SPREADSHEET_NAME || 'Test Spreadsheet for ak-sheets',
  localSpreadsheetFile: process.env.LOCAL_COPY_OF_TEST_SPREADSHEET || 'test-spreadsheet.xlsx',
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
if (!global.testConfig.skipIntegrationTests) {
  try {
    init({
      credentials: credentialsPath,
      environment: 'test'
    });
    console.log('ak-sheets initialized for integration testing');
  } catch (error) {
    console.warn(`Warning: Failed to initialize ak-sheets: ${error.message}`);
    console.warn('Integration tests will be skipped.');
    global.testConfig.skipIntegrationTests = true;
  }
}

// Helper function to read Excel file for comparison
global.readExcelFile = async function(filePath) {
  try {
    const xlsx = await import('xlsx');
    const workbook = xlsx.readFile(filePath);
    const sheets = {};
    
    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      sheets[sheetName] = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    });
    
    return sheets;
  } catch (error) {
    console.warn(`Warning: Could not read Excel file ${filePath}: ${error.message}`);
    return {};
  }
};

// Test data samples
global.testData = {
  simple: [
    { name: 'John Doe', age: 30, city: 'New York' },
    { name: 'Jane Smith', age: 25, city: 'San Francisco' },
    { name: 'Bob Johnson', age: 35, city: 'Chicago' }
  ],
  
  csvString: `Name,Age,City,Department
Alice Brown,28,Seattle,Engineering
Charlie Davis,32,Austin,Marketing
Diana Wilson,29,Boston,Sales`,
  
  arrayOfArrays: [
    ['Product', 'Price', 'Stock', 'Category'],
    ['Widget A', 29.99, 100, 'Electronics'],
    ['Gadget B', 49.99, 50, 'Electronics'],
    ['Tool C', 19.99, 200, 'Hardware']
  ],
  
  multiTab: {
    Users: [
      { id: 1, name: 'Admin User', role: 'admin' },
      { id: 2, name: 'Regular User', role: 'user' }
    ],
    Products: [
      { sku: 'PROD001', name: 'Product 1', price: 99.99 },
      { sku: 'PROD002', name: 'Product 2', price: 149.99 }
    ]
  }
};