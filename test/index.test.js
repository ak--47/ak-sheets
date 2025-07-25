import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { 
  init, 
  createSheet, 
  writeToSheet, 
  getSheet, 
  makeCSVFromData, 
  getURL,
  deleteSheet,
  csvToJson,
  jsonToCsv,
  appendToSheet,
  getRange,
  writeToRange,
  addTab,
  listTabs
} from '../src/index.js';

describe('ak-sheets Streamlined Integration Tests', () => {
  let testSpreadsheetId = null; // Single sheet for most tests

  beforeAll(async () => {
    if (global.testConfig.skipIntegrationTests) {
      console.warn('Skipping integration tests - missing credentials or configuration');
      return;
    }

    // Create one test spreadsheet to reuse for most tests
    console.log('Creating test spreadsheet for reuse...');
    testSpreadsheetId = await createSheet('ak-sheets-test-' + Date.now());
    console.log(`Created test spreadsheet: ${testSpreadsheetId}`);
  });

  afterAll(async () => {
    if (global.testConfig.skipIntegrationTests || !testSpreadsheetId) return;
    
    // Clean up the test spreadsheet
    console.log(`Cleaning up test spreadsheet: ${testSpreadsheetId}`);
    try {
      await deleteSheet(testSpreadsheetId);
      console.log('Test cleanup completed');
    } catch (error) {
      console.warn(`Failed to cleanup test sheet: ${error.message}`);
    }
  });

  describe('Core Functionality', () => {
    it('should initialize properly', () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      expect(() => {
        init({
          credentials: global.testConfig.credentialsFile,
          environment: 'test',
          maxRetries: 3, // Reduced for faster tests
          maxBackoffMs: 10000 // Reduced for faster tests
        });
      }).not.toThrow();
    });

    it('should perform complete CRUD workflow', async () => {
      if (global.testConfig.skipIntegrationTests) return;

      const testData = [
        { name: 'Alice', score: 95, department: 'Engineering' },
        { name: 'Bob', score: 87, department: 'Sales' },
        { name: 'Charlie', score: 92, department: 'Marketing' }
      ];

      // CREATE: Write data to sheet
      const writeResult = await writeToSheet(testSpreadsheetId, testData);
      expect(writeResult.updatedCells).toBeGreaterThan(0);

      // READ: Get data back
      const readData = await getSheet(testSpreadsheetId, undefined, 'json');
      expect(readData).toHaveLength(3);
      expect(readData[0].name).toBe('Alice');
      
      // UPDATE: Append more data
      const appendData = [{ name: 'David', score: 88, department: 'HR' }];
      const appendResult = await appendToSheet(testSpreadsheetId, appendData);
      expect(appendResult?.updates?.updatedCells).toBeGreaterThan(0);

      // VERIFY: Check total count
      const finalData = await getSheet(testSpreadsheetId, undefined, 'json');
      expect(finalData).toHaveLength(4);
    });

    it('should handle range operations', async () => {
      if (global.testConfig.skipIntegrationTests) return;

      // Write to specific range
      const rangeData = [['Product', 'Price'], ['Widget', '29.99'], ['Gadget', '49.99']];
      await writeToRange(testSpreadsheetId, 'E1:F3', rangeData);

      // Read from specific range
      const readRange = await getRange(testSpreadsheetId, 'E1:F3', undefined, 'array');
      expect(readRange).toHaveLength(3);
      expect(readRange[0]).toEqual(['Product', 'Price']);
      expect(readRange[1]).toEqual(['Widget', '29.99']);
    });

    it('should manage tabs', async () => {
      if (global.testConfig.skipIntegrationTests) return;

      // Add a tab
      const newTabId = await addTab(testSpreadsheetId, 'TestTab');
      expect(typeof newTabId).toBe('number');

      // List tabs to verify
      const tabs = await listTabs(testSpreadsheetId);
      const testTab = tabs.find(tab => tab.title === 'TestTab');
      expect(testTab).toBeDefined();
      expect(testTab.id).toBe(newTabId);

      // Write to the new tab
      const tabData = [{ item: 'Test', value: 123 }];
      await writeToSheet(testSpreadsheetId, tabData, 'TestTab');

      // Read from the tab
      const tabReadData = await getSheet(testSpreadsheetId, 'TestTab', 'json');
      expect(tabReadData).toHaveLength(1);
      expect(tabReadData[0].item).toBe('Test');
    });
  });

  describe('Data Conversion (Unit Tests)', () => {
    // These are pure functions - no API calls needed
    it('should convert CSV to JSON', () => {
      const csv = 'name,age\nJohn,30\nJane,25';
      const result = csvToJson(csv);
      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({ name: 'John', age: '30' });
    });

    it('should convert JSON to CSV', () => {
      const data = [{ name: 'John', age: 30 }, { name: 'Jane', age: 25 }];
      const result = jsonToCsv(data);
      expect(result).toContain('name,age');
      expect(result).toContain('John,30');
    });

    it('should generate proper CSV from data', () => {
      const data = [{ product: 'Widget', price: 29.99 }];
      const result = makeCSVFromData(data);
      expect(result).toContain('product,price');
      expect(result).toContain('"Widget","29.99"');
    });

    it('should generate correct URLs', () => {
      const url = getURL('test123');
      expect(url).toBe('https://docs.google.com/spreadsheets/d/test123');
    });
  });

  describe('Error Handling', () => {
    it('should handle invalid spreadsheet ID gracefully', async () => {
      if (global.testConfig.skipIntegrationTests) return;

      await expect(getSheet('invalid-id')).rejects.toThrow();
    });

    it('should handle missing tab gracefully', async () => {
      if (global.testConfig.skipIntegrationTests) return;

      await expect(getSheet(testSpreadsheetId, 'NonExistentTab')).rejects.toThrow();
    });
  });
});