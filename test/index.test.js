import { describe, it, expect, beforeAll, afterAll, beforeEach, afterEach } from 'vitest';
import { 
  init, 
  createSheet, 
  writeToSheet, 
  getSheet, 
  makeCSVFromData, 
  getURL,
  shareSheet,
  deleteSheet,
  listOwnedSpreadsheets,
  updateSheet,
  writeToSheetTabs,
  readXlsxFile,
  csvToJson,
  jsonToCsv,
  appendToSheet,
  clearSheet,
  getSheetInfo,
  getRange,
  writeToRange,
  addTab,
  deleteTab,
  renameTab,
  duplicateTab,
  listTabs
} from '../src/index.js';

describe('ak-sheets Integration Tests', () => {
  let createdSheetIds = []; // Track sheets created during tests for cleanup

  beforeAll(() => {
    if (global.testConfig.skipIntegrationTests) {
      console.warn('Skipping integration tests - missing credentials or configuration');
    }
  });

  afterAll(async () => {
    if (global.testConfig.skipIntegrationTests) return;
    
    // Clean up any sheets created during tests
    console.log(`Cleaning up ${createdSheetIds.length} test spreadsheets...`);
    const cleanupPromises = createdSheetIds.map(async (id) => {
      try {
        await deleteSheet(id);
      } catch (error) {
        console.warn(`Failed to cleanup sheet ${id}: ${error.message}`);
      }
    });
    
    await Promise.allSettled(cleanupPromises);
    console.log('Test cleanup completed');
  });

  beforeEach(() => {
    if (global.testConfig.skipIntegrationTests) {
      // Skip individual tests if integration tests are disabled
      return;
    }
  });

  describe('Configuration and Initialization', () => {
    it('should be properly initialized from setup', () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      expect(global.testConfig.credentialsFile).toBeTruthy();
      expect(global.testConfig.testSpreadsheetName).toBeTruthy();
    });

    it('should initialize with file path credentials', () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      expect(() => {
        init({
          credentials: global.testConfig.credentialsFile,
          environment: 'test'
        });
      }).not.toThrow();
    });

    it('should initialize with SHEETS_CREDENTIALS environment variable', () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      // Set the environment variable
      const originalEnv = process.env.SHEETS_CREDENTIALS;
      process.env.SHEETS_CREDENTIALS = global.testConfig.credentialsFile;
      
      expect(() => {
        init({ environment: 'test' });
      }).not.toThrow();
      
      // Restore original value
      if (originalEnv) {
        process.env.SHEETS_CREDENTIALS = originalEnv;
      } else {
        delete process.env.SHEETS_CREDENTIALS;
      }
    });
  });

  describe('File and Data Conversion Utilities', () => {
    it('should read Excel file and convert to CSV', () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const excelData = readXlsxFile(global.testConfig.localSpreadsheetFile);
      expect(typeof excelData).toBe('object');
      
      const sheetNames = Object.keys(excelData);
      expect(sheetNames.length).toBeGreaterThan(0);
      
      // Each sheet should be a CSV string
      sheetNames.forEach(sheetName => {
        expect(typeof excelData[sheetName]).toBe('string');
        expect(excelData[sheetName].length).toBeGreaterThan(0);
      });
      
      console.log(`Read Excel file with ${sheetNames.length} sheets: ${sheetNames.join(', ')}`);
    });

    it('should convert CSV to JSON', () => {
      const csvString = 'Name,Age,City\nJohn,30,New York\nJane,25,San Francisco';
      const jsonData = csvToJson(csvString);
      
      expect(Array.isArray(jsonData)).toBe(true);
      expect(jsonData).toHaveLength(2);
      expect(jsonData[0]).toEqual({ Name: 'John', Age: '30', City: 'New York' });
      expect(jsonData[1]).toEqual({ Name: 'Jane', Age: '25', City: 'San Francisco' });
    });

    it('should convert JSON to CSV', () => {
      const jsonData = [
        { Name: 'John', Age: 30, City: 'New York' },
        { Name: 'Jane', Age: 25, City: 'San Francisco' }
      ];
      
      const csvString = jsonToCsv(jsonData);
      expect(typeof csvString).toBe('string');
      expect(csvString).toContain('Name,Age,City');
      expect(csvString).toContain('John,30,New York');
      expect(csvString).toContain('Jane,25,San Francisco');
    });

    it('should handle CSV to JSON with custom options', () => {
      const csvString = 'name|age|city\nJohn|30|NYC\nJane|25|SF';
      const jsonData = csvToJson(csvString, { delimiter: '|' });
      
      expect(jsonData).toHaveLength(2);
      expect(jsonData[0].name).toBe('John');
      expect(jsonData[0].age).toBe('30');
    });

    it('should handle JSON to CSV with custom options', () => {
      const jsonData = [{ name: 'John', age: 30 }];
      const csvString = jsonToCsv(jsonData, { delimiter: '|' });
      
      expect(csvString).toContain('name|age');
      expect(csvString).toContain('John|30');
    });

    it('should handle empty data in conversions', () => {
      expect(csvToJson('')).toEqual([]);
      expect(jsonToCsv([])).toBe('');
    });
  });

  describe('Utility Functions', () => {
    it('should generate correct Google Sheets URL', () => {
      const spreadsheetId = 'test-spreadsheet-id';
      const expected = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
      const actual = getURL(spreadsheetId);
      expect(actual).toBe(expected);
    });

    it('should convert array of objects to CSV', () => {
      const data = [
        { name: 'John', age: 30, city: 'New York' },
        { name: 'Jane', age: 25, city: 'San Francisco' }
      ];
      
      const csv = makeCSVFromData(data);
      const lines = csv.split('\n');
      
      expect(lines[0]).toBe('name,age,city');
      expect(lines[1]).toBe('"John","30","New York"');
      expect(lines[2]).toBe('"Jane","25","San Francisco"');
    });

    it('should handle empty data in CSV conversion', () => {
      const csv = makeCSVFromData([]);
      expect(csv).toBe('');
    });

    it('should handle objects with different keys', () => {
      const data = [
        { name: 'John', age: 30 },
        { name: 'Jane', city: 'SF', age: 25 }
      ];
      
      const csv = makeCSVFromData(data);
      const lines = csv.split('\n');
      
      // Should include all unique keys
      const headers = lines[0].split(',');
      expect(headers).toEqual(expect.arrayContaining(['name', 'age', 'city']));
    });

    it('should escape CSV special characters', () => {
      const data = [
        { name: 'John "Johnny" Doe', description: 'Line 1\nLine 2', value: 'a,b,c' }
      ];
      
      const csv = makeCSVFromData(data);
      expect(csv).toContain('John ""Johnny"" Doe');
    });

    it('should handle null and undefined values in CSV', () => {
      const data = [
        { name: 'John', age: null, city: undefined },
        { name: null, age: 25, city: 'SF' }
      ];
      
      const csv = makeCSVFromData(data);
      const lines = csv.split('\n');
      
      expect(lines[1]).toBe(`"John",,`);
      expect(lines[2]).toBe(`,"25","SF"`);
    });

    it('should handle complex objects in CSV', () => {
      const data = [
        { 
          name: 'John', 
          metadata: { tags: ['developer', 'js'], active: true },
          scores: [95, 87, 92]
        }
      ];
      
      const csv = makeCSVFromData(data);
      expect(csv).toContain('John');
      // Complex objects should be JSON stringified with single quotes
      expect(csv).toContain("'developer'");
    });

    it('should respect character limit in CSV conversion', () => {
      const longString = 'a'.repeat(1000);
      const data = [{ content: longString }];
      
      const csv = makeCSVFromData(data, 100);
      const contentCell = csv.split('\n')[1];
      
      // Should be truncated to 100 chars plus quotes and escaping
      expect(contentCell.length).toBeLessThan(longString.length + 10);
    });
  });

  describe('Google Sheets API Operations', () => {
    it('should list owned spreadsheets', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const sheets = await listOwnedSpreadsheets();
      expect(Array.isArray(sheets)).toBe(true);
      
      // Each sheet should have expected properties
      if (sheets.length > 0) {
        const firstSheet = sheets[0];
        expect(firstSheet).toHaveProperty('id');
        expect(firstSheet).toHaveProperty('name');
        expect(typeof firstSheet.id).toBe('string');
        expect(typeof firstSheet.name).toBe('string');
      }
    });

    it('should create a new spreadsheet', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const testSheetName = `Test Sheet ${Date.now()}`;
      const spreadsheetId = await createSheet(testSheetName);
      
      expect(typeof spreadsheetId).toBe('string');
      expect(spreadsheetId.length).toBeGreaterThan(0);
      
      // Track for cleanup
      createdSheetIds.push(spreadsheetId);
      
      // Verify it was created
      const sheets = await listOwnedSpreadsheets();
      const createdSheet = sheets.find(sheet => sheet.id === spreadsheetId);
      expect(createdSheet).toBeDefined();
      expect(createdSheet.name).toBe(testSheetName);
    });

    it('should create spreadsheet with custom tabs', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const testSheetName = `Multi-Tab Sheet ${Date.now()}`;
      const tabs = ['Users', 'Products', 'Orders'];
      const spreadsheetId = await createSheet(testSheetName, tabs);
      
      createdSheetIds.push(spreadsheetId);
      
      // Verify tabs exist by trying to write to them
      for (const tab of tabs) {
        const result = await writeToSheet(spreadsheetId, [['Test']], tab);
        expect(result.updatedCells).toBeGreaterThan(0);
      }
    });

    it('should write array of objects to sheet', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Object Data Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      const result = await writeToSheet(spreadsheetId, global.testData.simple);
      expect(result.updatedCells).toBeGreaterThan(0);
      
      // Verify data was written correctly
      const readData = await getSheet(spreadsheetId, null, 'json');
      expect(readData).toHaveLength(3);
      expect(readData[0].name).toBe('John Doe');
      expect(readData[0].age).toBe('30');
      expect(readData[0].city).toBe('New York');
    });

    it('should write CSV string to sheet', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`CSV Data Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      const result = await writeToSheet(spreadsheetId, global.testData.csvString);
      expect(result.updatedCells).toBeGreaterThan(0);
      
      // Verify data
      const readData = await getSheet(spreadsheetId, null, 'json');
      expect(readData).toHaveLength(3);
      expect(readData[0].Name).toBe('Alice Brown');
      expect(readData[0].Department).toBe('Engineering');
    });

    it('should write array of arrays to sheet', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Array Data Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      const result = await writeToSheet(spreadsheetId, global.testData.arrayOfArrays);
      expect(result.updatedCells).toBeGreaterThan(0);
      
      // Verify data
      const readData = await getSheet(spreadsheetId, null, 'array');
      expect(readData).toHaveLength(4); // Header + 3 data rows
      expect(readData[0]).toEqual(['Product', 'Price', 'Stock', 'Category']);
      expect(readData[1][0]).toBe('Widget A');
    });

    it('should write to specific tab', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Tab Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      const tabName = 'TestTab';
      const result = await writeToSheet(spreadsheetId, global.testData.simple, tabName);
      expect(result.updatedCells).toBeGreaterThan(0);
      
      // Verify data in specific tab
      const readData = await getSheet(spreadsheetId, tabName, 'json');
      expect(readData).toHaveLength(3);
      expect(readData[0].name).toBe('John Doe');
    });

    it('should write to multiple tabs', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Multi-Tab Write Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      const results = await writeToSheetTabs(spreadsheetId, global.testData.multiTab);
      expect(results).toHaveLength(2);
      expect(results[0].updatedCells).toBeGreaterThan(0);
      expect(results[1].updatedCells).toBeGreaterThan(0);
      
      // Verify each tab
      const usersData = await getSheet(spreadsheetId, 'Users', 'json');
      const productsData = await getSheet(spreadsheetId, 'Products', 'json');
      
      expect(usersData).toHaveLength(2);
      expect(productsData).toHaveLength(2);
      expect(usersData[0].name).toBe('Admin User');
      expect(productsData[0].sku).toBe('PROD001');
    });

    it('should read data in different formats', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Read Format Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      // Write some data first
      await writeToSheet(spreadsheetId, global.testData.simple);
      
      // Test JSON format (default)
      const jsonData = await getSheet(spreadsheetId);
      expect(Array.isArray(jsonData)).toBe(true);
      expect(typeof jsonData[0]).toBe('object');
      expect(jsonData[0].name).toBe('John Doe');
      
      // Test array format
      const arrayData = await getSheet(spreadsheetId, null, 'array');
      expect(Array.isArray(arrayData)).toBe(true);
      expect(Array.isArray(arrayData[0])).toBe(true);
      expect(arrayData[0]).toEqual(['name', 'age', 'city']);
      
      // Test CSV format
      const csvData = await getSheet(spreadsheetId, null, 'csv');
      expect(typeof csvData).toBe('string');
      expect(csvData).toContain('name,age,city');
      expect(csvData).toContain('John Doe');
    });

    it('should update existing sheet data', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Update Test ${Date.now()}`, ['UpdateTab']);
      createdSheetIds.push(spreadsheetId);
      
      // Write initial data
      const initialData = [
        { name: 'John', age: 30 },
        { name: 'Jane', age: 25 }
      ];
      await writeToSheet(spreadsheetId, initialData, 'UpdateTab');
      
      // Update with new data
      const updatedData = [
        { name: 'John', age: 31 }, // Updated age
        { name: 'Bob', age: 35 }   // New person
      ];
      
      const result = await updateSheet(spreadsheetId, updatedData, 'UpdateTab');
      expect(result.updatedCells).toBeGreaterThan(0);
      
      // Verify updated data
      const readData = await getSheet(spreadsheetId, 'UpdateTab', 'json');
      expect(readData).toHaveLength(2);
      expect(readData[0].name).toBe('John');
      expect(readData[0].age).toBe('31'); // Should be updated
      expect(readData[1].name).toBe('Bob'); // Should be new
    });

    it('should share spreadsheet', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Share Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      // Test sharing (use a test email or skip if not available)
      const shareOptions = {
        userEmail: 'test@example.com',
        role: 'reader'
      };
      
      // This might fail depending on permissions, so we wrap in try-catch
      try {
        const result = await shareSheet(spreadsheetId, shareOptions);
        expect(result).toBeDefined();
      } catch (error) {
        console.warn(`Share test failed (expected in some environments): ${error.message}`);
      }
    });

    it('should append data to existing sheet', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Append Test ${Date.now()}`, ['AppendTab']);
      createdSheetIds.push(spreadsheetId);
      
      // Write initial data
      const initialData = [{ name: 'John', age: 30 }];
      await writeToSheet(spreadsheetId, initialData, 'AppendTab');
      
      // Append more data
      const appendData = [{ name: 'Jane', age: 25 }, { name: 'Bob', age: 35 }];
      const result = await appendToSheet(spreadsheetId, appendData, 'AppendTab');
      expect(result.updatedCells).toBeGreaterThan(0);
      
      // Verify all data is present
      const allData = await getSheet(spreadsheetId, 'AppendTab', 'json');
      expect(allData).toHaveLength(3);
      expect(allData[0].name).toBe('John');
      expect(allData[1].name).toBe('Jane');
      expect(allData[2].name).toBe('Bob');
    });

    it('should clear sheet data', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Clear Test ${Date.now()}`, ['ClearTab']);
      createdSheetIds.push(spreadsheetId);
      
      // Write some data
      await writeToSheet(spreadsheetId, global.testData.simple, 'ClearTab');
      
      // Verify data exists
      let data = await getSheet(spreadsheetId, 'ClearTab', 'json');
      expect(data).toHaveLength(3);
      
      // Clear the data
      await clearSheet(spreadsheetId, 'ClearTab');
      
      // Verify data is cleared
      data = await getSheet(spreadsheetId, 'ClearTab', 'json');
      expect(data).toHaveLength(0);
    });

    it('should get spreadsheet info', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const testName = `Info Test ${Date.now()}`;
      const tabs = ['Tab1', 'Tab2', 'Tab3'];
      const spreadsheetId = await createSheet(testName, tabs);
      createdSheetIds.push(spreadsheetId);
      
      const info = await getSheetInfo(spreadsheetId);
      expect(info).toHaveProperty('properties');
      expect(info).toHaveProperty('sheets');
      expect(info.properties.title).toBe(testName);
      expect(info.sheets).toHaveLength(3);
      
      const sheetTitles = info.sheets.map(sheet => sheet.properties.title);
      expect(sheetTitles).toEqual(expect.arrayContaining(tabs));
    });

    it('should handle non-existent spreadsheet gracefully', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const fakeId = 'non-existent-spreadsheet-id';
      
      await expect(getSheet(fakeId)).rejects.toThrow();
      await expect(deleteSheet(fakeId)).rejects.toThrow();
      await expect(getSheetInfo(fakeId)).rejects.toThrow();
      await expect(clearSheet(fakeId)).rejects.toThrow();
    });
  });

  describe('Range Operations', () => {
    it('should read and write to specific ranges', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Range Test ${Date.now()}`, ['RangeTab']);
      createdSheetIds.push(spreadsheetId);
      
      // Write data to specific range
      const data = [['Name', 'Age', 'City'], ['John', 30, 'NYC'], ['Jane', 25, 'SF']];
      await writeToRange(spreadsheetId, 'A1:C3', data, 'RangeTab');
      
      // Read full range back
      const fullRange = await getRange(spreadsheetId, 'A1:C3', 'RangeTab', 'array');
      expect(fullRange).toHaveLength(3);
      expect(fullRange[0]).toEqual(['Name', 'Age', 'City']);
      expect(fullRange[1]).toEqual(['John', '30', 'NYC']);
      
      // Read partial range
      const nameColumn = await getRange(spreadsheetId, 'A:A', 'RangeTab', 'array');
      expect(nameColumn[0]).toEqual(['Name']);
      expect(nameColumn[1]).toEqual(['John']);
      expect(nameColumn[2]).toEqual(['Jane']);
    });

    it('should write single values to specific cells', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Single Cell Test ${Date.now()}`, ['CellTab']);
      createdSheetIds.push(spreadsheetId);
      
      // Write single value
      await writeToRange(spreadsheetId, 'B2', 'Hello World', 'CellTab');
      
      // Read it back
      const cellValue = await getRange(spreadsheetId, 'B2', 'CellTab', 'array');
      expect(cellValue).toEqual([['Hello World']]);
    });

    it('should handle different data formats in ranges', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Data Format Test ${Date.now()}`, ['FormatTab']);
      createdSheetIds.push(spreadsheetId);
      
      // Write array of objects to range
      const objectData = [{ name: 'Alice', score: 95 }, { name: 'Bob', score: 87 }];
      await writeToRange(spreadsheetId, 'A1', objectData, 'FormatTab');
      
      // Read as JSON
      const jsonData = await getRange(spreadsheetId, 'A1:B3', 'FormatTab', 'json');
      expect(jsonData).toHaveLength(2);
      expect(jsonData[0].name).toBe('Alice');
      expect(jsonData[0].score).toBe('95');
      
      // Read as CSV
      const csvData = await getRange(spreadsheetId, 'A1:B3', 'FormatTab', 'csv');
      expect(typeof csvData).toBe('string');
      expect(csvData).toContain('name,score');
      expect(csvData).toContain('Alice,95');
    });

    it('should handle CSV string input to ranges', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`CSV Range Test ${Date.now()}`, ['CSVTab']);
      createdSheetIds.push(spreadsheetId);
      
      // Write CSV string to range
      const csvString = 'Product,Price,Stock\nWidget,29.99,100\nGadget,49.99,50';
      await writeToRange(spreadsheetId, 'D1', csvString, 'CSVTab');
      
      // Read it back
      const data = await getRange(spreadsheetId, 'D1:F3', 'CSVTab', 'json');
      expect(data).toHaveLength(2);
      expect(data[0].Product).toBe('Widget');
      expect(data[0].Price).toBe('29.99');
      expect(data[1].Product).toBe('Gadget');
    });

    it('should work with ranges without specifying tab', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`No Tab Range Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      // Write to default sheet without specifying tab
      await writeToRange(spreadsheetId, 'A1:B2', [['X', 'Y'], ['1', '2']]);
      
      // Read back without specifying tab
      const data = await getRange(spreadsheetId, 'A1:B2', undefined, 'array');
      expect(data).toEqual([['X', 'Y'], ['1', '2']]);
    });
  });

  describe('Tab Management Operations', () => {
    it('should add new tabs', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Tab Add Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      // Add a simple tab
      const sheetId = await addTab(spreadsheetId, 'NewTab');
      expect(typeof sheetId).toBe('number');
      
      // Verify tab was created
      const tabs = await listTabs(spreadsheetId);
      const newTab = tabs.find(tab => tab.title === 'NewTab');
      expect(newTab).toBeDefined();
      expect(newTab.id).toBe(sheetId);
    });

    it('should add tabs with options', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Tab Options Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      // Add tab with color and position
      await addTab(spreadsheetId, 'ColoredTab', {
        index: 0,
        tabColor: '#FF0000',
        hidden: false
      });
      
      // Verify tab exists
      const tabs = await listTabs(spreadsheetId);
      const coloredTab = tabs.find(tab => tab.title === 'ColoredTab');
      expect(coloredTab).toBeDefined();
      expect(coloredTab.index).toBe(0);
    });

    it('should prevent adding duplicate tabs', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Duplicate Tab Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      // Add first tab
      await addTab(spreadsheetId, 'TestTab');
      
      // Try to add duplicate - should fail
      await expect(addTab(spreadsheetId, 'TestTab')).rejects.toThrow('already exists');
    });

    it('should list all tabs', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const tabNames = ['Tab1', 'Tab2', 'Tab3'];
      const spreadsheetId = await createSheet(`List Tabs Test ${Date.now()}`, tabNames);
      createdSheetIds.push(spreadsheetId);
      
      const tabs = await listTabs(spreadsheetId);
      expect(tabs).toHaveLength(3);
      
      const tabTitles = tabs.map(tab => tab.title);
      expect(tabTitles).toEqual(expect.arrayContaining(tabNames));
      
      // Check tab properties
      tabs.forEach(tab => {
        expect(tab).toHaveProperty('id');
        expect(tab).toHaveProperty('title');
        expect(tab).toHaveProperty('index');
        expect(tab).toHaveProperty('hidden');
        expect(typeof tab.id).toBe('number');
        expect(typeof tab.title).toBe('string');
        expect(typeof tab.index).toBe('number');
        expect(typeof tab.hidden).toBe('boolean');
      });
    });

    it('should rename tabs', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Rename Tab Test ${Date.now()}`, ['OldName']);
      createdSheetIds.push(spreadsheetId);
      
      // Rename the tab
      await renameTab(spreadsheetId, 'OldName', 'NewName');
      
      // Verify rename worked
      const tabs = await listTabs(spreadsheetId);
      const renamedTab = tabs.find(tab => tab.title === 'NewName');
      const oldTab = tabs.find(tab => tab.title === 'OldName');
      
      expect(renamedTab).toBeDefined();
      expect(oldTab).toBeUndefined();
    });

    it('should prevent renaming to existing tab name', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Rename Conflict Test ${Date.now()}`, ['Tab1', 'Tab2']);
      createdSheetIds.push(spreadsheetId);
      
      // Try to rename Tab1 to Tab2 - should fail
      await expect(renameTab(spreadsheetId, 'Tab1', 'Tab2')).rejects.toThrow('already exists');
    });

    it('should duplicate tabs', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Duplicate Tab Test ${Date.now()}`, ['Template']);
      createdSheetIds.push(spreadsheetId);
      
      // Add some data to the template
      await writeToSheet(spreadsheetId, [['Header1', 'Header2'], ['Data1', 'Data2']], 'Template');
      
      // Duplicate the tab
      const newTabInfo = await duplicateTab(spreadsheetId, 'Template', 'Copy1');
      expect(newTabInfo).toHaveProperty('sheetId');
      expect(newTabInfo).toHaveProperty('title');
      expect(newTabInfo.title).toBe('Copy1');
      
      // Verify both tabs exist
      const tabs = await listTabs(spreadsheetId);
      expect(tabs).toHaveLength(2);
      
      const originalTab = tabs.find(tab => tab.title === 'Template');
      const duplicatedTab = tabs.find(tab => tab.title === 'Copy1');
      expect(originalTab).toBeDefined();
      expect(duplicatedTab).toBeDefined();
      
      // Verify data was copied
      const originalData = await getSheet(spreadsheetId, 'Template', 'array');
      const duplicatedData = await getSheet(spreadsheetId, 'Copy1', 'array');
      expect(duplicatedData).toEqual(originalData);
    });

    it('should duplicate tabs with auto-generated names', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Auto Name Duplicate Test ${Date.now()}`, ['Source']);
      createdSheetIds.push(spreadsheetId);
      
      // Duplicate without specifying new name
      const newTabInfo = await duplicateTab(spreadsheetId, 'Source');
      expect(newTabInfo.title).toBe('Copy of Source');
      
      // Verify it exists
      const tabs = await listTabs(spreadsheetId);
      const duplicatedTab = tabs.find(tab => tab.title === 'Copy of Source');
      expect(duplicatedTab).toBeDefined();
    });

    it('should delete tabs', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Delete Tab Test ${Date.now()}`, ['KeepMe', 'DeleteMe']);
      createdSheetIds.push(spreadsheetId);
      
      // Verify both tabs exist initially
      let tabs = await listTabs(spreadsheetId);
      expect(tabs).toHaveLength(2);
      
      // Delete one tab
      await deleteTab(spreadsheetId, 'DeleteMe');
      
      // Verify only one tab remains
      tabs = await listTabs(spreadsheetId);
      expect(tabs).toHaveLength(1);
      expect(tabs[0].title).toBe('KeepMe');
    });

    it('should handle tab operations with non-existent tabs gracefully', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Error Handling Test ${Date.now()}`, ['ExistingTab']);
      createdSheetIds.push(spreadsheetId);
      
      // Try operations on non-existent tab
      await expect(deleteTab(spreadsheetId, 'NonExistentTab')).rejects.toThrow('not found');
      await expect(renameTab(spreadsheetId, 'NonExistentTab', 'NewName')).rejects.toThrow('not found');
      await expect(duplicateTab(spreadsheetId, 'NonExistentTab')).rejects.toThrow('not found');
    });

    it('should integrate range operations with tab management', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const spreadsheetId = await createSheet(`Integration Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      // Add a new tab
      await addTab(spreadsheetId, 'DataTab');
      
      // Write data to specific range in the new tab
      const data = [['Quarter', 'Revenue'], ['Q1', '100K'], ['Q2', '150K']];
      await writeToRange(spreadsheetId, 'B2:C4', data, 'DataTab');
      
      // Read it back
      const readData = await getRange(spreadsheetId, 'B2:C4', 'DataTab', 'json');
      expect(readData).toHaveLength(2);
      expect(readData[0].Quarter).toBe('Q1');
      expect(readData[0].Revenue).toBe('100K');
      
      // Duplicate the tab with data
      await duplicateTab(spreadsheetId, 'DataTab', 'DataCopy');
      
      // Verify data exists in both tabs
      const originalData = await getRange(spreadsheetId, 'B2:C4', 'DataTab', 'json');
      const copiedData = await getRange(spreadsheetId, 'B2:C4', 'DataCopy', 'json');
      expect(copiedData).toEqual(originalData);
    });
  });

  describe('Data Verification with Local Excel File', () => {
    it('should read local Excel file', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const excelData = await global.readExcelFile(global.testConfig.localSpreadsheetFile);
      expect(typeof excelData).toBe('object');
      
      // Should have at least one sheet
      const sheetNames = Object.keys(excelData);
      expect(sheetNames.length).toBeGreaterThan(0);
      
      console.log(`Local Excel file has sheets: ${sheetNames.join(', ')}`);
    });

    it('should compare Google Sheets data with local Excel file', async () => {
      if (global.testConfig.skipIntegrationTests || !global.testConfig.testSpreadsheetId) return;
      
      // Read data from Google Sheets
      const googleSheetsData = await getSheet(global.testConfig.testSpreadsheetId, null, 'array');
      
      // Read data from local Excel file
      const excelData = await global.readExcelFile(global.testConfig.localSpreadsheetFile);
      
      if (Object.keys(excelData).length === 0) {
        console.warn('Skipping Excel comparison - could not read local file');
        return;
      }
      
      // Compare first sheet (assuming single sheet or main data is in first sheet)
      const firstExcelSheet = Object.values(excelData)[0];
      
      expect(googleSheetsData.length).toBeGreaterThan(0);
      expect(firstExcelSheet.length).toBeGreaterThan(0);
      
      // Compare headers if both exist
      if (googleSheetsData.length > 0 && firstExcelSheet.length > 0) {
        console.log('Google Sheets headers:', googleSheetsData[0]);
        console.log('Excel headers:', firstExcelSheet[0]);
        
        // Basic structure comparison
        expect(googleSheetsData[0].length).toBe(firstExcelSheet[0].length);
      }
    });

    it('should write Excel data to new Google Sheet and verify', async () => {
      if (global.testConfig.skipIntegrationTests) return;
      
      const excelData = await global.readExcelFile(global.testConfig.localSpreadsheetFile);
      
      if (Object.keys(excelData).length === 0) {
        console.warn('Skipping Excel upload test - could not read local file');
        return;
      }
      
      const spreadsheetId = await createSheet(`Excel Import Test ${Date.now()}`);
      createdSheetIds.push(spreadsheetId);
      
      // Use first sheet from Excel file
      const firstSheetName = Object.keys(excelData)[0];
      const firstSheetData = excelData[firstSheetName];
      
      // Write Excel data to Google Sheets
      const result = await writeToSheet(spreadsheetId, firstSheetData);
      expect(result.updatedCells).toBeGreaterThan(0);
      
      // Read back and verify
      const writtenData = await getSheet(spreadsheetId, null, 'array');
      expect(writtenData.length).toBe(firstSheetData.length);
      
      // Compare headers
      expect(writtenData[0]).toEqual(firstSheetData[0]);
      
      console.log(`Successfully uploaded ${firstSheetData.length} rows from Excel to Google Sheets`);
    });
  });
});