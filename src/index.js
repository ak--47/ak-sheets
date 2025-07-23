/**
 * @fileoverview Google Sheets API functions for ak-sheets
 * A modern, simple interface to work deeply with Google Sheets data
 */

import { google } from 'googleapis';
import Papa from 'papaparse';
import pino from 'pino';
import { readFileSync, existsSync } from 'fs';
import { resolve } from 'path';
import xlsx from 'xlsx';

let credentials = null;
// Default logger - will be reconfigured in init()
let logger = pino({ level: 'info' });
let environment = 'prod';
/** @type {any} */
let auth = null;
/** @type {any} */
let sheets = null;
/** @type {any} */
let drive = null;

// Retry configuration - can be overridden in init()
let maxRetries = 5;
let maxBackoffMs = 64000;

/**
 * Loads credentials from various sources
 * @param {any} credentialsInput - Can be object, file path, or undefined
 * @returns {any} Parsed credentials object
 */
function loadCredentials(credentialsInput) {
    // Check environment variable first
    const envCredentialsPath = process.env.SHEETS_CREDENTIALS;
    
    let credentialsToLoad = credentialsInput;
    
    // Use env var if no credentials passed
    if (!credentialsToLoad && envCredentialsPath) {
        credentialsToLoad = envCredentialsPath;
    }
    
    if (!credentialsToLoad) {
        throw new Error('No credentials provided. Pass credentials object/path or set SHEETS_CREDENTIALS env var.');
    }
    
    // If it's already an object, return it
    if (typeof credentialsToLoad === 'object' && credentialsToLoad !== null) {
        return credentialsToLoad;
    }
    
    // If it's a string, treat as file path
    if (typeof credentialsToLoad === 'string') {
        const credentialsPath = resolve(credentialsToLoad);
        
        if (!existsSync(credentialsPath)) {
            throw new Error(`Credentials file not found: ${credentialsPath}`);
        }
        
        try {
            const credentialsContent = readFileSync(credentialsPath, 'utf-8');
            return JSON.parse(credentialsContent);
        } catch (error) {
            throw new Error(`Failed to parse credentials file: ${error.message}`);
        }
    }
    
    throw new Error('Invalid credentials format. Must be object or file path string.');
}

/**
 * Exponential backoff retry mechanism for Google API calls
 * @param {Function} apiCall - The API function to call
 * @param {number} [customMaxRetries] - Override retry count (uses global config if not provided)
 * @param {number} [customMaxBackoff] - Override backoff time (uses global config if not provided)
 * @returns {Promise<any>} Result of the API call
 */
async function retryWithBackoff(apiCall, customMaxRetries, customMaxBackoff) {
    const maxRetriesConfig = customMaxRetries ?? maxRetries;
    const maxBackoffConfig = customMaxBackoff ?? maxBackoffMs;
    let currentRetry = 0;
    
    while (currentRetry <= maxRetriesConfig) {
        try {
            return await apiCall();
        } catch (error) {
            const isQuotaError = error.code === 429 || 
                                (error.message && error.message.includes('Quota exceeded')) ||
                                (error.message && error.message.includes('Too many requests'));
            
            const isRetryableError = isQuotaError || 
                                   error.code === 500 || 
                                   error.code === 502 || 
                                   error.code === 503 || 
                                   error.code === 504;
            
            if (!isRetryableError || currentRetry >= maxRetriesConfig) {
                logger.error({
                    error: error.message,
                    retryCount: currentRetry,
                    maxRetries: maxRetriesConfig,
                    isQuotaError,
                    isRetryableError
                }, 'API call failed after retries');
                throw error;
            }
            
            // Calculate exponential backoff: min(((2^n) + random), maxBackoff)
            const baseDelay = Math.pow(2, currentRetry) * 1000; // Start with 1 second
            const jitter = Math.random() * 1000; // Add up to 1 second of jitter
            const delay = Math.min(baseDelay + jitter, maxBackoffConfig);
            
            logger.warn({
                error: error.message,
                retryCount: currentRetry + 1,
                maxRetries: maxRetriesConfig,
                delayMs: Math.round(delay),
                isQuotaError
            }, 'API call failed, retrying with backoff');
            
            await new Promise(resolve => setTimeout(resolve, delay));
            currentRetry++;
        }
    }
}

/**
 * Initializes the ak-sheets library with configuration
 * @param {import('./index.d.ts').AkSheetsConfig} config - Configuration options
 * @example
 * import { init } from 'ak-sheets';
 * 
 * // Initialize with credentials object
 * init({
 *   credentials: {
 *     type: "service_account",
 *     project_id: "your-project",
 *     // ... other credential fields
 *   },
 *   environment: 'dev'
 * });
 * 
 * // Initialize with file path and custom retry settings
 * init({
 *   credentials: './credentials.json',
 *   environment: 'prod',
 *   maxRetries: 3,
 *   maxBackoffMs: 32000
 * });
 * 
 * // Initialize using environment variables
 * // Set SHEETS_CREDENTIALS=./credentials.json
 * init({});
 * 
 * // Custom logging level and retry configuration
 * // Set LOG_LEVEL=debug for verbose output
 * init({ 
 *   credentials: './credentials.json',
 *   maxRetries: 10,
 *   maxBackoffMs: 120000
 * });
 */
export function init(config) {
    credentials = loadCredentials(config.credentials);
    
    // Use passed environment, then NODE_ENV, then default to 'prod'
    environment = config.environment || process.env.NODE_ENV || 'prod';
    
    if (config.logger) {
        logger = config.logger;
    } else {
        const isDev = environment === 'dev' || environment === 'test';
        logger = pino({ 
            level: process.env.LOG_LEVEL || (isDev ? 'debug' : 'info'),
            transport: isDev
                ? {
                    target: 'pino-pretty',
                    options: { 
                        colorize: true, 
                        translateTime: true,
                        levelFirst: true,
                        messageFormat: '[ak-sheets] {msg}'
                    }
                }
                : undefined // In prod, keep as JSON for cloud logging
        });
    }

    // Configure retry settings
    if (typeof config.maxRetries === 'number') {
        maxRetries = config.maxRetries;
    }
    if (typeof config.maxBackoffMs === 'number') {
        maxBackoffMs = config.maxBackoffMs;
    }

    auth = new google.auth.GoogleAuth({
        credentials,
        scopes: [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/drive.file',
            'https://www.googleapis.com/auth/drive.metadata.readonly'
        ],
    });

    sheets = google.sheets({ version: 'v4', auth });
    drive = google.drive({ version: 'v3', auth });

    logger.info('ak-sheets initialized successfully');
}

/**
 * Generates a random name for a spreadsheet
 * @returns {string} Random name
 */
function makeName() {
    return `sheet-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`;
}

/**
 * Creates a new Google Spreadsheet
 * @param {string} [name] - Name of the spreadsheet (optional, generates random name if not provided)
 * @param {string[]} [tabs=[]] - Array of tab names to create (optional)
 * @returns {Promise<string>} Promise resolving to the spreadsheet ID
 * @example
 * import { createSheet } from 'ak-sheets';
 * 
 * // Create with auto-generated name
 * const id = await createSheet();
 * 
 * // Create with custom name
 * const id2 = await createSheet('My Data Sheet');
 * 
 * // Create with custom tabs
 * const id3 = await createSheet('Multi-Tab Sheet', ['Users', 'Products', 'Orders']);
 * console.log(`Created spreadsheet: ${id3}`);
 */
export async function createSheet(name = makeName(), tabs = []) {
    if (!sheets || !drive) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ name, tabs }, 'Creating spreadsheet');

    const existingSheets = await listOwnedSpreadsheets();

    // Check if sheet with the same name already exists
    const existingSheet = existingSheets.find(sheet => sheet.name === name);
    if (existingSheet) {
        logger.info({ spreadsheetId: existingSheet.id, name }, 'Found existing spreadsheet');
        return existingSheet.id;
    }

    try {
        // Create a new spreadsheet if no existing sheet is found
        const response = await retryWithBackoff(() => 
            sheets.spreadsheets.create({
                resource: {
                    properties: {
                        title: name,
                    },
                },
            })
        );

        const spreadsheetId = response?.data?.spreadsheetId;
        logger.info({ spreadsheetId, name }, 'Spreadsheet created successfully');

        // Add tabs to the sheet
        if (tabs.length > 0) {
            logger.debug({ tabs }, 'Adding tabs to spreadsheet');
            const addSheetRequests = tabs.map(tabName => ({
                addSheet: {
                    properties: {
                        title: tabName
                    }
                }
            }));

            await retryWithBackoff(() =>
                sheets.spreadsheets.batchUpdate({
                    spreadsheetId,
                    resource: {
                        requests: addSheetRequests
                    }
                })
            );

            // Remove default sheet if new tabs are added
            await retryWithBackoff(() =>
                sheets.spreadsheets.batchUpdate({
                    spreadsheetId,
                    resource: {
                        requests: [{
                            deleteSheet: {
                                sheetId: 0 // Default first sheet
                            }
                        }]
                    }
                })
            );

            logger.debug({ tabCount: tabs.length }, 'Tabs added and default sheet removed');
        }

        // Share the sheet with default user if in dev environment
        if (environment === 'dev') {
            await shareSheet(spreadsheetId);
        }

        return spreadsheetId;
    } catch (error) {
        logger.error({ error:  (error).message, name, tabs }, 'Failed to create spreadsheet');
        throw error;
    }
}

/**
 * Write data to a specific tab in a Google Spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {import('./index.d.ts').SpreadsheetData} [rows=""] - Data to write (can be CSV string, array of arrays, or array of objects)
 * @param {string} [tab] - Optional tab name to write to
 * @returns {Promise<import('./index.d.ts').SheetResponse>} Spreadsheet update response
 * @example
 * import { writeToSheet } from 'ak-sheets';
 * 
 * // Write array of objects
 * const data = [{ name: 'John', age: 30 }, { name: 'Jane', age: 25 }];
 * await writeToSheet(spreadsheetId, data);
 * 
 * // Write to specific tab
 * await writeToSheet(spreadsheetId, data, 'Users');
 * 
 * // Write CSV string
 * const csv = 'Name,Age\nBob,35\nAlice,28';
 * await writeToSheet(spreadsheetId, csv, 'Employees');
 */
export async function writeToSheet(spreadsheetId, rows = "", tab) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, tab, dataType: typeof rows }, 'Writing to sheet');

    // Convert rows to CSV if it's an array
    if (typeof rows === 'object' && Array.isArray(rows)) {
        if (rows.length > 0 && typeof rows[0] === 'object' && !Array.isArray(rows[0])) {
            // Array of objects
            rows = makeCSVFromData(rows);
        } else {
            // Array of arrays - convert to CSV
            rows = Papa.unparse(rows);
        }
    }

    try {
        // If tab is specified, first check and create if needed
        if (tab) {
            // Check if tab exists
            const spreadsheet = await retryWithBackoff(() => 
                sheets.spreadsheets.get({ spreadsheetId })
            );
            const existingSheets = spreadsheet.data.sheets?.map(( sheet) => sheet.properties.title) || [];

            // Create tab if it doesn't exist
            if (!existingSheets.includes(tab)) {
                logger.debug({ tab }, 'Creating new tab');
                await retryWithBackoff(() =>
                    sheets.spreadsheets.batchUpdate({
                        spreadsheetId,
                        resource: {
                            requests: [{
                                addSheet: {
                                    properties: {
                                        title: tab
                                    }
                                }
                            }]
                        }
                    })
                );
            }

            // Write to specific tab
            const response = await retryWithBackoff(() =>
                sheets.spreadsheets.values.update({
                    spreadsheetId,
                    range: `${tab}!A1`,
                    valueInputOption: 'USER_ENTERED',
                    resource: {
                        values: typeof rows === 'string' ? Papa.parse(rows).data : [],
                    },
                })
            );

            if (response?.data) {
                logger.info({ 
                    updatedCells: response.data.updatedCells, 
                    tab, 
                    spreadsheetId 
                }, 'Cells updated in tab');
                return response.data;
            }
        } else {
            // Original behavior if no tab specified
            const response = await retryWithBackoff(() =>
                sheets.spreadsheets.values.update({
                    spreadsheetId,
                    range: `A1`,
                    valueInputOption: 'USER_ENTERED',
                    resource: {
                        values: typeof rows === 'string' ? Papa.parse(rows).data : [],
                    },
                })
            );

            if (response?.data) {
                logger.info({ 
                    updatedCells: response.data.updatedCells, 
                    spreadsheetId 
                }, 'Cells updated');
                return response.data;
            }
        }
        
        return { updatedCells: 0 };
    } catch (error) {
        if ( (error).code === 404) {
            logger.warn({ spreadsheetId }, 'Spreadsheet not found, creating new one');
            const newSpreadsheetId = await createSheet();
            return await writeToSheet(newSpreadsheetId, rows, tab);
        }

        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            tab 
        }, 'Failed to write to sheet');
        throw error;
    }
}

/**
 * Writes data to multiple tabs in a Google Spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {Record<string, import('./index.d.ts').SpreadsheetData>} [assets={}] - Object with tab names as keys and data as values
 * @returns {Promise<import('./index.d.ts').SheetResponse[]>} Array of API responses
 * @example
 * import { writeToSheetTabs } from 'ak-sheets';
 * 
 * const multiTabData = {
 *   Users: [{ name: 'John', role: 'Admin' }, { name: 'Jane', role: 'User' }],
 *   Products: [{ sku: 'A001', name: 'Widget', price: 29.99 }],
 *   Orders: 'OrderID,UserID,Total\n1,123,29.99\n2,456,59.98'
 * };
 * 
 * const results = await writeToSheetTabs(spreadsheetId, multiTabData);
 * console.log(`Updated ${results.length} tabs`);
 */
export async function writeToSheetTabs(spreadsheetId, assets = {}) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    const tabs = Object.keys(assets);
    logger.debug({ spreadsheetId, tabCount: tabs.length }, 'Writing to multiple tabs');

    // Get existing tabs
    const existingTabs = await listTabs(spreadsheetId);
    const existingTabNames = existingTabs.map(tab => tab.title);

    // Create missing tabs
    for (const tabName of tabs) {
        if (!existingTabNames.includes(tabName)) {
            logger.debug({ tabName }, 'Creating missing tab');
            await addTab(spreadsheetId, tabName);
        }
    }

    const results = [];
    for (const tab of tabs) {
        const rows = assets[tab];
        
        // Convert rows to proper format
        let processedRows = rows;
        if (typeof rows === 'object' && Array.isArray(rows)) {
            if (rows.length > 0 && typeof rows[0] === 'object' && !Array.isArray(rows[0])) {
                processedRows = makeCSVFromData(rows);
            } else {
                processedRows = Papa.unparse(rows);
            }
        }

        try {
            const response = await retryWithBackoff(() =>
                sheets.spreadsheets.values.update({
                    spreadsheetId,
                    range: `${tab}!A1`,
                    valueInputOption: 'USER_ENTERED',
                    resource: {
                        values: typeof processedRows === 'string' ? Papa.parse(processedRows).data : [],
                    },
                })
            );

            if (response?.data) {
                logger.debug({ 
                    updatedCells: response.data.updatedCells, 
                    tab 
                }, 'Tab updated');
                results.push(response.data);
            }

        } catch (error) {
            if ( (error).code === 404) {
                logger.warn({ spreadsheetId }, 'Spreadsheet not found, creating new one');
                const newSpreadsheetId = await createSheet();
                return await writeToSheetTabs(newSpreadsheetId, assets);
            }

            logger.error({ 
                error:  (error).message, 
                spreadsheetId, 
                tab 
            }, 'Failed to write to tab');
            throw error;
        }
    }

    logger.info({ tabCount: results.length, spreadsheetId }, 'All tabs updated');
    return results;
}

/**
 * Shares a Google Spreadsheet with a user
 * @param {string} spreadsheetId - ID of the spreadsheet to share
 * @param {import('./index.d.ts').ShareOptions} [options] - Sharing options
 * @returns {Promise<any>} Promise resolving to the sharing result
 * @example
 * import { shareSheet } from 'ak-sheets';
 * 
 * // Share with default settings (writer access)
 * await shareSheet(spreadsheetId, { userEmail: 'colleague@company.com' });
 * 
 * // Share with read-only access
 * await shareSheet(spreadsheetId, {
 *   userEmail: 'viewer@company.com',
 *   role: 'reader'
 * });
 * 
 * // Share with anyone (public)
 * await shareSheet(spreadsheetId, {
 *   userEmail: '',
 *   type: 'anyone',
 *   role: 'reader'
 * });
 */
export async function shareSheet(spreadsheetId, options) {
    if (!drive) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    const {
        userEmail = 'aaron.krivitzky@mixpanel.com',
        role = 'writer',
        type = 'user'
    } = options || {};

    logger.debug({ spreadsheetId, userEmail, role, type }, 'Sharing spreadsheet');

    try {
        const result = await retryWithBackoff(() =>
            drive.permissions.create({
                fileId: spreadsheetId,
                requestBody: {
                    role,
                    type,
                    emailAddress: userEmail,
                },
            })
        );

        logger.info({ spreadsheetId, userEmail }, 'Spreadsheet shared successfully');
        return result;
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            userEmail 
        }, 'Failed to share spreadsheet');
        throw error;
    }
}

/**
 * Deletes a Google Spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet to delete
 * @returns {Promise<void>} Promise that resolves when deletion is complete
 * @example
 * import { deleteSheet } from 'ak-sheets';
 * 
 * // Delete a specific spreadsheet
 * await deleteSheet('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms');
 * console.log('Spreadsheet deleted successfully');
 */
export async function deleteSheet(spreadsheetId) {
    if (!drive) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId }, 'Deleting spreadsheet');

    try {
        await retryWithBackoff(() =>
            drive.files.delete({
                fileId: spreadsheetId,
            })
        );

        logger.info({ spreadsheetId }, 'Spreadsheet deleted successfully');
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId 
        }, 'Failed to delete spreadsheet');
        throw error;
    }
}

/**
 * Generates a Google Sheets URL for the given spreadsheet ID
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @returns {string} The Google Sheets URL
 * @example
 * import { getURL } from 'ak-sheets';
 * 
 * const url = getURL('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms');
 * console.log(url);
 * // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
 * 
 * // Open in browser
 * console.log(`View spreadsheet: ${url}`);
 */
export function getURL(spreadsheetId) {
    return `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
}

/**
 * Lists all spreadsheets owned by the authenticated user
 * @returns {Promise<import('./index.d.ts').SpreadsheetFile[]>} Promise resolving to array of spreadsheet metadata
 * @example
 * import { listOwnedSpreadsheets } from 'ak-sheets';
 * 
 * const sheets = await listOwnedSpreadsheets();
 * console.log(`You have ${sheets.length} spreadsheets`);
 * 
 * sheets.forEach(sheet => {
 *   console.log(`- ${sheet.name} (${sheet.id})`);
 * });
 */
export async function listOwnedSpreadsheets() {
    if (!drive) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug('Listing owned spreadsheets');

    try {
        const spreadsheets = [];
        let nextPageToken = null;

        do {
            
            const response = await retryWithBackoff(() =>
                drive.files.list({
                    q: "mimeType='application/vnd.google-apps.spreadsheet' and 'me' in owners",
                    fields: 'nextPageToken, files(id, name, owners)',
                    pageSize: 100,
                    pageToken: nextPageToken,
                })
            );

            spreadsheets.push(...(response?.data?.files || []));
            nextPageToken = response.data.nextPageToken;
        } while (nextPageToken);

        logger.debug({ count: spreadsheets.length }, 'Listed owned spreadsheets');
        return spreadsheets;
    } catch (error) {
        logger.error({ error:  (error).message }, 'Failed to list spreadsheets');
        throw error;
    }
}

/**
 * Deletes all spreadsheets owned by the authenticated user
 * @returns {Promise<import('./index.d.ts').SpreadsheetFile[]>} Promise resolving to array of deleted spreadsheet metadata
 * @example
 * import { deleteAllSheets } from 'ak-sheets';
 * 
 * // WARNING: This deletes ALL your spreadsheets!
 * const deletedSheets = await deleteAllSheets();
 * console.log(`Deleted ${deletedSheets.length} spreadsheets`);
 */
export async function deleteAllSheets() {
    logger.warn('Deleting all owned spreadsheets');

    try {
        const spreadsheets = await listOwnedSpreadsheets();
        const promises = spreadsheets.map((file) => deleteSheet(file.id));
        await Promise.all(promises);
        
        logger.info({ deletedCount: promises.length }, 'All spreadsheets deleted');
        return spreadsheets;
    } catch (error) {
        logger.error({ error:  (error).message }, 'Failed to delete all spreadsheets');
        throw error;
    }
}

/**
 * Converts array of objects to CSV string
 * @param {Record<string, any>[]} data - Array of objects to convert
 * @param {number} [charLimit=50000] - Maximum character limit per cell
 * @returns {string} CSV string representation
 * @example
 * import { makeCSVFromData } from 'ak-sheets';
 * 
 * const data = [
 *   { name: 'John', age: 30, city: 'NYC' },
 *   { name: 'Jane', age: 25, city: 'SF' }
 * ];
 * 
 * const csv = makeCSVFromData(data);
 * console.log(csv);
 * // name,age,city
 * // "John","30","NYC"
 * // "Jane","25","SF"
 */
export function makeCSVFromData(data, charLimit = 50000) {
    logger.debug({ dataLength: data?.length || 0, charLimit }, 'Converting data to CSV');

    // Handle empty data case
    if (!data || data.length === 0) return '';

    // Get all unique keys across all objects
    const columns = getUniqueKeys(data);

    // Create header row
    let csvString = columns.join(',') + '\n';

    // Process each data item
    data.forEach(item => {
        const row = columns.map(col => {
            // Handle undefined or null values
            if (item[col] === undefined || item[col] === null) return '';

            // Convert complex types to safe string representations
            const value = convertToSafeValue(item[col])
                ?.toString()
                ?.trim()
                ?.substring(0, charLimit);
            // Escape CSV-special characters
            return `"${value.toString().replace(/"/g, '""')}"`;
        }).join(',');

        csvString += row + '\n';
    });

    logger.debug({ csvLength: csvString.length }, 'CSV conversion completed');
    return csvString;
}

/**
 * Converts a value to a safe CSV-compatible string
 * @param {any} value - The value to convert
 * @returns {string} Safe string representation
 */
function convertToSafeValue(value) {
    // Handle different types of values
    if (value === null || value === undefined) return '';

    if (typeof value === 'object') {
        // For arrays or objects, use JSON.stringify with single quotes
        return JSON.stringify(value, null, 0)
            .replace(/"/g, "'");
    }

    // For primitive types, return as is
    return value;
}

/**
 * Gets unique keys from an array of objects
 * @param {Record<string, any>[]} data - Array of objects
 * @returns {string[]} Array of unique keys
 */
function getUniqueKeys(data) {
    const keysSet = new Set();
    data.forEach(item => {
        Object.keys(item).forEach(key => keysSet.add(key));
    });
    return Array.from(keysSet);
}

/**
 * Reads data from a Google Spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet to read from
 * @param {string} [tab] - Optional tab name to read from
 * @param {string} [format='json'] - Output format ('json', 'csv', or 'array')
 * @returns {Promise<any>} Promise resolving to the spreadsheet data in requested format
 * @example
 * import { getSheet } from 'ak-sheets';
 * 
 * // Get as array of objects (default)
 * const jsonData = await getSheet(spreadsheetId);
 * console.log(jsonData[0].name); // 'John'
 * 
 * // Get specific tab as CSV
 * const csvData = await getSheet(spreadsheetId, 'Users', 'csv');
 * 
 * // Get as 2D array
 * const arrayData = await getSheet(spreadsheetId, 'Products', 'array');
 * console.log(arrayData[0]); // ['Name', 'Price', 'Stock']
 */
export async function getSheet(spreadsheetId, tab, format = 'json') {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, tab, format }, 'Reading sheet data');

    try {
        const response = await retryWithBackoff(() =>
            sheets.spreadsheets.values.get({
                spreadsheetId,
                range: tab ? `${tab}!A:ZZ` : 'A:ZZ',
                majorDimension: 'ROWS'
            })
        );

        const values = response?.data?.values || [];
        logger.debug({ rowCount: values.length }, 'Sheet data retrieved');

        switch (format.toLowerCase()) {
            case 'csv':
                return makeCSVFromData(convertValuesToObjects(values));
            case 'array':
                return values;
            case 'json':
            default:
                return convertValuesToObjects(values);
        }
    } catch (error) {
        if ( (error).code === 404) {
            logger.error({ spreadsheetId }, 'Spreadsheet not found');
            throw error;
        }
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            tab 
        }, 'Failed to read sheet');
        throw error;
    }
}

/**
 * Merges existing data with new data
 * @param {any[][]} existingData - Existing spreadsheet data
 * @param {import('./index.d.ts').SpreadsheetData} newData - New data to merge
 * @returns {any[][]} Merged data
 */
function mergeData(existingData, newData) {
    // If newData is a CSV or string, convert to arrays
    if (typeof newData === 'string') {
        newData = Papa.parse(newData).data;
    }

    const headers = existingData[0];
    
    // Convert newData to array of arrays if it's JSON
    const processedNewData = Array.isArray(newData) && Array.isArray(newData[0]) 
        ? newData 
        : Array.isArray(newData) ? /** @type {any[][]} */ (newData).map(( item) => headers.map(header => item[header] || '')) : [];

    // Create mergedData starting with headers
    const mergedData = [headers];

    // Add processed new data rows
    for (let i = 0; i < existingData.length; i++) {
        if (i === 0) continue; // Skip header row processing
        if (i <= processedNewData.length) {
            // Replace rows with new data
            mergedData.push(/** @type {any[]} */ (processedNewData)[i - 1]);
        } else {
            // Keep existing rows after new data
            mergedData.push(existingData[i]);
        }
    }

    return mergedData;
}

/**
 * Updates existing data in a Google Spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet to update
 * @param {import('./index.d.ts').SpreadsheetData} newData - New data to merge/update
 * @param {string} [tab] - Tab name to update
 * @returns {Promise<import('./index.d.ts').SheetResponse>} Promise resolving to the API response
 * @example
 * import { updateSheet } from 'ak-sheets';
 * 
 * // Update with new data (merges with existing)
 * const updatedData = [
 *   { name: 'John', age: 31 }, // Updated age
 *   { name: 'Bob', age: 35 }   // New row
 * ];
 * 
 * const result = await updateSheet(spreadsheetId, updatedData, 'Users');
 * console.log(`Updated ${result.updatedCells} cells`);
 */
export async function updateSheet(spreadsheetId, newData, tab) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, tab }, 'Updating sheet data');

    try {
        // First, get existing sheet data
        const existingData = await getSheet(spreadsheetId, tab, 'array');
        
        // Determine which rows are new or different
        const updatedValues = mergeData(existingData, newData);

        // Write the updated data
        const response = await retryWithBackoff(() =>
            sheets.spreadsheets.values.update({
                spreadsheetId,
                range: `${tab}!A1`, 
                valueInputOption: 'USER_ENTERED',
                resource: {
                    values: updatedValues,
                },
            })
        );

        logger.info({ 
            updatedCells: response?.data?.updatedCells,
            spreadsheetId,
            tab
        }, 'Sheet updated successfully');
        return response.data;
    } catch (error) {
        if ( (error).code === 404) {
            logger.warn({ spreadsheetId, tab }, 'Spreadsheet or tab not found, creating new');
            const newSpreadsheetId = await createSheet(undefined, tab ? [tab] : undefined);
            return await writeToSheet(newSpreadsheetId, newData, tab);
        }
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            tab 
        }, 'Failed to update sheet');
        throw error;
    }
}

/**
 * Converts 2D array to array of objects
 * @param {any[][]} values - 2D array from sheets
 * @returns {Record<string, any>[]} Array of objects
 */
function convertValuesToObjects(values) {
    if (!values || values.length === 0) return [];
    
    const headers = values[0];
    return values.slice(1).map(row => 
        headers.reduce((obj, header, index) => {
            obj[header] = row[index] || '';
            return obj;
        }, {})
    );
}

/**
 * Reads an Excel (.xlsx) file and returns data as object with sheet names as keys
 * @param {string} filePath - Path to the Excel file
 * @returns {Record<string, string>} Object with sheet names as keys and CSV strings as values
 * @example
 * import { readXlsxFile } from 'ak-sheets';
 * 
 * const excelData = readXlsxFile('./data.xlsx');
 * console.log(excelData);
 * // { 'Sheet1': 'Name,Age\nJohn,30\nJane,25', 'Sheet2': 'Product,Price\nWidget,29.99' }
 */
export function readXlsxFile(filePath) {
    logger.debug({ filePath }, 'Reading Excel file');
    
    const resolvedPath = resolve(filePath);
    
    if (!existsSync(resolvedPath)) {
        throw new Error(`Excel file not found: ${resolvedPath}`);
    }

    try {
        const workbook = xlsx.readFile(resolvedPath);
        /** @type {Record<string, string>} */
        const sheets = {};
        
        workbook.SheetNames.forEach((/** @type {string} */ sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const csvData = xlsx.utils.sheet_to_csv(worksheet);
            sheets[/** @type {string} */ sheetName] = csvData;
        });
        
        logger.debug({ 
            filePath, 
            sheetCount: Object.keys(sheets).length,
            sheetNames: Object.keys(sheets)
        }, 'Excel file read successfully');
        
        return sheets;
    } catch ( error) {
        logger.error({ error: error.message, filePath }, 'Failed to read Excel file');
        throw new Error(`Failed to read Excel file: ${error.message}`);
    }
}

/**
 * Converts CSV string to array of objects (JSON)
 * @param {string} csvString - CSV string to convert
 * @param {Object} [options={}] - Papa Parse options
 * @returns {Record<string, any>[]} Array of objects
 * @example
 * import { csvToJson } from 'ak-sheets';
 * 
 * const csv = 'Name,Age\nJohn,30\nJane,25';
 * const json = csvToJson(csv);
 * console.log(json);
 * // [{ Name: 'John', Age: '30' }, { Name: 'Jane', Age: '25' }]
 */
export function csvToJson(csvString, options = {}) {
    logger.debug({ csvLength: csvString.length }, 'Converting CSV to JSON');
    
    const defaultOptions = {
        header: true,
        skipEmptyLines: true,
        transformHeader: (/** @type {string} */ header) => header.trim(),
        ...options
    };
    
    try {
        
        const result = Papa.parse(csvString, defaultOptions);
        
        // @ts-ignore
        if (result.errors.length > 0) {
            // @ts-ignore
            logger.warn({ errors: result.errors }, 'CSV parsing encountered errors');
        }
        
        logger.debug({ 
            // @ts-ignore
            rowCount: result.data.length,
            // @ts-ignore
            columnCount: result.meta.fields?.length || 0
        }, 'CSV converted to JSON successfully');
        
        // @ts-ignore
        return result.data;
    } catch ( error) {
        logger.error({ error: error.message }, 'Failed to convert CSV to JSON');
        throw new Error(`Failed to parse CSV: ${error.message}`);
    }
}

/**
 * Converts array of objects (JSON) to CSV string
 * @param {Record<string, any>[]} jsonData - Array of objects to convert
 * @param {Object} [options={}] - Papa Parse options
 * @returns {string} CSV string
 * @example
 * import { jsonToCsv } from 'ak-sheets';
 * 
 * const json = [{ Name: 'John', Age: 30 }, { Name: 'Jane', Age: 25 }];
 * const csv = jsonToCsv(json);
 * console.log(csv);
 * // 'Name,Age\nJohn,30\nJane,25'
 */
export function jsonToCsv(jsonData, options = {}) {
    logger.debug({ rowCount: jsonData.length }, 'Converting JSON to CSV');
    
    if (!Array.isArray(jsonData) || jsonData.length === 0) {
        return '';
    }
    
    const defaultOptions = {
        header: true,
        ...options
    };
    
    try {
        const csvString = Papa.unparse(jsonData, defaultOptions);
        
        logger.debug({ 
            csvLength: csvString.length,
            rowCount: jsonData.length
        }, 'JSON converted to CSV successfully');
        
        return csvString;
    } catch ( error) {
        logger.error({ error: error.message }, 'Failed to convert JSON to CSV');
        throw new Error(`Failed to convert JSON to CSV: ${error.message}`);
    }
}

/**
 * Appends data to an existing spreadsheet without overwriting
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {import('./index.d.ts').SpreadsheetData} rows - Data to append
 * @param {string} [tab] - Optional tab name to append to
 * @returns {Promise<import('./index.d.ts').SheetResponse>} Promise resolving to the API response
 * @example
 * import { appendToSheet } from 'ak-sheets';
 * 
 * const newData = [{ name: 'Bob', age: 35 }];
 * const result = await appendToSheet(spreadsheetId, newData, 'Users');
 * console.log(`Appended ${result.updatedCells} cells`);
 */
export async function appendToSheet(spreadsheetId, rows, tab) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, tab, dataType: typeof rows }, 'Appending to sheet');

    try {
        // Get existing data to find the next empty row
        const existingData = await getSheet(spreadsheetId, tab, 'array');
        const nextRow = existingData.length + 1;
        
        // Convert rows to proper format for appending
        let processedRows = rows;
        if (typeof rows === 'object' && Array.isArray(rows)) {
            if (rows.length > 0 && typeof rows[0] === 'object' && !Array.isArray(rows[0])) {
                // Array of objects - convert to array of arrays using existing headers
                const headers = existingData[0] || [];
                processedRows = rows.map(( obj) => headers.map((/** @type {string} */ header) => obj[header] || ''));
            } else if (typeof rows === 'string') {
                processedRows = Papa.parse(rows).data;
            }
        }

        const range = tab ? `${tab}!A${nextRow}` : `A${nextRow}`;
        const response = await retryWithBackoff(() =>
            sheets.spreadsheets.values.append({
                spreadsheetId,
                range,
                valueInputOption: 'USER_ENTERED',
                insertDataOption: 'INSERT_ROWS',
                resource: {
                    values: Array.isArray(processedRows[0]) ? processedRows : Papa.parse(/** @type {string} */ (processedRows)).data,
                },
            })
        );

        if (response?.data) {
            logger.info({ 
                updatedCells: response.data.updatedCells, 
                spreadsheetId, 
                tab 
            }, 'Data appended successfully');
            return response.data;
        }
        
        return { updatedCells: 0 };
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            tab 
        }, 'Failed to append to sheet');
        throw error;
    }
}

/**
 * Clears all data from a spreadsheet or specific tab
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {string} [tab] - Optional tab name to clear (clears all tabs if not specified)
 * @returns {Promise<import('./index.d.ts').SheetResponse>} Promise resolving to the API response
 * @example
 * import { clearSheet } from 'ak-sheets';
 * 
 * // Clear specific tab
 * await clearSheet(spreadsheetId, 'Users');
 * 
 * // Clear entire spreadsheet
 * await clearSheet(spreadsheetId);
 */
export async function clearSheet(spreadsheetId, tab) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, tab }, 'Clearing sheet data');

    try {
        const range = tab ? `${tab}!A:ZZ` : 'A:ZZ';
        const response = await retryWithBackoff(() =>
            sheets.spreadsheets.values.clear({
                spreadsheetId,
                range,
            })
        );

        logger.info({ spreadsheetId, tab }, 'Sheet cleared successfully');
        return response.data || { clearedRange: range };
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            tab 
        }, 'Failed to clear sheet');
        throw error;
    }
}

/**
 * Gets information about a spreadsheet (metadata, sheets, etc.)
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @returns {Promise<any>} Promise resolving to spreadsheet metadata
 * @example
 * import { getSheetInfo } from 'ak-sheets';
 * 
 * const info = await getSheetInfo(spreadsheetId);
 * console.log(`Spreadsheet: ${info.properties.title}`);
 * console.log(`Sheets: ${info.sheets.map(s => s.properties.title).join(', ')}`);
 */
export async function getSheetInfo(spreadsheetId) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId }, 'Getting sheet info');

    try {
        const response = await retryWithBackoff(() =>
            sheets.spreadsheets.get({
                spreadsheetId,
                fields: 'properties,sheets.properties'
            })
        );

        logger.debug({ 
            spreadsheetId, 
            title: response.data.properties?.title,
            sheetCount: response.data.sheets?.length || 0
        }, 'Sheet info retrieved');

        return response.data;
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId 
        }, 'Failed to get sheet info');
        throw error;
    }
}

/**
 * Reads data from a specific range in a Google Spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet to read from
 * @param {string} range - Range to read (e.g., 'A1:C10', 'B:B', 'A1:Z')
 * @param {string} [tab] - Optional tab name
 * @param {string} [format='json'] - Output format ('json', 'csv', or 'array')
 * @returns {Promise<any>} Promise resolving to the range data in requested format
 * @example
 * import { getRange } from 'ak-sheets';
 * 
 * // Get specific range as JSON
 * const data = await getRange(spreadsheetId, 'A1:C10');
 * 
 * // Get entire column B from specific tab
 * const columnB = await getRange(spreadsheetId, 'B:B', 'Users', 'array');
 * 
 * // Get range from specific tab as CSV
 * const csvData = await getRange(spreadsheetId, 'A1:E5', 'Products', 'csv');
 */
export async function getRange(spreadsheetId, range, tab, format = 'json') {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, range, tab, format }, 'Reading range data');

    try {
        const fullRange = tab ? `${tab}!${range}` : range;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: fullRange,
            majorDimension: 'ROWS'
        });

        const values = response?.data?.values || [];
        logger.debug({ rowCount: values.length, range: fullRange }, 'Range data retrieved');

        switch (format.toLowerCase()) {
            case 'csv':
                return makeCSVFromData(convertValuesToObjects(values));
            case 'array':
                return values;
            case 'json':
            default:
                return convertValuesToObjects(values);
        }
    } catch (error) {
        if ( (error).code === 404) {
            logger.error({ spreadsheetId, range }, 'Spreadsheet or range not found');
            throw error;
        }
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            range,
            tab 
        }, 'Failed to read range');
        throw error;
    }
}

/**
 * Writes data to a specific range in a Google Spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {string} range - Range to write to (e.g., 'A1:C10', 'B2', 'A1')
 * @param {import('./index.d.ts').SpreadsheetData} data - Data to write
 * @param {string} [tab] - Optional tab name
 * @returns {Promise<import('./index.d.ts').SheetResponse>} Promise resolving to the API response
 * @example
 * import { writeToRange } from 'ak-sheets';
 * 
 * // Write to specific range
 * const data = [['Name', 'Age'], ['John', 30], ['Jane', 25]];
 * await writeToRange(spreadsheetId, 'A1:B3', data);
 * 
 * // Write single value to cell
 * await writeToRange(spreadsheetId, 'D1', 'Total');
 * 
 * // Write to range in specific tab
 * await writeToRange(spreadsheetId, 'A1:C2', data, 'Users');
 */
export async function writeToRange(spreadsheetId, range, data, tab) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, range, tab, dataType: typeof data }, 'Writing to range');

    try {
        // Convert data to proper format
        let processedData = data;
        if (typeof data === 'string') {
            // Single value or CSV string
            if (data.includes('\n') || data.includes(',')) {
                processedData = Papa.parse(data).data;
            } else {
                processedData = [[data]]; // Single cell value
            }
        } else if (Array.isArray(data)) {
            if (data.length > 0 && typeof data[0] === 'object' && !Array.isArray(data[0])) {
                // Array of objects - convert to CSV then parse
                const csvString = makeCSVFromData(data);
                processedData = Papa.parse(csvString).data;
            }
            // Array of arrays - use as is
        }

        const fullRange = tab ? `${tab}!${range}` : range;
        const response = await retryWithBackoff(() =>
            sheets.spreadsheets.values.update({
                spreadsheetId,
                range: fullRange,
                valueInputOption: 'USER_ENTERED',
                resource: {
                    values: processedData,
                },
            })
        );

        if (response?.data) {
            logger.info({ 
                updatedCells: response.data.updatedCells, 
                range: fullRange,
                spreadsheetId 
            }, 'Range updated successfully');
            return response.data;
        }
        
        return { updatedCells: 0 };
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            range,
            tab 
        }, 'Failed to write to range');
        throw error;
    }
}

/**
 * Adds a new tab to an existing spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {string} tabName - Name of the new tab
 * @param {Object} [options={}] - Tab creation options
 * @param {number} [options.index] - Position to insert tab (0-based)
 * @param {boolean} [options.hidden=false] - Whether tab should be hidden
 * @param {string} [options.tabColor] - Hex color for tab (e.g., '#FF0000')
 * @returns {Promise<number>} Promise resolving to the new tab's sheet ID
 * @example
 * import { addTab } from 'ak-sheets';
 * 
 * // Add simple tab
 * await addTab(spreadsheetId, 'New Data');
 * 
 * // Add tab with options
 * await addTab(spreadsheetId, 'Reports', {
 *   index: 1,
 *   tabColor: '#FF0000',
 *   hidden: false
 * });
 */
export async function addTab(spreadsheetId, tabName, options = {}) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    const { index, hidden = false, tabColor } = options;
    logger.debug({ spreadsheetId, tabName, options }, 'Adding new tab');

    try {
        // Check if tab already exists
        const info = await getSheetInfo(spreadsheetId);
        const existingTabs = info.sheets?.map(sheet => sheet.properties.title) || [];
        
        if (existingTabs.includes(tabName)) {
            logger.warn({ spreadsheetId, tabName }, 'Tab already exists');
            throw new Error(`Tab '${tabName}' already exists in spreadsheet`);
        }

        const addSheetRequest = {
            addSheet: {
                properties: {
                    title: tabName,
                    index,
                    hidden,
                    ...(tabColor && {
                        tabColor: {
                            red: parseInt(tabColor.slice(1, 3), 16) / 255,
                            green: parseInt(tabColor.slice(3, 5), 16) / 255,
                            blue: parseInt(tabColor.slice(5, 7), 16) / 255
                        }
                    })
                }
            }
        };

        const response = await retryWithBackoff(() =>
            sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: {
                    requests: [addSheetRequest]
                }
            })
        );

        const newSheetId = response.data.replies?.[0]?.addSheet?.properties?.sheetId;
        logger.info({ 
            spreadsheetId, 
            tabName, 
            sheetId: newSheetId 
        }, 'Tab added successfully');

        return newSheetId;
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            tabName 
        }, 'Failed to add tab');
        throw error;
    }
}

/**
 * Deletes a tab from a spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {string} tabName - Name of the tab to delete
 * @returns {Promise<void>} Promise that resolves when deletion is complete
 * @example
 * import { deleteTab } from 'ak-sheets';
 * 
 * await deleteTab(spreadsheetId, 'Old Data');
 */
export async function deleteTab(spreadsheetId, tabName) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, tabName }, 'Deleting tab');

    try {
        // Get sheet info to find the sheet ID
        const info = await getSheetInfo(spreadsheetId);
        const targetSheet = info.sheets?.find(sheet => sheet.properties.title === tabName);
        
        if (!targetSheet) {
            throw new Error(`Tab '${tabName}' not found in spreadsheet`);
        }

        const sheetId = targetSheet.properties.sheetId;
        const deleteSheetRequest = {
            deleteSheet: {
                sheetId: sheetId
            }
        };

        await retryWithBackoff(() =>
            sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: {
                    requests: [deleteSheetRequest]
                }
            })
        );

        logger.info({ spreadsheetId, tabName, sheetId }, 'Tab deleted successfully');
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            tabName 
        }, 'Failed to delete tab');
        throw error;
    }
}

/**
 * Renames a tab in a spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {string} oldName - Current name of the tab
 * @param {string} newName - New name for the tab
 * @returns {Promise<void>} Promise that resolves when rename is complete
 * @example
 * import { renameTab } from 'ak-sheets';
 * 
 * await renameTab(spreadsheetId, 'Sheet1', 'User Data');
 */
export async function renameTab(spreadsheetId, oldName, newName) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, oldName, newName }, 'Renaming tab');

    try {
        // Get sheet info to find the sheet ID
        const info = await getSheetInfo(spreadsheetId);
        const targetSheet = info.sheets?.find(sheet => sheet.properties.title === oldName);
        
        if (!targetSheet) {
            throw new Error(`Tab '${oldName}' not found in spreadsheet`);
        }

        // Check if new name already exists
        const existingTabs = info.sheets?.map(sheet => sheet.properties.title) || [];
        if (existingTabs.includes(newName)) {
            throw new Error(`Tab '${newName}' already exists in spreadsheet`);
        }

        const sheetId = targetSheet.properties.sheetId;
        const updatePropertiesRequest = {
            updateSheetProperties: {
                properties: {
                    sheetId: sheetId,
                    title: newName
                },
                fields: 'title'
            }
        };

        await retryWithBackoff(() =>
            sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: {
                    requests: [updatePropertiesRequest]
                }
            })
        );

        logger.info({ spreadsheetId, oldName, newName, sheetId }, 'Tab renamed successfully');
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            oldName, 
            newName 
        }, 'Failed to rename tab');
        throw error;
    }
}

/**
 * Duplicates a tab within a spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @param {string} sourceTabName - Name of the tab to duplicate
 * @param {string} [newTabName] - Name for the duplicated tab (auto-generated if not provided)
 * @returns {Promise<{sheetId: number, title: string}>} Promise resolving to new tab info
 * @example
 * import { duplicateTab } from 'ak-sheets';
 * 
 * // Duplicate with auto-generated name
 * const newTab = await duplicateTab(spreadsheetId, 'Template');
 * 
 * // Duplicate with custom name
 * await duplicateTab(spreadsheetId, 'Template', 'January Data');
 */
export async function duplicateTab(spreadsheetId, sourceTabName, newTabName) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId, sourceTabName, newTabName }, 'Duplicating tab');

    try {
        // Get sheet info to find the source sheet ID
        const info = await getSheetInfo(spreadsheetId);
        const sourceSheet = info.sheets?.find(sheet => sheet.properties.title === sourceTabName);
        
        if (!sourceSheet) {
            throw new Error(`Source tab '${sourceTabName}' not found in spreadsheet`);
        }

        // Generate new name if not provided
        const finalNewName = newTabName || `Copy of ${sourceTabName}`;
        
        // Check if new name already exists
        const existingTabs = info.sheets?.map(sheet => sheet.properties.title) || [];
        if (existingTabs.includes(finalNewName)) {
            throw new Error(`Tab '${finalNewName}' already exists in spreadsheet`);
        }

        const sourceSheetId = sourceSheet.properties.sheetId;
        const duplicateSheetRequest = {
            duplicateSheet: {
                sourceSheetId: sourceSheetId,
                newSheetName: finalNewName
            }
        };

        const response = await retryWithBackoff(() =>
            sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: {
                    requests: [duplicateSheetRequest]
                }
            })
        );

        const newSheetInfo = response.data.replies?.[0]?.duplicateSheet?.properties;
        logger.info({ 
            spreadsheetId, 
            sourceTabName, 
            newTabName: finalNewName,
            newSheetId: newSheetInfo?.sheetId
        }, 'Tab duplicated successfully');

        return {
            sheetId: newSheetInfo?.sheetId,
            title: newSheetInfo?.title
        };
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId, 
            sourceTabName, 
            newTabName 
        }, 'Failed to duplicate tab');
        throw error;
    }
}

/**
 * Lists all tabs in a spreadsheet
 * @param {string} spreadsheetId - ID of the spreadsheet
 * @returns {Promise<Array<{id: number, title: string, index: number, hidden: boolean}>>} Promise resolving to array of tab info
 * @example
 * import { listTabs } from 'ak-sheets';
 * 
 * const tabs = await listTabs(spreadsheetId);
 * console.log(tabs);
 * // [
 * //   { id: 0, title: 'Sheet1', index: 0, hidden: false },
 * //   { id: 123, title: 'Users', index: 1, hidden: false }
 * // ]
 */
export async function listTabs(spreadsheetId) {
    if (!sheets) {
        throw new Error('ak-sheets not initialized. Call init() first.');
    }

    logger.debug({ spreadsheetId }, 'Listing tabs');

    try {
        const info = await getSheetInfo(spreadsheetId);
        const tabs = info.sheets?.map(sheet => ({
            id: sheet.properties.sheetId,
            title: sheet.properties.title,
            index: sheet.properties.index,
            hidden: sheet.properties.hidden || false
        })) || [];

        logger.debug({ spreadsheetId, tabCount: tabs.length }, 'Tabs listed successfully');
        return tabs;
    } catch (error) {
        logger.error({ 
            error:  (error).message, 
            spreadsheetId 
        }, 'Failed to list tabs');
        throw error;
    }
}

/**
 * Default export object with all sheet operations
 * @example
 * import sheet from 'ak-sheets';
 * import { init } from 'ak-sheets';
 * 
 * // Initialize first
 * init({ credentials: './credentials.json' });
 * 
 * // Use default export methods
 * const id = await sheet.create('My Spreadsheet');
 * await sheet.write(id, [{ name: 'John', age: 30 }]);
 * const data = await sheet.get(id);
 * const url = sheet.url(id);
 * await sheet.share(id, { userEmail: 'user@example.com' });
 * 
 * // List all spreadsheets
 * const allSheets = await sheet.list();
 * console.log(`You have ${allSheets.length} spreadsheets`);
 */
const sheet = {
    get: getSheet,
    update: updateSheet,
    create: createSheet,
    write: writeToSheet,
    writeTabs: writeToSheetTabs,
    append: appendToSheet,
    clear: clearSheet,
    share: shareSheet,
    delete: deleteSheet,
    url: getURL,
    list: listOwnedSpreadsheets,
    info: getSheetInfo,
    // Range operations
    getRange: getRange,
    writeRange: writeToRange,
    // Tab management
    addTab: addTab,
    deleteTab: deleteTab,
    renameTab: renameTab,
    duplicateTab: duplicateTab,
    listTabs: listTabs,
};

export default sheet;