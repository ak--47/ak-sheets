import { Logger } from 'pino';

/**
 * Configuration options for ak-sheets
 */
export interface AkSheetsConfig {
  /** Google API credentials object or file path to credentials JSON file. Can be omitted if SHEETS_CREDENTIALS env var is set. */
  credentials?: any | string;
  /** Optional pino logger instance */
  logger?: Logger;
  /** Environment (dev/prod/test) - affects logging verbosity. Uses NODE_ENV if not provided. */
  environment?: 'dev' | 'prod' | 'test';
  /** Maximum number of retries for failed API calls (default: 5) */
  maxRetries?: number;
  /** Maximum backoff time in milliseconds (default: 64000) */
  maxBackoffMs?: number;
}

/**
 * Spreadsheet data in various formats
 */
export type SpreadsheetData = string | any[][] | Record<string, any>[];

/**
 * Response from Google Sheets API operations
 */
export interface SheetResponse {
  spreadsheetId?: string;
  updatedCells?: number;
  updatedColumns?: number;
  updatedRows?: number;
}

/**
 * Spreadsheet file metadata
 */
export interface SpreadsheetFile {
  id: string;
  name: string;
  owners?: Array<{ emailAddress: string; displayName: string }>;
}

/**
 * Options for sharing a spreadsheet
 */
export interface ShareOptions {
  /** Email address to share with */
  userEmail: string;
  /** Permission role */
  role?: 'reader' | 'writer' | 'owner';
  /** Permission type */
  type?: 'user' | 'domain' | 'anyone';
}

/**
 * Creates a new Google Spreadsheet
 * @param name - Name of the spreadsheet (optional, generates random name if not provided)
 * @param tabs - Array of tab names to create (optional)
 * @returns Promise resolving to the spreadsheet ID
 */
export function createSheet(name?: string, tabs?: string[]): Promise<string>;

/**
 * Writes data to a Google Spreadsheet
 * @param spreadsheetId - ID of the target spreadsheet
 * @param rows - Data to write (CSV string, array of arrays, or array of objects)
 * @param tab - Optional tab name to write to
 * @returns Promise resolving to the API response
 */
export function writeToSheet(spreadsheetId: string, rows?: SpreadsheetData, tab?: string): Promise<SheetResponse>;

/**
 * Writes data to multiple tabs in a Google Spreadsheet
 * @param spreadsheetId - ID of the target spreadsheet
 * @param assets - Object with tab names as keys and data as values
 * @returns Promise resolving to array of API responses
 */
export function writeToSheetTabs(spreadsheetId: string, assets?: Record<string, SpreadsheetData>): Promise<SheetResponse[]>;

/**
 * Shares a Google Spreadsheet with a user
 * @param spreadsheetId - ID of the spreadsheet to share
 * @param options - Sharing options (email, role, type)
 * @returns Promise resolving to the sharing result
 */
export function shareSheet(spreadsheetId: string, options?: ShareOptions): Promise<any>;

/**
 * Deletes a Google Spreadsheet
 * @param spreadsheetId - ID of the spreadsheet to delete
 * @returns Promise that resolves when deletion is complete
 */
export function deleteSheet(spreadsheetId: string): Promise<void>;

/**
 * Generates a Google Sheets URL for the given spreadsheet ID
 * @param spreadsheetId - ID of the spreadsheet
 * @returns The Google Sheets URL
 */
export function getURL(spreadsheetId: string): string;

/**
 * Lists all spreadsheets owned by the authenticated user
 * @returns Promise resolving to array of spreadsheet metadata
 */
export function listOwnedSpreadsheets(): Promise<SpreadsheetFile[]>;

/**
 * Deletes all spreadsheets owned by the authenticated user
 * @returns Promise resolving to array of deleted spreadsheet metadata
 */
export function deleteAllSheets(): Promise<SpreadsheetFile[]>;

/**
 * Converts array of objects to CSV string
 * @param data - Array of objects to convert
 * @param charLimit - Maximum character limit per cell (default: 50000)
 * @returns CSV string representation
 */
export function makeCSVFromData(data: Record<string, any>[], charLimit?: number): string;

/**
 * Reads data from a Google Spreadsheet
 * @param spreadsheetId - ID of the spreadsheet to read from
 * @param tab - Optional tab name to read from
 * @param format - Output format ('json', 'csv', or 'array')
 * @returns Promise resolving to the spreadsheet data in requested format
 */
export function getSheet(spreadsheetId: string, tab?: string, format?: 'json' | 'csv' | 'array'): Promise<any>;

/**
 * Updates existing data in a Google Spreadsheet
 * @param spreadsheetId - ID of the spreadsheet to update
 * @param newData - New data to merge/update
 * @param tab - Tab name to update
 * @returns Promise resolving to the API response
 */
export function updateSheet(spreadsheetId: string, newData: SpreadsheetData, tab?: string): Promise<SheetResponse>;

/**
 * Reads an Excel (.xlsx) file and returns data as object with sheet names as keys
 * @param filePath - Path to the Excel file
 * @returns Object with sheet names as keys and CSV strings as values
 */
export function readXlsxFile(filePath: string): Record<string, string>;

/**
 * Converts CSV string to array of objects (JSON)
 * @param csvString - CSV string to convert
 * @param options - Papa Parse options
 * @returns Array of objects
 */
export function csvToJson(csvString: string, options?: any): Record<string, any>[];

/**
 * Converts array of objects (JSON) to CSV string
 * @param jsonData - Array of objects to convert
 * @param options - Papa Parse options
 * @returns CSV string
 */
export function jsonToCsv(jsonData: Record<string, any>[], options?: any): string;

/**
 * Appends data to an existing spreadsheet without overwriting
 * @param spreadsheetId - ID of the spreadsheet
 * @param rows - Data to append
 * @param tab - Optional tab name to append to
 * @returns Promise resolving to the API response
 */
export function appendToSheet(spreadsheetId: string, rows: SpreadsheetData, tab?: string): Promise<SheetResponse>;

/**
 * Clears all data from a spreadsheet or specific tab
 * @param spreadsheetId - ID of the spreadsheet
 * @param tab - Optional tab name to clear
 * @returns Promise resolving to the API response
 */
export function clearSheet(spreadsheetId: string, tab?: string): Promise<SheetResponse>;

/**
 * Gets information about a spreadsheet (metadata, sheets, etc.)
 * @param spreadsheetId - ID of the spreadsheet
 * @returns Promise resolving to spreadsheet metadata
 */
export function getSheetInfo(spreadsheetId: string): Promise<any>;

/**
 * Reads data from a specific range in a Google Spreadsheet
 * @param spreadsheetId - ID of the spreadsheet to read from
 * @param range - Range to read (e.g., 'A1:C10', 'B:B', 'A1:Z')
 * @param tab - Optional tab name
 * @param format - Output format ('json', 'csv', or 'array')
 * @returns Promise resolving to the range data in requested format
 */
export function getRange(spreadsheetId: string, range: string, tab?: string, format?: 'json' | 'csv' | 'array'): Promise<any>;

/**
 * Writes data to a specific range in a Google Spreadsheet
 * @param spreadsheetId - ID of the spreadsheet
 * @param range - Range to write to (e.g., 'A1:C10', 'B2', 'A1')
 * @param data - Data to write
 * @param tab - Optional tab name
 * @returns Promise resolving to the API response
 */
export function writeToRange(spreadsheetId: string, range: string, data: SpreadsheetData, tab?: string): Promise<SheetResponse>;

/**
 * Tab creation options
 */
export interface TabOptions {
  /** Position to insert tab (0-based) */
  index?: number;
  /** Whether tab should be hidden */
  hidden?: boolean;
  /** Hex color for tab (e.g., '#FF0000') */
  tabColor?: string;
}

/**
 * Tab information
 */
export interface TabInfo {
  /** Sheet ID */
  id: number;
  /** Tab title */
  title: string;
  /** Tab position (0-based) */
  index: number;
  /** Whether tab is hidden */
  hidden: boolean;
}

/**
 * Adds a new tab to an existing spreadsheet
 * @param spreadsheetId - ID of the spreadsheet
 * @param tabName - Name of the new tab
 * @param options - Tab creation options
 * @returns Promise resolving to the new tab's sheet ID
 */
export function addTab(spreadsheetId: string, tabName: string, options?: TabOptions): Promise<number>;

/**
 * Deletes a tab from a spreadsheet
 * @param spreadsheetId - ID of the spreadsheet
 * @param tabName - Name of the tab to delete
 * @returns Promise that resolves when deletion is complete
 */
export function deleteTab(spreadsheetId: string, tabName: string): Promise<void>;

/**
 * Renames a tab in a spreadsheet
 * @param spreadsheetId - ID of the spreadsheet
 * @param oldName - Current name of the tab
 * @param newName - New name for the tab
 * @returns Promise that resolves when rename is complete
 */
export function renameTab(spreadsheetId: string, oldName: string, newName: string): Promise<void>;

/**
 * Duplicates a tab within a spreadsheet
 * @param spreadsheetId - ID of the spreadsheet
 * @param sourceTabName - Name of the tab to duplicate
 * @param newTabName - Name for the duplicated tab (auto-generated if not provided)
 * @returns Promise resolving to new tab info
 */
export function duplicateTab(spreadsheetId: string, sourceTabName: string, newTabName?: string): Promise<{sheetId: number, title: string}>;

/**
 * Lists all tabs in a spreadsheet
 * @param spreadsheetId - ID of the spreadsheet
 * @returns Promise resolving to array of tab info
 */
export function listTabs(spreadsheetId: string): Promise<TabInfo[]>;

/**
 * Initialize ak-sheets with configuration
 * @param config - Configuration options
 */
export function init(config: AkSheetsConfig): void;

/**
 * Default export object with all sheet operations
 */
declare const sheet: {
  get: typeof getSheet;
  update: typeof updateSheet;
  create: typeof createSheet;
  write: typeof writeToSheet;
  writeTabs: typeof writeToSheetTabs;
  append: typeof appendToSheet;
  clear: typeof clearSheet;
  share: typeof shareSheet;
  delete: typeof deleteSheet;
  url: typeof getURL;
  list: typeof listOwnedSpreadsheets;
  info: typeof getSheetInfo;
  // Range operations
  getRange: typeof getRange;
  writeRange: typeof writeToRange;
  // Tab management
  addTab: typeof addTab;
  deleteTab: typeof deleteTab;
  renameTab: typeof renameTab;
  duplicateTab: typeof duplicateTab;
  listTabs: typeof listTabs;
};

export default sheet;