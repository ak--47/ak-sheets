# ak-sheets

A modern, simple interface to work deeply with Google Sheets data. Built with ESM, TypeScript support, and comprehensive logging.

## ✨ Features

- 🔧 **Modern ESM package** with full TypeScript support
- 📊 **Complete CRUD operations** for Google Sheets
- 📁 **Excel file support** - read .xlsx files directly
- 🔄 **Flexible data formats** - JSON, CSV, arrays
- 🚀 **Multiple initialization methods** - objects, files, environment variables
- 📝 **Comprehensive logging** with Pino
- 🧪 **Thoroughly tested** with integration tests
- 💡 **Excellent IntelliSense** with detailed JSDoc examples

## 🚀 Quick Start

```bash
npm install ak-sheets
```

```javascript
import { initSheets, createSheet, writeToSheet, getSheet } from 'ak-sheets';

// Initialize with credentials
await initSheets({
  credentials: './path/to/credentials.json',
  environment: 'dev'
});

// Create a spreadsheet
const spreadsheetId = await createSheet('My Data Sheet');

// Write data
const data = [
  { name: 'John', age: 30, city: 'NYC' },
  { name: 'Jane', age: 25, city: 'SF' }
];
await writeToSheet(spreadsheetId, data);

// Read data back
const readData = await getSheet(spreadsheetId);
console.log(readData); // [{ name: 'John', age: '30', city: 'NYC' }, ...]
```

## 📚 Documentation

### Initialization Options

```javascript
import { initSheets } from 'ak-sheets';

// Option 1: Credentials object
await initSheets({
  credentials: {
    type: "service_account",
    project_id: "your-project",
    // ... other credential fields
  }
});

// Option 2: File path
await initSheets({
  credentials: './credentials.json'
});

// Option 3: Environment variable
// Set SHEETS_CREDENTIALS=./credentials.json
await initSheets({});

// Option 4: Custom retry configuration
await initSheets({
  credentials: './credentials.json',
  maxRetries: 3,
  maxBackoffMs: 32000
});
```

### Core Functions

- **`createSheet(name?, tabs?)`** - Create new spreadsheets
- **`writeToSheet(id, data, tab?)`** - Write data to sheets
- **`getSheet(id, tab?, format?)`** - Read data from sheets
- **`updateSheet(id, data, tab?)`** - Update existing data
- **`appendToSheet(id, data, tab?)`** - Append without overwriting
- **`clearSheet(id, tab?)`** - Clear sheet data
- **`shareSheet(id, options?)`** - Share with users
- **`deleteSheet(id)`** - Delete spreadsheets

### Utility Functions

- **`readXlsxFile(path)`** - Read Excel files to CSV
- **`csvToJson(csv)`** - Convert CSV to JSON
- **`jsonToCsv(json)`** - Convert JSON to CSV
- **`makeCSVFromData(data)`** - Advanced CSV conversion
- **`getURL(id)`** - Generate Google Sheets URLs
- **`listOwnedSpreadsheets()`** - List your spreadsheets

### Default Export

```javascript
import sheet from 'ak-sheets';

const id = await sheet.create('My Sheet');
await sheet.write(id, data);
const readData = await sheet.get(id);
```

## 📊 Data Format Support

**Input formats:**
- Array of objects: `[{ name: 'John', age: 30 }]`
- CSV strings: `"Name,Age\nJohn,30"`
- Array of arrays: `[['Name', 'Age'], ['John', 30]]`

**Output formats:**
- `'json'` (default): Array of objects
- `'csv'`: CSV string
- `'array'`: 2D array

## 🔐 Authentication

1. Create a Google Cloud Project
2. Enable Google Sheets API and Google Drive API
3. Create a Service Account
4. Download credentials JSON file
5. Share your spreadsheets with the service account email

## 🧪 Environment Variables

- **`SHEETS_CREDENTIALS`** - Path to credentials file
- **`NODE_ENV`** - Environment (affects logging level and format)
- **`LOG_LEVEL`** - Logging level (fatal, error, warn, info, debug, trace)

### 📝 Logging

ak-sheets uses Pino for structured logging with environment-specific formatting:

- **Development/Test**: Pretty formatted logs with colors and timestamps
- **Production**: JSON formatted logs for cloud logging systems

```bash
# Set custom log level
export LOG_LEVEL=debug

# Development environment (pretty logs)
export NODE_ENV=development

# Production environment (JSON logs)
export NODE_ENV=production
```

## 🔄 Rate Limiting & Error Handling

ak-sheets implements robust exponential backoff retry logic to handle Google Sheets API quota limits:

- **Automatic retries** for quota exceeded (429) and server errors (500-504)
- **Exponential backoff** with jitter to avoid thundering herd
- **Configurable retry limits** and maximum backoff times
- **Detailed logging** of retry attempts and failures

### Quota Limits
- **Read requests**: 300/minute per project, 60/minute per user
- **Write requests**: 300/minute per project, 60/minute per user

### Retry Configuration
```javascript
await initSheets({
  credentials: './credentials.json',
  maxRetries: 5,        // Max retry attempts (default: 5)
  maxBackoffMs: 64000   // Max backoff time in ms (default: 64s)
});
```

The retry algorithm follows Google's recommended exponential backoff:
- **Base delay**: `2^n * 1000ms` (1s, 2s, 4s, 8s, 16s...)
- **Jitter**: Random 0-1000ms added to prevent synchronization
- **Max backoff**: Configurable ceiling (default 64 seconds)

## 📝 Examples

### Multi-tab Operations
```javascript
const multiTabData = {
  Users: [{ name: 'John', role: 'Admin' }],
  Products: [{ sku: 'A001', price: 29.99 }]
};
await writeToSheetTabs(spreadsheetId, multiTabData);
```

### Excel Integration
```javascript
const excelData = readXlsxFile('./data.xlsx');
// excelData = { 'Sheet1': 'Name,Age\nJohn,30', 'Sheet2': '...' }

const jsonData = csvToJson(excelData.Sheet1);
await writeToSheet(spreadsheetId, jsonData);
```

### Data Conversion
```javascript
const csv = 'Name,Age\nJohn,30\nJane,25';
const json = csvToJson(csv);
const backToCsv = jsonToCsv(json);
```

## 🛠️ Development

```bash
# Install dependencies
npm install

# Run all tests
npm test

# Run only unit tests (fast, no API calls)
npm run test:unit

# Run only integration tests (minimal API usage)
npm run test:integration

# Type checking
npm run typecheck

# Linting
npm run lint
```

**Testing Strategy**: ak-sheets uses a streamlined test approach to minimize API quota usage. See [TESTING.md](./TESTING.md) for details.

## 📄 License

ISC

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📞 Support

- GitHub Issues: [Report bugs or request features](https://github.com/ak/ak-sheets/issues)
- Documentation: [Full API documentation](https://github.com/ak/ak-sheets#readme)

---
