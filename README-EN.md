# 🔄 Automatic Folder Renamer - Google Drive

## 📋 Description

Automated script for renaming Google Drive folders based on Google Sheets data. The system performs automatic folder name updates as data is updated in the spreadsheet, ideal for organizations that need to keep folder structures synchronized with control systems.

## ✨ Features

- **Automatic Renaming**: Updates folder names based on spreadsheet data
- **Time Control**: Runs only during business hours (8 AM to 6 PM)
- **Smart Search**: Searches folders in configurable specific locations
- **State Management**: Handles folders with pending status and complete data
- **Scheduled Execution**: Runs automatically 3 times per day (8 AM, 12 PM, 4 PM)
- **Detailed Logging**: Complete reports of all operations

## 🏗️ Project Structure

```
📦 renomeador-pastas-drive/
├── 📄 script.gs              # Main Google Apps Script file
├── 📄 README.md              # Portuguese documentation  
├── 📄 README-EN.md           # English documentation
└── 📄 LICENSE                # Project license
```

## ⚙️ Configuration

### 1. Initial Settings

Edit the constants at the beginning of the script:

```javascript
// Google Sheets ID
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

// Sheet tab name
const SHEET_NAME = 'SHEET_NAME';

// Google Drive folder IDs
const PASTA_PRINCIPAL_ID = 'MAIN_FOLDER_ID';
const PASTA_SECUNDARIA_ID = 'SECONDARY_FOLDER_ID';
```

### 2. Spreadsheet Structure

| Column A | Column B |
|----------|----------|
| MAIN CODE | SECONDARY CODE |
| ABC123XYZ | DOC001/25 |
| X | DOC002/25 |

- **Column A**: Main code (or "X" for pending processes)
- **Column B**: Secondary code (customizable format)

### 3. Installation

1. Access [Google Apps Script](https://script.google.com)
2. Create a new project
3. Paste the code from `script.gs` file
4. Configure necessary permissions
5. Run `primeiraConfiguracao()` to validate setup

## 🚀 How to Use

### Manual Execution

```javascript
// Configuration test
primeiraConfiguracao();

// Single execution
renomearPastas();

// Test with specific cases
testeRapido();
```

### Automatic Execution

```javascript
// Configure automation (run once)
configurarExecucaoAutomatica();

// Check active triggers
listarTriggers();
```

## 📊 How It Works

### Renaming Flow

1. **Spreadsheet Reading**: Loads data from columns A and B
2. **Data Filtering**: Ignores empty rows and processes with "X"
3. **Folder Search**: Searches folders in configured locations
4. **Validation**: Checks if renaming is necessary
5. **Renaming**: Updates name to standard format
6. **Logging**: Records operation result

### Name Patterns

- **Before**: `DOC001_25_X` or `DOC001-25`
- **After**: `DOC001-25_ABC123XYZ`

### Search Locations

1. **Main Folder**: Main folder and configured subfolders
2. **Secondary Folder**: Department/sector specific folder
3. **General Search**: Fallback for entire Drive

## ⏰ Execution Schedule

- **08:00 AM**: Start of business day
- **12:00 PM**: Midday
- **04:00 PM**: End of afternoon

> **Note**: Script only executes between 8 AM and 6 PM, even with configured triggers

## 📈 Reports

### Example Output

```
=== STARTING RENAMING ===
⏰ 06/17/2025 08:00:00 AM (8h - Valid time)
📊 150 rows found

--- Row 5: ABC123XYZ | DOC001/25 ---
🔍 Folder found: DOC001_25_X
🔄 UPDATED: DOC001_25_X → DOC001-25_ABC123XYZ
✅ Renamed

=== SUMMARY ===
✅ Processed: 45 | Renamed: 12
⏸️ Pending (X): 98 | ❌ Errors: 0
```

## 🔒 Required Permissions

- **Google Sheets**: Read data spreadsheet
- **Google Drive**: Search and rename folders
- **Google Apps Script**: Execute time-based triggers

## 🐛 Troubleshooting

### Common Issues

**Permission Error**
```
❌ Error: Exception: You do not have permission to call DriveApp.getFolderById
```
- Solution: Run `primeiraConfiguracao()` and authorize permissions

**Spreadsheet Not Found**
```
❌ Sheet "SHEET_NAME" not found
```
- Solution: Check `SPREADSHEET_ID` and `SHEET_NAME`

**Folder Not Found**
```
❌ Folder not found: DOC001_25, DOC001-25, DOC001/25
```
- Solution: Check if folder exists in configured locations

### Debug

```javascript
// List active triggers
listarTriggers();

// Test with specific data
testeRapido();

// Check configuration
primeiraConfiguracao();
```

## 📝 Changelog

### v1.0.0 (Current)
- ✅ Automatic renaming based on spreadsheet
- ✅ Business hours control
- ✅ Multi-location search
- ✅ X state handling
- ✅ Scheduled execution
- ✅ Detailed logging system

## 🤝 Contributing

1. Fork the project
2. Create a branch for your feature (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is under the MIT license. See the `LICENSE` file for more details.

## 👥 Authors

- **Herberth Goldan Cardoso Junior** - *Initial development* - [HerberthGoldanJr](https://github.com/HerberthGoldanJr)

## 📞 Support

- 📧 Email: herberthgoldanjr@gmail.com

---

⭐ **Like the project? Leave a star!**
