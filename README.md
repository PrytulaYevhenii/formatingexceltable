# Excel Backup Journal Manager

A Node.js application for managing Excel backup journal files. This tool allows you to automatically add new rows to Excel worksheets with incremental dates and random time values.

## Features

- 📊 Process multiple Excel worksheets
- 📅 Add rows with incremental weekly dates
- ⏰ Assign random times from predefined list
- 🎨 Preserve cell formatting and styles
- 💾 Interactive command-line interface
- 🔒 Automatic backup creation before making changes
- 📁 Flexible file input (supports any Excel file path)

## Prerequisites

- Node.js (version 12 or higher)
- npm (Node Package Manager)

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd excel
```

2. Install dependencies:
```bash
npm install
```

## Usage

1. Run the application:
```bash
node addRow.js
```
2. Enter the Excel filename when prompted (you can use just the filename if it's in the same directory, or provide the full path)
3. Enter the target date in DD.MM.YYYY format when prompted
4. For each worksheet, choose whether to use the last valid row as a base for new rows

## Example

```bash
$ node addRow.js
Enter Excel filename (e.g., 'file.xlsx' or full path): ЖУРНАЛ резервного копіювання.xlsx
Enter target date (DD.MM.YYYY): 31.12.2025
📋 Created backup: ЖУРНАЛ резервного копіювання copy.xlsx
Sheet "Sheet1" last valid row:
   Row #15: Column2="15.08.2025", Column3="10:30"
Использовать эту строку как базу для добавления новых? (y/n): y
   ➕ Added styled row at #16: ...
✅ File updated successfully!
💾 Original backup saved as: ЖУРНАЛ резервного копіювання copy.xlsx
```

## Dependencies

- `exceljs` - Excel file manipulation library
- `readline` - Built-in Node.js module for command-line input

## File Structure

- `addRow.js` - Main application file
- `package.json` - Project dependencies and metadata
- `.gitignore` - Git ignore rules

## How It Works

1. The application prompts for an Excel filename (can be relative or absolute path)
2. It creates a backup copy of the Excel file (adds " copy" before the file extension)
3. The application reads the Excel file using ExcelJS
4. For each of the first 3 worksheets, it finds the last row with data in columns 2 and 3
5. It asks the user whether to use this row as a base
6. New rows are added weekly until the target date is reached
7. Each new row gets a random time from the predefined list
8. Cell formatting and styles are preserved from the base row

## Available Times

The application randomly selects from these predefined times:
9:35, 10:30, 9:50, 9:35, 10:00, 11:00, 11:20, 10:00, 9:50, 10:00, 11:00, 9:20, 9:50, 9:20, 10:00, 9:30, 10:30, 11:30, 12:30, 9:30, 11:30, 10:20
