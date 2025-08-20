# Excel Backup Journal Manager

A Node.js application for managing Excel backup journal files. This tool allows you to automatically add new rows to Excel worksheets with incremental dates and random time values.

## Features

- 📊 Process multiple Excel worksheets
- 📅 Add rows with incremental weekly dates
- ⏰ Assign random times from predefined list
- 🎨 Preserve cell formatting and styles
- 💾 Interactive command-line interface

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

1. Make sure you have the Excel file "ЖУРНАЛ резервного копіювання.xlsx" in the project directory
2. Run the application:
```bash
node addRow.js
```
3. Enter the target date in DD.MM.YYYY format when prompted
4. For each worksheet, choose whether to use the last valid row as a base for new rows

## Dependencies

- `exceljs` - Excel file manipulation library
- `readline` - Built-in Node.js module for command-line input

## File Structure

- `addRow.js` - Main application file
- `package.json` - Project dependencies and metadata
- `.gitignore` - Git ignore rules

## How It Works

1. The application reads the Excel file using ExcelJS
2. For each of the first 3 worksheets, it finds the last row with data in columns 2 and 3
3. It asks the user whether to use this row as a base
4. New rows are added weekly until the target date is reached
5. Each new row gets a random time from the predefined list
6. Cell formatting and styles are preserved from the base row

## Available Times

The application randomly selects from these predefined times:
9:35, 10:30, 9:50, 9:35, 10:00, 11:00, 11:20, 10:00, 9:50, 10:00, 11:00, 9:20, 9:50, 9:20, 10:00, 9:30, 10:30, 11:30, 12:30, 9:30, 11:30, 10:20
