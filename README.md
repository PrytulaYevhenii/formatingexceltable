# 📊 Excel Auto-Fill Tool

This project is designed to automatically fill out repetitive and boring Excel logs/reports where most of the content stays the same, but you still need to insert dates and times for a given period.

Instead of manually entering values into spreadsheets, the tool helps you generate them quickly and consistently, saving time and reducing mistakes.

> **Why use this?**
>
> - Automate routine Excel data entry for logs, journals, or schedules
> - Ensure consistent formatting and reduce human error
> - Great for IT, education, HR, or anyone who manages recurring Excel reports

A **Node.js** tool for automating the management of Excel backup journals. Easily add new rows to your Excel worksheets with weekly incremental dates and random time values, while preserving all formatting and styles.

---

## ✨ Features

- 📊 Process multiple Excel worksheets interactively
- 📅 Add rows with incremental weekly dates
- ⏰ Assign random times from a customizable list
- 🎨 Preserve cell formatting and styles (copies up to 5 columns to the right of the data)
- 💾 Automatic backup creation before changes
- 🖥️ Command-line interface (CLI) with smart file selection (manual entry or pick from list)
- 📁 Flexible file input (filename or full path)
- 🛡️ Improved error handling and user feedback

---

## 🛠 Prerequisites

- [Node.js](https://nodejs.org/) (v12 or higher)
- npm (comes with Node.js)

---

## ⚡ Installation

```bash
git clone https://github.com/PrytulaYevhenii/formattingexceltable.git
cd excel
npm install
```

---

## 🚦 Usage

```bash
node addRow.js
```

1. **Select the Excel file**: Choose from a list of `.xlsx` files in the folder or enter a full path manually.
2. **Enter the target date** in `DD.MM.YYYY` format.
3. **Select worksheets** to process (by number or `all`).
4. **Specify a time range**: Enter a start and end time (e.g., `10:00` to `13:00`). Only times within this range and ending in 0 or 5 will be used for random assignment.
5. **Confirm** using the last valid rows as a base for new rows.
6. The tool will always insert new rows at the end, copying formatting from the last two valid rows and up to 5 columns to the right.
7. A backup is automatically created before any changes.

---

## 💡 Example Session

```bash
$ node addRow.js
How do you want to select the Excel file?
1. Enter full path manually
2. Choose from files in this folder:
   1. ЖУРНАЛ резервного копіювання.xlsx
Enter 1 to write full path, or 2 to choose from list: 2
Enter file number (1-1): 1
Enter target date (DD.MM.YYYY): 31.12.2025
📋 Created backup: ЖУРНАЛ резервного копіювання copy.xlsx
📊 Available worksheets:
   1. Sheet1
   2. Sheet2
   3. Sheet3
Enter worksheet numbers to process (e.g., '1,2,3' or 'all'): 1,2
Sheet "Sheet1" last 2 valid row(s):
   Row #15: Column2="15.08.2025", Column3="10:30"
   Row #16: Column2="22.08.2025", Column3="11:00"
Использовать эти строки как базу для добавления новых? (y/n): y
   ➕ Added styled row at #17 with 13 columns formatted
✅ File updated successfully!
💾 Original backup saved as: ЖУРНАЛ резервного копіювання copy.xlsx
```

---

## 📦 Dependencies

- [`exceljs`](https://www.npmjs.com/package/exceljs) — Excel file manipulation
- [`readline`](https://nodejs.org/api/readline.html) — Node.js CLI input (built-in)
- [`fs`, `path`](https://nodejs.org/api/fs.html) — File system utilities (built-in)

---

## 📁 Project Structure

```
excel/
├── addRow.js         # Main application file
├── package.json      # Project dependencies and metadata
├── .gitignore        # Git ignore rules
└── README.md         # This documentation
```

---

## 🧠 How It Works

1. **Prompts** for an Excel filename (relative or absolute path, or pick from list)
2. **Creates a backup** of the Excel file (adds `copy` before the extension)
3. **Reads** the Excel file using ExcelJS
4. **Lists worksheets** and lets you select which to process
5. **Finds the last 2 valid rows** (with data in columns 2 and 3) in each worksheet
6. **Asks for confirmation** to use these rows as a base
7. **Adds new rows** weekly until the target date is reached (always at the end)
8. **Assigns random times** from the predefined list
9. **Preserves all formatting and styles** for up to 5 columns to the right of the data
10. **Improved error handling** and user feedback throughout the process

---

## ⏰ Available Times

The application randomly selects from these times:

```
9:35, 10:30, 9:50, 9:35, 10:00, 11:00, 11:20, 10:00, 9:50, 10:00, 11:00, 9:20, 9:50, 9:20, 10:00, 9:30, 10:30, 11:30, 12:30, 9:30, 11:30, 10:20
```

---

## 📝 License

MIT License. See [LICENSE](LICENSE) for details.

---

## 🙋‍♂️ Contributing

Pull requests and suggestions are welcome! For major changes, please open an issue first to discuss what you would like to change.

---

## 📣 Author

[Yevhenii Prytula](https://github.com/PrytulaYevhenii)