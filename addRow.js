const Excel = require("exceljs");
const readline = require("readline");

// --- New feature: Smart file selection ---
const fs = require('fs');
const path = require('path');

async function chooseExcelFile() {
  // Find all .xlsx files in the current directory
  const files = fs.readdirSync(process.cwd())
    .filter(f => f.endsWith('.xlsx'));

  if (files.length === 0) {
    // No files found, ask for full path
    return await new Promise(resolve => {
      rl.question("No .xlsx files found in this folder. Enter full path to Excel file: ", answer => {
        resolve(answer.trim());
      });
    });
  }

  // Show options
  console.log("\nHow do you want to select the Excel file?");
  console.log("1. Enter full path manually");
  console.log("2. Choose from files in this folder:");
  files.forEach((f, i) => {
    console.log(`   ${i + 1}. ${f}`);
  });

  const choice = await new Promise(resolve => {
    rl.question("Enter 1 to write full path, or 2 to choose from list: ", answer => {
      resolve(answer.trim());
    });
  });

  if (choice === '2') {
    const fileNum = await new Promise(resolve => {
      rl.question(`Enter file number (1-${files.length}): `, answer => {
        resolve(parseInt(answer.trim()));
      });
    });
    if (!isNaN(fileNum) && fileNum >= 1 && fileNum <= files.length) {
      return files[fileNum - 1];
    } else {
      console.log('❌ Invalid selection.');
      return process.exit(1);
    }
  } else {
    // Default to manual entry
    return await new Promise(resolve => {
      rl.question("Enter full path to Excel file: ", answer => {
        resolve(answer.trim());
      });
    });
  }
}
// --- End new feature ---

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// --- New feature: Ask user for time range and generate availableTimes dynamically ---
async function getAvailableTimes() {
  function parseTime(str) {
    const [h, m] = str.split(":").map(Number);
    if (isNaN(h) || isNaN(m) || h < 0 || h > 23 || m < 0 || m > 59) return null;
    return h * 60 + m;
  }
  function formatTime(minutes) {
    const h = String(Math.floor(minutes / 60)).padStart(2, "0");
    const m = String(minutes % 60).padStart(2, "0");
    return `${h}:${m.padStart(2, "0")}`;
  }
  const from = await new Promise(resolve => {
    rl.question("Enter start time (HH:MM): ", answer => resolve(answer.trim()));
  });
  const till = await new Promise(resolve => {
    rl.question("Enter end time (HH:MM): ", answer => resolve(answer.trim()));
  });
  const fromMin = parseTime(from);
  const tillMin = parseTime(till);
  if (fromMin === null || tillMin === null || fromMin >= tillMin) {
    console.error("❌ Invalid time range. Please use HH:MM format and ensure start < end.");
    rl.close();
    process.exit(1);
  }
  const times = [];
  for (let t = fromMin; t <= tillMin; t++) {
    const m = t % 60;
    if (m % 5 === 0) {
      const timeStr = formatTime(t);
      if (timeStr.endsWith("0") || timeStr.endsWith("5")) times.push(timeStr);
    }
  }
  return times;
}
// --- End new feature ---

// Remove hardcoded availableTimes, will be set in main()
let availableTimes = [];

// Безопасный парсер даты для ExcelJS
function parseDate(value) {
  if (!value) return null;

  if (value instanceof Date) return value;

  if (typeof value === "object") {
    if (value.text) value = value.text;
    else if (value.richText && value.richText.length > 0) {
      value = value.richText.map(t => t.text).join("");
    } else return null;
  }

  if (typeof value === "number") {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }

  if (typeof value === "string") {
    const parts = value.split(".");
    if (parts.length !== 3) return null;
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }

  return null;
}

// Форматирование даты
function formatDate(date) {
  const dd = String(date.getDate()).padStart(2,"0");
  const mm = String(date.getMonth()+1).padStart(2,"0");
  const yyyy = date.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

// Найти последние строки с текстом в колонках 2 и 3
function findLastValidRows(sheet, count = 2) {
  const validRows = [];
  for (let i = sheet.rowCount; i > 0 && validRows.length < count; i--) {
    const row = sheet.getRow(i);
    if (row.getCell(2).value && row.getCell(3).value) {
      validRows.push(row);
    }
  }
  return validRows.reverse(); // Возвращаем в хронологическом порядке
}

// Main function to process the Excel file: creates a backup, reads the file, and adds new rows to selected worksheets
async function processFile(targetDate, originalFileName) {
  const fs = require('fs');
  const path = require('path');
  
  // Check if file exists
  if (!fs.existsSync(originalFileName)) {
    console.error(`❌ File "${originalFileName}" not found!`);
    return;
  }
  
  // Generate backup filename by inserting "copy" before the file extension
  const ext = path.extname(originalFileName);
  const baseName = path.basename(originalFileName, ext);
  const dirName = path.dirname(originalFileName);
  const backupFileName = path.join(dirName, `${baseName} copy${ext}`);
  
  // Create a backup copy of the original file
  try {
    fs.copyFileSync(originalFileName, backupFileName);
    console.log(`📋 Created backup: ${backupFileName}`);
  } catch (error) {
    console.error(`❌ Failed to create backup: ${error.message}`);
    return;
  }

  const workbook = new Excel.Workbook();
  try {
    await workbook.xlsx.readFile(originalFileName);
  } catch (error) {
    console.error(`❌ Failed to read Excel file: ${error.message}`);
    return;
  }

  // Show available worksheets
  console.log("\n📊 Available worksheets:");
  workbook.worksheets.forEach((sheet, index) => {
    console.log(`   ${index + 1}. ${sheet.name}`);
  });

  // Ask user which worksheets to process
  const worksheetInput = await new Promise(resolve => {
    rl.question("\nEnter worksheet numbers to process (e.g., '1,2,3' or 'all'): ", answer => {
      resolve(answer.trim());
    });
  });

  let worksheetsToProcess = [];
  if (worksheetInput.toLowerCase() === 'all') {
    worksheetsToProcess = workbook.worksheets.map((_, index) => index);
  } else {
    const selectedNumbers = worksheetInput.split(',').map(num => parseInt(num.trim()) - 1);
    worksheetsToProcess = selectedNumbers.filter(index => 
      index >= 0 && index < workbook.worksheets.length
    );
  }

  if (worksheetsToProcess.length === 0) {
    console.log("❌ No valid worksheets selected");
    return;
  }

  console.log(`\n🎯 Processing ${worksheetsToProcess.length} worksheet(s)...\n`);

  for (const sheetIndex of worksheetsToProcess) {
    const sheet = workbook.worksheets[sheetIndex];
    if (!sheet) continue;

    let lastRows = findLastValidRows(sheet, 2);
    if (lastRows.length === 0) {
      console.log(`⚠️ Sheet "${sheet.name}" не содержит строк с текстом в колонках 2 и 3, пропускаем`);
      continue;
    }

    const lastRow = lastRows[lastRows.length - 1]; // Самая последняя строка
    console.log(`Sheet "${sheet.name}" last ${lastRows.length} valid row(s):`);
    lastRows.forEach((row, index) => {
      console.log(`   Row #${row.number}: Column2="${row.getCell(2).value}", Column3="${row.getCell(3).value}"`);
    });

    const useRow = await new Promise(resolve => {
      rl.question("Использовать эти строки как базу для добавления новых? (y/n): ", ans => {
        resolve(ans.trim().toLowerCase() === "y");
      });
    });

    if (!useRow) {
      console.log(`Пропускаем лист "${sheet.name}"`);
      continue;
    }

    // Remove feature: always insert at the end (default)
    let insertIndex = lastRow.number + 1;

    let lastDate = parseDate(lastRow.getCell(2).value);
    let currentBaseRow = lastRow; // Текущая базовая строка для копирования

    // Определяем максимальное количество колонок для копирования (базовое + 5 дополнительных)
    const maxColumns = Math.max(
      ...lastRows.map(row => row.cellCount)
    ) + 5;

    while (lastDate < targetDate) {
      const newDate = new Date(lastDate.getTime() + 7*24*60*60*1000);

      // создаём массив значений из текущей базовой строки
      const newRowValues = currentBaseRow.values.slice();
      newRowValues[2] = formatDate(newDate);
      newRowValues[3] = availableTimes[Math.floor(Math.random() * availableTimes.length)];

      // вставляем пустую строку на нужное место
      sheet.spliceRows(insertIndex, 0, []);

      const newRow = sheet.getRow(insertIndex);

      // копируем значения и стили с расширенным диапазоном колонок
      for (let colIndex = 1; colIndex <= maxColumns; colIndex++) {
        const cell = newRow.getCell(colIndex);
        const lastCell = currentBaseRow.getCell(colIndex);
        
        // Устанавливаем значение
        if (colIndex < newRowValues.length) {
          cell.value = newRowValues[colIndex];
        }

        // Копируем стили из текущей базовой строки или предпоследней (если доступна)
        let sourceCell = lastCell;
        if (lastRows.length > 1 && !lastCell.font && !lastCell.fill) {
          // Если в текущей строке нет стилей, берем из предпоследней
          const secondLastRow = lastRows[lastRows.length - 2];
          sourceCell = secondLastRow.getCell(colIndex);
        }

        // копируем стили
        if (sourceCell.font) cell.font = sourceCell.font;
        if (sourceCell.alignment) cell.alignment = sourceCell.alignment;
        if (sourceCell.border) cell.border = sourceCell.border;
        if (sourceCell.fill) cell.fill = sourceCell.fill;
        if (sourceCell.numFmt) cell.numFmt = sourceCell.numFmt;
      }

      console.log(`   ➕ Added styled row at #${insertIndex} with ${maxColumns} columns formatted`);

      currentBaseRow = newRow; // Обновляем базовую строку
      lastDate = parseDate(newRow.getCell(2).value);
      insertIndex++;
    }
  }

  await workbook.xlsx.writeFile(originalFileName);
  console.log("✅ File updated successfully!");
  console.log(`💾 Original backup saved as: ${backupFileName}`);
}

// Main execution
async function main() {
  // --- New feature: Use smart file selection ---
  const filename = await chooseExcelFile();
  // --- End new feature ---

  if (!filename) {
    console.error("❌ No filename provided");
    rl.close();
    return;
  }

  // Ask for target date
  const dateInput = await new Promise(resolve => {
    rl.question("Enter target date (DD.MM.YYYY): ", answer => {
      resolve(answer.trim());
    });
  });

  const targetDate = parseDate(dateInput);
  if (!targetDate) {
    console.error("❌ Invalid date format. Use DD.MM.YYYY");
    rl.close();
    return;
  }

  // --- New feature: Ask for time range ---
  let fromTime, toTime;
  while (true) {
    const fromInput = await new Promise(resolve => {
      rl.question("Enter start time (HH:MM, e.g. 10:00): ", answer => resolve(answer.trim()));
    });
    const toInput = await new Promise(resolve => {
      rl.question("Enter end time (HH:MM, e.g. 13:00): ", answer => resolve(answer.trim()));
    });
    fromTime = parseTimeString(fromInput);
    toTime = parseTimeString(toInput);
    if (fromTime !== null && toTime !== null && fromTime < toTime) break;
    console.log("❌ Invalid time range. Please enter valid times in HH:MM format, and make sure start < end.");
  }
  availableTimes = generateTimesInRange(fromTime, toTime);
  if (availableTimes.length === 0) {
    console.error("❌ No valid times in this range ending with 0 or 5. Try a different range.");
    rl.close();
    return;
  }
  console.log(`Using times: ${availableTimes.join(", ")}`);
  // --- End new feature ---

  await processFile(targetDate, filename);
  rl.close();
}

// Start the application
main().catch(error => {
  console.error("❌ Application error:", error.message);
  rl.close();
});
