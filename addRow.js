const Excel = require("exceljs");
const readline = require("readline");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// Список доступных времён
const availableTimes = [
  "9:35","10:30","9:50","9:35","10:00","11:00","11:20",
  "10:00","9:50","10:00","11:00","9:20","9:50","9:20",
  "10:00","9:30","10:30","11:30","12:30","9:30","11:30","10:20"
];

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

// Найти последнюю строку с текстом в колонках 2 и 3
function findLastValidRow(sheet) {
  for (let i = sheet.rowCount; i > 0; i--) {
    const row = sheet.getRow(i);
    if (row.getCell(2).value && row.getCell(3).value) {
      return row;
    }
  }
  return null;
}

async function processFile(targetDate) {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("ЖУРНАЛ резервного копіювання.xlsx");

  for (let sheetIndex = 0; sheetIndex < 3; sheetIndex++) {
    const sheet = workbook.worksheets[sheetIndex];
    if (!sheet) continue;

    let lastRow = findLastValidRow(sheet);
    if (!lastRow) {
      console.log(`⚠️ Sheet "${sheet.name}" не содержит строк с текстом в колонках 2 и 3, пропускаем`);
      continue;
    }

    console.log(`Sheet "${sheet.name}" last valid row:`);
    console.log(`   Row #${lastRow.number}: Column2="${lastRow.getCell(2).value}", Column3="${lastRow.getCell(3).value}"`);

    const useRow = await new Promise(resolve => {
      rl.question("Использовать эту строку как базу для добавления новых? (y/n): ", ans => {
        resolve(ans.trim().toLowerCase() === "y");
      });
    });

    if (!useRow) {
      console.log(`Пропускаем лист "${sheet.name}"`);
      continue;
    }

    let lastDate = parseDate(lastRow.getCell(2).value);
    let insertIndex = lastRow.number + 1; // следующая строка после найденной

    while (lastDate < targetDate) {
      const newDate = new Date(lastDate.getTime() + 7*24*60*60*1000);

      // создаём массив значений
      const newRowValues = lastRow.values.slice();
      newRowValues[2] = formatDate(newDate);
      newRowValues[3] = availableTimes[Math.floor(Math.random() * availableTimes.length)];

      // вставляем пустую строку на нужное место
      sheet.spliceRows(insertIndex, 0, []);

      const newRow = sheet.getRow(insertIndex);

      // копируем значения и стили
      newRowValues.forEach((val, idx) => {
        const cell = newRow.getCell(idx);
        const lastCell = lastRow.getCell(idx);

        cell.value = val;

        // копируем стили
        cell.font = lastCell.font;
        cell.alignment = lastCell.alignment;
        cell.border = lastCell.border;
        cell.fill = lastCell.fill;
        cell.numFmt = lastCell.numFmt;
      });

      console.log(`   ➕ Added styled row at #${insertIndex}: ${newRowValues}`);

      lastRow = newRow;
      lastDate = parseDate(newRow.getCell(2).value);
      insertIndex++;
    }
  }

  await workbook.xlsx.writeFile("ЖУРНАЛ резервного копіювання.xlsx");
  console.log("✅ File updated successfully!");
  rl.close();
}

rl.question("Enter target date (DD.MM.YYYY): ", async (answer) => {
  const targetDate = parseDate(answer.trim());
  if (!targetDate) {
    console.error("❌ Invalid date format. Use DD.MM.YYYY");
    rl.close();
    return;
  }
  await processFile(targetDate);
});
