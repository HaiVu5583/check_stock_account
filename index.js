const Excel = require("exceljs");

function isNumeric(value) {
  if (typeof value == "number") {
    return true;
  }
  if (typeof value != "string") {
    return false;
  }
  return !isNaN(parseInt(value));
}

function formatMoney(amount) {
  const formatter = new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
  });
  return formatter.format(amount);
}

async function getTotalCashInOutReport() {
  const FILE_PATH = "/Users/chaunv/Desktop/vnds_saoke.xlsx";
  const FIRST_WORKSHEET_NAME = "vnds";
  const STT_LABEL = "STT";
  const PARENT_CATEGORY_C0LUMN_INDEX = 7;
  const IN_AMOUNT_COLUMN_INDEX = 4;
  const OUT_AMOUNT_COLUMN_INDEX = 5;
  const LOAN_LABEL = "Cho vay";
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(FILE_PATH);
  const workSheet = workbook.getWorksheet(FIRST_WORKSHEET_NAME);
  const rowCount = workSheet.rowCount;
  let startRow = 0;
  const row = workSheet.getRow(11);
  const firstCell = row.getCell(1).value;
  // Find header row, start of table value
  for (let i = 1; i < rowCount; i++) {
    const row = workSheet.getRow(i);
    const firstCell = row.getCell(1);
    if (!row) {
      continue;
    }
    if (firstCell.value == STT_LABEL) {
      startRow = i + 1;
    }
  }
  if (startRow == 0) {
    return;
  }
  let totalCashIn = 0;
  let totalCashOut = 0;
  //
  for (i = startRow; i <= rowCount; i++) {
    const row = workSheet.getRow(i);
    if (!row) {
      continue;
    }
    const sttCell = row.getCell(1);
    if (!sttCell || !isNumeric(sttCell.value)) {
      continue;
    }
    const parentCategoryCell = row.getCell(PARENT_CATEGORY_C0LUMN_INDEX);
    // if parentCategoryCell, it's cash in or cash out record
    if (
      !parentCategoryCell ||
      !parentCategoryCell.value ||
      parentCategoryCell.value == LOAN_LABEL
    ) {
      const inAmountCell = row.getCell(IN_AMOUNT_COLUMN_INDEX);
      const outAmountCell = row.getCell(OUT_AMOUNT_COLUMN_INDEX);
      if (
        (!inAmountCell || !inAmountCell.value) &&
        (!outAmountCell || !outAmountCell.value)
      ) {
        continue;
      }
      if (inAmountCell && typeof inAmountCell.value == "number") {
        totalCashIn += inAmountCell.value;
      }
      if (outAmountCell && typeof outAmountCell.value == "number") {
        totalCashOut += outAmountCell.value;
      }
    }
  }
  console.log("Tổng số tiền nộp: ", formatMoney(totalCashIn));
  console.log("Tổng số tiền rút: ", formatMoney(totalCashOut));
  console.log(
    "Tổng số thực đã nộp vào tài khoản: ",
    formatMoney(totalCashIn - totalCashOut)
  );
}
getTotalCashInOutReport();
