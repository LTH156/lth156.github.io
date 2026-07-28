import fs from "node:fs/promises";

const source = JSON.parse(await fs.readFile("inspection.json", "utf8"));
const paymentBook = source.find((item) => item.file.endsWith("Phiếu chi 2023 - ĐNTT.xlsx"));
const paymentSheet = paymentBook.sheets.find((sheet) => sheet.name === "THU CHI TIỀN MẶT");
const rows = paymentSheet.values.slice(1, -1).filter((row) => Number.isFinite(row[4]));

function excelDate(serial) {
  const utcMs = Math.round((serial - 25569) * 86400 * 1000);
  return new Date(utcMs).toISOString().slice(0, 10);
}

function aggregate(index) {
  const totals = new Map();
  for (const row of rows) {
    const key = String(row[index] ?? "(trống)");
    const current = totals.get(key) || { count: 0, amount: 0 };
    current.count += 1;
    current.amount += row[4];
    totals.set(key, current);
  }
  return [...totals.entries()]
    .map(([key, value]) => ({ key, ...value }))
    .sort((a, b) => b.amount - a.amount);
}

const monthly = new Map();
for (const row of rows) {
  const month = excelDate(row[1]).slice(0, 7);
  const current = monthly.get(month) || { count: 0, amount: 0 };
  current.count += 1;
  current.amount += row[4];
  monthly.set(month, current);
}

const result = {
  count: rows.length,
  total: rows.reduce((sum, row) => sum + row[4], 0),
  minDate: excelDate(Math.min(...rows.map((row) => row[1]))),
  maxDate: excelDate(Math.max(...rows.map((row) => row[1]))),
  byReason: aggregate(6),
  byDocumentType: aggregate(7),
  byCombinedFlag: aggregate(8),
  topPayments: [...rows]
    .sort((a, b) => b[4] - a[4])
    .slice(0, 10)
    .map((row) => ({
      date: excelDate(row[1]),
      voucher: row[2],
      description: row[3],
      amount: row[4],
      counterparty: row[5],
    })),
  monthly: [...monthly.entries()]
    .map(([month, value]) => ({ month, ...value }))
    .sort((a, b) => a.month.localeCompare(b.month)),
  mealPayments: rows
    .filter((row) => String(row[3]).toLowerCase().includes("tiền ăn"))
    .map((row) => ({
      date: excelDate(row[1]),
      voucher: row[2],
      description: row[3],
      amount: row[4],
    })),
};

console.log(JSON.stringify(result, null, 2));
