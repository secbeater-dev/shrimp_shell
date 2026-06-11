const assert = require("assert");
const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
globalThis.XLSX = require(path.join(root, "vendor", "xlsx.full.min.js"));
require(path.join(root, "app.js"));

const api = globalThis.ShellCaseAnalyzer;

function totalAmount(summaries) {
  return Math.round(summaries.reduce((sum, summary) => sum + summary.totalAmount, 0));
}

function orderCount(summaries) {
  return summaries.reduce((sum, summary) => sum + summary.orderCount, 0);
}

function parseWorkbookFile(fileName) {
  const filePath = path.join(root, fileName);
  if (!fs.existsSync(filePath)) return null;
  const buffer = fs.readFileSync(filePath);
  const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
  const parsed = api.parseWorkbook(arrayBuffer);
  const aggregation = api.aggregateRecords(parsed.records, {});
  const summaries = api.applyDisplayFilters(aggregation.summaries, {
    sortKey: "totalAmount",
    sortDirection: "desc",
  });
  return {
    parsed,
    summaries,
    detailRows: api.buildDetailRows(summaries),
  };
}

function assertWorkbook(fileName, expected) {
  const result = parseWorkbookFile(fileName);
  if (!result) {
    console.log(`skip ${fileName}: workbook not found`);
    return;
  }

  assert.strictEqual(result.summaries.length, expected.buyers, `${fileName} buyer count`);
  assert.strictEqual(orderCount(result.summaries), expected.orders, `${fileName} order count`);
  assert.strictEqual(totalAmount(result.summaries), expected.amount, `${fileName} total amount`);
  assert.strictEqual(result.detailRows.length, expected.detailRows, `${fileName} detail row count`);
  assert.strictEqual(result.parsed.warnings.length, expected.warnings, `${fileName} warning count`);

  if (expected.warningText) {
    assert.ok(
      result.parsed.warnings.some((warning) => warning.includes(expected.warningText)),
      `${fileName} warning should include ${expected.warningText}`
    );
  }

  if (expected.buyerSource) {
    const sourceRows = api.getBuyerSourceRows(result.parsed.records, expected.buyerSource.buyer);
    const nonCompletedRows = api.getNonCompletedRows(sourceRows);
    const summary = result.summaries.find((item) => item.buyer === expected.buyerSource.buyer);
    assert.strictEqual(sourceRows.length, expected.buyerSource.sourceRows, `${fileName} source row count`);
    assert.strictEqual(nonCompletedRows.length, expected.buyerSource.nonCompletedRows, `${fileName} non-completed row count`);
    assert.deepStrictEqual(
      nonCompletedRows.map((row) => row.rowNumber),
      expected.buyerSource.nonCompletedRowNumbers,
      `${fileName} non-completed row numbers`
    );
    assert.ok(summary, `${fileName} buyer summary should exist`);
    assert.strictEqual(summary.orderCount, expected.buyerSource.completedOrders, `${fileName} buyer completed order count`);
    assert.strictEqual(Math.round(summary.totalAmount), expected.buyerSource.amount, `${fileName} buyer amount`);
  }

  console.log(
    `${fileName}: ${result.summaries.length} buyers, ${orderCount(result.summaries)} orders, ${totalAmount(
      result.summaries
    )} amount`
  );
}

function testSyntheticImportRules() {
  const rows = [
    [
      "買家帳號",
      "賣家帳號",
      "訂單編號",
      "訂單狀態",
      "訂單成立時間",
      "實際撥款額",
      "收件人姓名",
      "收件人電話",
      "商品資訊",
      "取件地址",
      "付款方式",
      "寄送方式",
    ],
    [
      "buyer_neg",
      "seller_a",
      "A001",
      "COMPLETED",
      "2026-01-01 10:00:00",
      "-45",
      "王小明",
      "886900000001",
      "退貨調整商品",
      "測試取件地址一",
      "信用卡付款",
      "新竹物流",
    ],
    [
      "buyer_pos",
      "seller_b",
      "B001",
      "COMPLETED",
      "2026-01-02 10:00:00",
      "100",
      "陳小美",
      "886900000002",
      "正額商品",
      "測試取件地址二",
      "貨到付款",
      "蝦皮店到店",
    ],
    [
      "buyer_pos",
      "seller_b",
      "B002",
      "CANCELLED",
      "2026-01-03 10:00:00",
      "查無",
      "陳小美",
      "886900000002",
      "取消商品",
      "測試取件地址二",
      "貨到付款",
      "蝦皮店到店",
    ],
    [
      "buyer_pos",
      "seller_c",
      "B003",
      "CANCELLED",
      "2026-01-04 10:00:00",
      "查無",
      "陳小美",
      "886900000002",
      "取消商品二",
      "測試取件地址二",
      "貨到付款",
      "蝦皮店到店",
    ],
  ];

  const parsed = api.parseRows(rows, "synthetic");
  const aggregation = api.aggregateRecords(parsed.records, {});
  const unfiltered = api.applyDisplayFilters(aggregation.summaries, {});
  const nonNegative = api.applyDisplayFilters(aggregation.summaries, {
    columnFilters: { totalAmount: ">=0" },
  });
  const explicitMinimum = api.applyDisplayFilters(aggregation.summaries, { amountMin: "0" });
  const blankMinimum = api.applyDisplayFilters(aggregation.summaries, { amountMin: "", countMin: "" });

  assert.deepStrictEqual(parsed.warnings, []);
  assert.strictEqual(unfiltered.length, 2, "blank filters should keep negative total buyers");
  assert.strictEqual(blankMinimum.length, 2, "blank amount/count filters should not constrain results");
  assert.strictEqual(nonNegative.length, 1, "column filter >=0 should exclude negative total buyers");
  assert.strictEqual(explicitMinimum.length, 1, "explicit amountMin 0 should exclude negative total buyers");

  const buyerSourceRows = api.getBuyerSourceRows(parsed.records, "buyer_pos", new Set([2026]));
  const nonCompletedRows = api.getNonCompletedRows(buyerSourceRows);
  assert.strictEqual(buyerSourceRows.length, 3, "source helper should keep all rows for the buyer");
  assert.strictEqual(nonCompletedRows.length, 2, "source helper should expose non-completed rows");
  assert.deepStrictEqual(
    nonCompletedRows.map((row) => row.orderIdDisplay),
    ["B002", "B003"],
    "non-completed rows should remain ordered by date"
  );

  const negativeSummary = unfiltered.find((summary) => summary.buyer === "buyer_neg");
  const negativeRecord = parsed.records.find((record) => record.buyer === "buyer_neg");
  assert.ok(negativeSummary, "negative buyer should be present by default");
  assert.ok(negativeRecord, "negative buyer source record should be present");
  assert.strictEqual(negativeSummary.totalAmount, -45);
  assert.strictEqual(negativeSummary.productsText, "退貨調整商品 x 1");
  assert.strictEqual(negativeSummary.addressText, "測試取件地址一");
  assert.strictEqual(negativeRecord.seller, "seller_a", "賣家帳號 alias should map to 賣家蝦皮帳號");

  const detailRows = api.buildDetailRows([negativeSummary]);
  assert.strictEqual(detailRows.length, 1);
  assert.strictEqual(detailRows[0][5], 1, "missing 購買數量 column should default to 1");
  assert.strictEqual(detailRows[0][6], "退貨調整商品", "商品資訊 alias should map to 商品名稱");
  assert.strictEqual(detailRows[0][9], "測試取件地址一", "取件地址 alias should map to 收件地址");
}

testSyntheticImportRules();
assertWorkbook("Sample.xlsx", {
  buyers: 1,
  orders: 5,
  amount: 672,
  detailRows: 5,
  warnings: 0,
});
assertWorkbook("114-115-新北.xlsx", {
  buyers: 1753,
  orders: 2359,
  amount: 1182412,
  detailRows: 2648,
  warnings: 0,
});
assertWorkbook("SHE1.xlsx", {
  buyers: 39017,
  orders: 44994,
  amount: 18005040,
  detailRows: 44994,
  warnings: 1,
  warningText: "第 14549 列",
  buyerSource: {
    buyer: "m0908970185",
    sourceRows: 4,
    nonCompletedRows: 3,
    nonCompletedRowNumbers: [2157, 2159, 2161],
    completedOrders: 1,
    amount: 141,
  },
});

console.log("import regression tests passed");
