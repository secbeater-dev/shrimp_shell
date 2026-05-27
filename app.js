(function () {
  "use strict";

  const REQUIRED_COLUMNS = [
    "訂單編號",
    "訂單狀態",
    "訂單成立時間",
    "實際撥款額",
    "買家蝦皮帳號",
    "收件人姓名",
    "收件人電話",
    "購買數量",
    "商品名稱",
  ];

  const OUTPUT_HEADERS = [
    "買家蝦皮帳號",
    "實際撥款額總額",
    "總購買次數",
    "訂單日期統整",
    "購買商品名稱統整",
    "收件電話統整",
    "收件人姓名統整",
    "收件地址統整",
    "付款方式統整",
    "寄送方式統整",
  ];

  const DETAIL_HEADERS = [
    "買家蝦皮帳號",
    "訂單編號",
    "訂單狀態",
    "訂單成立時間",
    "實際撥款額",
    "購買數量",
    "商品名稱",
    "收件人電話",
    "收件人姓名",
    "收件地址",
    "賣家蝦皮帳號",
    "付款方式",
    "寄送方式",
    "原始列號",
  ];

  const OPTIONAL_COLUMNS = [
    "收件地址",
    "賣家蝦皮帳號",
    "付款方式",
    "寄送方式",
  ];

  const numberFormatter = new Intl.NumberFormat("zh-TW");
  const currencyFormatter = new Intl.NumberFormat("zh-TW", {
    maximumFractionDigits: 0,
  });

  const STORAGE_KEYS = {
    theme: "shellCaseAnalyzerTheme",
    columnOrder: "shellCaseAnalyzerSummaryColumnOrder",
    columnWidths: "shellCaseAnalyzerSummaryColumnWidths",
    welcome: "shellCaseAnalyzerWelcomeSeen_20260527",
  };

  const MIN_SUMMARY_COLUMN_WIDTH = 118;

  const SUMMARY_COLUMNS = [
    {
      key: "buyer",
      label: "買家蝦皮帳號",
      type: "text",
      width: 180,
      display: (summary) => summary.buyer,
      sortValue: (summary) => summary.buyer,
    },
    {
      key: "totalAmount",
      label: "實際撥款額總額",
      type: "number",
      width: 158,
      filterPlaceholder: "篩選：>=1000",
      display: (summary) => currencyFormatter.format(Math.round(summary.totalAmount)),
      sortValue: (summary) => summary.totalAmount,
    },
    {
      key: "orderCount",
      label: "總購買次數",
      type: "number",
      width: 138,
      filterPlaceholder: "篩選：>=10",
      display: (summary) => numberFormatter.format(summary.orderCount),
      sortValue: (summary) => summary.orderCount,
    },
    {
      key: "dateText",
      label: "訂單日期統整",
      type: "date",
      width: 250,
      multiline: true,
      display: (summary) => summary.dateText,
      sortValue: (summary) => summary.latestDateMs || 0,
    },
    {
      key: "productsText",
      label: "購買商品名稱統整",
      type: "text",
      width: 360,
      multiline: true,
      display: (summary) => summary.productsText,
      sortValue: (summary) => summary.productsText,
    },
    {
      key: "phoneText",
      label: "收件電話統整",
      type: "text",
      width: 190,
      multiline: true,
      display: (summary) => summary.phoneText,
      sortValue: (summary) => summary.phoneText,
    },
    {
      key: "nameText",
      label: "收件人姓名統整",
      type: "text",
      width: 190,
      multiline: true,
      display: (summary) => summary.nameText,
      sortValue: (summary) => summary.nameText,
    },
    {
      key: "addressText",
      label: "收件地址統整",
      type: "text",
      width: 320,
      multiline: true,
      display: (summary) => summary.addressText,
      sortValue: (summary) => summary.addressText,
    },
    {
      key: "paymentText",
      label: "付款方式統整",
      type: "text",
      width: 190,
      multiline: true,
      display: (summary) => summary.paymentText,
      sortValue: (summary) => summary.paymentText,
    },
    {
      key: "shippingText",
      label: "寄送方式統整",
      type: "text",
      width: 190,
      multiline: true,
      display: (summary) => summary.shippingText,
      sortValue: (summary) => summary.shippingText,
    },
  ];

  const state = {
    records: [],
    warnings: [],
    fileName: "",
    sheetName: "",
    sourceRowCount: 0,
    currentSummaries: [],
    currentRows: [],
    charts: {},
    activeView: "buyers",
    sidebarCollapsed: false,
    activeClue: null,
    currentCaseSummary: null,
    theme: "dark",
    summaryColumnOrder: SUMMARY_COLUMNS.map((column) => column.key),
    summaryColumnWidths: Object.fromEntries(SUMMARY_COLUMNS.map((column) => [column.key, column.width])),
    summaryColumnFilters: {},
    summarySortKey: "totalAmount",
    summarySortDirection: "desc",
    draggedColumnKey: "",
    activeColumnFilterKey: "",
    activeColumnFilterStart: null,
    activeColumnFilterEnd: null,
  };

  const elements = {};

  function cleanCell(value) {
    if (value === null || value === undefined) return "";
    return String(value).replace(/^\uFEFF/, "").trim();
  }

  function normalizeHeader(value) {
    return cleanCell(value).replace(/\s+/g, "");
  }

  function normalizeSearch(value) {
    return cleanCell(value).toLowerCase();
  }

  function parseNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    const text = cleanCell(value)
      .replace(/,/g, "")
      .replace(/NT\$/gi, "")
      .replace(/\$/g, "");
    if (!text || text === "查無" || text === "-") return null;
    const parsed = Number(text);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function parseQuantity(value) {
    const parsed = parseNumber(value);
    if (parsed === null) return null;
    return Math.max(0, Math.trunc(parsed));
  }

  function excelSerialToDateText(value) {
    if (typeof XLSX === "undefined" || !XLSX.SSF || typeof value !== "number") {
      return cleanCell(value);
    }
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return cleanCell(value);
    const y = String(parsed.y).padStart(4, "0");
    const m = String(parsed.m).padStart(2, "0");
    const d = String(parsed.d).padStart(2, "0");
    const h = String(parsed.H || 0).padStart(2, "0");
    const min = String(parsed.M || 0).padStart(2, "0");
    const s = String(Math.floor(parsed.S || 0)).padStart(2, "0");
    return `${y}-${m}-${d} ${h}:${min}:${s}`;
  }

  function normalizeDateText(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return formatDateTime(value);
    }
    if (typeof value === "number") return excelSerialToDateText(value);
    return cleanCell(value);
  }

  function parseDateMs(value) {
    const text = normalizeDateText(value);
    const match = text.match(
      /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/
    );
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]) - 1;
    const day = Number(match[3]);
    const hour = Number(match[4] || 0);
    const minute = Number(match[5] || 0);
    const second = Number(match[6] || 0);
    const date = new Date(year, month, day, hour, minute, second);
    return Number.isNaN(date.getTime()) ? null : date.getTime();
  }

  function formatDateTime(date) {
    const y = String(date.getFullYear()).padStart(4, "0");
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    const h = String(date.getHours()).padStart(2, "0");
    const min = String(date.getMinutes()).padStart(2, "0");
    const s = String(date.getSeconds()).padStart(2, "0");
    return `${y}-${m}-${d} ${h}:${min}:${s}`;
  }

  function formatDateOnly(ms) {
    if (ms === null || ms === undefined) return "";
    const date = new Date(ms);
    if (Number.isNaN(date.getTime())) return "";
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(
      date.getDate()
    ).padStart(2, "0")}`;
  }

  function dateInputToRange(value, endOfDay) {
    if (!value) return null;
    const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return null;
    const date = new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      endOfDay ? 23 : 0,
      endOfDay ? 59 : 0,
      endOfDay ? 59 : 0,
      endOfDay ? 999 : 0
    );
    return date.getTime();
  }

  function isBlankRow(row) {
    return row.every((value) => cleanCell(value) === "");
  }

  function uniquePush(map, value) {
    const text = cleanCell(value);
    if (text && !map.has(text)) map.set(text, true);
  }

  function parseWorkbook(arrayBuffer) {
    if (typeof XLSX === "undefined") {
      throw new Error("找不到 Excel 解析函式庫，請確認 vendor/xlsx.full.min.js 已載入。");
    }
    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: false,
      raw: false,
      dateNF: "yyyy-mm-dd hh:mm:ss",
    });
    if (!workbook.SheetNames.length) {
      throw new Error("這個 Excel 沒有可讀取的工作表。");
    }
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
      raw: false,
      dateNF: "yyyy-mm-dd hh:mm:ss",
    });
    return parseRows(rows, sheetName);
  }

  function parseRows(rows, sheetName) {
    if (!rows.length) {
      throw new Error("這個工作表沒有資料。");
    }

    const headers = rows[0].map(normalizeHeader);
    const columnIndex = new Map();
    headers.forEach((header, index) => {
      if (header && !columnIndex.has(header)) columnIndex.set(header, index);
    });

    const missing = REQUIRED_COLUMNS.filter((column) => !columnIndex.has(normalizeHeader(column)));
    if (missing.length) {
      throw new Error(`缺少必要欄位：${missing.join("、")}`);
    }

    const allColumns = REQUIRED_COLUMNS.concat(OPTIONAL_COLUMNS);
    const get = (row, column) => {
      const index = columnIndex.get(normalizeHeader(column));
      return index === undefined ? "" : row[index];
    };

    const warnings = [];
    const records = [];
    const seenOrderAmounts = new Map();

    rows.slice(1).forEach((row, rowOffset) => {
      const rowNumber = rowOffset + 2;
      if (!Array.isArray(row) || isBlankRow(row)) return;

      const buyer = cleanCell(get(row, "買家蝦皮帳號"));
      const status = cleanCell(get(row, "訂單狀態"));
      const rawOrderId = cleanCell(get(row, "訂單編號"));
      const orderId = rawOrderId || `缺訂單編號-第${rowNumber}列`;
      const amountRaw = get(row, "實際撥款額");
      const amountParsed = parseNumber(amountRaw);
      const quantityRaw = get(row, "購買數量");
      const quantityParsed = parseQuantity(quantityRaw);
      const dateText = normalizeDateText(get(row, "訂單成立時間"));
      const record = {
        rowNumber,
        orderId,
        orderIdDisplay: rawOrderId,
        status,
        dateText,
        dateMs: parseDateMs(dateText),
        amount: amountParsed === null ? 0 : amountParsed,
        amountWasInvalid: amountParsed === null,
        buyer,
        name: cleanCell(get(row, "收件人姓名")),
        phone: cleanCell(get(row, "收件人電話")),
        address: cleanCell(get(row, "收件地址")),
        quantity: quantityParsed === null ? 0 : quantityParsed,
        quantityWasInvalid: quantityParsed === null,
        product: cleanCell(get(row, "商品名稱")),
        seller: cleanCell(get(row, "賣家蝦皮帳號")),
        payment: cleanCell(get(row, "付款方式")),
        shipping: cleanCell(get(row, "寄送方式")),
      };

      allColumns.forEach((column) => {
        if (!columnIndex.has(normalizeHeader(column)) && OPTIONAL_COLUMNS.includes(column)) {
          record[column] = "";
        }
      });

      if (!buyer) {
        warnings.push(`第 ${rowNumber} 列缺少買家蝦皮帳號，已略過彙整。`);
      }
      if (status === "COMPLETED" && !rawOrderId) {
        warnings.push(`第 ${rowNumber} 列缺少訂單編號，系統以列號作為暫時識別。`);
      }
      if (status === "COMPLETED" && amountParsed === null) {
        warnings.push(`第 ${rowNumber} 列完成訂單的實際撥款額不是數字，已以 0 計算。`);
      }
      if (status === "COMPLETED" && quantityParsed === null) {
        warnings.push(`第 ${rowNumber} 列完成訂單的購買數量不是數字，商品數量已以 0 計算。`);
      }
      if (status === "COMPLETED" && buyer) {
        const amountKey = `${buyer}::${orderId}`;
        if (seenOrderAmounts.has(amountKey)) {
          const previous = seenOrderAmounts.get(amountKey);
          if (Math.abs(previous - record.amount) > 0.001) {
            warnings.push(
              `訂單 ${rawOrderId || orderId} 在多列中有不同實際撥款額，彙整時採第一筆金額。`
            );
          }
        } else {
          seenOrderAmounts.set(amountKey, record.amount);
        }
      }

      records.push(record);
    });

    return {
      sheetName,
      records,
      warnings,
      sourceRowCount: Math.max(0, rows.length - 1),
      headers: rows[0].map(cleanCell),
    };
  }

  function aggregateRecords(records, filters) {
    const settings = filters || {};
    const dateFromMs = dateInputToRange(settings.dateFrom, false);
    const dateToMs = dateInputToRange(settings.dateTo, true);
    const productTerm = normalizeSearch(settings.productKeyword);
    const buyers = new Map();
    const includedRows = [];

    records.forEach((record) => {
      if (record.status !== "COMPLETED" || !record.buyer) return;
      if (dateFromMs !== null && (record.dateMs === null || record.dateMs < dateFromMs)) return;
      if (dateToMs !== null && (record.dateMs === null || record.dateMs > dateToMs)) return;
      if (productTerm && !normalizeSearch(record.product).includes(productTerm)) return;

      includedRows.push(record);
      if (!buyers.has(record.buyer)) {
        buyers.set(record.buyer, {
          buyer: record.buyer,
          orderMap: new Map(),
          productMap: new Map(),
          phoneMap: new Map(),
          nameMap: new Map(),
          addressMap: new Map(),
          paymentMap: new Map(),
          shippingMap: new Map(),
          detailRows: [],
          firstSeenIndex: buyers.size,
        });
      }

      const summary = buyers.get(record.buyer);
      summary.detailRows.push(record);

      if (!summary.orderMap.has(record.orderId)) {
        summary.orderMap.set(record.orderId, {
          orderId: record.orderId,
          orderIdDisplay: record.orderIdDisplay,
          dateText: record.dateText,
          dateMs: record.dateMs,
          amount: record.amount,
        });
      }

      if (record.product && record.quantity > 0) {
        if (!summary.productMap.has(record.product)) {
          summary.productMap.set(record.product, {
            name: record.product,
            quantity: 0,
            firstSeenIndex: summary.productMap.size,
          });
        }
        summary.productMap.get(record.product).quantity += record.quantity;
      }

      uniquePush(summary.phoneMap, record.phone);
      uniquePush(summary.nameMap, record.name);
      uniquePush(summary.addressMap, record.address);
      uniquePush(summary.paymentMap, record.payment);
      uniquePush(summary.shippingMap, record.shipping);
    });

    const summaries = Array.from(buyers.values()).map(finalizeSummary);
    return { summaries, includedRows };
  }

  function finalizeSummary(summary) {
    const orders = Array.from(summary.orderMap.values()).sort(compareOrdersByDate);
    const products = Array.from(summary.productMap.values()).sort((a, b) => {
      if (b.quantity !== a.quantity) return b.quantity - a.quantity;
      return a.firstSeenIndex - b.firstSeenIndex;
    });
    const phones = Array.from(summary.phoneMap.keys());
    const names = Array.from(summary.nameMap.keys());
    const addresses = Array.from(summary.addressMap.keys());
    const payments = Array.from(summary.paymentMap.keys());
    const shippings = Array.from(summary.shippingMap.keys());
    const totalAmount = orders.reduce((sum, order) => sum + order.amount, 0);
    const dateText = orders.map((order) => order.dateText).join("\n");
    const productsText = products.map((item) => `${item.name} x ${item.quantity}`).join("\n");
    const phoneText = phones.join("\n");
    const nameText = names.join("\n");
    const addressText = addresses.join("\n");
    const paymentText = payments.join("\n");
    const shippingText = shippings.join("\n");
    const latestDateMs = orders.reduce(
      (max, order) => (order.dateMs !== null && order.dateMs > max ? order.dateMs : max),
      -Infinity
    );
    const firstDateMs = orders.reduce(
      (min, order) => (order.dateMs !== null && order.dateMs < min ? order.dateMs : min),
      Infinity
    );

    return {
      ...summary,
      orders,
      products,
      phones,
      names,
      addresses,
      payments,
      shippings,
      totalAmount,
      orderCount: orders.length,
      dateText,
      productsText,
      phoneText,
      nameText,
      addressText,
      paymentText,
      shippingText,
      latestDateMs: latestDateMs === -Infinity ? null : latestDateMs,
      firstDateMs: firstDateMs === Infinity ? null : firstDateMs,
      searchBlob: [
        summary.buyer,
        dateText,
        productsText,
        phoneText,
        nameText,
        addressText,
        paymentText,
        shippingText,
      ]
        .join("\n")
        .toLowerCase(),
    };
  }

  function compareOrdersByDate(a, b) {
    const aTime = a.dateMs === null ? Number.MAX_SAFE_INTEGER : a.dateMs;
    const bTime = b.dateMs === null ? Number.MAX_SAFE_INTEGER : b.dateMs;
    if (aTime !== bTime) return aTime - bTime;
    return String(a.orderId).localeCompare(String(b.orderId), "zh-Hant");
  }

  function applyDisplayFilters(summaries, filters) {
    const settings = filters || {};
    const search = normalizeSearch(settings.search);
    const amountMin = parseNumber(settings.amountMin) || 0;
    const countMin = parseQuantity(settings.countMin) || 0;
    const columnFilters = settings.columnFilters || state.summaryColumnFilters;
    const sort = getSortSettings(settings);

    const filtered = summaries.filter((summary) => {
      if (summary.totalAmount < amountMin) return false;
      if (summary.orderCount < countMin) return false;
      if (search && !summary.searchBlob.includes(search)) return false;
      return SUMMARY_COLUMNS.every((column) => matchesColumnFilter(summary, column, columnFilters[column.key]));
    });

    return filtered.sort((a, b) => compareSummaries(a, b, sort.key, sort.direction));
  }

  function getSortSettings(filters) {
    if (filters && filters.sortKey) {
      return {
        key: getSummaryColumn(filters.sortKey) ? filters.sortKey : "totalAmount",
        direction: filters.sortDirection === "asc" ? "asc" : "desc",
      };
    }

    switch (filters && filters.sortBy) {
      case "countDesc":
        return { key: "orderCount", direction: "desc" };
      case "latestDesc":
        return { key: "dateText", direction: "desc" };
      case "buyerAsc":
        return { key: "buyer", direction: "asc" };
      case "amountDesc":
        return { key: "totalAmount", direction: "desc" };
      default:
        return { key: state.summarySortKey, direction: state.summarySortDirection };
    }
  }

  function matchesColumnFilter(summary, column, rawFilter) {
    const filterText = cleanCell(rawFilter);
    if (!filterText) return true;

    const displayValue = cleanCell(column.display(summary));
    if (column.type === "number") {
      return matchesNumericFilter(Number(column.sortValue(summary)) || 0, filterText, displayValue);
    }

    return normalizeSearch(displayValue).includes(normalizeSearch(filterText));
  }

  function matchesNumericFilter(value, filterText, displayValue) {
    const normalized = filterText.replace(/,/g, "").replace(/NT\$/gi, "").replace(/\$/g, "").trim();
    const match = normalized.match(/^(>=|<=|>|<|=)?\s*(-?\d+(?:\.\d+)?)$/);
    if (!match) return normalizeSearch(displayValue).includes(normalizeSearch(filterText));

    const operator = match[1] || ">=";
    const target = Number(match[2]);
    if (!Number.isFinite(target)) return true;

    switch (operator) {
      case ">":
        return value > target;
      case ">=":
        return value >= target;
      case "<":
        return value < target;
      case "<=":
        return value <= target;
      case "=":
        return value === target;
      default:
        return value >= target;
    }
  }

  function compareSummaries(a, b, sortKey, direction) {
    const column = getSummaryColumn(sortKey) || getSummaryColumn("totalAmount");
    const multiplier = direction === "asc" ? 1 : -1;
    const aValue = column.sortValue(a);
    const bValue = column.sortValue(b);
    let result = 0;

    if (column.type === "number" || column.type === "date") {
      result = (Number(aValue) || 0) - (Number(bValue) || 0);
    } else {
      result = cleanCell(aValue).localeCompare(cleanCell(bValue), "zh-Hant", { numeric: true });
    }

    if (result !== 0) return result * multiplier;
    return a.buyer.localeCompare(b.buyer, "zh-Hant", { numeric: true });
  }

  function buildSummaryRows(summaries) {
    return summaries.map((summary) => [
      summary.buyer,
      Math.round(summary.totalAmount),
      summary.orderCount,
      summary.dateText,
      summary.productsText,
      summary.phoneText,
      summary.nameText,
      summary.addressText,
      summary.paymentText,
      summary.shippingText,
    ]);
  }

  function buildDetailRows(summaries) {
    return summaries.flatMap((summary) =>
      summary.detailRows.map((row) => [
        row.buyer,
        row.orderIdDisplay || row.orderId,
        row.status,
        row.dateText,
        Math.round(row.amount),
        row.quantity,
        row.product,
        row.phone,
        row.name,
        row.address,
        row.seller,
        row.payment,
        row.shipping,
        row.rowNumber,
      ])
    );
  }

  function joinUniqueValues(items, key) {
    const values = [];
    const seen = new Set();
    items.forEach((item) => {
      const text = cleanCell(item[key]);
      if (text && !seen.has(text)) {
        seen.add(text);
        values.push(text);
      }
    });
    return values.join("\n");
  }

  function getSummaryColumn(key) {
    return SUMMARY_COLUMNS.find((column) => column.key === key);
  }

  function getOrderedSummaryColumns() {
    const ordered = state.summaryColumnOrder
      .map(getSummaryColumn)
      .filter(Boolean);
    const missing = SUMMARY_COLUMNS.filter((column) => !ordered.includes(column));
    return [...ordered, ...missing];
  }

  function getColumnWidth(column) {
    if (!column) return MIN_SUMMARY_COLUMN_WIDTH;
    const width = Number(state.summaryColumnWidths[column.key]);
    if (!Number.isFinite(width)) return column.width;
    return Math.max(MIN_SUMMARY_COLUMN_WIDTH, width);
  }

  function loadTablePreferences() {
    try {
      const storedOrder = JSON.parse(localStorage.getItem(STORAGE_KEYS.columnOrder) || "[]");
      if (Array.isArray(storedOrder) && storedOrder.some((key) => getSummaryColumn(key))) {
        state.summaryColumnOrder = [
          ...storedOrder.filter((key) => getSummaryColumn(key)),
          ...SUMMARY_COLUMNS.map((column) => column.key).filter((key) => !storedOrder.includes(key)),
        ];
      }
    } catch (error) {
      state.summaryColumnOrder = SUMMARY_COLUMNS.map((column) => column.key);
    }

    try {
      const storedWidths = JSON.parse(localStorage.getItem(STORAGE_KEYS.columnWidths) || "{}");
      if (storedWidths && typeof storedWidths === "object") {
        SUMMARY_COLUMNS.forEach((column) => {
          const width = Number(storedWidths[column.key]);
          if (Number.isFinite(width)) {
            state.summaryColumnWidths[column.key] = Math.max(MIN_SUMMARY_COLUMN_WIDTH, width);
          }
        });
      }
    } catch (error) {
      state.summaryColumnWidths = Object.fromEntries(SUMMARY_COLUMNS.map((column) => [column.key, column.width]));
    }
  }

  function saveTablePreferences() {
    try {
      localStorage.setItem(STORAGE_KEYS.columnOrder, JSON.stringify(state.summaryColumnOrder));
      localStorage.setItem(STORAGE_KEYS.columnWidths, JSON.stringify(state.summaryColumnWidths));
    } catch (error) {
      // Preferences are nice-to-have; the table still works when storage is unavailable.
    }
  }

  function getStoredTheme() {
    try {
      return localStorage.getItem(STORAGE_KEYS.theme) === "light" ? "light" : "dark";
    } catch (error) {
      return "dark";
    }
  }

  function init() {
    state.theme = getStoredTheme();
    loadTablePreferences();

    Object.assign(elements, {
      appShell: document.getElementById("appShell"),
      sidebar: document.getElementById("sidebar"),
      sidebarCollapseBtn: document.getElementById("sidebarCollapseBtn"),
      themeToggleBtn: document.getElementById("themeToggleBtn"),
      themeToggleIcon: document.getElementById("themeToggleIcon"),
      themeToggleLabel: document.getElementById("themeToggleLabel"),
      navItems: Array.from(document.querySelectorAll("[data-view-target]")),
      viewPanels: Array.from(document.querySelectorAll("[data-view]")),
      privacyModal: document.getElementById("privacyModal"),
      privacyConfirmBtn: document.getElementById("privacyConfirmBtn"),
      fileInput: document.getElementById("fileInput"),
      selectFileBtn: document.getElementById("selectFileBtn"),
      dropZone: document.getElementById("dropZone"),
      fileMeta: document.getElementById("fileMeta"),
      exportXlsxBtn: document.getElementById("exportXlsxBtn"),
      exportCsvBtn: document.getElementById("exportCsvBtn"),
      exportDetailsBtn: document.getElementById("exportDetailsBtn"),
      noticePanel: document.getElementById("noticePanel"),
      warningPanel: document.getElementById("warningPanel"),
      warningCount: document.getElementById("warningCount"),
      warningList: document.getElementById("warningList"),
      toggleWarningsBtn: document.getElementById("toggleWarningsBtn"),
      resetClueBtn: document.getElementById("resetClueBtn"),
      kpiBuyers: document.getElementById("kpiBuyers"),
      kpiOrders: document.getElementById("kpiOrders"),
      kpiAmount: document.getElementById("kpiAmount"),
      kpiDateRange: document.getElementById("kpiDateRange"),
      kpiProducts: document.getElementById("kpiProducts"),
      resultCount: document.getElementById("resultCount"),
      summaryTable: document.getElementById("summaryTable"),
      summaryColgroup: document.getElementById("summaryColgroup"),
      summaryHeadRow: document.getElementById("summaryHeadRow"),
      summaryBody: document.getElementById("summaryBody"),
      clueList: document.getElementById("clueList"),
      detailDrawer: document.getElementById("detailDrawer"),
      drawerBuyer: document.getElementById("drawerBuyer"),
      drawerContent: document.getElementById("drawerContent"),
      exportCaseXlsxBtn: document.getElementById("exportCaseXlsxBtn"),
      exportCaseCsvBtn: document.getElementById("exportCaseCsvBtn"),
      closeDrawerBtn: document.getElementById("closeDrawerBtn"),
    });

    applyTheme(state.theme, false);
    setupEvents();
    setupChartsDefaults();
    setActiveView(state.activeView);
    setSidebarCollapsed(state.sidebarCollapsed);
    maybeShowPrivacyModal();
    document.documentElement.dataset.appReady = "true";
    render();
  }

  function setupEvents() {
    elements.selectFileBtn.addEventListener("click", () => elements.fileInput.click());
    elements.dropZone.addEventListener("click", (event) => {
      if (event.target === elements.dropZone) elements.fileInput.click();
    });
    elements.dropZone.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        elements.fileInput.click();
      }
    });
    elements.fileInput.addEventListener("change", (event) => {
      const file = event.target.files && event.target.files[0];
      if (file) handleFile(file);
    });
    ["dragenter", "dragover"].forEach((type) => {
      elements.dropZone.addEventListener(type, (event) => {
        event.preventDefault();
        elements.dropZone.classList.add("dragging");
      });
    });
    ["dragleave", "drop"].forEach((type) => {
      elements.dropZone.addEventListener(type, (event) => {
        event.preventDefault();
        elements.dropZone.classList.remove("dragging");
      });
    });
    elements.dropZone.addEventListener("drop", (event) => {
      const file = event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) handleFile(file);
    });

    elements.themeToggleBtn.addEventListener("click", toggleTheme);
    elements.exportXlsxBtn.addEventListener("click", exportSummaryXlsx);
    elements.exportCsvBtn.addEventListener("click", exportSummaryCsv);
    elements.exportDetailsBtn.addEventListener("click", exportDetailsCsv);
    elements.exportCaseXlsxBtn.addEventListener("click", exportCurrentCaseXlsx);
    elements.exportCaseCsvBtn.addEventListener("click", exportCurrentCaseCsv);
    elements.closeDrawerBtn.addEventListener("click", closeDrawer);
    elements.toggleWarningsBtn.addEventListener("click", toggleWarnings);
    elements.resetClueBtn.addEventListener("click", resetClueSelection);
    elements.navItems.forEach((item) => {
      item.addEventListener("click", () => setActiveView(item.dataset.viewTarget));
    });
    elements.sidebarCollapseBtn.addEventListener("click", () => setSidebarCollapsed(!state.sidebarCollapsed));
    elements.privacyConfirmBtn.addEventListener("click", closePrivacyModal);
    elements.privacyModal.addEventListener("click", (event) => {
      if (event.target === elements.privacyModal) closePrivacyModal();
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && !elements.privacyModal.hidden) closePrivacyModal();
    });
  }

  function setActiveView(viewName) {
    const nextView = viewName || "buyers";
    state.activeView = nextView;
    elements.navItems.forEach((item) => {
      const isActive = item.dataset.viewTarget === nextView;
      item.classList.toggle("active", isActive);
      item.setAttribute("aria-current", isActive ? "page" : "false");
    });
    elements.viewPanels.forEach((panel) => {
      const isActive = panel.dataset.view === nextView;
      panel.classList.toggle("active", isActive);
      panel.hidden = !isActive;
    });
    requestAnimationFrame(resizeCharts);
  }

  function setSidebarCollapsed(collapsed) {
    state.sidebarCollapsed = Boolean(collapsed);
    elements.appShell.classList.toggle("sidebar-collapsed", state.sidebarCollapsed);
    elements.sidebarCollapseBtn.setAttribute("aria-expanded", String(!state.sidebarCollapsed));
    elements.sidebarCollapseBtn.setAttribute(
      "aria-label",
      state.sidebarCollapsed ? "展開側邊欄" : "縮放側邊欄"
    );
    requestAnimationFrame(resizeCharts);
  }

  function toggleTheme() {
    applyTheme(state.theme === "dark" ? "light" : "dark", true);
    render();
  }

  function applyTheme(theme, persist) {
    state.theme = theme === "light" ? "light" : "dark";
    document.documentElement.dataset.theme = state.theme;
    updateThemeToggle();
    setupChartsDefaults();

    if (persist) {
      try {
        localStorage.setItem(STORAGE_KEYS.theme, state.theme);
      } catch (error) {
        // Theme persistence is optional; the active theme is still applied.
      }
    }
  }

  function updateThemeToggle() {
    if (!elements.themeToggleBtn) return;
    const isDark = state.theme === "dark";
    elements.themeToggleBtn.setAttribute("aria-pressed", String(isDark));
    elements.themeToggleIcon.textContent = isDark ? "☾" : "☀";
    elements.themeToggleLabel.textContent = isDark ? "深色模式：開" : "深色模式：關";
  }

  function maybeShowPrivacyModal() {
    let hasSeen = false;
    try {
      hasSeen = localStorage.getItem(STORAGE_KEYS.welcome) === "true";
    } catch (error) {
      hasSeen = false;
    }
    if (!hasSeen) {
      elements.privacyModal.hidden = false;
      elements.privacyConfirmBtn.focus();
    }
  }

  function closePrivacyModal() {
    elements.privacyModal.hidden = true;
    try {
      localStorage.setItem(STORAGE_KEYS.welcome, "true");
    } catch (error) {
      // localStorage can be unavailable in private contexts; closing still works for this session.
    }
  }

  function resizeCharts() {
    Object.values(state.charts).forEach((chart) => {
      if (chart && typeof chart.resize === "function") chart.resize();
    });
  }

  function setupChartsDefaults() {
    if (typeof Chart === "undefined") return;
    Chart.defaults.color = getCssVariable("--chart-text", state.theme === "dark" ? "#ffffff" : "#171717");
    Chart.defaults.font.family =
      '-apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans TC", "Microsoft JhengHei", sans-serif';
    Chart.defaults.font.size = 14;
    Chart.defaults.borderColor = getCssVariable(
      "--chart-grid",
      state.theme === "dark" ? "rgba(255, 255, 255, 0.16)" : "rgba(0, 0, 0, 0.14)"
    );
  }

  function getCssVariable(name, fallback) {
    const value = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return value || fallback;
  }

  function getChartColor(name, fallback) {
    return getCssVariable(name, fallback);
  }

  async function handleFile(file) {
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      setNotice("檔案格式不符", "請匯入 .xlsx 或 .xls Excel 檔。", "error");
      return;
    }

    try {
      setNotice("解析中", "正在讀取 Excel 並建立買家統計。", "info");
      const buffer = await file.arrayBuffer();
      const parsed = parseWorkbook(buffer);
      state.records = parsed.records;
      state.warnings = parsed.warnings;
      state.fileName = file.name;
      state.sheetName = parsed.sheetName;
      state.sourceRowCount = parsed.sourceRowCount;
      elements.fileMeta.textContent = `${file.name} / 工作表：${parsed.sheetName} / ${numberFormatter.format(
        parsed.sourceRowCount
      )} 列`;
      setNotice(
        "已載入",
        `完成解析 ${numberFormatter.format(parsed.records.length)} 筆有效列，預設分析 COMPLETED 訂單。`,
        "success"
      );
      render();
    } catch (error) {
      state.records = [];
      state.warnings = [];
      state.currentSummaries = [];
      state.currentRows = [];
      elements.fileMeta.textContent = "尚未載入檔案";
      setNotice("解析失敗", error.message || String(error), "error");
      render();
    } finally {
      elements.fileInput.value = "";
    }
  }

  function resetClueSelection() {
    state.activeClue = null;
    render();
  }

  function render() {
    const aggregation = aggregateRecords(state.records, {});
    const summaries = applyDisplayFilters(aggregation.summaries, {
      columnFilters: state.summaryColumnFilters,
      sortKey: state.summarySortKey,
      sortDirection: state.summarySortDirection,
    });
    state.currentSummaries = summaries;
    state.currentRows = buildDetailRows(summaries);

    renderKpis(summaries);
    renderTable(summaries);
    renderCharts(summaries);
    renderWarnings(state.warnings);
    updateExportState(summaries.length > 0);
    requestAnimationFrame(resizeCharts);
  }

  function renderKpis(summaries) {
    const orderCount = summaries.reduce((sum, item) => sum + item.orderCount, 0);
    const totalAmount = summaries.reduce((sum, item) => sum + item.totalAmount, 0);
    const productSet = new Set();
    let minDate = Infinity;
    let maxDate = -Infinity;

    summaries.forEach((summary) => {
      summary.products.forEach((product) => productSet.add(product.name));
      if (summary.firstDateMs !== null && summary.firstDateMs < minDate) minDate = summary.firstDateMs;
      if (summary.latestDateMs !== null && summary.latestDateMs > maxDate) maxDate = summary.latestDateMs;
    });

    elements.kpiBuyers.textContent = numberFormatter.format(summaries.length);
    elements.kpiOrders.textContent = numberFormatter.format(orderCount);
    elements.kpiAmount.textContent = currencyFormatter.format(Math.round(totalAmount));
    elements.kpiProducts.textContent = numberFormatter.format(productSet.size);
    elements.kpiDateRange.textContent =
      minDate === Infinity || maxDate === -Infinity
        ? "-"
        : `${formatDateOnly(minDate)} ~ ${formatDateOnly(maxDate)}`;
    elements.resultCount.textContent = `${numberFormatter.format(summaries.length)} 筆`;
  }

  function renderTable(summaries) {
    const columns = getOrderedSummaryColumns();
    renderSummaryHeader(columns);
    elements.summaryBody.textContent = "";
    if (!summaries.length) {
      const row = document.createElement("tr");
      const cell = document.createElement("td");
      cell.colSpan = columns.length;
      cell.className = "empty-cell";
      cell.textContent = state.records.length ? "沒有符合目前篩選條件的買家" : "尚未載入資料";
      row.appendChild(cell);
      elements.summaryBody.appendChild(row);
      return;
    }

    const fragment = document.createDocumentFragment();
    summaries.forEach((summary) => {
      const row = document.createElement("tr");
      row.tabIndex = 0;
      row.className = "summary-row";
      row.addEventListener("click", () => openDrawer(summary));
      row.addEventListener("keydown", (event) => {
        if (event.key === "Enter") openDrawer(summary);
      });
      columns.forEach((column) => {
        const cell = document.createElement("td");
        cell.dataset.columnKey = column.key;
        cell.textContent = column.display(summary) || "-";
        if (column.multiline) cell.classList.add("multiline");
        if (column.key === "buyer") cell.classList.add("primary-cell");
        row.appendChild(cell);
      });
      fragment.appendChild(row);
    });
    elements.summaryBody.appendChild(fragment);
  }

  function renderSummaryHeader(columns) {
    elements.summaryHeadRow.textContent = "";
    elements.summaryColgroup.textContent = "";

    let tableWidth = 0;
    columns.forEach((column) => {
      const width = getColumnWidth(column);
      tableWidth += width;

      const col = document.createElement("col");
      col.dataset.columnKey = column.key;
      col.style.width = `${width}px`;
      elements.summaryColgroup.appendChild(col);

      const th = document.createElement("th");
      th.dataset.columnKey = column.key;
      th.className = "summary-header-cell";
      th.draggable = true;
      th.addEventListener("dragstart", (event) => startColumnDrag(event, column.key));
      th.addEventListener("dragover", (event) => handleColumnDragOver(event, column.key));
      th.addEventListener("dragleave", () => th.classList.remove("drag-over"));
      th.addEventListener("drop", (event) => finishColumnDrop(event, column.key));
      th.addEventListener("dragend", clearColumnDragState);

      const shell = document.createElement("div");
      shell.className = "column-shell";

      const titleRow = document.createElement("div");
      titleRow.className = "column-title-row";

      const dragHandle = document.createElement("span");
      dragHandle.className = "column-drag-handle";
      dragHandle.textContent = "⋮⋮";
      dragHandle.title = "拖曳欄位";

      const sortButton = document.createElement("button");
      sortButton.type = "button";
      sortButton.className = "column-sort";
      sortButton.setAttribute("aria-label", `依${column.label}排序`);
      sortButton.addEventListener("click", () => setSummarySort(column.key));

      const label = document.createElement("span");
      label.textContent = column.label;
      const indicator = document.createElement("span");
      indicator.className = "sort-indicator";
      indicator.textContent =
        state.summarySortKey === column.key ? (state.summarySortDirection === "asc" ? "↑" : "↓") : "";
      sortButton.append(label, indicator);

      const filter = document.createElement("input");
      filter.className = "column-filter";
      filter.type = "search";
      filter.value = state.summaryColumnFilters[column.key] || "";
      filter.placeholder = column.filterPlaceholder || "篩選";
      filter.setAttribute("aria-label", `篩選${column.label}`);
      filter.addEventListener("click", (event) => event.stopPropagation());
      filter.addEventListener("input", (event) => updateColumnFilter(column.key, event.target));

      const resizer = document.createElement("span");
      resizer.className = "column-resizer";
      resizer.tabIndex = 0;
      resizer.setAttribute("role", "separator");
      resizer.setAttribute("aria-orientation", "vertical");
      resizer.setAttribute("aria-label", `調整${column.label}欄寬`);
      resizer.addEventListener("mousedown", (event) => startColumnResize(event, column.key));
      resizer.addEventListener("keydown", (event) => handleColumnResizeKey(event, column.key));

      titleRow.append(dragHandle, sortButton);
      shell.append(titleRow, filter, resizer);
      th.appendChild(shell);
      elements.summaryHeadRow.appendChild(th);
    });

    elements.summaryTable.style.width = `${tableWidth}px`;
    restoreColumnFilterFocus();
  }

  function setSummarySort(columnKey) {
    state.activeColumnFilterKey = "";
    if (state.summarySortKey === columnKey) {
      state.summarySortDirection = state.summarySortDirection === "asc" ? "desc" : "asc";
    } else {
      const column = getSummaryColumn(columnKey);
      state.summarySortKey = columnKey;
      state.summarySortDirection = column && (column.type === "number" || column.type === "date") ? "desc" : "asc";
    }
    render();
  }

  function updateColumnFilter(columnKey, input) {
    state.activeColumnFilterKey = columnKey;
    state.activeColumnFilterStart = input.selectionStart;
    state.activeColumnFilterEnd = input.selectionEnd;
    const nextValue = cleanCell(input.value);
    if (nextValue) {
      state.summaryColumnFilters[columnKey] = nextValue;
    } else {
      delete state.summaryColumnFilters[columnKey];
    }
    render();
  }

  function restoreColumnFilterFocus() {
    if (!state.activeColumnFilterKey) return;
    requestAnimationFrame(() => {
      const input = elements.summaryHeadRow.querySelector(
        `th[data-column-key="${state.activeColumnFilterKey}"] .column-filter`
      );
      if (!input) return;
      input.focus({ preventScroll: true });
      if (
        state.activeColumnFilterStart !== null &&
        state.activeColumnFilterEnd !== null &&
        typeof input.setSelectionRange === "function"
      ) {
        input.setSelectionRange(state.activeColumnFilterStart, state.activeColumnFilterEnd);
      }
    });
  }

  function startColumnDrag(event, columnKey) {
    const target = event.target instanceof Element ? event.target : event.currentTarget;
    if (target.closest("input, button, .column-resizer")) {
      event.preventDefault();
      return;
    }
    state.draggedColumnKey = columnKey;
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", columnKey);
  }

  function handleColumnDragOver(event, columnKey) {
    if (!state.draggedColumnKey || state.draggedColumnKey === columnKey) return;
    event.preventDefault();
    event.currentTarget.classList.add("drag-over");
  }

  function finishColumnDrop(event, targetKey) {
    event.preventDefault();
    const draggedKey = event.dataTransfer.getData("text/plain") || state.draggedColumnKey;
    if (!draggedKey || draggedKey === targetKey) {
      clearColumnDragState();
      return;
    }

    const order = getOrderedSummaryColumns().map((column) => column.key);
    const fromIndex = order.indexOf(draggedKey);
    const toIndex = order.indexOf(targetKey);
    if (fromIndex >= 0 && toIndex >= 0) {
      order.splice(fromIndex, 1);
      order.splice(toIndex, 0, draggedKey);
      state.summaryColumnOrder = order;
      saveTablePreferences();
      render();
    }
    clearColumnDragState();
  }

  function clearColumnDragState() {
    state.draggedColumnKey = "";
    document.querySelectorAll(".summary-header-cell.drag-over").forEach((cell) => {
      cell.classList.remove("drag-over");
    });
  }

  function startColumnResize(event, columnKey) {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const startWidth = getColumnWidth(getSummaryColumn(columnKey));
    document.body.classList.add("is-resizing-column");

    const onMouseMove = (moveEvent) => {
      state.summaryColumnWidths[columnKey] = Math.max(
        MIN_SUMMARY_COLUMN_WIDTH,
        Math.round(startWidth + moveEvent.clientX - startX)
      );
      applyColumnWidths();
    };
    const onMouseUp = () => {
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
      document.body.classList.remove("is-resizing-column");
      saveTablePreferences();
    };

    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
  }

  function handleColumnResizeKey(event, columnKey) {
    if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") return;
    event.preventDefault();
    const delta = event.key === "ArrowRight" ? 16 : -16;
    const column = getSummaryColumn(columnKey);
    state.summaryColumnWidths[columnKey] = Math.max(MIN_SUMMARY_COLUMN_WIDTH, getColumnWidth(column) + delta);
    applyColumnWidths();
    saveTablePreferences();
  }

  function applyColumnWidths() {
    let tableWidth = 0;
    getOrderedSummaryColumns().forEach((column) => {
      const width = getColumnWidth(column);
      tableWidth += width;
      const col = elements.summaryColgroup.querySelector(`col[data-column-key="${column.key}"]`);
      if (col) col.style.width = `${width}px`;
    });
    elements.summaryTable.style.width = `${tableWidth}px`;
  }

  function renderCharts(summaries) {
    renderAmountChart(summaries);
    renderCountChart(summaries);
    renderProductChart(summaries);
    renderMonthChart(summaries);
    renderClueList(summaries);
  }

  function renderAmountChart(summaries) {
    const items = summaries.slice(0, 20);
    createBarChart("amountChart", "amountChart", {
      labels: items.map((item) => item.buyer),
      values: items.map((item) => Math.round(item.totalAmount)),
      label: "實際撥款額",
      color: getChartColor("--chart-series-1", "#ef2323"),
      indexAxis: "y",
    });
  }

  function renderCountChart(summaries) {
    const items = [...summaries]
      .sort((a, b) => b.orderCount - a.orderCount || b.totalAmount - a.totalAmount)
      .slice(0, 20);
    createBarChart("countChart", "countChart", {
      labels: items.map((item) => item.buyer),
      values: items.map((item) => item.orderCount),
      label: "購買次數",
      color: getChartColor("--chart-series-2", "#ffffff"),
      indexAxis: "y",
    });
  }

  function renderProductChart(summaries) {
    const products = new Map();
    summaries.forEach((summary) => {
      summary.products.forEach((item) => {
        if (!products.has(item.name)) products.set(item.name, 0);
        products.set(item.name, products.get(item.name) + item.quantity);
      });
    });
    const items = Array.from(products, ([name, quantity]) => ({ name, quantity }))
      .sort((a, b) => b.quantity - a.quantity || a.name.localeCompare(b.name))
      .slice(0, 20);
    createBarChart("productChart", "productChart", {
      labels: items.map((item) => item.name),
      values: items.map((item) => item.quantity),
      label: "商品數量",
      color: getChartColor("--chart-series-3", "#cfcfcf"),
      indexAxis: "y",
    });
  }

  function renderMonthChart(summaries) {
    const monthMap = new Map();
    summaries.forEach((summary) => {
      summary.orders.forEach((order) => {
        const month = order.dateText && order.dateText.length >= 7 ? order.dateText.slice(0, 7) : "未解析";
        if (!monthMap.has(month)) monthMap.set(month, { month, count: 0, amount: 0 });
        const bucket = monthMap.get(month);
        bucket.count += 1;
        bucket.amount += order.amount;
      });
    });
    const items = Array.from(monthMap.values()).sort((a, b) => a.month.localeCompare(b.month));
    destroyChart("monthChart");
    const canvas = document.getElementById("monthChart");
    if (!canvas || typeof Chart === "undefined") return;
    state.charts.monthChart = new Chart(canvas, {
      type: "bar",
      data: {
        labels: items.map((item) => item.month),
        datasets: [
          {
            type: "bar",
            label: "訂單數",
            data: items.map((item) => item.count),
            backgroundColor: withAlpha(getChartColor("--chart-series-1", "#ef2323"), 0.72),
            borderColor: getChartColor("--chart-series-1", "#ef2323"),
            borderWidth: 1,
            yAxisID: "y",
          },
          {
            type: "line",
            label: "撥款額",
            data: items.map((item) => Math.round(item.amount)),
            borderColor: getChartColor("--chart-series-2", "#ffffff"),
            backgroundColor: withAlpha(getChartColor("--chart-series-2", "#ffffff"), 0.16),
            tension: 0.3,
            yAxisID: "y1",
          },
        ],
      },
      options: chartBaseOptions({
        scales: {
          y: { beginAtZero: true, grid: { color: getChartColor("--chart-grid", "rgba(255, 255, 255, 0.16)") } },
          y1: {
            beginAtZero: true,
            position: "right",
            grid: { drawOnChartArea: false },
            ticks: { callback: (value) => numberFormatter.format(value) },
          },
        },
      }),
    });
  }

  function renderClueList(summaries) {
    destroyChart("linkChart");
    const clues = buildSharedClues(summaries);
    if (state.activeClue && !clues.some((clue) => isActiveClue(clue))) {
      state.activeClue = null;
    }

    const summaryByBuyer = new Map(summaries.map((summary) => [summary.buyer, summary]));
    elements.clueList.textContent = "";

    if (!clues.length) {
      const empty = document.createElement("p");
      empty.className = "empty-note";
      empty.textContent = summaries.length ? "目前篩選結果沒有共用電話或姓名線索。" : "載入資料後顯示線索。";
      elements.clueList.appendChild(empty);
      return;
    }

    const fragment = document.createDocumentFragment();
    clues.forEach((clue) => {
      const row = document.createElement("article");
      row.className = "clue-row";
      row.classList.toggle("expanded", isActiveClue(clue));

      const summaryButton = document.createElement("button");
      summaryButton.type = "button";
      summaryButton.className = "clue-row-summary";
      summaryButton.setAttribute("aria-expanded", String(isActiveClue(clue)));
      summaryButton.addEventListener("click", () => toggleClue(clue));

      const title = document.createElement("strong");
      title.textContent = `${clue.type} ${clue.value}`;
      const count = document.createElement("span");
      count.textContent = `${numberFormatter.format(clue.buyers.length)} 個共用帳號`;
      const indicator = document.createElement("span");
      indicator.className = "clue-row-indicator";
      indicator.textContent = isActiveClue(clue) ? "收合" : "展開";
      summaryButton.append(title, count, indicator);
      row.appendChild(summaryButton);

      if (isActiveClue(clue)) {
        row.appendChild(renderClueBuyers(clue, summaryByBuyer));
      }

      fragment.appendChild(row);
    });
    elements.clueList.appendChild(fragment);
  }

  function toggleClue(clue) {
    state.activeClue = isActiveClue(clue) ? null : { type: clue.type, value: clue.value };
    render();
  }

  function isActiveClue(clue) {
    return Boolean(state.activeClue && state.activeClue.type === clue.type && state.activeClue.value === clue.value);
  }

  function renderClueBuyers(clue, summaryByBuyer) {
    const detail = document.createElement("div");
    detail.className = "clue-row-detail";
    const buyers = document.createElement("div");
    buyers.className = "clue-buyer-list";

    clue.buyers.forEach((buyer) => {
      const summary = summaryByBuyer.get(buyer);
      if (!summary) return;
      const card = document.createElement("button");
      card.type = "button";
      card.className = "clue-buyer-card";
      card.addEventListener("click", () => openDrawer(summary));

      const name = document.createElement("strong");
      name.textContent = buyer;
      const meta = document.createElement("span");
      meta.textContent = `撥款 ${currencyFormatter.format(Math.round(summary.totalAmount))} / ${numberFormatter.format(
        summary.orderCount
      )} 筆完成訂單`;
      const people = document.createElement("span");
      people.textContent = `姓名：${summary.nameText || "-"} / 電話：${summary.phoneText || "-"}`;
      card.append(name, meta, people);
      buyers.appendChild(card);
    });

    detail.appendChild(buyers);
    return detail;
  }

  function buildSharedClues(summaries) {
    const map = new Map();
    const add = (type, value, buyer) => {
      if (!value) return;
      const key = `${type}:${value}`;
      if (!map.has(key)) map.set(key, { type, value, buyers: new Set() });
      map.get(key).buyers.add(buyer);
    };
    summaries.forEach((summary) => {
      summary.phones.forEach((phone) => add("電話", phone, summary.buyer));
      summary.names.forEach((name) => add("姓名", name, summary.buyer));
    });
    return Array.from(map.values())
      .map((item) => ({ ...item, buyers: Array.from(item.buyers).sort() }))
      .filter((item) => item.buyers.length > 1)
      .sort((a, b) => b.buyers.length - a.buyers.length || a.value.localeCompare(b.value));
  }

  function createBarChart(chartKey, canvasId, config) {
    destroyChart(chartKey);
    const canvas = document.getElementById(canvasId);
    if (!canvas || typeof Chart === "undefined") return;
    const labels = config.labels;
    state.charts[chartKey] = new Chart(canvas, {
      type: "bar",
      data: {
        labels: labels.map((label) => truncateLabel(label)),
        datasets: [
          {
            label: config.label,
            data: config.values,
            backgroundColor: withAlpha(config.color, 0.72),
            borderColor: config.color,
            borderWidth: 1,
          },
        ],
      },
      options: chartBaseOptions(
        mergeOptions(
          {
            indexAxis: config.indexAxis || "x",
            plugins: {
              tooltip: {
                callbacks: {
                  title: (items) => (items.length ? labels[items[0].dataIndex] : ""),
                  label: (item) =>
                    `${config.label}: ${numberFormatter.format(item.parsed.x || item.parsed.y || 0)}`,
                },
              },
            },
          },
          config.options || {}
        )
      ),
    });
  }

  function chartBaseOptions(extra) {
    return mergeOptions(
      {
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        plugins: {
          legend: { labels: { boxWidth: 12, boxHeight: 12 } },
          tooltip: { mode: "index", intersect: false },
        },
        scales: {
          x: {
            beginAtZero: true,
            grid: { color: getChartColor("--chart-grid", "rgba(255, 255, 255, 0.16)") },
            ticks: { callback: (value) => numberFormatter.format(value) },
          },
          y: {
            grid: { color: getChartColor("--chart-grid-soft", "rgba(255, 255, 255, 0.1)") },
          },
        },
      },
      extra || {}
    );
  }

  function mergeOptions(base, extra) {
    const merged = { ...base };
    Object.keys(extra).forEach((key) => {
      if (
        extra[key] &&
        typeof extra[key] === "object" &&
        !Array.isArray(extra[key]) &&
        base[key] &&
        typeof base[key] === "object"
      ) {
        merged[key] = mergeOptions(base[key], extra[key]);
      } else {
        merged[key] = extra[key];
      }
    });
    return merged;
  }

  function withAlpha(hex, alpha) {
    const value = hex.replace("#", "");
    const r = parseInt(value.slice(0, 2), 16);
    const g = parseInt(value.slice(2, 4), 16);
    const b = parseInt(value.slice(4, 6), 16);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
  }

  function truncateLabel(label) {
    const text = cleanCell(label);
    return text.length > 22 ? `${text.slice(0, 21)}...` : text;
  }

  function destroyChart(chartKey) {
    if (state.charts[chartKey]) {
      state.charts[chartKey].destroy();
      delete state.charts[chartKey];
    }
  }

  function openDrawer(summary) {
    state.currentCaseSummary = summary;
    elements.drawerBuyer.textContent = summary.buyer;
    elements.drawerContent.textContent = "";
    updateCaseExportState(true);

    const stats = document.createElement("div");
    stats.className = "drawer-stats";
    [
      ["實際撥款額", currencyFormatter.format(Math.round(summary.totalAmount))],
      ["完成訂單", `${summary.orderCount} 筆`],
      ["電話", summary.phoneText || "-"],
      ["姓名", summary.nameText || "-"],
    ].forEach(([label, value]) => {
      const block = document.createElement("div");
      const small = document.createElement("span");
      const strong = document.createElement("strong");
      small.textContent = label;
      strong.textContent = value;
      block.append(small, strong);
      stats.appendChild(block);
    });

    const products = document.createElement("div");
    products.className = "drawer-section";
    const productsTitle = document.createElement("h3");
    productsTitle.textContent = "商品統整";
    const productsList = document.createElement("ul");
    summary.products.forEach((item) => {
      const li = document.createElement("li");
      li.textContent = `${item.name} x ${item.quantity}`;
      productsList.appendChild(li);
    });
    if (!summary.products.length) {
      const li = document.createElement("li");
      li.textContent = "沒有商品數量資料";
      productsList.appendChild(li);
    }
    products.append(productsTitle, productsList);

    const orders = document.createElement("div");
    orders.className = "drawer-section";
    const ordersTitle = document.createElement("h3");
    ordersTitle.textContent = "訂單明細";
    const orderWrap = document.createElement("div");
    orderWrap.className = "mini-table-wrap";
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    ["日期", "訂單編號", "金額", "商品", "收件地址", "付款方式", "寄送方式"].forEach((header) => {
      const th = document.createElement("th");
      th.textContent = header;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    const tbody = document.createElement("tbody");
    summary.orders.forEach((order) => {
      const orderRows = summary.detailRows.filter((row) => row.orderId === order.orderId);
      const tr = document.createElement("tr");
      [
        order.dateText || "-",
        order.orderIdDisplay || order.orderId,
        currencyFormatter.format(Math.round(order.amount)),
        orderRows.map((row) => `${row.product || "-"} x ${row.quantity}`).join("\n"),
        joinUniqueValues(orderRows, "address") || "-",
        joinUniqueValues(orderRows, "payment") || "-",
        joinUniqueValues(orderRows, "shipping") || "-",
      ].forEach((value, index) => {
        const td = document.createElement("td");
        td.textContent = value;
        if (index >= 3) td.className = "multiline";
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.append(thead, tbody);
    orderWrap.appendChild(table);
    orders.append(ordersTitle, orderWrap);

    elements.drawerContent.append(stats, products, orders);
    elements.detailDrawer.classList.add("open");
    elements.detailDrawer.setAttribute("aria-hidden", "false");
  }

  function closeDrawer() {
    elements.detailDrawer.classList.remove("open");
    elements.detailDrawer.setAttribute("aria-hidden", "true");
    state.currentCaseSummary = null;
    updateCaseExportState(false);
  }

  function setNotice(title, body, type) {
    if (!elements.noticePanel) return;
    elements.noticePanel.classList.remove("success", "error", "info");
    if (type) elements.noticePanel.classList.add(type);
    const strong = elements.noticePanel.querySelector("strong");
    const span = elements.noticePanel.querySelector("span");
    strong.textContent = title;
    span.textContent = body;
  }

  function renderWarnings(warnings) {
    elements.warningCount.textContent = numberFormatter.format(warnings.length);
    elements.warningPanel.classList.toggle("hidden", warnings.length === 0);
    elements.warningList.textContent = "";
    warnings.slice(0, 80).forEach((warning) => {
      const item = document.createElement("li");
      item.textContent = warning;
      elements.warningList.appendChild(item);
    });
    if (warnings.length > 80) {
      const item = document.createElement("li");
      item.textContent = `另有 ${numberFormatter.format(warnings.length - 80)} 則提示未列出。`;
      elements.warningList.appendChild(item);
    }
  }

  function toggleWarnings() {
    const expanded = elements.toggleWarningsBtn.getAttribute("aria-expanded") === "true";
    elements.toggleWarningsBtn.setAttribute("aria-expanded", String(!expanded));
    elements.warningList.hidden = expanded;
  }

  function updateExportState(enabled) {
    elements.exportXlsxBtn.disabled = !enabled;
    elements.exportCsvBtn.disabled = !enabled;
    elements.exportDetailsBtn.disabled = !enabled;
  }

  function updateCaseExportState(enabled) {
    elements.exportCaseXlsxBtn.disabled = !enabled;
    elements.exportCaseCsvBtn.disabled = !enabled;
  }

  function exportSummaryXlsx() {
    if (!state.currentSummaries.length) return;
    const aoa = [OUTPUT_HEADERS, ...buildSummaryRows(state.currentSummaries)];
    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    worksheet["!cols"] = [
      { wch: 18 },
      { wch: 16 },
      { wch: 12 },
      { wch: 26 },
      { wch: 70 },
      { wch: 22 },
      { wch: 20 },
      { wch: 46 },
      { wch: 18 },
      { wch: 22 },
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "買家統計");
    XLSX.writeFile(workbook, `買家統計分析_${timestamp()}.xlsx`, { compression: true });
  }

  function exportSummaryCsv() {
    if (!state.currentSummaries.length) return;
    downloadCsv([OUTPUT_HEADERS, ...buildSummaryRows(state.currentSummaries)], `買家統計分析_${timestamp()}.csv`);
  }

  function exportDetailsCsv() {
    if (!state.currentSummaries.length) return;
    downloadCsv([DETAIL_HEADERS, ...buildDetailRows(state.currentSummaries)], `買家明細_${timestamp()}.csv`);
  }

  function exportCurrentCaseXlsx() {
    const summary = state.currentCaseSummary;
    if (!summary) return;

    const workbook = XLSX.utils.book_new();
    const summaryWorksheet = XLSX.utils.aoa_to_sheet([OUTPUT_HEADERS, ...buildSummaryRows([summary])]);
    summaryWorksheet["!cols"] = [
      { wch: 18 },
      { wch: 16 },
      { wch: 12 },
      { wch: 26 },
      { wch: 70 },
      { wch: 22 },
      { wch: 20 },
      { wch: 46 },
      { wch: 18 },
      { wch: 22 },
    ];

    const detailWorksheet = XLSX.utils.aoa_to_sheet([DETAIL_HEADERS, ...buildDetailRows([summary])]);
    detailWorksheet["!cols"] = [
      { wch: 18 },
      { wch: 20 },
      { wch: 14 },
      { wch: 22 },
      { wch: 14 },
      { wch: 10 },
      { wch: 70 },
      { wch: 18 },
      { wch: 18 },
      { wch: 44 },
      { wch: 18 },
      { wch: 16 },
      { wch: 18 },
      { wch: 10 },
    ];

    XLSX.utils.book_append_sheet(workbook, summaryWorksheet, "買家統計");
    XLSX.utils.book_append_sheet(workbook, detailWorksheet, "訂單明細");
    XLSX.writeFile(workbook, `Case_Detail_${safeFilename(summary.buyer)}_${timestamp()}.xlsx`, { compression: true });
  }

  function exportCurrentCaseCsv() {
    const summary = state.currentCaseSummary;
    if (!summary) return;
    downloadCsv([DETAIL_HEADERS, ...buildDetailRows([summary])], `Case_Detail_${safeFilename(summary.buyer)}_${timestamp()}.csv`);
  }

  function downloadCsv(aoa, filename) {
    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    const csv = XLSX.utils.sheet_to_csv(worksheet);
    const blob = new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function safeFilename(value) {
    const cleaned = cleanCell(value)
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, "_")
      .slice(0, 60);
    return cleaned || "buyer";
  }

  function timestamp() {
    const now = new Date();
    return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(
      now.getDate()
    ).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
  }

  const api = {
    REQUIRED_COLUMNS,
    OUTPUT_HEADERS,
    parseWorkbook,
    parseRows,
    aggregateRecords,
    applyDisplayFilters,
    buildSummaryRows,
    buildDetailRows,
  };

  if (typeof globalThis !== "undefined") {
    globalThis.ShellCaseAnalyzer = api;
  }
  if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", init);
  }
})();
