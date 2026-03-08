#!/usr/bin/env osascript -l JavaScript

// Write data to an Apple Numbers spreadsheet
// Usage: osascript -l JavaScript write-numbers.js "/path/to/file.numbers" '<json>'
//
// JSON format (single cell):
//   {"sheet":"Sheet 1","table":"Table 1","row":1,"col":1,"value":"Hello"}
//
// JSON format (batch):
//   [{"sheet":"Sheet 1","table":"Table 1","row":1,"col":1,"value":"A"},{"row":1,"col":2,"value":"B"}]
//
// JSON format (append one row):
//   {"sheet":"Sheet 1","table":"Table 1","appendRow":["AAPL","Apple Inc."]}
//
// JSON format (append multiple rows):
//   {"sheet":"Sheet 1","table":"Table 1","appendRows":[["AAPL","Apple Inc."],["MSFT","Microsoft Corporation"]]}
//
// row/col are 0-based. If sheet/table omitted, uses first sheet/table.

ObjC.import("Foundation");

function normalizePath(path) {
  return ObjC.unwrap($(path).stringByStandardizingPath);
}

function documentPath(doc) {
  try {
    return normalizePath(doc.file().toString());
  } catch (e) {
    return null;
  }
}

function findOpenDocument(Numbers, filePath) {
  const targetPath = normalizePath(filePath);
  const openDocs = Numbers.documents();

  for (let i = 0; i < openDocs.length; i++) {
    if (documentPath(openDocs[i]) === targetPath) {
      return openDocs[i];
    }
  }

  return null;
}

function openDocument(Numbers, filePath) {
  const targetPath = normalizePath(filePath);
  const existingDoc = findOpenDocument(Numbers, targetPath);

  if (existingDoc) {
    return { doc: existingDoc, wasOpen: true };
  }

  Numbers.open(Path(targetPath));

  for (let attempt = 0; attempt < 20; attempt++) {
    const openedDoc = findOpenDocument(Numbers, targetPath);
    if (openedDoc) {
      return { doc: openedDoc, wasOpen: false };
    }
    delay(0.1);
  }

  throw new Error("Document not found after opening: " + filePath);
}

function resolveSheet(doc, sheetName) {
  if (sheetName) {
    const sheets = doc.sheets();
    const sheet = sheets.find(s => s.name() === sheetName);
    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }
    return sheet;
  }

  return doc.sheets[0];
}

function resolveTable(sheet, tableName) {
  if (tableName) {
    const tables = sheet.tables();
    const table = tables.find(t => t.name() === tableName);
    if (!table) {
      throw new Error("Table not found: " + tableName);
    }
    return table;
  }

  return sheet.tables[0];
}

function normalizeRows(op) {
  if (Object.prototype.hasOwnProperty.call(op, "appendRow")) {
    if (!Array.isArray(op.appendRow)) {
      throw new Error("appendRow must be an array");
    }
    return [op.appendRow];
  }

  if (Object.prototype.hasOwnProperty.call(op, "appendRows")) {
    if (!Array.isArray(op.appendRows)) {
      throw new Error("appendRows must be an array of rows");
    }
    for (let i = 0; i < op.appendRows.length; i++) {
      if (!Array.isArray(op.appendRows[i])) {
        throw new Error("appendRows must contain only row arrays");
      }
    }
    return op.appendRows;
  }

  return null;
}

function appendRows(table, rows) {
  if (!rows || rows.length === 0) {
    return { rowsAppended: 0, cellsUpdated: 0 };
  }

  const colCount = table.columnCount();
  const startRow = table.rowCount();

  for (let i = 0; i < rows.length; i++) {
    if (rows[i].length > colCount) {
      throw new Error("Append row is wider than table: " + rows[i].length + " > " + colCount);
    }
  }

  table.rowCount = startRow + rows.length;

  let cellsUpdated = 0;
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      table.cells[(startRow + r) * colCount + c].value = rows[r][c];
      cellsUpdated++;
    }
  }

  return { rowsAppended: rows.length, cellsUpdated: cellsUpdated };
}

function run(argv) {
  const filePath = argv[0];
  const jsonStr = argv[1];

  if (!filePath || !jsonStr) {
    return JSON.stringify({ error: "Usage: write-numbers.js <file-path> <json>" });
  }

  let operations;
  try {
    const parsed = JSON.parse(jsonStr);
    operations = Array.isArray(parsed) ? parsed : [parsed];
  } catch (e) {
    return JSON.stringify({ error: "Invalid JSON: " + e.message });
  }

  const Numbers = Application("Numbers");
  const app = Application.currentApplication();
  app.includeStandardAdditions = true;

  let doc;
  let wasOpen = false;

  try {
    const opened = openDocument(Numbers, filePath);
    doc = opened.doc;
    wasOpen = opened.wasOpen;

    let updated = 0;
    let rowsAppended = 0;

    for (const op of operations) {
      const sheetName = op.sheet || null;
      const tableName = op.table || null;
      const sheet = resolveSheet(doc, sheetName);
      const table = resolveTable(sheet, tableName);
      const rows = normalizeRows(op);

      if (rows) {
        const appendResult = appendRows(table, rows);
        updated += appendResult.cellsUpdated;
        rowsAppended += appendResult.rowsAppended;
        continue;
      }

      const row = op.row;
      const col = op.col;
      const value = op.value;
      if (row === undefined || col === undefined || value === undefined) {
        continue;
      }

      const colCount = table.columnCount();
      const cellIndex = row * colCount + col;
      table.cells[cellIndex].value = value;
      updated++;
    }

    doc.save();

    return JSON.stringify({ success: true, cellsUpdated: updated, rowsAppended: rowsAppended });

  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally {
    try {
      if (doc && !wasOpen) doc.close({ saving: "no" });
    } catch (e2) {}
  }
}
