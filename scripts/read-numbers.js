#!/usr/bin/env osascript -l JavaScript

// Read data from an Apple Numbers spreadsheet
// Usage: osascript -l JavaScript read-numbers.js "/path/to/file.numbers" [sheetName] [tableName]

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

function run(argv) {
  const filePath = argv[0];
  if (!filePath) {
    return JSON.stringify({ error: "Usage: read-numbers.js <file-path> [sheet] [table]" });
  }

  const sheetFilter = argv[1] || null;
  const tableFilter = argv[2] || null;

  const Numbers = Application("Numbers");
  const app = Application.currentApplication();
  app.includeStandardAdditions = true;

  let doc;
  let wasOpen = false;

  try {
    const opened = openDocument(Numbers, filePath);
    doc = opened.doc;
    wasOpen = opened.wasOpen;

    const result = { file: filePath, sheets: [] };

    const sheets = doc.sheets();
    for (let s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      const sheetName = sheet.name();

      if (sheetFilter && sheetName !== sheetFilter) continue;

      const sheetData = { name: sheetName, tables: [] };

      const tables = sheet.tables();
      for (let t = 0; t < tables.length; t++) {
        const table = tables[t];
        const tableName = table.name();

        if (tableFilter && tableName !== tableFilter) continue;

        const rowCount = table.rowCount();
        const colCount = table.columnCount();
        const headerCount = table.headerRowCount();

        const headers = [];
        if (headerCount > 0) {
          for (let c = 0; c < colCount; c++) {
            try {
              headers.push(table.cells[c].formattedValue() || "");
            } catch (e) {
              headers.push("");
            }
          }
        }

        const rows = [];
        for (let r = headerCount; r < rowCount; r++) {
          const row = [];
          for (let c = 0; c < colCount; c++) {
            try {
              const cell = table.cells[r * colCount + c];
              row.push(cell.formattedValue() || "");
            } catch (e) {
              row.push("");
            }
          }
          // Skip entirely empty rows
          if (row.some(v => v !== "")) {
            rows.push(row);
          }
        }

        sheetData.tables.push({
          name: tableName,
          rowCount: rowCount,
          columnCount: colCount,
          headerRowCount: headerCount,
          headers: headers,
          rows: rows
        });
      }

      result.sheets.push(sheetData);
    }

    return JSON.stringify(result, null, 2);

  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally {
    try {
      if (doc && !wasOpen) doc.close({ saving: "no" });
    } catch (e2) {}
  }
}
