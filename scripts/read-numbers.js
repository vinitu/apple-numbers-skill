#!/usr/bin/env osascript -l JavaScript

// Read data from an Apple Numbers spreadsheet
// Usage: osascript -l JavaScript read-numbers.js "/path/to/file.numbers" [sheetName] [tableName]

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
    // Check if document is already open
    const openDocs = Numbers.documents();
    for (let i = 0; i < openDocs.length; i++) {
      try {
        if (openDocs[i].file().toString().includes(filePath.replace(/.*\//, ''))) {
          doc = openDocs[i];
          wasOpen = true;
          break;
        }
      } catch (e) {}
    }

    if (!doc) {
      Numbers.open(Path(filePath));
      doc = Numbers.documents[0];
    }

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
        for (let c = 0; c < colCount; c++) {
          try {
            headers.push(table.cells[c].formattedValue() || "");
          } catch (e) {
            headers.push("");
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

    if (!wasOpen) {
      doc.close({ saving: "no" });
    }

    return JSON.stringify(result, null, 2);

  } catch (e) {
    try {
      if (doc && !wasOpen) doc.close({ saving: "no" });
    } catch (e2) {}
    return JSON.stringify({ error: e.message });
  }
}
