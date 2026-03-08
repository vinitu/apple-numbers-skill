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
// row/col are 0-based. If sheet/table omitted, uses first sheet/table.

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

    let updated = 0;
    let lastSheet = null;
    let lastTable = null;

    for (const op of operations) {
      const sheetName = op.sheet || null;
      const tableName = op.table || null;
      const row = op.row;
      const col = op.col;
      const value = op.value;

      if (row === undefined || col === undefined || value === undefined) {
        continue;
      }

      let sheet;
      if (sheetName) {
        const sheets = doc.sheets();
        sheet = sheets.find(s => s.name() === sheetName);
        if (!sheet) {
          return JSON.stringify({ error: "Sheet not found: " + sheetName });
        }
      } else {
        sheet = doc.sheets[0];
      }

      let table;
      if (tableName) {
        const tables = sheet.tables();
        table = tables.find(t => t.name() === tableName);
        if (!table) {
          return JSON.stringify({ error: "Table not found: " + tableName });
        }
      } else {
        table = sheet.tables[0];
      }

      const colCount = table.columnCount();
      const cellIndex = row * colCount + col;
      table.cells[cellIndex].value = value;
      updated++;
    }

    doc.save();

    if (!wasOpen) {
      doc.close();
    }

    return JSON.stringify({ success: true, cellsUpdated: updated });

  } catch (e) {
    try {
      if (doc && !wasOpen) doc.close({ saving: "no" });
    } catch (e2) {}
    return JSON.stringify({ error: e.message });
  }
}
