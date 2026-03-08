#!/usr/bin/env osascript -l JavaScript

// List structure of an Apple Numbers spreadsheet (sheets, tables, dimensions)
// Usage: osascript -l JavaScript list-structure.js "/path/to/file.numbers"

function run(argv) {
  const filePath = argv[0];
  if (!filePath) {
    return JSON.stringify({ error: "Usage: list-structure.js <file-path>" });
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

    const result = { file: filePath, sheets: [] };

    const sheets = doc.sheets();
    for (let s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      const sheetInfo = { name: sheet.name(), tables: [] };

      const tables = sheet.tables();
      for (let t = 0; t < tables.length; t++) {
        const table = tables[t];
        sheetInfo.tables.push({
          name: table.name(),
          rowCount: table.rowCount(),
          columnCount: table.columnCount(),
          headerRowCount: table.headerRowCount(),
          headerColumnCount: table.headerColumnCount()
        });
      }

      result.sheets.push(sheetInfo);
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
