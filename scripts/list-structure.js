#!/usr/bin/env osascript -l JavaScript

// List structure of an Apple Numbers spreadsheet (sheets, tables, dimensions)
// Usage: osascript -l JavaScript list-structure.js "/path/to/file.numbers"

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
    return JSON.stringify({ error: "Usage: list-structure.js <file-path>" });
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

    return JSON.stringify(result, null, 2);

  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally {
    try {
      if (doc && !wasOpen) doc.close({ saving: "no" });
    } catch (e2) {}
  }
}
