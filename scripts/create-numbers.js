#!/usr/bin/env osascript -l JavaScript

// Create a new Apple Numbers spreadsheet
// Usage: osascript -l JavaScript create-numbers.js "/path/to/new-file.numbers" '<json>'
//
// JSON format:
// {
//   "sheets": [
//     {
//       "name": "Sheet Name",
//       "tables": [
//         {
//           "name": "Table 1",
//           "headers": ["Col A", "Col B", "Col C"],
//           "rows": [
//             ["val1", "val2", "val3"],
//             ["val4", "val5", "val6"]
//           ]
//         }
//       ]
//     }
//   ]
// }

function tableColumnCount(headers, rows) {
  let totalCols = Array.isArray(headers) ? headers.length : 0;

  if (Array.isArray(rows)) {
    for (let i = 0; i < rows.length; i++) {
      if (Array.isArray(rows[i]) && rows[i].length > totalCols) {
        totalCols = rows[i].length;
      }
    }
  }

  return Math.max(totalCols, 1);
}

function run(argv) {
  const filePath = argv[0];
  const jsonStr = argv[1];

  if (!filePath) {
    return JSON.stringify({ error: "Usage: create-numbers.js <file-path> [json]" });
  }

  let spec = null;
  if (jsonStr) {
    try {
      spec = JSON.parse(jsonStr);
    } catch (e) {
      return JSON.stringify({ error: "Invalid JSON: " + e.message });
    }
  }

  const Numbers = Application("Numbers");
  const app = Application.currentApplication();
  app.includeStandardAdditions = true;
  let doc;

  try {
    // Create new document
    doc = Numbers.Document().make();

    if (spec) {
      if (Array.isArray(spec.sheets) && spec.sheets.length > 0) {
        for (let s = 0; s < spec.sheets.length; s++) {
          const sheetSpec = spec.sheets[s];

          let sheet;
          if (s === 0) {
            // Use the default first sheet
            sheet = doc.sheets[0];
          } else {
            // Add new sheet
            const newSheet = Numbers.Sheet({ name: sheetSpec.name });
            doc.sheets.push(newSheet);
            sheet = doc.sheets[doc.sheets.length - 1];
          }

          if (sheetSpec.name) {
            sheet.name = sheetSpec.name;
          }

          if (Array.isArray(sheetSpec.tables)) {
            for (let t = 0; t < sheetSpec.tables.length; t++) {
              const tableSpec = sheetSpec.tables[t];
              let table;

              if (t === 0) {
                table = sheet.tables[0];
              } else {
                const newTable = Numbers.Table({ name: tableSpec.name });
                sheet.tables.push(newTable);
                table = sheet.tables[sheet.tables.length - 1];
              }

              if (tableSpec.name) {
                table.name = tableSpec.name;
              }

              const headers = Array.isArray(tableSpec.headers) ? tableSpec.headers : [];
              const rows = Array.isArray(tableSpec.rows) ? tableSpec.rows : [];
              const hasHeaders = headers.length > 0;
              const totalRows = Math.max((hasHeaders ? 1 : 0) + rows.length, 1);
              const totalCols = tableColumnCount(headers, rows);
              const maxHeaderColumns = Math.max(totalCols - 1, 0);

              // Numbers cannot keep one header column on a one-column table.
              if (table.headerColumnCount() > maxHeaderColumns) {
                table.headerColumnCount = maxHeaderColumns;
              }

              table.headerRowCount = hasHeaders ? 1 : 0;
              table.columnCount = totalCols;
              table.rowCount = totalRows;

              // Write headers
              if (hasHeaders) {
                for (let c = 0; c < headers.length; c++) {
                  table.cells[c].value = headers[c];
                }
              }

              // Write data rows
              const dataStartRow = hasHeaders ? 1 : 0;
              for (let r = 0; r < rows.length; r++) {
                if (!Array.isArray(rows[r])) continue;
                for (let c = 0; c < rows[r].length; c++) {
                  table.cells[(r + dataStartRow) * totalCols + c].value = rows[r][c];
                }
              }
            }
          }
        }
      }
    }

    // Save to specified path
    const saveFile = Path(filePath);
    doc.save({ in: saveFile });

    return JSON.stringify({ success: true, file: filePath });

  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally {
    try {
      if (doc) doc.close({ saving: "no" });
    } catch (e2) {}
  }
}
