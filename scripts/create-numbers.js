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

function run(argv) {
  const filePath = argv[0];
  const jsonStr = argv[1];

  if (!filePath) {
    return JSON.stringify({ error: "Usage: create-numbers.js <file-path> [json]" });
  }

  const Numbers = Application("Numbers");
  const app = Application.currentApplication();
  app.includeStandardAdditions = true;

  try {
    // Create new document
    const doc = Numbers.Document().make();

    if (jsonStr) {
      const spec = JSON.parse(jsonStr);

      if (spec.sheets && spec.sheets.length > 0) {
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

          if (sheetSpec.tables) {
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

              const headers = tableSpec.headers || [];
              const rows = tableSpec.rows || [];
              const totalRows = 1 + rows.length; // header + data
              const totalCols = headers.length || (rows[0] ? rows[0].length : 1);

              // Resize table
              table.rowCount = totalRows;
              table.columnCount = totalCols;

              // Write headers
              for (let c = 0; c < headers.length; c++) {
                table.cells[c].value = headers[c];
              }

              // Write data rows
              for (let r = 0; r < rows.length; r++) {
                for (let c = 0; c < rows[r].length; c++) {
                  table.cells[(r + 1) * totalCols + c].value = rows[r][c];
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
    doc.close();

    return JSON.stringify({ success: true, file: filePath });

  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
