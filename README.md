# apple-numbers-skill

AI agent skill for reading and editing Apple Numbers documents on macOS.

Requires macOS and Apple Numbers.app.

## Installation

```bash
npx skills add vinitu/apple-numbers-skill
```

## What it does

This skill gives AI agents (Claude Code, Cursor, Copilot, etc.) the ability to:

- **Read** data from `.numbers` files (all sheets/tables or specific ones)
- **Write** data to existing `.numbers` files (single cell, batch, or append rows)
- **Create** new `.numbers` spreadsheets with predefined structure
- **List** the structure of a spreadsheet (sheets, tables, dimensions)

## How it works

Uses JXA (JavaScript for Automation) via `osascript` to interact with Apple Numbers through its scripting bridge. No external dependencies required — works with any macOS system that has Numbers installed.

## Requirements

- macOS
- Apple Numbers (included with macOS)

## Behavior and Limitations

Apple Numbers is a GUI application. These scripts automate Numbers through `osascript` and JXA, but they do not provide a fully headless backend mode.

In normal usage, Numbers does not need to become the frontmost app, but macOS may still launch the application while a file is being processed.

## Permissions

On first run, macOS may ask for Automation permission so your terminal, editor, or AI agent can control Numbers.app.

If permission is denied, script execution may fail or appear to hang until access is granted in System Settings.

## Password-Protected Files

Password-protected `.numbers` files are not explicitly supported by this skill today.

Apple Numbers can protect spreadsheets with a password, but the current scripts do not accept a password argument and do not pass one when opening a document.

Because of that, a locked spreadsheet may trigger a password prompt in Numbers or fail to open in unattended automation.

This skill is most reliable with:
- unprotected spreadsheets;
- spreadsheets that are already open and unlocked in Numbers;
- spreadsheets that macOS can already unlock through saved credentials.

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/read-numbers.js` | Read cell data from a spreadsheet |
| `scripts/write-numbers.js` | Write/update cells in a spreadsheet |
| `scripts/create-numbers.js` | Create a new spreadsheet |
| `scripts/list-structure.js` | List sheets, tables, and dimensions |
| `scripts/read-numbers.sh` | Shell wrapper for reading |

## Write Coordinates

`row` and `col` in write operations are 0-based.

For example, `row: 0, col: 0` targets the top-left cell of the table.

## Append Rows

`write-numbers.js` also supports appending rows to an existing table.

Append one row:

```bash
osascript -l JavaScript scripts/write-numbers.js "/path/to/file.numbers" '{"sheet":"Stocks","table":"Table 1","appendRow":["NVDA","NVIDIA Corporation"]}'
```

Append multiple rows:

```bash
osascript -l JavaScript scripts/write-numbers.js "/path/to/file.numbers" '{"sheet":"Stocks","table":"Table 1","appendRows":[["AVGO","Broadcom Inc."],["TSM","Taiwan Semiconductor Manufacturing"]]}'
```

Notes:
- appended values are written from the first column of the new row;
- rows shorter than the table width leave the remaining cells blank;
- rows wider than the table return an error.

## Document Lifecycle

If a document is already open in Numbers, the scripts try to reuse that open document.

If a read-only script opens the document itself, it closes the document afterwards to avoid leaving extra windows or file locks behind.

## Performance Notes

Reading a full sheet or table can be slow for larger spreadsheets because values are fetched cell-by-cell through the Numbers scripting bridge.

When possible, target a specific sheet and table instead of reading the whole document.

## License

MIT
