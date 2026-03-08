# apple-numbers-skill

AI agent skill for reading and editing Apple Numbers documents on macOS.

## Installation

```bash
npx skills add vinitu/apple-numbers-skill
```

## What it does

This skill gives AI agents (Claude Code, Cursor, Copilot, etc.) the ability to:

- **Read** data from `.numbers` files (all sheets/tables or specific ones)
- **Write** data to existing `.numbers` files (single cell or batch)
- **Create** new `.numbers` spreadsheets with predefined structure
- **List** the structure of a spreadsheet (sheets, tables, dimensions)

## How it works

Uses JXA (JavaScript for Automation) via `osascript` to interact with Apple Numbers through its scripting bridge. No external dependencies required — works with any macOS system that has Numbers installed.

## Requirements

- macOS
- Apple Numbers (included with macOS)

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/read-numbers.js` | Read cell data from a spreadsheet |
| `scripts/write-numbers.js` | Write/update cells in a spreadsheet |
| `scripts/create-numbers.js` | Create a new spreadsheet |
| `scripts/list-structure.js` | List sheets, tables, and dimensions |
| `scripts/read-numbers.sh` | Shell wrapper for reading |

## License

MIT
