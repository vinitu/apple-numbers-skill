---
name: apple-numbers
description: Read and edit Apple Numbers spreadsheets on macOS. Requires macOS with Apple Numbers installed. Use when the user asks to view, modify, create, or analyze data in .numbers files. Supports reading cell values, writing data, creating new spreadsheets, and working with multiple sheets/tables.
---

# Apple Numbers

A skill for AI agents to read and edit Apple Numbers documents on macOS using JXA (JavaScript for Automation) via `osascript`.

This skill requires macOS and Apple Numbers.app.

## Prerequisites

- macOS with Apple Numbers installed
- `osascript` available (built-in on macOS)

## Password-Protected Files

Password-protected `.numbers` files are not explicitly supported by this skill today.

The current scripts do not accept a password argument and do not pass one when opening a document in Numbers.

For locked spreadsheets, Numbers may show a password prompt or fail to continue in unattended automation.

This skill is most reliable with:
- unprotected spreadsheets;
- spreadsheets that are already open and unlocked in Numbers;
- spreadsheets that macOS can already unlock through saved credentials.

## Reading a Spreadsheet

To read data from a Numbers file, use the `read-numbers.sh` helper script:

```bash
./scripts/read-numbers.sh "/path/to/file.numbers" [sheet_name] [table_name]
```

This outputs JSON with cell values, sheet names, and table structure.

### Reading via osascript directly

```bash
osascript -l JavaScript scripts/read-numbers.js "/path/to/file.numbers"
```

## Writing to a Spreadsheet

To write data to a Numbers file:

```bash
osascript -l JavaScript scripts/write-numbers.js "/path/to/file.numbers" '{"sheet":"Sheet 1","table":"Table 1","row":1,"col":1,"value":"Hello"}'
```

### Batch write (multiple cells):

```bash
osascript -l JavaScript scripts/write-numbers.js "/path/to/file.numbers" '[{"sheet":"Sheet 1","table":"Table 1","row":1,"col":1,"value":"A"},{"row":1,"col":2,"value":"B"}]'
```

### Append one row:

```bash
osascript -l JavaScript scripts/write-numbers.js "/path/to/file.numbers" '{"sheet":"Stocks","table":"Table 1","appendRow":["NVDA","NVIDIA Corporation"]}'
```

### Append multiple rows:

```bash
osascript -l JavaScript scripts/write-numbers.js "/path/to/file.numbers" '{"sheet":"Stocks","table":"Table 1","appendRows":[["AVGO","Broadcom Inc."],["TSM","Taiwan Semiconductor Manufacturing"]]}'
```

Append behavior:
- appended values start from the first column of the new row;
- shorter rows leave the remaining cells blank;
- rows wider than the table return an error.

## Creating a New Spreadsheet

```bash
osascript -l JavaScript scripts/create-numbers.js "/path/to/new-file.numbers" '{"sheets":[{"name":"Data","tables":[{"name":"Table 1","headers":["Name","Value","Date"],"rows":[["Item 1",100,"2026-01-01"]]}]}]}'
```

## Listing Sheets and Tables

```bash
osascript -l JavaScript scripts/list-structure.js "/path/to/file.numbers"
```

Returns JSON with all sheet names, table names, row/column counts.

## Guidelines

- Always close the Numbers document after operations to avoid file locks
- Use JSON output format for structured data exchange with the agent
- Handle errors gracefully — files may be open in Numbers.app already
- Treat password-protected files as a limited case unless the document is already unlocked
- For large spreadsheets, read specific ranges rather than entire sheets
- Respect the user's file paths — never modify files without explicit request
