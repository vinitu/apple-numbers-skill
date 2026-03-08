---
name: apple-numbers
description: Read and edit Apple Numbers spreadsheets on macOS. Use when the user asks to view, modify, create, or analyze data in .numbers files. Supports reading cell values, writing data, creating new spreadsheets, and working with multiple sheets/tables.
---

# Apple Numbers

A skill for AI agents to read and edit Apple Numbers documents on macOS using JXA (JavaScript for Automation) via `osascript`.

## Prerequisites

- macOS with Apple Numbers installed
- `osascript` available (built-in on macOS)

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
- For large spreadsheets, read specific ranges rather than entire sheets
- Respect the user's file paths — never modify files without explicit request
