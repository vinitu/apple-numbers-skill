#!/bin/bash
# Read data from an Apple Numbers spreadsheet
# Usage: read-numbers.sh "/path/to/file.numbers" [sheet_name] [table_name]

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [ -z "$1" ]; then
  echo '{"error": "Usage: read-numbers.sh <file-path> [sheet] [table]"}'
  exit 1
fi

FILE_PATH="$1"
SHEET="${2:-}"
TABLE="${3:-}"

if [ ! -f "$FILE_PATH" ] && [ ! -d "$FILE_PATH" ]; then
  echo "{\"error\": \"File not found: $FILE_PATH\"}"
  exit 1
fi

osascript -l JavaScript "$SCRIPT_DIR/read-numbers.js" "$FILE_PATH" "$SHEET" "$TABLE"
