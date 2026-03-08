# AGENTS.md

## Purpose

This repository provides an AI skill for working with Apple Numbers files on macOS through JXA scripts executed with `osascript`.

Primary goals:
- keep the scripts dependency-free;
- preserve predictable JSON I/O for agents;
- avoid leaving Numbers documents open or unexpectedly modified.

## Repository Layout

- `SKILL.md`: the skill contract and usage instructions for agents.
- `README.md`: public project overview and installation notes.
- `scripts/read-numbers.js`: reads spreadsheet data and returns JSON.
- `scripts/write-numbers.js`: writes one or more cell updates.
- `scripts/create-numbers.js`: creates a new spreadsheet from JSON spec.
- `scripts/list-structure.js`: lists sheets, tables, and dimensions.
- `scripts/read-numbers.sh`: thin shell wrapper around the read script.

## Working Rules

- Keep this repo macOS-first. The scripts depend on Apple Numbers and `osascript`; do not add cross-platform abstractions unless explicitly requested.
- Prefer small, direct JXA changes over adding external runtimes, packages, or build steps.
- Preserve CLI behavior. Existing script entrypoints and argument shapes should remain stable unless the task explicitly requires a breaking change.
- Preserve JSON output as the integration boundary. Success and error responses should stay machine-readable.
- If you change script behavior, update both `SKILL.md` and `README.md` when usage, arguments, or examples change.
- Respect document lifecycle handling. When opening a Numbers document in automation, close it when appropriate and avoid overwriting user state accidentally.

## Script Conventions

- `read-numbers.js` and `list-structure.js` should be read-only operations.
- `write-numbers.js` and `create-numbers.js` must fail clearly on invalid input and should not silently change interface semantics.
- Keep row/column indexing conventions explicit in comments and docs. Current write operations use 0-based `row`/`col`.
- Prefer returning structured errors like `{"error":"..."}` instead of plain text.
- Avoid broad refactors unless they reduce risk around Numbers automation behavior.

## Validation

There is no formal test suite in this repo today. After making changes:
- run the relevant script with `osascript -l JavaScript ...` if the environment has Numbers available;
- at minimum, syntax-check shell wrappers and re-read examples in `SKILL.md` and `README.md` for consistency;
- if you cannot run Numbers-dependent validation, state that clearly.

## Common Pitfalls

- Apple Numbers scripting can behave differently when a document is already open; keep `wasOpen` logic intact unless you are deliberately redesigning it.
- File paths may point to `.numbers` packages; do not assume plain files only.
- JSON passed through shell commands is quoting-sensitive. Keep examples copy-pastable.
- JXA indexing and Numbers table cell access are easy to break with off-by-one mistakes; verify row/column math carefully.
