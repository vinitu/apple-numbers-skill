.PHONY: dictionary dictionary-numbers dictionary-standard compile test test-dictionary test-smoke

dictionary:
	@printf '### Numbers.app\n'
	@sdef /Applications/Numbers.app
	@printf '\n### CocoaStandard.sdef\n'
	@cat /System/Library/ScriptingDefinitions/CocoaStandard.sdef

dictionary-numbers:
	@sdef /Applications/Numbers.app

dictionary-standard:
	@cat /System/Library/ScriptingDefinitions/CocoaStandard.sdef

compile:
	@set -euo pipefail; \
	find scripts -name '*.applescript' -print | while IFS= read -r file; do \
		osacompile -o /tmp/$$(echo "$$file" | tr '/' '_' | sed 's/\.applescript$$/.scpt/') "$$file"; \
	done

test: test-dictionary test-smoke

test-dictionary:
	@bash tests/dictionary_contract.sh

test-smoke:
	@bash tests/smoke_numbers.sh
