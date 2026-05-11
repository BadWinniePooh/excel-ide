# Excel Formula Beautifier

_Started 2026-05-11 08:19 UTC_

---

## User

I would like to create a application. The app must run client side only. There is no data persistence. Anyways it should be hosted as a web app centrally at a server. So each user should have their own state of the webapp without modifying other users states. The main purpose of the app is to beautify excel formulas. The first and main use case is, that a user adds a oneline formula directly from excel and the application will make it humanly readable by introducing new lines and indentations (pretty much like an IDE for a developer).

Additionally the user must be able to modify the pretty excel formula and in real time the one line formula should be updated. So that it can be copied back to excel.

## Assistant

_[tool: copy_starter_component]_

I'll build a client-side Excel formula prettifier with bidirectional editing. Let me commit to the design direction first:

**Aesthetic**: Warm paper IDE — IBM Plex Sans for UI, JetBrains Mono for code. Muted syntax-highlight palette (blue functions, green strings, amber numbers, purple references). Two-pane layout with the one-liner input on top and a syntax-highlighted multi-line editor below.

**Approach**:
- Custom Excel formula tokenizer (handles `""` escapes, `'sheet names'!ref`, `Table1[Column]` structured refs, `{array;literals}`, error tokens `#REF!`, etc.)
- Recursive group parser → pretty formatter with "inline if short" measuring
- Minifier (whitespace stripper that preserves string contents) for pretty→oneline sync
- Live two-way sync: edit oneline → re-format pretty; edit pretty → minify to oneline in real time

Let me set up the files.

_[tool: write_file]_

_[tool: set_project_title]_

_[tool: write_file]_

_[tool: read_file]_

_[tool: grep]_

_[tool: grep]_

_[tool: read_file]_

_[tool: grep]_

_[tool: grep]_

_[tool: snip]_

_[tool: read_file]_

_[tool: str_replace_edit]_

_[tool: write_file]_

_[tool: snip]_

_[tool: done]_

_[tool: fork_verifier_agent]_

Built **Formula Prettifier** — a client-side Excel/Sheets formula beautifier:

- **Top pane**: one-line input. Paste straight from Excel.
- **Bottom pane**: pretty-printed multi-line editor with line numbers + syntax highlighting (functions, strings, numbers, refs, structured `Table[Column]` refs, sheet-quoted names, errors like `#N/A`, rainbow-matched parens).
- **Live two-way sync**: edit either side and the other updates instantly. The one-line side is rebuilt by stripping whitespace (preserving string literals), so your edits to the pretty side flow back losslessly.
- **Smart formatter**: collapses short calls inline, expands long ones, special-cases `LET` (name/value pairs on one line, body on its own line).
- **Tweaks** (toolbar toggle): indent size, wrap column, uppercase function names, comma vs. semicolon separator.
- **Status bar**: balanced-paren check, function count, max depth, char count.

Everything runs in-browser, no persistence — each tab is its own session.

## User

Verifier agent check completed

## Assistant

_[tool: str_replace_edit]_

_[tool: str_replace_edit]_

_[tool: snip]_

_[tool: done]_

Fixed: scoped `min-height: 0` away from `.pane-one` and let the page scroll instead of forcing fixed viewport height, so the one-line pane is no longer squashed at short viewports.

## User

would you mind adding a settings pane, that let's me choose my Excel language and edit some more settings?
For now I would like to be able to choose between german and english (default german) excel. Additionally the Settings pane should cover Indent 1 to 8 (default 2), Wrap column 20 to 140 (default 60), Uppercase function names as toggle (default on), Separator (default Semicolon).

Additionally as a new feature it would be amazing if I can translate the functions between available languages, so for now between german and english.

## Assistant

_[tool: snip]_

_[tool: write_file]_

_[tool: write_file]_

_[tool: snip]_

_[tool: str_replace_edit]_

_[tool: snip]_

_[tool: done]_

Added a dedicated **Settings** drawer (gear in the top-right) covering:

- **Excel language** — German (default) / English, with auto-translation of all function names and decimal/separator conventions when you switch
- **Indent** 1–8 (default 2)
- **Wrap column** 20–140 (default 60)
- **Uppercase function names** toggle (default on)
- **Argument separator** — Comma `,` or Semicolon `;` (default `;` to match German Excel)

The translation feature is wired in two ways: changing the language in Settings translates the current formula in place, and there's a quick flag‑pair button (🇩🇪 → 🇬🇧) in the header for one-click swaps. The dictionary covers ~220 functions across logical, lookup, math, statistical, text, date, info, financial, and dynamic-array categories, plus `WAHR`/`FALSCH` ↔ `TRUE`/`FALSE` and decimal/separator conversion (`1,5;A1` ↔ `1.5,A1`).

## User

The Uppercase function names toggle has no function. Can you please implement this function?

## Assistant

_[tool: snip]_

_[tool: str_replace_edit]_

_[tool: snip]_

_[tool: done]_

Fixed. The uppercase toggle (and the other format settings) now reformats the pretty output immediately, even after you've edited the pretty pane.

## User

The argument separator switches only some separators and not all. Can you please verify that it correctly switches out all separators and not just some?

## Assistant

_[tool: str_replace_edit]_

_[tool: snip]_

_[tool: done]_

Fixed. Function-call and paren argument separators now always use the configured `options.sep`, so switching between `,` and `;` swaps every separator (array literals `{1,2;3,4}` keep their `,`/`;` since those mean column/row, not arg).

## User

Can you please add IDE functionality to the "pretty form"? I would like to have functionality like checking if the function syntax is correct, if my cursor is touching a bracket, the other part of the bracket should be highlighted. If the cursor is in a function than on the right side a permanent tool tip should appear that shows which part of the function the user is currently editing and which parts are missing.

## Assistant

_[tool: snip]_

_[tool: str_replace_edit]_

_[tool: grep]_

_[tool: read_file]_

_[tool: str_replace_edit]_

_[tool: grep]_

_[tool: snip]_

_[tool: read_file]_

_[tool: snip]_

_[tool: read_file]_

_[tool: grep]_

_[tool: snip]_

_[tool: list_files]_

_[tool: grep]_

_[tool: snip]_

_[tool: str_replace_edit]_

_[tool: str_replace_edit]_

_[tool: snip]_

_[tool: str_replace_edit]_

_[tool: str_replace_edit]_

_[tool: grep]_

_[tool: snip]_

_[tool: str_replace_edit]_

_[tool: snip]_

_[tool: str_replace_edit]_

_[tool: read_file]_

_[tool: snip]_

_[tool: str_replace_edit]_

_[tool: snip]_

_[tool: done]_

_[tool: fork_verifier_agent]_

Added full IDE features to the pretty pane:

**Bracket matching** — touching any `(`, `)`, `{`, `}` highlights the cursor bracket in solid blue and its partner in dashed blue. Works for nested rainbow-coloured parens and braces inside array literals.

**Function context panel** (right side, ~320px):
- **Function name** of the call your cursor is inside, with EN/DE alternates if known
- **Parameter list** — current argument marked with ▸ and a soft amber background; `required` vs `optional` labelled inline; arguments past the end of the signature show as overflow
- **Progress line** — "Argument 3 of 4 · 1 required still missing"
- **Diagnostics** — counts at the top, listed below with severity colour: missing required args, too many args, empty argument slots (`SUM(,,)`), stray closing `)`, unclosed `(`/`{`. Clean state shows "✓ Syntax looks clean"
- **Tips** section at the bottom

**Inline error squiggles** — error diagnostics underline the offending range in red wavy lines inside the pretty editor.

The panel collapses below the editor on screens < 980px. Cursor updates fire on click/key/select/focus, so the active argument tracks in real time.

## User

There is a bug, when I copy my function into there, the minify function strips also spaces from variable names. So for example if I provide "[@[End Date]]"  it will minify to "[@[EndDate]]" which effectively breaks my function. That is not acceptable, please fix this.

## Assistant

_[tool: snip]_

_[tool: grep]_

