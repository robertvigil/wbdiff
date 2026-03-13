# WBDiff

Compare two Excel workbooks cell-by-cell and generate a detailed differences report. Runs as a VBA macro inside Excel.

## What It Does

Opens two workbooks side by side and produces a "Differences Report" worksheet showing:

- **Cell-level differences** — worksheet name, cell address, primary value, compare value
- **Missing worksheets** — sheets present in one workbook but not the other (checks both directions)
- **Numeric threshold** — configurable minimum difference for numeric values (default: 1)
- **Per-worksheet breakdown** — difference counts grouped by sheet
- **Summary stats** — total differences and generation timestamp

Includes companion macros to highlight and unhighlight differing cells directly in the workbook.

## Usage

1. Open the two workbooks you want to compare
2. Open the VBA editor (`Alt+F11`), insert a new module, and paste `wbdiff.bas`
3. Make the "primary" workbook active (the one you want the report added to)
4. Run `WBDiff` from the macro menu (`Alt+F8`)
5. Review the "Differences Report" worksheet

### Optional: Highlight Differences

- `HighlightFromReport` — applies pink highlighting to every cell listed in the report
- `UnHighlightFromReport` — removes the highlighting

## Features

- **Array-based comparison** — reads entire sheet ranges into memory for 10-100x speed over cell-by-cell COM calls
- **Batched report writing** — buffers all differences and writes in one operation
- **Excel error handling** — correctly compares cells containing `#N/A`, `#REF!`, `#DIV/0!`, etc.
- **No row/column limits** — processes the entire used range per worksheet
- **Progress feedback** — real-time status bar updates during comparison
- **Configurable sensitivity** — adjust `minDifference` constant for numeric threshold

## Configuration

At the top of `WBDiff`:

```vba
Const minDifference As Double = 1  ' reports differences of 1 or greater
```

Set to `0.1` for higher sensitivity, `10` for lower, etc. Non-numeric differences (text, formulas) are always reported.

## Requirements

- Microsoft Excel with VBA support (desktop)
- No external dependencies

## Built With

Vibe-coded with [Claude Code](https://claude.ai/claude-code) (Anthropic)
