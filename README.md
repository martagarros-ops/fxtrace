# FxTrace

Excel add-in (Mac, Windows, Excel on the web) for financial-model auditing:

- **Precedents / Dependents**: navigable tree. Click or use the arrow keys to jump to each cell. Expand a node to keep tracing (ranges expand into their cells). Works across sheets, with named ranges and table (structured) references.
- **Formula Explorer**: breaks the selected formula into its parts, shows the value of every reference and, on demand, of every sub-expression.

## Keyboard shortcuts in the grid (no need to touch the panel)

| Mac | Windows | Action |
|---|---|---|
| ⌘⌥← | Ctrl+Alt+← | Go to precedents (press again to go one level deeper) |
| ⌘⌥→ | Ctrl+Alt+→ | Go to dependents (press again to go one level further) |
| ⌘⌥↓ / ⌘⌥↑ | Ctrl+Alt+↓ / ↑ | Next / previous cell in the current list |
| ⌘⌥B | Ctrl+Alt+B | Back one level (repeat to return to the start cell) |
| ⌘⌥F | Ctrl+Alt+F | Show the FxTrace panel |

If a shortcut clashes with an Excel one, Excel asks once which one to use. Pick **FxTrace**.

## Keyboard inside the panel (click inside the panel first)

| Key | Action |
|---|---|
| ↑ / ↓ | Move and jump to the cell in Excel |
| → / ← | Expand / collapse |
| Enter | Trace from this cell (it becomes the new origin) |
| Esc | Back to the origin cell |
| Backspace | Previous traced cell |

## Limitations (v1)

- Cells reached through `OFFSET` / `INDIRECT` are flagged but not listed.
- External-workbook references are shown but can't be opened.
- "Evaluate parts" writes to a temporary hidden sheet, and Excel clears the undo history when that happens.
- Dependents use Excel's own engine where available (Excel API 1.15+). Otherwise FxTrace scans all formulas in the workbook, and that scan doesn't cover structured references.

## Hosting (GitHub Pages)

1. Push this repo to GitHub. In **Settings → Pages**, choose *Deploy from a branch*, branch `main`, folder `/docs`.
2. Build the manifest: `node tools/build-manifest.js https://<user>.github.io/<repo>`

## Install on Mac

1. Quit Excel.
2. In Finder: **Go → Go to Folder…** and paste
   `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
   (if the `wef` folder doesn't exist, create it).
3. Copy `manifest.xml` into that folder.
4. Open Excel and a workbook. Then go to **Home → FxTrace**, or **Insert → Add-ins → My Add-ins → Developer Add-ins → FxTrace**.

Updates to the code on GitHub Pages are picked up the next time the panel opens. You don't need to reinstall.

## Development

- `node test/parser.test.js`: formula parser tests
- `node test/ui.test.js`: headless UI test against a mocked Excel API (Playwright)
