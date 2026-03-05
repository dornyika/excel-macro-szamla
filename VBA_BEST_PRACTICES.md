# VBA Macro Best Practices
Extracted from `szamla-v106.vb`. Apply these rules to all new macros in this project.

---

## Numeric ID / Scientific Notation Prevention

- Set `NumberFormat = "@"` (Text format) **BEFORE** writing any value to columns containing TAJ numbers, TSZ codes, or other long numeric IDs — never after.
- For sheets used as lookup sources (e.g. `vissz (2)`), apply Text format to **all cells**: `ws.Cells.NumberFormat = "@"` — a lookup can fail silently if the format of the source and target differ.

---

## Copy / Paste Rules

- Never paste formulas — always copy values only.
- For whole rows between sheets: `.Copy Destination:=` (one operation, no clipboard state issues).
- For individual cells: direct assignment `target.Value = source.Value`.
- Never use `PasteSpecial xlPasteFormulas`.

---

## Hungarian Locale (comma vs. point, accented names)

- Never let Excel interpret date or number strings freely. Hungarian locale replaces `.` with `,` in numeric values (e.g. `2025.04` becomes `2025,04`). Protect such columns with Text format upfront.
- Sheet names with accented characters (`"utólagos"`, `"regi_tsz_utolsó4honap"`) must match **exactly** — correct case and accents — or `Worksheets()` returns Nothing silently.

---

## Sheet & Workbook References

- Always use `ActiveWorkbook.Worksheets("name")` — never `Sheets("name")` or `ActiveSheet`.
- Always test `If ws Is Nothing` immediately after assigning a sheet reference, and exit with an optional message if missing.

---

## Column and Row Deletion Order

- Delete columns **right-to-left** (highest index first) — indices shift after each deletion.
- Delete rows **bottom-to-top** (`For i = lastRow To 2 Step -1`) — same reason.
- **Recalculate `lastCol` / `lastRow`** after any deletions or insertions before using them again.

```vb
' Correct column deletion (higher index first)
If intkodCol > megyeCol Then
    ws.Columns(intkodCol).Delete
    ws.Columns(megyeCol).Delete
Else
    ws.Columns(megyeCol).Delete
    ws.Columns(intkodCol).Delete
End If
```

---

## Find / Lookup

- Always reset `Set foundCell = Nothing` before calling `.Find()`.
- Always use `LookIn:=xlValues, LookAt:=xlWhole` — partial or formula-based matches cause silent wrong results.
- Always check `IsError(resultValue)` after reading a value from a found cell.

```vb
Set foundCell = Nothing
On Error Resume Next
Set foundCell = wsSource.Columns(1).Find(What:=lookupValue, LookIn:=xlValues, LookAt:=xlWhole)
On Error GoTo 0

If foundCell Is Nothing Then
    ws.Cells(i, col).Value = "HIÁNYZIK"
Else
    resultValue = wsSource.Cells(foundCell.Row, targetCol).Value
    If IsError(resultValue) Then
        ws.Cells(i, col).Value = "HIÁNYZIK"
    Else
        ws.Cells(i, col).Value = resultValue
    End If
End If
```

---

## AutoFilter

- Always call `ws.AutoFilterMode = False` before applying a new AutoFilter.
- After filtering, check visible rows with `If Not ws.Rows(i).Hidden` — never assume all rows are visible.

---

## Performance Optimization & Batch Save

- Wrap all processing with `OptimizePerformance(True)` at start and `OptimizePerformance(False)` at end.
- In batch mode: **restore all Application settings BEFORE `SaveAs`**, then re-enable optimizations for the next file. Saving while optimizations are active can corrupt the saved file.
- Call `DoEvents` after restoring settings and before saving.
- Always preserve original file format: `wb.SaveAs path, FileFormat:=wb.FileFormat`.
- Skip Excel temp files: `If Left(fileName, 1) <> "~"`.

```vb
' Correct batch save sequence
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayAlerts = True
DoEvents

wb.Activate
wb.SaveAs modifiedFolder & fileName, FileFormat:=wb.FileFormat
wb.Close SaveChanges:=False

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayAlerts = False
```

---

## Silent Parameter Pattern

- Every individual macro has `Optional silent As Boolean = False`.
- When called from `RunAll` or a batch loop, always pass `True` to suppress MsgBox popups.
- Only the top-level orchestrator (`RunAll`, batch summary) shows a final MsgBox.

```vb
Sub MyMacro(Optional silent As Boolean = False)
    ' ...
    If Not silent Then MsgBox "Done.", vbInformation
End Sub
```

---

## Dynamic Range Finding

```vb
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
```

- Re-evaluate both after any structural change (column/row insert or delete).

---

## Sorting

- Always call `.SortFields.Clear` before adding new sort keys.
- Always set `.Header = xlYes`.
- For multi-level sorts, add `SortFields` in priority order (primary first).
