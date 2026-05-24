# excel-macros

A small collection of Excel VBA macros I keep around so I don't have to write them from scratch each time.

## Macros

### `FillGaps` — `FillGaps.bas`

Walks down a column and fills every empty cell by linear interpolation between the nearest non-empty cell above and the nearest non-empty cell below. Handles many gaps in one go.

**Usage**

- Click any cell in the column and run — fills every gap between the first and last non-empty cells of that column.
- Or select a contiguous range in one column to scope it.

```
Before:        After:
A1  100        A1  100
A2             A2  110
A3             A3  120
A4  130        A4  130
A5             A5  145
A6  160        A6  160
A7             A7  170
A8             A8  180
A9             A9  190
A10 200        A10 200
```

Notes: non-numeric anchors (e.g. a column header) are skipped, so headers don't break it. Each gap's decimal places match the larger of its two anchors.

## Installing a macro

1. Open the workbook you want the macro in (or `PERSONAL.XLSB` for global use).
2. `Alt+F11` to open the VBA editor.
3. `File → Import File…` and pick the `.bas` file.
4. Back in Excel, `Alt+F8` lists the macro so you can run it or assign a shortcut/Quick Access button.
