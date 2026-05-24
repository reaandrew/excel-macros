# excel-macros

A small collection of Excel VBA macros I keep around so I don't have to write them from scratch each time.

## Macros

### `DistributeBetween` — `DistributeBetween.bas`

Linearly interpolates the cells between the top and bottom of a single-column selection.

**Usage**

1. In Excel, put a value in the top cell and a value in the bottom cell of a contiguous column range (e.g. `A2` = 100, `A10` = 180).
2. Select the full range, including both anchors (e.g. `A2:A10`).
3. Run the macro.

Before / after:

```
Before:        After:
A2  100        A2  100
A3             A3  110
A4             A4  120
A5             A5  130
A6             A6  140
A7             A7  150
A8             A8  160
A9             A9  170
A10 180        A10 180
```

The two anchor cells are preserved. Cells in between are overwritten. Interpolated values are rounded to the larger decimal-place count of the two anchors (so two integer anchors give integer steps; `1.50` and `2.00` give two-decimal steps).

**Validation**

The macro shows a message box and exits if:

- Nothing is selected, or the selection is not a `Range`.
- The selection spans more than one column.
- The selection is non-contiguous (multiple areas).
- Fewer than 3 cells are selected (no in-between cells to fill).
- The top or bottom cell is not numeric.

## Installing a macro

1. Open the workbook you want the macro in (or `PERSONAL.XLSB` for global use).
2. `Alt+F11` to open the VBA editor.
3. `File → Import File…` and pick the `.bas` file.
4. Back in Excel, `Alt+F8` lists the macro so you can run it or assign a shortcut/Quick Access button.
