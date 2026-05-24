Attribute VB_Name = "FillGaps"
Option Explicit

' For every empty cell in a column, fill it by linear interpolation between
' the nearest non-empty cell above and the nearest non-empty cell below.
' Decimal places per gap match the larger of the two anchors.
'
' Usage:
'   - Click any cell in the column, then run: fills all gaps between the
'     first and last non-empty cells of that column.
'   - Or select a contiguous range in a single column to limit the scope.
'
' Non-numeric anchors (e.g. a header) are skipped, so headers above the
' first numeric cell are safe.
Public Sub FillGaps()
    Dim col As Range
    Set col = ResolveScope()
    If col Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    Dim n As Long
    n = col.Rows.Count

    Dim i As Long, j As Long, k As Long
    i = 1
    Do While i < n
        If Not IsBlankCell(col.Cells(i, 1)) Then
            j = i + 1
            Do While j <= n
                If Not IsBlankCell(col.Cells(j, 1)) Then Exit Do
                j = j + 1
            Loop

            If j <= n And j > i + 1 Then
                If IsNumeric(col.Cells(i, 1).Value) And IsNumeric(col.Cells(j, 1).Value) Then
                    Dim topV As Double, botV As Double, stepSize As Double, dp As Long
                    topV = CDbl(col.Cells(i, 1).Value)
                    botV = CDbl(col.Cells(j, 1).Value)
                    stepSize = (botV - topV) / (j - i)
                    dp = MaxDecimalPlaces(col.Cells(i, 1), col.Cells(j, 1))
                    For k = i + 1 To j - 1
                        col.Cells(k, 1).Value = Application.WorksheetFunction.Round( _
                            topV + stepSize * (k - i), dp)
                    Next k
                End If
            End If

            i = j
        Else
            i = i + 1
        End If
    Loop

    Application.ScreenUpdating = True
End Sub

Private Function ResolveScope() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Click in a column first.", vbExclamation, "Fill Gaps"
        Exit Function
    End If

    Dim sel As Range
    Set sel = Selection

    If sel.Areas.Count > 1 Then
        MsgBox "Select a single contiguous range, or click one cell.", _
               vbExclamation, "Fill Gaps"
        Exit Function
    End If

    If sel.Columns.Count > 1 Then
        MsgBox "Selection must be in a single column.", _
               vbExclamation, "Fill Gaps"
        Exit Function
    End If

    If sel.Cells.Count = 1 Then
        Set ResolveScope = ColumnDataRange(sel)
        If ResolveScope Is Nothing Then
            MsgBox "No data found in this column.", vbExclamation, "Fill Gaps"
        End If
        Exit Function
    End If

    Set ResolveScope = sel
End Function

Private Function ColumnDataRange(c As Range) As Range
    Dim ws As Worksheet
    Set ws = c.Worksheet
    Dim colNum As Long
    colNum = c.Column

    Dim used As Range
    Set used = ws.UsedRange
    Dim usedTop As Long, usedBot As Long
    usedTop = used.Row
    usedBot = used.Row + used.Rows.Count - 1

    Dim firstRow As Long, lastRow As Long, r As Long
    firstRow = 0
    For r = usedTop To usedBot
        If Not IsBlankCell(ws.Cells(r, colNum)) Then
            firstRow = r
            Exit For
        End If
    Next r
    If firstRow = 0 Then Exit Function

    lastRow = firstRow
    For r = usedBot To firstRow Step -1
        If Not IsBlankCell(ws.Cells(r, colNum)) Then
            lastRow = r
            Exit For
        End If
    Next r

    If lastRow > firstRow Then
        Set ColumnDataRange = ws.Range(ws.Cells(firstRow, colNum), ws.Cells(lastRow, colNum))
    End If
End Function

Private Function IsBlankCell(c As Range) As Boolean
    Dim v As Variant
    v = c.Value
    IsBlankCell = IsEmpty(v) Or (VarType(v) = vbString And Len(v) = 0)
End Function

Private Function MaxDecimalPlaces(a As Range, b As Range) As Long
    Dim da As Long, db As Long
    da = DecimalPlaces(a)
    db = DecimalPlaces(b)
    If da > db Then
        MaxDecimalPlaces = da
    Else
        MaxDecimalPlaces = db
    End If
End Function

Private Function DecimalPlaces(c As Range) As Long
    Dim s As String
    s = Trim$(Str(c.Value2))
    Dim p As Long
    p = InStr(s, ".")
    If p = 0 Then
        DecimalPlaces = 0
    Else
        DecimalPlaces = Len(s) - p
    End If
End Function
