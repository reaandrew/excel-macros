Attribute VB_Name = "FillGaps"
Option Explicit

' For every empty cell in a column, fill it by linear interpolation between
' the nearest non-empty cell above and the nearest non-empty cell below.
' Decimal places per gap match the larger of the two anchors.
'
' Usage:
'   - Click any cell in the column and run: fills all gaps between the
'     first and last non-empty cells of that column.
'   - Or select a contiguous range in a single column to limit the scope.
'
' The column is read in one go as a Variant array and written back as one
' assignment, so this stays fast even on long columns or workbooks where
' UsedRange is unexpectedly large.
Public Sub FillGaps()
    Dim col As Range
    Set col = ResolveScope()
    If col Is Nothing Then Exit Sub

    Dim n As Long
    n = col.Rows.Count
    If n < 2 Then Exit Sub

    Dim prevCalc As Long
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo Cleanup

    Dim vals As Variant
    vals = col.Value  ' 2D Variant array, n rows x 1 col

    Dim i As Long, j As Long, k As Long
    i = 1
    Do While i < n
        If Not IsBlankVal(vals(i, 1)) Then
            j = i + 1
            Do While j <= n
                If Not IsBlankVal(vals(j, 1)) Then Exit Do
                j = j + 1
            Loop

            If j <= n And j > i + 1 Then
                If IsNumeric(vals(i, 1)) And IsNumeric(vals(j, 1)) Then
                    Dim topV As Double, botV As Double, stepSize As Double, dp As Long
                    topV = CDbl(vals(i, 1))
                    botV = CDbl(vals(j, 1))
                    stepSize = (botV - topV) / (j - i)
                    dp = MaxDpVal(vals(i, 1), vals(j, 1))
                    For k = i + 1 To j - 1
                        vals(k, 1) = Round2(topV + stepSize * (k - i), dp)
                    Next k
                End If
            End If

            i = j
        Else
            i = i + 1
        End If
    Loop

    col.Value = vals

Cleanup:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    If errNum <> 0 Then
        MsgBox "FillGaps error: " & errDesc, vbExclamation, "Fill Gaps"
    End If
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

' First-to-last non-empty cell in c's column, using Range.Find (fast even
' on whole columns).
Private Function ColumnDataRange(c As Range) As Range
    Dim ws As Worksheet
    Set ws = c.Worksheet
    Dim fullCol As Range
    Set fullCol = ws.Columns(c.Column)

    Dim firstCell As Range, lastCell As Range
    Set firstCell = fullCol.Find(What:="*", LookIn:=xlValues, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlNext)
    If firstCell Is Nothing Then Exit Function
    Set lastCell = fullCol.Find(What:="*", LookIn:=xlValues, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then Exit Function

    If lastCell.Row > firstCell.Row Then
        Set ColumnDataRange = ws.Range( _
            ws.Cells(firstCell.Row, c.Column), _
            ws.Cells(lastCell.Row, c.Column))
    End If
End Function

Private Function IsBlankVal(v As Variant) As Boolean
    IsBlankVal = IsEmpty(v) Or (VarType(v) = vbString And Len(v) = 0)
End Function

Private Function MaxDpVal(a As Variant, b As Variant) As Long
    Dim da As Long, db As Long
    da = DpVal(a)
    db = DpVal(b)
    If da > db Then MaxDpVal = da Else MaxDpVal = db
End Function

Private Function DpVal(v As Variant) As Long
    Dim s As String
    s = Trim$(Str(v))
    Dim p As Long
    p = InStr(s, ".")
    If p = 0 Then DpVal = 0 Else DpVal = Len(s) - p
End Function

' Round half away from zero (Excel's ROUND behaviour), not VBA's banker's
' rounding. Avoids a WorksheetFunction call per cell.
Private Function Round2(x As Double, digits As Long) As Double
    Dim f As Double
    f = 10# ^ digits
    If x >= 0 Then
        Round2 = Int(x * f + 0.5) / f
    Else
        Round2 = -Int(-x * f + 0.5) / f
    End If
End Function
