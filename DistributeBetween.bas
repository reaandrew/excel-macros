Attribute VB_Name = "DistributeBetween"
Option Explicit

' Linearly interpolate the cells between the top and bottom of a
' single-column selection. The top and bottom cells stay as the anchors;
' every cell in between is overwritten with the interpolated value,
' rounded to the larger decimal-place count of the two anchors.
'
' Usage: select e.g. A2:A10 (A2 and A10 already contain prices), then run.
Public Sub DistributeBetween()
    Dim rng As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range of cells first.", vbExclamation, "Distribute Between"
        Exit Sub
    End If
    Set rng = Selection

    If rng.Areas.Count > 1 Then
        MsgBox "Select a single contiguous range.", vbExclamation, "Distribute Between"
        Exit Sub
    End If

    If rng.Columns.Count <> 1 Then
        MsgBox "Selection must be in a single column.", vbExclamation, "Distribute Between"
        Exit Sub
    End If

    Dim n As Long
    n = rng.Rows.Count
    If n < 3 Then
        MsgBox "Select at least 3 cells (top anchor, middle cells, bottom anchor).", _
               vbExclamation, "Distribute Between"
        Exit Sub
    End If

    Dim topCell As Range, botCell As Range
    Set topCell = rng.Cells(1, 1)
    Set botCell = rng.Cells(n, 1)

    If Not IsNumeric(topCell.Value) Or Not IsNumeric(botCell.Value) Then
        MsgBox "The first and last cells must both contain numbers.", _
               vbExclamation, "Distribute Between"
        Exit Sub
    End If

    Dim topVal As Double, botVal As Double
    topVal = CDbl(topCell.Value)
    botVal = CDbl(botCell.Value)

    Dim intervals As Long
    intervals = n - 1
    Dim stepSize As Double
    stepSize = (botVal - topVal) / intervals

    Dim dp As Long
    dp = MaxDecimalPlaces(topCell, botCell)

    Application.ScreenUpdating = False
    Dim i As Long
    For i = 2 To n - 1
        rng.Cells(i, 1).Value = Application.WorksheetFunction.Round( _
            topVal + stepSize * (i - 1), dp)
    Next i
    Application.ScreenUpdating = True
End Sub

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
    ' Counts decimal places in the underlying numeric value, not the
    ' display format. Uses Str() so the result is locale-independent
    ' ("." as decimal separator).
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
