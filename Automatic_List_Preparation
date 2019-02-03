Option Explicit
Sub BS()

Dim i As Integer, m As Integer, n As Integer, k As Integer, l As Integer, j As Integer, max As Integer


Application.ScreenUpdating = False
Sheets("BS").Range("A2:E104876").Clear

n = Sheets("Tweet Data").Range("T1048576").End(xlUp).Row

For i = 2 To n
    k = InStr(Sheets("Tweet Data").Range("T" & i), "ban")
    l = InStr(Sheets("Tweet Data").Range("T" & i), "Ban")
    j = InStr(Sheets("Tweet Data").Range("T" & i), "BAN")
    max = WorksheetFunction.max(k, l, j)
    
    If k > 0 Or l > 0 Or j > 0 Then
        m = Sheets("BS").Range("A1048576").End(xlUp).Row
        Sheets("BS").Range("A" & m + 1) = Sheets("Tweet Data").Range("G" & i)
        Sheets("BS").Range("B" & m + 1) = Sheets("Tweet Data").Range("K" & i)
        'Sheets("BS").Range("C" & m + 1) = Application.WorksheetFunction.CountIf(Sheets("Tweet Data").Range("G:G"), Sheets("BS").Range("A" & m + 1))
        Sheets("BS").Range("C" & m + 1) = Application.WorksheetFunction.CountIfs(Sheets("Tweet Data").Range("G:G"), Sheets("BS").Range("A" & m + 1), _
        Sheets("Tweet Data").Range("T:T"), max > 0)
    End If

Next

m = Sheets("BS").Range("A1048576").End(xlUp).Row

'Columns("A:C").Select
Sheets("BS").Range("$A$1:$C$" & m).RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
    
m = Sheets("BS").Range("A1048576").End(xlUp).Row

Columns("A:C").Select
    ActiveWorkbook.Worksheets("BS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BS").Sort.SortFields.Add Key:=Range("B2:B" & m), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BS").Sort
        .SetRange Range("A1:E" & m)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Application.ScreenUpdating = True
End Sub
