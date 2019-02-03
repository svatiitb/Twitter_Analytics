Option Explicit

Sub TI()

Dim i As Integer, j As Integer, n As Integer, m As Integer, ID As Integer, k As Integer
Dim TH As String, link As String
Dim chromePath As String

Application.ScreenUpdating = False
Sheets("TI").Range("A2:F104876").Clear

For i = 2 To Sheets("Tweet Data").Range("A1048576").End(xlUp).Row
    
    TH = UCase(Right(Sheets("Tweet Data").Range("G" & i), _
    Len(Sheets("Tweet Data").Range("G" & i)) - 1))
    On Error GoTo Step_1
    ID = Application.WorksheetFunction.VLookup(UCase(Right(Sheets("Tweet Data").Range("G" & i), _
    Len(Sheets("Tweet Data").Range("G" & i)) - 1)), Sheets("master employee list").Range("B:D"), 2, False)

Next

n = Sheets("TI").Range("A1048576").End(xlUp).Row
'Sheets("TI").Columns("A:A").Select
'ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
Sheets("TI").Range("$A$1:$C$" & n).RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes

For i = 2 To Sheets("TI").Range("A1048576").End(xlUp).Row

    Sheets("TI").Range("B" & i) = Application.WorksheetFunction.VLookup(Sheets("TI").Range("A" & i), Sheets("Tweet Data").Range("G:K"), 5, False)
    Sheets("TI").Range("C" & i) = Application.WorksheetFunction.CountIf(Sheets("Tweet Data").Range("G:G"), Sheets("TI").Range("A" & i))
    Sheets("TI").Range("G" & i) = Application.WorksheetFunction.Index(Sheets("Tweet Data").Range("C:G"), Application.WorksheetFunction.Match( _
    Sheets("TI").Range("A" & i), Sheets("Tweet Data").Range("G:G"), 0), 1)
    
Next

n = Sheets("TI").Range("A1048576").End(xlUp).Row
Sheets("TI").Range("A:E").Select
    Worksheets("TI").Sort.SortFields.Clear
    Worksheets("TI").Sort.SortFields.Add Key:=Range("B2:B" & n), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With Worksheets("TI").Sort
        .SetRange Range("A1:E" & n)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

n = Sheets("TI").Range("A1048576").End(xlUp).Row
Sheets("TI").Range("J2") = n - 1
m = Sheets("TI").Range("G1048576").End(xlUp).Row
On Error Resume Next
For i = 2 To m
    Sheets("TI").Range("XFD1").Select
    j = Application.WorksheetFunction.Match(Sheets("TI").Range("H" & i), Sheets("TI").Range("A:A"), 0)
    Sheets("TI").Range("A" & j & ":" & "G" & j).Delete Shift:=xlUp
    j = 0
    
Next

chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
On Error Resume Next
For n = 2 To 10
On Error GoTo 0
link = Sheets("TI").Range("G" & n)
Shell (chromePath & " -url " & link)
Next

Exit Sub
Step_1:
    
    m = Sheets("TI").Range("A1048576").End(xlUp).Row
    Sheets("TI").Range("A" & m + 1) = Sheets("Tweet Data").Range("G" & i)
    
    Resume Next

Application.ScreenUpdating = True

End Sub
