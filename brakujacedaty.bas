Attribute VB_Name = "Module1"
Sub brakujacedni()
Dim first As Date
Dim last As Date
Dim diff As Integer
Dim num As Integer
Dim col As New Collection
first = Range("H2").Value
last = Range("H274").Value
diff = last - first
For count1 = 1 To diff
g = first + count1
num = Application.WorksheetFunction.CountIf(Range("H2:H274"), g)
If num = 0 Then
col.Add g
End If
Next
Range("K2").Select
For Each el In col
ActiveCell.Value = el
ActiveCell.Offset(1, 0).Select
Next
End Sub
