Attribute VB_Name = "Module1"
Sub usunwiersze()
Dim lastRow As Long
Dim i As Long
lastRow = Range("B2").End(xlDown).Row
For i = lastRow To 2 Step -1
If Range("C" & i).Value = 0 Then
Range("C" & i).EntireRow.Delete
End If
Next
End Sub
