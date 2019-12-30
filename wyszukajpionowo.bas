Attribute VB_Name = "UsingVlookup"
Sub VlookupExample()
Dim goalsWs As Worksheet
Dim dataWs As Worksheet
Dim goalsLastRow As Long
Dim dataLastRow As Long
Dim dataRng As Range
Set goalsWs = ThisWorkbook.Worksheets("Goalscorers")
Set dataWs = ThisWorkbook.Worksheets("PlacesOfBirth")
goalsLastRow = goalsWs.Range("B" & Rows.Count).End(xlUp).Row
dataLastRow = dataWs.Range("A" & Rows.Count).End(xlUp).Row
Set dataRng = dataWs.Range("A2:B" & dataLastRow)

For x = 2 To goalsLastRow
    On Error Resume Next
        goalsWs.Range("C" & x).Value = WorksheetFunction.VLookup( _
        goalsWs.Range("A" & x).Value, dataRng, 2, 0)
Next
End Sub
