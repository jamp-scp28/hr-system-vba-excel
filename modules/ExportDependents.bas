Attribute VB_Name = "ExportDependents"
Public RngEmployees As Range
Public ShtDependents As Worksheet
Public ShtResult As Worksheet
Public RngInner
Public ExportData
Sub ExportDependentsData()

Set ShtDependents = Sheets("Pdep")
Set ShtResult = Sheets("Result")
Set RngEmployees = ShtDependents.Range("A:A")

For Each cell In RngEmployees
    If cell.Offset(0, 2) <> "" Then
         'MsgBox Cell.Offset(0, 2).Column
    End If
Next
End Sub



