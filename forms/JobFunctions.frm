VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JobFunctions 
   Caption         =   "Job_Functions"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3930
   OleObjectBlob   =   "JobFunctions.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "JobFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ShData As Worksheet
Public DaRange As Range
Public lrShData As Long
Public Jobs As Variant
Public JobFunctions As Variant

Private Sub RegisterData_Click()
JobFunctions = _
        Application.WorksheetFunction.VLookup(Me.Job_Name.Value, ShData.Range("B:C"), 2, False)
Me.Hide
End Sub

Private Sub UserForm_Initialize()

Set ShData = Sheets("Functions")
lrShData = ShData.Cells(Rows.Count, 1).End(xlUp).Row
Set DaRange = ShData.Range("A1:C" & lrShData)

'load data to deparment code
Me.Department_Code.List = Array("72", "51", "52", "20")

End Sub
Private Sub Department_Code_Change()
    Call LoadJobData
End Sub

Sub LoadJobData()
For Each cell In ShData.Range("A1:A" & lrShData)
    If cell.Value = "72" Then
        Jobs = cell.Offset(0, 1).Value
        Me.Job_Name.AddItem Jobs
    ElseIf cell.Value = "51" Then
        Jobs = cell.Offset(0, 1).Value
        Me.Job_Name.AddItem Jobs
    ElseIf cell.Value = "52" Then
        Jobs = cell.Offset(0, 1).Value
        Me.Job_Name.AddItem Jobs
    ElseIf cell.Value = "20" Then
        Jobs = cell.Offset(0, 1).Value
        Me.Job_Name.AddItem Jobs
    End If
Next
End Sub

