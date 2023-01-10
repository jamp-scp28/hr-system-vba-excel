VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} News 
   Caption         =   "NOVEDADES"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   OleObjectBlob   =   "News.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "News"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NewData As Boolean
Public wsP As Worksheet
Public Function CalledCode(NewData As Boolean)
Dim gotData As Boolean
gotData = NewData
Set wsP = Sheets("PData") 'set sheet with the information
'If the userform is open from the update form then next happen
News.NDATE.Value = CDate(Date)
If gotData = True Then
News.BUSCADORN.Value = FGPersonal.ComboBox1.Value
'select the prev value depend on option selected
News.Show
End If
End Function
Private Sub NCANCEL_Click()
Unload Me
FGPersonal.ComboBox1.ListIndex = -1
MsgBox "Actualización Cancelada"
End Sub
'============================
' VALIDATION FOR THE DATE
'============================
Private Sub NDATE_Change()
If NDATE.TextLength > 1 And NDATE.TextLength < 3 Then
    NDATE.Value = NDATE.Value & "/"
End If
If NDATE.TextLength > 4 And NDATE.TextLength < 6 Then
    NDATE.Value = NDATE.Value & "/"
End If
End Sub
'============================
' END OF VALIDATION
'============================
'============================
' ASSIGN THE INF. TO THE CELLS
'============================
Private Sub NREG_Click()
Dim lastrow1 As Long
Set WSS = Sheets("SData")
lastrow1 = WSS.Cells(Rows.Count, 1).End(xlUp).Row + 1 'declare variable to get last row
'============================
' ASSIGN THE INF. TO THE CELLS
'============================
WSS.Cells(lastrow1, 1).Value = CDate(News.NDATE.Value)
WSS.Cells(lastrow1, 2).Value = News.BUSCADORN.Value
WSS.Cells(lastrow1, 3).Value = FGPersonal.IDENTIFICACION.Value
WSS.Cells(lastrow1, 4).Value = News.TNEW.Value
WSS.Cells(lastrow1, 5).Value = News.RBEF.Value
WSS.Cells(lastrow1, 6).Value = News.RNEW.Value
MsgBox "Novedad Registrada"
Unload Me
Call FGPersonal.UpdateInf
Unload Me
End Sub
Private Sub TNEW_Change()
'============================
' FIND THE VALUE
'============================
Dim currentRowN As Long
currentRowN = Sheets("PData").UsedRange.Find(What:=Me.BUSCADORN.Value, after:=Sheets("PData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
Dim currentValue As Long 'declare the variable that is going to store data
If News.TNEW.Value = "SALARIAL" Then
    currentValue = 23
    News.RNEW.Value = FGPersonal.SBASE.Value
    ElseIf News.TNEW.Value = "RODAMIENTO" Then
    currentValue = 24
    News.RNEW.Value = FGPersonal.RODAMIENTO.Value
    ElseIf News.TNEW.Value = "OTROS AUXILIOS" Then
    currentValue = 25
    News.RNEW.Value = FGPersonal.OAUX.Value
    ElseIf News.TNEW.Value = "TIPO DE CONTRATO" Then
    currentValue = 22
    News.RNEW.Value = FGPersonal.TCONTRATO.Value
    ElseIf News.TNEW.Value = "CARGO" Then
    currentValue = 21
    News.RNEW.Value = FGPersonal.CARGO
    Else
    News.RNEW.Value = "Sin datos"
End If
News.RBEF.Value = wsP.Cells(currentRowN, currentValue)
End Sub
Private Sub UserForm_Initialize()
News.TNEW.AddItem "SALARIAL"
News.TNEW.AddItem "RODAMIENTO"
News.TNEW.AddItem "OTROS AUXILIOS"
News.TNEW.AddItem "TIPO DE CONTRATO"
News.TNEW.AddItem "CARGO"
End Sub
