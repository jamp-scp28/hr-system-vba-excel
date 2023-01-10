VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VacationsI 
   Caption         =   "SEGUIMIENTO VACACIONES"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   OleObjectBlob   =   "VacationsI.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "VacationsI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================
' DECLARE THE VARIABLES & SET VALUES
'====================================
Public wsPa As Worksheet
Public wsVa As Worksheet
Public lastrowP As Long
Public lastrowV As Long
Public myRangeP As Range
Public myRangeV As Range
Public CurrentRow As Long
Public vv As UserForm
Private Sub BUSCADORV_Change()
Call VlookUp_Value
End Sub
Private Sub CANCEL_Click()
Unload Me 'CLOSE FORM
Sheets("PPrincipal").Select
End Sub

Private Sub FCIV_Change()
'VALIDATION FOR DATE OF DATE OF CONTRACT I------------------
If FCIV.TextLength > 1 And FCIV.TextLength < 3 Then
FCIV.Value = FCIV.Value & "/"
End If
If FCIV.TextLength > 10 Then
    FCIV.Value = Mid(FCIV.Text, 1, Len(FCIV.Text) - 1)
End If
If FCIV.TextLength > 4 And FCIV.TextLength < 6 Then
FCIV.Value = FCIV.Value & "/"
End If
End Sub
Private Sub FCIV_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       FCIV.Value = vbNullString
    End If
End Sub
Private Sub FCIV_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FCIV.TextLength > 1 And FCIV.TextLength < 10 And FCIV.Value <> vbNullString Then
MsgBox "Ingrese una fecha en formato DD/MM/AAAA"
FCIV.Value = vbNullString
CANCEL = True
End If
'END OF VALIDATION OF DATE OF FCIV DATE OF CONTRACT I------------------
End Sub
Private Sub FDL_Change()
'VALIDATION FOR DATE OF LIQUIDATION------------------
If FDL.TextLength > 1 And FDL.TextLength < 3 Then
FDL.Value = FDL.Value & "/"
End If
If FDL.TextLength > 10 Then
    FDL.Value = Mid(FDL.Text, 1, Len(FDL.Text) - 1)
End If
If FDL.TextLength > 4 And FDL.TextLength < 6 Then
FDL.Value = FDL.Value & "/"
End If
End Sub
Private Sub FDL_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       FDL.Value = vbNullString
    End If
End Sub
Private Sub FDL_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FDL.TextLength > 1 And FDL.TextLength < 10 And FDL.Value <> vbNullString Then
MsgBox "Ingrese una fecha en formato DD/MM/AAAA"
FDL.Value = vbNullString
CANCEL = True
End If
'END OF VALIDATION OF DATE OF LIQUIDATION------------------
End Sub
Private Sub UpdateV_Click()
Set vv = VacationsI
'====================================
' FIND THE CURRENT POSITION OF VALUE SELECTED IN COMBOBOX
'====================================
CurrentRow = wsVa.UsedRange.Find(What:=Me.BUSCADORV, after:=wsVa.Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
'====================================
' ADD THE VALUES OF THE FIELD TO THE CELLS
'====================================
wsVa.Cells(CurrentRow, VData.[vac_und_contract_dated].Column).Value = CDate(VacationsI.FCIV.Value)
wsVa.Cells(CurrentRow, VData.[vac_liquidation_dated].Column).Value = CDate(VacationsI.FDL.Value)
wsVa.Cells(CurrentRow, VData.[vac_days_emp_bef].Column).Value = VacationsI.VAD.Value
MsgBox "Datos actualizados"
Call VlookUp_Value
BUSCADORV.SetFocus
End Sub
Private Sub UserForm_Initialize()
Me.MultiPage1.Value = 0
'====================================
' SET VALUES TO VARIABLES
'====================================
Set wsPa = Sheets("PData") 'sheet to load the names
Set wsVa = Sheets("VData") 'sheet where the information is going to be modify
lastrowP = wsPa.Cells(Rows.Count, 1).End(xlUp).Row 'get last name from sheet PData
lastrowV = wsVa.Cells(Rows.Count, 1).End(xlUp).Row 'get the last row of sheet with the vacation data
'====================================
' ADD THE NAMES TO THE COMBOBOX TO SEARCH
'====================================
VacationsI.BUSCADORV.List = wsPa.Range("b2:b" & lastrowP).Value
End Sub
Sub VlookUp_Value() 'Sub to find data and refresh it
Set vv = VacationsI 'select userform
Set myRangeP = wsPa.Range("B1:bo" & lastrowP) 'select range to search data
Set myRangeV = wsVa.Range("B1:Q" & lastrowV) 'select range to search data
On Error Resume Next 'if an error happen go to the next
'====================================
' FIND THE DATA FROM THE SHEET AND PUT IT IN THE FIELDS
'====================================
vv.NOMBRESVA.Value = _
    Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeP, PData.[EMPNAME].Column - 1, False)
If Err.Number <> 0 Then vv.NOMBRESVA.Value = "no encontrado"
vv.IDENTIFICACIONV.Value = _
    Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeP, PData.[ID].Column - 1, False)
If Err.Number <> 0 Then vv.IDENTIFICACIONV.Value = "no encontrado"
vv.CDEPV.Value = _
    Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeP, PData.[DEPARTNAME].Column - 1, False)
If Err.Number <> 0 Then vv.CDEPV.Value = "no encontrado"
vv.CARGOV.Value = _
    Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeP, PData.[JOBNAME].Column - 1, False)
If Err.Number <> 0 Then vv.CARGOV.Value = "no encontrado"
vv.SBASEV.Value = _
    Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeP, PData.[wage].Column - 1, False)
If Err.Number <> 0 Then vv.SBASEV.Value = "no encontrado"
vv.DVC.Value = _
    Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeV, VData.[vac_days_emp].Column - 1, False)
If Err.Number <> 0 Then vv.DVC.Value = "no encontrado"
vv.DVD.Value = _
    Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeV, VData.[vac_taken_days].Column - 1, False)
If Err.Number <> 0 Then vv.DVD.Value = "no encontrado"
vv.DVP.Value = _
    Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeV, VData.[vac_days_aval].Column - 1, False)
If Err.Number <> 0 Then vv.DVP.Value = "no encontrado"
vv.VVAC.Value = _
    Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeV, VData.[vac_cost].Column - 1, False)
If Err.Number <> 0 Then vv.VVAC.Value = "no encontrado"
vv.FDL.Value = _
    CDate(Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeV, VData.[vac_liquidation_dated].Column - 1, False))
If Err.Number <> 0 Then vv.FDL.Value = "no encontrado"
vv.FCIV.Value = _
    CDate(Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeV, VData.[vac_und_contract_dated].Column - 1, False))
If Err.Number <> 0 Then vv.FCIV.Value = "no encontrado"
vv.VAD.Value = _
    Format(CDate(Application.WorksheetFunction.VLookup(vv.BUSCADORV, myRangeV, VData.[vac_days_emp_bef].Column - 1, False)), "0")
If Err.Number <> 0 Then vv.VAD.Value = "no encontrado"
End Sub
