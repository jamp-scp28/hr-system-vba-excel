VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IllTracking 
   Caption         =   "Seguimiento de Incapacidades"
   ClientHeight    =   9165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6285
   OleObjectBlob   =   "IllTracking.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "IllTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public wsP As Worksheet
Public wsD As Worksheet
Public wsE As Worksheet
Public wsOD As Worksheet
Public lrwsD As Long
Public lrwsP As Long
Public lrwsE As Long
Public FindRow As Long
Public myRange As Range
Public CurrentRowP As Long
Public CurrentRow As Long
Public CurrentRowOD As Long
Public wageS As Long
Public valueSelect As Long
Public FoundCells As Range
'========================================================
'STATIC CODE
'========================================================
Private Sub ISEARCH_Results_Click()
'==Get the position of the selected value on listbox
Dim lr As String
Dim Index As Integer
Dim strAddress As String
Dim ValueTFind As String
Dim l As Long
    For l = 0 To Me.ISEARCH_Results.ListCount
        If Me.ISEARCH_Results.Selected(l) = True Then
            On Error Resume Next
            strAddress = Me.ISEARCH_Results.List(l, IData.[inc_name].Column)
            ValueTFind = wsD.Range(strAddress).Value
            valueSelect = wsD.Range(strAddress).Row
            GoTo EndLoop
        End If
    Next l
EndLoop:
'==Get data and asign it to the combobox fields
Set myRange = wsD.Range("A2:O" & lrwsD)
On Error Resume Next
Me.SEMP.Value = ValueTFind
Me.SDATE.Value = wsD.Cells(valueSelect, 5).Value
Me.SENT.Value = wsD.Cells(valueSelect, 6).Value
Me.SDATEI.Value = wsD.Cells(valueSelect, 10).Value
Me.SDATEF.Value = wsD.Cells(valueSelect, 11).Value
Me.SCOST.Value = wsD.Cells(valueSelect, 12).Value
Me.SFEC.Value = wsD.Cells(valueSelect, 13).Value
Me.SVREI.Value = wsD.Cells(valueSelect, 14).Value
Me.SOBS.Value = wsD.Cells(valueSelect, 15).Value
End Sub
'========Begin Validation For Date of Request
Private Sub SDATE_Change()
If SDATE.TextLength > 1 And SDATE.TextLength < 3 Then
    SDATE.Value = SDATE.Value & "/"
End If
If SDATE.TextLength > 10 Then
    SDATE.Value = Mid(SDATE.Text, 1, Len(SDATE.Text) - 1)
End If
    If SDATE.TextLength > 4 And SDATE.TextLength < 6 Then
    SDATE.Value = SDATE.Value & "/"
End If
End Sub
Private Sub SDATE_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       SDATE.Value = vbNullString
    End If
End Sub
Private Sub SDATE_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If SDATE.TextLength > 1 And SDATE.TextLength < 10 And SDATE.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    SDATE.Value = vbNullString
    SDATE.SetFocus
    Exit Sub
End If
End Sub
'========End Validation For Date of Request
'========Begin Validation For Initial Date of Ill
Private Sub SDATEI_Change()
If SDATEI.TextLength > 1 And SDATEI.TextLength < 3 Then
    SDATEI.Value = SDATEI.Value & "/"
End If
If SDATEI.TextLength > 10 Then
    SDATEI.Value = Mid(SDATEI.Text, 1, Len(SDATEI.Text) - 1)
End If
    If SDATEI.TextLength > 4 And SDATEI.TextLength < 6 Then
    SDATEI.Value = SDATEI.Value & "/"
End If
If Me.SDATEI.TextLength = 10 Then
    Me.SDATEF.Enabled = True
    Else
    Me.SDATEF.Enabled = False
End If
End Sub
Private Sub SDATEI_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       SDATEI.Value = vbNullString
    End If
End Sub
Private Sub SDATEI_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If SDATEI.TextLength > 1 And SDATEI.TextLength < 10 And SDATEI.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    SDATEI.Value = vbNullString
    SDATEI.SetFocus
    Exit Sub
End If
End Sub
'========End Validation For Initial Date of Ill
'========Begin Validation For End Date of Ill
Private Sub SDATEF_Change()
If SDATEF.TextLength > 1 And SDATEF.TextLength < 3 Then
    SDATEF.Value = SDATEF.Value & "/"
End If
If SDATEF.TextLength > 10 Then
    SDATEF.Value = Mid(SDATEF.Text, 1, Len(SDATEF.Text) - 1)
End If
    If SDATEF.TextLength > 4 And SDATEF.TextLength < 6 Then
    SDATEF.Value = SDATEF.Value & "/"
End If
End Sub
Private Sub SDATEF_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       SDATEF.Value = vbNullString
    End If
End Sub
Private Sub SDATEF_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If SDATEF.TextLength > 1 And SDATEF.TextLength < 10 And SDATEF.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    SDATEF.Value = vbNullString
    SDATEF.SetFocus
    Exit Sub
End If
'==Get the days of leave for ill
Dim Days As Long: Days = CDate(Me.SDATEF) - CDate(Me.SDATEI)
Me.SCOST.Value = Int((((wageS / 30) * (Days + 1)) - ((wageS / 30) * 2)) * (66.667 / 100))
End Sub
'========End Validation For End Date of Ill
'========Begin Validation For Date of Transfer
Private Sub SFEC_Change()
If SFEC.TextLength > 1 And SFEC.TextLength < 3 Then
    SFEC.Value = SFEC.Value & "/"
End If
If SFEC.TextLength > 10 Then
    SFEC.Value = Mid(SFEC.Text, 1, Len(SFEC.Text) - 1)
End If
    If SFEC.TextLength > 4 And SFEC.TextLength < 6 Then
    SFEC.Value = SFEC.Value & "/"
End If
End Sub
Private Sub SFEC_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       SFEC.Value = vbNullString
    End If
End Sub
Private Sub SFEC_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If SFEC.TextLength > 1 And SFEC.TextLength < 10 And SFEC.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    SFEC.Value = vbNullString
    SFEC.SetFocus
    Exit Sub
End If
End Sub
'========End Validation For Date of Transfer
'*****************************Begin Functionality of Buttons
Private Sub ISEARCH_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Calls the FindAllMatches routine as user types text in the textbox
    Call FindallMatches
End Sub
Sub FindallMatches()
'Find all matches on activesheet
'Called by: TextBox_Find_KeyUp event
Dim SearchRange As Range, arrResults() As Variant, FindWhat As Variant
Dim FoundCell As Range, lFound As Long
Dim lSearchCol As Long, lLastRow As Long
    If Len(ISEARCH.Value) > 1 Then 'Do search if text in find box is longer than 1 character.
        Set SearchRange = wsD.Range("A:A")
        FindWhat = Me.ISEARCH.Value
        'Calls the FindAll function
        Set FoundCells = FindAll(SearchRange:=SearchRange, _
                                FindWhat:=FindWhat, _
                                LookIn:=xlValues, _
                                LookAt:=xlPart, _
                                SearchOrder:=xlByColumns, _
                                MatchCase:=False, _
                                BeginsWith:=vbNullString, _
                                EndsWith:=vbNullString, _
                                BeginEndCompare:=vbTextCompare)
        If FoundCells Is Nothing Then
            ReDim arrResults(1 To 1, 1 To 2)
            arrResults(1, 1) = "Sin resultados"
        Else
            'Add results of FindAll to an array
            ReDim arrResults(1 To FoundCells.Count, 1 To 2)
            lFound = 1
            For Each FoundCell In FoundCells
                arrResults(lFound, 1) = FoundCell.Value
                arrResults(lFound, 2) = FoundCell.Address
                lFound = lFound + 1
            Next FoundCell
        End If
        'Populate the listbox with the array
        Me.ISEARCH_Results.List = arrResults
    Else
        Me.ISEARCH_Results.Clear
    End If
End Sub
'*****************************End Find Data
Private Sub UserForm_Initialize()
Set wsP = Sheets("PData"): Set wsD = Sheets("IData"): Set wsE = Sheets("C_CIE10")
Set wsOD = Sheets("C_CIE10")
'==Get last Row from both sheets
lrwsP = wsP.Cells(Rows.Count, 1).End(xlUp).Row
lrwsD = wsD.Cells(Rows.Count, 2).End(xlUp).Row
lrwsE = wsE.Cells(Rows.Count, 4).End(xlUp).Row
'==Add data to listbox employees
Me.ISEARCH.List = wsP.Range("A2:A" & lrwsP).Value
Me.SENT.List = wsE.Range("D2:D" & lrwsE).Value
Me.SEMP.List = wsP.Range("A2:A" & lrwsP).Value
'==Disable buttons to avoid errors
Me.SUPD.Enabled = False
Me.SDATEF.Enabled = False
End Sub
Private Sub ISEARCH_Change()
'Go to selection on sheet when result is clicked
Dim myRange As Range
Set myRange = wsD.Range("A:B")
'Disabled button to avoid errors
If Me.ISEARCH.Value = "" Then
    Me.SNING.Enabled = True
    Call DeleteData
    Else
    Me.SNING.Enabled = False
End If
'==Validate value of listbox to avoid error when buttons clicked
If Me.ISEARCH.Value <> "" Then
    Me.SUPD.Enabled = True
    Me.SNING.Enabled = False
    Else
    Me.SUPD.Enabled = False
    Me.SNING.Enabled = True
End If
'==Get the currentrow of the value selected
On Error Resume Next
CurrentRowP = wsP.UsedRange.Find(What:=Me.ISEARCH, after:=wsP.Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
wageS = wsP.Cells(CurrentRow, wsP.[wage].Column).Value
End Sub
Private Sub SEMP_Change()
'==Code to get the current row of the value selected on ISEARCH
On Error Resume Next
CurrentRow = wsP.UsedRange.Find(What:=Me.SEMP, after:=wsP.Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
wageS = wsP.Cells(CurrentRow, wsP.[wage].Column).Value
Me.SENT.Value = wsP.Cells(CurrentRow, wsP.[EPS].Column).Value
End Sub

Private Sub SENT_Change()
'==Get the CurrentRow on the sheet with the data of the social security administrators
On Error Resume Next
CurrentRowOD = wsOD.Range("D2:J200").Find(What:=Me.SENT, after:=wsOD.Range("D2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
End Sub
Private Sub SCAN_Click()
Unload Me
Sheets("PPrincipal").Select
End Sub
'===========================================================
'*****************************Start Functionality of Buttons
'===========================================================
Private Sub SNING_Click()
'====Do not allow empty fields in key data
Dim msgE As String: msgE = "Diligencie el Dato"
If Me.SEMP.Value = "" Then
    MsgBox msgE
    Me.SEMP.SetFocus
    Exit Sub
End If
If Me.SDATE.Value = "" Then
    MsgBox msgE
    Me.SDATE.SetFocus
    Exit Sub
End If
If Me.SENT.Value = "" Then
    MsgBox msgE
    Me.SENT.SetFocus
    Exit Sub
End If
If Me.SDATEI.Value = "" Then
    MsgBox msgE
    Me.SDATEI.SetFocus
    Exit Sub
End If
If Me.SDATEF.Value = "" Then
    MsgBox msgE
    Me.SDATEF.SetFocus
    Exit Sub
End If
'==========================================
'==Add the data to the sheets of tracking
wsD.Cells(lrwsD + 1, IData.[inc_name].Column).Value = Me.SEMP.Value
wsD.Cells(lrwsD + 1, IData.[inc_id].Column).Value = wsP.Cells(CurrentRow, wsP.[ID].Column).Value 'Get the id Number
wsD.Cells(lrwsD + 1, IData.[id_jobname].Column).Value = wsP.Cells(CurrentRow, wsP.[JOBNAME].Column).Value 'Get the charge name
wsD.Cells(lrwsD + 1, IData.[inc_wage].Column).Value = wsP.Cells(CurrentRow, wsP.[wage].Column).Value 'Get the wage
wsD.Cells(lrwsD + 1, IData.[inc_dated_register].Column).Value = CDate(Me.SDATE.Value)
wsD.Cells(lrwsD + 1, IData.[inc_eps].Column).Value = wsP.Cells(CurrentRow, wsP.[EPS].Column).Value
wsD.Cells(lrwsD + 1, IData.[inc_nit].Column).Value = wsOD.Cells(CurrentRowOD, 7).Value 'Get the Nit
wsD.Cells(lrwsD + 1, IData.[inc_address].Column).Value = wsOD.Cells(CurrentRowOD, 8).Value 'Get Address
wsD.Cells(lrwsD + 1, IData.[inc_phone].Column).Value = wsOD.Cells(CurrentRowOD, 10).Value 'Get number phone
wsD.Cells(lrwsD + 1, IData.[inc_initial_dated].Column).Value = CDate(Me.SDATEI.Value)
wsD.Cells(lrwsD + 1, IData.[inc_final_dated].Column).Value = CDate(Me.SDATEF.Value)
wsD.Cells(lrwsD + 1, IData.[inc_cost].Column).Value = Me.SCOST.Value
'**********End days of leave for ill
If Me.SFEC.Value <> "" Then
    wsD.Cells(lrwsD + 1, IData.[inc_dated_devolution].Column).Value = CDate(Me.SFEC.Value)
Else
    wsD.Cells(lrwsD + 1, IData.[inc_dated_devolution].Column).Value = ""
End If
wsD.Cells(lrwsD + 1, IData.[inc_payment].Column).Value = Me.SVREI.Value
wsD.Cells(lrwsD + 1, IData.[inc_observation].Column).Value = Me.SOBS.Value

MsgBox "Datos Ingresados Correctamente"
Call DeleteData
Call UserForm_Initialize
Me.ISEARCH.SetFocus
Me.Hide
ITrack = False
End Sub
'==============================================Update Data
Private Sub SUPD_Click()
'====Do not allow empty fields in key data
Dim msgE As String: msgE = "Diligencie el Dato"
If Me.SEMP.Value = "" Then
    MsgBox msgE
    Me.SEMP.SetFocus
    Exit Sub
End If
If Me.SDATE.Value = "" Then
    MsgBox msgE
    Me.SDATE.SetFocus
    Exit Sub
End If
If Me.SENT.Value = "" Then
    MsgBox msgE
    Me.SENT.SetFocus
    Exit Sub
End If
If Me.SDATEI.Value = "" Then
    MsgBox msgE
    Me.SDATEI.SetFocus
    Exit Sub
End If
If Me.SDATEF.Value = "" Then
    MsgBox msgE
    Me.SDATEF.SetFocus
    Exit Sub
End If
'==========================================
'==================================Add Data
wsD.Cells(valueSelect, IData.[inc_name].Column).Value = Me.SEMP.Value
wsD.Cells(valueSelect, 2).Value = wsP.Cells(CurrentRow, 2).Value 'Get the id Number
wsD.Cells(valueSelect, 3).Value = wsP.Cells(CurrentRow, 21).Value 'Get the charge name
wsD.Cells(valueSelect, 4).Value = wageS 'Get the wage
wsD.Cells(valueSelect, 5).Value = CDate(Me.SDATE.Value)
wsD.Cells(valueSelect, 6).Value = Me.SENT
wsD.Cells(valueSelect, 7).Value = wsOD.Cells(CurrentRowOD, 7).Value 'Get the Nit
wsD.Cells(valueSelect, 8).Value = wsOD.Cells(CurrentRowOD, 8).Value 'Get Address
wsD.Cells(valueSelect, 9).Value = wsOD.Cells(CurrentRowOD, 10).Value 'Get number phone
wsD.Cells(valueSelect, 10).Value = CDate(Me.SDATEI.Value)
wsD.Cells(valueSelect, 11).Value = CDate(Me.SDATEF.Value)
wsD.Cells(valueSelect, 12).Value = Int(Me.SCOST.Value)
'**********End days of leave for ill
If Me.SFEC.Value <> "" Then
    wsD.Cells(valueSelect, 13).Value = CDate(Me.SFEC.Value)
Else
    wsD.Cells(valueSelect, 13).Value = ""
End If
wsD.Cells(valueSelect, 14).Value = Me.SVREI.Value
wsD.Cells(valueSelect, 15).Value = Me.SOBS.Value
'=============================================
MsgBox "Datos Actualizados Correctamente"
Call DeleteData
Call UserForm_Initialize
Me.ISEARCH.SetFocus
End Sub
'==========================================End update Data
'=========================================================
'*****************************End Functionality of Buttons
'=========================================================
'*****************************Get Data and assign to textbox
Private Sub ISEARCH_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyBack Then
        Me.ISEARCH.Value = vbNullString
    End If
End Sub
Sub DeleteData()
Me.ISEARCH.Value = ""
Me.ISEARCH_Results.Clear
Me.SEMP.Value = ""
Me.SDATE.Value = ""
Me.SENT.Value = ""
Me.SDATEI.Value = ""
Me.SDATEF.Value = ""
Me.SCOST.Value = ""
Me.SFEC.Value = ""
Me.SVREI.Value = ""
Me.SOBS.Value = ""
End Sub

