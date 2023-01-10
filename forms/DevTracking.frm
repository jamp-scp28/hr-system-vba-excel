VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DevTracking 
   Caption         =   "Seguimiento Devolución de Aportes"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6015
   OleObjectBlob   =   "DevTracking.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "DevTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lrEmployeeSheet As Long
Public lr As Long
Public lrEntity As Long
Public employeeSheet As Worksheet
Public dataSheet As Worksheet
Public infoSheet As Worksheet
Public rnData As Range
Public rnInfo As Range
Public crData As Long
Public crInfo As Long
Public FindValue As String
Public ValueTFind As String
Public strAddress As String
Public KeyBack1 As Boolean
Public KeyBack2 As Boolean
Public myRangeEmployee As Range
Public currentrowEmployee As Long
Public myRangeInfo As Range
'Begin of validation for Date field
Private Sub DDATE_Change()
If KeyBack1 = False Then
    If Me.DDATE.TextLength > 1 And Me.DDATE.TextLength < 3 Then
        Me.DDATE.Value = Me.DDATE.Value & "/"
    End If
    If Me.DDATE.TextLength > 4 And Me.DDATE.TextLength < 6 Then
        Me.DDATE.Value = Me.DDATE & "/"
    End If
    If Me.DDATE.TextLength > 10 Then
        Me.DDATE.Value = Mid(Me.DDATE.Value, 1, Len(Me.DDATE.Text) - 1)
    End If
    Else
End If
End Sub
Private Sub DDATE_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If Me.DDATE.TextLength < 10 And Me.DDATE.Value <> "" Then
    MsgBox "La fecha debe ser en formato DD/MM/YYYY"
    CANCEL = True
End If
End Sub
Private Sub DDATE_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyBack And Me.DDATE.Value <> "" Then
    KeyBack1 = True
    Me.DDATE.Value = Mid(Me.DDATE.Value, 1, Len(Me.DDATE.Text) - 1)
    Else
    KeyBack1 = False
End If
End Sub
'Begin of validation for date field
Private Sub DFEC_Change()
If KeyBack2 = False Then
    If Me.DFEC.TextLength > 1 And Me.DFEC.TextLength < 3 Then
        Me.DFEC.Value = Me.DFEC.Value & "/"
    End If
    If Me.DFEC.TextLength > 4 And Me.DFEC.TextLength < 6 Then
        Me.DFEC.Value = Me.DFEC.Value & "/"
    End If
    If Me.DFEC.TextLength > 10 Then
        Me.DFEC.Value = Mid(Me.DFEC.Value, 1, Len(Me.DFEC.Text) - 1)
    End If
End If
End Sub
Private Sub DFEC_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyBack And Me.DFEC.Value <> "" Then
    KeyBack2 = True
    Me.DFEC.Value = Mid(Me.DFEC.Value, 1, Len(Me.DFEC.Value) - 1)
End If
End Sub
Private Sub DFEC_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If Me.DFEC.TextLength < 10 And Me.DFEC.Value <> "" Then
    MsgBox "La fecha debe estar en formato DD/MM/AAAA"
    CANCEL = True
End If
End Sub
'End of validation for date field
'End of validation for Date field
Private Sub DSEARCH_Change()
If Me.DSEARCH.Value <> "" Then
    Me.SNING.Enabled = False
    Me.SUPD.Enabled = True
    Else
    Me.SUPD.Enabled = False
    Me.SNING.Enabled = True
End If
End Sub
Private Sub UserForm_Initialize()
Application.EnableCancelKey = xlDisabled
Set employeeSheet = Sheets("PData")
Set dataSheet = Sheets("APData")
Set infoSheet = Sheets("C_CIE10")
lrEntity = infoSheet.Cells(Rows.Count, 4).End(xlUp).Row
lrEmployeeSheet = employeeSheet.Cells(Rows.Count, 1).End(xlUp).Row
lr = dataSheet.Cells(Rows.Count, 1).End(xlUp).Row
Me.DENT.List = infoSheet.Range("D2:D" & lrEntity).Value
Me.DSEARCH.List = employeeSheet.Range("A2:A" & lrEmployeeSheet).Value
Me.DEMP.List = employeeSheet.Range("A2:A" & lrEmployeeSheet).Value
'Set ranges for next process
Set rnData = dataSheet.Range("A:N")
Set rnInfo = infoSheet.Range("D:J")
Me.SUPD.Enabled = False
Set myRangeEmployee = employeeSheet.Range("A1:AO" & lrEmployeeSheet)
Set myRangeInfo = infoSheet.Range("D1:J" & lrEntity)
End Sub
Private Sub SNING_Click()
'Do not allow blank spaces in some of the fields
Dim msg As String
msg = "Diligencie Campo"
If Me.DSEARCH.Value <> "" Then
    Me.DSEARCH.Value = ""
End If
If Me.DSEARCH_Results.Value <> "" Then
    Me.DSEARCH_Results.Clear
End If
If Me.DDATE.Value = "" Then
    MsgBox msg
    Me.DDATE.SetFocus
    Exit Sub
End If
If Me.DEMP.Value = "" Then
    MsgBox msg
    Me.DEMP.SetFocus
    Exit Sub
End If
If Me.DENT.Value = "" Then
    MsgBox msg
    Me.DENT.SetFocus
    Exit Sub
End If
If Me.DSOL.Value = "" Then
    MsgBox msg
    Me.DSOL.SetFocus
    Exit Sub
End If
If Me.DOBS.Value = "" Then
    MsgBox msg
    Me.DOBS.SetFocus
    Exit Sub
End If
'Register the new data in the lastrow
dataSheet.Cells(lr + 1, 1).Value = Me.DEMP
dataSheet.Cells(lr + 1, 2).Value = _
    Application.WorksheetFunction.VLookup(Me.DEMP, myRangeEmployee, 2, False)
    If Err.Number <> 0 Then Exit Sub
dataSheet.Cells(lr + 1, 3).Value = _
    Application.WorksheetFunction.VLookup(Me.DEMP, myRangeEmployee, 21, False)
    If Err.Number <> 0 Then Exit Sub
dataSheet.Cells(lr + 1, 4).Value = _
    Application.WorksheetFunction.VLookup(Me.DEMP, myRangeEmployee, 23, False)
    If Err.Number <> 0 Then Exit Sub
dataSheet.Cells(lr + 1, 5).Value = Me.DDATE
dataSheet.Cells(lr + 1, 6).Value = Me.REQMOT
dataSheet.Cells(lr + 1, 7).Value = Me.DENT
dataSheet.Cells(lr + 1, 8).Value = _
    Application.WorksheetFunction.VLookup(Me.DENT, myRangeInfo, 4, False)
    If Err.Number <> 0 Then Exit Sub
dataSheet.Cells(lr + 1, 9).Value = _
    Application.WorksheetFunction.VLookup(Me.DENT, myRangeInfo, 5, False)
    If Err.Number <> 0 Then Exit Sub
dataSheet.Cells(lr + 1, 10).Value = _
    Application.WorksheetFunction.VLookup(Me.DENT, myRangeInfo, 7, False)
    If Err.Number <> 0 Then Exit Sub
dataSheet.Cells(lr + 1, 11).Value = Me.DSOL
dataSheet.Cells(lr + 1, 12).Value = Me.DDEV
dataSheet.Cells(lr + 1, 13).Value = Me.DFEC
dataSheet.Cells(lr + 1, 14).Value = Me.DOBS
End Sub
Private Sub SCAN_Click()
Unload Me
Sheets("PPrincipal").Select
End Sub
'CODE TO SEARCH ALL MATCHES
Private Sub DSEARCH_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Calls the FindAllMatches routine as user types text in the textbox
    Call FindallMatches
End Sub
Private Sub DSEARCH_Results_Click()
Dim ValuetoFind As String
Dim l As Long
    For l = 0 To Me.DSEARCH_Results.ListCount
        If Me.DSEARCH_Results.Selected(l) = True And Me.DSEARCH_Results.Value <> "Sin Resultados" Then
            On Error Resume Next
            strAddress = Me.DSEARCH_Results.List(l, 1)
            ValueTFind = dataSheet.Range(strAddress).Value
            FindValue = dataSheet.Range(strAddress).Row
            GoTo EndLoop
        End If
    Next l
EndLoop:
Me.DDATE.Value = dataSheet.Cells(FindValue, 5).Value
Me.DEMP.Value = Me.DSEARCH.Value
Me.DENT.Value = dataSheet.Cells(FindValue, 3).Value
Me.DSOL.Value = dataSheet.Cells(FindValue, 11).Value
Me.DFEC.Value = dataSheet.Cells(FindValue, 13).Value
Me.DDEV.Value = dataSheet.Cells(FindValue, 12).Value
Me.DOBS.Value = dataSheet.Cells(FindValue, 14).Value
End Sub
Private Sub SUPD_Click()
Dim msg As String
msg = "Diligencie Campo"
If Me.DSEARCH.Value <> "" Then
    Me.DSEARCH.Value = ""
End If
If Me.DSEARCH_Results.Value <> "" Then
    Me.DSEARCH_Results.Clear
End If
If Me.DDATE.Value = "" Then
    MsgBox msg
    Me.DDATE.SetFocus
    Exit Sub
End If
If Me.DEMP.Value = "" Then
    MsgBox msg
    Me.DEMP.SetFocus
    Exit Sub
End If
If Me.DENT.Value = "" Then
    MsgBox msg
    Me.DENT.SetFocus
    Exit Sub
End If
If Me.DSOL.Value = "" Then
    MsgBox msg
    Me.DSOL.SetFocus
    Exit Sub
End If
If Me.DOBS.Value = "" Then
    MsgBox msg
    Me.DOBS.SetFocus
    Exit Sub
End If
dataSheet.Cells(FindValue, 1).Value = Me.DEMP
dataSheet.Cells(FindValue, 2).Value = 2
dataSheet.Cells(FindValue, 3).Value = 3
dataSheet.Cells(FindValue, 4).Value = 4
dataSheet.Cells(FindValue, 5).Value = Me.DDATE
dataSheet.Cells(FindValue, 6).Value = 6
dataSheet.Cells(FindValue, 7).Value = Me.DENT
dataSheet.Cells(FindValue, 8).Value = 8
dataSheet.Cells(FindValue, 9).Value = 9
dataSheet.Cells(FindValue, 10).Value = 10
dataSheet.Cells(FindValue, 11).Value = Me.DSOL
dataSheet.Cells(FindValue, 12).Value = Me.DDEV
dataSheet.Cells(FindValue, 13).Value = Me.DFEC
dataSheet.Cells(FindValue, 14).Value = Me.DOBS
End Sub
Sub FindallMatches()
'Find all matches on activesheet
'Called by: TextBox_Find_KeyUp event
Dim SearchRange As Range, FindWhat As Variant, FoundCells As Range
Dim FoundCell As Range, arrResults() As Variant, lFound As Long
Dim lSearchCol As Long, lLastRow As Long
    If Len(DSEARCH.Value) > 1 Then 'Do search if text in find box is longer than 1 character.
        Set SearchRange = dataSheet.Range("A:A")
        FindWhat = Me.DSEARCH.Value
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
        Me.DSEARCH_Results.List = arrResults
    Else
        Me.DSEARCH_Results.Clear
    End If
End Sub
'*****************************End Find Data
