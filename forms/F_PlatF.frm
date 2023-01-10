VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_PlatF 
   Caption         =   "GESTION DE PLATAFORMAS"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6885
   OleObjectBlob   =   "F_PlatF.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "F_PlatF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public wspf As Worksheet
Public CurrentRow As Long
Public Link As String
Public lastrow As Long
Private Sub UserForm_Initialize()
    '==DECLARE VARIABLES
Set wspf = Sheets("PlatF")
lastrow = wspf.Cells(Rows.Count, 1).End(xlUp).Row
'==ADD DATA TO THE COMBOBOX
N_PLAT.List = wspf.Range("A2:A" & lastrow).Value
'==DISABLED LINK BUTTON AND LINK TO AVOID ERR
Me.LINK_PLAT.Enabled = False
Me.UP_PLAT.Enabled = False
End Sub
'================================
Private Sub N_PLAT_Change()
'==DECLARE VARIABLES
Dim myRangeP As Range
Set myRangeP = wspf.Range("A:K")
'==DISPLAY IMAGES
If Me.N_PLAT.ListIndex > -2 Then
On Error Resume Next
Me.IMG_PLAT.PictureSizeMode = fmPictureSizeModeStretch
   Me.IMG_PLAT.Picture = LoadPicture(ThisWorkbook.Path & "\Img Plataformas\" & _
   Me.N_PLAT.Value & ".jpg")
   Else
    On Error Resume Next
    Me.IMG_PLAT.Picture = LoadPicture(ThisWorkbook.Path & "\Fotos Colaboradores\" & "noimage.jpg")
End If
If Err Then
    Err.Clear
    Me.IMG_PLAT.Picture = LoadPicture(ThisWorkbook.Path & "\Fotos Colaboradores\" & "noimage.jpg")
End If
'==FIND DATA ON VALUE SELETEC IN N_PLAT
On Error Resume Next
CurrentRow = wspf.UsedRange.Find(What:=Me.N_PLAT, after:=wspf.Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
'==ASSIGN CURRENTDATA TO COMBOBOX
T_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 2, False)
    If Err.Number <> 0 Then T_PLAT.Value = ""
A_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 3, False)
    If Err.Number <> 0 Then A_PLAT.Value = ""
E_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 4, False)
    If Err.Number <> 0 Then E_PLAT.Value = ""
TD1_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 5, False)
    If Err.Number <> 0 Then TD1_PLAT.Value = ""
U1_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 6, False)
    If Err.Number <> 0 Then U1_PLAT.Value = ""
TD2_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 7, False)
    If Err.Number <> 0 Then TD2_PLAT.Value = ""
U2_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 8, False)
    If Err.Number <> 0 Then U2_PLAT.Value = ""
PAS_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 9, False)
    If Err.Number <> 0 Then PAS_PLAT.Value = ""
O_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 10, False)
    If Err.Number <> 0 Then O_PLAT.Value = ""
NLINK_PLAT.Value = _
    Application.WorksheetFunction.VLookup(Me.N_PLAT, myRangeP, 11, False)
    If Err.Number <> 0 Then NLINK_PLAT.Value = ""
'==CALL LINK
Link = wspf.Cells(CurrentRow, 11)
'==DISABLED LINK_PLAT TO AVOID ERR
If Me.N_PLAT.Value <> vbNullString And Me.N_PLAT.ListIndex > -2 Then
    Me.LINK_PLAT.Enabled = True
    Me.NLINK_PLAT.Enabled = False
    Me.NN_PLAT.Enabled = False
    Me.UP_PLAT.Enabled = True
    Me.NR_PLAT.Enabled = False
    Else
    Me.LINK_PLAT.Enabled = False
    Me.NLINK_PLAT.Enabled = True
    Me.NN_PLAT.Enabled = True
    Me.UP_PLAT.Enabled = False
    Me.NR_PLAT.Enabled = True
End If

End Sub
'=========================================
Private Sub LINK_PLAT_Click()
'==GO TO WEBSITE OF THE PLAFORM ON CALLED LINK
ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
End Sub
'=============================
Private Sub UP_PLAT_Click()
'==UPDATE DATA
wspf.Cells(CurrentRow, 1).Value = N_PLAT.Value
wspf.Cells(CurrentRow, 2).Value = T_PLAT.Value
wspf.Cells(CurrentRow, 3).Value = A_PLAT.Value
wspf.Cells(CurrentRow, 4).Value = E_PLAT.Value
wspf.Cells(CurrentRow, 5).Value = TD1_PLAT.Value
wspf.Cells(CurrentRow, 6).Value = U1_PLAT.Value
wspf.Cells(CurrentRow, 7).Value = TD2_PLAT.Value
wspf.Cells(CurrentRow, 8).Value = U2_PLAT.Value
wspf.Cells(CurrentRow, 9).Value = PAS_PLAT.Value
wspf.Cells(CurrentRow, 10).Value = O_PLAT.Value
wspf.Cells(CurrentRow, 11).Value = NLINK_PLAT.Value
MsgBox ("Datos actualizados correctamente")
End Sub
'============================
Private Sub NR_PLAT_Click()
'==AVOID EMPTY CELLS
Dim msg As String
msg = "Diligencie el campo" 'MSGBOX
If NN_PLAT.Value = vbNullString Then
    MsgBox msg
    NN_PLAT.SetFocus
    Exit Sub
End If
If T_PLAT.Value = vbNullString Then
    MsgBox msg
    T_PLAT.SetFocus
    Exit Sub
End If
If A_PLAT.Value = vbNullString Then
    MsgBox msg
    A_PLAT.SetFocus
    Exit Sub
End If
If E_PLAT.Value = vbNullString Then
    MsgBox msg
    E_PLAT.SetFocus
    Exit Sub
End If
If TD1_PLAT.Value = vbNullString Then
    MsgBox msg
    TD1_PLAT.SetFocus
    Exit Sub
End If
If U1_PLAT.Value = vbNullString Then
    MsgBox msg
    U1_PLAT.SetFocus
    Exit Sub
End If
If TD2_PLAT.Value = vbNullString Then
    MsgBox msg
    TD2_PLAT.SetFocus
    Exit Sub
End If
If U2_PLAT.Value = vbNullString Then
    MsgBox msg
    U2_PLAT.SetFocus
    Exit Sub
End If
If PAS_PLAT.Value = vbNullString Then
    MsgBox msg
    PAS_PLAT.SetFocus
    Exit Sub
End If
If O_PLAT.Value = vbNullString Then
    MsgBox msg
    O_PLAT.SetFocus
    Exit Sub
End If
If NLINK_PLAT.Value = vbNullString Then
    MsgBox msg
    NLINK_PLAT.SetFocus
    Exit Sub
End If
'==REGISTER THE DATA IN THE CELLS
wspf.Cells(lastrow + 1, 1).Value = NN_PLAT.Value
wspf.Cells(lastrow + 1, 2).Value = T_PLAT.Value
wspf.Cells(lastrow + 1, 3).Value = A_PLAT.Value
wspf.Cells(lastrow + 1, 4).Value = E_PLAT.Value
wspf.Cells(lastrow + 1, 5).Value = TD1_PLAT.Value
wspf.Cells(lastrow + 1, 6).Value = U1_PLAT.Value
wspf.Cells(lastrow + 1, 7).Value = TD2_PLAT.Value
wspf.Cells(lastrow + 1, 8).Value = U2_PLAT.Value
wspf.Cells(lastrow + 1, 9).Value = PAS_PLAT.Value
wspf.Cells(lastrow + 1, 10).Value = O_PLAT.Value
wspf.Cells(lastrow + 1, 11).Value = NLINK_PLAT.Value
'==FINAL MSG
If MsgBox("Datos ingresados correctamente, desea permanecer en el formulario", vbYesNo) = vbYes Then
    N_PLAT.SetFocus
    Call UserForm_Initialize
    Call EmptyBox
    Else
    Call EmptyBox
    Unload Me
End If
End Sub
'===================================
Private Sub CAN_PLAT_Click()
Unload Me
End Sub
'====================================
Sub EmptyBox()
'==CLEAR DATA IN THE BOXES
NN_PLAT.Value = vbNullString
T_PLAT.Value = vbNullString
A_PLAT.Value = vbNullString
E_PLAT.Value = vbNullString
TD1_PLAT.Value = vbNullString
U1_PLAT.Value = vbNullString
TD2_PLAT.Value = vbNullString
U2_PLAT.Value = vbNullString
PAS_PLAT.Value = vbNullString
O_PLAT.Value = vbNullString
NLINK_PLAT.Value = vbNullString
End Sub

