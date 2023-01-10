VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "Login"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3555
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Counter As Long
Private Sub LOGINI_Click()
If Me.USER.Value = "SUPERUSER" And Me.PASSWORD.Value = "SCP2810" Then
Unload Me
Exit Sub
Else
MsgBox "Usuario y/o Contraseña Incorrecta"
Counter = Counter + 1
MsgBox Counter
Me.USER.SetFocus
End If
End Sub
Private Sub LOGOUT_Click()
DataFLogin = True
ThisWorkbook.Close True
End Sub
Private Sub PASSWORD_Change()
Me.PASSWORD.Value = UCase(Me.PASSWORD.Value)
End Sub
Private Sub USER_Change()
Me.USER.Value = UCase(Me.USER.Value)
End Sub
Private Sub UserForm_QueryClose(CANCEL As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        CANCEL = True
        MsgBox "Ingrese Datos", vbCritical
    End If
End Sub
