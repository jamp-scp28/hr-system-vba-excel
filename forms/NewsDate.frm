VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewsDate 
   Caption         =   "FECHA DE NOVEDAD"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "NewsDate.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "NewsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
Me.NDateb.SetFocus
End Sub
Private Sub NDateb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyReturn Then
    
    
    NewsDateD = Me.NDateb
    Unload Me
End If
If KeyCode.Value = vbKeyBack Then
    Me.NDateb.Value = vbNullString
End If
End Sub
Private Sub NDateb_Change()
If Me.NDateb.TextLength > 1 And Me.NDateb.TextLength < 3 Then
    Me.NDateb.Value = Me.NDateb.Value + "/"
End If
If Me.NDateb.TextLength > 4 And Me.NDateb.TextLength < 6 Then
    Me.NDateb.Value = Me.NDateb.Value + "/"
End If
If Me.NDateb.TextLength > 10 Then
    Me.NDateb.Value = Mid(Me.NDateb, 1, Len(Me.NDateb.Text) - 1)
End If
End Sub

