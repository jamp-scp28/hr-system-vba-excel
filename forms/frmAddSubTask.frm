VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddSubTask 
   Caption         =   "Add a Sub Task"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   OleObjectBlob   =   "frmAddSubTask.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmAddSubTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCal_Click()
sSender = "frmAddSubTask"
frmCal.Show
End Sub

Private Sub btnClose_Click()
Me.Hide
Unload frmAddSubTask
End Sub



Private Sub UserForm_Activate()

Me.cbStatus.AddItem "Not Started"
Me.cbStatus.AddItem "In Progress"
Me.cbStatus.AddItem "Frozen"
Me.cbStatus.AddItem "Trash"

End Sub

Private Sub btnAddTask_Click()

Dim oConn As New ADODB.Connection
Dim ssql As String, sConn As String, sPath As String
Dim sMisc As String, sMisc2 As String, myvar As Variant

'**************************************************************
'test all fields are completed correctly ..
    ' Due Date ...
    If Me.txtDueDate.Text = "" Then
    MsgBox "You need to add a due date", vbCritical, "subtask form not complete"
    Exit Sub
    ElseIf IsDate(Me.txtDueDate.Text) = False Then
    MsgBox "You need to add a due date in the following format (MM/DD/YYYY)", vbCritical, "subtask form not complete"
    Exit Sub
    Else
    End If
   
    
    ' description ...
    If Me.txtDesc.Text = "" Then
    MsgBox "You need to add a subtask description", vbCritical, "subtask form not complete"
    Exit Sub
    End If

'**************************************************************

sPath = ActiveWorkbook.Path & "\ToDo.accdb"

sConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & sPath & ";" & _
"Persist Security Info=False;"
oConn.Open sConn

myvar = Split(Me.lblsubTaskNo.Caption)
sMisc2 = CStr(myvar(UBound(myvar)))
sMisc = CStr(myvar(2))

If Me.Caption <> "Edit a SubTask" Then
    'code to add another subtask ..
    ssql = "INSERT INTO SubTasks (SubTaskNb, TaskNb, Date_Created, Date_Due, " & _
    "Description, Status) Values('" & _
    sMisc2 & "', '" & sMisc & "',#" & Me.lblED.Caption & "#, #" & _
    Me.txtDueDate.Text & "#, '" & _
    Me.txtDesc.Text & "', '" & _
    Me.cbStatus.Value & "');"
Else
    ssql = "UPDATE Subtasks SET " & _
    "Date_Due = #" & Me.txtDueDate.Text & "#, " & _
    "Description = '" & Me.txtDesc.Text & "', " & _
    "Status = '" & Me.cbStatus.Value & "' " & _
    "Where TaskNb = '" & sMisc & "' AND " & _
    "SubTaskNb = '" & sMisc2 & "'"

End If


oConn.Execute ssql
oConn.Close

Me.Hide
Unload frmAddSubTask

End Sub

Private Sub cbStatus_Change()

If Me.cbStatus.Value = "Completed" Then
    Me.lblcdt.Visible = True
    Me.txtdtComp.Visible = True
    If Me.txtdtComp.Text = "" Then
    Me.txtdtComp.Text = Format(Now, "mm/dd/yy")
    End If
Else
Me.txtdtComp.Visible = False
Me.lblcdt.Visible = False
End If

End Sub

Private Sub cbDelete_Click()
Dim ans As Variant

Dim oConn As New ADODB.Connection
Dim ssql As String, sConn As String, sPath As String
Dim sMisc As String, myvar As Variant

sPath = ActiveWorkbook.Path & "\ToDo.accdb"

sConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & sPath & ";" & _
"Persist Security Info=False;"

oConn.Open sConn

ans = MsgBox("Are You sure you want to delete this sub Task", vbYesNo, "Delete this subtask?")

If ans = vbYes Then
sMisc = Me.lblsubTaskNo.Caption
myvar = Split(sMisc)
ssql = "DELETE FROM SubTasks WHERE TaskNb = '" & myvar(2) & "' " & _
"AND SubTaskNb = '" & myvar(3) & "'"
oConn.Execute ssql

' clean up ...
oConn.Close
Me.Hide
End If

End Sub
