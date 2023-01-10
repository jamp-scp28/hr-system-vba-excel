VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindC 
   Caption         =   "BUSCAR CODIGO CIE10"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4245
   OleObjectBlob   =   "FindC.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FINDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Sub TextBox_Find_Change()
'Go to selection on sheet when result is clicked
Dim myRange As Range
Set myRange = Worksheets("C_CIE10").Range("A:B")

Dim lRow As Long

On Error Resume Next

'VlookUp the values of the boxes

FINDC.ShowCD.Value = _
Application.WorksheetFunction.VLookup(TextBox_Find, myRange, 2, False)
If Err.Number <> 0 Then FINDC.ShowCD.Value = "No Results"

End Sub

Private Sub TextBox_Find_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Calls the FindAllMatches routine as user types text in the textbox
    Call FindallMatches
End Sub

Private Sub Label_ClearFind_Click()
'Clears the find text box and sets focus

    Me.TextBox_Find.Text = ""
    Me.TextBox_Find.SetFocus
    
End Sub

Sub FindallMatches()
'Find all matches on activesheet
'Called by: TextBox_Find_KeyUp event

Dim SearchRange As Range
Dim FindWhat As Variant
Dim FoundCells As Range
Dim FoundCell As Range
Dim arrResults() As Variant
Dim lFound As Long
Dim lSearchCol As Long
Dim lLastRow As Long
   
    If Len(TextBox_Find.Value) > 1 Then 'Do search if text in find box is longer than 1 character.
        
        Set SearchRange = Sheets("C_CIE10").Range("A:A")
        
        FindWhat = FINDC.TextBox_Find.Value
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
        Me.ListBox_Results.List = arrResults
        
    Else
        Me.ListBox_Results.Clear
    End If
        
End Sub

Private Sub ListBox_Results_Click()
Dim strAddress As String
strAddress = ListBox_Results.Value

TextBox_Find.Value = strAddress


'Dim l As Long

    'For l = 0 To ListBox_Results.ListCount
        'If ListBox_Results.Selected(l) = True Then
           ' strAddress = ListBox_Results.List(l, 1)
            'ActiveSheet.Range(strAddress).Select
            'GoTo EndLoop
       ' End If
    'Next l

'EndLoop:
    
End Sub

Private Sub ListBox_Results_DblClick(ByVal CANCEL As MSForms.ReturnBoolean)
Dim strAddress As String
strAddress = ListBox_Results.Value
AbsDescription = Me.ShowCD.Value
RAusentismos.VC_Selected.Value = strAddress

Unload Me
RAusentismos.FECHAI.SetFocus
End Sub

Private Sub UserForm_Click()

End Sub
