VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SDocumentacion 
   Caption         =   "DOCUMENTACIÓN"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16260
   OleObjectBlob   =   "SDocumentacion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "SDocumentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActualizarD_Click()
'Variables for find values
 Dim lRow As Long
 Dim wsD As Worksheet
 Set wsD = Sheets("DData")
 
lRow = Sheets("DData").UsedRange.Find(What:=Me.BUSCADORD, after:=Sheets("DData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row

Dim CurrentRow As Long
CurrentRow = lRow


'Code to get the data from the cells and in button click replace it for the info in the texbox on form

wsD.Cells(CurrentRow, DData.[doc_state].Column).Value = FENTREVISTA.Value
wsD.Cells(CurrentRow + 1, DData.[doc_state].Column).Value = HV.Value
wsD.Cells(CurrentRow + 2, DData.[doc_state].Column).Value = CC.Value
wsD.Cells(CurrentRow + 3, DData.[doc_state].Column).Value = ANTECEDENTES.Value
wsD.Cells(CurrentRow + 4, DData.[doc_state].Column).Value = SESTUDIOS.Value
wsD.Cells(CurrentRow + 5, DData.[doc_state].Column).Value = CLABORALES.Value
wsD.Cells(CurrentRow + 6, DData.[doc_state].Column).Value = RPERSONALES.Value
wsD.Cells(CurrentRow + 7, DData.[doc_state].Column).Value = CSEGURIDAD.Value
wsD.Cells(CurrentRow + 8, DData.[doc_state].Column).Value = ASEGURIDAD.Value
wsD.Cells(CurrentRow + 9, DData.[doc_state].Column).Value = DFAMILIAR.Value
wsD.Cells(CurrentRow + 10, DData.[doc_state].Column).Value = EINGRESO.Value
wsD.Cells(CurrentRow + 11, DData.[doc_state].Column).Value = DCONTRATO.Value


wsD.Cells(CurrentRow + 12, DData.[doc_state].Column).Value = Me.AEXTT.Value
wsD.Cells(CurrentRow + 13, DData.[doc_state].Column).Value = Me.OSS.Value
wsD.Cells(CurrentRow + 14, DData.[doc_state].Column).Value = Me.OSDD.Value
wsD.Cells(CurrentRow + 15, DData.[doc_state].Column).Value = Me.CCDD.Value
wsD.Cells(CurrentRow + 16, DData.[doc_state].Column).Value = Me.DPRR.Value

wsD.Cells(CurrentRow + 17, DData.[doc_state].Column).Value = Me.DACCOUNT.Value
wsD.Cells(CurrentRow + 18, DData.[doc_state].Column).Value = AINDUCCION.Value
wsD.Cells(CurrentRow + 19, DData.[doc_state].Column).Value = MEMORANDOS.Value
wsD.Cells(CurrentRow + 20, DData.[doc_state].Column).Value = EDESEMPEÑO.Value
wsD.Cells(CurrentRow + 21, DData.[doc_state].Column).Value = PAZYSALVO.Value
wsD.Cells(CurrentRow + 22, DData.[doc_state].Column).Value = EEGRESO.Value
wsD.Cells(CurrentRow + 23, DData.[doc_state].Column).Value = LIQUIDACION.Value

'Code to get the data from the cells and in button click replace it for the info in the texbox on form observation document

wsD.Cells(CurrentRow, DData.[doc_observation].Column).Value = OFENTREVISTA.Value
wsD.Cells(CurrentRow + 1, DData.[doc_observation].Column).Value = OHV.Value
wsD.Cells(CurrentRow + 2, DData.[doc_observation].Column).Value = OCC.Value
wsD.Cells(CurrentRow + 3, DData.[doc_observation].Column).Value = OANTECEDENTES.Value
wsD.Cells(CurrentRow + 4, DData.[doc_observation].Column).Value = OSESTUDIOS.Value
wsD.Cells(CurrentRow + 5, DData.[doc_observation].Column).Value = OCLABORALES.Value
wsD.Cells(CurrentRow + 6, DData.[doc_observation].Column).Value = ORPERSONALES.Value
wsD.Cells(CurrentRow + 7, DData.[doc_observation].Column).Value = OCSEGURIDAD.Value
wsD.Cells(CurrentRow + 8, DData.[doc_observation].Column).Value = OASEGURIDAD.Value
wsD.Cells(CurrentRow + 9, DData.[doc_observation].Column).Value = ODFAMILIAR.Value
wsD.Cells(CurrentRow + 10, DData.[doc_observation].Column).Value = OEINGRESO.Value
wsD.Cells(CurrentRow + 11, DData.[doc_observation].Column).Value = ODCONTRATO.Value

wsD.Cells(CurrentRow + 12, DData.[doc_observation].Column).Value = Me.OAEXT.Value
wsD.Cells(CurrentRow + 13, DData.[doc_observation].Column).Value = Me.OOS.Value
wsD.Cells(CurrentRow + 14, DData.[doc_observation].Column).Value = Me.OOSD.Value
wsD.Cells(CurrentRow + 15, DData.[doc_observation].Column).Value = Me.OCCD.Value
wsD.Cells(CurrentRow + 16, DData.[doc_observation].Column).Value = Me.ODPR.Value

wsD.Cells(CurrentRow + 17, DData.[doc_observation].Column).Value = Me.DOACCOUNT.Value
wsD.Cells(CurrentRow + 18, DData.[doc_observation].Column).Value = OAINDUCCION.Value
wsD.Cells(CurrentRow + 19, DData.[doc_observation].Column).Value = OMEMORANDOS.Value
wsD.Cells(CurrentRow + 20, DData.[doc_observation].Column).Value = OEDESEMPEÑO.Value
wsD.Cells(CurrentRow + 21, DData.[doc_observation].Column).Value = OPAZYSALVO.Value
wsD.Cells(CurrentRow + 22, DData.[doc_observation].Column).Value = OEEGRESO.Value
wsD.Cells(CurrentRow + 23, DData.[doc_observation].Column).Value = OLIQUIDACION.Value

If MsgBox("Datos actualizados, desea continuar en el formulario?", vbYesNo) = vbYes Then

Else

End If
    
End Sub

Private Sub BUSCADORD_Change()
'==DISABLED BUTTON TO REGISTER NEW DATA
If BUSCADORD.ListIndex > -1 Then
    Me.ActualizarD.Enabled = True
    Else
    Me.ActualizarD.Enabled = False
End If
'allow button actualizar data
ActualizarD.Enabled = True
'Set reference to look up the concept of every document
Dim myRange As Range
Set myRange = Worksheets("DData").Range("D:G")
On Error Resume Next

'VlookUp the values of the boxes
'______________________________________________________________________________
'Declare the value of combobox + the document to be search in every box in state of document

BFE = BUSCADORD.Value & "-" & "Formato Entrevista"
BHV = BUSCADORD.Value & "-" & "Hoja de Vida"
BCC = BUSCADORD.Value & "-" & "Fotocopia Documento Identidad"
BA = BUSCADORD.Value & "-" & "Antecedentes"
BSE = BUSCADORD.Value & "-" & "Soportes de Estudio"
BCLL = BUSCADORD.Value & "-" & "Certificaciones Laborales"
BRP = BUSCADORD.Value & "-" & "Referencias Personales"
BCS = BUSCADORD.Value & "-" & "Certificacion EPS-AFP"
BAS = BUSCADORD.Value & "-" & "Afiliación EPS-ARL-AFP-CCF"
BF = BUSCADORD.Value & "-" & "Documentos Conyuge e Hijos"
BEI = BUSCADORD.Value & "-" & "Examen Ingreso"
BCL = BUSCADORD.Value & "-" & "Contrato"
ODAC = BUSCADORD.Value & "-" & "Cuenta Bancaria"
BAI = BUSCADORD.Value & "-" & "Acta de Inducción"
BM = BUSCADORD.Value & "-" & "Memorandos"
BED = BUSCADORD.Value & "-" & "Evaluaciones de Desempeño"
BPZ = BUSCADORD.Value & "-" & "Paz y Salvo"
BEE = BUSCADORD.Value & "-" & "Examen de Egreso"
BL = BUSCADORD.Value & "-" & "Liquidacion"
AEXT = BUSCADORD.Value & "-" & "Auxilio Extralegal"
OS = BUSCADORD.Value & "-" & "Otro Sí"
OSD = BUSCADORD.Value & "-" & "Otro Sí Datos Personales"
CCD = BUSCADORD.Value & "-" & "Compromiso Confidencialidad"
DPR = BUSCADORD.Value & "-" & "Documento de Precisión y Ratificación"

'Look up code to get the data from the sheet
On Error Resume Next
FENTREVISTA.Value = _
    Application.WorksheetFunction.VLookup(BFE, myRange, 2, False)
    If Err.Number <> 0 Then FENTREVISTA.Value = vbNullString

HV.Value = _
    Application.WorksheetFunction.VLookup(BHV, myRange, 2, False)
    If Err.Number <> 0 Then HV.Value = vbNullString

CC.Value = _
    Application.WorksheetFunction.VLookup(BCC, myRange, 2, False)
    If Err.Number <> 0 Then CC.Value = vbNullString

ANTECEDENTES.Value = _
    Application.WorksheetFunction.VLookup(BA, myRange, 2, False)
    If Err.Number <> 0 Then ANTECEDENTES.Value = vbNullString

SESTUDIOS.Value = _
    Application.WorksheetFunction.VLookup(BSE, myRange, 2, False)
    If Err.Number <> 0 Then SESTUDIOS.Value = vbNullString

CLABORALES.Value = _
    Application.WorksheetFunction.VLookup(BCLL, myRange, 2, False)
    If Err.Number <> 0 Then CLABORALES.Value = vbNullString

RPERSONALES.Value = _
    Application.WorksheetFunction.VLookup(BRP, myRange, 2, False)
    If Err.Number <> 0 Then RPERSONALES.Value = vbNullString

CSEGURIDAD.Value = _
    Application.WorksheetFunction.VLookup(BCS, myRange, 2, False)
    If Err.Number <> 0 Then CSEGURIDAD.Value = vbNullString

ASEGURIDAD.Value = _
    Application.WorksheetFunction.VLookup(BAS, myRange, 2, False)
    If Err.Number <> 0 Then ASEGURIDAD.Value = vbNullString

DFAMILIAR.Value = _
    Application.WorksheetFunction.VLookup(BF, myRange, 2, False)
    If Err.Number <> 0 Then DFAMILIAR.Value = vbNullString

EINGRESO.Value = _
    Application.WorksheetFunction.VLookup(BEI, myRange, 2, False)
    If Err.Number <> 0 Then EINGRESO.Value = vbNullString

DCONTRATO.Value = _
    Application.WorksheetFunction.VLookup(BCL, myRange, 2, False)
    If Err.Number <> 0 Then DCONTRATO.Value = vbNullString
'======

OSS.Value = _
    Application.WorksheetFunction.VLookup(OS, myRange, 2, False)
    If Err.Number <> 0 Then OS.Value = vbNullString

OSDD.Value = _
    Application.WorksheetFunction.VLookup(OSD, myRange, 2, False)
    If Err.Number <> 0 Then OSD.Value = vbNullString
    
AEXTT.Value = _
    Application.WorksheetFunction.VLookup(AEXT, myRange, 2, False)
    If Err.Number <> 0 Then AEXT.Value = vbNullString

CCDD.Value = _
    Application.WorksheetFunction.VLookup(CCD, myRange, 2, False)
    If Err.Number <> 0 Then CCD.Value = vbNullString

DPRR.Value = _
    Application.WorksheetFunction.VLookup(DPR, myRange, 2, False)
    If Err.Number <> 0 Then DPR.Value = vbNullString
'================
Me.DACCOUNT.Value = _
    Application.WorksheetFunction.VLookup(ODAC, myRange, 2, False)
    If Err.Number <> 0 Then Me.DACCOUNT.Value = vbNullString

AINDUCCION.Value = _
    Application.WorksheetFunction.VLookup(BAI, myRange, 2, False)
    If Err.Number <> 0 Then AINDUCCION.Value = vbNullString

MEMORANDOS.Value = _
    Application.WorksheetFunction.VLookup(BM, myRange, 2, False)
    If Err.Number <> 0 Then MEMORANDOS.Value = vbNullString

EDESEMPEÑO.Value = _
    Application.WorksheetFunction.VLookup(BED, myRange, 2, False)
    If Err.Number <> 0 Then EDESEMPEÑO.Value = vbNullString

PAZYSALVO.Value = _
    Application.WorksheetFunction.VLookup(BPZ, myRange, 2, False)
    If Err.Number <> 0 Then PAZYSALVO.Value = vbNullString

EEGRESO.Value = _
    Application.WorksheetFunction.VLookup(BEE, myRange, 2, False)
    If Err.Number <> 0 Then EEGRESO.Value = vbNullString

LIQUIDACION.Value = _
    Application.WorksheetFunction.VLookup(BL, myRange, 2, False)
    If Err.Number <> 0 Then LIQUIDACION.Value = vbNullString
'________________________________________________________________________

'Vlooup function to get the data from the sheet observation of document

OFENTREVISTA.Value = _
    Application.WorksheetFunction.VLookup(BFE, myRange, 3, False)
    If Err.Number <> 0 Then OFENTREVISTA.Value = "NA"

OHV.Value = _
    Application.WorksheetFunction.VLookup(BHV, myRange, 3, False)
    If Err.Number <> 0 Then OHV.Value = "NA"

OCC.Value = _
    Application.WorksheetFunction.VLookup(BCC, myRange, 3, False)
    If Err.Number <> 0 Then OCC.Value = "NA"

OANTECEDENTES.Value = _
    Application.WorksheetFunction.VLookup(BA, myRange, 3, False)
    If Err.Number <> 0 Then OANTECEDENTES.Value = "NA"

OSESTUDIOS.Value = _
    Application.WorksheetFunction.VLookup(BSE, myRange, 3, False)
    If Err.Number <> 0 Then OSESTUDIOS.Value = "NA"

OCLABORALES.Value = _
    Application.WorksheetFunction.VLookup(BCLL, myRange, 3, False)
    If Err.Number <> 0 Then OCLABORALES.Value = "NA"

ORPERSONALES.Value = _
    Application.WorksheetFunction.VLookup(BRP, myRange, 3, False)
    If Err.Number <> 0 Then ORPERSONALES.Value = "NA"

OCSEGURIDAD.Value = _
    Application.WorksheetFunction.VLookup(BCS, myRange, 3, False)
    If Err.Number <> 0 Then OCSEGURIDAD.Value = "NA"

OASEGURIDAD.Value = _
    Application.WorksheetFunction.VLookup(BAS, myRange, 3, False)
    If Err.Number <> 0 Then OASEGURIDAD.Value = "NA"

ODFAMILIAR.Value = _
    Application.WorksheetFunction.VLookup(BF, myRange, 3, False)
    If Err.Number <> 0 Then ODFAMILIAR.Value = "NA"

OEINGRESO.Value = _
    Application.WorksheetFunction.VLookup(BEI, myRange, 3, False)
    If Err.Number <> 0 Then OEINGRESO.Value = "NA"

ODCONTRATO.Value = _
    Application.WorksheetFunction.VLookup(BCL, myRange, 3, False)
    If Err.Number <> 0 Then ODCONTRATO.Value = "NA"

Me.DOACCOUNT.Value = _
    Application.WorksheetFunction.VLookup(ODAC, myRange, 3, False)
    If Err.Number <> 0 Then Me.DOACCOUNT.Value = "NA"

OAINDUCCION.Value = _
    Application.WorksheetFunction.VLookup(BAI, myRange, 3, False)
    If Err.Number <> 0 Then OAINDUCCION.Value = "NA"

OMEMORANDOS.Value = _
    Application.WorksheetFunction.VLookup(BM, myRange, 3, False)
    If Err.Number <> 0 Then OMEMORANDOS.Value = "NA"

OEDESEMPEÑO.Value = _
    Application.WorksheetFunction.VLookup(BED, myRange, 3, False)
    If Err.Number <> 0 Then OEDESEMPEÑO.Value = "NA"

OPAZYSALVO.Value = _
    Application.WorksheetFunction.VLookup(BPZ, myRange, 3, False)
    If Err.Number <> 0 Then OPAZYSALVO.Value = "NA"

OEEGRESO.Value = _
    Application.WorksheetFunction.VLookup(BEE, myRange, 3, False)
    If Err.Number <> 0 Then OEEGRESO.Value = "NA"

OLIQUIDACION.Value = _
    Application.WorksheetFunction.VLookup(BL, myRange, 3, False)
    If Err.Number <> 0 Then OLIQUIDACION.Value = "NA"
'======
OOS.Value = _
    Application.WorksheetFunction.VLookup(OS, myRange, 3, False)
    If Err.Number <> 0 Then OS.Value = vbNullString

OOSD.Value = _
    Application.WorksheetFunction.VLookup(OSD, myRange, 3, False)
    If Err.Number <> 0 Then OSD.Value = vbNullString
    
OAEXT.Value = _
    Application.WorksheetFunction.VLookup(AEXT, myRange, 3, False)
    If Err.Number <> 0 Then AEXT.Value = vbNullString

OCCD.Value = _
    Application.WorksheetFunction.VLookup(CCD, myRange, 3, False)
    If Err.Number <> 0 Then CCD.Value = vbNullString

ODPR.Value = _
    Application.WorksheetFunction.VLookup(DPR, myRange, 3, False)
    If Err.Number <> 0 Then DPR.Value = vbNullString
'======
End Sub

Private Sub BUSCADORD_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
End If
End Sub

Private Sub CANCEL_Click()
Unload Me
Sheets("PPrincipal").Select
End Sub

Private Sub NINGRESO_Click()
'Validation: Do not allow blank spaces

If BUSCADORD.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    BUSCADORD.SetFocus
    BUSCADORD.DropDown
Exit Sub
End If


If FENTREVISTA.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    FENTREVISTA.SetFocus
    FENTREVISTA.DropDown
Exit Sub
End If

If HV.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    HV.SetFocus
    HV.DropDown
Exit Sub
End If

If CC.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    CC.SetFocus
    CC.DropDown
Exit Sub
End If

If ANTECEDENTES.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    ANTECEDENTES.SetFocus
    ANTECEDENTES.DropDown
Exit Sub
End If

If SESTUDIOS.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    SESTUDIOS.SetFocus
    SESTUDIOS.DropDown
Exit Sub
End If

If CLABORALES.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    CLABORALES.SetFocus
    CLABORALES.DropDown
Exit Sub
End If

If RPERSONALES.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    RPERSONALES.SetFocus
    RPERSONALES.DropDown
Exit Sub
End If

If CSEGURIDAD.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    CSEGURIDAD.SetFocus
    CSEGURIDAD.DropDown
Exit Sub
End If

If ASEGURIDAD.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    ASEGURIDAD.SetFocus
    ASEGURIDAD.DropDown
Exit Sub
End If

If DFAMILIAR.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    DFAMILIAR.SetFocus
    DFAMILIAR.DropDown
Exit Sub
End If

If EINGRESO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    EINGRESO.SetFocus
    EINGRESO.DropDown
Exit Sub
End If

If DCONTRATO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    DCONTRATO.SetFocus
    DCONTRATO.DropDown
Exit Sub
End If

If AINDUCCION.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    AINDUCCION.SetFocus
    AINDUCCION.DropDown
Exit Sub
End If

If MEMORANDOS.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    MEMORANDOS.SetFocus
    MEMORANDOS.DropDown
Exit Sub
End If

If EDESEMPEÑO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    EDESEMPEÑO.SetFocus
    EDESEMPEÑO.DropDown
Exit Sub
End If

If PAZYSALVO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    PAZYSALVO.SetFocus
    PAZYSALVO.DropDown
Exit Sub
End If

If EEGRESO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    EEGRESO.SetFocus
    EEGRESO.DropDown
Exit Sub
End If

If LIQUIDACION.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    LIQUIDACION.SetFocus
    LIQUIDACION.DropDown
Exit Sub
End If

'Do not allow blank spaces in observations

If OFENTREVISTA.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OFENTREVISTA.SetFocus
Exit Sub
End If

If OHV.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OHV.SetFocus
Exit Sub
End If

If OCC.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OCC.SetFocus
Exit Sub
End If

If OANTECEDENTES.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OANTECEDENTES.SetFocus
Exit Sub
End If

If OSESTUDIOS.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OSESTUDIOS.SetFocus
Exit Sub
End If

If OCLABORALES.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OCLABORALES.SetFocus
Exit Sub
End If

If ORPERSONALES.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    ORPERSONALES.SetFocus
Exit Sub
End If

If OCSEGURIDAD.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OCSEGURIDAD.SetFocus
Exit Sub
End If

If OASEGURIDAD.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OASEGURIDAD.SetFocus
Exit Sub
End If

If ODFAMILIAR.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    ODFAMILIAR.SetFocus
Exit Sub
End If

If OEINGRESO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OEINGRESO.SetFocus
Exit Sub
End If

If ODCONTRATO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    ODCONTRATO.SetFocus
Exit Sub
End If

If OAINDUCCION.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OAINDUCCION.SetFocus
Exit Sub
End If

If OMEMORANDOS.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OMEMORANDOS.SetFocus
Exit Sub
End If

If OEDESEMPEÑO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OEDESEMPEÑO.SetFocus
Exit Sub
End If

If OPAZYSALVO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OPAZYSALVO.SetFocus
Exit Sub
End If

If OEEGRESO.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OEEGRESO.SetFocus
Exit Sub
End If

If OLIQUIDACION.Value = vbNullString Then
    MsgBox "Ingrese Dato"
    OLIQUIDACION.SetFocus
Exit Sub
End If

'End of validation of fields (blank spaces)

Dim lastrow As Long
Dim lastr As Long
Dim wsdd As Worksheet
Dim Duplicado As Boolean

lastrow = Sheets("DData").Cells(Rows.Count, 1).End(xlUp).Row + 1

Duplicado = False

Dim Comprobar As String

Comprobar = BUSCADORD.Value & "-" & "Formato Entrevista"

For i = 1 To lastrow
    'If Application.WorksheetFunction.CountIf(Sheets("DData").Cells, Comprobar) > 0 Then
    If Sheets("DData").Cells(i, DData.[doc_merge].Column).Value = Comprobar Then
    Duplicado = True
    MsgBox "Datos Duplicados"
    Call DeleteDataD
    BUSCADORD.SetFocus
    Exit Sub
    End If
Next i

If Not Duplicado Then

Dim wsD As Worksheet
Set wsD = Sheets("DData")

'Assign style to the row
wsD.Cells(lastrow, 1).Style = "Énfasis5"
wsD.Cells(lastrow, DData.[doc_documentslist].Column).Style = "Énfasis5"
wsD.Cells(lastrow, DData.[doc_merge].Column).Style = "Énfasis5"
wsD.Cells(lastrow, DData.[doc_state].Column).Style = "Énfasis5"
wsD.Cells(lastrow, DData.[doc_observation].Column).Style = "Énfasis5"
'Repeat the name 18 times in the last row
Dim ncell As Range
Dim rangeName As Range

Dim r As Long
For r = 0 To 23
wsD.Cells(lastrow + r, 1).Value = Me.BUSCADORD.Value
wsD.Cells(lastrow + r, 2).Value = BUSCADORD.Value
Next r

'Add the name of the document to the adjacent cell

wsD.Cells(lastrow, DData.[doc_documentslist].Column).Value = "Formato Entrevista"
wsD.Cells(lastrow + 1, DData.[doc_documentslist].Column).Value = "Hoja de Vida"
wsD.Cells(lastrow + 2, DData.[doc_documentslist].Column).Value = "Fotocopia Documento Identidad"
wsD.Cells(lastrow + 3, DData.[doc_documentslist].Column).Value = "Antecedentes"
wsD.Cells(lastrow + 4, DData.[doc_documentslist].Column).Value = "Soportes de Estudio"
wsD.Cells(lastrow + 5, DData.[doc_documentslist].Column).Value = "Certificaciones Laborales"
wsD.Cells(lastrow + 6, DData.[doc_documentslist].Column).Value = "Referencias Personales"
wsD.Cells(lastrow + 7, DData.[doc_documentslist].Column).Value = "Certificacion EPS-AFP"
wsD.Cells(lastrow + 8, DData.[doc_documentslist].Column).Value = "Afiliación EPS-ARL-AFP-CCF"
wsD.Cells(lastrow + 9, DData.[doc_documentslist].Column).Value = "Documentos Conyuge e Hijos"
wsD.Cells(lastrow + 10, DData.[doc_documentslist].Column).Value = "Examen Ingreso"
wsD.Cells(lastrow + 11, DData.[doc_documentslist].Column).Value = "Contrato"

wsD.Cells(lastrow + 12, DData.[doc_documentslist].Column).Value = "Auxilio Extralegal"
wsD.Cells(lastrow + 13, DData.[doc_documentslist].Column).Value = "Otro Sí"
wsD.Cells(lastrow + 14, DData.[doc_documentslist].Column).Value = "Otro Sí Datos Personales"
wsD.Cells(lastrow + 15, DData.[doc_documentslist].Column).Value = "Compromiso Confidencialidad"
wsD.Cells(lastrow + 16, DData.[doc_documentslist].Column).Value = "Documento de Precisión y Ratificación"

wsD.Cells(lastrow + 17, DData.[doc_documentslist].Column).Value = "Cuenta Bancaria"
wsD.Cells(lastrow + 18, DData.[doc_documentslist].Column).Value = "Acta de Inducción"
wsD.Cells(lastrow + 19, DData.[doc_documentslist].Column).Value = "Memorandos"
wsD.Cells(lastrow + 20, DData.[doc_documentslist].Column).Value = "Evaluaciones de Desempeño"
wsD.Cells(lastrow + 21, DData.[doc_documentslist].Column).Value = "Paz y Salvo"
wsD.Cells(lastrow + 22, DData.[doc_documentslist].Column).Value = "Examen de Egreso"
wsD.Cells(lastrow + 23, DData.[doc_documentslist].Column).Value = "Liquidacion"

'Add data to the state of the document

wsD.Cells(lastrow, DData.[doc_state].Column).Value = FENTREVISTA.Value
wsD.Cells(lastrow + 1, DData.[doc_state].Column).Value = HV.Value
wsD.Cells(lastrow + 2, DData.[doc_state].Column).Value = CC.Value
wsD.Cells(lastrow + 3, DData.[doc_state].Column).Value = ANTECEDENTES.Value
wsD.Cells(lastrow + 4, DData.[doc_state].Column).Value = SESTUDIOS.Value
wsD.Cells(lastrow + 5, DData.[doc_state].Column).Value = CLABORALES.Value
wsD.Cells(lastrow + 6, DData.[doc_state].Column).Value = RPERSONALES.Value
wsD.Cells(lastrow + 7, DData.[doc_state].Column).Value = CSEGURIDAD.Value
wsD.Cells(lastrow + 8, DData.[doc_state].Column).Value = ASEGURIDAD.Value
wsD.Cells(lastrow + 9, DData.[doc_state].Column).Value = DFAMILIAR.Value
wsD.Cells(lastrow + 10, DData.[doc_state].Column).Value = EINGRESO.Value
wsD.Cells(lastrow + 11, DData.[doc_state].Column).Value = DCONTRATO.Value


wsD.Cells(lastrow + 12, DData.[doc_state].Column).Value = Me.AEXTT.Value
wsD.Cells(lastrow + 13, DData.[doc_state].Column).Value = Me.OSS.Value
wsD.Cells(lastrow + 14, DData.[doc_state].Column).Value = Me.OSDD.Value
wsD.Cells(lastrow + 15, DData.[doc_state].Column).Value = Me.CCDD.Value
wsD.Cells(lastrow + 16, DData.[doc_state].Column).Value = Me.DPRR.Value

wsD.Cells(lastrow + 17, DData.[doc_state].Column).Value = Me.DACCOUNT.Value
wsD.Cells(lastrow + 18, DData.[doc_state].Column).Value = AINDUCCION.Value
wsD.Cells(lastrow + 19, DData.[doc_state].Column).Value = MEMORANDOS.Value
wsD.Cells(lastrow + 20, DData.[doc_state].Column).Value = EDESEMPEÑO.Value
wsD.Cells(lastrow + 21, DData.[doc_state].Column).Value = PAZYSALVO.Value
wsD.Cells(lastrow + 22, DData.[doc_state].Column).Value = EEGRESO.Value
wsD.Cells(lastrow + 23, DData.[doc_state].Column).Value = LIQUIDACION.Value


'Concatenate data in the 3 column to search the data

Dim rr As Long
For rr = 0 To 23
    wsD.Cells(lastrow + rr, DData.[doc_merge].Column).FormulaR1C1 = "=CONCATENATE(RC[-2],""-"",RC[-1])"
Next rr

'Add data to the observation of the document

wsD.Cells(lastrow, DData.[doc_observation].Column).Value = OFENTREVISTA.Value
wsD.Cells(lastrow + 1, DData.[doc_observation].Column).Value = OHV.Value
wsD.Cells(lastrow + 2, DData.[doc_observation].Column).Value = OCC.Value
wsD.Cells(lastrow + 3, DData.[doc_observation].Column).Value = OANTECEDENTES.Value
wsD.Cells(lastrow + 4, DData.[doc_observation].Column).Value = OSESTUDIOS.Value
wsD.Cells(lastrow + 5, DData.[doc_observation].Column).Value = OCLABORALES.Value
wsD.Cells(lastrow + 6, DData.[doc_observation].Column).Value = ORPERSONALES.Value
wsD.Cells(lastrow + 7, DData.[doc_observation].Column).Value = OCSEGURIDAD.Value
wsD.Cells(lastrow + 8, DData.[doc_observation].Column).Value = OASEGURIDAD.Value
wsD.Cells(lastrow + 9, DData.[doc_observation].Column).Value = ODFAMILIAR.Value
wsD.Cells(lastrow + 10, DData.[doc_observation].Column).Value = OEINGRESO.Value
wsD.Cells(lastrow + 11, DData.[doc_observation].Column).Value = ODCONTRATO.Value

wsD.Cells(lastrow + 12, DData.[doc_observation].Column).Value = Me.OAEXT.Value
wsD.Cells(lastrow + 13, DData.[doc_observation].Column).Value = Me.OOS.Value
wsD.Cells(lastrow + 14, DData.[doc_observation].Column).Value = Me.OOSD.Value
wsD.Cells(lastrow + 15, DData.[doc_observation].Column).Value = Me.OCCD.Value
wsD.Cells(lastrow + 16, DData.[doc_observation].Column).Value = Me.ODPR.Value

wsD.Cells(lastrow + 17, DData.[doc_observation].Column).Value = Me.DOACCOUNT.Value
wsD.Cells(lastrow + 18, DData.[doc_observation].Column).Value = OAINDUCCION.Value
wsD.Cells(lastrow + 19, DData.[doc_observation].Column).Value = OMEMORANDOS.Value
wsD.Cells(lastrow + 20, DData.[doc_observation].Column).Value = OEDESEMPEÑO.Value
wsD.Cells(lastrow + 21, DData.[doc_observation].Column).Value = OPAZYSALVO.Value
wsD.Cells(lastrow + 22, DData.[doc_observation].Column).Value = OEEGRESO.Value
wsD.Cells(lastrow + 23, DData.[doc_observation].Column).Value = OLIQUIDACION.Value


Call DeleteDataD

If MsgBox("Datos ingresados, desea continuar en el formulario?", vbYesNo) = vbYes Then

Else
Unload Me
Sheets("PPrincipal").Select
End If
End If
End Sub
Private Sub UserForm_Initialize()
Application.EnableCancelKey = xlDisabled
'==DISABLED NINGRESO & UPDATE BUTTON TO AVOID ERROR
Me.ActualizarD.Enabled = False
'Load data to the BuscadoD
Dim lastrow
lastrow = Sheets("PData").Cells(Rows.Count, 1).End(xlUp).Row
BUSCADORD.List = Sheets("PData").Range("A2:A" & lastrow).Value
Dim OData As Variant
OData = Array("SÍ", "NO", "NO APLICA", "NO TIENE", "DESACTUALIZADO")
Dim ctrl As MSForms.control
    For Each ctrl In Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            If ctrl.Name Like "FENTREVISTA*" Then ctrl.List = OData
            If ctrl.Name Like "HV*" Then ctrl.List = OData
            If ctrl.Name Like "CC*" Then ctrl.List = OData
            If ctrl.Name Like "ANTECEDENTES*" Then ctrl.List = OData
            If ctrl.Name Like "SESTUDIOS*" Then ctrl.List = OData
            If ctrl.Name Like "CLABORALES*" Then ctrl.List = OData
            If ctrl.Name Like "RPERSONALES*" Then ctrl.List = OData
            If ctrl.Name Like "CSEGURIDAD*" Then ctrl.List = OData
            If ctrl.Name Like "ASEGURIDAD*" Then ctrl.List = OData
            If ctrl.Name Like "DFAMILIAR*" Then ctrl.List = OData
            If ctrl.Name Like "EINGRESO*" Then ctrl.List = OData
            If ctrl.Name Like "DCONTRATO*" Then ctrl.List = OData
            If ctrl.Name Like "AINDUCCION*" Then ctrl.List = OData
            If ctrl.Name Like "MEMORANDOS*" Then ctrl.List = OData
            If ctrl.Name Like "EDESEMPEÑO*" Then ctrl.List = OData
            If ctrl.Name Like "PAZYSALVO*" Then ctrl.List = OData
            If ctrl.Name Like "EEGRESO*" Then ctrl.List = OData
            If ctrl.Name Like "LIQUIDACION*" Then ctrl.List = OData
            If ctrl.Name Like "DACCOUNT*" Then ctrl.List = OData
        End If
    Next ctrl
End Sub

Sub DeleteDataD()
'Delete data from state of document
BUSCADORD.Value = vbNullString
FENTREVISTA.Value = vbNullString
HV.Value = vbNullString
CC.Value = vbNullString
ANTECEDENTES.Value = vbNullString
SESTUDIOS.Value = vbNullString
CLABORALES.Value = vbNullString
RPERSONALES.Value = vbNullString
CSEGURIDAD.Value = vbNullString
ASEGURIDAD.Value = vbNullString
DFAMILIAR.Value = vbNullString
EINGRESO.Value = vbNullString
DCONTRATO.Value = vbNullString
AINDUCCION.Value = vbNullString
MEMORANDOS.Value = vbNullString
EDESEMPEÑO.Value = vbNullString
PAZYSALVO.Value = vbNullString
EEGRESO.Value = vbNullString
LIQUIDACION.Value = vbNullString

'Delete Data from the observation of document

OFENTREVISTA.Value = "NA"
OHV.Value = "NA"
OCC.Value = "NA"
OANTECEDENTES.Value = "NA"
OSESTUDIOS.Value = "NA"
OCLABORALES.Value = "NA"
ORPERSONALES.Value = "NA"
OCSEGURIDAD.Value = "NA"
OASEGURIDAD.Value = "NA"
ODFAMILIAR.Value = "NA"
OEINGRESO.Value = "NA"
ODCONTRATO.Value = "NA"
OAINDUCCION.Value = "NA"
OMEMORANDOS.Value = "NA"
OEDESEMPEÑO.Value = "NA"
OPAZYSALVO.Value = "NA"
OEEGRESO.Value = "NA"
OLIQUIDACION.Value = "NA"

End Sub

'When close then select sheets principal
Private Sub UserForm_QueryClose(CANCEL As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        CANCEL = True
        SDocumentacion.Hide
        Sheets("PPrincipal").Select
    End If
End Sub

'Validations in fields
Private Sub FENTREVISTA_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FENTREVISTA.ListIndex < 0 And FENTREVISTA.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    FENTREVISTA.Value = vbNullString
    FENTREVISTA.DropDown
    Exit Sub
End If
End Sub
Private Sub HV_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If HV.ListIndex < 0 And HV.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    HV.Value = vbNullString
    HV.DropDown
    Exit Sub
End If
End Sub
Private Sub CC_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If CC.ListIndex < 0 And CC.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    CC.Value = vbNullString
    CC.DropDown
    Exit Sub
End If
End Sub
Private Sub ANTECEDENTES_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If ANTECEDENTES.ListIndex < 0 And ANTECEDENTES.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    ANTECEDENTES.Value = vbNullString
    ANTECEDENTES.DropDown
    Exit Sub
End If
End Sub
Private Sub SESTUDIOS_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If SESTUDIOS.ListIndex < 0 And SESTUDIOS.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    SESTUDIOS.Value = vbNullString
    SESTUDIOS.DropDown
    Exit Sub
End If
End Sub
Private Sub CLABORALES_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If CLABORALES.ListIndex < 0 And CLABORALES.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    CLABORALES.Value = vbNullString
    CLABORALES.DropDown
    Exit Sub
End If
End Sub
Private Sub RPERSONALES_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If RPERSONALES.ListIndex < 0 And RPERSONALES.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    RPERSONALES.Value = vbNullString
    RPERSONALES.DropDown
    Exit Sub
End If
End Sub
Private Sub CSEGURIDAD_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If CSEGURIDAD.ListIndex < 0 And CSEGURIDAD.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    CSEGURIDAD.Value = vbNullString
    CSEGURIDAD.DropDown
    Exit Sub
End If
End Sub
Private Sub ASEGURIDAD_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If ASEGURIDAD.ListIndex < 0 And ASEGURIDAD.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    ASEGURIDAD.Value = vbNullString
    ASEGURIDAD.DropDown
    Exit Sub
End If
End Sub
Private Sub DFAMILIAR_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If DFAMILIAR.ListIndex < 0 And DFAMILIAR.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    DFAMILIAR.Value = vbNullString
    DFAMILIAR.DropDown
    Exit Sub
End If
End Sub
Private Sub EINGRESO_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If EINGRESO.ListIndex < 0 And EINGRESO.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    EINGRESO.Value = vbNullString
    EINGRESO.DropDown
    Exit Sub
End If
End Sub
Private Sub DCONTRATO_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If DCONTRATO.ListIndex < 0 And DCONTRATO.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    DCONTRATO.Value = vbNullString
    DCONTRATO.DropDown
    Exit Sub
End If
End Sub
Private Sub AINDUCCION_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If AINDUCCION.ListIndex < 0 And AINDUCCION.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    AINDUCCION.Value = vbNullString
    AINDUCCION.DropDown
    Exit Sub
End If
End Sub
Private Sub MEMORANDOS_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If MEMORANDOS.ListIndex < 0 And MEMORANDOS.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    MEMORANDOS.Value = vbNullString
    MEMORANDOS.DropDown
    Exit Sub
End If
End Sub
Private Sub EDESEMPEÑO_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If EDESEMPEÑO.ListIndex < 0 And EDESEMPEÑO.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    EDESEMPEÑO.Value = vbNullString
    EDESEMPEÑO.DropDown
    Exit Sub
End If
End Sub
Private Sub PAZYSALVO_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If PAZYSALVO.ListIndex < 0 And PAZYSALVO.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    PAZYSALVO.Value = vbNullString
    PAZYSALVO.DropDown
    Exit Sub
End If
End Sub
Private Sub EEGRESO_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If EEGRESO.ListIndex < 0 And EEGRESO.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    EEGRESO.Value = vbNullString
    EEGRESO.DropDown
    Exit Sub
End If
End Sub
Private Sub LIQUIDACION_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If LIQUIDACION.ListIndex < 0 And LIQUIDACION.Value <> vbNullString Then
    MsgBox "Seleccione un Item de la lista"
    CANCEL = True
    LIQUIDACION.Value = vbNullString
    LIQUIDACION.DropDown
    Exit Sub
End If
End Sub
