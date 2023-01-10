VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FGPersonal 
   Caption         =   "GESTIÓN RRHH"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11475
   OleObjectBlob   =   "FGPersonal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FGPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public NewData As Boolean
Public CalcState As Long
Public wsPD As Worksheet
Public lRow As Long 'Get current row on PData sheet
Public ws As Worksheet
Public wsNews As Worksheet
Public lrNews As Long
Public ExpD As Boolean
'=====================
'Variable to Speed up the code
'=====================
Public EventState As Boolean
Public PageBreakState As Boolean
Public CurrentRow As Long
'References for dependents
Public DepSheet As Worksheet
Public DepSlastrow As Long
Public DepCRow As Long
Public DepSRowState As Long
Public RetRowState As Long
Public ShNewsReport As Worksheet
Public LastrowNR As Long
Public reportdate As Variant
Public EmpRowState As Long
Public EmpRowStateVac As Long
Public EnterpriseState As Variant
'initialize code
Sub OptimizeCode_Begin()
    Application.ScreenUpdating = False
    EventState = Application.EnableEvents
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    PageBreakState = ActiveSheet.DisplayPageBreaks
    ActiveSheet.DisplayPageBreaks = False
End Sub
Sub OptimizeCode_End()
    ActiveSheet.DisplayPageBreaks = PageBreakState
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = EventState
    Application.ScreenUpdating = True
End Sub

Public Sub AutoDocuments_Click()
Dim DSheet As Worksheet
Set DSheet = Worksheets("Auto_Docs")

DSheet.[DATED_REGISTER].Offset(1, 0) = CDate(Now())
DSheet.[EMP_NAME].Offset(1, 0) = Me.NOMBRE.Value
DSheet.[EMP_ID].Offset(1, 0) = Me.IDENTIFICACION.Value
DSheet.[EMP_ID_EXP].Offset(1, 0) = CDate(Me.DATEDEXP.Value)
DSheet.[EMP_ID_PLACE_EXP].Offset(1, 0) = Me.PLACEEXP.Value
DSheet.[EMP_DOB].Offset(1, 0) = CDate(Me.FDN.Value)
DSheet.[EMP_EMAIL].Offset(1, 0) = Me.EMAILP.Value
DSheet.[EMP_PHONE].Offset(1, 0) = Me.TELMP.Value
DSheet.[EMP_DOI].Offset(1, 0) = CDate(Me.FDI.Value)
DSheet.[EMP_DEPARTMENT].Offset(1, 0) = Me.DEP.Value
DSheet.[EMP_JOBNAME].Offset(1, 0) = Me.CARGO.Value
DSheet.[EMP_TYPE_CONTRACT].Offset(1, 0) = Me.TCONTRATO.Value
DSheet.[EMP_WAGE].Offset(1, 0) = Me.SBASE.Value
DSheet.[EMP_AUXI1].Offset(1, 0) = Me.RODAMIENTO.Value
DSheet.[EMP_AUXI2].Offset(1, 0) = Me.OAUX.Value
DSheet.[EMP_EPS].Offset(1, 0) = Me.EPS.Value
DSheet.[EMP_AFP].Offset(1, 0) = Me.AFP.Value
DSheet.[EMP_CCF].Offset(1, 0) = Me.CCF.Value
DSheet.[EMP_ARL].Offset(1, 0) = Me.ARL.Value
If Me.FechaR.Value = "" Then
    DSheet.[EMP_DOR].Offset(1, 0) = ""
Else
    DSheet.[EMP_DOR].Offset(1, 0) = CDate(Me.FechaR.Value)
End If
DSheet.[EMP_RET_REA].Offset(1, 0) = Me.MotivoR.Value

On Error Resume Next
DSheet.[dt_uc].Offset(1, 0) = CDate(Me.FCONTRATOI.Value)

Dim DTvalue As Variant

DTvalue = CDate(InputBox("Insert date in format dd/mm/yy", "User date", CDate(Format(Now, "dd/mm/yyyy"))))

DSheet.[dt_os].Offset(1, 0) = DTvalue

DSheet.[Day_Exam].Offset(1, 0) = DTvalue

DSheet.[hour_examT].Offset(1, 0) = InputBox("Insert date in format dd/mm/yy", "User date", Format(Now(), "hh:mm"))

Call SelectDocumentType

End Sub

'=====================
'End Variable to Speed up the code
'=====================
Private Sub ComboBox1_Change()
Application.EnableCancelKey = xlDisabled
'==DISABLED BUTTON IS NOBODY IS SELECTED IN SEARCH BOX
'Set value for wsPD
Set DepSheet = Sheets("PDep")
Set wsNews = Sheets("SData")
lrNews = wsNews.Cells(Rows.Count, 1).End(xlUp).Row + 1
'load images to the employees
If FGPersonal.ComboBox1.ListIndex > -2 Then
Me.REntry_Button.Enabled = True
On Error Resume Next
'FGPersonal.Employees.PictureSizeMode = fmPictureSizeModeStretch
 '  FGPersonal.Employees.Picture = LoadPicture(ThisWorkbook.Path & "\Fotos Colaboradores\" & _
 '  FGPersonal.ComboBox1.Value & ".jpg")
   Else
    On Error Resume Next
 '   FGPersonal.Employees.Picture = LoadPicture(ThisWorkbook.Path & "\Fotos Colaboradores\" & "noimage.jpg")
End If
If Err Then
    Err.Clear
  '  FGPersonal.Employees.Picture = LoadPicture(ThisWorkbook.Path & "\Fotos Colaboradores\" & "noimage.jpg")
End If
'Enable button to update data
CommandButton1.Enabled = True
NINGRESO.Enabled = False
' Variables for look up
Set wsPD = Sheets("PData")
Set ws = Sheets("PData")
Dim myRange As Range
Set myRange = Worksheets("PData").Range("A:BQ")
Dim myRange2 As Range
Set myRange2 = Worksheets("VData").Range("A:Q")
Dim DepRange As Range
    Set DepRange = DepSheet.Range("A:BC")
On Error Resume Next
lRow = Sheets("PData").UsedRange.Find(What:=Me.ComboBox1, after:=Sheets("PData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
CurrentRow = lRow
Dim lRowV As Long 'Get current row on VData sheet
lRowV = Sheets("VData").UsedRange.Find(What:=Me.ComboBox1, after:=Sheets("VData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
On Error Resume Next
'VlookUp the values of the boxes
If Sheets("VData").Cells(lRowV, VData.[vac_und_contract_dated].Column).Value = "" Then
    FCONTRATOI.Value = ""
    Else
    FCONTRATOI.Value = _
        CDate(Application.WorksheetFunction.VLookup(ComboBox1.Value, myRange2, VData.[vac_und_contract_dated].Column, False))
        If Err.Number <> 0 Then FCONTRATOI.Value = vbNullString
End If

NOMBRE.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, (wsPD.[EMPNAME].Column), False)
    If Err.Number <> 0 Then NOMBRE.Value = "NO EN"
Me.Enterprise.Value = _
    Application.WorksheetFunction.Index(PData.Range("A:A"), Application.WorksheetFunction.Match(Me.NOMBRE.Value, PData.Range("B:B"), 0))
    If Err.Number <> 0 Then Me.Enterprise.Value = ""
Me.DATEDEXP.Value = _
    CDate(Application.WorksheetFunction.VLookup(ComboBox1, myRange, (wsPD.[DATEDEXP].Column), False))
    If Err.Number <> 0 Then Me.DATEDEXP.Value = ""
Me.PLACEEXP.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, (wsPD.[PLACEEXP].Column), False)
    If Err.Number <> 0 Then Me.PLACEEXP.Value = ""
IDENTIFICACION.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[ID].Column, False)
    If Err.Number <> 0 Then IDENTIFICACION.Value = vbNullString
emp_proffession.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[p_proffession].Column, False)
    If Err.Number <> 0 Then emp_proffession.Value = vbNullString
Me.emp_proffession_card.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[e_proffesional_card].Column, False)
    If Err.Number <> 0 Then emp_proffession.Value = vbNullString
RH.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[BLOODT].Column, False)
    If Err.Number <> 0 Then Me.emp_proffession_card = vbNullString
CIVILS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[CIVILSTATUS].Column, False)
    If Err.Number <> 0 Then CIVILS.Value = vbNullString
DEGREE.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[DEGREE].Column, False)
    If Err.Number <> 0 Then DEGREE.Value = vbNullString
'==========Avoid 0:00:00 error on format
'If wsPD.Cells(lRow, wsPD.[DATEDOB].Column).Value = vbNullString Then
    'FDN.Value = ""
    'Else
    FDN.Value = _
        CDate(Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[DATEDOB].Column, False))
        If Err.Number <> 0 Then FDN.Value = vbNullString
'End If
'=======================================
CIUDAD.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[CITY].Column, False)
    If Err.Number <> 0 Then CIUDAD.Value = vbNullString
DIRECCION.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[EADDRESS].Column, False)
    If Err.Number <> 0 Then DIRECCION.Value = vbNullString
Me.NHoodList.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[NHOOD].Column, False)
    If Err.Number <> 0 Then Me.NHoodList.Value = vbNullString
Me.LISTDISTRICT.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[DISTRICT].Column, False)
    If Err.Number <> 0 Then Me.LISTDISTRICT.Value = vbNullString
EMAILP.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[EMAILP].Column, False)
    If Err.Number <> 0 Then EMAILP.Value = vbNullString
TELMP.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[EPHONEM].Column, False)
    If Err.Number <> 0 Then TELMP.Value = vbNullString
TELFP.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[EPHONES].Column, False)
    If Err.Number <> 0 Then TELFP.Value = vbNullString
EMAILCOR.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[EMAILCO].Column, False)
    If Err.Number <> 0 Then EMAILCOR.Value = vbNullString
TELMC.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[PHONEMC].Column, False)
    If Err.Number <> 0 Then TELMC.Value = vbNullString
TELFC.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[PHONESC].Column, False)
    If Err.Number <> 0 Then TELFC.Value = vbNullString
FDI.Value = _
    CDate(Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[DOI].Column, False))
    If Err.Number <> 0 Then FDI.Value = vbNullString
CDEP.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[DEPARTCODE].Column, False)
    If Err.Number <> 0 Then CDEP.Value = vbNullString
DEP.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[DEPARTNAME].Column, False)
    If Err.Number <> 0 Then DEP.Value = vbNullString
CARGO.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[JOBNAME].Column, False)
    If Err.Number <> 0 Then CARGO.Value = vbNullString
TCONTRATO.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[TContract].Column, False)
    If Err.Number <> 0 Then TCONTRATO.Value = vbNullString
SBASE.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[wage].Column, False)
    If Err.Number <> 0 Then SBASE.Value = vbNullString
RODAMIENTO.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[Auxi1].Column, False)
    If Err.Number <> 0 Then RODAMIENTO.Value = vbNullString
OAUX.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[Auxi2].Column, False)
    If Err.Number <> 0 Then OAUX.Value = vbNullString
EPS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[EPS].Column, False)
    If Err.Number <> 0 Then EPS.Value = vbNullString
AFP.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[AFP].Column, False)
    If Err.Number <> 0 Then AFP.Value = vbNullString
CCF.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[CCF].Column, False)
    If Err.Number <> 0 Then CCF.Value = vbNullString
ARL.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[ARL].Column, False)
    If Err.Number <> 0 Then ARL.Value = vbNullString
WORKC.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[JOBCENTER].Column, False)
    If Err.Number <> 0 Then WORKC.Value = vbNullString
CLASS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[RISKCLASS].Column, False)
    If Err.Number <> 0 Then CLASS.Value = vbNullString
FARE.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[FARE].Column, False)
    If Err.Number <> 0 Then FARE.Value = vbNullString
FCOBER.Value = _
    CDate(Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[doc].Column, False))
    If Err.Number <> 0 Then FCOBER.Value = vbNullString
'==CODE TO GET DATE AND TEXT FROM A CELL THAT HAS THE VALUES MERGED
'If Len(Sheets("PData").Cells(CurrentRow, wsPD.[LASTME].Column)) = 18 Then
    TEXAM.Value = _
        Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[LASTME].Column, False)
        If Err.Number <> 0 Then TEXAM.Value = vbNullString
'ElseIf Len(Sheets("PData").Cells(CurrentRow, wsPD.[LASTME].Column)) = 17 Then
 '   TEXAM.Value = _
  '      Mid((Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[LASTME].Column, False)), 1, 6)
   '     If Err.Number <> 0 Then TEXAM.Value = vbNullString
'ElseIf Len(Sheets("PData").Cells(CurrentRow, wsPD.[LASTME].Column)) = 20 Then
'    TEXAM.Value = _
 '       Mid((Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[LASTME].Column, False)), 1, 9)
 '       If Err.Number <> 0 Then TEXAM.Value = vbNullString
'End If
EXADATE.Value = _
    CDate(Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[p_exadated].Column, False))
    If Err.Number <> 0 Then EXADATE.Value = vbNullString
HCON.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[MEDICALCON].Column, False)
    If Err.Number <> 0 Then HCON.Value = vbNullString
RECOM.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[RECOM].Column, False)
    If Err.Number <> 0 Then RECOM.Value = vbNullString
REST.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[RESTRICTIONS].Column, False)
    If Err.Number <> 0 Then REST.Value = vbNullString
'================AVOID FORMAT 0:00:00 WITH EMPTY DATE
'If ws.Cells(lRow, wsPD.[DATEDOR].Column).Value = "" Then
   ' Me.FechaR.Value = ""
    'Else
    FechaR.Value = _
        CDate(Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[DATEDOR].Column, False))
        If Err.Number <> 0 Then FechaR.Value = ""
'End If
MotivoR.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[CAUSEOFRET].Column, False)
    If Err.Number <> 0 Then MotivoR.Value = ""
  
ReDATEDR.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[ReDATEDR].Column, False)
    If Err.Number <> 0 Then ReDATEDR.Value = ""

PNS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[PNS].Column, False)
    If Err.Number <> 0 Then PNS.Value = ""

RLETTERS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[RLETTERS].Column, False)
    If Err.Number <> 0 Then RLETTERS.Value = ""

LIQUIDATION.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[LIQUIDATION].Column, False)
    If Err.Number <> 0 Then LIQUIDATION.Value = ""

RDOCSSCAN.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[RDOCSSCAN].Column, False)
    If Err.Number <> 0 Then RDOCSSCAN.Value = ""

Me.retOBS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[retOBS].Column, False)
    If Err.Number <> 0 Then retOBS.Value = ""

Me.EPSS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[EPSS].Column, False)
    If Err.Number <> 0 Then Me.EPSS.Value = ""
    
Me.EPSD.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[EPSBE].Column, False)
    If Err.Number <> 0 Then Me.EPSD.Value = ""
    
Me.EPSO.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[EPSOB].Column, False)
    If Err.Number <> 0 Then Me.EPSO.Value = ""

Me.AFPS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[AFPS].Column, False)
    If Err.Number <> 0 Then Me.AFPS.Value = ""
    
Me.AFPO.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[AFPOB].Column, False)
    If Err.Number <> 0 Then Me.AFPO.Value = ""
    
Me.CCFS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[CCFS].Column, False)
    If Err.Number <> 0 Then Me.CCFS.Value = ""

Me.CCFD.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[CCFBE].Column, False)
    If Err.Number <> 0 Then Me.CCFD.Value = ""

Me.CCFO.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[CCFOB].Column, False)
    If Err.Number <> 0 Then Me.CCFO.Value = ""
    
Me.ARLS.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[ARLS].Column, False)
    If Err.Number <> 0 Then Me.ARLS.Value = ""
    
Me.ARLO.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[ARLOB].Column, False)
    If Err.Number <> 0 Then Me.ARLO.Value = "NA"

Me.CONTRACTSTATE.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[CONTRACTSTATE].Column, False)
    If Err.Number <> 0 Then Me.CONTRACTSTATE.Value = ""
    
Me.CONTRACTDATEDR.Value = _
    CDate(Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[CONTRACTDATEDR].Column, False))
    If Err.Number <> 0 Then Me.CONTRACTDATEDR.Value = ""

'=========Call dependents data
'FIRST DEPENDENT
Me.DREL1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEREL1].Column, False)
    If Err.Number <> 0 Then Me.DREL1.Value = ""
Me.DTID1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DETID1].Column, False)
    If Err.Number <> 0 Then Me.DTID1.Value = ""
Me.DID1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEID1].Column, False)
    If Err.Number <> 0 Then Me.DID1.Value = ""
Me.DFN1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEFN1].Column, False)
    If Err.Number <> 0 Then Me.DFN1.Value = ""
Me.DSN1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DESN1].Column, False)
    If Err.Number <> 0 Then Me.DSN1.Value = ""
Me.DELN1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DELN1].Column, False)
    If Err.Number <> 0 Then Me.DELN1.Value = ""
Me.DESLN1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DESLN1].Column, False)
    If Err.Number <> 0 Then Me.DESLN1.Value = ""
Me.DDOB1.Value = _
    CDate(Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEDOB1].Column, False))
    If Err.Number <> 0 Then Me.DDOB1.Value = ""
Me.DCIVILR1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DCIVILR1].Column, False)
    If Err.Number <> 0 Then Me.DCIVILR1.Value = ""
Me.TICC1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[TICC1].Column, False)
    If Err.Number <> 0 Then Me.TICC1.Value = ""
Me.STUCER1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[STUCER1].Column, False)
    If Err.Number <> 0 Then Me.STUCER1.Value = "No encontrado"
Me.MSUPPORT1.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[MSUPPORT1].Column, False)
    If Err.Number <> 0 Then Me.MSUPPORT1.Value = "No encontrado"
'SECOND DEPENDENT
Me.DREL2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEREL2].Column, False)
    If Err.Number <> 0 Then Me.DREL2.Value = ""
Me.DTID2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DETID2].Column, False)
    If Err.Number <> 0 Then Me.DTID2.Value = ""
Me.DID2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEID2].Column, False)
    If Err.Number <> 0 Then Me.DID2.Value = ""
Me.DFN2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEFN2].Column, False)
    If Err.Number <> 0 Then Me.DFN2.Value = ""
Me.DSN2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DESN2].Column, False)
    If Err.Number <> 0 Then Me.DSN2.Value = ""
Me.DELN2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DELN2].Column, False)
    If Err.Number <> 0 Then Me.DELN2.Value = ""
Me.DESLN2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DESLN2].Column, False)
    If Err.Number <> 0 Then Me.DESLN2.Value = ""
Me.DDOB2.Value = _
    CDate(Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEDOB2].Column, False))
    If Err.Number <> 0 Then Me.DDOB2.Value = ""
Me.DCIVILR2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DCIVILR2].Column, False)
    If Err.Number <> 0 Then Me.DCIVILR2.Value = ""
Me.TICC2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[TICC2].Column, False)
    If Err.Number <> 0 Then Me.TICC2.Value = ""
Me.STUCER2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[STUCER2].Column, False)
    If Err.Number <> 0 Then Me.STUCER2.Value = ""
Me.MSUPPORT2.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[MSUPPORT2].Column, False)
    If Err.Number <> 0 Then Me.MSUPPORT2.Value = ""
'THIRD DEPENDENT
Me.DREL3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEREL3].Column, False)
    If Err.Number <> 0 Then Me.DREL3.Value = ""
Me.DTID3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DETID3].Column, False)
    If Err.Number <> 0 Then Me.DTID3.Value = ""
Me.DID3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEID3].Column, False)
    If Err.Number <> 0 Then Me.DID3.Value = ""
Me.DFN3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEFN3].Column, False)
    If Err.Number <> 0 Then Me.DFN3.Value = ""
Me.DSN3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DESN3].Column, False)
    If Err.Number <> 0 Then Me.DSN3.Value = ""
Me.DELN3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DELN3].Column, False)
    If Err.Number <> 0 Then Me.DELN3.Value = ""
Me.DESLN3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DESLN3].Column, False)
    If Err.Number <> 0 Then Me.DESLN3.Value = ""
Me.DDOB3.Value = _
    CDate(Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEDOB3].Column, False))
    If Err.Number <> 0 Then Me.DDOB3.Value = ""
Me.DCIVILR3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DCIVILR3].Column, False)
    If Err.Number <> 0 Then Me.DCIVILR3.Value = ""
Me.TICC3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[TICC3].Column, False)
    If Err.Number <> 0 Then Me.TICC3.Value = ""
Me.STUCER3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[STUCER3].Column, False)
    If Err.Number <> 0 Then Me.STUCER3.Value = ""
Me.MSUPPORT3.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[MSUPPORT3].Column, False)
    If Err.Number <> 0 Then Me.MSUPPORT3.Value = ""
'FORTH DEPENDENT
Me.DREL4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEREL4].Column, False)
    If Err.Number <> 0 Then Me.DREL4.Value = ""
Me.DTID4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DETID4].Column, False)
    If Err.Number <> 0 Then Me.DTID4.Value = ""
Me.DID4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEID4].Column, False)
    If Err.Number <> 0 Then Me.DID4.Value = ""
Me.DFN4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEFN4].Column, False)
    If Err.Number <> 0 Then Me.DFN4.Value = ""
Me.DSN4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DESN4].Column, False)
    If Err.Number <> 0 Then Me.DSN4.Value = ""
Me.DELN4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DELN4].Column, False)
    If Err.Number <> 0 Then Me.DELN4.Value = ""
Me.DESLN4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DESLN4].Column, False)
    If Err.Number <> 0 Then Me.DESLN4.Value = ""
Me.DDOB4.Value = _
    CDate(Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DEDOB4].Column, False))
    If Err.Number <> 0 Then Me.DDOB4.Value = ""
Me.DCIVILR4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[DCIVILR4].Column, False)
    If Err.Number <> 0 Then Me.DCIVILR4.Value = ""
Me.TICC4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[TICC4].Column, False)
    If Err.Number <> 0 Then Me.TICC4.Value = ""
Me.STUCER4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[STUCER4].Column, False)
    If Err.Number <> 0 Then Me.STUCER4.Value = ""
Me.MSUPPORT4.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, DepRange, DepSheet.[MSUPPORT4].Column, False)
    If Err.Number <> 0 Then Me.MSUPPORT4.Value = ""
    
Me.e_branch.Value = _
    Application.WorksheetFunction.VLookup(ComboBox1, myRange, (wsPD.[e_branch].Column), False)
    If Err.Number <> 0 Then Me.e_branch.Value = "NA"
'=================AVOID FORMAT: 0:00:00 WITH EMPTY DATE
If ws.Cells(lRow, wsPD.[DATERARL].Column).Value = "" Then
    Me.FRARL.Value = ""
    Else
    On Error Resume Next
    FRARL.Value = _
        CDate(Application.WorksheetFunction.VLookup(ComboBox1, myRange, wsPD.[DATERARL].Column, False))
        If Err.Number <> 0 Then FRARL.Value = vbNullString
End If
' Assign true or false for checkbox, taking into account the data of the texbox
If wsPD.Cells(CurrentRow, wsPD.[RETIRED].Column).Value = True Then
        Me.Retirado = True
        Me.Retirado.Enabled = True
        Me.REntry_Button.Enabled = True
    ElseIf Cells(CurrentRow, wsPD.[RETIRED].Column).Value = False Then
        Me.Retirado = False
        Me.Retirado.Enabled = True
        Me.REntry_Button.Enabled = False
End If
End Sub
'===================================================
Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'=============================================
'ENABLED "INGRESO" WHEN COMBOBOX IS NULL
'=============================================
If KeyCode.Value = vbKeyBack Then
    CommandButton1.Enabled = False
    NINGRESO.Enabled = True
    Me.REntry_Button.Enabled = False
    Me.Retirado = False
End If
End Sub
'=============================================
Sub UpdateInf()
Call OptimizeCode_Begin
'============================================
' UPDATE THE INFORMATION
'=============================================
'Variables for find values
Dim ws As Worksheet
Set ws = Sheets("PData")
Dim lRow As Long
EmpRowState = Sheets("PData").UsedRange.Find(What:=Me.ComboBox1, after:=Sheets("PData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row

Dim lRowV As Long

EmpRowStateVac = Sheets("VData").UsedRange.Find(What:=Me.ComboBox1, after:=Sheets("VData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row

Call AddEmpInformation

MsgBox "Datos Actualizados"
Me.ComboBox1.SetFocus
Call UserForm_Initialize
Call OptimizeCode_End
End Sub

Private Sub CommandButton2_Click()
'=======================================
' CLOSE USERFORM
'=======================================
Unload Me
'Sheets("PPrincipal").Select
End Sub

Private Sub Create_DepState_Click()

End Sub

Private Sub DATEDEXP_Change()
'==============================
'VALIDATION FOR DATE EXP
'==============================
If Me.DATEDEXP.TextLength > 1 And Me.DATEDEXP.TextLength < 3 Then
    Me.DATEDEXP.Value = Me.DATEDEXP.Value & "/"
End If
If Me.DATEDEXP.TextLength > 10 Then
    Me.DATEDEXP.Value = Mid(Me.DATEDEXP.Text, 1, Len(Me.DATEDEXP.Text) - 1)
End If
    If Me.DATEDEXP.TextLength > 4 And Me.DATEDEXP.TextLength < 6 Then
    DATEDEXP.Value = Me.DATEDEXP.Value & "/"
End If
End Sub
Private Sub DATEDEXP_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       DATEDEXP.Value = vbNullString
    End If
End Sub
Private Sub DATEDEXP_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If DATEDEXP.TextLength > 1 And DATEDEXP.TextLength < 10 And DATEDEXP.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    DATEDEXP.Value = vbNullString
    DATEDEXP.SetFocus
    Exit Sub
End If
End Sub
Private Sub EMAILCOR_Change()
Me.EMAILCOR.Value = LCase(Me.EMAILCOR.Value)
End Sub

Private Sub EMAILP_Change()
Me.EMAILP.Value = LCase(Me.EMAILP.Value)
End Sub

Private Sub emp_proffession_Change()

End Sub

'==================================
'END OF VALIDATION OF DATE EXP
'===================================
Private Sub FRARL_Change()
'====================================
'VALIDATION FOR DATE OF RETIREMENT
'==================================
If FRARL.TextLength > 1 And FRARL.TextLength < 3 Then
    FRARL.Value = FRARL.Value & "/"
End If
If FRARL.TextLength > 10 Then
    FRARL.Value = Mid(FRARL.Text, 1, Len(FRARL.Text) - 1)
End If
If FRARL.TextLength > 4 And FRARL.TextLength < 6 Then
    FRARL.Value = FRARL.Value & "/"
End If
End Sub
Private Sub FRARL_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       FRARL.Value = vbNullString
    End If
End Sub
Private Sub FRARL_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FRARL.TextLength > 1 And FRARL.TextLength < 10 And FRARL.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    FRARL.Value = vbNullString
    FRARL.SetFocus
Exit Sub
End If
'====================================
'END VALIDATION FOR DATE OF RETIREMENT
'==================================
End Sub

Private Sub INDEXP_Change()
If Me.ExpD = False Then
'    Select Case Me.INDEXP.TextLength
 '       Case 2 To 2
  '          Me.ExpD = False
   '         Me.INDEXP.Value = Me.INDEXP.Value & "/"
    '
     '   Case 5 To 5
      '      Me.ExpD = False
       ''     Me.INDEXP.Value = Me.INDEXP.Value & "/"
       '
       ' Case 11 To 11
       '     Me.INDEXP.Value = Mid(Me.INDEXP.Value, 1, Len(Me.INDEXP.Value) - 1)
       '     Me.ExpD = True
       ' Case 12 To 12
       '     Me.ExpD = False
    'End Select
    
End If
End Sub

Private Sub INDEXP_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
    Me.INDEXP.Value = ""
End If
End Sub

Private Sub INDEXP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    Me.INDEXP.Value = ""
End If
End Sub

'=============================================
Private Sub IRetiro_Click()
'=======================================
'DO NOT ALLOW BLANK SPACE FOR RETIREMENT
'=======================================
If FechaR.Value = vbNullString Then
    MsgBox "Ingrese un Valor"
    FechaR.SetFocus
    Exit Sub
End If
If MotivoR.Value = vbNullString Then
    MsgBox "Ingrese un Valor"
    MotivoR.SetFocus
    Exit Sub
End If
If FRARL.Value = vbNullString Then
    MsgBox "Ingrese un Valor"
    FRARL.SetFocus
    Exit Sub
End If
'=======================================
' REGISTER RETIRED
'=======================================
If FGPersonal.Retirado.Value = True Then
PData.Cells(CurrentRow, wsPD.[DATERARL].Column).Value = CDate(FGPersonal.FRARL)
PData.Cells(CurrentRow, wsPD.[RETIRED].Column).Value = FGPersonal.Retirado
    If Me.Retirado.Value = True Then
        PData.Cells(CurrentRow, wsPD.[DATEDOR].Column).Value = CDate(FGPersonal.FechaR.Value)
    Else
        PData.Cells(CurrentRow, wsPD.[DATEDOR].Column).Value = vbNullString
    End If
PData.Cells(CurrentRow, wsPD.[CAUSEOFRET].Column).Value = FGPersonal.MotivoR.Value
MsgBox "Novedad Registrada"
Else
MsgBox "Marque la Casilla Retirado"
End If
'Generate CSV file for letters
Dim myFile As String, Rng As Range, cellValue As Variant, i As Integer, j As Integer
myFile = ActiveWorkbook.Path & "\data.csv"

Dim WageWord As Variant
Dim Auxi1 As Variant
Dim Auxi2 As Variant
Dim TContract As Variant

TContract = LCase(Me.TCONTRATO.Value)
WageWord = NumLetras(Me.SBASE.Value)
Auxi1 = NumLetras(Me.RODAMIENTO.Value)
Auxi2 = NumLetras(Me.OAUX.Value)

'Add News Type
Call AutoDocuments_Click

NewsType = 2
Call NewsReportS
Call SelectDocumentType
'Create File

Open myFile For Output As #1


Write #1, "Nombre", "CC", "FI", "FR", "FCesantias", "Cargo", "Salario", "SalarioW", "Auxilio", "AuxilioW", "Auxilio2", "AuxilioW2", "Contrato"
Write #1, Me.NOMBRE, Me.IDENTIFICACION, Me.FDI, Me.FechaR, Me.AFP, Me.CARGO, Me.SBASE, WageWord, Me.RODAMIENTO, Auxi1, Me.OAUX, Auxi2, TContract

Close #1

End Sub



'==================================
Private Sub MALE_Click()
'Code to desactive the selection of the two option at the same time
If MALE = True Then
    FEMALE.Enabled = False
Else
    FEMALE.Enabled = True
End If
End Sub
Private Sub FEMALE_Click()
'Code to desactive the selection of the two option at the same time
If FEMALE = True Then
    MALE.Enabled = False
Else
    MALE.Enabled = True
End If
End Sub
Private Sub CommandButton1_Click()
Call OptimizeCode_Begin
Application.EnableCancelKey = xlDisabled
Dim typeC As Range
'=======================================
' ASK THE USER IF A VALUE CHANGE AND NEED TO BE REPORT
'=======================================
'UPDATE INFO THAT DOESN'T REQUIRED DATE
'wsPD.Cells(lRow, wsPD.[RECOM].Column).Value = Me.RECOM.Value
'wsPD.Cells(lRow, wsPD.[RESTRICTIONS].Column).Value = Me.REST.Value

If MsgBox("¿El cambio realizado necesita ser reportado como novedad?", vbYesNo) = vbYes Then
    reportdate = Application.InputBox("register dated", Type:=2)
    If Me.DEP.Value <> wsPD.Cells(lRow, wsPD.[DEPARTNAME].Column).Value Then
        'Call UpdateInfo2
        NewsType = 31
        Call NewsReportS
    End If
    If Me.CARGO.Value <> wsPD.Cells(lRow, wsPD.[JOBNAME].Column).Value Then
        NewsType = 32
        Call NewsReportS
    End If
    If Me.TCONTRATO.Value <> wsPD.Cells(lRow, wsPD.[TContract].Column).Value Then
        NewsType = 33
        Call NewsReportS
    End If
    '========================
    If CLng(Me.SBASE.Value) <> CLng(wsPD.Cells(lRow, wsPD.[wage].Column).Value) Then
        NewsType = 34
        Call NewsReportS
    End If
    If CLng(Me.RODAMIENTO) <> CLng(wsPD.Cells(lRow, wsPD.[Auxi1].Column)) Then
        NewsType = 35
        Call NewsReportS
    End If
    If CLng(Me.OAUX.Value) <> CLng(wsPD.Cells(lRow, wsPD.[Auxi2].Column).Value) Then
        NewsType = 36
        Call NewsReportS
    End If
    If Me.EPS.Value <> wsPD.Cells(lRow, wsPD.[EPS].Column).Value Then
        NewsType = 37
        Call NewsReportS
    End If
    If Me.AFP.Value <> wsPD.Cells(lRow, wsPD.[AFP].Column).Value Then
        NewsType = 38
        Call NewsReportS
    End If
    
Else

End If

EnterpriseState = Me.Enterprise.Value

EmpRowState = Sheets("PData").UsedRange.Find(What:=Me.ComboBox1, after:=Sheets("PData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row

EmpRowStateVac = Sheets("VData").UsedRange.Find(What:=Me.ComboBox1, after:=Sheets("VData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row



'update dependents info
DepCRow = Sheets("PDep").UsedRange.Find(What:=Me.ComboBox1, after:=Sheets("PDep").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
DepSRowState = DepCRow

Call AddEmpInformation
Call AddDependentsInfo

    Call OptimizeCode_End
End Sub
Sub UpdateInfo2()
        wsNews.Cells(lrNews, 1).Value = NewsDateD
        wsNews.Cells(lrNews, 2).Value = Me.ComboBox1.Value
        wsNews.Cells(lrNews, 3).Value = Me.IDENTIFICACION.Value
End Sub

'=========================================
Private Sub NINGRESO_Click()

EnterpriseState = Me.Enterprise.Value & Me.NOMBRE.Value

EmpRowStateVac = Sheets("VData").Cells(Rows.Count, 1).End(xlUp).Row + 1
EmpRowState = Sheets("PData").Cells(Rows.Count, 1).End(xlUp).Row + 1

DepSRowState = DepSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1

Call AddEmpInformation

NewsType = 1
Call NewsReportS

MsgBox ("Datos Actualizados")
Call DeleteData
Me.ComboBox1.SetFocus

Application.Calculation = xlCalculationAutomatic
Call OptimizeCode_End
End Sub



'==========================================
Private Sub Retirado_Click()
'Register retire of a employee
If Me.Retirado = False Then
    Me.MotivoR.Enabled = False
    Me.IRetiro.Enabled = False
    Me.FechaR.Enabled = False
    Me.FRARL.Enabled = False
    Else
    Me.MotivoR.Enabled = True
    Me.IRetiro.Enabled = True
    Me.FechaR.Enabled = True
    Me.FRARL.Enabled = True
End If
End Sub

Private Sub RETTRACKING_Click()
Dim rowstate As Long
rowstate = Sheets("PData").UsedRange.Find(What:=Me.ComboBox1, after:=Sheets("PData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
RetRowState = rowstate
Call TrackingRetirementInfo
End Sub

Private Sub SpinButton1_Change()

End Sub

Private Sub TCONTRATO_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If TCONTRATO.ListIndex < 0 Then
    MsgBox "Seleccione un dato de la lista"
    CANCEL = True
    TCONTRATO.Value = vbNullString
    TCONTRATO.DropDown
End If
End Sub
Private Sub UserForm_Initialize()
Application.EnableCancelKey = xlDisabled
'create sheet reference
Set DepSheet = Sheets("PDep")
Set wsPD = Sheets("PData")
'other code

Me.ExpD = False
Retirado = False 'When userform is initialize then retirado is false
Me.REntry_Button.Enabled = False
If ComboBox1.Value = Null Or ComboBox1.Value = vbNullString Then
    CommandButton1.Enabled = False
End If
If ComboBox1.ListIndex > 1 Then
    CommandButton1.Enabled = True
End If
Me.ComboBox1.SetFocus
Me.MultiPage1.Value = 0
'==ADD DATA TO THE COMBOBOX BASE ON TYPE OF ENTITY=======================
Dim r As Long, rngEntity As Range, lrEntity As Long, rngEntity2 As Range
lrEntity = Sheets("C_CIE10").Cells(Rows.Count, 4).End(xlUp).Row
    Set rngEntity = Sheets("C_CIE10").Range("E2:E" & lrEntity)
For r = rngEntity.Count To 1 Step -1
    If rngEntity(r).Value = "EPS" Then
            Set rngEntity2 = rngEntity(r).Offset(0, -1)
            Me.EPS.AddItem rngEntity2.Value
        ElseIf rngEntity(r).Value = "AFP" Then
            Set rngEntity2 = rngEntity(r).Offset(0, -1)
            Me.AFP.AddItem rngEntity2.Value
        ElseIf rngEntity(r).Value = "CCF" Then
            Set rngEntity2 = rngEntity(r).Offset(0, -1)
            Me.CCF.AddItem rngEntity2.Value
        ElseIf rngEntity(r).Value = "ARL" Then
            Set rngEntity2 = rngEntity(r).Offset(0, -1)
            Me.ARL.AddItem rngEntity2
    End If
Next r
'===========================================================================


Dim Status As Variant
Status = Array("AFILIADO", "PENDIENTE", "RADICADO", "DILIGENCIADO", "LOGISTICA", "PENDIENTE GESTION")
Dim Dependents_Status As Variant
Dependents_Status = Array("NO REPORTA", "PENDIENTE DOC", "AFILIADOS")
Dim Branchs As Variant: Branchs = Array("1 DE MAYO", "CALLE 26", "CALLE 63", "CALLE 100", "CL GIRARDOT", "CM BUCARAMANGA", "CM GIRARDTO", "FUNZA", "INFANTIL", "MAZUREN", "PEREIRA", "ROMA", "SUBA", "USAQUEN", "VILLAVICENCIO", "MEDPLUS", "CLINICA NUEVA", "OFICINA PRINCIPAL")

Dim ctrl As MSForms.control
    For Each ctrl In Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            If ctrl.Name Like "MotivoR" Then ctrl.List = Array("RENUNCIA VOLUNTARIA", "TERMINACION DE CONTRATO FIJO CON JUSTA CAUSA", "TERMINACION DE CONTRATO FIJO SIN JUSTA CAUSA", "TERMINACION DE CONTRATO INDEFINIDO CON JUSTA CAUSA", "TERMINACION DE CONTRATO INDEFINIDO SIN JUSTA CAUSA", "TERMINACION DE CONTRATO FIJO DENTRO DEL PERIODO DE PRUEBA", "TERMINACION DE CONTRATO INDEFINIDO DENTRO DEL PERIODO DE PRUEBA", "TERMINACION DE CONTRATO DE APRENDIZAJE", "APLAZAMIENTO DE CONTRATO")
            If ctrl.Name Like "CIVILS" Then ctrl.List = Array("SOLTERO", "UNION LIBRE", "CASADO", "SEPARADO", "VIUDO(A)")
            If ctrl.Name Like "DEGREE" Then ctrl.List = Array("DOCTORADO EN FORMACION", "DOCTORADO", "MAESTRIA EN FORMACION", "MAESTRIA", "ESPECIALIZACION EN FORMACION", "ESPECIALIZACION", "NIVEL PROFESIONAL EN FORMACION", "NIVEL PROFESIONAL", "NIVEL TECNOLOGICO EN FORMACION", "NIVEL TECNOLOGICO", "NIVEL TECNICO EN FORMACION", "NIVEL TECNICO", "EDUCACION MEDIA EN FORMACION", "EDUCACION MEDIA", "EDUCACION BASICA EN FORMACION", "EDUCACION BASICA", "PREESCOLAR")
            If ctrl.Name Like "WORKC" Then ctrl.List = Array("ADMINISTRATIVO BOGOTA", "ADMINISTRATIVO MEDELLIN", "INGENIERIA CLINICA", "COMERCIAL", "MENSAJERIA BOGOTA D.C.", "INGENIERIA Y SOPORTE MEDELLIN", "RADIOLOGIA BARRANQUILLA", "ADMINISTRATIVO COLSUBSIDIO", "RADIOLOGIA COLSUBSIDIO", "GINECOLOGIA COLSUBSIDIO")
            If ctrl.Name Like "CLASS" Then ctrl.List = Array("1", "2", "3", "4", "5", "6")
            If ctrl.Name Like "TEXAM" Then ctrl.List = Array("INGRESO", "PERIODICO", "EGRESO")
            If ctrl.Name Like "HCON" Then ctrl.List = Array("APTO", "NO APTO")
            If ctrl.Name Like "CIUDAD" Or ctrl.Name Like "PLACEEXP" Then ctrl.List = Array("BOGOTA D.C", "MEDELLIN", "BARRANQUILLA", "MONTERIA", "CIUDAD DE MEXICO", "LIMA", "CHIA", "FUNZA", "MOSQUERA", "ZIPAQUIRA", "CAJICA", "SOACHA", "LA CALERA", "PEREIRA", "VILLAVICENCIO", "GIRARDOT", "MADRID", "FACATATIVA", "BUCARAMANGA", "TUNJA")
            If ctrl.Name Like "DEP" Then ctrl.List = Array("ADMINISTRACION", "CALIDAD", "COMERCIAL", "CONTABILIDAD", "DISEÑO Y DESARROLLO", "GERENCIA", "GESTION IT", "INGENIERIA Y SOPORTE", "INVESTIGACION Y DESARROLLO", "LOGISTICA", "PROYECTOS", "MARKETING", "RADIOLOGIA", "COLSUBSIDIO ADMINISTRATIVO", "COLSUBSIDIO ASISTENCIAL")
            If ctrl.Name Like "TCONTRATO" Then ctrl.List = Array("INDEFINIDO", "INDEFINIDO - S INTEGRAL", "FIJO 3 MESES", "FIJO 6 MESES", "PRESTACION DE SERVICIOS", "APRENDIZ SENA - ETAPA LECTIVA", "APRENDIZ SENA - ETAPA PRODUCTIVA", "APRENDIZ UNIVERSITARIO", "OBRA O LABOR", "FIJO 4 MESES", "FIJO 2 MESES", "FIJO 6 MESES", "FIJO 4 MESES")
            If ctrl.Name Like "EPSS" Then ctrl.List = Status
            If ctrl.Name Like "AFPS" Then ctrl.List = Status
            If ctrl.Name Like "CCFS" Then ctrl.List = Status
            If ctrl.Name Like "ARLS" Then ctrl.List = Status
            If ctrl.Name Like "EPSD" Then ctrl.List = Dependents_Status
            If ctrl.Name Like "CCFD" Then ctrl.List = Dependents_Status
            If ctrl.Name Like "e_branch" Then ctrl.List = Branchs
        End If
    Next ctrl
'Add name and lastname of the employees to the listbox to be selected
Dim lastrow As Long
lastrow = Sheets("PData").Cells(Rows.Count, 1).End(xlUp).Row
ComboBox1.List = Sheets("PData").Range("A2:A" & lastrow).Value
'Add data to the districts
Dim lastRowNHood As Long
lastRowNHood = Sheets("C_Cie10").Cells(Rows.Count, 18).End(xlUp).Row
Me.NHoodList.List = Sheets("C_Cie10").Range("R2:R" & lastRowNHood).Value
'Other
Me.Retirado.Enabled = False
End Sub
Private Sub REntry_Button_Click()
Dim lastRowREntry As Long
lastRowREntry = Sheets("RData").Cells(Rows.Count, 1).End(xlUp).Row + 1
Dim wsRange As Range
Set wsRange = ws.Range(ws.Cells(lRow, 1), ws.Cells(lRow, wsPD.[CAUSEOFRET].Column))
wsRange.Copy Sheets("RData").Cells(lastRowREntry, 1)
ws.Cells(lRow, wsPD.[DATERARL].Column).Value = vbNullString
ws.Cells(lRow, wsPD.[RETIRED].Column).Value = False
ws.Cells(lRow, wsPD.[DATEDOR].Column).Value = vbNullString
ws.Cells(lRow, wsPD.[CAUSEOFRET].Column).Value = vbNullString
'======================================
'CALL NEWS REPORT FOR ACCOUNTING REPORT
'=======================================
NewsType = 1
Call NewsReportS
'====================
'End adding news
'====================
MsgBox "Colaborador registrado como reingreso, ingrese los nuevos datos de ingreso"
Call UserForm_Initialize
End Sub
'===============================
Function DeleteData()
'Delete the data of all the texbox
Me.NOMBRE.Value = vbNullString
Me.IDENTIFICACION.Value = vbNullString
Me.emp_proffession.Value = ""
Me.emp_proffession_card.Value = ""
Me.RH.Value = vbNullString
Me.FDN.Value = vbNullString
Me.CIUDAD.Value = vbNullString
Me.DIRECCION.Value = vbNullString
Me.EMAILP.Value = vbNullString
Me.TELMP.Value = vbNullString
Me.TELFP.Value = vbNullString
Me.EMAILCOR.Value = vbNullString
Me.TELMC.Value = vbNullString
Me.TELFC.Value = vbNullString
Me.FDI.Value = vbNullString
Me.CDEP.Value = vbNullString
Me.DEP.Value = vbNullString
Me.CARGO.Value = vbNullString
Me.TCONTRATO.Value = vbNullString
Me.SBASE.Value = vbNullString
Me.RODAMIENTO.Value = vbNullString
Me.OAUX.Value = vbNullString
Me.EPS.Value = vbNullString
Me.AFP.Value = vbNullString
Me.CCF.Value = vbNullString
Me.ARL.Value = vbNullString
Me.MALE = False
Me.FEMALE = False
Me.DEGREE.Value = vbNullString
Me.WORKC.Value = vbNullString
Me.CLASS.Value = vbNullString
Me.FARE.Value = vbNullString
Me.FCOBER.Value = vbNullString
Me.TEXAM.Value = vbNullString
Me.EXADATE.Value = vbNullString
Me.HCON.Value = vbNullString
Me.RECOM.Value = vbNullString
Me.REST.Value = vbNullString
Me.Retirado.Value = False
Me.MotivoR.Value = vbNullString
Me.ComboBox1.Value = vbNullString
End Function
'VALIDATIONS OF THE FIELDS
'==================================
Private Sub DEP_Change()
If DEP.Value = vbNullString Then
    CDEP.Value = vbNullString
ElseIf DEP.Value = "ADMINISTRACION" Then
    CDEP.Value = "02A"
ElseIf DEP.Value = "CALIDAD" Then
    CDEP.Value = "02A"
ElseIf DEP.Value = "COMERCIAL" Then
    CDEP.Value = "04C"
ElseIf DEP.Value = "CONTABILIDAD" Then
    CDEP.Value = "02A"
ElseIf DEP.Value = "DISEÑO Y DESARROLLO" Then
    CDEP.Value = "09IS"
ElseIf DEP.Value = "GERENCIA" Then
    CDEP.Value = "02A"
ElseIf DEP.Value = "GESTION IT" Then
    CDEP.Value = "05IC"
ElseIf DEP.Value = "INGENIERIA Y SOPORTE" Then
    CDEP.Value = "05IC"
ElseIf DEP.Value = "INVESTIGACION Y DESARROLLO" Then
    CDEP.Value = "08IS"
ElseIf DEP.Value = "LOGISTICA" Then
    CDEP.Value = "02A"
ElseIf DEP.Value = "PROYECTOS" Then
    CDEP.Value = "10P"
    ElseIf DEP.Value = "MARKETING" Then
    CDEP.Value = "04C"
ElseIf DEP.Value = "RADIOLOGIA" Then
    CDEP.Value = "05IC"
ElseIf DEP.Value = "COLSUBSIDIO ADMINISTRATIVO" Then
    CDEP.Value = "0065CA"
ElseIf DEP.Value = "COLSUBSIDIO ASISTENCIAL" Then
    CDEP.Value = "0066CAS"
End If
End Sub
Private Sub DEP_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
    If DEP.ListIndex < 0 And DEP.Value <> vbNullString Then
    MsgBox "Seleccione un dato de la lista"
    DEP.Value = vbNullString
    CANCEL = True
    DEP.SetFocus
    DEP.DropDown
    End If
End Sub
Private Sub CARGO_Change()
    CARGO.Value = UCase(CARGO.Value)
End Sub
Private Sub EPS_Change()
    EPS.Value = UCase(EPS.Value)
End Sub
Private Sub ARL_Change()
    ARL.Value = UCase(ARL.Value)
End Sub
Private Sub AFP_Change()
    AFP.Value = UCase(AFP.Value)
End Sub
Private Sub CCF_Change()
    CCF.Value = UCase(CCF.Value)
End Sub
Private Sub RH_Change()
If RH.TextLength > 3 Then
MsgBox "No se permiten más datos"
    RH.Text = Mid(RH.Text, 1, Len(RH.Text) - 1)
End If
    RH.Text = UCase(RH.Text)
End Sub
Private Sub NOMBRE_Change()
    NOMBRE.Text = UCase(NOMBRE.Text)
End Sub
Private Sub CIUDAD_Change()
    CIUDAD.Text = UCase(CIUDAD.Text)
End Sub
Private Sub WORKC_Change()
If WORKC.Value = "ADMINISTRATIVO BOGOTA" Or WORKC.Value = "ADMINISTRATIVO MEDELLIN" _
    Or WORKC.Value = "ADMINISTRATIVO MEXICO" Then
            CLASS.Value = 1
            FARE.Value = 0.00522
        ElseIf WORKC.Value = "INGENIERIA CLINICA" Or WORKC.Value = "INGENIERIA Y SOPORTE MEDELLIN" Then
            CLASS.Value = 3
            FARE.Value = 0.02436
        ElseIf WORKC.Value = "MENSAJERIA BOGOTA D.C." Then
            CLASS.Value = 4
            FARE.Value = 0.0435
        ElseIf WORKC.Value = "COMERCIAL" Then
            CLASS.Value = 2
            FARE.Value = 0.00144
        ElseIf WORKC.Value = "RADIOLOGIA BARRANQUILLA" Then
            CLASS.Value = 5
            FARE.Value = 0.0696
        ElseIf WORKC.Value = "ADMINISTRATIVO COLSUBSIDIO" Then
            CLASS.Value = 1
            FARE.Value = 0.00522
        ElseIf WORKC.Value = "RADIOLOGIA COLSUBSIDIO" Then
            CLASS.Value = 5
            FARE.Value = 0.0696
        ElseIf WORKC.Value = "GINECOLOGIA COLSUBSIDIO" Then
            CLASS.Value = 3
            FARE.Value = "0.02436"
    End If
End Sub
'==========================================
'DO NO ALLOW USER TO SELECT DIFFERENT VALUE
'==========================================
Private Sub CIUDAD_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
    If CIUDAD.ListIndex < 0 And CIUDAD.Value <> vbNullString Then
        MsgBox "Seleccione un dato de la lista"
        CANCEL = True
        CIUDAD.Value = vbNullString
        CIUDAD.SetFocus
        CIUDAD.DropDown
    End If
End Sub
Private Sub CIVILS_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
    If CIVILS.ListIndex < 0 And CIVILS.Value <> vbNullString Then
        MsgBox "Seleccione un datos de la lista"
        CIVILS.ListIndex = Null
        CIVILS.DropDown
        CANCEL = True
    End If
End Sub
Private Sub DEGREE_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
    If DEGREE.ListIndex < 0 And DEGREE.Value <> vbNullString Then
        MsgBox "Seleccione un datos de la lista"
        DEGREE.ListIndex = Null
        DEGREE.DropDown
        CANCEL = True
    End If
End Sub
Private Sub WORKC_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
    If WORKC.ListIndex < 0 And WORKC.Value <> vbNullString Then
        MsgBox "Seleccione un datos de la lista"
        WORKC.ListIndex = Null
        WORKC.DropDown
        CANCEL = True
    End If
End Sub
Private Sub CLASS_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
    If CLASS.ListIndex < 0 And CLASS.Value <> vbNullString Then
        MsgBox "Seleccione un datos de la lista"
        CLASS.ListIndex = Null
        CLASS.DropDown
        CANCEL = True
    End If
End Sub
Private Sub TEXAM_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
    If TEXAM.ListIndex < 0 And TEXAM.Value <> vbNullString Then
        MsgBox "Seleccione un datos de la lista"
        TEXAM.ListIndex = Null
        TEXAM.DropDown
        CANCEL = True
    End If
End Sub
Private Sub HCON_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
    If HCON.ListIndex < 0 And HCON.Value <> vbNullString Then
        MsgBox "Seleccione un datos de la lista"
        HCON.ListIndex = Null
        HCON.DropDown
        CANCEL = True
    End If
End Sub
Private Sub MotivoR_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
    If MotivoR.ListIndex < 0 And MotivoR.Value <> vbNullString Then
        MsgBox "Seleccione un dato de la lista"
        MotivoR.Value = vbNullString
        CANCEL = True
        MotivoR.SetFocus
        MotivoR.DropDown
    End If
    If MotivoR.Value = vbNullString Then
        MotivoR.Width = 120
    End If
End Sub
'==========================================
'END DO NOT ALLOW USER TO SELECT DIFFERENT VALUE
'==========================================
Private Sub DIRECCION_Change()
    DIRECCION.Text = UCase(DIRECCION.Text)
End Sub
Private Sub IDENTIFICACION_Change()
If Not IsNumeric(IDENTIFICACION.Text) And IDENTIFICACION.Text <> vbNullString Then
    Beep
    MsgBox "Ingrese un valor numerico"
    IDENTIFICACION.Value = Mid(IDENTIFICACION.Text, 1, Len(IDENTIFICACION.Text) - 1)
    Exit Sub
End If
End Sub
Private Sub FDN_Change()
'==============================
'VALIDATION FOR DATE OF BIRTH
'==============================
If FDN.TextLength > 1 And FDN.TextLength < 3 Then
    FDN.Value = FDN.Value & "/"
End If
If FDN.TextLength > 10 Then
    FDN.Value = Mid(FDN.Text, 1, Len(FDN.Text) - 1)
End If
    If FDN.TextLength > 4 And FDN.TextLength < 6 Then
    FDN.Value = FDN.Value & "/"
End If
End Sub
Private Sub FDN_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       FDN.Value = vbNullString
    End If
End Sub
Private Sub FDN_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FDN.TextLength > 1 And FDN.TextLength < 10 And FDN.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    FDN.Value = vbNullString
    FDN.SetFocus
    Exit Sub
End If
'==================================
'END OF VALIDATION OF DATE OF BIRTH
'===================================
End Sub
Private Sub FDI_Change()
'VALIDATION FOR DATE OF INCORPORATION------------------
If FDI.TextLength > 1 And FDI.TextLength < 3 Then
    FDI.Value = FDI.Value & "/"
End If
If FDI.TextLength > 10 Then
    FDI.Value = Mid(FDI.Text, 1, Len(FDI.Text) - 1)
End If
If FDI.TextLength > 4 And FDI.TextLength < 6 Then
    FDI.Value = FDI.Value & "/"
End If
End Sub
Private Sub FDI_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
    FDI.Value = vbNullString
End If
End Sub
Private Sub FDI_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FDI.TextLength > 1 And FDI.TextLength < 10 And FDI.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    FDI.Value = vbNullString
    FDI.SetFocus
    Exit Sub
End If
'END OF VALIDATION OF DATE OF INCORPORATION------------------
End Sub
Private Sub FCONTRATOI_Change()
'=======================================
'START OF VALIDATION OF DATE OF CONTRACT
'========================================
If FCONTRATOI.TextLength > 1 And FCONTRATOI.TextLength < 3 And FCONTRATOI.Value <> "N" And FCONTRATOI.Value <> "NA" Then
    FCONTRATOI.Value = FCONTRATOI.Value & "/"
End If
If FCONTRATOI.TextLength > 10 And FCONTRATOI.Value <> "N" And FCONTRATOI.Value <> "NA" Then
    FCONTRATOI.Value = Mid(FCONTRATOI.Text, 1, Len(FCONTRATOI.Text) - 1)
End If
If FCONTRATOI.TextLength > 4 And FCONTRATOI.TextLength < 6 And FCONTRATOI.Value <> "N" And FCONTRATOI.Value <> "NA" Then
    FCONTRATOI.Value = FCONTRATOI.Value & "/"
End If
End Sub
Private Sub FCONTRATOI_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyBack Then
       FCONTRATOI.Value = vbNullString
    End If
End Sub
Private Sub FCONTRATOI_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FCONTRATOI.TextLength >= 1 And FCONTRATOI.TextLength < 10 And FCONTRATOI.Text <> vbNullString And FCONTRATOI.Value <> "NA" And FCONTRATOI.Value <> "N" Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    FCONTRATOI.Value = vbNullString
    CANCEL = True
    Exit Sub
End If
'=======================================
'END OF VALIDATION OF DATE OF CONTRACT
'========================================
End Sub
Private Sub FechaR_Change()
'====================================
'VALIDATION FOR DATE OF RETIREMENT
'==================================
If FechaR.TextLength > 1 And FechaR.TextLength < 3 Then
    FechaR.Value = FechaR.Value & "/"
End If
If FechaR.TextLength > 10 Then
    FechaR.Value = Mid(FechaR.Text, 1, Len(FechaR.Text) - 1)
End If
If FechaR.TextLength > 4 And FechaR.TextLength < 6 Then
    FechaR.Value = FechaR.Value & "/"
End If
End Sub
Private Sub FechaR_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       FechaR.Value = vbNullString
    End If
End Sub
Private Sub FechaR_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FechaR.TextLength > 1 And FechaR.TextLength < 10 And FechaR.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    FechaR.Value = vbNullString
    CANCEL = True
    FechaR.SetFocus
    Exit Sub
End If
'====================================
'END VALIDATION FOR DATE OF RETIREMENT
'==================================
End Sub
Private Sub FCOBER_Change()
'====================================
'VALIDATION FOR DATE OF FCOBERT
'====================================
If FCOBER.TextLength > 1 And FCOBER.TextLength < 3 Then
    FCOBER.Value = FCOBER.Value & "/"
End If
If FCOBER.TextLength > 10 Then
    FCOBER.Value = Mid(FCOBER.Text, 1, Len(FCOBER.Text) - 1)
End If
If FCOBER.TextLength > 4 And FCOBER.TextLength < 6 Then
    FCOBER.Value = FCOBER.Value & "/"
End If
End Sub
Private Sub FCOBER_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       FCOBER.Value = vbNullString
    End If
End Sub
Private Sub FCOBER_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FCOBER.TextLength > 1 And FCOBER.TextLength < 10 And FCOBER.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    FCOBER.Value = vbNullString
    FCOBER.SetFocus
    Exit Sub
End If
'====================================
'END VALIDATION FOR DATE OF FCOBERT
'==================================
End Sub
Private Sub EXADATE_Change()
'====================================
'VALIDATION FOR DATE OF EXADATET
'====================================
If EXADATE.TextLength > 1 And EXADATE.TextLength < 3 Then
    EXADATE.Value = EXADATE.Value & "/"
End If
If EXADATE.TextLength > 10 Then
    EXADATE.Value = Mid(EXADATE.Text, 1, Len(EXADATE.Text) - 1)
End If
If EXADATE.TextLength > 4 And EXADATE.TextLength < 6 Then
    EXADATE.Value = EXADATE.Value & "/"
End If
End Sub
Private Sub EXADATE_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       EXADATE.Value = vbNullString
    End If
End Sub
Private Sub EXADATE_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If EXADATE.TextLength > 1 And EXADATE.TextLength < 10 And EXADATE.Text <> vbNullString Then
    MsgBox "ingrese fecha en formato DD/MM/AAAA"
    EXADATE.Value = vbNullString
    EXADATE.SetFocus
    Exit Sub
End If
'====================================
'END VALIDATION FOR DATE OF EXADATET
'====================================
End Sub
Private Sub SBASE_Change()
   SBASE.Value = SBASE.Value
   If Not IsNumeric(SBASE.Value) And SBASE.Value <> vbNullString Then
   MsgBox "Solo se permiten números"
   SBASE.Value = Mid(SBASE.Value, 1, Len(SBASE.Value) - 1)
   'Me.Lwage.Value =
   End If
End Sub
Private Sub SBASE_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyBack Then
       SBASE.Value = vbNullString
    End If
End Sub
Private Sub RODAMIENTO_Change()
    RODAMIENTO.Value = RODAMIENTO.Value
    If Not IsNumeric(RODAMIENTO.Value) And RODAMIENTO.Value <> vbNullString Then
        MsgBox "Solo se permiten números"
        RODAMIENTO.Value = Mid(RODAMIENTO.Value, 1, Len(RODAMIENTO.Value) - 1)
    End If
End Sub
Private Sub RODAMIENTO_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyBack Then
       RODAMIENTO.Value = vbNullString
    End If
End Sub
Private Sub OAUX_Change()
    OAUX.Value = OAUX.Value
    If Not IsNumeric(OAUX.Value) And OAUX.Value <> vbNullString Then
        MsgBox "Solo se permiten números"
        OAUX.Value = Mid(OAUX.Value, 1, Len(OAUX.Value) - 1)
    End If
End Sub
Private Sub OAUX_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyBack Then
       OAUX.Value = vbNullString
    End If
End Sub

'When close then select sheets principal
Private Sub UserForm_QueryClose(CANCEL As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        CANCEL = True
        FGPersonal.Hide
        Sheets("PPrincipal").Select
    End If
End Sub

Sub AddEmpInformation()
Call OptimizeCode_Begin
'==================================
If FGPersonal.NOMBRE.Value = vbNullString Then
    MsgBox "Ingrese el nombre completo"
    MultiPage1.Value = 0
    NOMBRE.SetFocus
    Exit Sub
End If
If FGPersonal.IDENTIFICACION.Value = vbNullString Then
    MsgBox "Ingrese Identificación"
    MultiPage1.Value = 0
    IDENTIFICACION.SetFocus
    Exit Sub
End If
If FGPersonal.RH.Value = vbNullString Then
    MsgBox "Ingrese RH"
MultiPage1.Value = 0
RH.SetFocus
Exit Sub
End If

If FGPersonal.FDN.Value = vbNullString Then
MsgBox "Ingrese fecha de nacimiento"
MultiPage1.Value = 0
FDN.SetFocus
Exit Sub
End If

If FGPersonal.CIVILS = vbNullString Then
MsgBox "Seleccione un Estado Civil"
MultiPage1.Value = 0
FDN.SetFocus
Exit Sub
End If

If FGPersonal.DEGREE = vbNullString Then
MsgBox "Seleccione un Grado de Escolaridad"
MultiPage1.Value = 0
FDN.SetFocus
Exit Sub
End If

If FGPersonal.CIUDAD.Value = vbNullString Then
MsgBox "Ingrese Ciudad"
MultiPage1.Value = 0
CIUDAD.SetFocus
Exit Sub
End If

If FGPersonal.DIRECCION.Value = vbNullString Then
MsgBox "Ingrese Dirección"
MultiPage1.Value = 0
DIRECCION.SetFocus
Exit Sub
End If

If FGPersonal.EMAILP.Value = vbNullString Then
MsgBox "Ingrese E-mail personal"
MultiPage1.Value = 0
EMAILP.SetFocus
Exit Sub
End If

If FGPersonal.TELMP.Value = vbNullString Then
MsgBox "Ingrese teléfono móvil"
MultiPage1.Value = 0
TELMP.SetFocus
Exit Sub
End If

If FGPersonal.TELFP.Value = vbNullString Then
MsgBox "Ingrese teléfono fijo"
MultiPage1.Value = 0
TELFP.SetFocus
Exit Sub
End If

If FGPersonal.EMAILCOR.Value = vbNullString Then
MsgBox "Ingrese E-mail corporativo"
MultiPage1.Value = 1
EMAILCOR.SetFocus
Exit Sub
End If

If FGPersonal.TELMC.Value = vbNullString Then
MsgBox "Ingrese teléfono móvil corporativo"
MultiPage1.Value = 1
TELMC.SetFocus
Exit Sub
End If

If FGPersonal.TELFC.Value = vbNullString Then
MsgBox "Ingrese teléfono fijo corporativo"
MultiPage1.Value = 1
TELFC.SetFocus
Exit Sub
End If

If FGPersonal.FDI.Value = vbNullString Then
MsgBox "Ingrese fecha de ingreso"
MultiPage1.Value = 1
FDI.SetFocus
Exit Sub
End If

If FGPersonal.DEP.Value = vbNullString Then
MsgBox "Ingrese departamento"
MultiPage1.Value = 1
DEP.SetFocus
Exit Sub
End If

If FGPersonal.CDEP.Value = vbNullString Then
MsgBox "Ingrese código de departamento"
MultiPage1.Value = 1
CDEP.SetFocus
Exit Sub
End If

If FGPersonal.CARGO.Value = vbNullString Then
MsgBox "Ingrese cargo"
MultiPage1.Value = 1
CARGO.SetFocus
Exit Sub
End If

If FGPersonal.TCONTRATO.Value = vbNullString Then
MsgBox "Ingrese tipo de contrato"
MultiPage1.Value = 1
TCONTRATO.SetFocus
Exit Sub
End If

If FGPersonal.SBASE.Value = vbNullString Then
MsgBox "Ingrese salario base"
MultiPage1.Value = 2
SBASE.SetFocus
Exit Sub
End If

If FGPersonal.RODAMIENTO.Value = vbNullString Then
MsgBox "Ingrese auxilio de rodamiento"
MultiPage1.Value = 2
RODAMIENTO.SetFocus
Exit Sub
End If

If FGPersonal.OAUX.Value = vbNullString Then
MsgBox "Ingrese otros auxilios"
MultiPage1.Value = 2
OAUX.SetFocus
Exit Sub
End If

If FGPersonal.EPS.Value = vbNullString Then
MsgBox "Ingrese EPS"
MultiPage1.Value = 2
EPS.SetFocus
Exit Sub
End If

If FGPersonal.AFP.Value = vbNullString Then
MsgBox "Ingrese fondo de pensiones"
MultiPage1.Value = 2
AFP.SetFocus
Exit Sub
End If

If FGPersonal.CCF.Value = vbNullString Then
MsgBox "Ingrese Caja de compensación"
MultiPage1.Value = 2
CCF.SetFocus
Exit Sub
End If

If FGPersonal.ARL.Value = vbNullString Then
MsgBox "Ingrese ARL"
MultiPage1.Value = 2
ARL.SetFocus
Exit Sub
End If

If FGPersonal.WORKC.Value = vbNullString Then
MsgBox "Ingrese Centro de Trabajo"
MultiPage1.Value = 2
WORKC.SetFocus
Exit Sub
End If

If FGPersonal.CLASS.Value = vbNullString Then
MsgBox "Ingrese Clase"
MultiPage1.Value = 2
CLASS.SetFocus
Exit Sub
End If

If FGPersonal.FARE.Value = vbNullString Then
MsgBox "Ingrese Tasa"
MultiPage1.Value = 2
FARE.SetFocus
Exit Sub
End If

If FGPersonal.FCOBER.Value = vbNullString Then
MsgBox "Ingrese Fecha de Cobertura"
MultiPage1.Value = 2
FCOBER.SetFocus
Exit Sub
End If

If FGPersonal.TEXAM.Value = vbNullString Then
MsgBox "Ingrese Tipo de Exámen"
MultiPage1.Value = 3
TEXAM.SetFocus
Exit Sub
End If

If FGPersonal.EXADATE.Value = vbNullString Then
MsgBox "Ingrese Fecha de Exámen"
MultiPage1.Value = 3
EXADATE.SetFocus
Exit Sub
End If

If FGPersonal.HCON.Value = vbNullString Then
MsgBox "Ingrese Condicion de Salud"
MultiPage1.Value = 3
HCON.SetFocus
Exit Sub
End If

If FGPersonal.RECOM.Value = vbNullString Then
MsgBox "Ingrese Recomendaciones"
MultiPage1.Value = 3
RECOM.SetFocus
Exit Sub
End If

If FGPersonal.REST.Value = vbNullString Then
MsgBox "Ingrese Restricciones"
MultiPage1.Value = 3
REST.SetFocus
Exit Sub
End If

' Variables to find the last row with information
Dim duplicados As Boolean
'=======================================
'START: REGISTER DATA ON VACATION SHEET
'=======================================
Dim wsVa As Worksheet
Set wsVa = Sheets("VData")

duplicados = False
'loop, and conditional to register data

'Add next information

wsVa.Cells(EmpRowStateVac, VData.[Vdata_Enterprise].Column).Value = EnterpriseState
wsVa.Cells(EmpRowStateVac, VData.[vac_name].Column).Value = FGPersonal.NOMBRE.Value
wsVa.Cells(EmpRowStateVac, VData.[vac_id].Column).Value = FGPersonal.IDENTIFICACION.Value
wsVa.Cells(EmpRowStateVac, VData.[vac_department_code].Column).Value = FGPersonal.CDEP.Value
wsVa.Cells(EmpRowStateVac, VData.[vac_jobname].Column).Value = FGPersonal.CARGO.Value
'Validate Tcontrato to enable Fcontratoi
Dim indefined As Boolean 'to enable field
    If FGPersonal.TCONTRATO.Value = "INDEFINIDO" Then
        If FGPersonal.FCONTRATOI.Value = vbNullString Then 'DO NOT ALLOW BLANK FIELD
            MsgBox "Ingrese fecha de contrato indefinido"
            MultiPage1.Value = 1
            FCONTRATOI.SetFocus
         Exit Sub
        End If
    End If
If Me.FCONTRATOI.Value = "" Then
    wsVa.Cells(EmpRowStateVac, VData.[vac_und_contract_dated].Column).Value = ""
Else
    wsVa.Cells(EmpRowStateVac, VData.[vac_und_contract_dated].Column).Value = CDate(Me.FCONTRATOI.Value)
End If

wsVa.Cells(EmpRowStateVac, VData.[vac_wage].Column).Value = FGPersonal.SBASE.Value
wsVa.Cells(EmpRowStateVac, VData.[vac_liquidation_dated].Column).Value = Date
'VALIDATE IF DATE OF INDEFINED CONTRAT IS NA OR N
wsVa.Cells(EmpRowStateVac, VData.[vac_worked_days].Column).FormulaR1C1 = "=IF(RC[-3]="""",0,DAYS360(RC[-3],RC[-1]))"
wsVa.Cells(EmpRowStateVac, VData.[vac_days_emp].Column).FormulaR1C1 = "=RC[-1]*0.0417"
wsVa.Cells(EmpRowStateVac, VData.[vac_days_emp_bef].Column).FormulaR1C1 = 0
'Formula to get the sumatory of the vacation days
wsVa.Cells(EmpRowStateVac, VData.[vac_taken_days].Column).Formula = _
    "=SUMIF(AData!C[-7],VData!RC[5],AData!C[2])+RC[-1]"
'----------------------------------------------------------
wsVa.Cells(EmpRowStateVac, VData.[vac_days_aval].Column).FormulaR1C1 = "=RC[-3]-RC[-1]"
wsVa.Cells(EmpRowStateVac, VData.[vac_cost].Column).FormulaR1C1 = "=(RC[-1]*RC[-7])/30"
wsVa.Cells(EmpRowStateVac, VData.[vac_state].Column).Value = FGPersonal.Retirado.Value
    If FGPersonal.TCONTRATO.Value <> "INDEFINIDO" Then
        wsVa.Cells(EmpRowStateVac, VData.[vac_contract].Column).Value = False
    Else
        wsVa.Cells(EmpRowStateVac, VData.[vac_contract].Column).Value = True
    End If
    wsVa.Cells(EmpRowStateVac, VData.[vac_utility].Column).Value = FGPersonal.NOMBRE.Value & "V."
'=====================================
'END: REGISTER DATA ON VACATION SHEET

duplicados = False

'loop, and conditional to register data
Dim O As Long
'For O = 1 To EmpRowState
'
'    If Cells(O, wsPD.[EMPNAME].Column).Value = FGPersonal.NOMBRE.Value Then
'        If FGPersonal.ComboBox1.Value > 1 Then
'            MsgBox "Dato duplicado o Error: Elimine los datos del buscador"
'            duplicados = True
'        End If
'    End If

'Next O

'If the condition is accomplish then continue with the code

If Not duplicados Then

'Uncheck Retirado Box

'FGPersonal.Retirado = False

'assign data to cells
Dim wsP As Worksheet
Set wsP = Sheets("PData")
wsP.Cells(EmpRowState, wsPD.[Pdata_emp_enterprise].Column).Value = EnterpriseState
wsP.Cells(EmpRowState, wsPD.[EMPNAME].Column).Value = FGPersonal.NOMBRE.Value
wsP.Cells(EmpRowState, wsPD.[DATEDEXP].Column).Value = CDate(Me.DATEDEXP.Value)
wsP.Cells(EmpRowState, wsPD.[PLACEEXP].Column).Value = Me.PLACEEXP.Value
wsP.Cells(EmpRowState, wsPD.[ID].Column).Value = FGPersonal.IDENTIFICACION.Value
wsP.Cells(EmpRowState, wsPD.[BLOODT].Column).Value = FGPersonal.RH.Value
If MALE = True Then 'Validate if is male or female
        wsP.Cells(EmpRowState, wsPD.[GENDER].Column).Value = "MASCULINO"
    ElseIf FEMALE = True Then
        wsP.Cells(EmpRowState, wsPD.[GENDER].Column).Value = "FEMENINO"
End If
wsP.Cells(EmpRowState, wsPD.[CIVILSTATUS].Column).Value = CIVILS.Value
wsP.Cells(EmpRowState, wsPD.[DEGREE].Column).Value = DEGREE.Value
wsP.Cells(EmpRowState, wsPD.[p_proffession].Column).Value = Me.emp_proffession.Value
wsP.Cells(EmpRowState, wsPD.[e_proffesional_card].Column).Value = Me.emp_proffession_card.Value
wsP.Cells(EmpRowState, wsPD.[DATEDOB].Column).Value = IIf(Me.FDN.Value = Empty, "", CDate(Me.FDN.Value))
wsP.Cells(EmpRowState, wsPD.[EAGE].Column).FormulaR1C1 = "=((TODAY()-RC[-1])/365)"
wsP.Cells(EmpRowState, wsPD.[CITY].Column).Value = FGPersonal.CIUDAD.Value
wsP.Cells(EmpRowState, wsPD.[EADDRESS].Column).Value = FGPersonal.DIRECCION.Value
wsP.Cells(EmpRowState, wsPD.[NHOOD].Column).Value = Me.NHoodList.Value
wsP.Cells(EmpRowState, wsPD.[DISTRICT].Column).Value = Me.LISTDISTRICT.Value
wsP.Cells(EmpRowState, wsPD.[EMAILP].Column).Value = FGPersonal.EMAILP.Value
wsP.Cells(EmpRowState, wsPD.[EPHONEM].Column).Value = FGPersonal.TELMP.Value
wsP.Cells(EmpRowState, wsPD.[EPHONES].Column).Value = FGPersonal.TELFP.Value
wsP.Cells(EmpRowState, wsPD.[EMAILCO].Column).Value = FGPersonal.EMAILCOR.Value
wsP.Cells(EmpRowState, wsPD.[PHONEMC].Column).Value = FGPersonal.TELMC.Value
wsP.Cells(EmpRowState, wsPD.[PHONESC].Column).Value = FGPersonal.TELFC.Value
wsP.Cells(EmpRowState, wsPD.[DOI].Column).Value = CDate(FGPersonal.FDI.Value)
wsP.Cells(EmpRowState, wsPD.[TIC].Column).FormulaR1C1 = _
    "=IF(RC[-1]="""","""",IF(RC[21]=FALSE,CONCATENATE(DATEDIF(RC[-1],TODAY(),""Y""),"" "",""AÑOS"","" - "",DATEDIF(RC[-1],TODAY(),""YM""),"" "",""MESES""),CONCATENATE(DATEDIF(RC[-1],RC[13],""Y""),"" "",""AÑOS"",""-"",DATEDIF(RC[-1],RC[13],""YM""),"" "",""MESES"")))"
wsP.Cells(EmpRowState, wsPD.[DEPARTCODE].Column).Value = FGPersonal.CDEP.Value
wsP.Cells(EmpRowState, wsPD.[DEPARTNAME].Column).Value = FGPersonal.DEP.Value
wsP.Cells(EmpRowState, wsPD.[JOBNAME].Column).Value = FGPersonal.CARGO.Value
wsP.Cells(EmpRowState, wsPD.[TContract].Column).Value = FGPersonal.TCONTRATO.Value
wsP.Cells(EmpRowState, wsPD.[wage].Column).Value = FGPersonal.SBASE.Value
wsP.Cells(EmpRowState, wsPD.[Auxi1].Column).Value = FGPersonal.RODAMIENTO.Value
wsP.Cells(EmpRowState, wsPD.[Auxi2].Column).Value = FGPersonal.OAUX.Value
wsP.Cells(EmpRowState, wsPD.[EPS].Column).Value = FGPersonal.EPS.Value
wsP.Cells(EmpRowState, wsPD.[AFP].Column).Value = FGPersonal.AFP.Value
wsP.Cells(EmpRowState, wsPD.[CCF].Column).Value = FGPersonal.CCF.Value
wsP.Cells(EmpRowState, wsPD.[ARL].Column).Value = FGPersonal.ARL.Value
wsP.Cells(EmpRowState, wsPD.[JOBCENTER].Column).Value = WORKC.Value
wsP.Cells(EmpRowState, wsPD.[RISKCLASS].Column).Value = CLASS.Value
wsP.Cells(EmpRowState, wsPD.[FARE].Column).Value = FARE.Value
wsP.Cells(EmpRowState, wsPD.[doc].Column).Value = CDate(FCOBER.Value)
If Me.FRARL.Value = "" Then
    wsP.Cells(EmpRowState, wsPD.[DATERARL].Column).Value = Me.FRARL
Else
    wsP.Cells(EmpRowState, wsPD.[DATERARL].Column).Value = CDate(Me.FRARL)
End If
wsP.Cells(EmpRowState, wsPD.[LASTME].Column).Value = TEXAM.Value
wsP.Cells(EmpRowState, wsPD.[p_exadated].Column).Value = CDate(EXADATE.Value)
wsP.Cells(EmpRowState, wsPD.[MEDICALCON].Column).Value = HCON.Value
wsP.Cells(EmpRowState, wsPD.[RECOM].Column).Value = RECOM.Value
wsP.Cells(EmpRowState, wsPD.[RESTRICTIONS].Column).Value = REST.Value
If Me.Retirado = False Then
        wsP.Cells(EmpRowState, wsPD.[RETIRED].Column).Value = False
    Else
        wsP.Cells(EmpRowState, wsPD.[RETIRED].Column).Value = True
End If
    'wsp.Cells(fila, 40).Value = FGPersonal.FechaR.Value
If Me.FechaR.Value = "" Then
    wsP.Cells(EmpRowState, wsPD.[DATEDOR].Column).Value = ""
    Else
    wsP.Cells(EmpRowState, wsPD.[DATEDOR].Column).Value = CDate(Me.FechaR)
End If

wsP.Cells(EmpRowState, wsPD.[CAUSEOFRET].Column).Value = FGPersonal.MotivoR.Value
'Insert data regarding social affiliations status
wsP.Cells(EmpRowState, wsPD.[EPSS].Column).Value = Me.EPSS.Value
wsP.Cells(EmpRowState, wsPD.[EPSBE].Column).Value = Me.EPSD.Value
wsP.Cells(EmpRowState, wsPD.[EPSOB].Column).Value = Me.EPSO.Value

wsP.Cells(EmpRowState, wsPD.[AFPS].Column).Value = Me.AFPS.Value
wsP.Cells(EmpRowState, wsPD.[AFPOB].Column).Value = Me.AFPO.Value

wsP.Cells(EmpRowState, wsPD.[CCFS].Column).Value = Me.CCFS.Value
wsP.Cells(EmpRowState, wsPD.[CCFBE].Column).Value = Me.CCFD.Value
wsP.Cells(EmpRowState, wsPD.[CCFOB].Column).Value = Me.CCFO.Value

wsP.Cells(EmpRowState, wsPD.[ARLS].Column).Value = Me.ARLS.Value
wsP.Cells(EmpRowState, wsPD.[ARLOB].Column).Value = Me.ARLO.Value
'Success Message
'Add contract info
wsP.Cells(EmpRowState, wsPD.[CONTRACTSTATE].Column).Value = Me.CONTRACTSTATE.Value
wsP.Cells(EmpRowState, wsPD.[CONTRACTDATEDR].Column).Value = IIf(IsNull(Me.CONTRACTDATEDR.Value), "", CDate(Me.CONTRACTDATEDR.Value))

'Add Employee_branch
wsP.Cells(EmpRowState, wsPD.[e_branch].Column).Value = Me.e_branch.Value


'Empty RowState

EmpRowState = 0
EmpRowStateVac = 0

'Add dependents

'ADD INFO FOR DEPENDENTS

DepSheet.Cells(DepSRowState, DepSheet.[dep_enterprise].Column).Value = EnterpriseState
DepSheet.Cells(DepSRowState, DepSheet.[APELLIDOS_Y_NOMBRES].Column).Value = Me.NOMBRE.Value
DepSheet.Cells(DepSRowState, DepSheet.[IDENTIFICACION].Column).Value = FGPersonal.IDENTIFICACION.Value

Call AddDependentsInfo

'finish code
'======================================
'CALL NEWS REPORT FOR ACCOUNTING REPORT
'=======================================
End If
End Sub


Sub AddDependentsInfo()
'FIRST DEPENDENT
DepSheet.Cells(DepSRowState, DepSheet.[DEREL1].Column).Value = Me.DREL1.Value
DepSheet.Cells(DepSRowState, DepSheet.[DETID1].Column).Value = Me.DTID1.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEID1].Column).Value = Me.DID1.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEFN1].Column).Value = Me.DFN1.Value
DepSheet.Cells(DepSRowState, DepSheet.[DESN1].Column).Value = Me.DSN1.Value
DepSheet.Cells(DepSRowState, DepSheet.[DELN1].Column).Value = Me.DELN1.Value
DepSheet.Cells(DepSRowState, DepSheet.[DESLN1].Column).Value = Me.DESLN1.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEDOB1].Column).Value = Me.DDOB1.Value
DepSheet.Cells(DepSRowState, DepSheet.[DCIVILR1].Column).Value = Me.DCIVILR1.Value
DepSheet.Cells(DepSRowState, DepSheet.[TICC1].Column).Value = Me.TICC1.Value
DepSheet.Cells(DepSRowState, DepSheet.[STUCER1].Column).Value = Me.STUCER1.Value
DepSheet.Cells(DepSRowState, DepSheet.[MSUPPORT1].Column).Value = Me.MSUPPORT1.Value
If Me.DM1 = True Then
    DepSheet.Cells(DepSRowState, DepSheet.[DEGEN1].Column).Value = "M"
Else
    DepSheet.Cells(DepSRowState, DepSheet.[DEGEN1].Column).Value = "F"
End If
'SECOND DEPENDENT
DepSheet.Cells(DepSRowState, DepSheet.[DEREL2].Column).Value = Me.DREL2.Value
DepSheet.Cells(DepSRowState, DepSheet.[DETID2].Column).Value = Me.DTID2.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEID2].Column).Value = Me.DID2.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEFN2].Column).Value = Me.DFN2.Value
DepSheet.Cells(DepSRowState, DepSheet.[DESN2].Column).Value = Me.DSN2.Value
DepSheet.Cells(DepSRowState, DepSheet.[DELN2].Column).Value = Me.DELN2.Value
DepSheet.Cells(DepSRowState, DepSheet.[DESLN2].Column).Value = Me.DESLN2.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEDOB2].Column).Value = Me.DDOB2.Value
DepSheet.Cells(DepSRowState, DepSheet.[DCIVILR2].Column).Value = Me.DCIVILR2.Value
DepSheet.Cells(DepSRowState, DepSheet.[TICC2].Column).Value = Me.TICC2.Value
DepSheet.Cells(DepSRowState, DepSheet.[STUCER2].Column).Value = Me.STUCER2.Value
DepSheet.Cells(DepSRowState, DepSheet.[MSUPPORT2].Column).Value = Me.MSUPPORT2.Value
If Me.DM2 = True Then
    DepSheet.Cells(DepSRowState, DepSheet.[DEGEN2].Column).Value = "M"
Else
    DepSheet.Cells(DepSRowState, DepSheet.[DEGEN2].Column).Value = "F"
End If
'THIRD DEPENDENT
DepSheet.Cells(DepSRowState, DepSheet.[DEREL3].Column).Value = Me.DREL3.Value
DepSheet.Cells(DepSRowState, DepSheet.[DETID3].Column).Value = Me.DTID3.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEID3].Column).Value = Me.DID3.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEFN3].Column).Value = Me.DFN3.Value
DepSheet.Cells(DepSRowState, DepSheet.[DESN3].Column).Value = Me.DSN3.Value
DepSheet.Cells(DepSRowState, DepSheet.[DELN3].Column).Value = Me.DELN3.Value
DepSheet.Cells(DepSRowState, DepSheet.[DESLN3].Column).Value = Me.DESLN3.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEDOB3].Column).Value = Me.DDOB3.Value
DepSheet.Cells(DepSRowState, DepSheet.[DCIVILR3].Column).Value = Me.DCIVILR3.Value
DepSheet.Cells(DepSRowState, DepSheet.[TICC3].Column).Value = Me.TICC3.Value
DepSheet.Cells(DepSRowState, DepSheet.[STUCER3].Column).Value = Me.STUCER3.Value
DepSheet.Cells(DepSRowState, DepSheet.[MSUPPORT3].Column).Value = Me.MSUPPORT3.Value
If Me.DM3 = True Then
    DepSheet.Cells(DepSRowState, DepSheet.[DEGEN3].Column).Value = "M"
Else
    DepSheet.Cells(DepSRowState, DepSheet.[DEGEN3].Column).Value = "F"
End If
'FORTH DEPENDENT
DepSheet.Cells(DepSRowState, DepSheet.[DEREL4].Column).Value = Me.DREL4.Value
DepSheet.Cells(DepSRowState, DepSheet.[DETID4].Column).Value = Me.DTID4.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEID4].Column).Value = Me.DID4.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEFN4].Column).Value = Me.DFN4.Value
DepSheet.Cells(DepSRowState, DepSheet.[DESN4].Column).Value = Me.DSN4.Value
DepSheet.Cells(DepSRowState, DepSheet.[DELN4].Column).Value = Me.DELN4.Value
DepSheet.Cells(DepSRowState, DepSheet.[DESLN4].Column).Value = Me.DESLN4.Value
DepSheet.Cells(DepSRowState, DepSheet.[DEDOB4].Column).Value = Me.DDOB4.Value
DepSheet.Cells(DepSRowState, DepSheet.[DCIVILR4].Column).Value = Me.DCIVILR4.Value
DepSheet.Cells(DepSRowState, DepSheet.[TICC4].Column).Value = Me.TICC4.Value
DepSheet.Cells(DepSRowState, DepSheet.[STUCER4].Column).Value = Me.STUCER4.Value
DepSheet.Cells(DepSRowState, DepSheet.[MSUPPORT4].Column).Value = Me.MSUPPORT4.Value
If Me.DM4 = True Then
    DepSheet.Cells(DepSRowState, DepSheet.[DEGEN4].Column).Value = "M"
Else
    DepSheet.Cells(DepSRowState, DepSheet.[DEGEN4].Column).Value = "F"
End If
End Sub

Sub TrackingRetirementInfo()
PData.Cells(RetRowState, wsPD.[ReDATEDR].Column).Value = Me.ReDATEDR.Value
PData.Cells(RetRowState, wsPD.[PNS].Column).Value = Me.PNS.Value
PData.Cells(RetRowState, wsPD.[RLETTERS].Column).Value = Me.RLETTERS.Value
PData.Cells(RetRowState, wsPD.[LIQUIDATION].Column).Value = Me.LIQUIDATION.Value
PData.Cells(RetRowState, wsPD.[retOBS].Column).Value = Me.retOBS.Value
End Sub
Sub NewsReportS()


Set ShNewsReport = Sheets("News_P")
LastrowNR = ShNewsReport.Cells(Rows.Count, 1).End(xlUp).Row + 1
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_enterprise].Column).Value = Me.Enterprise
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_no].Column).Value = LastrowNR - 1
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_emp_name].Column).Value = Me.NOMBRE.Value
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_id].Column).Value = Me.IDENTIFICACION.Value
Select Case NewsType
    Case 1: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_incorporation].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(Me.FDI.Value)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(Me.FDI.Value)
    
    Case 2: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_retired].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(Me.FechaR.Value)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(Me.FechaR.Value)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_observation].Column).Value = Me.MotivoR.Value
    
    'PAYROLLL UPDATES
    
    Case 31: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_contract].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[prevreg].Column).Value = wsPD.Cells(lRow, wsPD.[DEPARTNAME].Column)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nextreg].Column).Value = Me.DEP.Value
       
    Case 32: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_jobname].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[prevreg].Column).Value = wsPD.Cells(lRow, wsPD.[JOBNAME].Column)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nextreg].Column).Value = Me.CARGO.Value
    
    Case 33: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_contract].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[prevreg].Column).Value = wsPD.Cells(lRow, wsPD.[TContract].Column)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nextreg].Column).Value = Me.TCONTRATO.Value
    
    Case 34: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_wage].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[prevreg].Column).Value = wsPD.Cells(lRow, wsPD.[wage].Column)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nextreg].Column).Value = Me.SBASE.Value
    
    Case 35: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_auxi1].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[prevreg].Column).Value = wsPD.Cells(lRow, wsPD.[Auxi1].Column)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nextreg].Column).Value = Me.RODAMIENTO.Value
    
    Case 36: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_auxi2].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[prevreg].Column).Value = wsPD.Cells(lRow, wsPD.[Auxi2].Column)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nextreg].Column).Value = Me.OAUX.Value
    
    Case 37: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_eps].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[prevreg].Column).Value = wsPD.Cells(lRow, wsPD.[EPS].Column)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nextreg].Column).Value = Me.EPS.Value
    
    Case 38: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_afp].Column).Value = "X"
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(reportdate)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[prevreg].Column).Value = wsPD.Cells(lRow, wsPD.[AFP].Column)
    ShNewsReport.Cells(LastrowNR, ShNewsReport.[nextreg].Column).Value = Me.AFP.Value

End Select
'ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(Me.FDI.Value)

End Sub
