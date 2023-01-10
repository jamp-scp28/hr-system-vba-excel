VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RAusentismos 
   Caption         =   "AUSENTISMOS"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
   OleObjectBlob   =   "RAusentismos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "RAusentismos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lndDay As Long
Public EG As Boolean
Public DIng As Boolean
Public lngDay As Long
Public EPSN As String 'Get the name of EPS of every employee
'=====================
'Variable to Speed up the code
'=====================
Public EventState As Boolean
Public PageBreakState As Boolean
Public wsPD As Worksheet
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
'=====================
'End Variable to Speed up the code
'=====================
Private Sub B_FINDC_Click()
FINDC.Show
End Sub
Private Sub BUSCADOR_Change()
'==DISABLED UPDATE BUTTON TO AVOID ERR
If BUSCADOR.ListIndex > -1 And BUSCADOR.Value <> vbNullString Then
    Me.RAUSENTISMO.Enabled = True
    Else
    Me.RAUSENTISMO.Enabled = False
End If
' Variables for look up
Dim myRange As Range
Set myRange = Worksheets("PData").Range("b:bl")
On Error Resume Next
Set wsPD = Sheets("PData")
'VlookUp the values of the boxes
'MsgBox wsPD.[EMPNAME].Column

Me.Enterprise.Value = _
Application.WorksheetFunction.Index(wsPD.Range("A:A"), Application.WorksheetFunction.Match(Me.BUSCADOR.Value, wsPD.Range("B:B"), 0))
If Err.Number <> 0 Then Me.Enterprise.Value = ""

NOMBRES.Value = _
Application.WorksheetFunction.VLookup(BUSCADOR, myRange, wsPD.[EMPNAME].Column - 1, False)
If Err.Number <> 0 Then NOMBRES.Value = vbNullString

IDENTIFICACION.Value = _
Application.WorksheetFunction.VLookup(BUSCADOR, myRange, wsPD.[ID].Column - 1, False)
If Err.Number <> 0 Then IDENTIFICACION.Value = vbNullString

CDEP.Value = _
Application.WorksheetFunction.VLookup(BUSCADOR, myRange, wsPD.[DEPARTCODE].Column - 1, False)
If Err.Number <> 0 Then CDEP.Value = vbNullString

CARGO.Value = _
Application.WorksheetFunction.VLookup(BUSCADOR, myRange, wsPD.[JOBNAME].Column - 1, False)
If Err.Number <> 0 Then CARGO.Value = vbNullString

SBASE.Value = _
Application.WorksheetFunction.VLookup(BUSCADOR, myRange, wsPD.[wage].Column - 1, False)
If Err.Number <> 0 Then SBASE.Value = vbNullString

EPSN = _
Application.WorksheetFunction.VLookup(BUSCADOR, myRange, wsPD.[EPS].Column - 1, False)
    If Err.Number <> 0 Then EPSN = ""
'==========================
'DISABLED SEARCH FOR ID
'==========================
If BUSCADOR.ListIndex > 0 Or BUSCADOR.Value <> vbNullString Then
    ID_A.Enabled = False
    Else
    ID_A.Enabled = True
End If
End Sub
'==============================
Private Sub CANCELAR_Click()
Call ClearDataF
Unload Me
Sheets("PPrincipal").Select
End Sub
Private Sub CONVENCIONES_Change()
'======================================
'EVALUATE IF THE VALUE SELECT IS "E.G."
'=======================================
If CONVENCIONES.Value = "E.G." Or CONVENCIONES.Value = "A.T." Or CONVENCIONES.Value = "L.M." Then
    VC_Selected.Enabled = True
    B_FINDC.Enabled = True
    EG = True
    Else
    VC_Selected.Enabled = False
    B_FINDC.Enabled = False
    EG = False
End If
End Sub


Private Sub DESCRIPCION_Change()
DESCRIPCION.Value = UCase(DESCRIPCION)
End Sub

Private Sub FECHAA_Change()
'VALIDATION FOR DATE OF REGISTER------------------
If FECHAA.TextLength > 1 And FECHAA.TextLength < 3 Then
FECHAA.Value = FECHAA.Value & "/"
End If

If FECHAA.TextLength > 10 Then
    FECHAA.Value = Mid(FECHAA.Text, 1, Len(FECHAA.Text) - 1)
End If

If FECHAA.TextLength > 4 And FECHAA.TextLength < 6 Then
FECHAA.Value = FECHAA.Value & "/"
End If

End Sub
Private Sub FECHAA_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode.Value = vbKeyBack Then
       FECHAA.Value = vbNullString
    End If
End Sub

Private Sub FECHAA_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FECHAA.TextLength > 1 And FECHAA.TextLength < 10 And FECHAA.Value <> vbNullString Then
MsgBox "Ingrese una fecha en formato DD/MM/AAAA"
FECHAA.Value = vbNullString
CANCEL = True
End If
'END OF VALIDATION OF DATE OF REGISTER------------------
End Sub
'VALIDATION OF INITIAL DATE
'==

Private Sub FECHAI_Change()
If OHOUR = True Then 'IN CASE THE HOUR NEED TO BE MODIFY

If FECHAI.TextLength > 1 And FECHAI.TextLength < 3 Then
    FECHAI.Value = FECHAI.Value & "/"
End If
If FECHAI.TextLength > 4 And FECHAI.TextLength < 6 Then
    FECHAI.Value = FECHAI.Value & "/"
End If
If FECHAI.TextLength > 9 And FECHAI.TextLength < 11 Then
    FECHAI.Value = FECHAI.Value & " "
End If
If FECHAI.TextLength > 12 And FECHAI.TextLength < 14 Then
    FECHAI.Value = FECHAI.Value & ":"
End If
Else 'IF NO NEED TO BE DIFFERENT THAN 7 A.M. THEN
If FECHAI.TextLength > 1 And FECHAI.TextLength < 3 Then
    FECHAI.Value = FECHAI.Value & "/"
End If
If FECHAI.TextLength > 4 And FECHAI.TextLength < 6 Then
    FECHAI.Value = FECHAI.Value & "/"
End If
If FECHAI.TextLength > 9 And FECHAI.TextLength < 11 Then
    FECHAI.Value = FECHAI.Value & " "
End If
If FECHAI.TextLength > 10 And FECHAI.TextLength < 12 Then
    FECHAI.Value = FECHAI.Value & "0"
End If
If FECHAI.TextLength > 11 And FECHAI.TextLength < 13 Then
    FECHAI.Value = FECHAI.Value & "7"
End If
If FECHAI.TextLength > 12 And FECHAI.TextLength < 14 Then
    FECHAI.Value = FECHAI.Value & ":"
End If
If FECHAI.TextLength > 13 And FECHAI.TextLength < AData.[abs_hours].Column Then
    FECHAI.Value = FECHAI.Value & "0"
End If
If FECHAI.TextLength > 14 And FECHAI.TextLength < 16 Then
    FECHAI.Value = FECHAI.Value & "0"
End If

End If
If FECHAI.TextLength > 16 Then
    FECHAI.Value = Mid(FECHAI, 1, Len(FECHAI.Value) - 1)
End If

'Enable FECHAF is FECHAI is ok, to avoid err
If FECHAI.Value <> vbNullString Then
    FECHAF.Enabled = True
Else
    FECHAF.Enabled = False
End If
'===ADD 0 VALUE WHEN THE START DAY IS IN THE MORNING AND THE CODE DELETE THE 0
Dim StrFI
Dim StrFI2
i = 12
StrFI = Mid(FECHAI.Value, 12, 1)
StrFI2 = 0 & StrFI
If DIng = True And StrFI > 2 Then
    FECHAI.Value = Mid(FECHAI.Value, 1, i - 1) & Replace(FECHAI.Value, StrFI, StrFI2, Start:=i)
End If
End Sub
Private Sub FECHAI_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
If FECHAI.TextLength > 1 And FECHAI.TextLength < 16 And FECHAI.Value <> vbNullString Then
    MsgBox "El formato de fecha debe ser DD/MM/AAAA HH:MM"
    FECHAI.Value = vbNullString
    CANCEL = True
End If
End Sub
Private Sub FECHAI_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyBack Then
        FECHAI.Value = vbNullString
    End If
End Sub


'END OF VALIDATION OF INITIAL DATE

'VALIDATION OF FINAL DATE
Private Sub FECHAF_Change()
If FECHAI.Value = vbNullString And FECHAF.Value <> vbNullString Then
MsgBox "Debe diligenciar datos en fecha inicial primeramente"
FECHAF.Value = vbNullString
FECHAI.SetFocus
Else
If FECHAF.TextLength > 1 And FECHAF.TextLength < 3 Then
FECHAF.Value = FECHAF.Value & "/"
End If
If FECHAF.TextLength > 4 And FECHAF.TextLength < 6 Then
FECHAF.Value = FECHAF.Value & "/"
End If
If FECHAF.TextLength > 9 And FECHAF.TextLength < 11 Then
FECHAF.Value = FECHAF.Value & " "
End If
If FECHAF.TextLength > 12 And FECHAF.TextLength < 14 Then
FECHAF.Value = FECHAF.Value & ":"
End If
If FECHAF.TextLength > 16 Then
FECHAF.Value = Mid(FECHAF.Value, 1, Len(FECHAF.Value) - 1)
End If
End If
'===ADD 0 VALUE WHEN THE START DAY IS IN THE MORNING AND THE CODE DELETE THE 0
Dim StrFI
Dim StrFI2
i = 12
StrFI = Mid(FECHAF.Value, 12, 1)
StrFI2 = 0 & StrFI
If DIng = True And StrFI > 2 Then
    FECHAF.Value = Mid(FECHAF.Value, 1, i - 1) & Replace(FECHAF.Value, StrFI, StrFI2, Start:=i)
End If
End Sub
'============================================
Private Sub FECHAF_Exit(ByVal CANCEL As MSForms.ReturnBoolean)
'Declare variable to get hour of FECHAF
Call OptimizeCode_Begin
If FECHAF.TextLength > 1 And FECHAF.TextLength < 16 And FECHAF.Value <> vbNullString Then
    MsgBox "El formato de fecha debe ser DD/MM/AAAA HH:MM"
    FECHAF.Value = vbNullString
    CANCEL = True
End If
If FECHAF.Value > 2 And FECHAF.Value <> vbNullString Then
    Dim FinalHour As Long
    FinalHour = DatePart("h", FECHAF.Value)
    If FinalHour > 18 Then
        MsgBox ("la hora no puede ser mayor a 17:30 p.m.")
        FECHAF.Value = vbNullString
        CANCEL = True
    ElseIf FinalHour < 7 Then
        MsgBox ("la hora no puede ser inferior a 07:00 a.m.")
        FECHAF.Value = Mid(FECHAF.Value, 1, Len(FECHAF.Value) - 6)
        CANCEL = True
    End If
End If
If FECHAI.Value <> vbNullString And FECHAF.Value <> vbNullString Then
lngDay = WorksheetFunction.NetworkDays_Intl(CDate(FECHAI.Value), CDate(FECHAF.Value))
End If
If lngDay > 1 Then
    DABS.Value = lngDay & " DÍA(S)"
    Else
    DABS.Value = lngDay - 1 & " DÍA(S)"
End If
Call OptimizeCode_End
End Sub
Private Sub FECHAF_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyBack Then
        FECHAF.Value = vbNullString
    End If
End Sub
'===================================
Private Sub ID_A_Change()
'==FIND CURRENT ROW
On Error Resume Next
Dim lRow As Long
Dim CurrentRow As Long
lRow = Sheets("AData").Range("A:A").Find(What:=Me.ID_A, after:=Sheets("AData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row

CurrentRow = lRow
'==DISABLED BUTTON TO AVOID ERR
If ID_A.ListIndex > -1 And ID_A.Value <> vbNullString Then
    Me.UPDATEA.Enabled = True
    Else
    Me.UPDATEA.Enabled = False
End If
'=========================================
'GET THE DATA TO MODIFY IT
'=========================================
Dim myRange As Range
Set myRange = Worksheets("AData").Range("A:P")
On Error Resume Next

'VlookUp the values of the boxes

FECHAA.Value = _
CDate(Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_dated].Column, False))
If Err.Number <> 0 Then FECHAA.Value = vbNullString

NOMBRES.Value = _
Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_emp_name].Column, False)
If Err.Number <> 0 Then NOMBRES.Value = vbNullString

IDENTIFICACION.Value = _
Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_emp_id].Column, False)
If Err.Number <> 0 Then IDENTIFICACION.Value = vbNullString

CDEP.Value = _
Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_department].Column, False)
If Err.Number <> 0 Then CDEP.Value = vbNullString

CARGO.Value = _
Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_jobname].Column, False)
If Err.Number <> 0 Then CARGO.Value = vbNullString

Me.SBASE.Value = _
Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_wage].Column, False)
If Err.Number <> 0 Then Me.SBASE.Value = vbNullString

CONVENCIONES.Value = _
Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_type_abs].Column, False)
If Err.Number <> 0 Then CONVENCIONES.Value = vbNullString

If Sheets("ADATA").Cells(CurrentRow, AData.[abs_type_abs].Column).Value = "E.G." Then
    Me.VC_Selected.Enabled = True
    Me.VC_Selected.Value = _
        Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_CIE10].Column, False)
        If Err.Number <> 0 Then Me.VC_Selected.Value = vbNullString
    Else
    Me.VC_Selected.Value = vbNullString
    Me.VC_Selected.Enabled = False
End If

FECHAI.Value = _
CDate(Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_initial_dated].Column, False))
If Err.Number <> 0 Then FECHAI.Value = vbNullString

FECHAF.Value = _
CDate(Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_final_dated].Column, False))
If Err.Number <> 0 Then FECHAF.Value = vbNullString

DESCRIPCION.Value = _
Application.WorksheetFunction.VLookup(ID_A, myRange, AData.[abs_cause].Column, False)
If Err.Number <> 0 Then DESCRIPCION.Value = vbNullString
'========================
'DISABLED "BUSCADOR"
'========================
If ID_A.ListIndex > 0 Or ID_A.Value <> vbNullString Then
    BUSCADOR.Enabled = False
    Else
    BUSCADOR.Enabled = True
End If
DIng = True
End Sub
Private Sub MINUSD_Click()
'IF ADD DAYS IS TRUE THEN THE OTHER CHECKBOX IS ENABLE IS FALSE
If MINUSD = True Then
    PLUSD.Enabled = False
    MMDAYS.Enabled = True
    Else
    PLUSD.Enabled = True
    MMDAYS.Enabled = False
End If
End Sub
Private Sub PLUSD_Click()
'IF ADD DAYS IS TRUE THEN THE OTHER CHECKBOX IS ENABLE IS FALSE
If PLUSD = True Then
    MINUSD.Enabled = False
    MMDAYS.Enabled = True
    Else
    MINUSD.Enabled = True
    MMDAYS.Enabled = False
End If
End Sub

'END OF VALIDATION OF FINAL DATE
Private Sub RAUSENTISMO_Click()
Call OptimizeCode_Begin
' Add data to the sheet AData

'Code to block the register in blank
If Me.BUSCADOR.Value = vbNullString Then
MsgBox "Seleccione un Colaborador"
Me.BUSCADOR.SetFocus
Exit Sub
End If

If EG = True And VC_Selected.Value = vbNullString Then
MsgBox "Debe ingresar el código de la enfermedad"
VC_Selected.SetFocus
Exit Sub
End If

If RAusentismos.BUSCADOR.Value = vbNullString Then
MsgBox "Seleccione un Colaborador de Lista"
BUSCADOR.SetFocus
Exit Sub
End If

If RAusentismos.FECHAA.Value = vbNullString Then
MsgBox "Ingrese Fecha"
FECHAA.SetFocus
Exit Sub
End If

If RAusentismos.CONVENCIONES.Value = vbNullString Then
MsgBox "Seleccione la causa del ausentismo"
CONVENCIONES.SetFocus
Exit Sub
End If

If RAusentismos.FECHAI.Value = vbNullString Then
MsgBox "Ingrese Fecha"
FECHAI.SetFocus
Exit Sub
End If

If RAusentismos.FECHAF.Value = vbNullString Then
MsgBox "Ingrese Fecha"
FECHAF.SetFocus
Exit Sub
End If

If RAusentismos.DESCRIPCION.Value = vbNullString Then
MsgBox "Ingrese la descripción de la solicitud del permiso"
DESCRIPCION.SetFocus
Exit Sub
End If

'========== PROCESS TO ADD ID
Dim StartNum As Long
Dim ID As String
Dim ID_S As String

ID_S = Left(FECHAA.Value, 2) & Left(NOMBRES.Value, 2) & Right(NOMBRES.Value, 1) & Mid(FECHAA.Value, 4, 2) & _
        Right(IDENTIFICACION, 2) & Left(CONVENCIONES.Value, 1) & Left(FECHAI.Value, 2)
ID = ID_S
'=========== END PROCESS TO ADD ID

'Code to calculate the works days and the hours
    Dim DHours As Long
    Dim Hour1 As Long
    Dim Hour2 As Long
    Dim lngHours As Long
    Dim DateD As Date
    Dim HourF As Double
    Dim Minutes As Double
    Dim StoreHourM As Double
    Dim HourWM As Double
    DateD = CDate(FECHAF.Value) - CDate(FECHAI.Value)
    Dim inc As Boolean
    HourF = Hour(DateD)
    
    'Conditional to value on lngDay
    If DateD >= 1.4 And DateD <= 1.6 And HourF >= 8 Then
        lngDay = WorksheetFunction.NetworkDays_Intl(CDate(FECHAI.Value), CDate(FECHAF.Value))
        inc = True
        ElseIf DateD >= 1.4 Then
        lngDay = WorksheetFunction.NetworkDays_Intl(CDate(FECHAI.Value), CDate(FECHAF.Value))
    End If
    Minutes = Minute(DateD) / 60
    HourWM = HourF + Minutes
    If HourF > 8 Then 'Always make the work hours equal to 8
        StoreHourM = HourWM - 8
        HourF = HourF - StoreHourM
    End If
    If Int(DateD) > 1.3 Or inc = True Then 'Validate if Number of Days isn't 0
        If MMDAYS >= 1 And PLUSD = True Then
            lngDay = WorksheetFunction.NetworkDays_Intl(CDate(FECHAI.Value), CDate(FECHAF.Value)) + MMDAYS
        ElseIf MMDAYS >= 1 And MINUSD = True Then
            lngDay = WorksheetFunction.NetworkDays_Intl(CDate(FECHAI.Value), CDate(FECHAF.Value)) - MMDAYS
        End If
    'ElseIf Int(DateD) > 1.3 Or inc = True Then
    HourF = (8 * lngDay) - Minutes
    'if number of days es inferior to 1, then number of days equal to 0
    End If
'================================================================
'=======Validation to avoid error with dates without proper format
Dim msgE As String: msgE = "Fecha con formato o valores erroneos"
Dim FDate1 As Boolean
If Me.FECHAA.TextLength < 10 Then
    MsgBox msgE
    Me.FECHAA.SetFocus
    Exit Sub
End If
'=================================================================
'Code to put the data in the form of ill tracking
If (lndDay >= 3 Or lngDay >= 3) And (CONVENCIONES.Value = "E.G." Or CONVENCIONES.Value = "A.T." Or CONVENCIONES.Value = "L.M.") Then
    MsgBox "No olvide realizar la solicitud de reintegro de la incapacidad"
    ITrack = True
    IllTracking.SEMP.Value = Me.BUSCADOR.Value
    IllTracking.SDATE.Value = CDate(Me.FECHAA.Value)
    IllTracking.SDATEI.Value = CDate(Me.FECHAI.Value)
    IllTracking.SDATEF.Value = CDate(Me.FECHAF.Value)
    IllTracking.SENT.Value = EPSN
    IllTracking.Show
End If
'=================================================================
'REGISTER THE DATA IN THE SHEET
'=================================================================
Dim lastRowA As Long
    Dim duplicados As Boolean
    'Obtener la LastRowA disponible
    lastRowA = Sheets("AData").Cells(Rows.Count, AData.[abs_id].Column).End(xlUp).Row + 1

duplicados = False
For i = 2 To lastRowA
    If CLng(CDate(Sheets("AData").Cells(i, AData.[abs_initial_dated].Column).Value)) = CLng(CDate(Me.FECHAI.Value)) Then
        If CLng(CDate(Sheets("AData").Cells(i, AData.[abs_final_dated].Column).Value)) = CLng(CDate(Me.FECHAF.Value)) Then
            If Sheets("AData").Cells(i, AData.[abs_emp_name].Column).Value = Me.BUSCADOR.Value Then
                MsgBox "Datos duplicados en la fila " & i & " revise las fechas"
                duplicados = True
                Exit Sub
            End If
        End If
    End If
Next i
Dim Ingresado As Boolean
Ingresado = False

If Not duplicados Then

Sheets("AData").Cells(lastRowA, AData.[abs_id].Column).Value = ID
Sheets("AData").Cells(lastRowA, AData.[abs_enterprise].Column).Value = Me.Enterprise.Value
Sheets("AData").Cells(lastRowA, AData.[abs_dated].Column).Value = CDate(RAusentismos.FECHAA.Value)
Sheets("AData").Cells(lastRowA, AData.[abs_emp_name].Column).Value = RAusentismos.NOMBRES.Value
Sheets("AData").Cells(lastRowA, AData.[abs_utility].Column).Value = RAusentismos.NOMBRES.Value & RAusentismos.CONVENCIONES.Value
Sheets("AData").Cells(lastRowA, AData.[abs_emp_id].Column).Value = RAusentismos.IDENTIFICACION.Value
Sheets("AData").Cells(lastRowA, AData.[abs_department].Column).Value = RAusentismos.CDEP.Value
Sheets("AData").Cells(lastRowA, AData.[abs_jobname].Column).Value = RAusentismos.CARGO.Value
Sheets("AData").Cells(lastRowA, AData.[abs_wage].Column).Value = RAusentismos.SBASE.Value
Sheets("AData").Cells(lastRowA, AData.[abs_type_abs].Column).Value = RAusentismos.CONVENCIONES.Value
If Me.VC_Selected.Value = vbNullString Then
    Sheets("AData").Cells(lastRowA, AData.[abs_CIE10].Column).Value = "N"
    Sheets("AData").Cells(lastRowA, AData.[abs_CIE10Des].Column).Value = "N"
Else
    Sheets("AData").Cells(lastRowA, AData.[abs_CIE10].Column).Value = RAusentismos.VC_Selected.Value
    Sheets("AData").Cells(lastRowA, AData.[abs_CIE10Des].Column).Value = AbsDescription
End If
Sheets("AData").Cells(lastRowA, AData.[abs_initial_dated].Column).Value = CDate(RAusentismos.FECHAI.Value)
Sheets("AData").Cells(lastRowA, AData.[abs_final_dated].Column).Value = CDate(RAusentismos.FECHAF.Value)
Sheets("AData").Cells(lastRowA, AData.[abs_days].Column).Value = lngDay
Sheets("AData").Cells(lastRowA, AData.[abs_hours].Column).Value = HourF + Minutes
Sheets("AData").Cells(lastRowA, AData.[abs_cause].Column).Value = RAusentismos.DESCRIPCION.Value
 
    If Me.VC_Selected.Value <> vbNullString And HourF > 24 Then
        'Get salary per hour
        Dim wagePH As Long
        wagePH = (Me.SBASE / 240)
        'Get total value of the ill without the 66.67%.
        MsgBox wagePH
        Dim BValue As Long
        BValue = (wagePH * (HourF + Minutes)) - (wagePH * 16)
        MsgBox BValue
        'Get the value that is pay for the administrator
        Dim NetValue As Long
        NetValue = BValue * (66.67 / 100)
        'Put value in data sheet
        Sheets("AData").Cells(lastRowA, AData.[abs_cost].Column).Value = NetValue
    Else
        Sheets("AData").Cells(lastRowA, AData.[abs_cost].Column).Value = (SBASE.Value / 240) * (HourF + Minutes)
    End If

End If
'Add to reports sheet
Select Case CONVENCIONES
    Case "E.G.": NewsType = 3: Call NewsReportSABS
    Case "L.M.": NewsType = 4: Call NewsReportSABS
    Case "A.T.": NewsType = 5: Call NewsReportSABS
    Case "V.": NewsType = 6: Call NewsReportSABS
    Case "C.D.": NewsType = 7: Call NewsReportSABS
    Case "P.S.T.N.R": NewsType = 8: Call NewsReportSABS
End Select

'NewsType


'CONVENCIONES.Value = "E.G." Or CONVENCIONES.Value = "A.T." Or CONVENCIONES.Value = "L.M."

'=================================================================
'Continue with the code
If ITrack = False Then
    If MsgBox("INGRESAR NUEVO AUSENTISMO", vbYesNo) = vbYes Then
        Call ClearDataF
        BUSCADOR.SetFocus
    Else
        Unload Me
        'Sheets("PPrincipal").Select
    End If
End If
Application.Calculation = xlCalculationAutomatic
Call OptimizeCode_End
End Sub
'========================================
Private Sub UPDATEA_Click()
Call OptimizeCode_Begin
'=================================
'UPDATE INFO OF ABS
'=================================
Dim ws As Worksheet
Set ws = Sheets("AData")
Dim lRow As Long
Dim CurrentRow As Long
lRow = Sheets("AData").Range("A:A").Find(What:=Me.ID_A, after:=Sheets("AData").Range("A2"), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row

CurrentRow = lRow

ws.Cells(CurrentRow, AData.[abs_enterprise].Column).Value = Me.Enterprise.Value
ws.Cells(CurrentRow, AData.[abs_dated].Column).Value = Me.FECHAA
ws.Cells(CurrentRow, AData.[abs_type_abs].Column).Value = Me.CONVENCIONES
If Me.VC_Selected.Value = vbNullString Then
    ws.Cells(CurrentRow, AData.[abs_CIE10].Column).Value = "N"
    ws.Cells(CurrentRow, AData.[abs_CIE10Des].Column).Value = "N"
Else
    ws.Cells(CurrentRow, AData.[abs_CIE10].Column).Value = Me.VC_Selected.Value
    ws.Cells(CurrentRow, AData.[abs_CIE10Des].Column).Value = AbsDescription
End If
ws.Cells(CurrentRow, AData.[abs_initial_dated].Column).Value = CDate(FECHAI)
ws.Cells(CurrentRow, AData.[abs_final_dated].Column).Value = CDate(FECHAF)

    Dim DHours As Long
    Dim Hour1 As Long
    Dim Hour2 As Long
    Dim lngHours As Long
    Dim lngDay As Long
    Dim DateD As Date
    Dim HourF As Double
    Dim Minutes As Double
    Dim StoreHourM As Double
    Dim HourWM As Double
    
    DateD = CDate(FECHAF.Value) - CDate(FECHAI.Value)
    Dim inc As Boolean
    HourF = Hour(DateD)
    
    'Conditional to value on lngDay
    If DateD >= 1.4 And DateD <= 1.6 And HourF >= 8 Then
        lngDay = WorksheetFunction.NetworkDays_Intl(CDate(FECHAI.Value), CDate(FECHAF.Value))
        inc = True
        ElseIf DateD >= 1.4 Then
        lngDay = WorksheetFunction.NetworkDays_Intl(CDate(FECHAI.Value), CDate(FECHAF.Value))
    End If
    
    Minutes = Minute(DateD) / 60
    HourWM = HourF + Minutes
    If HourF > 8 Then 'Always make the work hours equal to 8
        StoreHourM = HourWM - 8
        HourF = HourF - StoreHourM
    End If
    
    If Int(DateD) > 1.3 Or inc = True Then 'Validate if Number of Days isn't 0
        If MMDAYS >= 1 And PLUSD = True Then
            lngDay = WorksheetFunction.NetworkDays_Intl(CDate(FECHAI.Value), CDate(FECHAF.Value)) + MMDAYS
        ElseIf MMDAYS >= 1 And MINUSD = True Then
            lngDay = WorksheetFunction.NetworkDays_Intl(CDate(FECHAI.Value), CDate(FECHAF.Value)) - MMDAYS
        End If
    'ElseIf Int(DateD) > 1.3 Or inc = True Then
    HourF = (8 * lngDay) - Minutes
    'if number of days es inferior to 1, then number of days equal to 0
    End If
    
    ws.Cells(CurrentRow, AData.[abs_days].Column).Value = lngDay
    ws.Cells(CurrentRow, AData.[abs_hours].Column).Value = HourF + Minutes
    ws.Cells(CurrentRow, AData.[abs_cause].Column).Value = RAusentismos.DESCRIPCION.Value
    
    If Me.VC_Selected.Value <> vbNullString And HourF > 24 Then
        'Get salary per hour
        Dim wagePH As Long
        wagePH = (Me.SBASE / 240)
        'Get total value of the ill without the 66.67%.
        MsgBox wagePH
        Dim BValue As Long
        BValue = (wagePH * (HourF + Minutes)) - (wagePH * 16)
        MsgBox BValue
        'Get the value that is pay for the administrator
        Dim NetValue As Long
        NetValue = BValue * (66.67 / 100)
        'Put value in data sheet
        Sheets("AData").Cells(CurrentRow, AData.[abs_cost].Column).Value = NetValue
    Else
        Sheets("AData").Cells(CurrentRow, AData.[abs_cost].Column).Value = (SBASE.Value / 240) * (HourF + Minutes)
    End If
    
    Call ClearDataF
    ID_A.SetFocus
Call OptimizeCode_End
End Sub
'========================================
Private Sub UserForm_Initialize()
Set wsPD = Sheets("PData")
Me.BUSCADOR.SetFocus 'setfocus on this field to avoid err
Call OptimizeCode_Begin
'==DISABLED UPDATE VALUE TO AVOID ERR
Me.RAUSENTISMO.Enabled = False
Me.UPDATEA.Enabled = False
'Disable FECHAF to avoid Err: No coinciden los tipos
If FECHAI.Value = vbNullString Then
    FECHAF.Enabled = False
End If
'When Form is open Disabled Search Code of ill
VC_Selected.Enabled = False
B_FINDC.Enabled = False
'Add data to the ComboBox for select the employee
Dim lastrow As Long
lastrow = Sheets("PData").Cells(Rows.Count, 1).End(xlUp).Row
BUSCADOR.List = Sheets("PData").Range("B2:b" & lastrow).Value
Me.Enterprise.List = Array("IMEXHS", "RIMAB")
'ADD DATA TO THE ID
Dim lastRowR As Long
lastRowR = Sheets("AData").Cells(Rows.Count, 1).End(xlUp).Row
ID_A.List = Sheets("AData").Range("A2:A" & lastRowR).Value
    
    Dim ctrl As MSForms.control
    For Each ctrl In Controls
        If TypeOf ctrl Is MSForms.ListBox Then
        If ctrl.Name Like "CONVENCIONES" Then ctrl.List = Array("E.P.", "E.L.", "A.T.", "E.G.", "C.D.", "P.S.T.", "F.S.P.", "V.", "L.M.", "P.S.T.N.R")
        If ctrl.Name Like "CONVENCIONESD" Then ctrl.List = Array("ENFERMEDAD PROFESIONAL", "ENFERMEDAD LABORAL", "ACCIDENTE DE TRABAJO", "ENFERMEDAD GENERAL", "CALAMIDAD DOMESTICA", "PERMISO SOLICITADO POR EL TRABAJADOR", "FALLA SIN PERMISO", "VACACIONES", "LICENCIA MATERNIDAD", "PER SOL TRA NR")
        End If
    Next ctrl

'INITIALIZE FORM WITH MMDAYS DISABLED
MMDAYS.Enabled = False
Call OptimizeCode_End
End Sub
'===================================
Function ClearDataF()
Call OptimizeCode_Begin
'Delete the data from all the boxes

RAusentismos.BUSCADOR.Value = vbNullString
RAusentismos.FECHAA.Value = vbNullString
RAusentismos.NOMBRES.Value = vbNullString
RAusentismos.IDENTIFICACION.Value = vbNullString
RAusentismos.CDEP.Value = vbNullString
RAusentismos.CARGO.Value = vbNullString
RAusentismos.SBASE.Value = vbNullString
RAusentismos.CONVENCIONES.Value = vbNullString
RAusentismos.FECHAI.Value = vbNullString
RAusentismos.FECHAF.Value = vbNullString
RAusentismos.DESCRIPCION.Value = vbNullString
RAusentismos.MMDAYS.Value = vbNullString
RAusentismos.DABS.Value = vbNullString
RAusentismos.VC_Selected.Value = vbNullString
RAusentismos.OHOUR.Value = False
Call OptimizeCode_End
End Function

'When close then select sheets principal
Private Sub UserForm_QueryClose(CANCEL As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        CANCEL = True
        RAusentismos.Hide
        Sheets("PPrincipal").Select
    End If
End Sub
Private Sub OHOUR_Click()
'==Select field when option is selected
FECHAI.SetFocus
End Sub
Sub NewsReportSABS()

Set ShNewsReport = Sheets("News_P")
LastrowNR = ShNewsReport.Cells(Rows.Count, 1).End(xlUp).Row + 1
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_enterprise].Column).Value = Me.Enterprise.Value
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_no].Column).Value = LastrowNR - 1
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_emp_name].Column).Value = Me.NOMBRES.Value
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_id].Column).Value = Me.IDENTIFICACION.Value
Select Case NewsType
    Case 3: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_illleave].Column).Value = "X"
    Case 4: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_maternitylicense].Column).Value = "X"
    Case 5: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_wkaccident].Column).Value = "X"
    Case 6: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_vacations].Column).Value = "X"
    Case 7: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_calam].Column).Value = "X"
    Case 8: ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_pstnr].Column).Value = "X"
End Select
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_initialdated].Column).Value = CDate(Me.FECHAI.Value)
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_finaldated].Column).Value = CDate(Me.FECHAF.Value)
ShNewsReport.Cells(LastrowNR, ShNewsReport.[nw_observation].Column).Value = Me.DESCRIPCION.Value
End Sub
