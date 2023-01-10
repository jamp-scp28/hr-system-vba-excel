Attribute VB_Name = "Mod_Active_Report"
Public lastrow As Long 'Get lastrow from the sheet with the employee data
Public lastColumn As Long 'Get lastcolumn from the sheet with the employee data
Public bCallCode As Boolean 'Variable to avoid msg when creating extern xls report
'=====================
'Preset to speed up the code
'=====================
Global Active_Emp As Boolean
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean
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
Sub SelectExportOption()
If MsgBox("¿Desea Exportar el Reporte?", vbYesNo) = vbYes Then
MsgBoxAnswer = MsgBoxCB("Seleccione el formato", "EXCEL")
    If MsgBoxAnswer = 1 Then
        ReportsI.Hide
        If MsgBox("¿El reporte es para personal ajeno al departamento?", vbYesNo) = vbYes Then
            Call exportPivotT
            Else
            Call exportPivot
        End If
    End If
Else
Exit Sub
End If
End Sub
Sub exportPivot()
Call OptimizeCode_Begin 'optimize code
'==========================================
'EXPORT IN FORMAT XLM
'===========================================
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("PData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    Dim r As Integer
    For r = wbdoel.Sheets("PData").UsedRange.Rows.Count To 1 Step -1
        If Cells(r, shtbron2.[RETIRED].Column) = True Then
            wbdoel.Sheets("PData").Rows(r).EntireRow.Delete
        End If
    Next
Worksheets.Add().Name = "RPData"
'Declare Variables

'DECLARE VARIABLES
Dim PvtTbl      As PivotTable
Dim wsData      As Worksheet 'THIS SHEET HAS THE DATA OF THE EMPLOYEES
Dim rngData     As Range 'GET THE DATA FROM THE USEDRANGE
Dim PvtTblCache As PivotCache
Dim wsPvtTbl    As Worksheet 'THIS SHEET IS WHERE THE REPORT IS GOING TO BE CREATE

'SET VALUE FOR VARIABLES
Set wsData = Worksheets("PData")
Set wsPvtTbl = Worksheets("RPData")
lastrow = wsData.Cells(Rows.Count, 2).End(xlUp).Row 'GET LASTROW
lastColumn = wsData.Cells(1, Columns.Count).End(xlToLeft).Column 'GET LASTCOLUMN
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)

'DELETE THE PIVOTTABLE IF IT ALREADY EXIST
For Each PvtTbl In wsPvtTbl.PivotTables
    If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
        PvtTbl.TableRange2.Clear
    End If
Next PvtTbl
'CREATE NEW PIVOTTABLE
Set nuevo = Worksheets("RPData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=("PData!R1C1:R" & lastrow & "C" & lastColumn), Version:=xlPivotTableVersion11)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion11)
'SELECT FIELDS TO DISPLAY AND ASSIGN THEM A POSITION
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("IDENTIFICACION")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("E-MAIL CORPORATIVO")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("TELEFONO MOVIL CORPORATIVO")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("TELEFONO OFICINA - EXT")
    .Orientation = xlRowField: .Position = 5
End With
With PvtTbl.PivotFields("FECHA DE INGRESO")
    .Orientation = xlRowField: .Position = 6
End With
With PvtTbl.PivotFields("CODIGO DEPARTAMENTO")
    .Orientation = xlRowField: .Position = 7
End With
With PvtTbl.PivotFields("DEPARTAMENTO")
    .Orientation = xlRowField: .Position = 8
End With
With PvtTbl.PivotFields("CARGO")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("TIPO DE CONTRATO")
    .Orientation = xlRowField: .Position = 10
End With
'*******************************
'When the code
'*******************************
If bCallCode = True Then
Else
With PvtTbl.PivotFields("RETIRADO")
    .Orientation = xlRowField: .Position = 9
End With
End If
'*******************************
'END OF THIS VALIDATION
'*******************************
With PvtTbl
    For Each PvtFld In .PivotFields
        PvtFld.Subtotals(1) = False
        PvtFld.Subtotals(1) = False
    Next PvtFld
End With
'Create calculate fields for pivottable style
 PvtTbl.AddDataField PvtTbl _
        .PivotFields("SALARIO"), "SALARIOS", xlSum
    With PvtTbl.PivotFields("SALARIOS")
        .Function = xlSum: .NumberFormat = "_($ * #,##0_);_($ * (#,##0);_($ * ""-""_);_(@_)"
    End With
  
  PvtTbl.AddDataField PvtTbl _
        .PivotFields("RODAMIENTO"), "Cuenta de RODAMIENTO", xlCount
    With PvtTbl.DataPivotField
        .Orientation = xlColumnField: .Position = 1
    End With
    With PvtTbl.PivotFields("Cuenta de RODAMIENTO")
        .Caption = "Suma de RODAMIENTO": .Function = xlSum: .NumberFormat = "_($ * #,##0_);_($ * (#,##0);_($ * ""-""_);_(@_)"
    End With
    PvtTbl.AddDataField PvtTbl _
        .PivotFields("O AUXILIOS"), "Suma de O AUXILIOS", xlSum
    With PvtTbl.PivotFields("suma de O AUXILIOS")
        .NumberFormat = "_($ * #,##0_);_($ * (#,##0);_($ * ""-""_);_(@_)"
    End With
'Desactivate the calculation of row and column grand
With PvtTbl
    .ColumnGrand = True: .RowGrand = False
End With
'PivotTable.ShowTableStyleColumnHeaders Property. Set to True to display the column headers in the PivotTable.
PvtTbl.ShowTableStyleColumnHeaders = True
'PivotTable.ShowTableStyleRowHeaders Property. Set to True to display the row headers in the PivotTable.
PvtTbl.ShowTableStyleRowHeaders = False
'PivotTable.ShowTableStyleColumnStripes Property. Set to True to display the banded columns in the PivotTable.
PvtTbl.ShowTableStyleColumnStripes = True
'PivotTable.ShowTableStyleRowStripes Property. Set to True to display the banded rows in the PivotTable.
PvtTbl.ShowTableStyleRowStripes = True
'Assign the style to the pivotable
PvtTbl.TableStyle2 = "PivotStyleMedium9"

With Sheets("RPData").Range("B6:M" & lastrow)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
'Set space to rows
Dim u As Long
Dim LastRowRPD As Long
LastRowRPD = Sheets("RPData").Cells(Rows.Count, 2).End(xlUp).Row
With Sheets("RPData")
    For u = 7 To LastRowRPD
        .Cells(u, 1).RowHeight = 38
    Next u
End With
'add width to row six
Dim wsAC As Worksheet
Set wsAC = Sheets("RPData")
wsAC.Cells(7, 2).RowHeight = 27.5
'define width to colums
wsAC.Range("B:B").ColumnWidth = 26.57
wsAC.Range("C:C").ColumnWidth = 16.57
wsAC.Range("D:D").ColumnWidth = 29.43
wsAC.Range("E:E").ColumnWidth = 16
wsAC.Range("F:F").ColumnWidth = 19.29
wsAC.Range("G:G").ColumnWidth = 14
wsAC.Range("H:H").ColumnWidth = 32.57
wsAC.Range("I:I").ColumnWidth = 17.71
wsAC.Range("K:K").ColumnWidth = 13.14
wsAC.Range("L:L").ColumnWidth = 12.86
wsAC.Range("M:M").ColumnWidth = 11.86
ReportsI.Active_Emp = False

shtbron.Copy wbdoel.Sheets("RPData").Range("A1") 'COPY HEADERS
'==Add name to the report
wbdoel.Sheets("RPData").Range("D1").Value = "REPORTE PERSONAL ACTIVO DE IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
wbdoel.Sheets("PData").Visible = False
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True
Call OptimizeCode_End 'optimize code


End Sub

Sub exportPivotT()
Call OptimizeCode_Begin 'optimize code
'====================================
' REPORT TO PEOPLE THAT DOES NOT BELONG TO THE HUMAN RESOURCES
'===========================================================
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("PData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    'Delete the false values
    Dim r As Integer
    For r = wbdoel.Sheets("PData").UsedRange.Rows.Count To 1 Step -1
        If Cells(r, "AM") = True Then
            wbdoel.Sheets("PData").Rows(r).EntireRow.Delete
        End If
    Next
Worksheets.Add().Name = "RPData"
'delete columns with private information
Sheets("PData").Columns(PData.[wage].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[Auxi1].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[Auxi2].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[CIVILSTATUS].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[DEGREE].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[DATEDOB].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[EAGE].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[EADDRESS].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[NHOOD].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[DISTRICT].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[EPHONEM].Column).Delete Shift:=xlToLeft
Sheets("PData").Columns(PData.[EPHONES].Column).Delete Shift:=xlToLeft
'Declare Variables
Dim PvtTbl As PivotTable
Dim wsData As Worksheet
Dim rngData As Range
Dim PvtTblCache As PivotCache
Dim wsPvtTbl As Worksheet
Dim lastrow As Long
Dim lastRowR As Long
'Set the value of variables
Set wsData = Worksheets("PData")
Set wsPvtTbl = Worksheets("RPData")
lastrow = wsData.Cells(Rows.Count, 2).End(xlUp).Row
lastColumn = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RPData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=("PData!R1C1:R" & lastrow & "C" & lastColumn), Version:=xlPivotTableVersion15)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)
'Asign the position to every field
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("IDENTIFICACION")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("E-MAIL CORPORATIVO")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("TELEFONO MOVIL CORPORATIVO")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("TELEFONO OFICINA - EXT")
    .Orientation = xlRowField: .Position = 5
End With
With PvtTbl.PivotFields("FECHA DE INGRESO")
    .Orientation = xlRowField: .Position = 6
End With
With PvtTbl.PivotFields("CARGO")
    .Orientation = xlRowField: .Position = 7
End With
With PvtTbl.PivotFields("TIPO DE CONTRATO")
    .Orientation = xlRowField: .Position = 8
End With

PvtTbl.RowAxisLayout xlTabularRow
With PvtTbl
    For Each PvtFld In .PivotFields
        PvtFld.Subtotals(1) = False
        PvtFld.Subtotals(1) = False
    Next PvtFld
End With
'Desactivate the calculation of row and column grand
With PvtTbl
    .ColumnGrand = True: .RowGrand = False
End With
'PivotTable.ShowTableStyleColumnHeaders Property. Set to True to display the column headers in the PivotTable.
PvtTbl.ShowTableStyleColumnHeaders = True
'PivotTable.ShowTableStyleRowHeaders Property. Set to True to display the row headers in the PivotTable.
PvtTbl.ShowTableStyleRowHeaders = False
'PivotTable.ShowTableStyleColumnStripes Property. Set to True to display the banded columns in the PivotTable.
PvtTbl.ShowTableStyleColumnStripes = True
'PivotTable.ShowTableStyleRowStripes Property. Set to True to display the banded rows in the PivotTable.
PvtTbl.ShowTableStyleRowStripes = True
'Assign the style to the pivotable
PvtTbl.TableStyle2 = "PivotStyleMedium9"
lastRowR = wbdoel.Sheets("RPData").Cells(Rows.Count, 2).End(xlUp).Row
With Sheets("RPData").Range("B6:M" & lastRowR)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
'Set space to rows
Dim u As Long
With Sheets("RPData")
    For u = 7 To lastRowR
        .Cells(u, 1).RowHeight = 38
    Next u
End With
'add width to row six
Dim wsAC As Worksheet
Set wsAC = Sheets("RPData")
wsAC.Cells(7, 2).RowHeight = 27.5
'define width to colums
wsAC.Range("B:B").ColumnWidth = 26.57
wsAC.Range("C:C").ColumnWidth = 16.57
wsAC.Range("D:D").ColumnWidth = 29.43
wsAC.Range("E:E").ColumnWidth = 16
wsAC.Range("F:F").ColumnWidth = 19.29
wsAC.Range("G:G").ColumnWidth = 14
wsAC.Range("H:H").ColumnWidth = 32.57
wsAC.Range("I:I").ColumnWidth = 17.71
wsAC.Range("K:K").ColumnWidth = 13.14
wsAC.Range("L:L").ColumnWidth = 12.86
wsAC.Range("M:M").ColumnWidth = 11.86
shtbron.Copy wbdoel.Sheets("RPData").Range("A1")
wbdoel.Sheets("RPData").Range("D1").Value = "REPORTE PERSONAL ACTIVO DE IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
wbdoel.Sheets("PData").Visible = False
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True
Call OptimizeCode_End 'optimize code
End Sub

