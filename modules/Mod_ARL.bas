Attribute VB_Name = "Mod_ARL"
'=====================
'Variable to Speed up the code
'=====================
Public bCallCode As Boolean
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean
Global Active_Emp As Boolean
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
Sub ARL_ReportOptions()
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
'=====================
'End Variable to Speed up the code
'=====================
Sub ARL_Report()
Call OptimizeCode_Begin 'Optimize Code
'Declare Variables
Dim PvtTbl As PivotTable, wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet, PvtFld As PivotField
'Set the value of variables
Set wsData = Worksheets("PData")
Set wsPvtTbl = Worksheets("RARLData")
lastrow = wsData.Cells(Rows.Count, 2).End(xlUp).Row
lastColumn = wsData.Cells(1, 41).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RARLData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData, Version:=xlPivotTableVersion15)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)
'Asign the position to every field
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("IDENTIFICACION")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("FECHA DE INGRESO")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("DEPARTAMENTO")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("CARGO")
    .Orientation = xlRowField: .Position = 5
End With
With PvtTbl.PivotFields("TIPO DE CONTRATO")
    .Orientation = xlRowField: .Position = 6
End With
With PvtTbl.PivotFields("EPS")
    .Orientation = xlRowField: .Position = 7
End With
With PvtTbl.PivotFields("AFP")
    .Orientation = xlRowField
    .Position = 8
End With
With PvtTbl.PivotFields("CCF")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("ARL")
    .Orientation = xlRowField: .Position = 10
End With
With PvtTbl.PivotFields("CENTRO DE TRABAJO")
    .Orientation = xlRowField: .Position = 11
End With
With PvtTbl.PivotFields("CLASE")
    .Orientation = xlRowField: .Position = 12
End With
With PvtTbl.PivotFields("TASA")
    .Orientation = xlRowField: .Position = 13
End With
With PvtTbl.PivotFields("FECHA DE COBERTURA")
    .Orientation = xlRowField: .Position = 14
End With
With PvtTbl.PivotFields("RETIRADO")
    .Orientation = xlRowField: .Position = 15: .PivotItems("true").Visible = False 'Filter the data, no selecting TRUE values
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
    .ColumnGrand = False: .RowGrand = False
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
PvtTbl.TableStyle2 = "ReportStyle"
Sheets("RARLData").Columns("P:P").EntireColumn.Hidden = True
With Sheets("RARLData").Range("B6:O" & lastrow)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
'Set space to rows
Dim u As Long
With Sheets("RARLData")
    For u = 7 To lastrow
        .Cells(u, 1).RowHeight = 38
    Next u
End With
'add width to row six
Dim wsAC As Worksheet
Set wsAC = Sheets("RARLData")
wsAC.Cells(7, 2).RowHeight = 27.5
'define width to colums
wsAC.Range("B:B").ColumnWidth = 24
wsAC.Range("C:C").ColumnWidth = 16.43
wsAC.Range("D:D").ColumnWidth = 14.71
wsAC.Range("E:E").ColumnWidth = 21.29
wsAC.Range("F:F").ColumnWidth = 22.71
wsAC.Range("G:G").ColumnWidth = 14
wsAC.Range("H:H").ColumnWidth = 14
wsAC.Range("I:I").ColumnWidth = 16.71
wsAC.Range("J:J").ColumnWidth = 15.71
wsAC.Range("K:K").ColumnWidth = 11.71
wsAC.Range("L:L").ColumnWidth = 17.57
wsAC.Range("M:M").ColumnWidth = 7.43
wsAC.Range("N:N").ColumnWidth = 7.73
wsAC.Range("O:O").ColumnWidth = 12.57
ReportsI.ARL = False

Call OptimizeCode_End 'Optimize Code
End Sub
Sub Export_ARLReport()
Call OptimizeCode_Begin 'Optimize Code
Dim wsEA As Worksheet
Set wsEA = Sheets("RARLData")
'Export PivotTable as PDF
Dim fName As String, FPath As String
With wsEA
    fName = wsEA.Range("D1").Value
    lr = wsEA.Range("B" & Rows.Count).End(xlUp).Row
End With
    wsEA.PageSetup.PrintArea = "B1:O" & lr
    ChDir ActiveWorkbook.Path
    wsEA.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        fName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
MsgBox "El documento de guardo se: " & ActiveWorkbook.Path
Call OptimizeCode_End 'Optimize Code
End Sub
Sub exportPivot()
Call OptimizeCode_Begin 'Optimize Code
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
    With ActiveSheet
        .Name = "PData": .Visible = True
    End With
Worksheets.Add().Name = "RARLData"
'Declare Variables

'Declare Variables
Dim PvtTbl As PivotTable, wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet
'Set the value of variables
Set wsData = Worksheets("PData")
Set wsPvtTbl = Worksheets("RARLData")
lastrow = wsData.Cells(Rows.Count, 2).End(xlUp).Row
lastColumn = wsData.Cells(1, 41).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RARLData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData, Version:=xlPivotTableVersion15)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)
'Asign the position to every field
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("IDENTIFICACION")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("FECHA DE INGRESO")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("DEPARTAMENTO")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("CARGO")
    .Orientation = xlRowField: .Position = 5
End With
With PvtTbl.PivotFields("TIPO DE CONTRATO")
    .Orientation = xlRowField: .Position = 6
End With
With PvtTbl.PivotFields("EPS")
    .Orientation = xlRowField: .Position = 7
End With
With PvtTbl.PivotFields("AFP")
    .Orientation = xlRowField
    .Position = 8
End With
With PvtTbl.PivotFields("CCF")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("ARL")
    .Orientation = xlRowField: .Position = 10
End With
With PvtTbl.PivotFields("CENTRO DE TRABAJO")
    .Orientation = xlRowField: .Position = 11
End With
With PvtTbl.PivotFields("CLASE")
    .Orientation = xlRowField: .Position = 12
End With
With PvtTbl.PivotFields("TASA")
    .Orientation = xlRowField: .Position = 13
End With
With PvtTbl.PivotFields("FECHA DE COBERTURA")
    .Orientation = xlRowField: .Position = 14
End With
With PvtTbl.PivotFields("RETIRADO")
    .Orientation = xlRowField: .Position = 15: .PivotItems("true").Visible = False 'Filter the data, no selecting TRUE values
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
    .ColumnGrand = False: .RowGrand = False
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
PvtTbl.TableStyle2 = "ReportStyle"
Sheets("RARLData").Columns("P:P").EntireColumn.Hidden = True
With Sheets("RARLData").Range("B6:O" & lastrow)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
'Set space to rows
Dim u As Long
With Sheets("RARLData")
    For u = 7 To lastrow
        .Cells(u, 1).RowHeight = 38
    Next u
End With
'add width to row six
Dim wsAC As Worksheet
Set wsAC = Sheets("RARLData")
wsAC.Cells(7, 2).RowHeight = 27.5
'define width to colums
wsAC.Range("B:B").ColumnWidth = 24
wsAC.Range("C:C").ColumnWidth = 16.43
wsAC.Range("D:D").ColumnWidth = 14.71
wsAC.Range("E:E").ColumnWidth = 21.29
wsAC.Range("F:F").ColumnWidth = 22.71
wsAC.Range("G:G").ColumnWidth = 14
wsAC.Range("H:H").ColumnWidth = 14
wsAC.Range("I:I").ColumnWidth = 16.71
wsAC.Range("J:J").ColumnWidth = 15.71
wsAC.Range("K:K").ColumnWidth = 11.71
wsAC.Range("L:L").ColumnWidth = 17.57
wsAC.Range("M:M").ColumnWidth = 7.43
wsAC.Range("N:N").ColumnWidth = 7.73
wsAC.Range("O:O").ColumnWidth = 12.57
ReportsI.ARL = False

shtbron.Copy wbdoel.Sheets("RARLData").Range("A1")
wbdoel.Sheets("RPDataT").Range("D1").Value = "REPORTE ARL - IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
wbdoel.Sheets("PData").Visible = False
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True

bCallCode = False
Call OptimizeCode_End 'Optimize Code
End Sub
Sub exportPivotT()
Call OptimizeCode_Begin 'Optimize Code
'==========================================
'EXPORT IN FORMAT XLM TO OTHERS DEPAR
'===========================================
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("PData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    With ActiveSheet
        .Name = "PData": .Visible = True
    End With
Worksheets.Add().Name = "RARLData"
Sheets("PData").Columns("C:J").Delete Shift:=xlToLeft
Sheets("PData").Columns("O:Q").Delete Shift:=xlToLeft
'Declare Variables
Dim PvtTbl As PivotTable, wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet
'Set the value of variables
Set wsData = Worksheets("PData")
Set wsPvtTbl = Worksheets("RARLData")
lastrow = wsData.Cells(Rows.Count, 2).End(xlUp).Row
lastColumn = wsData.Cells(1, 30).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RARLData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData, Version:=xlPivotTableVersion15)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)
'Asign the position to every field
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("IDENTIFICACION")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("FECHA DE INGRESO")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("DEPARTAMENTO")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("CARGO")
    .Orientation = xlRowField: .Position = 5
End With
With PvtTbl.PivotFields("TIPO DE CONTRATO")
    .Orientation = xlRowField: .Position = 6
End With
With PvtTbl.PivotFields("EPS")
    .Orientation = xlRowField: .Position = 7
End With
With PvtTbl.PivotFields("AFP")
    .Orientation = xlRowField: .Position = 8
End With
With PvtTbl.PivotFields("CCF")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("ARL")
    .Orientation = xlRowField: .Position = 10
End With
With PvtTbl.PivotFields("CENTRO DE TRABAJO")
    .Orientation = xlRowField: .Position = 11
End With
With PvtTbl.PivotFields("CLASE")
    .Orientation = xlRowField: .Position = 12
End With
With PvtTbl.PivotFields("TASA")
    .Orientation = xlRowField: .Position = 13
End With
With PvtTbl.PivotFields("FECHA DE COBERTURA")
    .Orientation = xlRowField: .Position = 14
End With
With PvtTbl.PivotFields("RETIRADO")
    .Orientation = xlRowField: .Position = 15: .PivotItems("True").Visible = False 'Filter the data, no selecting TRUE values
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
    .ColumnGrand = False: .RowGrand = False
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
Sheets("RARLData").Columns("P:P").EntireColumn.Hidden = True
With Sheets("RARLData").Range("B6:O" & lastrow)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
'Set space to rows
Dim u As Long
With Sheets("RARLData")
    For u = 7 To lastrow
        .Cells(u, 1).RowHeight = 38
    Next u
End With
shtbron.Copy wbdoel.Sheets("RARLData").Range("A1") 'copy headers
wbdoel.Sheets("RARLData").Range("D1").Value = "REPORTE ARL - IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
ActiveWorkbook.Sheets("PData").Visible = False
'add width to row six
Dim wsAC As Worksheet
Set wsAC = Sheets("RARLData")
wsAC.Cells(7, 2).RowHeight = 27.5
'define width to colums
wsAC.Range("B:B").ColumnWidth = 24
wsAC.Range("C:C").ColumnWidth = 16.43
wsAC.Range("D:D").ColumnWidth = 14.71
wsAC.Range("E:E").ColumnWidth = 21.29
wsAC.Range("F:F").ColumnWidth = 22.71
wsAC.Range("G:G").ColumnWidth = 14
wsAC.Range("H:H").ColumnWidth = 14
wsAC.Range("I:I").ColumnWidth = 16.71
wsAC.Range("J:J").ColumnWidth = 15.71
wsAC.Range("K:K").ColumnWidth = 11.71
wsAC.Range("L:L").ColumnWidth = 17.57
wsAC.Range("M:M").ColumnWidth = 7.43
wsAC.Range("N:N").ColumnWidth = 7.73
wsAC.Range("O:O").ColumnWidth = 12.57
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True
Call OptimizeCode_End 'Optimize Code
End Sub

