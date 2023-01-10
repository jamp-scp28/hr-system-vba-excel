Attribute VB_Name = "Mod_Documentation_Report"
'=====================
'Variable to Speed up the code
'=====================
Public bCallCode As Boolean
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean
Global All_Emp As Boolean
Global SL As Slicer
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
Sub Documentation_ReportOption()
    If MsgBox("¿Desea Exportar el Reporte?", vbYesNo) = vbYes Then
    MsgBoxAnswer = MsgBoxCB("Seleccione el formato", "EXCEL")
        If MsgBoxAnswer = 1 Then
            ReportsI.Hide
            Call exportPivot
        End If
    Else
    Exit Sub
    End If
End Sub
'=====================
'End Variable to Speed up the code
'=====================
Sub Documentation_Report()
Call OptimizeCode_Begin 'Optimize Code

End Sub
Sub create_slicer()
Dim i As SlicerCaches, j As Slicers, k As Slicer
Set i = ActiveWorkbook.SlicerCaches
Set j = i.Add(ActiveSheet.PivotTables(1), "APELLIDOS Y NOMBRES", "My_Region").Slicers
Set k = j.Add(ActiveSheet, , "My_Region", "APELLIDOS Y NOMBRES", 0, 0, 200, 200)
MsgBox "Created Slicer"
End Sub
Sub Export_DocumentationReport()
Call OptimizeCode_Begin 'Optimize Code
Dim wsEA As Worksheet
Set wsEA = Sheets("RDData")
'Export PivotTable as PDF
Dim fName As String, FPath As String
With wsEA
    fName = wsEA.Range("D1").Value
    lr = wsEA.Range("B" & Rows.Count).End(xlUp).Row
End With
    wsEA.PageSetup.PrintArea = "B1:E" & lr
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
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("DData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
Worksheets.Add().Name = "RDData"
'Declare Variables
Dim PvtTbl As PivotTable, wb As Workbook
Set wb = ActiveWorkbook

'Declare Variables
Dim ws As Worksheet
Set ws = Sheets("RDData")
Set wb = ActiveWorkbook
Dim wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet, SLcache As SlicerCache, SL As Slicer
'Set the value of variables
Set wsData = Worksheets("DData")
Set wsPvtTbl = Worksheets("RDData")
lastrow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = wsData.Cells(1, 6).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RDData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData, Version:=xlPivotTableVersion15)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)
'Asign the position to every field
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("DOCUMENTOS")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("ESTADO")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("OBSERVACION DEL DOCUMENTO")
    .Orientation = xlRowField: .Position = 4
End With
PvtTbl.RowAxisLayout xlTabularRow
With PvtTbl
    For Each PvtFld In .PivotFields
        PvtFld.Subtotals(1) = False
        PvtFld.Subtotals(1) = False
    Next PvtFld
End With
Dim wsD As Worksheet
Set wsD = Sheets("RDData")
'set the width of the column E to display slicer
wsD.Range("E:E").ColumnWidth = 75.71
'wrap text
Dim lastRowD As Long
lastRowD = wsD.Cells(Rows.Count, 5).End(xlUp).Row
With wsD.Range("E1:E" & lastRowD)
    .WrapText = True
End With
'Desactivate the calculation of row and column grand
With PvtTbl
    .ColumnGrand = False: .RowGrand = False
End With
Dim shp As Shape
For Each shp In Sheets("RDData").Shapes
    If shp.Type = msoSlicer Then shp.Delete
Next shp
'Assign the style to the pivotable
PvtTbl.TableStyle2 = "PivotStyleMedium9"
'CREATE SLICER CACHE
Set SLcache = wb.SlicerCaches.Add2(PvtTbl, "APELLIDOS Y NOMBRES", _
"EMPLOYEES")
'CREATE SLICER
Set SL = SLcache.Slicers.Add( _
Sheets("RDData"), , "EMPLOYEES", "APELLIDOS Y NOMBRES", 85, 1250, 200, 350)
ReportsI.Doc_Report = False

Dim wsAC As Worksheet
Set wsAC = Sheets("RDData")
'Add style
'define width to colums
wsAC.Range("B:B").ColumnWidth = 41.86
wsAC.Range("C:C").ColumnWidth = 33.71
wsAC.Range("D:D").ColumnWidth = 25.71
wsAC.Range("E:E").ColumnWidth = 75.43
ReportsI.Doc_Report = False

'Assign the style to the pivotable
shtbron.Copy wbdoel.Sheets("RDData").Range("A1")
wbdoel.Sheets("RDData").Range("D1").Value = "REPORTE DOCUMENTACION - IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
wbdoel.Sheets("DData").Visible = False
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True
Call OptimizeCode_End 'Optimize Code
bCallCode = False
End Sub
