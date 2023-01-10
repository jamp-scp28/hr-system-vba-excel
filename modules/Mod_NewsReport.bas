Attribute VB_Name = "Mod_NewsReport"
'=====================
'Variable to Speed up the code
'=====================
Public NewsDateD As Date
Public bCallCode As Boolean
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
Sub News_ReportOptions()
    If MsgBox("¿Desea Exportar el Reporte?", vbYesNo) = vbYes Then
    MsgBoxAnswer = MsgBoxCB("Seleccione el formato", "EXCEL")
        If MsgBoxAnswer = 1 Then
            ReportsI.Hide
            Call exportPivotS
        End If
    Else
        Exit Sub
    End If
End Sub
Sub Export_NewsReport()
Call OptimizeCode_Begin 'Optimize Code
Dim wsEA As Worksheet
Set wsEA = Sheets("RSData")
'Export PivotTable as PDF
Dim fName As String, FPath As String
With wsEA
Dim lr As Long
    fName = wsEA.Range("D1").Value
    lr = wsEA.Range("B" & Rows.Count).End(xlUp).Row + 6
End With
    wsEA.PageSetup.PrintArea = "B1:G" & lr
    ChDir ActiveWorkbook.Path
    wsEA.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        fName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
ReportsI.NewsReport = False
MsgBox "El documento de guardo se: " & ActiveWorkbook.Path
Call OptimizeCode_End 'Optimize Code
End Sub
Sub exportPivotS()
Call OptimizeCode_Begin 'Optimize Code
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("SData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    With ActiveSheet
        .Name = "SData": .Visible = True
    End With
Worksheets.Add().Name = "RSData"
'Declare Variables

Dim PvtTbl As PivotTable, ws As Worksheet
Set ws = Sheets("RSData")
Dim wb As Workbook
Set wb = ActiveWorkbook
Dim wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet, SLcache As SlicerCache, SL As Slicer
'Set the value of variables
Set wsData = Worksheets("SData")
Set wsPvtTbl = Worksheets("RSData")
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
Set nuevo = Worksheets("RSData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData, Version:=xlPivotTableVersion15)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)
'Asign the position to every field
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField
    .Position = 1
End With
With PvtTbl.PivotFields("IDENTIFICACION")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("FECHA DE NOVEDAD")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("TIPO DE NOVEDAD")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("DATO ANTERIOR")
    .Orientation = xlRowField: .Position = 5
End With
With PvtTbl.PivotFields("DATO NUEVO")
    .Orientation = xlRowField: .Position = 6
End With
PvtTbl.RowAxisLayout xlTabularRow
With PvtTbl
    For Each PvtFld In .PivotFields
        PvtFld.Subtotals(1) = False
        PvtFld.Subtotals(1) = False
    Next PvtFld
End With
Dim WSS As Worksheet
Set WSS = Sheets("RSData")
'wrap text
Dim lastRowD As Long
lastRowD = WSS.Cells(Rows.Count, 5).End(xlUp).Row
With WSS.Range("E1:E" & lastRowD)
    .WrapText = True
End With
'Desactivate the calculation of row and column grand
With PvtTbl
    .ColumnGrand = False: .RowGrand = False
End With
Dim lastRowA As Long
lastRowA = Sheets("RSData").Cells(Rows.Count, 2).End(xlUp).Row
With Sheets("RSData").Range("B6:G" & lastRowA)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
WSS.Range("B:B").ColumnWidth = 20
WSS.Range("C:C").ColumnWidth = 26.57
WSS.Range("D:D").ColumnWidth = 17.71
WSS.Range("E:E").ColumnWidth = 19.43
WSS.Range("F:F").ColumnWidth = 17.43
WSS.Range("G:G").ColumnWidth = 17.43
'Assign the style to the pivotable
PvtTbl.TableStyle2 = "PivotStyleMedium9"
ReportsI.NewsReport = False

'deselected report
shtbron.Copy wbdoel.Sheets("RSData").Range("A1")
wbdoel.Sheets("RSData").Range("D1").Value = "REPORTE NOVEDADES - IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
wbdoel.Sheets("SData").Visible = False
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True
Call OptimizeCode_End 'Optimize Code
bCallCode = False
End Sub
