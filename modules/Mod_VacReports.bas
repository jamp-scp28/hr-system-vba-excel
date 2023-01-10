Attribute VB_Name = "Mod_VacReports"
'=====================
'Variable to Speed up the code
'=====================
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
Sub Vac_ReportOptions()
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
Sub Export_VacReport()
Call OptimizeCode_Begin 'Optimize Code
Dim wsEA As Worksheet
Set wsEA = Sheets("RVData")
'Export PivotTable as PDF
Dim fName As String, FPath As String
With wsEA
    fName = wsEA.Range("D1").Value
    lr = wsEA.Range("B" & Rows.Count).End(xlUp).Row
End With
    wsEA.PageSetup.PrintArea = "B1:N" & lr
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
Set shtbron2 = wbbron.Sheets("VData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    With ActiveSheet
        .Name = "VData"
    End With
Worksheets.Add().Name = "RVData"
'Declare Variables

Dim PvtTbl As PivotTable, wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet
'Set the value of variables
Set wsData = Worksheets("VData")
Set wsPvtTbl = Worksheets("RVData")
lastrow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = wsData.Cells(1, 15).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RVData").Cells(6, 2)
Set PvtTbl = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rngData, Version:=xlPivotTableVersion15).CreatePivotTable( _
        TableDestination:=nuevo, TableName:="PivotTable", _
        DefaultVersion:=xlPivotTableVersion15)
PvtTbl.RowAxisLayout xlTabularRow
'Asign the position to every field
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("CODIGO DEPARTAMENTO")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("CARGO")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("FECHA CONTRATO INDEFINIDO")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("SALARIO BASE")
    .Orientation = xlRowField: .Position = 5: .NumberFormat = "_($ * #,##0_);_($ * (#,##0);_($ * ""-""_);_(@_)"
End With
With PvtTbl.PivotFields("FECHA DE LIQUIDACION")
    .Orientation = xlRowField: .Position = 6
End With
With PvtTbl.PivotFields("DIAS TRABAJADOS")
    .Orientation = xlRowField: .Position = 7
End With
With PvtTbl.PivotFields("VACACIONES QUE CORRESPONDE")
    .Orientation = xlRowField: .Position = 8
End With
With PvtTbl.PivotFields("DIAS VACACIONES DISFRUTADAS")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("DIAS VACACIONES PENDIENTES")
    .Orientation = xlRowField: .Position = 10
End With
With PvtTbl.PivotFields("APLICA POR ESTADO")
    .Orientation = xlRowField: .Position = 11: .PivotItems("TRUE").Visible = False 'Filter the data, no selecting TRUE values
End With
With PvtTbl.PivotFields("APLICA POR CONTRATO")
    .Orientation = xlRowField: .Position = 12: .PivotItems("FALSE").Visible = False 'Filter the data, no selecting TRUE values
End With
With PvtTbl
    For Each PvtFld In .PivotFields
        PvtFld.Subtotals(1) = False
        PvtFld.Subtotals(1) = False
    Next PvtFld
End With
'Create calculate fields for pivottable style
 PvtTbl.AddDataField PvtTbl _
        .PivotFields("VALOR DE VACACIONES"), "VALOR", xlSum
    With PvtTbl.PivotFields("VALOR")
        .Caption = "VALOR": .Function = xlSum: .NumberFormat = "_($ * #,##0_);_($ * (#,##0);_($ * ""-""_);_(@_)"
    End With
'Desactivate the calculation of row and column grand
With PvtTbl
    .ColumnGrand = True: .RowGrand = False
End With
'Assign the style to the pivotable
PvtTbl.TableStyle2 = "PivotStyleMedium9"
'PivotTable.ShowTableStyleColumnHeaders Property. Set to True to display the column headers in the PivotTable.
PvtTbl.ShowTableStyleColumnHeaders = True
'PivotTable.ShowTableStyleRowHeaders Property. Set to True to display the row headers in the PivotTable.
PvtTbl.ShowTableStyleRowHeaders = False
'PivotTable.ShowTableStyleColumnStripes Property. Set to True to display the banded columns in the PivotTable.
PvtTbl.ShowTableStyleColumnStripes = True
'PivotTable.ShowTableStyleRowStripes Property. Set to True to display the banded rows in the PivotTable.
PvtTbl.ShowTableStyleRowStripes = True
Dim lastrowV As Long, wsV As Worksheet
Set wsV = Sheets("RVData")
lastrowV = wsV.Cells(Rows.Count, 2).End(xlUp).Row
Dim O As Long
With Sheets("RVData")
    For O = 6 To lastrowV
        .Cells(O, 1).RowHeight = 27
    Next O
End With
With wsV.Range("B6:N" & lastrowV)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
'define width to colums
wsV.Range("B:B").ColumnWidth = 26.57
wsV.Range("C:C").ColumnWidth = 16.29
wsV.Range("D:D").ColumnWidth = 31.14
wsV.Range("E:E").ColumnWidth = 17.71
wsV.Range("F:F").ColumnWidth = 15.29
wsV.Range("G:G").ColumnWidth = 14.71
wsV.Range("H:H").ColumnWidth = 13
wsV.Range("I:I").ColumnWidth = 17.57
wsV.Range("J:J").ColumnWidth = 17.57
wsV.Range("K:K").ColumnWidth = 17.57
wsV.Range("N:N").ColumnWidth = 14.3
wsV.Columns("L:L").EntireColumn.Hidden = True
wsV.Columns("M:M").EntireColumn.Hidden = True
'deselected report
ReportsI.VAC_BReport = False

'deselected report
shtbron.Copy wbdoel.Sheets("RVData").Range("A1")
wbdoel.Sheets("RVData").Range("D1").Value = "REPORTE VACACIONES - IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
wbdoel.Sheets("VData").Visible = False
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True

Call OptimizeCode_End 'Optimize Code
End Sub
