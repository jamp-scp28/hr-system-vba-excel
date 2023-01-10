Attribute VB_Name = "Mod_absReportes"
'=====================
'Variable to Speed up the code
'=====================
Option Explicit
Global ABS_BReport As Boolean
Public bCallCode As Boolean
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean
Public MsgBoxAnswer As Long
Global AbsDescription As String 'Get the description of code CEI10 and register in the data
Global ITrack As Boolean 'Send code to register the data in absenteeism form to illtrack
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
Sub ABS_ExportOption()
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
Sub Export_ABSReport1()
Call OptimizeCode_Begin 'Optimize Code
Dim wsEA As Worksheet: Set wsEA = Sheets("RAData")
'Export PivotTable as PDF
Dim fName       As String
Dim FPath       As String
With wsEA
Dim lr As Long
    fName = wsEA.Range("D1").Value
    lr = wsEA.Range("B" & Rows.Count).End(xlUp).Row + 6
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
'===================================
' EXPORT DATA TO A NEW WORKBOOK
'===================================
Dim wbbron      As Workbook
Dim wbdoel      As Workbook
Dim shtbron     As Range
Dim shtdoel     As Worksheet
Dim shtbron2    As Worksheet
Dim PvtFld      As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("AData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
Worksheets.Add().Name = "RAData"
'Declare Variables
'Sheets("AData").Columns("H:H").Delete Shift:=xlToLeft

'delete column with sensible data
Sheets("AData").Columns(Sheets("AData").[abs_wage].Column).Delete Shift:=xlToLeft

Dim PvtTbl      As PivotTable
Dim wsData      As Worksheet
Dim rngData     As Range
Dim PvtTblCache As PivotCache
Dim wsPvtTbl    As Worksheet
'Set the value of variables
Set wsData = Worksheets("AData")
Set wsPvtTbl = Worksheets("RAData")
Dim lastrow As Long
Dim lastColumn As Long
lastrow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Dim nuevo As Range
Set nuevo = Worksheets("RAData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=("AData!R1C1:R" & lastrow & "C" & lastColumn), Version:=xlPivotTableVersion15)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTableDd", DefaultVersion:=xlPivotTableVersion15)
PvtTbl.RowAxisLayout xlTabularRow
'Asign the position to every field
With PvtTbl.PivotFields("FECHA")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("COLABORADOR")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("IDENTIFICACION")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("ÁREA")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("CARGO")
    .Orientation = xlRowField: .Position = 5
End With
With PvtTbl.PivotFields("TIPO AUSENCIA")
    .Orientation = xlRowField: .Position = 6
End With
With PvtTbl.PivotFields("CIE10")
    .Orientation = xlRowField: .Position = 7
End With
With PvtTbl.PivotFields("FECHA INICIAL")
    .Orientation = xlRowField: .Position = 8
End With
With PvtTbl.PivotFields("FECHA FINAL")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("NO. DIAS AUSENCIA")
    .Orientation = xlRowField: .Position = 10
End With
With PvtTbl.PivotFields("NO. HORAS AUSENCIA")
    .Orientation = xlRowField: .Position = 11
End With
With PvtTbl.PivotFields("CAUSA")
    .Orientation = xlRowField: .Position = 12
End With
With PvtTbl
    For Each PvtFld In .PivotFields
        PvtFld.Subtotals(1) = False
        PvtFld.Subtotals(1) = False
    Next PvtFld
End With
'Create calculate fields for pivottable style
 'PvtTbl.AddDataField PvtTbl _
 '       .PivotFields("COSTO POR AUSENCIA"), "COSTO", xlSum
 '   With PvtTbl.PivotFields("COSTO")
 '       .Caption = "COSTO": .Function = xlSum: .NumberFormat = "_($ * #,##0_);_($ * (#,##0);_($ * ""-""_);_(@_)"
 '   End With
'Desactivate the calculation of row and column grand
With PvtTbl
    .ColumnGrand = True: .RowGrand = False
End With
'Autofit the columns to its content
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
'======================================================================
'Code to generate and create the charts
'======================================================================
 Worksheets.Add().Name = "DashBoard"
 Sheets("DashBoard").Select
 
'=======================================================================
'Create firts chart to analize the data of absents with the hours
'========================================================================
Dim ChPvtTbl      As PivotTable
Dim ChwsData      As Worksheet
Dim ChrngData     As Range
Dim ChPvtTblCache As PivotCache
Dim ChwsPvtTbl    As Worksheet
'Set the value of variables
Set ChwsData = Worksheets("AData")
Set ChwsPvtTbl = Worksheets("DashBoard")
Dim ChlastRow As Long
Dim ChlastColumn As Long
ChlastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
ChlastColumn = wsData.Cells(1, 15).Column
Set ChrngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist

For Each ChPvtTbl In ChwsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
ChPvtTbl.TableRange2.Clear
End If
Next ChPvtTbl
'Create the new pivotable
Dim Chnuevo       As Range
Set Chnuevo = Worksheets("DashBoard").Cells(7, 1)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=("AData!R1C1:R" & lastrow & "C" & lastColumn), Version:=xlPivotTableVersion15)
Set ChPvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=Chnuevo, TableName:="PivotTableDash", DefaultVersion:=xlPivotTableVersion15)
ChPvtTbl.RowAxisLayout xlTabularRow
 
With ChwsPvtTbl.PivotTables("PivotTableDash")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
 
    ChwsPvtTbl.PivotTables("PivotTableDash").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = True
    ChwsPvtTbl.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("DashBoard!$A$7:$C$19")
    With ActiveChart.PivotLayout.PivotTable.PivotFields("FECHA INICIAL")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.PivotFields("FECHA INICIAL").AutoGroup
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("NO. HORAS AUSENCIA"), "Suma de NO. HORAS AUSENCIA", _
        xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("NO. HORAS AUSENCIA"), "Suma de NO. HORAS AUSENCIA2", _
        xlSum
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "Suma de NO. HORAS AUSENCIA")
        .NumberFormat = "0"
        .Caption = "TOTAL HORAS AUSENTES"
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "Suma de NO. HORAS AUSENCIA2")
        .Function = xlCount
        .Caption = "CANTIDAD DE AUSENCIAS"
    End With
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlLine
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).AxisGroup = 2
    ActiveChart.ApplyLayout (4)
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "Ausentismos * Mes"
    With ActiveSheet.ChartObjects("Gráfico 1")
        .Left = Sheets("DashBoard").Range("D7").Left
        .Top = Sheets("DashBoard").Range("D7").Top
    End With
    
'=======================================================================
'Create second chart to analize the data of absents with the type of absent
'========================================================================
Dim ChPvtTbl2      As PivotTable
Dim ChPvtTblCache2 As PivotCache
Dim ChrngData2     As Range
Dim Chnuevo2 As Range
Set Chnuevo2 = Worksheets("DashBoard").Cells(22, 1)
Set ChPvtTblCache2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=("AData!R1C1:R" & lastrow & "C" & lastColumn), Version:=xlPivotTableVersion15)
Set ChPvtTbl2 = PvtTblCache.CreatePivotTable(TableDestination:=Chnuevo2, TableName:="PivotTableDash2", DefaultVersion:=xlPivotTableVersion15)
ChPvtTbl2.RowAxisLayout xlTabularRow
 
With ChwsPvtTbl.PivotTables("PivotTableDash2")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
 
    ChwsPvtTbl.PivotTables("PivotTableDash2").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = True
    ChwsPvtTbl.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("DashBoard!$A$23:$C$27")
    With ActiveChart.PivotLayout.PivotTable.PivotFields("TIPO AUSENCIA")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("NO. HORAS AUSENCIA"), "Suma de NO. HORAS AUSENCIA2", _
        xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("NO. HORAS AUSENCIA"), "Suma de NO. HORAS AUSENCIA3", _
        xlSum
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "Suma de NO. HORAS AUSENCIA2")
        .NumberFormat = "0"
        .Caption = "TOTAL HORAS AUSENTES"
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "Suma de NO. HORAS AUSENCIA3")
        .Function = xlCount
        .Caption = "CANTIDAD DE AUSENCIAS"
    End With
    
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlLine
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).AxisGroup = 2
    ActiveChart.ApplyLayout (4)
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "Ausentismos * Tipo"
    With ActiveSheet.ChartObjects("Gráfico 2")
        .Left = Sheets("DashBoard").Range("J7").Left
        .Top = Sheets("DashBoard").Range("J7").Top
    End With
'=======================================================================
'Create third chart to analize the data of absents with the CEI_10 description
'========================================================================

Dim ChPvtTbl3      As PivotTable
Dim ChPvtTblCache3 As PivotCache
Dim ChrngData3     As Range
Dim Chnuevo3 As Range
Set Chnuevo3 = Worksheets("DashBoard").Cells(35, 1)
Set ChPvtTblCache3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=("AData!R1C1:R" & lastrow & "C" & lastColumn), Version:=xlPivotTableVersion15)
Set ChPvtTbl3 = PvtTblCache.CreatePivotTable(TableDestination:=Chnuevo3, TableName:="PivotTableDash3", DefaultVersion:=xlPivotTableVersion15)
ChPvtTbl3.RowAxisLayout xlTabularRow
 
With ChwsPvtTbl.PivotTables("PivotTableDash3")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
 
    ChwsPvtTbl.PivotTables("PivotTableDash3").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = True
    ChwsPvtTbl.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("DashBoard!$A$35:$C$70")
    With ActiveChart.PivotLayout.PivotTable.PivotFields("DESCRIPCION CIE10")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("NO. HORAS AUSENCIA"), "Suma de NO. HORAS AUSENCIA4", _
        xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("NO. HORAS AUSENCIA"), "Suma de NO. HORAS AUSENCIA5", _
        xlSum
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "Suma de NO. HORAS AUSENCIA4")
        .NumberFormat = "0"
        .Caption = "TOTAL HORAS AUSENTES"
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields( _
        "Suma de NO. HORAS AUSENCIA5")
        .Function = xlCount
        .Caption = "CANTIDAD DE AUSENCIAS"
    End With
    
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlLine
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).AxisGroup = 2
    ActiveChart.ApplyLayout (4)
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "Enf. General * Tipo"
    
    With ActiveSheet.ChartObjects("Gráfico 3")
        .Left = Sheets("DashBoard").Range("D22").Left
        .Top = Sheets("DashBoard").Range("D22").Top
    End With
    ActiveSheet.Shapes("Gráfico 3").ScaleWidth 1.9965277778, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Gráfico 3").ScaleHeight 1.9965277778, msoFalse, _
        msoScaleFromTopLeft
    '===================================================
    'ADD SLICERCACHE FO FILTER DATA FOR MONTH 'error to fix ===============
    '===================================================
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTableDash"), _
        "Años").Slicers.Add ActiveSheet, , "Años", "Años", 0, 213, 144, 198
    
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Años").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("PivotTableDash2"))
    
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Años").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("PivotTableDash3"))
        
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Años").Slicers("Años"). _
        NumberOfColumns = 12
        
    ActiveSheet.Shapes("Años").ScaleWidth 3.34375, msoFalse, msoScaleFromTopLeft
    
    ActiveSheet.Shapes("Años").ScaleHeight 0.3596226415, msoFalse, _
        msoScaleFromTopLeft
        
        
    '===================================================
    'ADD SLICERCACHE FO FILTER DATA FOR MONTH
    '===================================================
    
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTableDash"), _
        "TIPO AUSENCIA").Slicers.Add ActiveSheet, , "TIPO AUSENCIA", "TIPO AUSENCIA", _
        0, 45, 144, 198.75
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_TIPO_AUSENCIA").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("PivotTableDash2"))
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_TIPO_AUSENCIA").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("PivotTableDash3"))
    
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_TIPO_AUSENCIA").Slicers("TIPO AUSENCIA"). _
        NumberOfColumns = 2
    ActiveSheet.Shapes("TIPO AUSENCIA").ScaleHeight 0.3596226415, msoFalse, _
        msoScaleFromTopLeft
        
    '===================================================
    'ADD SLICERCACHE FO FILTER DATA FOR MONTH
    '===================================================
    
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PivotTableDash"), _
        "ÁREA").Slicers.Add ActiveSheet, , "ÁREA", "ÁREA", 0, 720, 144, 198.75
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_ÁREA").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("PivotTableDash2"))
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_ÁREA").PivotTables. _
        AddPivotTable (ActiveSheet.PivotTables("PivotTableDash3"))
    
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_ÁREA").Slicers("ÁREA"). _
        NumberOfColumns = 4
    ActiveSheet.Shapes("ÁREA").ScaleWidth 1.4201357536, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("ÁREA").ScaleHeight 0.3596226415, msoFalse, _
        msoScaleFromTopLeft
'====================================================
'Code to give format to the sheet of report
'====================================================
Dim lastRowA        As Long
Dim lastRowA2       As Long
Dim copyC           As Boolean
Dim wsA             As Worksheet
lastRowA = Sheets("RAData").Cells(Rows.Count, 2).End(xlUp).Row
'set variables to copy the conventions
copyC = True
Set wsA = Sheets("RAData")
lastRowA2 = lastRowA + 2
'conditional to copy the conventions
If copyC = True Then
wsA.Cells(lastRowA2, 2).Value = "CONVENCIONES"
wsA.Cells(lastRowA2 + 1, 2).Value = "E.P: ENFERMEDAD PROFESIONAL"
wsA.Cells(lastRowA2 + 2, 2).Value = "E.L.: LICENCIA DE MATERNIDAD"
wsA.Cells(lastRowA2 + 3, 2).Value = "A.T.: ACCIDENTE DE TRABAJO"
wsA.Cells(lastRowA2 + 4, 2).Value = "E.G.: ENFERMEDAD GENERAL"
'SECOND COLUMN
wsA.Cells(lastRowA2 + 1, 4).Value = "C.D.: CALAMIDAD DÓMESTICA"
wsA.Cells(lastRowA2 + 2, 4).Value = "P.S.T.: PERMISO SOLICITADO POR EL TRABAJADOR"
wsA.Cells(lastRowA2 + 3, 4).Value = "F.S.P.: FALLA SIN PERMISO"
wsA.Cells(lastRowA2 + 4, 4).Value = "V.: VACACIONES"
End If
Dim u As Long
With Sheets("RAData")
    For u = 7 To lastRowA
        .Cells(u, 1).RowHeight = 38
    Next u
End With
With Sheets("RAData").Range("B6:M" & lastRowA)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
    'Copy format
    Dim wsC     As Range
    Set wsC = Sheets("RAData").Range("B6")
    wsC.Copy
    Dim wsP     As Range
    Set wsP = Sheets("RAData").Range("B" & lastRowA + 2 & ":F" & lastRowA + 6)
    wsP.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    With Sheets("RAData").Range("B" & lastRowA + 2 & ":F" & lastRowA + 2)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Font.Bold = True
        .WrapText = False: .Orientation = 0: .AddIndent = False: .IndentLevel = 0
        .ShrinkToFit = False: .ReadingOrder = xlContext: .MergeCells = False
    End With
    With Sheets("RAData").Range("B" & lastRowA + 3 & ":F" & lastRowA + 6)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .Font.Bold = False
        .WrapText = False: .Orientation = 0: .AddIndent = False: .IndentLevel = 0
        .ShrinkToFit = False: .ReadingOrder = xlContext: .MergeCells = False
    End With
'add width to row six
Sheets("RAData").Cells(6, 2).RowHeight = 31.5
'define width to colums
Sheets("RAData").Range("B:B").ColumnWidth = 13.5
Sheets("RAData").Range("C:C").ColumnWidth = 26.57
Sheets("RAData").Range("D:D").ColumnWidth = 17
Sheets("RAData").Range("E:E").ColumnWidth = 7.71
Sheets("RAData").Range("F:F").ColumnWidth = 26.29
Sheets("RAData").Range("G:G").ColumnWidth = 10.29
Sheets("RAData").Range("H:H").ColumnWidth = 13
Sheets("RAData").Range("I:I").ColumnWidth = 18.23
Sheets("RAData").Range("J:J").ColumnWidth = 18.23
Sheets("RAData").Range("K:K").ColumnWidth = 10.43
Sheets("RAData").Range("L:L").ColumnWidth = 10.86
Sheets("RAData").Range("M:M").ColumnWidth = 35.29
Sheets("RAData").Range("N:N").ColumnWidth = 15.57

shtbron.Copy wbdoel.Sheets("RAData").Range("A1")
wbdoel.Sheets("RAData").Range("D1").Value = "AUNSENTISMOS IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S."
wbdoel.Sheets("AData").Visible = False
'deselected report
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True
'====================================================
'Code to give format to the sheet of report
'====================================================
bCallCode = False
Call OptimizeCode_End 'Optimize Code
End Sub

