Attribute VB_Name = "Mod_Retired_Report"
'=====================
'Variable to Speed up the code
'=====================
Global Retired_Emp As Boolean
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
Sub Retired_ReportOptions()
    If MsgBox("¿Desea Exportar el Reporte?", vbYesNo) = vbYes Then
    MsgBoxAnswer = MsgBoxCB("Seleccione el formato", "EXCEL")
        If MsgBoxAnswer = 1 Then
            ReportsI.Hide
            If MsgBox("¿El reporte es para personal ajeno al departamento de RRHH?", vbYesNo) = vbYes Then
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
Sub Retired_Report()
Call OptimizeCode_Begin 'Optimize Code

Call OptimizeCode_End 'Optimize Code
End Sub
Sub RemoveCalculatedFields()
Dim pt As PivotTable, pf As PivotField, pfNew As PivotField, strSource As String, strFormula As String
Set pt = Sheets("RPDataR").PivotTables(1)
    For Each pf In pt.CalculatedFields
        strSource = pf.SourceName
        strFormula = pf.Formula
        pf.Delete
        Set pfNew = pt.CalculatedFields.Add(strSource, strFormula)
    Next pf
End Sub
Sub Export_RetiredReport()
Call OptimizeCode_Begin 'Optimize Code
Dim wsEA As Worksheet
Set wsEA = Sheets("RPDataR")
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
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("PData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    Dim r As Integer
    For r = wbdoel.Sheets("PData").UsedRange.Rows.Count To 1 Step -1
        If Cells(r, "AM") = False Then
            wbdoel.Sheets("PData").Rows(r).EntireRow.Delete
        End If
    Next
Worksheets.Add().Name = "RPDataR"

'Declare Variables
Dim PvtTbl As PivotTable, wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet
'Set the value of variables
Set wsData = Worksheets("PData")
Set wsPvtTbl = Worksheets("RPDataR")
lastrow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = wsData.Cells(1, 57).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RPDataR").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData, Version:=xlPivotTableVersion12)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion12)
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
With PvtTbl.PivotFields("FECHA DE RETIRO")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("MOTIVO")
    .Orientation = xlRowField: .Position = 10
End With
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
'Autofit the columns to its content
Worksheets("RPDataR").Columns("B:J").AutoFit
'PivotTable.ShowTableStyleColumnHeaders Property. Set to True to display the column headers in the PivotTable.
PvtTbl.ShowTableStyleColumnHeaders = True
'PivotTable.ShowTableStyleRowHeaders Property. Set to True to display the row headers in the PivotTable.
PvtTbl.ShowTableStyleRowHeaders = False
'PivotTable.ShowTableStyleColumnStripes Property. Set to True to display the banded columns in the PivotTable.
PvtTbl.ShowTableStyleColumnStripes = True
'PivotTable.ShowTableStyleRowStripes Property. Set to True to display the banded rows in the PivotTable.
PvtTbl.ShowTableStyleRowStripes = True
'HIDE THE COLUMN WITH OTHER DATA
'CENTER TEXT
With Sheets("RPDataR").Range("B7:O" & lastrow)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
Dim u As Long
With Sheets("RPDataR")
    For u = 7 To lastrow
        .Cells(u, 1).RowHeight = 27.5
    Next u
End With
'add width to row six
Dim wsAC As Worksheet
Set wsAC = Sheets("RPDataR")
wsAC.Cells(7, 2).RowHeight = 27.5
'define width to colums
wsAC.Range("B:B").ColumnWidth = 26.57
wsAC.Range("C:C").ColumnWidth = 15.57
wsAC.Range("D:D").ColumnWidth = 29.43
wsAC.Range("E:E").ColumnWidth = 16
wsAC.Range("F:F").ColumnWidth = 19.29
wsAC.Range("G:G").ColumnWidth = 14
wsAC.Range("H:H").ColumnWidth = 32.57
wsAC.Range("I:I").ColumnWidth = 17.71
wsAC.Range("K:K").ColumnWidth = 29
wsAC.Range("J:J").ColumnWidth = 14
wsAC.Range("L:L").ColumnWidth = 17
wsAC.Range("M:M").ColumnWidth = 13.14
wsAC.Range("N:N").ColumnWidth = 12.86
wsAC.Range("O:O").ColumnWidth = 11.86
'Assign the style to the pivotable
PvtTbl.TableStyle2 = "PivotStyleMedium9"
ReportsI.Retired_Emp = False

shtbron.Copy wbdoel.Sheets("RPDataR").Range("A1")
wbdoel.Sheets("RPDataR").Range("D1").Value = "PERSONAL RETIRADO - IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
wbdoel.Sheets("PData").Visible = False
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True

bCallCode = False
Call OptimizeCode_End 'Optimize Code
End Sub
Sub exportPivotT()
Call OptimizeCode_Begin 'Optimize Code
'=======================================================
'EXPORT DATA TO OTHER STAFF NOT RRHH
'=======================================================
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("PData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    Dim r As Integer
    For r = wbdoel.Sheets("PData").UsedRange.Rows.Count To 1 Step -1
        If Cells(r, "AM") = False Then
            wbdoel.Sheets("PData").Rows(r).EntireRow.Delete
        End If
    Next
Worksheets.Add().Name = "RPDataR"
Sheets("PData").Columns("C:j").Delete Shift:=xlToLeft
Sheets("PData").Columns("O:Q").Delete Shift:=xlToLeft
'Declare Variables
Dim PvtTbl As PivotTable, wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet
'Set the value of variables
Set wsData = Worksheets("PData")
Set wsPvtTbl = Worksheets("RPDataR")
lastrow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = wsData.Cells(1, 30).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RPDataR").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData, Version:=xlPivotTableVersion15)
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
With PvtTbl.PivotFields("FECHA DE RETIRO")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("MOTIVO")
    .Orientation = xlRowField: .Position = 10
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
Dim lastRowR As Long
lastRowR = Sheets("RPDataR").Cells(Rows.Count, 2).End(xlUp).Row
'CENTER TEXT
With Sheets("RPDataR").Range("B7:O" & lastRowR)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
Dim u As Long
With Sheets("RPDataR")
    For u = 7 To lastRowR
        .Cells(u, 1).RowHeight = 27.5
    Next u
End With
'add width to row six
Dim wsAC As Worksheet
Set wsAC = Sheets("RPDataR")
wsAC.Cells(7, 2).RowHeight = 27.5
'define width to colums
wsAC.Range("B:B").ColumnWidth = 26.57
wsAC.Range("C:C").ColumnWidth = 15.57
wsAC.Range("D:D").ColumnWidth = 29.43
wsAC.Range("E:E").ColumnWidth = 16
wsAC.Range("F:F").ColumnWidth = 19.29
wsAC.Range("G:G").ColumnWidth = 14
wsAC.Range("H:H").ColumnWidth = 32.57
wsAC.Range("I:I").ColumnWidth = 17.71
wsAC.Range("K:K").ColumnWidth = 29
wsAC.Range("J:J").ColumnWidth = 14
wsAC.Range("M:M").ColumnWidth = 13.14
wsAC.Range("N:N").ColumnWidth = 12.86
wsAC.Range("O:O").ColumnWidth = 11.86
'Assign the style to the pivotable
PvtTbl.TableStyle2 = "PivotStyleMedium9"
shtbron.Copy wbdoel.Sheets("RPDataR").Range("A1")
wbdoel.Sheets("RPDataR").Range("D1").Value = "REPORTE RETIRADO - IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
wbdoel.Sheets("PData").Visible = False
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True
Call OptimizeCode_End 'Optimize Code
End Sub
