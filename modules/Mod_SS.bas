Attribute VB_Name = "Mod_SS"
'=====================
'Variable to Speed up the code
'=====================
Public bCallCode As Boolean
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean
Global Active_Emp As Boolean
Public SensitiveInfo As Variant

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
Sub SocEco_ReportOptions()
    If MsgBox("¿Desea Exportar el Reporte?", vbYesNo) = vbYes Then
    MsgBoxAnswer = MsgBoxCB("Seleccione el formato", "EXCEL")
        If MsgBoxAnswer = 1 Then
            ReportsI.Hide
            If MsgBox("¿El reporte es para personal ajeno al departamento?", vbYesNo) = vbYes Then
                Call ExportSSinfo
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
Sub SocEco_Report()


Call OptimizeCode_End 'Optimize Code
End Sub
Sub Export_SSReport()
Call OptimizeCode_Begin 'Optimize Code
Dim wsEA As Worksheet
Set wsEA = Sheets("RSSData")
'Export PivotTable as PDF
Dim fName As String, FPath As String
With wsEA
    fName = wsEA.Range("D1").Value
    lr = wsEA.Range("B" & Rows.Count).End(xlUp).Row
End With
    wsEA.PageSetup.PrintArea = "B1:V" & lr
    ChDir ActiveWorkbook.Path
    wsEA.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        fName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
MsgBox "El documento de guardo se: " & ActiveWorkbook.Path
Call OptimizeCode_End 'Optimize Code
End Sub


Sub ExportSSinfo()
Call OptimizeCode_Begin 'Optimize Code
'==========================================
'EXPORT IN FORMAT XLM TO OTHER DEPAR
'===========================================
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("PData")

SensitiveInfo = Array("FECHA EXPEDICIÓN", "LUGAR EXPEDICIÓN", "RH", "PROFESIÓN", "TARJETA PROFESIONAL", "BARRIO", "LOCALIDAD", "TELEFONO FIJO", "E-MAIL CORPORATIVO", "TELEFONO MOVIL CORPORATIVO", "TELEFONO OFICINA - EXT", "ANTIGÜEDAD", "CODIGO DEPARTAMENTO", "TIPO DE CONTRATO", "SALARIO", "RODAMIENTO", "O AUXILIOS", "CENTRO DE TRABAJO", "CLASE", "FECHA DE COBERTURA", "FECHA RETIRO ARL")

Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    With ActiveSheet
        .Name = "PData": .Visible = True
    End With
'Worksheets.Add().Name = "RSSData"
Dim myRange As Range
Set myRange = shtbron2.Range("A1:BR1")
Call DelAsterisk
For Each cell In myRange
    
    If IsInArray(cell.Value, SensitiveInfo) Then
       'MsgBox cell.reference
       'MsgBox cell.Value
       'Columns(cell.Column).Delete
    End If
Next cell

End Sub



Sub ExportContractState()
Call OptimizeCode_Begin 'Optimize Code
'==========================================
'EXPORT IN FORMAT XLM TO OTHER DEPAR
'===========================================
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("PData")

SensitiveInfo = Array("FECHA EXPEDICIÓN", "LUGAR EXPEDICIÓN", "RH", "PROFESIÓN", "TARJETA PROFESIONAL", "BARRIO", "LOCALIDAD", "TELEFONO FIJO", "E-MAIL CORPORATIVO", "TELEFONO MOVIL CORPORATIVO", "TELEFONO OFICINA - EXT", "ANTIGÜEDAD", "CODIGO DEPARTAMENTO", "TIPO DE CONTRATO", "SALARIO", "RODAMIENTO", "O AUXILIOS", "CENTRO DE TRABAJO", "CLASE", "FECHA DE COBERTURA", "FECHA RETIRO ARL")

Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    With ActiveSheet
        .Name = "PData": .Visible = True
    End With
'Worksheets.Add().Name = "RSSData"
Dim myRange As Range
Set myRange = shtbron2.Range("A1:BR1")
Call DelAsteriskOS
For Each cell In myRange
    
    If IsInArray(cell.Value, SensitiveInfo) Then
       'MsgBox cell.reference
       'MsgBox cell.Value
       'Columns(cell.Column).Delete
    End If
Next cell

End Sub


Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function




Sub DelAsterisk()
    Dim i As Long, lc As Long
    'Assumes header is in row 1
    With Sheets("PData")
        lc = .Cells(1, Columns.Count).End(xlToLeft).Column
        For i = lc To 1 Step -1
            If InStr(.Cells(1, i), "*") > 0 Then
                .Cells(1, i).EntireColumn.Delete
            End If
        Next i
    End With
End Sub

Sub DelAsteriskOS()
    Dim i As Long, lc As Long
    'Assumes header is in row 1
    With Sheets("PData")
        lc = .Cells(1, Columns.Count).End(xlToLeft).Column
        For i = lc To 1 Step -1
            If InStr(.Cells(1, i), "-") > 0 Then
                .Cells(1, i).EntireColumn.Delete
            End If
        Next i
    End With
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
Worksheets.Add().Name = "RSSData"
'Declare Variables

Call OptimizeCode_Begin 'Optimize Code
'Declare Variables
Dim PvtTbl As PivotTable, wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet
'Set the value of variables
Set wsData = Worksheets("PData")
Set wsPvtTbl = Worksheets("RSSData")
lastrow = wsData.Cells(Rows.Count, 2).End(xlUp).Row
lastColumn = wsData.Cells(1, 57).Column
Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RSSData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData, Version:=xlPivotTableVersion15)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)
'Asign the position to every field
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("IDENTIFICACION")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("SEXO")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("EDAD")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("ESTADO CIVIL")
    .Orientation = xlRowField: .Position = 5
End With
With PvtTbl.PivotFields("ESCOLARIDAD")
    .Orientation = xlRowField: .Position = 6
End With
With PvtTbl.PivotFields("CIUDAD")
    .Orientation = xlRowField: .Position = 7
End With
With PvtTbl.PivotFields("DIRECCION")
    .Orientation = xlRowField: .Position = 8
End With
With PvtTbl.PivotFields("TELEFONO MOVIL")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("FECHA DE INGRESO")
    .Orientation = xlRowField: .Position = 10
End With
With PvtTbl.PivotFields("DEPARTAMENTO")
    .Orientation = xlRowField: .Position = 11
End With
With PvtTbl.PivotFields("CARGO")
    .Orientation = xlRowField: .Position = 12
End With
With PvtTbl.PivotFields("TIPO DE CONTRATO")
    .Orientation = xlRowField: .Position = 13
End With
With PvtTbl.PivotFields("EPS")
    .Orientation = xlRowField: .Position = 14
End With
With PvtTbl.PivotFields("AFP")
    .Orientation = xlRowField: .Position = 15
End With
With PvtTbl.PivotFields("CCF")
    .Orientation = xlRowField: .Position = 16
End With
With PvtTbl.PivotFields("ARL")
    .Orientation = xlRowField: .Position = 17
End With
With PvtTbl.PivotFields("ULTIMO EXAMEN MEDICO")
    .Orientation = xlRowField: .Position = 18
End With
With PvtTbl.PivotFields("CONDICION MEDICA")
    .Orientation = xlRowField: .Position = 19
End With
With PvtTbl.PivotFields("RECOMENDACIONES")
    .Orientation = xlRowField: .Position = 20
End With
With PvtTbl.PivotFields("RESTRICCIONES")
    .Orientation = xlRowField: .Position = 21
End With
With PvtTbl.PivotFields("RETIRADO")
    .Orientation = xlRowField: .Position = 22: .PivotItems("True").Visible = False 'Filter the data, no selecting TRUE values
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
Sheets("RSSData").Columns("W:W").EntireColumn.Hidden = True
With Sheets("RSSData").Range("B6:V" & lastrow)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
'Set space to rows
Dim u As Long
With Sheets("RSSData")
    For u = 7 To lastrow
        .Cells(u, 1).RowHeight = 38
    Next u
End With
'add width to row six
Dim wsAC As Worksheet
Set wsAC = Sheets("RSSData")
wsAC.Cells(7, 2).RowHeight = 27.5
'define width to colums
wsAC.Range("B:B").ColumnWidth = 24
wsAC.Range("C:C").ColumnWidth = 16.43
wsAC.Range("D:D").ColumnWidth = 13.14
wsAC.Range("E:E").ColumnWidth = 7.71
wsAC.Range("F:F").ColumnWidth = 12.3
wsAC.Range("G:G").ColumnWidth = 22
wsAC.Range("H:H").ColumnWidth = 15.86
wsAC.Range("I:I").ColumnWidth = 20.86
wsAC.Range("J:J").ColumnWidth = 14.43
wsAC.Range("K:K").ColumnWidth = 13.2
wsAC.Range("L:L").ColumnWidth = 19.57
wsAC.Range("M:M").ColumnWidth = 21.29
wsAC.Range("N:N").ColumnWidth = 15.14
wsAC.Range("O:O").ColumnWidth = 14
wsAC.Range("P:P").ColumnWidth = 16.14
wsAC.Range("Q:Q").ColumnWidth = 15
wsAC.Range("R:R").ColumnWidth = 12.52
wsAC.Range("Q:Q").ColumnWidth = 15
wsAC.Range("S:S").ColumnWidth = 12.29
wsAC.Range("T:T").ColumnWidth = 11.57
wsAC.Range("U:U").ColumnWidth = 73.43
wsAC.Range("V:V").ColumnWidth = 15
ReportsI.SDCS = False
shtbron.Copy wbdoel.Sheets("RSSData").Range("A1") 'copy headers
wbdoel.Sheets("RSSData").Range("D1").Value = "REPORTE INFORMACIÓN SOCIODEMOGRAFICA Y CONDICIÓN DE SALUD - IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
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
'EXPORT IN FORMAT XLM TO OTHER DEPAR
'===========================================
Dim wbbron As Workbook, wbdoel As Workbook, wbdole As Workbook, shtbron As Range, shtdoel As Worksheet, shtbron2 As Worksheet, SrcData As Range, pvt As PivotTable, PvtCache As PivotCache, PvtFld As PivotField
Application.ScreenUpdating = False
Set wbbron = ActiveWorkbook
Set shtbron2 = wbbron.Sheets("PData")
Set shtbron = wbbron.Sheets("BG").Rows("36:40")
SensitiveInfo = Array("WAGE", "AUXI1", "AUXI2")

Set wbdoel = Workbooks.Add
    shtbron2.Copy after:=wbdoel.Sheets(wbdoel.Sheets.Count)
    With ActiveSheet
        .Name = "PData": .Visible = True
    End With
Worksheets.Add().Name = "RSSData"

For Each register In SensitiveInfo
12
Next
'arrCol = Array("A", "C", "G", "H") 'always in ascending order
'For i = UBound(arrCol, 1) To 0 Step -1
'Columns(arrCol(i)).EntireColumn.Delete
'Next


'Declare Variables
Dim PvtTbl As PivotTable, wsData As Worksheet, rngData As Range, PvtTblCache As PivotCache, wsPvtTbl As Worksheet
'Set the value of variables
Set wsData = Worksheets("PData")
Set wsPvtTbl = Worksheets("RSSData")
lastrow = wsData.Cells(Rows.Count, 2).End(xlUp).Row
lastColumn = wsData.Cells(1, 63).Column
'Set rngData = wsData.Cells(1, 1).Resize(lastrow, lastColumn)
'Set rngData = "PData!R1C1:R345C67"
'Delete the pivottable if it already exist
For Each PvtTbl In wsPvtTbl.PivotTables
If MsgBox("Eliminar Tabla Dinámica Existente!", vbOKOnly) = vbOK Then
PvtTbl.TableRange2.Clear
End If
Next PvtTbl
'Create the new pivotable
Set nuevo = Worksheets("RSSData").Cells(6, 2)
Set PvtTblCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="PData!R1C1:R345C67", Version:=xlPivotTableVersion15)
Set PvtTbl = PvtTblCache.CreatePivotTable(TableDestination:=nuevo, TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)
'Asign the position to every field
With PvtTbl.PivotFields("APELLIDOS Y NOMBRES")
    .Orientation = xlRowField: .Position = 1
End With
With PvtTbl.PivotFields("IDENTIFICACION")
    .Orientation = xlRowField: .Position = 2
End With
With PvtTbl.PivotFields("SEXO")
    .Orientation = xlRowField: .Position = 3
End With
With PvtTbl.PivotFields("EDAD")
    .Orientation = xlRowField: .Position = 4
End With
With PvtTbl.PivotFields("ESTADO CIVIL")
    .Orientation = xlRowField: .Position = 5
End With
With PvtTbl.PivotFields("ESCOLARIDAD")
    .Orientation = xlRowField: .Position = 6
End With
With PvtTbl.PivotFields("CIUDAD")
    .Orientation = xlRowField: .Position = 7
End With
With PvtTbl.PivotFields("DIRECCION")
    .Orientation = xlRowField: .Position = 8
End With
With PvtTbl.PivotFields("TELEFONO MOVIL")
    .Orientation = xlRowField: .Position = 9
End With
With PvtTbl.PivotFields("FECHA DE INGRESO")
    .Orientation = xlRowField: .Position = 10
End With
With PvtTbl.PivotFields("DEPARTAMENTO")
    .Orientation = xlRowField: .Position = 11
End With
With PvtTbl.PivotFields("CARGO")
    .Orientation = xlRowField: .Position = 12
End With
With PvtTbl.PivotFields("TIPO DE CONTRATO")
    .Orientation = xlRowField: .Position = 13
End With
With PvtTbl.PivotFields("EPS")
    .Orientation = xlRowField: .Position = 14
End With
With PvtTbl.PivotFields("AFP")
    .Orientation = xlRowField: .Position = 15
End With
With PvtTbl.PivotFields("CCF")
    .Orientation = xlRowField: .Position = 16
End With
With PvtTbl.PivotFields("ARL")
    .Orientation = xlRowField: .Position = 17
End With
With PvtTbl.PivotFields("ULTIMO EXAMEN MEDICO")
    .Orientation = xlRowField: .Position = 18
End With
With PvtTbl.PivotFields("CONDICION MEDICA")
    .Orientation = xlRowField: .Position = 19
End With
With PvtTbl.PivotFields("RECOMENDACIONES")
    .Orientation = xlRowField: .Position = 20
End With
With PvtTbl.PivotFields("RESTRICCIONES")
    .Orientation = xlRowField: .Position = 21
End With
With PvtTbl.PivotFields("RETIRADO")
    .Orientation = xlRowField: .Position = 22: .PivotItems("True").Visible = False 'Filter the data, no selecting TRUE values
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
'Autofit the columns to its content
'Assign the style to the pivotable
shtbron.Copy wbdoel.Sheets("RSSData").Range("A1") 'copy headers
PvtTbl.TableStyle2 = "PivotStyleMedium9"
Sheets("RSSData").Columns("W:W").EntireColumn.Hidden = True
With Sheets("RSSData").Range("B6:V" & lastrow)
    .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .WrapText = True
End With
'Set space to rows
Dim u As Long
With Sheets("RSSData")
    For u = 7 To lastrow
        .Cells(u, 1).RowHeight = 38
    Next u
End With
'add width to row six
Dim wsAC As Worksheet
Set wsAC = Sheets("RSSData")
wsAC.Cells(7, 2).RowHeight = 27.5
'define width to colums
wsAC.Range("B:B").ColumnWidth = 24
wsAC.Range("C:C").ColumnWidth = 16.43
wsAC.Range("D:D").ColumnWidth = 13.14
wsAC.Range("E:E").ColumnWidth = 7.71
wsAC.Range("F:F").ColumnWidth = 12.3
wsAC.Range("G:G").ColumnWidth = 22
wsAC.Range("H:H").ColumnWidth = 15.86
wsAC.Range("I:I").ColumnWidth = 20.86
wsAC.Range("J:J").ColumnWidth = 14.43
wsAC.Range("K:K").ColumnWidth = 13.2
wsAC.Range("L:L").ColumnWidth = 19.57
wsAC.Range("M:M").ColumnWidth = 21.29
wsAC.Range("N:N").ColumnWidth = 15.14
wsAC.Range("O:O").ColumnWidth = 14
wsAC.Range("P:P").ColumnWidth = 16.14
wsAC.Range("Q:Q").ColumnWidth = 15
wsAC.Range("R:R").ColumnWidth = 12.52
wsAC.Range("Q:Q").ColumnWidth = 15
wsAC.Range("S:S").ColumnWidth = 12.29
wsAC.Range("T:T").ColumnWidth = 11.57
wsAC.Range("U:U").ColumnWidth = 73.43
wsAC.Range("V:V").ColumnWidth = 15

shtbron.Copy wbdoel.Sheets("RSSData").Range("A1")
wbdoel.Sheets("RSSData").Range("D1").Value = "REPORTE INFORMACIÓN SOCIODEMOGRAFICA Y CONDICIÓN DE SALUD - IMAGING EXPERTS AND HEALTHCARE SERVICES S.A.S"
wbdoel.Sheets("PData").Visible = False
Application.DisplayAlerts = False
Worksheets("hoja1").Delete
Application.DisplayAlerts = True
Call OptimizeCode_End 'Optimize Code
End Sub

