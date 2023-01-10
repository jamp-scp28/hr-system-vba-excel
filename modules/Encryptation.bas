Attribute VB_Name = "Encryptation"
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean
Public DataFLogin As Boolean
Sub Encrypt()
'Speed Up the Code
    Application.ScreenUpdating = False
    EventState = Application.EnableEvents
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    PageBreakState = ActiveSheet.DisplayPageBreaks
    ActiveSheet.DisplayPageBreaks = False
'===============================================

Dim sKey As String
sKey = 28
Dim SheetsFG As Worksheet
Dim RangeFG As Range
Dim lastRowFG
Dim rFG As Range
'Variables for Email
Dim PeRange As Range
Dim CeRange As Range
Dim rPe As Range
Dim rCe As Range
'Encrypt the first group of data bases
Dim mySheetsFG As Sheets
Set mySheetsFG = Sheets(Array(PData.Name, RData.Name, VData.Name, PlatF.Name, IData.Name, APData.Name))

For Each SheetsFG In mySheetsFG
    
        lastRowFG = SheetsFG.Cells(Rows.Count, 1).End(xlUp).Row
        Set RangeFG = SheetsFG.Range("A2:B" & lastRowFG)
        For Each rFG In RangeFG
            rFG.Value = XorC(rFG.Value, sKey)
        Next rFG
        If SheetsFG.Name Like "PData" Then
            Set PeRange = Sheets("PData").Range("K2:K" & lastRowFG)
            Set CeRange = Sheets("PData").Range("N2:N" & lastRowFG)
            For Each rPe In PeRange
                rPe.Value = XorC(rPe.Value, sKey)
            Next rPe
            For Each rCe In CeRange
                rCe.Value = XorC(rCe.Value, sKey)
            Next rCe
        End If
Next SheetsFG


'Finish Speed Up - Reset the value to default
    ActiveSheet.DisplayPageBreaks = PageBreakState
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = EventState
    Application.ScreenUpdating = True
'================================================
End Sub

Sub Encrypt2()
'Speed Up the Code
    Application.ScreenUpdating = False
    EventState = Application.EnableEvents
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    PageBreakState = ActiveSheet.DisplayPageBreaks
    ActiveSheet.DisplayPageBreaks = False
'===============================================

'Encrypt the Second Group of Sheets
Dim sKey2 As String
sKey2 = 27
Dim SheetsSG As Worksheet
Dim lastRowSG As Long
Dim RangeSG As Range
Dim rSG As Range
Dim mySheetsSG As Sheets
Set mySheetsSG = Sheets(Array(DData.Name, SData.Name, AData.Name))

For Each SheetsSG In mySheetsSG
    
        lastRowSG = SheetsSG.Cells(Rows.Count, 1).End(xlUp).Row
        If SheetsSG.Name Like "DData" Then
            Set RangeSG = SheetsSG.Range("A2:C" & lastRowSG)
        ElseIf SheetsSG.Name Like "AData" Then
            Set RangeSG = SheetsSG.Range("C2:E" & lastRowSG)
        ElseIf SheetsSG.Name Like "SData" Then
            Set RangeSG = SheetsSG.Range("B2:C" & lastRowSG)
        End If
        
        For Each rSG In RangeSG
            rSG.Value = XorC(rSG.Value, sKey2)
        Next rSG
        
Next SheetsSG


'Finish Speed Up - Reset the value to default
    ActiveSheet.DisplayPageBreaks = PageBreakState
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = EventState
    Application.ScreenUpdating = True
'================================================
End Sub

Function XorC(ByVal ssData As String, ByVal sKey As String) As String
    Dim l As Long, i As Long, byIn() As Byte, byOut() As Byte, byKey() As Byte
    Dim bEncOrDec As Boolean
    'confirm valid string and key input:
    If Len(ssData) = 0 Or Len(sKey) = 0 Then XorC = "Invalid argument(s) used": Exit Function
    'check whether running encryption or decryption (flagged by presence of "xxx" at start of ssData):
    If Left$(ssData, 3) = "xxx" Then
        bEncOrDec = False   'decryption
        ssData = Mid$(ssData, 4)
    Else
        bEncOrDec = True   'encryption
    End If
    'assign strings to byte arrays (unicode)
    byIn = ssData
    byOut = ssData
    byKey = sKey
    l = LBound(byKey)
    For i = LBound(byIn) To UBound(byIn) - 1 Step 2
        byOut(i) = ((byIn(i) + Not bEncOrDec) Xor byKey(l)) - bEncOrDec 'avoid Chr$(0) by using bEncOrDec flag
        l = l + 2
        If l > UBound(byKey) Then l = LBound(byKey)  'ensure stay within bounds of Key
    Next i
    XorC = byOut
    If bEncOrDec Then XorC = "xxx" & XorC  'add "xxx" onto encrypted text
End Function
Sub CreateDocument()
Dim myFile As String, Rng As Range, cellValue As Variant, i As Integer, j As Integer
myFile = ActiveWorkbook.Path & "\data.csv"

Open myFile For Output As #1


Write #1, "Nombre", "CC", "FI", "FR", "FCesantias", "Cargo", "Salario", "Auxilio"
Write #1, "darlum", "ABC", "DCS", "ss", "ss", "fsfsd", "dfsd", "gsfs"

Close #1

End Sub
