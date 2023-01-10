Attribute VB_Name = "ModToolsForData"
Sub ReplaceData()
'
' Macro1 Macro
'
' Acceso directo: Ctrl+Mayús+Q
'
Dim lr As Long
Dim StrName As String

StrName = "Actas"
lr = Sheets("DData").Cells(Rows.Count, 1).End(xlUp).Row

Dim Rng As Range, r As Long, rgn As Range
Set Rng = Sheets("DData").Range("A2:E" & lr)
For r = Rng.Count To 1 Step -1
    If Rng(r).Value = "Certificaciones Laborales" Then
        Rng(r + 1).EntireRow.Insert
        'rng(r + 1).Value = StrName
        'rng(r).Offset(0, -1).Copy rgn.Offset(1, 0)
        Set rgn = Rng(r).Offset(-1, -1)
        Rng(r).Offset(0, -1).Select
        Selection.Value = rgn.Value
        Rng(r).Offset(0, 0).Select
        Selection.Value = StrName
        Rng(r).Offset(0, 1).Select
        Selection.Value = rgn.Value & "-" & StrName
        Rng(r).Offset(0, 2).Select
        Selection.Value = "REV"
        Rng(r).Offset(0, 3).Select
        Selection.Value = "-"
    End If
Next r
End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' Macro1 Macro
'
' Acceso directo: Ctrl+Mayús+Q
'
Dim Rng As Range, r As Long, rgn As Range
Set Rng = Sheets("DData").Range("A2:E2262")
For r = Rng.Count To 1 Step -1
    If Rng(r).Value = "Cuenta Bancaria" Then
        Rng(r + 1).EntireRow.Insert
        Rng(r).Value = "Documento de Precisión y Ratificación"
        Set rgn = Rng(r).Offset(-1, -1)
        Rng(r).Offset(0, -1).Select
        Selection.Value = rgn.Value
        Rng(r).Offset(0, 2).Select
        Selection.Value = "NO APLICA"
        Rng(r).Offset(0, 3).Select
        Selection.Value = "-"
    End If
Next r
End Sub
Sub Macro2()
'
' Macro1 Macro
'
' Acceso directo: Ctrl+Mayús+Q
'
Dim Rng As Range, r As Long, rgn As Range, rng3 As Long
Set Rng = Sheets("DData").Range("A2:E1837")
For r = Rng.Count To 1 Step -1
    If Rng(r).Value = "Acta de Inducción" Then
        Rng(r + 1).EntireRow.Insert
        Rng(r).Value = "Cuenta Bancaria"
        Set rgn = Rng(r).Offset(-1, -1)
        Rng(r).Offset(0, -1).Select
        Selection.Value = rgn.Value
        Rng(r).Offset(0, 2).Select
        Selection.Value = "Sí"
        Rng(r).Offset(0, 3).Select
        Selection.Value = "NA"
        'ADD NEW LINE FOR BANK ACCOUNT
        Rng(r).Offset(0, 1).Select
        Selection.Value = rgn.Offset(-1, 2).FormulaR1C1
    End If
Next r
End Sub
Sub Macro3()
'
' Macro1 Macro
'
' Acceso directo: Ctrl+Mayús+Q
'
Dim Rng As Range, r As Long, rgn As Range
Set Rng = Sheets("DData").Range("A2:E1855")
For r = Rng.Count To 1 Step -1
    If Rng(r).Value = "Formato Entrevista" Then
        Set rgn = Rng(r).Offset(0, -1)
        rgn.Style = "Énfasis5"
        Set rgn1 = Rng(r).Offset(0, 0)
        rgn1.Style = "Énfasis5"
        Set rgn2 = Rng(r).Offset(0, 1)
        rgn2.Style = "Énfasis5"
        Set rgn3 = Rng(r).Offset(0, 2)
        rgn3.Style = "Énfasis5"
        Set rgn4 = Rng(r).Offset(0, 3)
        rgn4.Style = "Énfasis5"
    End If
Next r
End Sub
Sub DeleteBadRefs()
  Dim nm As Name
  
  For Each nm In ActiveWorkbook.Names
    If nm = "dsfsa" Then
        nm.Delete
        MsgBox "complete"
    End If
    
    'If InStr(1, nm.RefersTo, "#REF!") > 0 Then
     ' 'List the name before deleting
     ' Debug.Print nm.Name & ": deleted"
     ' nm.Delete
    'End If
  Next nm
End Sub
Sub Delete_Names()
ActiveWorkbook.Names("nw_nospecific").Delete
End Sub

Sub ExtractRangeName()
Dim page As Worksheet
Dim nm As Name

Dim nmSheet As Worksheet
Set nmSheet = Sheets("references")
Dim lastrow As Long

lastrow = nmSheet.Cells(Rows.Count, 1).End(xlUp).Row

For Each nm In ActiveWorkbook.Names
    nmSheet.Cells(lastrow, 1).Value = nm.Name
    nmSheet.Cells(lastrow, 2).Value = nm
    lastrow = lastrow + 1
Next

End Sub

Public Function ToColNum(ColN)
    ToColNum = Range(ColN & 1).Column
End Function
