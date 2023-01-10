Attribute VB_Name = "Contract_Creator"
Dim Enterprise As Variant
Sub EnterpriseChoosed()
If MsgBox("¿Desea Exportar el Reporte?", vbYesNo) = vbYes Then
MsgBoxAnswer = MsgBoxCB("Seleccione la empresa", "RIMAB SAS", "IMEXHS SAS")
    If MsgBoxAnswer = 1 Then
        Enterprise = "RIMAB"
        Call Word_Certification_Acive_Personal_auxi1
    ElseIf MsgBoxAnswer = 2 Then
        Enterprise = "IMEXHS"
    End If
Else
Exit Sub
End If
End Sub

Sub Word_Certification_Acive_Personal()

'Codigo escrito por Manuel Vizcarra - wwww.combito.com
Dim data(0 To 1, 0 To 10) As String '(columna,fila)

patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " Certificado_Laboral_Activos.dotx"
Set objword = CreateObject("Word.Application")
objword.Visible = True
objword.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0

data(0, 0) = "[employee_name]"
data(1, 0) = Auto_Docs.Cells(2, Auto_Docs.[EMP_NAME].Column) '(fila,columna)
data(0, 1) = "[employee_id]"
data(1, 1) = Auto_Docs.Cells(2, Auto_Docs.[EMP_ID].Column)
data(0, 2) = "[day_word]"
data(1, 2) = Auto_Docs.Cells(2, Auto_Docs.[inc_day_word].Column)
data(0, 3) = "[day]"
data(1, 3) = Auto_Docs.Cells(2, Auto_Docs.[inc_day].Column)
data(0, 4) = "[inc_month]"
data(1, 4) = Auto_Docs.Cells(2, Auto_Docs.[inc_month].Column)
data(0, 5) = "[inc_dated]"
data(1, 5) = Auto_Docs.Cells(2, Auto_Docs.[inc_dated].Column)
data(0, 6) = "[job_name]"
data(1, 6) = Auto_Docs.Cells(2, Auto_Docs.[EMP_JOBNAME].Column)
data(0, 7) = "[word_wage]"
data(1, 7) = Auto_Docs.Cells(2, Auto_Docs.[word_wage].Column)
data(0, 8) = "[wage]"
data(1, 8) = Auto_Docs.Cells(2, Auto_Docs.[EMP_WAGE].Column)
data(0, 9) = "[type_contract]"
data(1, 9) = Auto_Docs.Cells(2, Auto_Docs.[type_contract].Column)
data(0, 10) = "[exp_dated]"
data(1, 10) = Format(Auto_Docs.Cells(2, Auto_Docs.[DATED_REGISTER].Column), "dd"" de ""mmmm"" de ""YYYY")
data(0, 11) = Auto_Docs.Cells(2, Auto_Docs.[EMP_AFP].Column)
data(1, 11) = "[emp_afp]"
data(0, 12) = Auto_Docs.Cells(2, Auto_Docs.[EMP_DORE].Column)
data(1, 12) = "[emp_dor]"
data(0, 13) = Auto_Docs.Cells(2, Auto_Docs.[word_emp_dor].Column)
data(1, 13) = "[word_emp_dor]"
data(0, 14) = Auto_Docs.Cells(2, Auto_Docs.[month_retired].Column)
data(1, 14) = "[month_retired]"
data(0, 15) = Auto_Docs.Cells(2, Auto_Docs.[year_retired].Column)
data(1, 15) = "[year_retired]"

For i = 0 To UBound(data, 2)

    textobuscar = data(0, i)
    objword.Selection.Move 6, -1
    objword.Selection.Find.Execute FindText:=textobuscar

    While objword.Selection.Find.found = True
        objword.Selection.Text = data(1, i) 'texto a reemplazar
        objword.Selection.Move 6, -1
        objword.Selection.Find.Execute FindText:=textobuscar
    Wend

Next i

objword.Activate

End Sub
Sub Word_Certification_Acive_Personal_auxi1()

'Codigo escrito por Manuel Vizcarra - wwww.combito.com
Dim data(0 To 1, 0 To 12) As String '(columna,fila)

patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " Certificado_Laboral_Activos - rodamiento.dotx"
Set objword = CreateObject("Word.Application")
objword.Visible = True
objword.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0

data(0, 0) = "[employee_name]"
data(1, 0) = Auto_Docs.Cells(2, Auto_Docs.[EMP_NAME].Column) '(fila,columna)
data(0, 1) = "[employee_id]"
data(1, 1) = Auto_Docs.Cells(2, Auto_Docs.[EMP_ID].Column)
data(0, 2) = "[day_word]"
data(1, 2) = Auto_Docs.Cells(2, Auto_Docs.[inc_day_word].Column)
data(0, 3) = "[day]"
data(1, 3) = Auto_Docs.Cells(2, Auto_Docs.[inc_day].Column)
data(0, 4) = "[inc_month]"
data(1, 4) = Auto_Docs.Cells(2, Auto_Docs.[inc_month].Column)
data(0, 5) = "[inc_dated]"
data(1, 5) = Auto_Docs.Cells(2, Auto_Docs.[inc_dated].Column)
data(0, 6) = "[job_name]"
data(1, 6) = Auto_Docs.Cells(2, Auto_Docs.[EMP_JOBNAME].Column)
data(0, 7) = "[word_wage]"
data(1, 7) = Auto_Docs.Cells(2, Auto_Docs.[word_wage].Column)
data(0, 8) = "[wage]"
data(1, 8) = Auto_Docs.Cells(2, Auto_Docs.[EMP_WAGE].Column)
data(0, 9) = "[word_auxi1]"
data(1, 9) = Auto_Docs.Cells(2, Auto_Docs.[word_auxi1].Column)
data(0, 10) = "[auxi1]"
data(1, 10) = Auto_Docs.Cells(2, Auto_Docs.[EMP_AUXI1].Column)
data(0, 11) = "[type_contract]"
data(1, 11) = Auto_Docs.Cells(2, Auto_Docs.[type_contract].Column)
data(0, 12) = "[exp_dated]"
data(1, 12) = Format(Auto_Docs.Cells(2, Auto_Docs.[DATED_REGISTER].Column), "dd"" de ""mmmm"" de ""YYYY")


For i = 0 To UBound(data, 2)

    textobuscar = data(0, i)
    objword.Selection.Move 6, -1
    objword.Selection.Find.Execute FindText:=textobuscar

    While objword.Selection.Find.found = True
        objword.Selection.Text = data(1, i) 'texto a reemplazar
        objword.Selection.Move 6, -1
        objword.Selection.Find.Execute FindText:=textobuscar
    Wend

Next i

objword.Activate

End Sub
Sub Word_contract_asistencial2()

'Codigo escrito por Manuel Vizcarra - wwww.combito.com
Dim data(0 To 1, 0 To 12) As String '(columna,fila)

patharch = ThisWorkbook.Path & "\Templates" & "\IMEXHS Contratos_Medicos_Ginecologos.dotx"
Set objword = CreateObject("Word.Application")
objword.Visible = True
objword.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0

data(0, 0) = "[employee_name]"
data(1, 0) = Auto_Docs.Cells(2, Auto_Docs.[EMP_NAME].Column) '(fila,columna)
data(0, 1) = "[employee_id]"
data(1, 1) = Auto_Docs.Cells(2, Auto_Docs.[EMP_ID].Column)
data(0, 2) = "[day_word]"
data(1, 2) = Auto_Docs.Cells(2, Auto_Docs.[inc_day_word].Column)
data(0, 3) = "[day]"
data(1, 3) = Auto_Docs.Cells(2, Auto_Docs.[inc_day].Column)
data(0, 4) = "[inc_month]"
data(1, 4) = Auto_Docs.Cells(2, Auto_Docs.[inc_month].Column)
data(0, 5) = "[inc_dated]"
data(1, 5) = Auto_Docs.Cells(2, Auto_Docs.[inc_dated].Column)
data(0, 6) = "[job_name]"
data(1, 6) = Auto_Docs.Cells(2, Auto_Docs.[EMP_JOBNAME].Column)
data(0, 7) = "[word_wage]"
data(1, 7) = Auto_Docs.Cells(2, Auto_Docs.[word_wage].Column)
data(0, 8) = "[wage]"
data(1, 8) = Auto_Docs.Cells(2, Auto_Docs.[EMP_WAGE].Column)
data(0, 9) = "[word_auxi1]"
data(1, 9) = Auto_Docs.Cells(2, Auto_Docs.[word_auxi1].Column)
data(0, 10) = "[auxi1]"
data(1, 10) = Auto_Docs.Cells(2, Auto_Docs.[EMP_AUXI1].Column)
data(0, 11) = "[type_contract]"
data(1, 11) = Auto_Docs.Cells(2, Auto_Docs.[type_contract].Column)
data(0, 12) = "[exp_dated]"
data(1, 12) = Format(Auto_Docs.Cells(2, Auto_Docs.[DATED_REGISTER].Column), "dd"" de ""mmmm"" de ""YYYY")


For i = 0 To UBound(data, 2)

    textobuscar = data(0, i)
    objword.Selection.Move 6, -1
    objword.Selection.Find.Execute FindText:=textobuscar

    While objword.Selection.Find.found = True
        objword.Selection.Text = data(1, i) 'texto a reemplazar
        objword.Selection.Move 6, -1
        objword.Selection.Find.Execute FindText:=textobuscar
    Wend

Next i

objword.Activate

End Sub

Sub Word_contract_asistencial1()

'Codigo escrito por Manuel Vizcarra - wwww.combito.com
Dim data(0 To 1, 0 To 12) As String '(columna,fila)

patharch = ThisWorkbook.Path & "\Templates" & "\IMEXHS Contratos_Medicos_Radiólogos.dotx"
Set objword = CreateObject("Word.Application")
objword.Visible = True
objword.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0

data(0, 0) = "[employee_name]"
data(1, 0) = Auto_Docs.Cells(2, Auto_Docs.[EMP_NAME].Column) '(fila,columna)
data(0, 1) = "[employee_id]"
data(1, 1) = Auto_Docs.Cells(2, Auto_Docs.[EMP_ID].Column)
data(0, 2) = "[day_word]"
data(1, 2) = Auto_Docs.Cells(2, Auto_Docs.[inc_day_word].Column)
data(0, 3) = "[day]"
data(1, 3) = Auto_Docs.Cells(2, Auto_Docs.[inc_day].Column)
data(0, 4) = "[inc_month]"
data(1, 4) = Auto_Docs.Cells(2, Auto_Docs.[inc_month].Column)
data(0, 5) = "[inc_dated]"
data(1, 5) = Auto_Docs.Cells(2, Auto_Docs.[inc_dated].Column)
data(0, 6) = "[job_name]"
data(1, 6) = Auto_Docs.Cells(2, Auto_Docs.[EMP_JOBNAME].Column)
data(0, 7) = "[word_wage]"
data(1, 7) = Auto_Docs.Cells(2, Auto_Docs.[word_wage].Column)
data(0, 8) = "[wage]"
data(1, 8) = Auto_Docs.Cells(2, Auto_Docs.[EMP_WAGE].Column)
data(0, 9) = "[word_auxi1]"
data(1, 9) = Auto_Docs.Cells(2, Auto_Docs.[word_auxi1].Column)
data(0, 10) = "[auxi1]"
data(1, 10) = Auto_Docs.Cells(2, Auto_Docs.[EMP_AUXI1].Column)
data(0, 11) = "[type_contract]"
data(1, 11) = Auto_Docs.Cells(2, Auto_Docs.[type_contract].Column)
data(0, 12) = "[exp_dated]"
data(1, 12) = Format(Auto_Docs.Cells(2, Auto_Docs.[DATED_REGISTER].Column), "dd"" de ""mmmm"" de ""YYYY")


For i = 0 To UBound(data, 2)

    textobuscar = data(0, i)
    objword.Selection.Move 6, -1
    objword.Selection.Find.Execute FindText:=textobuscar

    While objword.Selection.Find.found = True
        objword.Selection.Text = data(1, i) 'texto a reemplazar
        objword.Selection.Move 6, -1
        objword.Selection.Find.Execute FindText:=textobuscar
    Wend

Next i

objword.Activate

End Sub

Sub Word_contract_transcriptor()

'Codigo escrito por Manuel Vizcarra - wwww.combito.com
Dim data(0 To 1, 0 To 12) As String '(columna,fila)

patharch = ThisWorkbook.Path & "\Templates" & "\IMEXHS Contrato Transcriptora.dotx"
Set objword = CreateObject("Word.Application")
objword.Visible = True
objword.documents.Add Template:=patharch, NewTemplate:=False, DocumentType:=0

data(0, 0) = "[employee_name]"
data(1, 0) = Auto_Docs.Cells(2, Auto_Docs.[EMP_NAME].Column) '(fila,columna)
data(0, 1) = "[employee_id]"
data(1, 1) = Auto_Docs.Cells(2, Auto_Docs.[EMP_ID].Column)
data(0, 2) = "[day_word]"
data(1, 2) = Auto_Docs.Cells(2, Auto_Docs.[inc_day_word].Column)
data(0, 3) = "[day]"
data(1, 3) = Auto_Docs.Cells(2, Auto_Docs.[inc_day].Column)
data(0, 4) = "[inc_month]"
data(1, 4) = Auto_Docs.Cells(2, Auto_Docs.[inc_month].Column)
data(0, 5) = "[inc_dated]"
data(1, 5) = Auto_Docs.Cells(2, Auto_Docs.[inc_dated].Column)
data(0, 6) = "[job_name]"
data(1, 6) = Auto_Docs.Cells(2, Auto_Docs.[EMP_JOBNAME].Column)
data(0, 7) = "[word_wage]"
data(1, 7) = Auto_Docs.Cells(2, Auto_Docs.[word_wage].Column)
data(0, 8) = "[wage]"
data(1, 8) = Auto_Docs.Cells(2, Auto_Docs.[EMP_WAGE].Column)
data(0, 9) = "[word_auxi1]"
data(1, 9) = Auto_Docs.Cells(2, Auto_Docs.[word_auxi1].Column)
data(0, 10) = "[auxi1]"
data(1, 10) = Auto_Docs.Cells(2, Auto_Docs.[EMP_AUXI1].Column)
data(0, 11) = "[type_contract]"
data(1, 11) = Auto_Docs.Cells(2, Auto_Docs.[type_contract].Column)
data(0, 12) = "[exp_dated]"
data(1, 12) = Format(Auto_Docs.Cells(2, Auto_Docs.[DATED_REGISTER].Column), "dd"" de ""mmmm"" de ""YYYY")


For i = 0 To UBound(data, 2)

    textobuscar = data(0, i)
    objword.Selection.Move 6, -1
    objword.Selection.Find.Execute FindText:=textobuscar

    While objword.Selection.Find.found = True
        objword.Selection.Text = data(1, i) 'texto a reemplazar
        objword.Selection.Move 6, -1
        objword.Selection.Find.Execute FindText:=textobuscar
    Wend

Next i

objword.Activate

End Sub

