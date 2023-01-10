Attribute VB_Name = "Documents_Automatic"
Public patharch As String
Public Enterprise As Variant
Public AskFor As Long
Public FolderName As String
Public SaveDocs As Long
Public DocumentName As String
Global NewsType As Long
'Global JobFunctions As Variant
Public Sub SelectDocumentType()
Application.Calculation = xlCalculationAutomatic
    'MsgBox JobFunctions.JobFunctions
    strMsg1 = "Select Document"
    strMsg2 = "1" & vbTab & "CL Activos"
    strMsg3 = "2" & vbTab & "CL Retirado"
    strMsg4 = "3" & vbTab & "Retirado"
    strMsg5 = "4" & vbTab & "CT Colsubsidio"
    strMsg6 = "5" & vbTab & "CT Ingenieros"
    strMsg7 = "6" & vbTab & "CT Desarrolladores"
    strMsg8 = "7" & vbTab & "CT Administrativo"
    strMsg9 = "8" & vbTab & "CT RIMAB"
    strMsg10 = "9" & vbTab & "OTRO SÍ"
    strMsg11 = "10" & vbTab & "IME-RIM UT"
    strMsg12 = "11" & vbTab & "Med. Exams"
    
    AskFor = InputBox(strMsg1 & vbCrLf & vbCrLf & strMsg2 & vbCrLf & strMsg3 & _
        vbCrLf & strMsg4 & vbCrLf & strMsg5 & vbCrLf & strMsg6 & vbCrLf & strMsg7 & _
        vbCrLf & strMsg8 & vbCrLf & strMsg9 & vbCrLf & strMsg10 & vbCrLf & strMsg11 & vbCrLf & strMsg12, strDefault, 1, 1)
    Call MsgBoxCB_Test
End Sub
Sub MsgBoxCB_Test()
Application.Calculation = xlCalculationAutomatic
' * ' Initialize
      On Error Resume Next

' * ' Define variable
      Dim MsgBoxAnswer As Long

' * ' Ask for a month
      If AskFor >= 1 And AskFor <= 11 Then
            Select Case AskFor
                  Case 1:  MsgBoxAnswer = MsgBoxCB("Seleccione Tipo de Certificado Laboral", "CL Activo", "CL Activo Auxilio", "CL Activo Servicios", vbQuestion) + 0      ' Show MsgBox with custom buttons
                  Case 2:  MsgBoxAnswer = MsgBoxCB("Seleccione Tipo de Certificado Laboral", "CL Retirado", "CL Retirado Auxilio", "empty", vbQuestion) + 3              ' Show MsgBox with custom buttons
                  Case 3:  MsgBoxAnswer = MsgBoxCB("Retirado", "Retirados", "CT Ingenieros", "September", vbQuestion) + 6       ' Show MsgBox with custom buttons
                  Case 4:  MsgBoxAnswer = MsgBoxCB("Contratos Proyecto Colsubsidio", "Transcriptora", "Radiologo", "Ginecologo", vbQuestion) + 9   ' Show MsgBox with custom buttons
                  'Create Engineer Contracts
                  Case 5:  MsgBoxAnswer = MsgBoxCB("Contratos Ingenieros", "CT Ing.", "CT Ing. Auxilio", "Empty", vbQuestion) + 12   ' Show MsgBox with custom buttons
                  Case 6:  MsgBoxAnswer = MsgBoxCB("Contratos Desarrolladores", "CT Dev.", "CT Dev. Auxilio", "Empty", vbQuestion) + 15   ' Show MsgBox with custom buttons
                  Case 7:  MsgBoxAnswer = MsgBoxCB("Contratos Administrativo", "CT Administrativo", "CT Administratiov Auxilio", "Empty", vbQuestion) + 18   ' Show MsgBox with custom buttons
                  Case 8:  MsgBoxAnswer = MsgBoxCB("Contratos RIMAB", "CT Indefinido", "CT Indefinido Auxilio", "Empty", vbQuestion) + 21  ' Show MsgBox with custom buttons
                  Case 9:  MsgBoxAnswer = MsgBoxCB("Otros Sí", "Otro Sí", "Empty", "Empty", vbQuestion) + 24  ' Show MsgBox with custom buttons
                  Case 10: MsgBoxAnswer = MsgBoxCB("IME RIM UT", "CT Admon", "Empty", "Empty", vbQuestion) + 27  ' Show MsgBox with custom buttons
                  Case 11: MsgBoxAnswer = MsgBoxCB("AE", "Authorization", "Empty", "Empty", vbQuestion) + 30  ' Show MsgBox with custom buttons
            End Select

            Select Case MsgBoxAnswer
                  'Create Work Certification
                  Case 1:  Call ChooseEnterprise
                            DocumentName = "CL " & Auto_Docs.Range("C2")
                             SaveDocs = 3
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CL Activo.dotx"
                             Call TemplateGenerator
                  'Create Work Certification Auxi
                  Case 2:  Call ChooseEnterprise
                            DocumentName = "CL " & Auto_Docs.Range("C2")
                             SaveDocs = 3
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CL Activo Auxilio.dotx"
                             Call TemplateGenerator
                  'CL Activo Asistencial
                  Case 3: Call ChooseEnterprise
                            DocumentName = "CL " & Auto_Docs.Range("C2")
                             SaveDocs = 3
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CL Activo Servicios.dotx"
                             Call TemplateGenerator
                  'Retired
                  Case 4: Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CL Retirado.dotx"
                             Call TemplateGenerator
                  'Retired Auxilio
                  Case 5:  Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CL Retirado Auxilio.dotx"
                             Call TemplateGenerator
                  'empty
                  Case 6:  MsgBox "empty"
                  'Create Retired Documents
                  Case 7:  Call ChooseEnterprise
                            'Assign Folder Name
                             FolderName = Auto_Docs.Range("C2")
                             Call CreateFolderRetirements
                             'Generate Documents
                             'Payment Authorization
                             SaveDocs = 1
                             DocumentName = "Autorización de Pago"
                             patharch = ThisWorkbook.Path & "\Templates" & "\Retired\" & Enterprise & " Autorizacion_Pago.dotx"
                             Call TemplateGenerator
                             'Retirement Medical Exam
                             SaveDocs = 1
                             DocumentName = "Examen Egreso"
                             patharch = ThisWorkbook.Path & "\Templates" & "\Retired\" & Enterprise & " Examen Egreso.dotx"
                             Call TemplateGenerator
                             'Work Certificate IMEXHS Certificado_Laboral_Retirado
                             SaveDocs = 1
                             DocumentName = "Certificado Laboral Retirado"
                             patharch = ThisWorkbook.Path & "\Templates" & "\Retired\" & Enterprise & " Certificado_Laboral_Retirado.dotx"
                             Call TemplateGenerator
                             'Work Certificate IMEXHS Certificado_Laboral_Retirado - rodamiento
                             SaveDocs = 1
                             DocumentName = "Certificado Laboral Retirado - rodamiento"
                             patharch = ThisWorkbook.Path & "\Templates" & "\Retired\" & Enterprise & " Certificado_Laboral_Retirado - rodamiento.dotx"
                             Call TemplateGenerator
                             'Cesantias Authorization IMEXHS Autorizacion_Cesantias
                             SaveDocs = 1
                             DocumentName = "Autorización Cesantías"
                             patharch = ThisWorkbook.Path & "\Templates" & "\Retired\" & Enterprise & " Autorizacion_Cesantias.dotx"
                             Call TemplateGenerator
                 'empty
                  Case 8:   MsgBox "empty"
                  'empty
                  Case 9:   MsgBox "empty"
                  'Create Transcriptor Contract
                  Case 10:  Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Transcriptora.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 2
                             Call TemplateGenerator
                  'Create Radiology Contract
                  Case 11: Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Radiologo.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 2
                             Call TemplateGenerator
                  'Create Gine... Contract
                  Case 12: Call ChooseEnterprise
                             patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Ginecologo.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 2
                             Call TemplateGenerator
                  'Empty
                  Case 13: JobFunctions.Show
                            Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Ingeniero.dotx"
                            DocumentName = Auto_Docs.Range("C2")
                            SaveDocs = 2
                            Call TemplateGenerator
                  Case 14: JobFunctions.Show
                            Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Ingeniero Auxilio.dotx"
                            DocumentName = Auto_Docs.Range("C2")
                            SaveDocs = 2
                            Call TemplateGenerator
                  'Empty
                  Case 15: MsgBox "empty"
                  'Create Developer Contract
                  Case 16:  JobFunctions.Show
                            Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Desarrollador.dotx"
                            DocumentName = Auto_Docs.Range("C2")
                            SaveDocs = 2
                            Call TemplateGenerator
                  'Create Developer Contract Auxilio
                  Case 17:  JobFunctions.Show
                            Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Desarrollador Auxilio.dotx"
                            DocumentName = Auto_Docs.Range("C2")
                            SaveDocs = 2
                            Call TemplateGenerator
                  'Empty
                   Case 18: MsgBox "empty"
                   'Create Administrative Contract
                   Case 19:
                   'MsgBox "CALLING ME"
                            JobFunctions.Show
                            Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Administativo.dotx"
                            DocumentName = Auto_Docs.Range("C2")
                            SaveDocs = 2
                            Call TemplateGenerator
                   'Create Administrative Contract Aux
                   Case 20: JobFunctions.Show
                            Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Indefinido Auxilio.dotx"
                            DocumentName = Auto_Docs.Range("C2")
                            SaveDocs = 2
                            Call TemplateGenerator
                    'Empty
                   Case 21: JobFunctions.Show
                            Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Fijo.dotx"
                            DocumentName = Auto_Docs.Range("C2")
                            SaveDocs = 2
                            Call TemplateGenerator
                   'Create RIMAB Contracts
                   Case 22: JobFunctions.Show
                            Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Fijo.dotx"
                            DocumentName = Auto_Docs.Range("C2")
                            SaveDocs = 2
                            'MsgBox JobFunctions.JobFunctions
                            Call TemplateGenerator
                   'Create Administrative Contract Aux
                   Case 23: JobFunctions.Show
                            Call ChooseEnterprise
                            patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Indefinido Auxilio.dotx"
                            DocumentName = Auto_Docs.Range("C2")
                            SaveDocs = 2
                            Call TemplateGenerator
                    'Empty
                   Case 24:
                             Call ChooseEnterprise
                             patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT RadiologoSeg.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 2
                             Call TemplateGenerator
            '==============OTROS SÍ
                    Case 25: JobFunctions.Show
                             Call ChooseEnterprise
                             patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " OtroSí.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 4
                             Call TemplateGenerator
                    Case 26: MsgBox "empty"
                    Case 27: MsgBox "empty"
                    
            '==============IME RIM UT
                    Case 28: JobFunctions.Show
                             Call ChooseEnterprise
                             patharch = ThisWorkbook.Path & "\Templates\IMERIMUT" & "\" & Enterprise & " CT Admon.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 5
                             Call TemplateGenerator
                    Case 29: Call ChooseEnterprise
                             patharch = ThisWorkbook.Path & "\Templates\IMERIMUT" & "\" & Enterprise & " CL Activo.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 5
                             Call TemplateGenerator
                    Case 30: JobFunctions.Show
                            Call ChooseEnterprise
                             patharch = ThisWorkbook.Path & "\Templates\" & Enterprise & " CT Indefinido Auxilio.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 5
                             Call TemplateGenerator
            '==============examenes
                    Case 31: Call ChooseEnterprise
                             patharch = ThisWorkbook.Path & "\Templates\" & Enterprise & " Autorizacion Examen.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 6
                             Call TemplateGenerator
                    Case 32: Call ChooseEnterprise
                             patharch = ThisWorkbook.Path & "\Templates\" & Enterprise & " ME.dotx"
                             DocumentName = Auto_Docs.Range("C2")
                             SaveDocs = 7
                             Call TemplateGenerator
                    Case 33: MsgBox "empty"
            End Select
      End If

End Sub

Sub CreateFolderRetirements()

MkDir ThisWorkbook.Path & "\Templates\Retired\" & FolderName

End Sub

Public Sub Docs_Creator()
Application.Calculation = xlCalculationAutomatic
If MsgBox("Generar Documento", vbYesNo) = vbYes Then
MsgBoxAnswer = MsgBoxCB("Seleccione el tipo de documento", "CL Activos", "CL Activos Auxilio", "CT Ginecologo")
    If MsgBoxAnswer = 1 Then
        Call ChooseEnterprise
        patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CL Activo.dotx"
        Call TemplateGenerator
    ElseIf MsgBoxAnswer = 2 Then
        Call ChooseEnterprise
        patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CL Activo Auxilio.dotx"
        Call TemplateGenerator
    ElseIf MsgBoxAnswer = 3 Then
        Call ChooseEnterprise
        patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Ginecologo.dotx"
        Call TemplateGenerator
    ElseIf msgboxanwer = 4 Then
        Call ChooseEnterprise
        patharch = ThisWorkbook.Path & "\Templates" & "\" & Enterprise & " CT Radiologo.dotx"
        Call TemplateGenerator
    End If
Else
Exit Sub
End If
End Sub

Public Sub ChooseEnterprise()

MsgBoxAnswer = MsgBoxCB("Seleccione el tipo de documento", "IMEXHS", "RIMAB", "IMERIMUT")
    If MsgBoxAnswer = 1 Then
        Enterprise = "IMEXHS"
    ElseIf MsgBoxAnswer = 2 Then
        Enterprise = "RIMAB"
    ElseIf MsgBoxAnswer = 3 Then
        Enterprise = "IMERIMUT"
    End If

End Sub

Public Sub TemplateGenerator()
Application.Calculation = xlCalculationAutomatic
'Codigo escrito por Manuel Vizcarra - wwww.combito.com
Dim data(0 To 1, 0 To 32) As String '(columna,fila)

Set objword = CreateObject("word.Application")
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
data(0, 11) = "[emp_afp]"
data(1, 11) = Auto_Docs.Cells(2, Auto_Docs.[EMP_AFP].Column)
data(0, 12) = "[emp_dor]"
data(1, 12) = Auto_Docs.Cells(2, Auto_Docs.[EMP_DORE].Column)
data(0, 13) = "[word_emp_dor]"
data(1, 13) = Auto_Docs.Cells(2, Auto_Docs.[word_emp_dor].Column)
data(0, 14) = "[month_retired]"
data(1, 14) = Auto_Docs.Cells(2, Auto_Docs.[month_retired].Column)
data(0, 15) = "[year_retired]"
data(1, 15) = Auto_Docs.Cells(2, Auto_Docs.[year_retired].Column)
data(0, 16) = "[auxi1]"
data(1, 16) = Auto_Docs.Cells(2, Auto_Docs.[EMP_AUXI1].Column)
data(0, 17) = "[word_auxi1]"
data(1, 17) = Auto_Docs.Cells(2, Auto_Docs.[word_auxi1].Column)
data(0, 18) = "[JobFunctions]"
data(1, 18) = JobFunctions.JobFunctions
data(0, 19) = "[inc_year_word]"
data(1, 19) = Auto_Docs.Cells(2, Auto_Docs.[inc_year_word].Column)
data(0, 20) = "[dtuc_day]"
data(1, 20) = Auto_Docs.Cells(2, Auto_Docs.[dtuc_day].Column)
data(0, 21) = "[dtuc_dayw]"
data(1, 21) = Auto_Docs.Cells(2, Auto_Docs.[dtuc_dayw].Column)
data(0, 22) = "[dtuc_dayw]"
data(1, 22) = Auto_Docs.Cells(2, Auto_Docs.[dtuc_dayw].Column)
data(0, 23) = "[dtuc_month]"
data(1, 23) = Auto_Docs.Cells(2, Auto_Docs.[dtuc_month].Column)
data(0, 24) = "[dtuc_year]"
data(1, 24) = Auto_Docs.Cells(2, Auto_Docs.[dtuc_year].Column)
data(0, 25) = "[dtuc_year_]"
data(1, 25) = Auto_Docs.Cells(2, Auto_Docs.[dtuc_year_].Column)

data(0, 26) = "[os_day]"
data(1, 26) = Auto_Docs.Cells(2, Auto_Docs.[os_day].Column)
data(0, 27) = "[os_dayw]"
data(1, 27) = Auto_Docs.Cells(2, Auto_Docs.[os_dayw].Column)
data(0, 28) = "[os_month]"
data(1, 28) = Auto_Docs.Cells(2, Auto_Docs.[os_month].Column)
data(0, 29) = "[os_year]"
data(1, 29) = Auto_Docs.Cells(2, Auto_Docs.[os_year].Column)
data(0, 30) = "[os_year_]"
data(1, 30) = Auto_Docs.Cells(2, Auto_Docs.[os_year_].Column)
data(0, 31) = "[Day_Exam]"
data(1, 31) = Auto_Docs.Cells(2, Auto_Docs.[Day_Exam].Column)
data(0, 32) = "[Hour_Exam]"
data(1, 32) = Auto_Docs.Cells(2, Auto_Docs.[Hour_Exam].Column)


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

If SaveDocs = 1 Then
    'MsgBox "called"
    objword.activedocument.SaveAs Filename:=ThisWorkbook.Path & "\templates\Retired\" & FolderName & "\" & DocumentName
    SaveDocs = 0
ElseIf SaveDocs = 2 Then
    'MsgBox "2"
    objword.activedocument.SaveAs Filename:=ThisWorkbook.Path & "\templates\Contracts\" & "CT " & DocumentName
    SaveDocs = 0
ElseIf SaveDocs = 3 Then
    'MsgBox "2"
    'objdoc.ExportAsFixedFormat OutputFileName:=ThisWorkbook.Path & "\templates\Work Certifications\" & DocumentName & ".pdf", ExportFormat:=wdExportFormatPDF
    objword.activedocument.SaveAs Filename:=ThisWorkbook.Path & "\templates\Work Certifications\" & DocumentName
    SaveDocs = 0
ElseIf SaveDocs = 4 Then
    'MsgBox "2"
    'objdoc.ExportAsFixedFormat OutputFileName:=ThisWorkbook.Path & "\templates\Work Certifications\" & DocumentName & ".pdf", ExportFormat:=wdExportFormatPDF
    objword.activedocument.SaveAs Filename:=ThisWorkbook.Path & "\templates\OtroSí\" & "OS" & DocumentName
    SaveDocs = 0
ElseIf SaveDocs = 5 Then
    objword.activedocument.SaveAs Filename:=ThisWorkbook.Path & "\templates\IMERIMUT\Contracts\" & "CT " & DocumentName
    SaveDocs = 0
ElseIf SaveDocs = 6 Then
    objword.activedocument.SaveAs Filename:=ThisWorkbook.Path & "\templates\ME\" & "AE" & DocumentName
    SaveDocs = 0
ElseIf SaveDocs = 7 Then
    objword.activedocument.SaveAs Filename:=ThisWorkbook.Path & "\templates\ME\" & "EX " & DocumentName
    SaveDocs = 0
Else
    objword.Activate
End If

End Sub
