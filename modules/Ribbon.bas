Attribute VB_Name = "Ribbon"
Option Explicit
Option Base 1
Dim CANCEL As Boolean
Public Cinta As IRibbonUI
Public retVal(18) As Boolean
#If VBA7 And Win64 Then
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If
Sub CargarCinta(CintaDeExcel As IRibbonUI)
Set Cinta = CintaDeExcel
End Sub
Sub GUARDAR(control As IRibbonControl)
Application.EnableCancelKey = xlDisabled
ThisWorkbook.Save
Application.Calculation = xlCalculationAutomatic
End Sub
Sub SALIR(control As IRibbonControl)
ThisWorkbook.Save
End Sub
Sub EMPLOYEESDATA(control As IRibbonControl)
'Sheets("BG").Select
Call FGPersonal.Show(vbModeless)
End Sub
Sub DOCUMENTATIONDATA(control As IRibbonControl)
'Sheets("BG").Select
SDocumentacion.Show (vbModeless)
End Sub
Sub ABSENTEEISM(control As IRibbonControl)
'Sheets("BG").Select
RAusentismos.Show vbModeless
End Sub
Sub VACATIONS(control As IRibbonControl)
'Sheets("BG").Select
VacationsI.Show vbModeless
End Sub
Sub REPORTS(control As IRibbonControl)
'Sheets("BG").Select
ReportsI.Show
End Sub
Sub ILLREQ(control As IRibbonControl)
'Sheets("BG").Select
IllTracking.Show
End Sub
Sub MONEYRET(control As IRibbonControl)
'Sheets("BG").Select
DevTracking.Show
End Sub
Sub G_PLAT(control As IRibbonControl)
'Sheets("BG").Select
F_PlatF.Show
End Sub
Sub IRPRINCIPAL(control As IRibbonControl)
Sheets("PPrincipal").Select
End Sub

