VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportsI 
   Caption         =   "REPORTES"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6300
   OleObjectBlob   =   "ReportsI.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ReportsI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ABS_BReport_Click()
     If ABS_BReport = True Then
        ReportsI.Active_Emp.Enabled = False
        ReportsI.Retired_Emp.Enabled = False
        ReportsI.All_Emp.Enabled = False
        ReportsI.Doc_Report.Enabled = False
        ReportsI.VAC_BReport.Enabled = False
        ReportsI.NewsReport.Enabled = False
        ReportsI.ARL.Enabled = False
        ReportsI.SDCS.Enabled = False
    Else
        ReportsI.Active_Emp.Enabled = True
        ReportsI.Retired_Emp.Enabled = True
        ReportsI.All_Emp.Enabled = True
        ReportsI.Doc_Report.Enabled = True
        ReportsI.VAC_BReport.Enabled = True
        ReportsI.NewsReport.Enabled = True
        ReportsI.ARL.Enabled = True
        ReportsI.SDCS.Enabled = True
    End If
End Sub

Private Sub Active_Emp_Click()
    If Active_Emp = True Then
        ReportsI.Retired_Emp.Enabled = False
        ReportsI.All_Emp.Enabled = False
        ReportsI.Doc_Report.Enabled = False
        ReportsI.ABS_BReport.Enabled = False
        ReportsI.VAC_BReport.Enabled = False
        ReportsI.NewsReport.Enabled = False
        ReportsI.ARL.Enabled = False
        ReportsI.SDCS.Enabled = False
    Else
        ReportsI.Retired_Emp.Enabled = True
        ReportsI.All_Emp.Enabled = True
        ReportsI.Doc_Report.Enabled = True
        ReportsI.ABS_BReport.Enabled = True
        ReportsI.VAC_BReport.Enabled = True
        ReportsI.NewsReport.Enabled = True
        ReportsI.ARL.Enabled = True
        ReportsI.SDCS.Enabled = True
    End If
End Sub

Private Sub Can_Reporte_Click()
Sheets("PPrincipal").Select
Unload Me
End Sub

Private Sub Doc_Report_Click()
    If Doc_Report = True Then
        ReportsI.Active_Emp.Enabled = False
        ReportsI.Retired_Emp.Enabled = False
        ReportsI.All_Emp.Enabled = False
        ReportsI.ABS_BReport.Enabled = False
        ReportsI.VAC_BReport.Enabled = False
        ReportsI.NewsReport.Enabled = False
        ReportsI.ARL.Enabled = False
        ReportsI.SDCS.Enabled = False
    Else
        ReportsI.Active_Emp.Enabled = True
        ReportsI.Retired_Emp.Enabled = True
        ReportsI.All_Emp.Enabled = True
        ReportsI.ABS_BReport.Enabled = True
        ReportsI.VAC_BReport.Enabled = True
        ReportsI.NewsReport.Enabled = True
        ReportsI.ARL.Enabled = True
        ReportsI.SDCS.Enabled = True
    End If
End Sub

Private Sub NewsReport_Click()
    If NewsReport = True Then
        ReportsI.Active_Emp.Enabled = False
        ReportsI.Retired_Emp.Enabled = False
        ReportsI.All_Emp.Enabled = False
        ReportsI.Doc_Report.Enabled = False
        ReportsI.VAC_BReport.Enabled = False
        ReportsI.ABS_BReport.Enabled = False
        ReportsI.ARL.Enabled = False
        ReportsI.SDCS.Enabled = False
    Else
        ReportsI.Active_Emp.Enabled = True
        ReportsI.Retired_Emp.Enabled = True
        ReportsI.All_Emp.Enabled = True
        ReportsI.Doc_Report.Enabled = True
        ReportsI.VAC_BReport.Enabled = True
        ReportsI.ABS_BReport.Enabled = True
        ReportsI.ARL.Enabled = True
        ReportsI.SDCS.Enabled = True
    End If
End Sub

Private Sub Retired_Emp_Click()
    If Retired_Emp = True Then
        ReportsI.Active_Emp.Enabled = False
        ReportsI.All_Emp.Enabled = False
        ReportsI.Doc_Report.Enabled = False
        ReportsI.ABS_BReport.Enabled = False
        ReportsI.VAC_BReport.Enabled = False
        ReportsI.NewsReport.Enabled = False
        ReportsI.ARL.Enabled = False
        ReportsI.SDCS.Enabled = False
    Else
        ReportsI.Active_Emp.Enabled = True
        ReportsI.All_Emp.Enabled = True
        ReportsI.Doc_Report.Enabled = True
        ReportsI.ABS_BReport.Enabled = True
        ReportsI.VAC_BReport.Enabled = True
        ReportsI.NewsReport.Enabled = True
        ReportsI.ARL.Enabled = True
        ReportsI.SDCS.Enabled = True
    End If
End Sub
Private Sub All_Emp_Click()
    If All_Emp = True Then
        ReportsI.Active_Emp.Enabled = False
        ReportsI.Retired_Emp.Enabled = False
        ReportsI.Doc_Report.Enabled = False
        ReportsI.ABS_BReport.Enabled = False
        ReportsI.VAC_BReport.Enabled = False
        ReportsI.NewsReport.Enabled = False
        ReportsI.ARL.Enabled = False
        ReportsI.SDCS.Enabled = False
    Else
        ReportsI.Active_Emp.Enabled = True
        ReportsI.Retired_Emp.Enabled = True
        ReportsI.Doc_Report.Enabled = True
        ReportsI.ABS_BReport.Enabled = True
        ReportsI.VAC_BReport.Enabled = True
        ReportsI.NewsReport.Enabled = True
        ReportsI.ARL.Enabled = True
        ReportsI.SDCS.Enabled = True
    End If
End Sub
Private Sub Create_Report_Emp_Click()
If Active_Emp = True Then
        Call SelectExportOption
    ElseIf Retired_Emp = True Then
        Call Retired_ReportOptions
    ElseIf All_Emp = True Then
        Call AE_SelectExportOption
    ElseIf Doc_Report = True Then
        Call Documentation_ReportOption
    ElseIf ABS_BReport = True Then
        Call ABS_ExportOption
    ElseIf VAC_BReport = True Then
        Call Vac_ReportOptions
    ElseIf NewsReport = True Then
        Call News_ReportOptions
    ElseIf ARL = True Then
        Call ARL_ReportOptions
    ElseIf SDCS = True Then
        Call SocEco_ReportOptions
    ElseIf Ctrs = True Then
        Call ExportContractState
Else
    MsgBox "Seleccione un reporte para crear"
End If
End Sub

Private Sub SDCS_Click()
    If SDCS = True Then
        ReportsI.Active_Emp.Enabled = False
        ReportsI.Retired_Emp.Enabled = False
        ReportsI.All_Emp.Enabled = False
        ReportsI.Doc_Report.Enabled = False
        ReportsI.ABS_BReport.Enabled = False
        ReportsI.VAC_BReport.Enabled = False
        ReportsI.NewsReport.Enabled = False
        ReportsI.ARL.Enabled = False
    Else
        ReportsI.Active_Emp.Enabled = True
        ReportsI.Retired_Emp.Enabled = True
        ReportsI.All_Emp.Enabled = True
        ReportsI.Doc_Report.Enabled = True
        ReportsI.ABS_BReport.Enabled = True
        ReportsI.VAC_BReport.Enabled = True
        ReportsI.NewsReport.Enabled = True
        ReportsI.ARL.Enabled = True
    End If
End Sub
Private Sub ARL_Click()
    If ARL = True Then
        ReportsI.Active_Emp.Enabled = False
        ReportsI.Retired_Emp.Enabled = False
        ReportsI.All_Emp.Enabled = False
        ReportsI.Doc_Report.Enabled = False
        ReportsI.ABS_BReport.Enabled = False
        ReportsI.VAC_BReport.Enabled = False
        ReportsI.NewsReport.Enabled = False
        ReportsI.SDCS.Enabled = False
    Else
        ReportsI.Active_Emp.Enabled = True
        ReportsI.Retired_Emp.Enabled = True
        ReportsI.All_Emp.Enabled = True
        ReportsI.Doc_Report.Enabled = True
        ReportsI.ABS_BReport.Enabled = True
        ReportsI.VAC_BReport.Enabled = True
        ReportsI.NewsReport.Enabled = True
        ReportsI.SDCS.Enabled = True
    End If
End Sub
Private Sub VAC_BReport_Click()
    If VAC_BReport = True Then
        ReportsI.Active_Emp.Enabled = False
        ReportsI.Retired_Emp.Enabled = False
        ReportsI.All_Emp.Enabled = False
        ReportsI.Doc_Report.Enabled = False
        ReportsI.ABS_BReport.Enabled = False
        ReportsI.NewsReport.Enabled = False
        ReportsI.ARL.Enabled = False
        ReportsI.SDCS.Enabled = False
    Else
        ReportsI.Active_Emp.Enabled = True
        ReportsI.Retired_Emp.Enabled = True
        ReportsI.All_Emp.Enabled = True
        ReportsI.Doc_Report.Enabled = True
        ReportsI.ABS_BReport.Enabled = True
        ReportsI.NewsReport.Enabled = True
        ReportsI.ARL.Enabled = True
        ReportsI.SDCS.Enabled = True
    End If
End Sub
