Attribute VB_Name = "Mod_PersonalizeMsgBox"

' (C) Dan Elgaard   (www.EXCELGAARD.dk)
' Visit vebsite for full explanation and documentation


' MsgBox Buttons/Answers ID Constants
  Private Const MsgBox_Button_ID_OK     As Long = 1
  Private Const MsgBox_Button_ID_Cancel As Long = 2
  Private Const MsgBox_Button_ID_Abort  As Long = 3
  Private Const MsgBox_Button_ID_Retry  As Long = 4
  Private Const MsgBox_Button_ID_Ignore As Long = 5
  Private Const MsgBox_Button_ID_Yes    As Long = 6
  Private Const MsgBox_Button_ID_No     As Long = 7

' MsgBox Buttons/Answers Text Variables
  Private MsgBox_Button_Text_OK         As String
  Private MsgBox_Button_Text_Cancel     As String
  Private MsgBox_Button_Text_Abort      As String
  Private MsgBox_Button_Text_Retry      As String
  Private MsgBox_Button_Text_Ignore     As String
  Private MsgBox_Button_Text_Yes        As String
  Private MsgBox_Button_Text_No         As String

' Handle to the Hook procedure
 #If VBA7 Then
      Private MsgBoxHookHandle          As LongPtr                                        ' 64-bit handle
 #Else
      Private MsgBoxHookHandle          As Long
 #End If

' Windows API functions
 #If VBA7 Then
      Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As LongPtr
      Private Declare PtrSafe Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As LongPtr, ByVal lpString As String) As LongPtr
      Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As LongPtr, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As LongPtr) As LongPtr
      Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
 #Else
      Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
      Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
      Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
      Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
 #End If
 
Option Private Module   ' To prevent the function(s) from appearing the worksheet list of functions (it's a 'for macros only' function)
Option Explicit
Function MsgBoxCB(MsgBox_Text As String, Button1 As String, Optional Button2 As String, Optional Button3 As String, Optional MsgBox_Icon As Long, Optional MsgBox_Title As String) As Long

' * ' Initialize
      On Error Resume Next


' * ' Define variables
      If Button1 = vbNullString Then
            Button1 = Button2
            Button2 = vbNullString
      End If
      If Button2 = vbNullString Then
            Button2 = Button3
            Button3 = vbNullString
      End If

      Dim ButtonsToUse As Long
      ButtonsToUse = vbAbortRetryIgnore
      If Button3 = vbNullString Then ButtonsToUse = vbYesNo
      If Button2 = vbNullString Then ButtonsToUse = vbOKOnly

      Select Case MsgBox_Icon
            Case vbCritical:        ButtonsToUse = ButtonsToUse + MsgBox_Icon
            Case vbExclamation:     ButtonsToUse = ButtonsToUse + MsgBox_Icon
            Case vbInformation:     ButtonsToUse = ButtonsToUse + MsgBox_Icon
            Case vbQuestion:        ButtonsToUse = ButtonsToUse + MsgBox_Icon
      End Select

      If MsgBox_Title = vbNullString Then MsgBox_Title = " Microsoft Excel"               ' Default MsgBox title

      Dim MsgBoxAnswer As Long


' * ' Set custom buttons texts
      MsgBox_Button_Text_OK = Button1
      MsgBox_Button_Text_Cancel = vbNullString                                           ' Not used
      MsgBox_Button_Text_Abort = Button1
      MsgBox_Button_Text_Retry = Button2
      MsgBox_Button_Text_Ignore = Button3
      MsgBox_Button_Text_Yes = Button1
      MsgBox_Button_Text_No = Button2

      MsgBoxHookHandle = SetWindowsHookEx(5, AddressOf MsgBoxHook, 0, GetCurrentThreadId) ' Set MsgBox Hook


' * ' Show hooked MsgBox
      MsgBoxAnswer = MsgBox(MsgBox_Text, ButtonsToUse, MsgBox_Title)


EF: ' End of Function
      UnhookWindowsHookEx MsgBoxHookHandle                                                ' Unhook MsgBox again

      Select Case MsgBoxAnswer
            Case vbOK:        MsgBoxCB = 1
            Case vbCancel:    MsgBoxCB = 0                                                ' Not used
            Case vbAbort:     MsgBoxCB = 1
            Case vbRetry:     MsgBoxCB = 2
            Case vbIgnore:    MsgBoxCB = 3
            Case vbYes:       MsgBoxCB = 1
            Case vbNo:        MsgBoxCB = 2
      End Select

End Function
 
' Below, same sub-function for the MsgBoxHook, depending on weather or not Excel is 32 or 64 bit version
 #If VBA7 Then
      Private Function MsgBoxHook(ByVal LM As LongPtr, ByVal WP As LongPtr, ByVal LP As LongPtr) As LongPtr
            SetDlgItemText WP, MsgBox_Button_ID_OK, MsgBox_Button_Text_OK
            SetDlgItemText WP, MsgBox_Button_ID_Cancel, MsgBox_Button_Text_Cancel         ' Not used
            SetDlgItemText WP, MsgBox_Button_ID_Abort, MsgBox_Button_Text_Abort
            SetDlgItemText WP, MsgBox_Button_ID_Retry, MsgBox_Button_Text_Retry
            SetDlgItemText WP, MsgBox_Button_ID_Ignore, MsgBox_Button_Text_Ignore
            SetDlgItemText WP, MsgBox_Button_ID_Yes, MsgBox_Button_Text_Yes
            SetDlgItemText WP, MsgBox_Button_ID_No, MsgBox_Button_Text_No
      End Function
 #Else
      Private Function MsgBoxHook(ByVal LM As Long, ByVal WP As Long, ByVal LP As Long) As Long
            SetDlgItemText WP, MsgBox_Button_ID_OK, MsgBox_Button_Text_OK
            SetDlgItemText WP, MsgBox_Button_ID_Cancel, MsgBox_Button_Text_Cancel         ' Not used
            SetDlgItemText WP, MsgBox_Button_ID_Abort, MsgBox_Button_Text_Abort
            SetDlgItemText WP, MsgBox_Button_ID_Retry, MsgBox_Button_Text_Retry
            SetDlgItemText WP, MsgBox_Button_ID_Ignore, MsgBox_Button_Text_Ignore
            SetDlgItemText WP, MsgBox_Button_ID_Yes, MsgBox_Button_Text_Yes
            SetDlgItemText WP, MsgBox_Button_ID_No, MsgBox_Button_Text_No
      End Function
 #End If

Sub MsgBoxCB_Test(Optional AskFor As Long = 0)

' AskFor  =  0  =  Website
' AskFor  =  1  =  Month in 1. quarter
' AskFor  =  2  =  Month in 2. quarter
' AskFor  =  3  =  Month in 3. quarter
' AskFor  =  4  =  Month in 4. quarter
' AskFor  =  5  =  Car brand


' * ' Initialize
      On Error Resume Next


' * ' Define variable
      Dim MsgBoxAnswer As Long

      If AskFor < 1 Or AskFor > 5 Then AskFor = 0


' 0 ' Ask for website
      If AskFor < 1 Or AskFor > 5 Then
            MsgBoxAnswer = MsgBoxCB("What's the world's best Excel website?", "EXCELGAARD", "Another", "Don't know", vbQuestion)                ' Show MsgBox with custom buttons
            Select Case MsgBoxAnswer
                  Case 1:  MsgBox "You're so right :-)", vbOKOnly + vbInformation                                                               ' Here we use normal MsgBox
                  Case 2:  MsgBox "You're so wrong!", vbOKOnly + vbExclamation                                                                  ' Here we use normal MsgBox
                  Case 3:  MsgBox "It's www.EXCELGAARD.dk", vbOKOnly + vbInformation                                                            ' Here we use normal MsgBox
            End Select
      End If


' 5 ' Ask for car brand
      If AskFor = 5 Then
            MsgBoxAnswer = MsgBoxCB("What car are you driving?", "Chrysler", "EXCELGAARD", "Mazda", vbQuestion)                                 ' Show MsgBox with custom buttons
            Select Case MsgBoxAnswer
                  Case 1:  MsgBox "I think that's an American car...", vbOKOnly + vbInformation                                                 ' Here we use normal MsgBox
                  Case 2:  MsgBox "That's not a car brand!", vbOKOnly + vbExclamation                                                           ' Here we use normal MsgBox
                  Case 3:  MsgBox "Good old Japanese reliability and speed :-)", vbOKOnly + vbInformation                                       ' Here we use normal MsgBox
            End Select
      End If


' * ' Ask for a month
      If AskFor >= 1 And AskFor <= 4 Then
            Select Case AskFor
                  Case 1:  MsgBoxAnswer = MsgBoxCB("Select a month in the 1st quarter...", "January", "February", "March", vbQuestion) + 0      ' Show MsgBox with custom buttons
                  Case 2:  MsgBoxAnswer = MsgBoxCB("Select a month in the 2nd quarter...", "April", "May", "June", vbQuestion) + 3              ' Show MsgBox with custom buttons
                  Case 3:  MsgBoxAnswer = MsgBoxCB("Select a month in the 3rd quarter...", "July", "August", "September", vbQuestion) + 6       ' Show MsgBox with custom buttons
                  Case 4:  MsgBoxAnswer = MsgBoxCB("Select a month in the 4th quarter...", "October", "November", "December", vbQuestion) + 9   ' Show MsgBox with custom buttons
            End Select

            Select Case MsgBoxAnswer
                  Case 1:  MsgBox "Both 1st month of the year and 1st day of the 1st quarter.", vbOKOnly + vbInformation, " January"
                  Case 2:  MsgBox "There are " & Day(DateSerial(Year(Date), 3, 0)) & " days in February this year (" & Year(Date) & ").", vbOKOnly + vbInformation, " February"
                  Case 3:  MsgBox "Spring is finally here :-)", vbOKOnly + vbInformation, " March"
                  Case 4:  MsgBox "Don't be fooled by any Aprils Fool trick :-)", vbOKOnly + vbInformation, " April"
                  Case 5:  MsgBox "May the force (or the fourth?) be with you :-)", vbOKOnly + vbInformation, " May"
                  Case 6:  MsgBox "www.EXCELGAARD.dk was founded on 15 June 2004 :-)", vbOKOnly + vbInformation, " June"
                  Case 7:  MsgBox "Independence Day occurs this month - not just in the US, but in many countries all around the world actually...", vbOKOnly + vbInformation, " July"
                  Case 8:  MsgBox "Last month of Summer :-(", vbOKOnly + vbInformation, " August"
                  Case 9:  MsgBox "NFL Season is upon us :-)", vbOKOnly + vbInformation, " September"
                  Case 10: MsgBox "Dan Elgaard has a birthday on 27 October :-)", vbOKOnly + vbInformation, " October"
                  Case 11: MsgBox "Or is it 'Movember'?  :-)", vbOKOnly + vbInformation, " November"
                  Case 12: MsgBox "Both last month of the year and last month of last quarter!", vbOKOnly + vbInformation, " December"
            End Select
      End If

End Sub

