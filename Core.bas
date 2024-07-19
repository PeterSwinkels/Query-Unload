Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'This procedure returns the description for the specified unload mode.
Public Function GetUnloadModeDescription(UnloadMode As Integer) As String
On Error GoTo ErrorTrap
Dim Description As String

   Description = vbNullString
   
   Select Case UnloadMode
      Case QueryUnloadConstants.vbAppTaskManager
         Description = "Task Manager."
      Case QueryUnloadConstants.vbAppWindows
         Description = "Windows is closing."
      Case QueryUnloadConstants.vbFormCode
         Description = "Unload statement."
      Case QueryUnloadConstants.vbFormControlMenu
         Description = "Form control menu."
      Case QueryUnloadConstants.vbFormMDIForm
         Description = "Parent form is closing."
      Case QueryUnloadConstants.vbFormOwner
         Description = "This form's owner is closing this form."
      Case Else
         Description = "Unknow shutdown mode #" & CStr(UnloadMode) & "."
   End Select
   
EndProcedure:
   GetUnloadModeDescription = Description
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure handles any errors that occur.
Public Function HandleError(Optional ReturnPreviousChoice As Boolean = False) As Long
Dim Description As String
Dim ErrorCode As Long
Static Choice As Long

   Description = Err.Description
   ErrorCode = Err.Number
   On Error Resume Next
   If Not ReturnPreviousChoice Then
      Choice = MsgBox(Description & "." & vbCr & "Error code: " & CStr(ErrorCode), vbAbortRetryIgnore Or vbDefaultButton2 Or vbExclamation)
   End If
   
   If Choice = vbAbort Then End
   
   HandleError = Choice
End Function

'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
   
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path

   MDIWindow.Show

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

'This procedure returns information about this program.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String

   With App
      Information = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With

EndProcedure:
   ProgramInformation = Information
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function
