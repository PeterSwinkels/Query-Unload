Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'This procedure returns the description for the specified unload mode.
Public Function GetUnloadModeDescription(UnloadMode As Integer) As String
Dim Description As String

   Description = vbNullString
   
   Select Case UnloadMode
      Case QueryUnloadConstants.vbFormControlMenu
         Description = "Title bar close button or system menu."
      Case QueryUnloadConstants.vbFormCode
         Description = "The ""Unload"" statement has been executed."
      Case QueryUnloadConstants.vbAppWindows
         Description = "This Windows session is closing."
      Case QueryUnloadConstants.vbAppTaskManager
         Description = "The Task Manager is closing this program."
      Case QueryUnloadConstants.vbFormMDIForm
         Description = "The parent MDI form is closing."
      Case QueryUnloadConstants.vbFormOwner
         Description = "This form's owner is closing this form."
      Case Else
         Description = "Unknow shutdown mode #" & CStr(UnloadMode) & "."
   End Select
   
   GetUnloadModeDescription = Description
End Function


