VERSION 5.00
Begin VB.MDIForm MDIWindow 
   BackColor       =   &H8000000C&
   ClientHeight    =   6540
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13008
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MDIWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the MDI window.
Option Explicit

'This procedure initializes this window.
Private Sub MDIForm_Load()
On Error GoTo ErrorTrap

   Me.Caption = ProgramInformation()
   MDIChildWindow.Show

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

'This procedure displays the unload mode used to close this window.
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrorTrap

   MsgBox GetUnloadModeDescription(UnloadMode), vbInformation

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub


