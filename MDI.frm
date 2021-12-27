VERSION 5.00
Begin VB.MDIForm MDIWindow 
   BackColor       =   &H8000000C&
   Caption         =   "QueryUnload event example."
   ClientHeight    =   6540
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13005
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MDIWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the MDI window.
Option Explicit

Private Sub MDIForm_Load()
WindowWindow.Show
End Sub


'This procedure retrieves and displays the unload mode when this window is closed.
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   MsgBox GetUnloadModeDescription(UnloadMode), vbInformation
End Sub


