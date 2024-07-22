VERSION 5.00
Begin VB.Form MDIChildWindow 
   Caption         =   "MDI Child Window."
   ClientHeight    =   2325
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3630
   MDIChild        =   -1  'True
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   Begin VB.CommandButton CloseMeButton 
      Caption         =   "&Close me."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1452
   End
End
Attribute VB_Name = "MDIChildWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's MDI window's main window.
Option Explicit

'This procedure closes this window.
Private Sub CloseMeButton_Click()
On Error GoTo ErrorTrap

   Unload Me

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub


'This procedure retrieves and displays the unload mode when this window is closed.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrorTrap

   MsgBox GetUnloadModeDescription(UnloadMode), vbInformation
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

