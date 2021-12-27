VERSION 5.00
Begin VB.Form WindowWindow 
   Caption         =   "A window."
   ClientHeight    =   2325
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3630
   MDIChild        =   -1  'True
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   Begin VB.CommandButton CloseMeButton 
      Caption         =   "&Close Me"
      Default         =   -1  'True
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1452
   End
End
Attribute VB_Name = "WindowWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's MDI window's main window.
Option Explicit

'This procedure closes this window.
Private Sub CloseMeButton_Click()
   Unload Me
End Sub


'This procedure retrieves and displays the unload mode when this window is closed.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   MsgBox GetUnloadModeDescription(UnloadMode), vbInformation
End Sub

