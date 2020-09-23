VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Preview"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pictOutput 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   90
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   90
      Width           =   1335
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Call Me.pictOutput.Move(Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight)
End Sub

Private Sub pictOutput_DblClick()
    Me.Width = Me.Height
End Sub

