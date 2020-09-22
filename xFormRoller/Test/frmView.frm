VERSION 5.00
Object = "*\A..\..\..\..\..\..\..\..\MYDEVE~1\WEB\VTECHS~3\DEVZONE\CONTROLS\CONTROLS\XFORMR~1\Component\ExtendedFormRoller.vbp"
Begin VB.Form frmView 
   Caption         =   "View"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   Begin ExtendedFormRoller.xFormRoller xFormRoller1 
      Left            =   3240
      Top             =   1200
      _ExtentX        =   1111
      _ExtentY        =   1111
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    Picture1.Move 0, 0, Width, Height
End Sub
