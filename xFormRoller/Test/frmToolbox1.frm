VERSION 5.00
Object = "*\A..\..\..\..\..\MYDEVE~1\VB\COMPON~1\XFORMR~1\Component\ExtendedFormRoller.vbp"
Begin VB.Form frmToolBox1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Folders"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   121
   ShowInTaskbar   =   0   'False
   Begin ExtendedFormRoller.xFormRoller xFormRoller1 
      Left            =   480
      Top             =   960
      _ExtentX        =   1111
      _ExtentY        =   1111
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmToolBox1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
