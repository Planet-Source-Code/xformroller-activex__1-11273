VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000007&
      Height          =   2175
      Left            =   0
      Picture         =   "tstMDIForm.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   7245
      TabIndex        =   0
      Top             =   2475
      Width           =   7305
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Sub MDIForm_Load()

    'frmToolBox1.Show
    'frmToolBox1.Move ScaleWidth - frmToolBox1.Width, 0
    
    'frmToolbox2.Show
    'frmToolbox2.Move ScaleWidth - frmToolbox2.Width, frmToolBox1.Height + 30
    
    frmView.Show
    
End Sub

Private Sub Picture1_Click()
    Call ShellExecute(GetDesktopWindow(), vbNullString, "http://vtech.ifrance.com/vtech/devzone", vbNullString, vbNullString, vbNormalFocus)
End Sub
