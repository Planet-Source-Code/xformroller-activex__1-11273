VERSION 5.00
Begin VB.UserControl xFormRoller 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   44
   ToolboxBitmap   =   "xFormRoller.ctx":0000
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   0
      Picture         =   "xFormRoller.ctx":0312
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "xFormRoller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

'---------- API declarations ----------

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Const GWL_WNDPROC = (-4)
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const MF_STRING = &H0&
Private Const SM_CXFRAME = 32

' Big caption
Private Const SM_CXSIZE = 30
Private Const SM_CYSIZE = 31
Private Const SM_CYCAPTION = 4

' Small caption
Private Const SM_CYSMCAPTION = 51
Private Const SM_CXSMSIZE = 52
Private Const SM_CYSMSIZE = 53

' Frame thickness
Private Const SM_CXFIXEDFRAME = 7
Private Const SM_CYFIXEDFRAME = 8
Private Const SM_CXSIZEFRAME = 32
Private Const SM_CYSIZEFRAME = 33


Private Const SM_CYDLGFRAME = 8
Private Const TPM_LEFTALIGN = &H0&

' Messages
Private Const WM_ACTIVATE = &H6
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MENUSELECT = &H11F
Private Const WM_NCHITTEST = &H84
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_PAINT = &HF
Private Const WM_NCPAINT = &H85

Private Const WM_SIZE = &H5
Private Const SIZE_MAXHIDE = 4
Private Const SIZE_MAXIMIZED = 2
Private Const SIZE_MAXSHOW = 3
Private Const SIZE_MINIMIZED = 1
Private Const SIZE_RESTORED = 0

Private Const WM_NCLBUTTONDBLCLK = &HA3

' Draw frame constants
Private Const DFC_SCROLL = 3
Private Const DFCS_SCROLLDOWN = &H1
Private Const DFCS_SCROLLUP = &H0

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'---------- Variables ----------

Private CursorInBOX As Boolean
Private CursorLoc As POINTAPI
Private oldWndProc As Long
'Private R As RECT
Private m_IsRolledUp As Boolean

Private m_NormalSize As Long
Private m_MaximizedSize As Long
Private m_IsRolling As Boolean
Private m_IsOnMDI As Boolean
Private m_IsPainting As Boolean
Private m_ButtonRectInitialized As Boolean
Private m_ButtonRect As RECT

Private WithEvents m_Form As Form
Attribute m_Form.VB_VarHelpID = -1

Private counter As Long

'============================================================
'                       Subclassing
'============================================================
Private Sub AttachMessages()

    Call AttachMessage(Me, m_Form.hwnd, WM_SIZE)
    Call AttachMessage(Me, m_Form.hwnd, WM_NCHITTEST)
    Call AttachMessage(Me, m_Form.hwnd, WM_NCLBUTTONDOWN)
    Call AttachMessage(Me, m_Form.hwnd, WM_ACTIVATE)
    Call AttachMessage(Me, m_Form.hwnd, WM_MDIACTIVATE)
    Call AttachMessage(Me, m_Form.hwnd, WM_NCLBUTTONDBLCLK)
    Call AttachMessage(Me, m_Form.hwnd, WM_MENUSELECT)
    Call AttachMessage(Me, m_Form.hwnd, WM_PAINT)
    Call AttachMessage(Me, m_Form.hwnd, WM_NCPAINT)

End Sub

Private Sub m_Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DetachMessages
End Sub

Private Sub DetachMessages()

    Call DetachMessage(Me, m_Form.hwnd, WM_SIZE)
    Call DetachMessage(Me, m_Form.hwnd, WM_NCHITTEST)
    Call DetachMessage(Me, m_Form.hwnd, WM_NCLBUTTONDOWN)
    Call DetachMessage(Me, m_Form.hwnd, WM_ACTIVATE)
    Call DetachMessage(Me, m_Form.hwnd, WM_MDIACTIVATE)
    Call DetachMessage(Me, m_Form.hwnd, WM_NCLBUTTONDBLCLK)
    Call DetachMessage(Me, m_Form.hwnd, WM_MENUSELECT)
    Call DetachMessage(Me, m_Form.hwnd, WM_PAINT)
    Call DetachMessage(Me, m_Form.hwnd, WM_NCPAINT)

End Sub

Private Property Get ISubClass_MsgResponse() As EMsgResponse
    ISubClass_MsgResponse = emrPostProcess
End Property

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim ButRgn As Long
    Dim lHeight As Long
    
    Select Case iMsg
        'Case WM_WININICHANGED
            ' On Windows 95, the text is not centered and the
            ' user can choose the Font. In addition, your
            ' application might want to monitor the
            ' WM_WININICHANGED message, because the user can
            ' change titlebar widths, and so forth,
            ' dynamically. When this happens, the application
            ' should take the new system metrics into account,
            ' and force a window redraw.

        Case WM_NCHITTEST
            '---- Verify if the mouse is under the button
            CopyMemory ByVal VarPtr(CursorLoc.X), ByVal VarPtr(lParam), 2
            CopyMemory ByVal VarPtr(CursorLoc.Y), ByVal VarPtr(lParam) + 2, 2
            
            Dim TT As RECT
            GetWindowRect m_Form.hwnd, TT
            ButRgn = CreateRectRgn(m_ButtonRect.Left, m_ButtonRect.Top, m_ButtonRect.Right, m_ButtonRect.Bottom)
            If PtInRegion(ButRgn, (CursorLoc.X - TT.Left), (CursorLoc.Y - TT.Top)) Then
                CursorInBOX = True
            Else
                CursorInBOX = False
            End If
            DeleteObject ButRgn
            
        Case WM_PAINT, WM_NCPAINT
            RefreshButtonAndMenu
            
        Case WM_SIZE
            
            '---- Save the size for the good window state
            If (Not m_IsRolling And Not m_IsRolledUp) Then
    
                Select Case m_Form.WindowState
                Case vbNormal
                    m_NormalSize = ScaleY(m_Form.Height, m_Form.ScaleMode, vbPixels)
                Case vbMinimized
                Case vbMaximized
                    m_MaximizedSize = ScaleY(m_Form.Height, m_Form.ScaleMode, vbPixels)
                End Select
                
            End If
            
            '---- Restore the information if necessary
            If (wParam = SIZE_RESTORED And m_IsRolledUp) Then
                RollDownForm
            End If
            
            RefreshButtonAndMenu
                        
        Case WM_ACTIVATE, WM_MDIACTIVATE
            RefreshButtonAndMenu
       
        Case WM_NCLBUTTONDBLCLK
            If CursorInBOX = True Then
                Exit Function
            End If
            
        Case WM_MENUSELECT
            Static Clik%
            If Clik% = 1 And wParam = -65536 Then
                If m_IsRolledUp = False Then
                    RollUpForm
                Else
                    RollDownForm
                End If
            End If
            If -1602224128 = wParam Then
                Clik% = 1
            Else
                Clik% = 0
            End If
        
        Case WM_NCLBUTTONDOWN
            If CursorInBOX = True Then
                PaintPressedButton
                If m_IsRolledUp = False Then
                    RollUpForm
                Else
                    RollDownForm
                End If
            End If
            
            RefreshButtonAndMenu
            
    End Select
        
End Function

'============================================================
'                       Properties
'============================================================

Private Sub UserControl_Resize()
   UserControl.Width = 42 * Screen.TwipsPerPixelX
   UserControl.Height = 42 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' Only at run time
    If (Not Ambient.UserMode) Then Exit Sub

    ' Verify if this is a form !!!
    If (Not TypeOf UserControl.Parent Is Form) Then
        Exit Sub
    End If
    
    ' Not yet supported
    If (TypeOf UserControl.Parent Is MDIForm) Then
        Exit Sub
    End If
    
    Set m_Form = UserControl.Parent
    m_IsOnMDI = m_Form.MDIChild
    m_IsRolling = False
    
    AttachMessages

    m_IsRolledUp = False
    RefreshButtonAndMenu
    
End Sub

'============================================================
'                      Button drawing
'============================================================
Private Function GetButtonRect() As RECT

    Dim buttonWidth As Long, buttonHeight As Long, CaptionHeight As Long
    Dim nbStdButtons As Integer, lCaptionTop As Long, lButtonOffset As Long
    
    ' Initial gap on the right
    lButtonOffset = 2
    
    Select Case m_Form.BorderStyle
        Case vbBSNone
        Case vbFixedSingle
            ' Big button
            buttonWidth = GetSystemMetrics(SM_CXSIZE)
            buttonHeight = GetSystemMetrics(SM_CYSIZE)
            CaptionHeight = GetSystemMetrics(SM_CYCAPTION)
            nbStdButtons = 1
            lCaptionTop = GetSystemMetrics(SM_CYFIXEDFRAME)
            lButtonOffset = lButtonOffset + GetSystemMetrics(SM_CXFIXEDFRAME)
            
        Case vbSizable
            ' Big button
            buttonWidth = GetSystemMetrics(SM_CXSIZE)
            buttonHeight = GetSystemMetrics(SM_CYSIZE)
            CaptionHeight = GetSystemMetrics(SM_CYCAPTION)
            lCaptionTop = GetSystemMetrics(SM_CYSIZEFRAME)
            lButtonOffset = lButtonOffset + GetSystemMetrics(SM_CXSIZEFRAME)
        
            If (m_Form.MaxButton Or m_Form.MinButton) Then
                ' The three buttons are present
                nbStdButtons = 3
                lButtonOffset = lButtonOffset + 2
            Else
                ' Only the close button is present
                nbStdButtons = 1
            End If
            
        Case vbFixedDouble ' Fixed dialog
            ' Big button
            buttonWidth = GetSystemMetrics(SM_CXSIZE)
            buttonHeight = GetSystemMetrics(SM_CYSIZE)
            CaptionHeight = GetSystemMetrics(SM_CYCAPTION)
            lCaptionTop = GetSystemMetrics(SM_CYFIXEDFRAME)
            lButtonOffset = lButtonOffset + GetSystemMetrics(SM_CXFIXEDFRAME)
            nbStdButtons = 1
            
        Case vbFixedToolWindow
            ' Small button
            buttonWidth = GetSystemMetrics(SM_CXSMSIZE)
            buttonHeight = GetSystemMetrics(SM_CYSMSIZE)
            CaptionHeight = GetSystemMetrics(SM_CYSMCAPTION)
            lCaptionTop = GetSystemMetrics(SM_CYFIXEDFRAME)
            lButtonOffset = lButtonOffset + GetSystemMetrics(SM_CXFIXEDFRAME)
            nbStdButtons = 1
            
        Case vbSizableToolWindow
            ' Small button
            buttonWidth = GetSystemMetrics(SM_CXSMSIZE)
            buttonHeight = GetSystemMetrics(SM_CYSMSIZE)
            CaptionHeight = GetSystemMetrics(SM_CYSMCAPTION)
            lCaptionTop = GetSystemMetrics(SM_CYSIZEFRAME)
            lButtonOffset = lButtonOffset + GetSystemMetrics(SM_CXFIXEDFRAME)
            nbStdButtons = 1
            
    End Select
    
    ' GetSystemMetrics returns one pixel more than the
    ' actual height of a title bar when a version 3.x
    ' application requests the SM_CYCAPTION system
    ' metric value. A version 4.0 application receives
    ' the actual height of the title bar.
    CaptionHeight = CaptionHeight - 1
    
    '---- Why ???, a border of the button ?
    buttonHeight = buttonHeight - 4
    buttonWidth = buttonWidth - 1
    
    Dim frmRect As RECT
    GetWindowRect m_Form.hwnd, frmRect
    
    Dim iLeft As Long, iTop As Double
    iLeft = (frmRect.Right - frmRect.Left - lButtonOffset) - ((nbStdButtons + 1) * buttonWidth)
    iTop = lCaptionTop + (CaptionHeight - buttonHeight) / 2
    SetRect GetButtonRect, iLeft, iTop, iLeft + buttonWidth, iTop + buttonHeight
    
    m_ButtonRect = GetButtonRect
    
End Function

Private Sub PaintButton()

    Dim DestHeight As Long
    Dim ButWidth As Long
    Dim tmpDC As Long
    Dim hBr As Long
    Dim TT As RECT
    
    '---- Get the button size
    GetWindowRect m_Form.hwnd, TT
    hBr = CreateSolidBrush(RGB(255, 0, 0))
    tmpDC = GetWindowDC(m_Form.hwnd)
     
    Dim r As RECT
    r = GetButtonRect
    
    '---- Draw the button
    m_IsPainting = True
    If m_IsRolledUp = False Then
        DrawFrameControl tmpDC, r, DFC_SCROLL, DFCS_SCROLLUP
    Else
        DrawFrameControl tmpDC, r, DFC_SCROLL, DFCS_SCROLLDOWN
    End If
    counter = counter + 1
    Debug.Print counter
    m_IsPainting = False
    
    ReleaseDC m_Form.hwnd, tmpDC
    
End Sub

Private Sub PaintPressedButton()

    Dim DestHeight As Long
    Dim ButWidth As Long
    Dim tmpDC As Long
    Dim hBr As Long
    Dim TT As RECT
    
    '---- Get the button size
    GetWindowRect m_Form.hwnd, TT
    hBr = CreateSolidBrush(RGB(255, 0, 0))
    tmpDC = GetWindowDC(m_Form.hwnd)
    
    Dim r As RECT
    r = GetButtonRect
    
    '---- Draw the button
    m_IsPainting = True
    If m_IsRolledUp = False Then
        DrawFrameControl tmpDC, r, DFC_SCROLL, 528
    Else
        DrawFrameControl tmpDC, r, DFC_SCROLL, 1 'Oops
    End If
    m_IsPainting = False
    
    ReleaseDC m_Form.hwnd, tmpDC
    
End Sub


Private Sub RefreshButtonAndMenu()

    Dim hSysMenu As Long
    Dim nCnt As Long
    Dim ret As Long
    
    If (m_Form.WindowState = vbMinimized Or m_Form.WindowState = vbMaximized) Then Exit Sub

    '---- Change the menu item
    hSysMenu = GetSystemMenu(m_Form.hwnd, False)
    If hSysMenu Then
    
        ' If the menu item exist, remove it
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            ret = RemoveMenu(hSysMenu, nCnt - 3, MF_BYPOSITION Or MF_REMOVE)
        End If
        
        DrawMenuBar m_Form.hwnd
        
    End If
    
    PaintButton
    
    ' Add the menu item
    AddItemToMenu
    
End Sub

'============================================================
'                      Frame menu management
'============================================================

Private Sub AddItemToMenu()

    Dim hSysMenu As Long
    Dim nCnt As Long
    hSysMenu = GetSystemMenu(m_Form.hwnd, False)

    If hSysMenu Then
    
        '---- Get System menu's menu count
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
        
            ' Add the right menu item
            If m_IsRolledUp = False Then
                InsertMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_STRING, ByVal 0&, "Roll up"
            Else
                InsertMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_STRING, ByVal 0&, "Roll down"
            End If
            
            ' Repaint
            DrawMenuBar m_Form.hwnd
            PaintButton
            
        End If
        
    End If
    
End Sub

'============================================================
'                      Shading management
'============================================================

Private Sub RollUpForm()

    If (m_Form.WindowState = vbMinimized) Then Exit Sub
    
    '---- During shadowing
    m_IsRolling = True
    
    '---- Save the window height
    Select Case m_Form.WindowState
        Case vbNormal
            m_NormalSize = ScaleY(m_Form.Height, m_Form.ScaleMode, vbPixels)
        Case vbMinimized
        Case vbMaximized
            m_MaximizedSize = ScaleY(m_Form.Height, m_Form.ScaleMode, vbPixels)
    End Select
                    
    '---- Roll up
    ' Use move window because with form.height we can resize
    ' a form that is maximized.
    Dim lDestHeight As Long
    Dim rctWin As RECT
    
    '----- Calculate the caption size
    Call GetWindowRect(m_Form.hwnd, rctWin)
    If (m_Form.BorderStyle = vbFixedToolWindow Or m_Form.BorderStyle = vbSizableToolWindow) Then
        lDestHeight = GetSystemMetrics(SM_CYSMCAPTION)
    Else
        lDestHeight = GetSystemMetrics(SM_CYCAPTION)
    End If
    
    ' Add frame thickness
    If (m_Form.BorderStyle = vbSizable Or m_Form.BorderStyle = vbSizableToolWindow) Then
        lDestHeight = lDestHeight + 2 * GetSystemMetrics(SM_CYSIZEFRAME)
    Else
        lDestHeight = lDestHeight + 2 * GetSystemMetrics(SM_CYFIXEDFRAME)
    End If
    
    '---- On MDI convert screen coordinates to client coordinates
    If (m_IsOnMDI) Then
        Dim paTL As POINTAPI, paBR As POINTAPI
        paTL.X = rctWin.Left
        paTL.Y = rctWin.Top
        paBR.X = rctWin.Right
        paBR.Y = rctWin.Bottom
        
        Dim mdiHWnd As Long
        mdiHWnd = GetParent(m_Form.hwnd)
        Call ScreenToClient(mdiHWnd, paTL)
        Call ScreenToClient(mdiHWnd, paBR)

        rctWin.Left = paTL.X
        rctWin.Top = paTL.Y
        rctWin.Right = paBR.X
        rctWin.Bottom = paBR.Y
    End If
    
    Call MoveWindow(m_Form.hwnd, rctWin.Left, rctWin.Top, (rctWin.Right - rctWin.Left), lDestHeight, 1)
    
    '---- Set the shadow mode
    m_IsRolledUp = True
    
    '---- Change the menu item and repaint the button
    RefreshButtonAndMenu
    
    '---- Shadowing is finished
    m_IsRolling = False
    
End Sub

Private Sub RollDownForm()

    If (m_Form.WindowState = vbMinimized) Then Exit Sub

    '---- During shadowing
    m_IsRolling = True
    
    '---- Roll down
    Dim lDestHeight As Long
    Select Case m_Form.WindowState
    Case vbNormal
        lDestHeight = m_NormalSize
    Case vbMinimized
    Case vbMaximized
        lDestHeight = m_MaximizedSize
    End Select
    
    m_Form.Height = lDestHeight
     
    '---- Roll mode
    m_IsRolledUp = False
    
    '---- Change the menu item and repaint the button
    RefreshButtonAndMenu
    
    '---- Shadowing is finished
    m_IsRolling = False
    
End Sub
