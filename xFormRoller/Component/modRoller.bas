Attribute VB_Name = "modRoller"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
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
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = (-4)
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const MF_STRING = &H0&
Private Const SM_CXFRAME = 32
Private Const SM_CXSMSIZE = 30
Private Const SM_CYCAPTION = 4
Private Const SM_CYDLGFRAME = 8
Private Const TPM_LEFTALIGN = &H0&
Private Const WM_ACTIVATE = &H6
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MENUSELECT = &H11F
Private Const WM_NCHITTEST = &H84
Private Const WM_NCLBUTTONDOWN = &HA1

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

Private ButhWnd As Long
Private CursorInBOX As Boolean
Private CursorLoc As POINTAPI
Private oldWndProc As Long
Private R As RECT
Private RollMode As Boolean

Private m_NormalSize As Long
Private m_MaximizedSize As Long
Private m_IsShadowing As Boolean
Private m_IsOnMDI As Boolean

Private m_FrmName As Form

'============================================================
'                      Sub classing
'============================================================

Private Function HOOK() As Long

    oldWndProc = GetWindowLong(ButhWnd, GWL_WNDPROC)
    HOOK = SetWindowLong(ButhWnd, GWL_WNDPROC, AddressOf CallBack)
    
End Function

Private Function UNHOOK() As Long

    If oldWndProc Then
        UNHOOK = SetWindowLong(ButhWnd, GWL_WNDPROC, oldWndProc)
    End If
    
End Function

Private Function CallBack(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long

    Dim ButRgn As Long
    Dim lHeight As Long
    
    Select Case uMsg
        Case WM_NCHITTEST
            CopyMemory ByVal VarPtr(CursorLoc.X), ByVal VarPtr(lparam), 2
            CopyMemory ByVal VarPtr(CursorLoc.Y), ByVal VarPtr(lparam) + 2, 2
            
            Dim TT As RECT
            GetWindowRect ButhWnd, TT
            ButRgn = CreateRectRgn(R.Left, R.Top, R.Right, R.Bottom)
            If PtInRegion(ButRgn, (CursorLoc.X - TT.Left), (CursorLoc.Y - TT.Top)) Then
                CursorInBOX = True
            Else
                CursorInBOX = False
            End If
            DeleteObject ButRgn
        
        Case WM_SIZE
            If GetActiveWindow = ButhWnd Then PaintButton
            PaintButton
            
            ' Save the size for the good window state
            If (Not m_IsShadowing And Not RollMode) Then
                ' Get the client height
                lHeight = lparam \ &H10000 And &HFFFF&
    
                Select Case m_FrmName.WindowState
                Case vbNormal
                    m_NormalSize = lHeight
                Case vbMinimized
                Case vbMaximized
                    m_MaximizedSize = lHeight
                End Select
            End If
            
            ' Restore the information if necessary
            Select Case wParam
                Case SIZE_MAXIMIZED
                    If (RollMode) Then
                        RollMode = False
                        RefreshButtonAndMenu
                    End If

                Case SIZE_MINIMIZED
                    
                Case SIZE_RESTORED
                    If (RollMode) Then
                        RollDownForm
                        RefreshButtonAndMenu
                    End If
                    
            End Select
            
            
        Case WM_ACTIVATE, WM_MDIACTIVATE
            PaintButton
       
        Case WM_NCLBUTTONDBLCLK
            If CursorInBOX = True Then
                Exit Function
            End If
            
        Case WM_MENUSELECT
            Static Clik%
            If Clik% = 1 And wParam = -65536 Then
                If RollMode = False Then
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
                If RollMode = False Then
                    RollUpForm
                Else
                    RollDownForm
                End If
            End If
            
    End Select
    
    
    CallBack = CallWindowProc(oldWndProc, hwnd, uMsg, wParam, lparam)
    
End Function

'============================================================
'                      Drawing
'============================================================

Public Sub DrawButton(FrmName As Form)

    'FrmName.ScaleMode = vbPixels
    ButhWnd = FrmName.hwnd
    HOOK
    RollMode = False
    AddItemToMenu
    
    Set m_FrmName = FrmName
    m_IsShadowing = False
    
    m_IsOnMDI = True
    
End Sub

Private Sub PaintButton()

    Dim DestHeight As Long
    Dim ButWidth As Long
    Dim tmpDC As Long
    Dim hBr As Long
    Dim TT As RECT
    
    '---- Get the button size
    GetWindowRect ButhWnd, TT
    hBr = CreateSolidBrush(RGB(255, 0, 0))
    DestHeight = GetSystemMetrics(SM_CYCAPTION)
    ButWidth = (DestHeight - 3) * 3
    SetRect R, 4, 4, (TT.Right - TT.Left - 8) - ButWidth, DestHeight + 2
    tmpDC = GetWindowDC(ButhWnd)
    SetRect R, (TT.Right - TT.Left - 8) - ButWidth - 17, 6, ((TT.Right - TT.Left - 8) - ButWidth) - 1, DestHeight + 1
    
    '---- Draw the button
    If RollMode = False Then
        DrawFrameControl tmpDC, R, DFC_SCROLL, DFCS_SCROLLUP
    Else
        DrawFrameControl tmpDC, R, DFC_SCROLL, DFCS_SCROLLDOWN
    End If
    
    ReleaseDC ButhWnd, tmpDC
    
End Sub

Private Sub PaintPressedButton()

    Dim DestHeight As Long
    Dim ButWidth As Long
    Dim tmpDC As Long
    Dim hBr As Long
    Dim TT As RECT
    
    '---- Get the button size
    GetWindowRect ButhWnd, TT
    hBr = CreateSolidBrush(RGB(255, 0, 0))
    DestHeight = GetSystemMetrics(SM_CYCAPTION)
    ButWidth = (DestHeight - 3) * 3
    SetRect R, 4, 4, (TT.Right - TT.Left - 8) - ButWidth, DestHeight + 2
    tmpDC = GetWindowDC(ButhWnd)
    SetRect R, (TT.Right - TT.Left - 8) - ButWidth - 17, 6, ((TT.Right - TT.Left - 8) - ButWidth) - 1, DestHeight + 1
    
    '---- Draw the button
    If RollMode = False Then
        DrawFrameControl tmpDC, R, 3, 528
    Else
        DrawFrameControl tmpDC, R, 3, 1 'Oops
    End If
    
    ReleaseDC ButhWnd, tmpDC
    
End Sub

'============================================================
'                      Frame menu management
'============================================================

Private Sub AddItemToMenu()

    Dim hSysMenu As Long
    Dim nCnt As Long
    hSysMenu = GetSystemMenu(ButhWnd, False)

    If hSysMenu Then
    
        '---- Get System menu's menu count
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
        
            ' Add the right menu item
            If RollMode = False Then
                InsertMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_STRING, ByVal 0&, "Roll up"
            Else
                InsertMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_STRING, ByVal 0&, "Roll down"
            End If
            
            ' Repaint
            DrawMenuBar ButhWnd
            PaintButton
            
        End If
        
    End If
    
End Sub

'============================================================
'                      Shading management
'============================================================

Private Sub RollUpForm()

    Dim ret As Long
    Dim TT As RECT
    Dim hSysMenu As Long
    Dim nCnt As Long
    Dim DestHeight As Long
    
    m_IsShadowing = True
    
    '---- Roll the window
    ' Save the height
    ret = GetWindowRect(ButhWnd, TT)
    
    Select Case m_FrmName.WindowState
        Case vbNormal
            m_NormalSize = TT.Bottom - TT.Top
        Case vbMinimized
        Case vbMaximized
            m_MaximizedSize = TT.Bottom - TT.Top
    End Select
                    
    ' Roll up
    DestHeight = GetSystemMetrics(SM_CYCAPTION)
    ret = MoveWindow(ButhWnd, TT.Left, TT.Top, (TT.Right - TT.Left), DestHeight, 1)
    
    '---- Set the shadow mode
    RollMode = True
    
    '---- Change the menu item and repaint the button
    RefreshButtonAndMenu
    
    m_IsShadowing = False
    
End Sub

Private Sub RollDownForm(Optional lHeight As Long = -1)

    m_IsShadowing = True

    Dim hSysMenu As Long
    Dim nCnt As Long
    Dim ret As Long
    Dim TT As RECT
    Dim DestHeight As Long
    
    '---- Reset the window height
    ret = GetWindowRect(ButhWnd, TT)
    If (lHeight = -1) Then
        Select Case m_FrmName.WindowState
        Case vbNormal
            DestHeight = m_NormalSize - TT.Top
        Case vbMinimized
        Case vbMaximized
            DestHeight = m_MaximizedSize - TT.Top
        End Select
    Else
        DestHeight = lHeight
    End If
    
    ret = MoveWindow(ButhWnd, TT.Left, TT.Top, (TT.Right - TT.Left), DestHeight, 1)
    
    RollMode = False
    
    '---- Change the menu item and repaint the button
    RefreshButtonAndMenu
    
    m_IsShadowing = False
    
End Sub

Private Sub RefreshButtonAndMenu()

    Dim hSysMenu As Long
    Dim nCnt As Long
    Dim ret As Long

    '---- Change the menu item
    hSysMenu = GetSystemMenu(ButhWnd, False)
    If hSysMenu Then
    
        ' If the menu item exist, remove it
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            ret = RemoveMenu(hSysMenu, nCnt - 3, MF_BYPOSITION Or MF_REMOVE)
        End If
        
        DrawMenuBar ButhWnd
        PaintButton
        
    End If
    
    ' Add the menu item
    AddItemToMenu
    
End Sub
