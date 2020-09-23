Attribute VB_Name = "ModMenu"
'*********Module Copyright PSST Software 2003**************
'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

Option Explicit
'API - stripped to absolute minimum required
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemID As Long
    itemAction As Long
    itemState As Long
    hWndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type
Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemID As Long
    ItemWidth As Long
    ItemHeight As Long
    ItemData As Long
End Type
Private Type MENUITEMINFO
     cbSize As Long
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As String
     cch As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function PolylineTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hbr As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal byPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As Any) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_SUNKENOUTER = &H2
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_NOCLIP = &H100
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const ODT_MENU = 1
Private Const ODS_SELECTED = &H1
Private Const ODS_DISABLED = &H4
Private Const ODS_CHECKED = &H8
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_MENU = 4
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const DSS_DISABLED = &H20
Private Const DST_ICON = &H3
Private Const MIIM_ID As Long = &H2
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_SUBMENU As Long = &H4
Private Const MIIM_TYPE = &H10
Private Const MF_OWNERDRAW = &H100&
Private Const MF_SEPARATOR = &H800
Private Const GWL_WNDPROC = (-4)
Private Const WM_CLOSE = &H10
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
Private gOldProc As Long
Private BoboMenu As Collection
Private mVBMenus As Collection
Private ParForm As Form
Private MenuBarBase As Long
Private FormScaleMode As Long
Public IL As ImageList

'In order to have images a menu must be 'ownerdrawn'
'If a menu is 'ownerdrawn' the application must respond to
'the windows messages 'WM_MEASUREITEM' to set the dimensions
'of the menu and 'WM_DRAWITEM' to print the caption and
'draw the icon. So we need to subclass the form containing
'the menu. Subclassing is automatically removed when the
'form unloads
Private Sub SubClass(mHwnd As Long)
    gOldProc& = GetWindowLong(mHwnd, GWL_WNDPROC)
    Call SetWindowLong(mHwnd, GWL_WNDPROC, AddressOf MenuProc)
End Sub
Private Sub UnSubClass(mHwnd As Long)
    Call SetWindowLong(mHwnd, GWL_WNDPROC, gOldProc&)
End Sub
Private Function MenuProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim z As Long
    Dim MeasureInfo As MEASUREITEMSTRUCT
    Dim DrawInfo As DRAWITEMSTRUCT
    Dim BB As BBMenu
    Dim R As RECT
    Dim chkR As RECT
    Dim IsSelected As Boolean
    Dim IsChecked As Boolean
    Dim IsDisabled As Boolean
    Select Case wMsg&
        Case WM_MEASUREITEM
            '************************************
            'Set menu size
            FormScaleMode = ParForm.ScaleMode
            ParForm.ScaleMode = vbPixels 'API uses pixels
            Call CopyMemory(MeasureInfo, ByVal lParam, Len(MeasureInfo)) 'get the UDT for the menuitem's dimensions
            If MeasureInfo.CtlType <> ODT_MENU Then GoTo Done 'bail - not 'ownerdrawn'
            Set BB = BoboMenu(Str(MeasureInfo.ItemID)) 'get menu details from the class
            'Adjust dimensions as neccessary
            MeasureInfo.ItemHeight = ParForm.TextHeight(BB.Caption) + 6
            MeasureInfo.ItemWidth = ParForm.TextWidth(BB.Caption) + IIf(Len(BB.KeyAccel) = 0, 36, ParForm.TextWidth(BB.KeyAccel) + 4) + 0
            
            'Return the UDT with the new values
            Call CopyMemory(ByVal lParam, MeasureInfo, Len(MeasureInfo))
            ParForm.ScaleMode = FormScaleMode 'Return the forms' original scalemode
        Case WM_DRAWITEM
            '************************************
            'Gather the information about a menu item
            FormScaleMode = ParForm.ScaleMode
            ParForm.ScaleMode = vbPixels 'API uses pixels
            Call CopyMemory(DrawInfo, ByVal lParam, LenB(DrawInfo)) 'get the UDT for the menuitem's appearance
            If DrawInfo.CtlType <> ODT_MENU Then GoTo Done 'bail - not 'ownerdrawn'
            IsSelected = ((DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED) 'selected ?
            IsDisabled = ((DrawInfo.itemState And ODS_DISABLED) = ODS_DISABLED) 'disabled ?
            IsChecked = ((DrawInfo.itemState And ODS_CHECKED) = ODS_CHECKED) 'checked ?
            If BoboMenu(Str(DrawInfo.ItemID)) Is Nothing Then GoTo Done
            Set BB = BoboMenu(Str(DrawInfo.ItemID)) 'get menu details from the class
            R = DrawInfo.rcItem
            
            '************************************
            'Set background and forecolor appropriately
            If IsSelected And Not IsDisabled Then
                FillRect DrawInfo.hdc, R, GetSysColorBrush(COLOR_HIGHLIGHT) ' paint blue background for selection
                SetTextColor DrawInfo.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT) ' write in a color that will be readable through blue background
            ElseIf IsDisabled Then
                FillRect DrawInfo.hdc, R, GetSysColorBrush(COLOR_MENU) ' paint gray background
                SetTextColor DrawInfo.hdc, vbWhite ' text white for disabled (we'll write this text again as gray, offset by one pixel to look 'disabled')
            Else
                FillRect DrawInfo.hdc, R, GetSysColorBrush(COLOR_MENU) ' paint gray background
                SetTextColor DrawInfo.hdc, GetSysColor(COLOR_MENUTEXT) ' normal text color
            End If
            SetBkMode DrawInfo.hdc, 1 ' write text transparent
            
            '*****************************************
            'Do the accellerator keys first
            If Len(BB.KeyAccel) > 0 Then
                DrawText DrawInfo.hdc, BB.KeyAccel & "    ", Len(BB.KeyAccel) + 4, R, DT_SINGLELINE Or DT_RIGHT Or DT_NOCLIP Or DT_VCENTER
                If IsDisabled Then
                    'Do it again with gray and offset by one pixel
                    SetTextColor DrawInfo.hdc, GetSysColor(COLOR_GRAYTEXT)
                    OffsetRect R, -1, -1
                    DrawText DrawInfo.hdc, BB.KeyAccel & "    ", Len(BB.KeyAccel) + 4, R, DT_SINGLELINE Or DT_RIGHT Or DT_NOCLIP Or DT_VCENTER
                    OffsetRect R, 1, 1
                End If
            End If
            
            '*****************************************
            'Do the caption next
            OffsetRect R, 26, 0
            SetTextColor DrawInfo.hdc, GetSysColor(COLOR_MENUTEXT)
            If IsSelected Then SetTextColor DrawInfo.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
            If IsDisabled Then SetTextColor DrawInfo.hdc, vbWhite
            DrawText DrawInfo.hdc, BB.Caption, Len(BB.Caption), R, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
            If IsDisabled Then
                'Do it again with gray and offset by one pixel
                SetTextColor DrawInfo.hdc, GetSysColor(COLOR_GRAYTEXT)
                OffsetRect R, -1, -1
                DrawText DrawInfo.hdc, BB.Caption, Len(BB.Caption), R, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
            End If
            
            '*****************************************
            'Draw the icon - if any
            If Val(BB.vbMenuID.Tag) > 0 Then
                If IsChecked Then
                    'draw a sunken rectangle before painting the icon
                    SetRect chkR, 2, R.Top, 22, R.Top + 19
                    DrawEdge DrawInfo.hdc, chkR, BDR_SUNKENOUTER, BF_RECT
                    SetRect chkR, 3, R.Top + 1, 21, R.Top + 18
                    'FillRect DrawInfo.hdc, chkR, IIf(IsSelected, GetSysColorBrush(COLOR_HIGHLIGHT), GetSysColorBrush(COLOR_BTNHIGHLIGHT))
                End If
                If IsDisabled Then
                    DrawState DrawInfo.hdc, 0, 0, IL.ListImages(Val(BB.vbMenuID.Tag)).Picture, 0, 4, R.Top + 2, 16, 16, DST_ICON Or DSS_DISABLED
                Else
                    IL.ListImages(Val(BB.vbMenuID.Tag)).Draw DrawInfo.hdc, 4, R.Top + 2, 1
                End If
            Else
                If IsChecked Then
                    'No icon - do a checkmark
                    'This checkmark is a bit cheesy
                    'Do your own adjustments...
                    Dim Pt(0 To 2) As POINTAPI
                    Dim hRPen As Long
                    If IsDisabled Then
                        Pt(0).x = 7
                        Pt(0).y = 9
                        Pt(1).x = 9
                        Pt(1).y = 13
                        Pt(2).x = 13
                        Pt(2).y = 5
                        hRPen = CreatePen(0, 2, vbWhite)
                        DeleteObject SelectObject(DrawInfo.hdc, hRPen)
                        MoveToEx DrawInfo.hdc, 7, 9, ByVal 0&
                        PolylineTo DrawInfo.hdc, Pt(0), 3
                        DeleteObject hRPen
                    End If
                    Pt(0).x = 6
                    Pt(0).y = 8
                    Pt(1).x = 8
                    Pt(1).y = 12
                    Pt(2).x = 12
                    Pt(2).y = 4
                    If IsSelected And Not IsDisabled Then
                        hRPen = CreatePen(0, 2, GetSysColor(COLOR_HIGHLIGHTTEXT))
                    ElseIf IsDisabled Then
                        hRPen = CreatePen(0, 2, GetSysColor(COLOR_GRAYTEXT))
                    Else
                        hRPen = CreatePen(0, 2, GetSysColor(COLOR_MENUTEXT))
                    End If
                    DeleteObject SelectObject(DrawInfo.hdc, hRPen)
                    MoveToEx DrawInfo.hdc, 6, 8, ByVal 0&
                    PolylineTo DrawInfo.hdc, Pt(0), 3
                    DeleteObject hRPen
                End If
            End If
            'Return the scalemode to original
            ParForm.ScaleMode = FormScaleMode
            'Also...
            'If you wanted to add a tooltip in the status bar this is where to start
        Case WM_CLOSE
            'Form is unloading - unsubclass
            Call SetWindowLong(hwnd&, GWL_WNDPROC, gOldProc&)
    End Select
    MenuProc = CallWindowProc(gOldProc&, hwnd&, wMsg&, wParam&, lParam&)
    Exit Function
Done:
    ParForm.ScaleMode = FormScaleMode 'Return the forms' original scalemode

End Function
Public Sub ConvertOD(mForm As Form)
    Dim BB As BBMenu
    Dim cnt As Long
    Dim cnt2 As Long
    Dim z As Long
    Dim mmID As Long
    Dim mnuAcc As String
    Dim temp As String
    Dim mnu As Control
    Dim InVisMenus As Collection
    Set InVisMenus = New Collection
    Set mVBMenus = New Collection
    Set ParForm = mForm
    Set BoboMenu = New Collection
    'If a VBmenu is not visible the the API wont see it
    'so make them all visible
    For Each mnu In ParForm.Controls
        If TypeOf mnu Is Menu Then
            If Not mnu.Visible Then
                InVisMenus.Add mnu
                mnu.Visible = True
            End If
            mVBMenus.Add mnu
        End If
    Next
    MenuBarBase = GetMenu(ParForm.hwnd)
    If MenuBarBase = 0 Then Exit Sub ' form has no menu - bail out !!
    cnt = GetMenuItemCount(MenuBarBase)
    For z = 0 To cnt - 1
        cnt2 = GetSubMenu(MenuBarBase, z)
        Set BB = New BBMenu
        temp = GetCaption(MenuBarBase, z, True, mmID, mnuAcc)
        With BB
            .Caption = temp
            .Handle = mmID
            .ParentHandle = MenuBarBase
            .KeyAccel = mnuAcc
            BoboMenu.Add BB, Str(cnt2)
        End With
        'This is a recursive sub and will get all menu items
        GetSubs cnt2
    Next
    'Reset any menus we made visible
    If InVisMenus.Count > 0 Then
        For z = 1 To InVisMenus.Count
            InVisMenus(z).Visible = False
        Next
    End If
    'Add to each class the VBmenu
    If BoboMenu.Count > 0 Then
        For z = 1 To mVBMenus.Count
            Set BoboMenu(z).vbMenuID = mVBMenus(z)
        Next
        'OK hook the form
        SubClass ParForm.hwnd
    End If
End Sub
Private Sub GetSubs(mPar As Long)
    Dim cnt As Long
    Dim cnt2 As Long
    Dim z As Long
    Dim temp As String
    Dim mmID As Long
    Dim mnuAcc As String
    Dim BB As BBMenu
    cnt = GetMenuItemCount(mPar)
    For z = 0 To cnt - 1
        cnt2 = GetSubMenu(mPar, z)
        'If the menu's parent is the Forms menu it must be a main menu
        'and therefore has no icon - (mPar = MenuBarBase)
        temp = GetCaption(mPar, z, (mPar = MenuBarBase), mmID, mnuAcc)
        Set BB = New BBMenu
        'Grab all the info we'll need from each menu
        'and place in the class
        With BB
            .Caption = temp
            .Handle = mmID
            .ParentHandle = mPar
            .KeyAccel = mnuAcc
            BoboMenu.Add BB, Str(mmID)
            If BoboMenu.Count = mVBMenus.Count Then Exit Sub
            GetSubs cnt2
        End With
NextPlease:
    Next
End Sub

Private Function GetCaption(hMnu As Long, mIndex As Long, NoOD As Boolean, ByRef mnuID As Long, ByRef mnuAcc As String) As String
    Dim MII As MENUITEMINFO
    Dim MI(0 To 1023) As Byte
    Dim sBuffer As String * 80
    Dim temp As String
    Dim IsSep As Boolean
    MII.cbSize = Len(MII)
    MII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
    MII.fType = 0
    MII.dwTypeData = sBuffer
    MII.cch = 80
    GetMenuItemInfo hMnu, mIndex, True, MII
    mnuID = MII.wID
    IsSep = ((MII.fType And MF_SEPARATOR) = MF_SEPARATOR)
    temp = Left(Replace(MII.dwTypeData, Chr(0), ""), MII.cch)
    If InStr(1, temp, Chr(9)) Then
        GetCaption = Split(temp, Chr(9))(0)
        mnuAcc = Split(temp, Chr(9))(1)
    Else
        GetCaption = temp
        mnuAcc = ""
    End If
    'If it's not a main menu or a separator then convert to 'ownerdrawn'
    If Not NoOD And IsSep = False Then
        MII.fType = MII.fType Or MF_OWNERDRAW
        SetMenuItemInfo hMnu, mIndex, True, MII
    End If
End Function
Public Sub AddODMenu(NewMenu As Menu, ParMenu As Menu)
    Dim cnt As Long
    Dim z As Long
    Dim temp As String
    Dim mmID As Long
    Dim mnuAcc As String
    Dim BB As BBMenu
    For z = 1 To mVBMenus.Count
        If mVBMenus(z) Is ParMenu Then Exit For
    Next
    cnt = GetMenuItemCount(BoboMenu(z).Handle)
    Set BB = New BBMenu
    temp = GetCaption(BoboMenu(z).Handle, cnt - 1, False, mmID, mnuAcc)
    With BB
        .Caption = temp
        .Handle = mmID
        .ParentHandle = BoboMenu(z).Handle
        .KeyAccel = mnuAcc
        mVBMenus.Add NewMenu
        Set .vbMenuID = NewMenu
        BoboMenu.Add BB, Str(mmID)
    End With

End Sub
