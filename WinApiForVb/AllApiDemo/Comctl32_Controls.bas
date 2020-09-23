Attribute VB_Name = "Comctl32_Controls"
Option Explicit

''''''''''''''''''''''''''''''''''''
'''''''''Comctl32 controls
''''''''''''''''''''''''''''''''''''

'Put together by MH

'Main Reference used: Individual Control Information (MSDN)
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/indivcontrol.asp

'Omitted controls
'   New Comctl32 v6.0 Control; "SysLink Controls" (works only in XP)
'   Flat Scrollbars (Not supported in v6.0. Why O why???)

'Note Form MSDN:
'   You should not subclass the updated common controls with an ANSI window procedure.

'Unicode Support
'   For Shell and Common Controls Versions 5.80 and later of comctl32.dll,
'   common controls notifications support both ANSI and Unicode formats on Windows 95 systems or later.
'   The system determines which format to use by sending your window a WM_NOTIFYFORMAT message.
'   To specify a format, return NFR_ANSI(=1) for ANSI notifications or NFR_UNICODE(=2) for Unicode notifications.
'   If you do not handle this message, the system calls IsWindowUnicode to determine the format.
'   Since Windows 95 and Windows 98 always return FALSE to this function call,
'   they use ANSI notifications by default.

'Version DLL_Distribution   Platform
'======= ================   =========================================================
'4.0     All                Microsoft Windows 95/Microsoft Windows NT 4.0.
'4.7     All                Microsoft Internet Explorer 3.x.
'4.71    All                Internet Explorer 4.0.
'4.72    All                Internet Explorer 4.01 and Windows 98.
'5.0     Shlwapi.dll        Internet Explorer 5.
'6.0     Shlwapi.dll        Internet Explorer 6 and Windows XP.
'5.0     Shell32.dll        Windows 2000 and Windows Millennium Edition (Windows Me).
'6.0     Shell32.dll        Windows XP.
'5.8     Comctl32.dll       Internet Explorer 5.
'5.81    Comctl32.dll       Windows 2000 and Windows Me.
'6.0     Comctl32.dll       Windows XP.

'The following code example illustrates how you can use GetDllVersion to
'test if Comctl32.dll is version 4.71 or later.
'if GetDllVersion("comctl32.dll") >= VERSION(4,71) then
'    Proceed.
'Else
'    Use an alternate approach for older DLL versions.
'End If

'Furhter reading regarding DLL versions:
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/programmersguide/versions.asp

'Note:
'   Some comctl32 msgs can not be used with any operating system
'   except Microsoft Windows XP because it requires ComCtl32.dll version 6.00 or later.
'   ComCtl32.dll version 6.00 is not redistributable; therefore you must have Windows XP installed,
'   which contains this particular dynamic-link library (DLL). To ensure that ComCtl32.dll version 6.00
'   is available you must provide a manifest.


'WS_BORDER Or WS_THICKFRAME, gives any control a resizeable grab handle
'==========Local
Private hFont As Long
Private lStyle As Long
Private hTemp As Long

'Assigns default font ot controls as they are created
Private Sub SetCtlDefaultGuiFont(hWnd As Long)
    If hFont = NUM_ZERO Then hFont = GetStockObject(DEFAULT_GUI_FONT)
    A_SendMessage hWnd, WM_SETFONT, hFont, 1&
End Sub

'Clean up, must be called before exiting main app
Public Sub Comctl32_Terminate()
    Dim lCount As Long
    
    On Error Resume Next
    
    If hFont <> NUM_ZERO Then DeleteObject hFont
    hFont = NUM_ZERO
    lStyle = NUM_ZERO
    hTemp = NUM_ZERO
    
    'Do we have any image list
    If hTvImgList16 <> 0 Then
        If hTree > 0 Then
            'Pass 0 to clear imagelist
            TreeView_SetImageList hTree, 0
            'A_SendMessageAnyRef hTree, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal 0&
        End If
        ImageList_Destroy hTvImgList16
    End If
    If Not cNodesD Is Nothing Then Set cNodesD = Nothing
    
    If hTvImgListState16 <> 0 Then
        If hTree <> 0 Then
            'state image list
            TreeView_SetImageList hTree, 0, False
        End If
        ImageList_Destroy hTvImgListState16
    End If
    If hIeImgList <> 0 Then ImageList_Destroy hIeImgList
    If hComboImgList16 <> 0 Then ImageList_Destroy hComboImgList16
    
    
    If hMenuTree <> 0 Then DestroyMenu hMenuTree
    If hLvMenu <> 0 Then DestroyMenu hLvMenu
    If hBrush <> 0 Then DeleteObject hBrush
    
    If hMainToolTip <> 0 Then DestroyWindow hMainToolTip
    
    'Delete all fonts created if listbox form was created
    If hFrmListBox > 0 Then
        For lCount = 0 To UBound(arrFonts)
            If arrFonts(lCount) <> 0 Then DeleteObject arrFonts(lCount)
        Next
        Erase arrFonts
        If hBrushListBox <> 0 Then DeleteObject hBrushListBox
        If hListImgList <> 0 Then ImageList_Destroy hListImgList
    End If
    
    sToolbarTip = CHAR_ZERO_LENGTH_STRING
    sTreeTip = CHAR_ZERO_LENGTH_STRING
    sLvTip = CHAR_ZERO_LENGTH_STRING
    
End Sub

'Can create => PushButton, Checkbox, Radio, and Group (General)
Public Function CreateButton(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, _
                            Height As Long, Style As Long, Optional WS_EX_flags As Long = 0) As Long

    If hParent = 0 Then Exit Function
    hTemp = A_CreateWindowEx(WS_EX_flags, WNDCTRL_BUTTON, strCaption, Style, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    SetCtlDefaultGuiFont hTemp
    CreateButton = hTemp

End Function

'BS_GROUPBOX Creates a rectangle in which other controls can be grouped.
'Any text associated with this style is displayed in the rectangle's upper left corner.
Public Function CreateGroupBox(hParent As Long, strCaption As String, X As Long, Y As Long, _
                                Width As Long, _
                                Height As Long) As Long
    If hParent = 0 Then Exit Function

    hTemp = A_CreateWindowEx(0&, WNDCTRL_BUTTON, strCaption, BS_GROUPBOX Or WS_CHILD Or WS_VISIBLE, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    SetCtlDefaultGuiFont hTemp
    CreateGroupBox = hTemp
End Function

Public Function CreateCmdButton(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional blnTabStop As Boolean = True, _
                                Optional lTextorBitmaporIcon As Long = BS_TEXT, _
                                Optional BS_TextAlignement As Long = BS_CENTER, _
                                Optional blnMultiLineCaption As Boolean = False, _
                                Optional blnFlat As Boolean = False, _
                                Optional bDefault As Boolean = False, _
                                Optional bOwnerDran As Boolean = False, _
                                Optional WS_EX_flags As Long = 0) As Long
                                
    'Create a command button
    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD Or WS_VISIBLE
    If bDefault = True Then
        lStyle = lStyle Or BS_DEFPUSHBUTTON
    Else
        lStyle = lStyle Or BS_PUSHBUTTON
    End If
    'Ownerdrawn
    If lTextorBitmaporIcon = BS_OWNERDRAW Then
        lStyle = lStyle Or BS_OWNERDRAW
    Else
        If blnTabStop Then lStyle = lStyle Or WS_TABSTOP
        If blnMultiLineCaption Then lStyle = lStyle Or BS_MULTILINE
        If blnFlat Then lStyle = lStyle Or BS_FLAT
        lStyle = lStyle Or BS_TextAlignement
    End If
    'WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE or WS_EX_DLGMODALFRAME
    CreateCmdButton = CreateButton(hParent, strCaption, X, Y, Width, Height, lStyle, WS_EX_flags)
End Function

Public Function CreateCheckbox(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, Height As Long, _
                               Optional bln3States As Boolean = False, _
                               Optional blnBeginGroup As Boolean = False, _
                               Optional blnTabStop As Boolean = True, _
                               Optional blnMultiLineCaption As Boolean = True, _
                               Optional BS_TextAlignement As Long = BS_LEFT) As Long

    'Create a checkbox
    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD Or WS_VISIBLE

    If blnTabStop Then lStyle = lStyle Or WS_TABSTOP
    If blnBeginGroup Then lStyle = lStyle Or WS_GROUP
    If blnMultiLineCaption Then lStyle = lStyle Or BS_MULTILINE
    If bln3States Then
        lStyle = lStyle Or BS_AUTO3STATE
    Else
        lStyle = lStyle Or BS_AUTOCHECKBOX
    End If
    lStyle = lStyle Or BS_TextAlignement
    
    CreateCheckbox = CreateButton(hParent, strCaption, X, Y, Width, Height, lStyle)

End Function

Public Function CreateRadioButton(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, Height As Long, _
                                  Optional bAutoRadio As Boolean = True, _
                                  Optional blnBeginGroup As Boolean = False, _
                                  Optional blnTabStop As Boolean = True, _
                                  Optional blnMultiLineCaption As Boolean = True, _
                                  Optional BS_TextAlignement As Long = BS_LEFT) As Long

    'Create a radio button
    If hParent = 0 Then Exit Function
    lStyle = WS_CHILD Or WS_VISIBLE
    
    If bAutoRadio Then
        lStyle = lStyle Or BS_AUTORADIOBUTTON
    Else
        'Simple radio, system does not automaticaly uncheck other radio btns in the same group
        lStyle = lStyle Or BS_RADIOBUTTON
    End If
    If blnTabStop Then lStyle = lStyle Or WS_TABSTOP
    'WS_GROUP   Specifies the first control of a group of controls in which the user can move
    'from one control to the next with the arrow keys. All controls defined with the WS_GROUP
    'style FALSE after the first control belong to the same group. The next control with the
    'WS_GROUP style starts the next group (that is, one group ends where the next begins).
    If blnBeginGroup Then lStyle = lStyle Or WS_GROUP
    If blnMultiLineCaption Then lStyle = lStyle Or BS_MULTILINE
'BS_LEFTTEXT = Places text on the left side of the radio button or check box when combined with a radio button or check box style. Same as the BS_RIGHTBUTTON style.
'BS_RIGHTBUTTON = Positions a radio button's circle or a check box's square on the right side of the button rectangle. Same as the BS_LEFTTEXT style.
    lStyle = lStyle Or BS_TextAlignement
    
    CreateRadioButton = CreateButton(hParent, strCaption, X, Y, Width, Height, lStyle)
End Function

'Image type                             Static control style
'==========                             ====================
'IMAGE_BITMAP                           SS_BITMAP
'IMAGE_CURSOR (incl Animated)           SS_ICON
'IMAGE_ENHMETAFILE                      SS_ENHMETAFILE
'IMAGE_ICON                             SS_ICON
'W + H are automatically adjusted by control
Public Function CreateImageBox(hParent As Long, X As Long, Y As Long, _
                                Optional bOwnerDrawn As Boolean = False, _
                                Optional bBorder As Boolean = False, _
                                Optional SS_ImgType As Long = SS_ICON, _
                                Optional lInitialImg As Long = 0) As Long
    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD Or WS_VISIBLE
    If bOwnerDrawn Then
        lStyle = lStyle Or SS_OWNERDRAW
    Else
        If bBorder Then lStyle = lStyle Or WS_BORDER
        lStyle = lStyle Or SS_ImgType
    End If
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_STATIC, vbNullString, lStyle, X, Y, 0&, 0&, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    CreateImageBox = hTemp
    If bOwnerDrawn = False Then
        If lInitialImg <> 0 Then
            If SS_ImgType = SS_ICON Then
                A_SendMessage hTemp, STM_SETICON, lInitialImg, 0&
            Else
                A_SendMessage hTemp, STM_SETIMAGE, SS_ImgType, lInitialImg
            End If
        End If
    End If
End Function

Public Function CreateLabel(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, Height As Long, _
                            Optional SS_Alignment As Long = SS_LEFT, _
                            Optional SS_AdditionalStyles As Long = 0) As Long

'Usefull styles::
'SS_NOPREFIX = Prevents interpretation of any ampersand (&) characters in the control's text as accelerator prefix characters.
'SS_NOTIFY = Sends the parent window STN_CLICKED, STN_DBLCLK, STN_DISABLE, and STN_ENABLE notification messages when the user clicks or double-clicks the control.

'Very usefull styles!!!
'SS_ENDELLIPSIS
'SS_PATHELLIPSIS
'SS_WORDELLIPSIS

'SS_SUNKEN

'These styles allows us to create a Frame like window with various borders, no text
'To create a frame with text use
'CreateButton passing << BS_GROUPBOX Or WS_CHILD Or WS_VISIBLE >> as WS_EX_flags param
'SS_ETCHEDFRAME = Draws the frame of the static control using the EDGE_ETCHED edge style
'SS_ETCHEDHORZ
'SS_ETCHEDVERT

'SS_BLACKFRAME
'SS_BLACKRECT
'SS_WHITEFRAME
'SS_WHITERECT
'SS_GRAYFRAME
'SS_GRAYRECT

    'Creates a label and returns its handle
    If hParent = 0 Then Exit Function
    lStyle = WS_CHILD Or WS_VISIBLE Or SS_Alignment
    If SS_AdditionalStyles > NUM_ZERO Then lStyle = lStyle Or SS_AdditionalStyles
    hTemp = A_CreateWindowEx(0&, WNDCTRL_STATIC, strCaption, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    SetCtlDefaultGuiFont hTemp
    CreateLabel = hTemp

End Function

'WS_EX_flags controls the kind of border
'to get a flat border, pass 0
Public Function CreateTextbox(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional sText As String = "Text", _
                                Optional WS_EX_flags As Long = WS_EX_CLIENTEDGE, _
                                Optional bReadOnly As Boolean = False, _
                                Optional blnTabStop As Boolean = True, _
                                Optional bFlatBorder As Boolean = False, _
                                Optional ES_TextAlignment As Long = ES_LEFT, _
                                Optional bMultiLine As Boolean = False, _
                                Optional bAutoHScroll As Boolean = True, _
                                Optional bAutoVScroll As Boolean = False, _
                                Optional bPassword As Boolean = False, _
                                Optional bUpperCase As Boolean = False, _
                                Optional bLowerCase As Boolean = False, _
                                Optional bNumber As Boolean = False, _
                                Optional bShowVScrollbar As Boolean = True, _
                                Optional bShowHScrollbar As Boolean = True) As Long

    'Creates a textbox and returns its handle
    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD Or WS_VISIBLE
    If blnTabStop Then lStyle = lStyle Or WS_TABSTOP
    If bFlatBorder Then lStyle = lStyle Or WS_BORDER
    If bReadOnly Then lStyle = lStyle Or ES_READONLY
    lStyle = lStyle Or ES_TextAlignment
    If bMultiLine Then
        lStyle = lStyle Or ES_MULTILINE Or ES_WANTRETURN
        If bShowVScrollbar Then lStyle = lStyle Or WS_VSCROLL
        If bShowHScrollbar Then lStyle = lStyle Or WS_HSCROLL
    End If
    If bAutoHScroll Then lStyle = lStyle Or ES_AUTOHSCROLL
    If bAutoVScroll Then lStyle = lStyle Or ES_AUTOVSCROLL
    If bPassword Then lStyle = lStyle Or ES_PASSWORD
    If bUpperCase Then lStyle = lStyle Or ES_UPPERCASE
    If bLowerCase Then lStyle = lStyle Or ES_LOWERCASE
    If bNumber Then lStyle = lStyle Or ES_NUMBER
    
    hTemp = A_CreateWindowEx(WS_EX_flags, WNDCTRL_EDIT, sText, lStyle, X, Y, Width, Height, hParent, 0, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    SetCtlDefaultGuiFont hTemp
    CreateTextbox = hTemp

End Function

'setting x = CW_USEDEFAULT is valid only for WS_OVERLAPPEDWINDOW style
Public Function CreateForm(X As Long, Y As Long, Width As Long, Height As Long, _
                            lWndProcAddress As Long, Optional hWndParent As Long = 0, _
                            Optional winStyle As Long = WS_OVERLAPPEDWINDOW, _
                            Optional WS_EX_flags As Long = WS_EX_CONTROLPARENT, _
                            Optional sCLASSNAME As String = "CustomClass", _
                            Optional sCaption As String = "Form", _
                            Optional COLOR_BackColor As Long = COLOR_BTNFACE, _
                            Optional lIcon As Long = NUM_ZERO, _
                            Optional lCursor As Long = NUM_ZERO) As Long

'WS_EX_CONTROLPARENT:
'   The window itself contains child windows that should take part in dialog box navigation.
'   If this style is specified, the dialog manager recurses into children of this window when
'   performing navigation operations such as handling the TAB key, an arrow key, or a keyboard mnemonic.
    
    'if using WNDCLASSEX, we need to set the cbSize member first unlike WNDCLASS
    'and use RegisterClassEx
    Dim wc As WNDCLASS
    
    'Creates a form and returns its handle
    With wc
        .hInstance = App.hInstance
        .lpfnWndProc = lWndProcAddress 'GetAdd(AddressOf WndProc)
        .hbrBackground = COLOR_BackColor + 1
        .lpszClassName = StrPtr(StrConv(sCLASSNAME, vbFromUnicode))
        'This member must be a handle to an icon resource.
        If lIcon > NUM_ZERO Then .hIcon = lIcon
        'This member must be a handle to an cursor resource.
        If lCursor > NUM_ZERO Then .hCursor = lCursor
    End With

    hTemp = A_RegisterClass(wc)
    'Debug.Print "hTemp:" & hTemp
    If hTemp <= 0 Then Exit Function
    'winStyle = winStyle Or WS_TABSTOP
    hTemp = A_CreateWindowEx(WS_EX_flags, sCLASSNAME, sCaption, winStyle, X, Y, Width, Height, hWndParent, 0, App.hInstance, ByVal 0&)

    CreateForm = hTemp
    
'WS_BORDER   Creates a window that has a border.
'WS_CAPTION   Creates a window that has a title bar (implies the WS_BORDER style). Cannot be used with the WS_DLGFRAME style.
'WS_CHILD   Creates a child window. Cannot be used with the WS_POPUP style.
'WS_CHILDWINDOW   Same as the WS_CHILD style.
'WS_CLIPCHILDREN   Excludes the area occupied by child windows when you draw within the parent window. Used when you create the parent window.
'WS_CLIPSIBLINGS   Clips child windows relative to each other; that is, when a particular child window receives a paint message, the WS_CLIPSIBLINGS style clips all other overlapped child windows out of the region of the child window to be updated. (If WS_CLIPSIBLINGS is not given and child windows overlap, when you draw within the client area of a child window, it is possible to draw within the client area of a neighboring child window.) For use with the WS_CHILD style only.
'WS_DISABLED   Creates a window that is initially disabled.
'WS_DLGFRAME   Creates a window with a double border but no title.
'WS_GROUP   Specifies the first control of a group of controls in which the user can move from one control to the next with the arrow keys. All controls defined with the WS_GROUP style FALSE after the first control belong to the same group. The next control with the WS_GROUP style starts the next group (that is, one group ends where the next begins).
'WS_HSCROLL   Creates a window that has a horizontal scroll bar.
'WS_ICONIC   Creates a window that is initially minimized. Same as the WS_MINIMIZE style.
'WS_MAXIMIZE   Creates a window of maximum size.
'WS_MAXIMIZEBOX   Creates a window that has a Maximize button.
'WS_MINIMIZE   Creates a window that is initially minimized. For use with the WS_OVERLAPPED style only.
'WS_MINIMIZEBOX   Creates a window that has a Minimize button.
'WS_OVERLAPPED   Creates an overlapped window. An overlapped window usually has a caption and a border.
'WS_OVERLAPPEDWINDOW   Creates an overlapped window with the WS_OVERLAPPED, WS_CAPTION, WS_SYSMENU, WS_THICKFRAME, WS_MINIMIZEBOX, and WS_MAXIMIZEBOX styles.
'WS_POPUP   Creates a pop-up window. Cannot be used with the WS_CHILD style.
'WS_POPUPWINDOW   Creates a pop-up window with the WS_BORDER, WS_POPUP, and WS_SYSMENU styles. The WS_CAPTION style must be combined with the WS_POPUPWINDOW style to make the Control menu visible.
'WS_SIZEBOX   Creates a window that has a sizing border. Same as the WS_THICKFRAME style.
'WS_SYSMENU   Creates a window that has a Control-menu box in its title bar. Used only for windows with title bars.
'WS_TABSTOP   Specifies one of any number of controls through which the user can move by using the TAB key. The TAB key moves the user to the next control specified by the WS_TABSTOP style.
'WS_THICKFRAME   Creates a window with a thick frame that can be used to size the window.
'WS_TILED   Creates an overlapped window. An overlapped window has a title bar and a border. Same as the WS_OVERLAPPED style.
'WS_TILEDWINDOW   Creates an overlapped window with the WS_OVERLAPPED, WS_CAPTION, WS_SYSMENU, WS_THICKFRAME, WS_MINIMIZEBOX, and WS_MAXIMIZEBOX styles. Same as the WS_OVERLAPPEDWINDOW style.
'WS_VISIBLE   Creates a window that is initially visible.
'WS_VSCROLL   Creates a window that has a vertical scroll bar.

'To use system cursors use IDC_ flags with loadcursor
'A_LoadCursor App.hInstance, IDC_WAIT

End Function

Public Function CreateIPField(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional sIP1 As Byte = 0, Optional sIP2 As Byte = 0, Optional sIP3 As Byte = 0, Optional sIP4 As Byte = 0) As Long

    Dim lIp As Long
    If hParent = 0 Then Exit Function
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_IPADDRESS, vbNullString, WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_BORDER, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    SetCtlDefaultGuiFont hTemp

    If sIP1 > NUM_ZERO And hTemp > NUM_ZERO Then
        lIp = IPADDRESS_MAKEIPADDRESS(sIP1, sIP2, sIP3, sIP4)
        A_SendMessageAnyRef hTemp, IPM_SETADDRESS, 0, ByVal lIp
    End If
    
    CreateIPField = hTemp

End Function

Public Function CreateMenuItem(strCaption As String, ID As Long, hSubMenu As Long, _
                               Optional blnSeparator As Boolean = False, Optional blnDisabled As Boolean = False, _
                               Optional blnChecked As Boolean = False) As MENUITEMINFO

Dim mnuInfo As MENUITEMINFO

    With mnuInfo
        .cbSize = Len(mnuInfo)
        .fMask = MIIM_ID Or MIIM_STRING
        If hSubMenu Then
            .fMask = .fMask Or MIIM_SUBMENU
            .hSubMenu = hSubMenu
        End If
        If blnSeparator Then
            .fMask = .fMask Or MIIM_FTYPE
            .fType = MFT_SEPARATOR
        End If
        If blnDisabled Or blnChecked Then
            .fMask = .fMask Or MIIM_STATE
            If blnDisabled Then
                .fState = MFS_DISABLED
            Else
                .fState = MFS_CHECKED
            End If
        End If
        .wID = ID
        .dwTypeData = StrPtr(StrConv(strCaption, vbFromUnicode))
        .cch = Len(.dwTypeData)
    End With
    CreateMenuItem = mnuInfo

End Function

'Set up styles in the parameter list
Public Function CreateUpDown(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bBorder As Boolean = True, Optional bTabStop As Boolean = True, _
                                Optional bAutoChooseBuddy As Boolean = False, _
                                Optional bLeftAlign As Boolean = True, Optional bNoThousands As Boolean = True, _
                                Optional bUpdateBuddy As Boolean = True, Optional bUseArrowKeys As Boolean = True, _
                                Optional bHorizental As Boolean = False, Optional bWrap As Boolean = False, _
                                Optional bHotTrack As Boolean = False, Optional WS_EX_flags As Long = 0) As Long

'"msctls_updown32"
    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD Or WS_VISIBLE 'Or WS_TABSTOP Or WS_BORDER
    
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If bTabStop = True Then lStyle = lStyle Or WS_TABSTOP
    If bAutoChooseBuddy Then lStyle = lStyle Or UDS_AUTOBUDDY
    If bLeftAlign Then
        lStyle = lStyle Or UDS_ALIGNLEFT
    Else
        lStyle = lStyle Or UDS_ALIGNRIGHT
    End If
    If bNoThousands Then lStyle = lStyle Or UDS_NOTHOUSANDS
    If bUpdateBuddy Then lStyle = lStyle Or UDS_SETBUDDYINT
    If bUseArrowKeys Then lStyle = lStyle Or UDS_ARROWKEYS
    If bHorizental Then lStyle = lStyle Or UDS_HORZ
    If bHotTrack Then lStyle = lStyle Or UDS_HOTTRACK
    If bWrap Then lStyle = lStyle Or UDS_WRAP
    
    hTemp = A_CreateWindowEx(WS_EX_flags, WNDCTRL_UPDOWN_CLASS, vbNullString, lStyle, X, Y, Width, Height, hParent, 0, App.hInstance, ByVal 0&)
    CreateUpDown = hTemp
    
End Function

Public Function CreateTreeView(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bBorder As Boolean = True, Optional bTabStop As Boolean = True, _
                                Optional WS_EX_flags As Long = WS_EX_CLIENTEDGE, Optional bHasLines As Boolean = True, _
                                Optional bLinesAsRoot As Boolean = True, Optional bHasButtons As Boolean = True, _
                                Optional bEditLabels As Boolean = True, Optional bInfoTip As Boolean = True, _
                                Optional bSingleExpand As Boolean = False, Optional bTrackSelected As Boolean = False, _
                                Optional bShowAlways As Boolean = False, Optional bNoToolTips As Boolean = False, _
                                Optional bNoSroll As Boolean = False, Optional bNoHScroll As Boolean = False, _
                                Optional bFullRowSelect As Boolean = False, Optional bDisableDragDrop As Boolean = False, _
                                Optional lAdditionalStyles As Long = NUM_ZERO) As Long

    '"SysTreeView32"
    'basic styles
    lStyle = WS_CHILD Or WS_VISIBLE
    
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If bTabStop = True Then lStyle = lStyle Or WS_TABSTOP
    If bHasLines = True Then lStyle = lStyle Or TVS_HASLINES
    If bLinesAsRoot = True Then lStyle = lStyle Or TVS_LINESATROOT
    If bHasButtons = True Then lStyle = lStyle Or TVS_HASBUTTONS
    If bEditLabels = True Then lStyle = lStyle Or TVS_EDITLABELS
    If bInfoTip = True Then
        If bNoToolTips = False Then lStyle = lStyle Or TVS_INFOTIP
    End If
    If bSingleExpand = True Then lStyle = lStyle Or TVS_SINGLEEXPAND
    If bTrackSelected = True Then lStyle = lStyle Or TVS_TRACKSELECT
    If bShowAlways = True Then lStyle = lStyle Or TVS_SHOWSELALWAYS
    If bNoToolTips = True Then lStyle = lStyle Or TVS_NOTOOLTIPS
    If bNoSroll = True Then lStyle = lStyle Or TVS_NOSCROLL
    If bNoHScroll = True Then lStyle = lStyle Or TVS_NOHSCROLL
    If bFullRowSelect = True Then
        If bHasLines = False Then lStyle = lStyle Or TVS_FULLROWSELECT
    End If
    If bDisableDragDrop = True Then lStyle = lStyle Or TVS_DISABLEDRAGDROP
    If lAdditionalStyles > NUM_ZERO Then lStyle = lStyle Or lAdditionalStyles

    If hParent = 0 Then Exit Function
    hTemp = A_CreateWindowEx(WS_EX_flags, WNDCTRL_TREEVIEW, vbNullString, lStyle, X, Y, Width, Height, hParent, 0, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    CreateTreeView = hTemp

'TVS_INFOTIP = Version 4.71. Obtains ToolTip information by sending the TVN_GETINFOTIP notification.
'TVS_FULLROWSELECT = This style cannot be used in conjunction with the TVS_HASLINES style.
'TVS_EDITLABELS
'TVS_DISABLEDRAGDROP
'TVS_CHECKBOXES
'Version 4.70. Enables check boxes for items in a tree-view control.
'   A check box is displayed only if an image is associated with the item.
'   When set to this style, the control effectively uses DrawFrameControl
'   to create and set a state image list containing two images.
'   State image 1 is the unchecked box and state image 2 is the checked box.
'   Setting the state image to zero removes the check box altogether.
'Version 5.80. Displays a check box even if no image is associated with the item.
'Once a tree-view control is created with this style, the style cannot be removed.
'   Instead, you must destroy the control and create a new one in its place.
'   Destroying the tree-view control does not destroy the check box state image list.
'   You must destroy it explicitly. Get the handle to the state image list by sending
'   the tree-view control a TVM_GETIMAGELIST message. Then destroy the image list with ImageList_Destroy.
'If you want to use this style, you must set the TVS_CHECKBOXES style with SetWindowLong
'   after you create the treeview control, and before you populate the tree.
'   Otherwise, the checkboxes might appear unchecked, depending on timing issues.

End Function

Public Function CreateDateTimePicker(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, Height As Long, _
                                    Optional bBorder As Boolean = True, Optional bTabStop As Boolean = True, _
                                    Optional bDisplayLongDateFormat As Boolean = True, _
                                    Optional bRightAligned As Boolean = True, _
                                    Optional bUseUpDown As Boolean = False, _
                                    Optional bShowNoInitSelectedDate As Boolean = False) As Long
                                    
    If hParent = 0 Then Exit Function
    'DTS_UPDOWN
    lStyle = WS_CHILD Or WS_VISIBLE 'Or WS_BORDER Or WS_TABSTOP
    
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If bTabStop = True Then lStyle = lStyle Or WS_TABSTOP
    If bDisplayLongDateFormat = True Then lStyle = lStyle Or DTS_LONGDATEFORMAT
    If bRightAligned = True Then lStyle = lStyle Or DTS_RIGHTALIGN
    If bUseUpDown = True Then lStyle = lStyle Or DTS_UPDOWN
    If bShowNoInitSelectedDate = True Then lStyle = lStyle Or DTS_SHOWNONE
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_DATETIMEPICK_CLASS, strCaption, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    SetCtlDefaultGuiFont hTemp
    CreateDateTimePicker = hTemp

End Function

Public Function CreateCalendar(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bBorder As Boolean = True, _
                                Optional lMultiSelect As Long = -1)
    If hParent = 0 Then Exit Function
    
'The month calendar will send MCN_GETDAYSTATE notifications to request information about which days should be displayed in bold.
'MCS_DAYSTATE
    
    lStyle = WS_CHILD Or WS_VISIBLE
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If lMultiSelect > -1 Then lStyle = lStyle Or MCM_SETMAXSELCOUNT
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_MONTHCAL_CLASS, 0&, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    'After creating the control, you can change all of the styles except for MCS_DAYSTATE and
    'MCS_MULTISELECT. To change these styles, you will need to destroy the existing control
    'and create a new one that has the desired styles.
    
    'Month calendar controls that use the MCS_MULTISELECT style allow the user to select
    'a range of days. By default, the control allows the user to
    'select seven contiguous days
    A_SendMessage hTemp, MCM_SETMAXSELCOUNT, lMultiSelect, 0&
    
    CreateCalendar = hTemp

'Day of Week
'0 Monday
'1 Tuesday
'2 Wednesday
'3 Thursday
'4 Friday
'5 Saturday
'6 Sunday

End Function


Public Function CreateCombo(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bBorder As Boolean = True, Optional bTabStop As Boolean = True, _
                                Optional WS_CBS_flags As Long = CBS_DROPDOWN Or CBS_AUTOHSCROLL Or WS_VSCROLL) As Long

    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD Or WS_VISIBLE
    
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If bTabStop = True Then lStyle = lStyle Or WS_TABSTOP
    lStyle = lStyle Or WS_CBS_flags

    hTemp = A_CreateWindowEx(0&, "COMBOBOX", strCaption, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    SetCtlDefaultGuiFont hTemp
    CreateCombo = hTemp

'Usefull
'CB_DIR Message

'Styles
'
'CBS_AUTOHSCROLL
'   Automatically scrolls the text in an edit control to the right when the user types a character at the end of the line. If this style is not set, only text that fits within the rectangular boundary is allowed.
'CBS_DISABLENOSCROLL
'   Shows a disabled vertical scroll bar in the list box when the box does not contain enough items to scroll. Without this style, the scroll bar is hidden when the list box does not contain enough items.
'CBS_DROPDOWN
'   Similar to CBS_SIMPLE, except that the list box is not displayed unless the user selects an icon next to the edit control.
'CBS_DROPDOWNLIST
'   Similar to CBS_DROPDOWN, except that the edit control is replaced by a static text item that displays the current selection in the list box.
'CBS_LOWERCASE
'   Converts to lowercase all text in both the selection field and the list.
'CBS_NOINTEGRALHEIGHT
'   Specifies that the size of the combo box is exactly the size specified by the application when it created the combo box. Normally, the system sizes a combo box so that it does not display partial items.
'CBS_SIMPLE
'   Displays the list box at all times. The current selection in the list box is displayed in the edit control.
'CBS_SORT
'   Automatically sorts strings added to the list box.
'CBS_UPPERCASE
'   Converts to uppercase all text in both the selection field and the list.

End Function

Public Function CreateComboEx(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bBorder As Boolean = True, Optional bTabStop As Boolean = True, _
                                Optional WS_CBS_flags As Long = CBS_DROPDOWN Or CBS_AUTOHSCROLL Or WS_VSCROLL, _
                                Optional CBES_EX_flags As Long = CBES_EX_CASESENSITIVE) As Long

    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD Or WS_VISIBLE
    
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If bTabStop = True Then lStyle = lStyle Or WS_TABSTOP
    lStyle = lStyle Or WS_CBS_flags

    hTemp = A_CreateWindowEx(CBES_EX_flags, WNDCTRL_COMBOBOXEX, strCaption, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    SetCtlDefaultGuiFont hTemp
    CreateComboEx = hTemp

'ComboBoxEx controls support most standard combo box control styles.
'Additionally, ComboBoxEx controls support the extended styles that are listed in this section.

'CBES_EX_CASESENSITIVE
'   BSTR searches in the list will be case sensitive. This includes searches as a result of text being typed in the edit box and the CB_FINDSTRINGEXACT message.
'CBES_EX_NOEDITIMAGE
'   The edit box and the dropdown list will not display item images.
'CBES_EX_NOEDITIMAGEINDENT
'   The edit box and the dropdown list will not display item images.
'CBES_EX_NOSIZELIMIT
'   Allows the ComboBoxEx control to be vertically sized smaller than its contained combo box control. If the ComboBoxEx is sized smaller than the combo box, the combo box will be clipped.
'CBES_EX_PATHWORDBREAKPROC
'   Microsoft Windows NT only. The edit box will use the slash (/), backslash (\), and period (.) characters as word delimiters. This makes keyboard shortcuts for word-by-word cursor movement () effective in path names and URLs.

End Function

Public Function CreatePageScroller(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bHoriz As Boolean = True, _
                                Optional bAutoScroll As Boolean = True, _
                                Optional bDragDrop As Boolean = False) As Long

    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD Or WS_VISIBLE 'Or WS_BORDER Or WS_TABSTOP
    
    If bHoriz = True Then
        lStyle = lStyle Or PGS_HORZ
    Else
        lStyle = lStyle Or PGS_VERT
    End If
    If bAutoScroll = True Then lStyle = lStyle Or PGS_AUTOSCROLL
    If bDragDrop = True Then lStyle = lStyle Or PGS_DRAGNDROP
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_PAGESCROLLER, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    CreatePageScroller = hTemp

End Function

'/* Richedit2.0 Window Class. */
'
'#define RICHEDIT_CLASSA     "RichEdit20A"
'#define RICHEDIT_CLASS10A   "RICHEDIT"          // Richedit 1.0
'
'#ifndef MACPORT
'#define RICHEDIT_CLASSW     L"RichEdit20W"
'#else   /*----------------------MACPORT */
'#define RICHEDIT_CLASSW     TEXT("RichEdit20W") /* MACPORT change */
'#endif /* MACPORT  */
'
'#if (_RICHEDIT_VER >= 0x0200 )
'#ifdef UNICODE
'#define RICHEDIT_CLASS      RICHEDIT_CLASSW
'#Else
'#define RICHEDIT_CLASS      RICHEDIT_CLASSA
'#endif /* UNICODE */
'#Else
'#define RICHEDIT_CLASS      RICHEDIT_CLASS10A
'#endif /* _RICHEDIT_VER >= 0x0200 */
'Rich Edit version DLL
'1.0 Riched32.dll
'2.0 Riched20.dll
'3.0 Riched20.dll
'4.1 Msftedit.dll - non-distributable, included with XP - emulates RichEdit2.0
Public Function CreateRichEdit(hParent As Long, strCaption As String, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bBorder As Boolean = True, Optional bTabStop As Boolean = True, _
                                Optional bMultiLine As Boolean = True, Optional bWantReturn As Boolean = True, _
                                Optional bDISABLENOSCROLL As Boolean = False, Optional bAutoVScroll As Boolean = True, _
                                Optional bAutoHScroll As Boolean = True, Optional bReadOnly As Boolean = False, _
                                Optional lAdditionalStyles As Long = NUM_ZERO) As Long

    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_HSCROLL
    
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If bTabStop = True Then lStyle = lStyle Or WS_TABSTOP
    If bMultiLine = True Then lStyle = lStyle Or ES_MULTILINE
    If bWantReturn = True Then lStyle = lStyle Or ES_WANTRETURN
    'Disables scroll bars instead of hiding them when they are not needed.
    If bDISABLENOSCROLL = True Then lStyle = lStyle Or ES_DISABLENOSCROLL
    If bAutoVScroll = True Then lStyle = lStyle Or ES_AUTOVSCROLL
    If bAutoHScroll = True Then lStyle = lStyle Or ES_AUTOHSCROLL
    If bReadOnly = True Then lStyle = lStyle Or ES_READONLY
    If lAdditionalStyles > NUM_ZERO Then lStyle = lStyle Or lAdditionalStyles
    
    'This call is necessary specially in XP, as XP does not load older libraries
    If A_LoadLibrary("C:\Windows\system32\riched20.dll") <> 0 Then
        '"RichEdit20A"
        hTemp = A_CreateWindowEx(0&, WNDCTRL_RICHEDIT_CLASS20A, strCaption, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
        If hTemp <= NUM_ZERO Then Exit Function
        SetCtlDefaultGuiFont hTemp
        CreateRichEdit = hTemp
    'Else
        'Debug.Print "Unable to load library"
    End If
End Function

Public Function CreateAnimation(hParent As Long, X As Long, Y As Long, _
                                Optional bBorder As Boolean = True, _
                                Optional ACS_Flags As Long = ACS_CENTER Or ACS_TRANSPARENT) As Long

    If hParent = 0 Then Exit Function
    
    lStyle = WS_CHILD
    If bBorder Then lStyle = lStyle Or WS_BORDER
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_ANIMATE_CLASS, vbNullString, lStyle, X, Y, 0&, 0&, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    CreateAnimation = hTemp

'This control does not create a separate thread any more??
'ACS_TIMER is obsolete, do not use (MSDN)

'ACS_AUTOPLAY
'Starts playing the animation as soon as the AVI clip is opened.
End Function

'The TTM_ACTIVATE message activates and deactivates a ToolTip control.
Public Function CreateToolTip(Optional hParent As Long = 0, _
                                Optional bBaloon As Boolean = False, _
                                Optional bMultiLine As Boolean = False, _
                                Optional sTitle As String = "", _
                                Optional lIcon As Long = 0, _
                                Optional lForeColor As Long = -1, _
                                Optional lBackColor As Long = -1) As Long
    
    lStyle = WS_POPUP Or TTS_NOPREFIX Or TTS_ALWAYSTIP
    If bBaloon = True Then lStyle = lStyle Or TTS_BALLOON
    
    hTemp = A_CreateWindowEx(WS_EX_TOPMOST, WNDCTRL_TOOLTIPS_CLASS, _
                            vbNullString, lStyle, 0&, 0&, 0&, 0&, _
                            0&, 0&, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    'The window procedure for the ToolTip control automatically sets the size, position, and
    'visibility of the control. The height of the ToolTip window is based on the height of th
    'font currently selected into the device context for the ToolTip control.
    'The width varies based on the length of the string currently in the ToolTip window.
    SetWindowPos hTemp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    ' Set max width
    If bMultiLine = True Then A_SendMessage hTemp, TTM_SETMAXTIPWIDTH, 0&, MAXSHORT   'ByVal &H7FFF&) '32767
'Icon
'   Set wParam to one of the following values to specify the icon to be displayed.
'   As of Microsoft Windows XP Service Pack 2 (SP2) and later,
'   this parameter can also contain an HICON value. Any value greater than 3 is assumed to be an HICON.
'0
'   No icon.
'1
'   Info icon.
'2
'   Warning Icon
'3
'Error Icon
'pszTitle
'   Pointer to the title string. You must assign a value to pszTitle.

'When icon contains an HICON, a copy of the icon is created by the ToolTip window.
'The caller is responsible for freeing that copied icon.
'When calling TTM_SETTITLE, the string pointed to by pszTitle cannot exceed 100 characters.
    If LenB(sTitle) > 0 Then A_SendMessageStr hTemp, TTM_SETTITLE, CLng(lIcon), sTitle
    If lIcon > 0 Then
        'Document states that we need a pointer to a string
        A_SendMessageStr hTemp, TTM_SETTITLE, CLng(lIcon), sTitle
    End If
    If lForeColor > -1 Then A_SendMessage hTemp, TTM_SETTIPTEXTCOLOR, lForeColor, 0&
    If lBackColor > -1 Then A_SendMessage hTemp, TTM_SETTIPBKCOLOR, lBackColor, 0&

    CreateToolTip = hTemp
End Function

'Common control styles which apply to header controls, toolbar controls, and status windows.
'CCS_ADJUSTABLE = Enables a toolbar's built-in customization features, which enable the user to drag a button to a new position or to remove a button by dragging it off the toolbar. In addition, the user can double-click the toolbar to display the Customize Toolbar dialog box, which enables the user to add, delete, and rearrange toolbar buttons.
'CCS_BOTTOM = Causes the control to position itself at the bottom of the parent window's client area and sets the width to be the same as the parent window's width. Status windows have this style by default.
'CCS_LEFT = Version 4.70. Causes the control to be displayed vertically on the left side of the parent window.
'CCS_NODIVIDER = Prevents a two-pixel highlight from being drawn at the top of the control.
'CCS_NOMOVEX = Version 4.70. Causes the control to resize and move itself vertically, but not horizontally, in response to a WM_SIZE message. If CCS_NORESIZE is used, this style does not apply.
'CCS_NOMOVEY = Causes the control to resize and move itself horizontally, but not vertically, in response to a WM_SIZE message. If CCS_NORESIZE is used, this style does not apply. Header windows have this style by default.
'CCS_NOPARENTALIGN = Prevents the control from automatically moving to the top or bottom of the parent window. Instead, the control keeps its position within the parent window despite changes to the size of the parent. If CCS_TOP or CCS_BOTTOM is also used, the height is adjusted to the default, but the position and width remain unchanged.
'CCS_NORESIZE = Prevents the control from using the default width and height when setting its initial size or a new size. Instead, the control uses the width and height specified in the request for creation or sizing.
'CCS_RIGHT = Version 4.70. Causes the control to be displayed vertically on the right side of the parent window.
'CCS_TOP = Causes the control to position itself at the top of the parent window's client area and sets the width to be the same as the parent window's width. Toolbars have this style by default.
'CCS_VERT = Version 4.70. Causes the control to be displayed vertically.

'The toolbar window procedure automatically adjusts the size of the toolbar whenever
'it receives a WM_SIZE or TB_AUTOSIZE message. For example, a TB_SETBUTTONSIZE message.
'The toolbar default sizing and positioning behaviors can be turned off by setting the CCS_NORESIZE
'and CCS_NOPARENTALIGN common control styles. Toolbar controls that are hosted by REBAR controls must
'set these styles because the rebar control sizes and positions the toolbar.
Public Function CreateToolBar(hParent As Long, bHostedByRebar As Boolean, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bBorder As Boolean = False, Optional CCS_Orientation As Long = CCS_NOPARENTALIGN, _
                                Optional TBSTYLE_FlatList As Long = TBSTYLE_FLAT Or TBSTYLE_LIST Or CCS_NODIVIDER, _
                                Optional bDRAWDDARROWS As Boolean = True, Optional bMIXEDBUTTONS As Boolean = True, _
                                Optional bHIDECLIPPEDBUTTONS As Boolean = False, _
                                Optional bToolTips As Boolean = True, _
                                Optional bWRAPABLE As Boolean = False, _
                                Optional ByVal IDB_ImageList As Long = -1) As Long

    If hParent = 0 Then Exit Function
'
'TBSTYLE_LIST: Creates a flat toolbar with button text to the right of the bitmap. TBSTYLE_LIST + TBSTYLE_FLAT = FlatList style,
'TBSTYLE_EX_MIXEDBUTTONS: This style allows you to set text for all buttons, but only display it for those buttons with the BTNS_SHOWTEXT button style.
'TBSTYLE_EX_DRAWDDARROWS: This style allows buttons to have a separate dropdown arrow. Buttons that have the BTNS_DROPDOWN style will be drawn with a drop-down arrow in a separate section, to the right of the button.
'TBSTYLE_TRANSPARENT: Creates a transparent toolbar. In a transparent toolbar,
'           the toolbar is transparent but the buttons are not.
'           Button text appears under button bitmaps. + TBSTYLE_FLAT = FlatTransparent (MS Win style)
    
    'Basic styles
    If bHostedByRebar = True Then lStyle = CCS_NORESIZE
    lStyle = lStyle Or WS_CHILD Or WS_VISIBLE Or CCS_Orientation
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    lStyle = lStyle Or TBSTYLE_FlatList
    If bToolTips Then lStyle = lStyle Or TBSTYLE_TOOLTIPS
    If bWRAPABLE Then lStyle = lStyle Or TBSTYLE_WRAPABLE
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_TOOLBARCLASSNAME, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function

    Dim tBB As TBBUTTON
    'Dim tbAddb As TBADDBITMAP
'
    'Set up extended styles
    lStyle = 0
    If bMIXEDBUTTONS Then lStyle = TBSTYLE_EX_MIXEDBUTTONS
    If bDRAWDDARROWS Then lStyle = lStyle Or TBSTYLE_EX_DRAWDDARROWS
    If bHIDECLIPPEDBUTTONS Then lStyle = lStyle Or TBSTYLE_EX_HIDECLIPPEDBUTTONS
    If lStyle > -1 Then A_SendMessage hTemp, TB_SETEXTENDEDSTYLE, 0&, lStyle
    'Debug.Print "ex: " & A_SendMessage(hTemp, TB_GETEXTENDEDSTYLE, 0&, 0&)
    
    'Use windows provided images, nice and a whole bunch of them
    'IDB_HIST_LARGE_COLOR
    '   Microsoft Windows Explorer bitmaps in large size.
    'IDB_HIST_SMALL_COLOR
    '   Microsoft Windows Explorer bitmaps in small size.
    'IDB_STD_LARGE_COLOR
    '   Standard bitmaps in large size.
    'IDB_STD_SMALL_COLOR
    '   Standard bitmaps in small size.
    'IDB_VIEW_LARGE_COLOR
    '   View bitmaps in large size.
    'IDB_VIEW_SMALL_COLOR
    '   View bitmaps in small size.
    If IDB_ImageList > -1 Then A_SendMessage hTemp, TB_LOADIMAGES, IDB_ImageList, HINST_COMMCTRL

'Second method of adding bitmaps to the toolbar
'        tbAddb.hInst = HINST_COMMCTRL
'        tbAddb.nID = IDB_ImageList
'        'Returns the index of the first new image if successful, or -1 otherwise.
'        A_SendMessageAnyRef hTemp, TB_ADDBITMAP, 0&, tbAddb


    'If an application uses the CreateWindowEx function to create the toolbar,
    'the application must send TB_BUTTONSTRUCTSIZE msg to the toolbar before sending the
    'TB_ADDBITMAP or TB_ADDBUTTONS msgs.
    A_SendMessage hTemp, TB_BUTTONSTRUCTSIZE, Len(tBB), 0&
    'Return hwnd
    CreateToolBar = hTemp
End Function

Public Function CreateReBar(hParent As Long, _
                            Optional CCS_Orientation As Long = CCS_TOP, _
                            Optional WS_EX_flags As Long = 0, Optional lImageList As Long = 0) As Long
'WS_EX_flags default was WS_EX_TOOLWINDOW
'Which I saw in a demo in MSDN, removed it, there was no need for it
    If hParent = 0 Then Exit Function

    'Basic styles
    lStyle = WS_CHILD Or WS_VISIBLE Or WS_BORDER Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or RBS_VARHEIGHT Or RBS_BANDBORDERS Or CCS_Orientation
    
    hTemp = A_CreateWindowEx(WS_EX_flags, WNDCTRL_REBARCLASSNAME, vbNullString, lStyle, 0&, 0&, 0&, 0&, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    Dim rbi As REBARINFO
    'Initialize and send the REBARINFO structure.
    rbi.cbSize = Len(rbi)  'Required when using this structure.
    'fMask
    '   Flag values that describe characteristics of the rebar control.
    '   Currently, rebar controls support only one value:
    'RBIM_IMAGELIST
    '   The himl member is valid or must be filled.
    If lImageList > 0 Then
        rbi.fMask = RBIM_IMAGELIST
    Else
        rbi.fMask = 0
    End If
    rbi.himl = lImageList
    If CBool(A_SendMessageAnyRef(hTemp, RB_SETBARINFO, 0&, rbi)) = False Then Exit Function
    CreateReBar = hTemp
    
'RBS_AUTOSIZE = Version 4.71. The rebar control will automatically change the layout of the bands when the size or position of the control changes. An RBN_AUTOSIZE notification will be sent when this occurs.
'RBS_BANDBORDERS =Version 4.71. The rebar control displays narrow lines to separate adjacent bands.
'RBS_DBLCLKTOGGLE = Version 4.71. The rebar band will toggle its maximized or minimized state when the user double-clicks the band. Without this style, the maximized or minimized state is toggled when the user single-clicks on the band.
'RBS_FIXEDORDER =Version 4.70. The rebar control always displays bands in the same order. You can move bands to different rows, but the band order is static.
'RBS_REGISTERDROP = Version 4.71. The rebar control generates RBN_GETOBJECT notification messages when an object is dragged over a band in the control. To receive the RBN_GETOBJECT notifications, initialize OLE with a call to OleInitialize or CoInitialize.
'RBS_TOOLTIPS = Version 4.71. Not yet supported.
'RBS_VARHEIGHT = Version 4.71. The rebar control displays bands at the minimum required height, when possible. Without this style, the rebar control displays all bands at the same height, using the height of the tallest visible band to determine the height of other bands.
'RBS_VERTICALGRIPPER = Version 4.71. The size grip will be displayed vertically instead of horizontally in a vertical rebar control. This style is ignored for rebar controls that do not have the CCS_VERT style.

End Function

'If you do not set the range values, the system sets the
'minimum value to 0 and the maximum value to 100.
'By default, the step increment is set to 10.
'An application can now control the colors used in a progress bar control with
'the PBM_SETBARCOLOR and PBM_SETBKCOLOR messages.
Public Function CreateProgressBar(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                    Optional bBorder As Boolean = True, Optional lStep As Long = -1, _
                                    Optional lMinValue As Long = -1, Optional lMaxValue As Long = -1, _
                                    Optional lBackColor As Long = -1, Optional lBarColor As Long = -1, _
                                    Optional bSmooth As Boolean = True, Optional bMarquee As Boolean = False, _
                                    Optional bVertical As Boolean = False, _
                                    Optional bTabStop As Boolean = False) As Long
    If hParent = 0 Then Exit Function

    'Basic styles
    lStyle = WS_CHILD Or WS_VISIBLE
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If bTabStop = True Then lStyle = lStyle Or WS_TABSTOP
    If bSmooth = True Then lStyle = lStyle Or PBS_SMOOTH
    If bVertical = True Then lStyle = lStyle Or PBS_VERTICAL
'Sets the progress bar to marquee mode. This causes the progress bar to move like a marquee.
'Use PBM_SETMARQUEE message when you do not know the amount of progress
'toward completion but wish to indicate that progress is being made.
'Send the PBM_SETMARQUEE message to start or stop the animation. Requires XP
    If bMarquee = True Then lStyle = lStyle Or PBS_MARQUEE
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_PROGRESS_CLASS, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    If lMinValue > -1 Or lMaxValue > -1 Then A_SendMessage hTemp, PBM_SETRANGE, 0&, MAKELONG(lMinValue, lMaxValue)
    If lStep > -1 Then A_SendMessage hTemp, PBM_SETSTEP, lStep, 0&
    'Specify the CLR_DEFAULT value to cause the progress bar to use its default background color.
    If lBackColor > -1 Then A_SendMessage hTemp, PBM_SETBKCOLOR, 0&, lBackColor
    If lBarColor > -1 Then A_SendMessage hTemp, PBM_SETBARCOLOR, 0&, lBarColor
    
    CreateProgressBar = hTemp
    'Note:
    '   You use the Theme API to apply visual styles to applications.
    '   If you are using a progress bar control with the Theme API on
    '   Microsoft Windows XP the control must be at least 10 pixels high.
End Function

'Note:
'   Combining the CCS_TOP and SBARS_SIZEGRIP styles is not recommended
'   because the resulting sizing grip is not functional.
'If your application uses a status bar that has only one part,
'you can use the WM_SETTEXT, WM_GETTEXT, and WM_GETTEXTLENGTH messages
'to perform text operations. These messages deal only with the part that
'has an index of zero, allowing you to treat the status bar much like a static text control.
'
'sText = array of strings, if no text then pass "" for that specifiec part
'SBT_TxtDrawingStyle = How to draw each part
'lPartRightEdgeCoord = Right coordinates of each part (rigth/width)
Public Function CreateStatusBar(hParent As Long, lNumOfParts As Long, _
                                sText() As String, _
                                SBT_TxtDrawingStyle() As Long, _
                                lPartRightEdgeCoord() As Long, _
                                Optional X As Long = 0, Optional Y As Long = 0, _
                                Optional Width As Long = 0, Optional Height As Long = 0, _
                                Optional bBorder As Boolean = False, _
                                Optional CCS_Orientation As Long = CCS_BOTTOM, _
                                Optional bSizigGrip As Boolean = False, _
                                Optional bToolTips As Boolean = False, _
                                Optional lBackColor As Long = -1) As Long
    
    Dim rcRect As RECT
    Dim lCount As Long
    
    
    If hParent = 0 Then Exit Function
    'Basic styles
    lStyle = WS_CHILD Or WS_VISIBLE
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If bSizigGrip = True Then lStyle = lStyle Or SBARS_SIZEGRIP
    'SBT_TOOLTIPS:
    '   Version 4.71.Use this style to enable ToolTips.
    'SBARS_TOOLTIPS:
    '   Version 5.80.Identical to SBT_TOOLTIPS. Use this flag for versions 5.00 or later.
    If bToolTips = True Then lStyle = lStyle Or SBARS_TOOLTIPS
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_STATUSCLASSNAMEW, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    'Returns the previous background color, or CLR_DEFAULT if the background color is the default color.
    If lBackColor > -1 Then A_SendMessage hTemp, SB_SETBKCOLOR, 0&, lBackColor
    
''If you just want the status width be divided into even parts
'    If CCS_Orientation = CCS_NOPARENTALIGN Then
'        SetRect rcRect, x, Y, Width, Height
'    Else
'    'Otherwise get client rect
'        GetClientRect hParent, rcRect
'    End If
'    'Initialize array of parts
'    Dim arrParts() as long
'    Dim nWidth As Long
'    ReDim arrParts(lNumOfParts - 1)
'    'Calculate the right edge coordinate for each part, and
'    'copy the coordinates to the array.
'    'If an element is -1, the right edge of the corresponding part extends to the border of the window.
'    nWidth = rcRect.Right / lNumOfParts
'    For lCount = 0 To lNumOfParts - 1
'        arrParts(lCount) = nWidth
'        nWidth = nWidth + nWidth
'    Next
    
    'Tell the status bar to create the window parts. Pass -1 to the last part to fill in the rest
    A_SendMessageAnyRef hTemp, SB_SETPARTS, lNumOfParts, lPartRightEdgeCoord(0)

'<wParam can be>
'iPart:
'   Zero-based index of the part to set. If this parameter is set to SB_SIMPLEID,
'   the status window is assumed to be a simple window with only one part.
'Or
'uType:
'   Type of drawing operation. This parameter can be one of the following values:
'0:
'   The text is drawn with a border to appear lower than the plane of the window.
'SBT_NOBORDERS:
'   The text is drawn without borders.
'SBT_OWNERDRAW:
'   The text is drawn by the parent window.
'SBT_POPOUT:
'   The text is drawn with a border to appear higher than the plane of the window.
'SBT_RTLREADING:
'   The text will be displayed in the opposite direction to the text in the parent window.
    For lCount = 0 To UBound(sText)
        If LenB(sText(lCount)) > 0 Then
            A_SendMessageStr hTemp, SB_SETTEXT, lCount Or SBT_TxtDrawingStyle(lCount), sText(lCount)
        End If
    Next
    
    CreateStatusBar = hTemp
End Function

Public Function CreateStatusBarSimple(hParent As Long, Optional sText As String = "", _
                                Optional X As Long = 0, Optional Y As Long = 0, _
                                Optional Width As Long = 0, Optional Height As Long = 0, _
                                Optional bBorder As Boolean = False, _
                                Optional CCS_Orientation As Long = CCS_BOTTOM, _
                                Optional SBT_TxtDrawingStyle As Long = SBT_POPOUT, _
                                Optional bSizigGrip As Boolean = False, _
                                Optional bToolTips As Boolean = False, _
                                Optional lBackColor As Long = -1) As Long
    If hParent = 0 Then Exit Function
    'Basic styles
    lStyle = WS_CHILD Or WS_VISIBLE
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    If bSizigGrip = True Then lStyle = lStyle Or SBARS_SIZEGRIP
    If bToolTips = True Then lStyle = lStyle Or SBARS_TOOLTIPS
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_STATUSCLASSNAMEW, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    'Set as simple
    A_SendMessage hTemp, SB_SIMPLE, 0&, 0&
    'Returns the previous background color, or CLR_DEFAULT if the background color is the default color.
    If lBackColor > -1 Then A_SendMessage hTemp, SB_SETBKCOLOR, 0&, lBackColor

    'If simple status, then we treat it as any static control (label)
    'unless we want to apply a type
    'A_SetWindowText hTemp, sText
    If LenB(sText) > 0 Then A_SendMessageStr hTemp, SB_SETTEXT, SBT_TxtDrawingStyle, sText
    
    CreateStatusBarSimple = hTemp
End Function

'Note that you cannot set a global hot key for a window that has the WS_CHILD window style.
Public Function CreateHotKey(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long) As Long
    
    'Basic styles
    lStyle = WS_CHILD Or WS_VISIBLE
    hTemp = A_CreateWindowEx(0&, WNDCTRL_HOTKEY_CLASS, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    SetFocusApi hTemp
    
'    Set rules for invalid key combinations. If the user does not supply a
'    modifier key, use ALT as a modifier. If the user supplies SHIFT as a
'    modifier key, use SHIFT + ALT instead.

'    HKCOMB_NONE or HKCOMB_S = invalid key combinations
'    MAKELong(HOTKEYF_ALT, 0) = add ALT to invalid entries
    A_SendMessage hTemp, HKM_SETRULES, HKCOMB_NONE Or HKCOMB_S, MAKELONG(HOTKEYF_ALT, 0)

'    Set CTRL + ALT + A as the default hot key for this window.
'    0x41 is the virtual key code for 'A'.
    A_SendMessage hTemp, HKM_SETHOTKEY, MAKELONG(VK_A, HOTKEYF_CONTROL Or HOTKEYF_ALT), 0&

    CreateHotKey = hTemp
End Function

Public Function CreateTabCtl(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bToolTips As Boolean = True, _
                                Optional bButtons As Boolean = False, _
                                Optional bFlatButtons As Boolean = False, _
                                Optional bFlatSeparators As Boolean = False, _
                                Optional bMultiSelect As Boolean = False, _
                                Optional bMultiLine As Boolean = False, _
                                Optional bRightJustified As Boolean = False, _
                                Optional bScrollOpposite As Boolean = False, _
                                Optional bHotTrack As Boolean = False, _
                                Optional bFocusOnButtonDown As Boolean = True, _
                                Optional bFixedWidth As Boolean = True, _
                                Optional bForceLabelLeft As Boolean = False, _
                                Optional bForceIconLeft As Boolean = False, _
                                Optional lMinTabWidth As Long = 70) As Long
    'Basic styles
    lStyle = WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or TCS_RAGGEDRIGHT  'WS_CLIPSIBLINGS
    If bButtons = True Then
        lStyle = lStyle Or TCS_BUTTONS
        'Only valid if style is Buttons
        If bFlatButtons = True Then lStyle = lStyle Or TCS_FLATBUTTONS
        If bMultiSelect = True Then lStyle = lStyle Or TCS_MULTISELECT
    Else
        lStyle = lStyle Or TCS_TABS
    End If
    'cannot be combined with the TCS_RIGHTJUSTIFY style.
    If bFixedWidth = True Then
        lStyle = lStyle Or TCS_FIXEDWIDTH
        If bForceLabelLeft = True Then lStyle = lStyle Or TCS_FORCELABELLEFT
        If bForceIconLeft = True Then lStyle = lStyle Or TCS_FORCEICONLEFT
    End If
    If bMultiLine = True Then
        lStyle = lStyle Or TCS_MULTILINE
        'The width of each tab is increased, if necessary,
        'so that each row of tabs fills the entire width of the tab control.
        'This window style is ignored unless the TCS_MULTILINE style is also specified.
        If bRightJustified = True And bFixedWidth = False Then lStyle = lStyle Or TCS_RIGHTJUSTIFY
        If bScrollOpposite = True Then lStyle = lStyle Or TCS_SCROLLOPPOSITE
    Else
        lStyle = lStyle Or TCS_SINGLELINE
    End If
    If bHotTrack = True Then lStyle = lStyle Or TCS_HOTTRACK
    If bFocusOnButtonDown = True Then
        lStyle = lStyle Or TCS_FOCUSONBUTTONDOWN
    Else
        lStyle = lStyle Or TCS_FOCUSNEVER
    End If
    If bToolTips = True Then lStyle = lStyle Or TCS_TOOLTIPS

'Styles that are not supported if you use ComCtl32.dll version 6.
'   TCS_VERTICAL
'   TCS_BOTTOM
'Styles that are not supported if you use visual styles.
'   TCS_RIGHT
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_TABCONTROLA, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    If lMinTabWidth > 0 Then
        A_SendMessage hTemp, TCM_SETMINTABWIDTH, 0&, lMinTabWidth
    End If
    If bButtons = True Then
        'The tab control will draw separators between the tab items.
        'This extended style only affects tab controls that have the
        'TCS_BUTTONS and TCS_FLATBUTTONS styles. By default,
        'creating the tab control with the TCS_FLATBUTTONS style sets this extended style.
        'If you do not require separators, you should remove this extended style after creating the control.
        If bFlatSeparators = False Then
            'Clear Ex style
            A_SendMessage hTemp, TCM_SETEXTENDEDSTYLE, TCS_EX_FLATSEPARATORS, 0&
        End If
    End If
    SetCtlDefaultGuiFont hTemp
    CreateTabCtl = hTemp
End Function

'To get a flat border pass 0 to WS_EX_flags
'To have subitems display images add LVS_EX_SUBITEMIMAGES flag to LVS_EX_Styles
'To enable hottracking add LVS_EX_TRACKSELECT to LVS_EX_Styles
'To get a flat border remove WS_EX_CLIENTEDGE from WS_EX_flags
'To get auto label edit add LVS_EDITLABELS to LVS_Styles flags
Public Function CreateListView(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bBorder As Boolean = True, Optional WS_EX_flags As Long = WS_EX_CLIENTEDGE, _
                                Optional LVS_Styles As Long = LVS_AUTOARRANGE Or LVS_REPORT, _
                                Optional LVS_EX_Styles As Long = LVS_EX_HEADERDRAGDROP Or LVS_EX_INFOTIP) As Long
    
    lStyle = WS_CHILD Or WS_VISIBLE Or WS_TABSTOP
    If bBorder = True Then lStyle = lStyle Or WS_BORDER
    lStyle = lStyle Or LVS_Styles
    
    hTemp = A_CreateWindowEx(WS_EX_flags, WNDCTRL_LISTVIEW, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    A_SendMessage hTemp, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_Styles, LVS_EX_Styles
    SetCtlDefaultGuiFont hTemp
    CreateListView = hTemp
End Function

'Or TBS_ENABLESELRANGE
Public Function CreateTrackBar(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional TBS_flags As Long = TBS_AUTOTICKS, _
                                Optional bToolTips As Boolean = True, _
                                Optional TBTS_tipDisplayFlag As TrackBarTipSideFlags = TBTS_TOP, _
                                Optional lMinRange As Long = 0, _
                                Optional lMaxRange As Long = 100, _
                                Optional lLineSize As Long = 1, _
                                Optional lPageSize As Long = 4, _
                                Optional lFrequency As Long = 1, _
                                Optional lBuddy As Long = -1, _
                                Optional bAlignBuddyRight As Boolean = True) As Long
    
    lStyle = WS_CHILD Or WS_VISIBLE Or TBS_flags
    If bToolTips Then lStyle = lStyle Or TBS_TOOLTIPS
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_TRACKBAR_CLASS, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    
    'Set range
    TrackBar_SetRange hTemp, lMinRange, lMaxRange
    'Set linesize, default is 1
    If lLineSize <> 1 Then TrackBar_SetLineSize hTemp, lLineSize
    'Set pagesize
    TrackBar_SetPageSize hTemp, lPageSize
    'set initial pos
    TrackBar_SetPos hTemp, lMinRange
    'Set frquency, default is 1
    If lFrequency <> 1 Then A_SendMessage hTemp, TBM_SETTICFREQ, lFrequency, 0&
    'Set buddy (Can have two buddy
    If lBuddy > -1 Then
        A_SendMessage hTemp, TBM_SETBUDDY, CLng(bAlignBuddyRight), lBuddy
    End If
    If bToolTips Then A_SendMessage hTemp, TBM_SETTIPSIDE, TBTS_tipDisplayFlag, 0&
    CreateTrackBar = hTemp
End Function

'LBS_STANDARD:
'   Sorts strings in the list box alphabetically.
'   The parent window receives an input message whenever
'   the user clicks or double-clicks a string. The list box has borders on all sides.
'LBS_NOINTEGRALHEIGHT:
'   Specifies that the size of the list box is exactly the size specified by the application
'   when it created the list box. Normally, the system sizes a list box so that the list box
'   does not display partial items.
Public Function CreateListBox(hParent As Long, X As Long, Y As Long, Width As Long, Height As Long, _
                                Optional bSTANDARD As Boolean = True, _
                                Optional bBorder As Boolean = False, _
                                Optional bSCROLLBARS As Boolean = False, _
                                Optional bDISABLENOSCROLL As Boolean = False, _
                                Optional bHASSTRINGS As Boolean = False, _
                                Optional bMULTIPLESEL As Boolean = False, _
                                Optional bEXTENDEDSEL As Boolean = False, _
                                Optional bMULTICOLUMN As Boolean = False, _
                                Optional bNOINTEGRALHEIGHT As Boolean = False, _
                                Optional bNOSEL As Boolean = False, _
                                Optional bNOTIFY As Boolean = False, _
                                Optional bSORT As Boolean = False, _
                                Optional bUSETABSTOPS As Boolean = True, _
                                Optional bWANTKEYBOARDINPUT As Boolean = False, _
                                Optional bOWNERDRAW As Boolean = False, _
                                Optional bOD_VARIABLE_HEIGHT As Boolean = False) As Long
    'Start with basic style
    lStyle = WS_CHILD Or WS_VISIBLE
    
    If bSTANDARD Then lStyle = lStyle Or LBS_STANDARD
    If bSCROLLBARS Then lStyle = lStyle Or WS_VSCROLL Or WS_HSCROLL
    If bBorder Then lStyle = lStyle Or WS_BORDER
    If bDISABLENOSCROLL Then lStyle = lStyle Or LBS_DISABLENOSCROLL
    If bHASSTRINGS Then lStyle = lStyle Or LBS_HASSTRINGS
    If bMULTIPLESEL Then lStyle = lStyle Or LBS_MULTIPLESEL
    If bEXTENDEDSEL Then lStyle = lStyle Or LBS_EXTENDEDSEL
    If bMULTICOLUMN Then lStyle = lStyle Or LBS_MULTICOLUMN
    If bNOINTEGRALHEIGHT Then lStyle = lStyle Or LBS_NOINTEGRALHEIGHT
    If bNOSEL Then lStyle = lStyle Or LBS_NOSEL
    If bNOTIFY Then lStyle = lStyle Or LBS_NOTIFY
    If bSORT Then lStyle = lStyle Or LBS_SORT
    If bUSETABSTOPS Then lStyle = lStyle Or LBS_USETABSTOPS
    If bOWNERDRAW Then
        lStyle = lStyle Or LBS_STANDARD
        If bOD_VARIABLE_HEIGHT Then
            lStyle = lStyle Or LBS_OWNERDRAWVARIABLE Or WS_HSCROLL
        Else
            lStyle = lStyle Or LBS_OWNERDRAWFIXED Or WS_HSCROLL
        End If
    End If
    
    hTemp = A_CreateWindowEx(0&, WNDCTRL_LISTBOX, vbNullString, lStyle, X, Y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= NUM_ZERO Then Exit Function
    SetCtlDefaultGuiFont hTemp
    CreateListBox = hTemp
End Function

''Say bye bye to flat scrollbars in Comctl32 version 6.0 (XP)
''Note  Flat scroll bar APIs are implemented in Comctl32.dll versions 4.71 through 5.82.
'Comctl32.dll versions 6.00 and higher do not support flat scroll bars.
'Public Function CreateFlatScrollBar(hParent As Long, x As Long, y As Long, Width As Long, Height As Long) As Long
'End Function


