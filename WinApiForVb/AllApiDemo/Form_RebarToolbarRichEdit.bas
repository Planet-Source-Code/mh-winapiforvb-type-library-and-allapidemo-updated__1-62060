Attribute VB_Name = "Form_RebarToolbarRichEdit"
Option Explicit

'Form rebar(coolbar) + toolbar + richedit
Public hFrmToolBar As Long
Public hRebar As Long
Public hToolbar As Long
Public sToolbarTip As String
Public hRebarCombo As Long
Public hRichEdit As Long
Public hMenuRichEdit As Long

Public Sub BtnShowToolbar_Click()
    If hFrmToolBar = 0 Then
        hFrmToolBar = CreateForm(300, 300, 500, 250, AddressOf ToolBarWndProc, hMainForm, , , "FormToolbar", "Rebar Toolbar RichEdit2")
        If hFrmToolBar > 0 Then
            hRebar = CreateReBar(hFrmToolBar)
            If hRebar > 0 Then
                hToolbar = Comctl32_Controls.CreateToolBar(hRebar, True, 0, 0, 100, 35, , , , , , , , , IDB_STD_SMALL_COLOR)
                If hToolbar > 0 Then
                    'Add buttons
                    Toolbar_AddButton hToolbar, 1000, STD_FILENEW, "New", , BTNS_BUTTON Or BTNS_AUTOSIZE, , True, BTNS_WHOLEDROPDOWN     ', "New"
                    Toolbar_AddButton hToolbar, 1001, STD_FILEOPEN, "Open"
                    Toolbar_AddButton hToolbar, 1002, STD_FILESAVE, "Save"
                    Toolbar_AddButton hToolbar, 1003, STD_CUT, "Cut"
                    Toolbar_AddButton hToolbar, 1004, STD_COPY, "Copy"
                    Toolbar_AddButton hToolbar, 1005, STD_PASTE, "Paste"
                    Toolbar_AddButton hToolbar, 1006, , "", TBSTATE_ENABLED, BTNS_SEP
                    'Create a group
                    Toolbar_AddButton hToolbar, 1007, STD_PRINT, "", TBSTATE_ENABLED Or TBSTATE_CHECKED, BTNS_CHECKGROUP
                    Toolbar_AddButton hToolbar, 1008, STD_DELETE, "", TBSTATE_ENABLED, BTNS_CHECKGROUP
                    Toolbar_AddButton hToolbar, 1009, STD_FIND, "", TBSTATE_ENABLED, BTNS_CHECKGROUP
                    'A_SendMessage hToolbar, TB_SETSTATE, 1009, MAKELONG(TBSTATE_ENABLED Or TBSTATE_CHECKED, 0&)
                    Toolbar_AddButton hToolbar, 1010, , "", TBSTATE_ENABLED, BTNS_SEP
                    Toolbar_AddButton hToolbar, 1011, STD_HELP, "", TBSTATE_ENABLED Or TBSTATE_CHECKED, BTNS_CHECK
                    'you must send the TB_AUTOSIZE message after all the items and strings have been inserted into the control to cause the toolbar to recalculate its size based on its content.
                    A_SendMessage hToolbar, TB_AUTOSIZE, 0, 0
                    'Add toolbar and a combo to rebar bands
                    Dim dwBtnSize As Long
                    Dim rc As RECT
                    'Get the height of the toolbar.
                    dwBtnSize = A_SendMessage(hToolbar, TB_GETBUTTONSIZE, 0&, 0&)
                    'Add toolbar to rebar
                    Rebar_InsertBand hRebar, hToolbar, 0, HiWord(dwBtnSize), 100
                    
                    'Create combo
                    hRebarCombo = CreateCombo(hRebar, "Item one", 0, 0, 100, 250)
                    If hRebarCombo > 0 Then
                        ComboBox_AddString hRebarCombo, "Item one"
                        ComboBox_AddString hRebarCombo, "Item Two"
                        ComboBox_AddString hRebarCombo, "Item Three", 1
                        A_SendMessage hRebarCombo, CB_SETCURSEL, 0&, 0&
                        GetWindowRect hRebarCombo, rc
                        'Add combo to the rebar
                        Rebar_InsertBand hRebar, hRebarCombo, 100, rc.Bottom - rc.Top, 100, True, , , , "Address:"
                    End If
                    ShowWindow hToolbar, SW_SHOW
                End If 'End hToolbar
                
                'Create rich edit 2.0
                hRichEdit = CreateRichEdit(hFrmToolBar, "", 0, 50, 490, 250 - dwBtnSize)   ', , , , , , , , True)
                If hRichEdit > 0 Then
                    'Make richedit URL aware
                    A_SendMessage hRichEdit, EM_AUTOURLDETECT, 1&, 0&
                    'If we are not using RTF, we can simply use the textbox macros
                    'to work with the richedit
                    TextBox_SetText hRichEdit, "Sample Text, right click for context menu" & vbCrLf & "http://www.google.com"
                    'Register messages that we want to receive from richedit
                    'The default event mask (before any is set) is ENM_NONE (no msgs are send).
                    A_SendMessage hRichEdit, EM_SETEVENTMASK, 0&, ENM_MOUSEEVENTS
                    'A_SendMessage hRichEdit, EM_SETEVENTMASK, 0&, ENM_KEYEVENTS or ENM_MOUSEEVENTS or ENM_SCROLLEVENTS
                    'The EN_MSGFILTER message notifies a rich edit control's parent window of a keyboard or mouse event in the control.
                    'A rich edit control sends this notification message in the form of a WM_NOTIFY message.
                    'To receive EN_MSGFILTER notifications for events, specify one or more of the following flags in the mask sent
                    'with the EM_SETEVENTMASK message.
                End If
                'Set up context menu for RichEdit
                hMenuRichEdit = CreatePopupMenu()
                If hMenuRichEdit > 0 Then
                    Dim mnuItem As MENUITEMINFO
                    Dim lSubMenu As Long

                    mnuItem = CreateMenuItem("Redo", 1040, 0&)
                    A_InsertMenuItem hMenuRichEdit, 1040, 0, mnuItem
                    
                    mnuItem = CreateMenuItem("Undo", 1041, 0&)
                    A_InsertMenuItem hMenuRichEdit, 1041, 0, mnuItem
                    
                    mnuItem = CreateMenuItem("-", 1042, 0&, True)
                    A_InsertMenuItem hMenuRichEdit, 1042, 0, mnuItem
                    
                    mnuItem = CreateMenuItem("Cut", 1043, 0&)
                    A_InsertMenuItem hMenuRichEdit, 1043, 0, mnuItem
                    
                    mnuItem = CreateMenuItem("Copy", 1044, 0&)
                    A_InsertMenuItem hMenuRichEdit, 1044, 0, mnuItem
                    
                    mnuItem = CreateMenuItem("Paste", 1045, 0&)
                    A_InsertMenuItem hMenuRichEdit, 1045, 0, mnuItem
                    
                    mnuItem = CreateMenuItem("Delete", 1046, 0&)
                    A_InsertMenuItem hMenuRichEdit, 1046, 0, mnuItem
                    
                    mnuItem = CreateMenuItem("-", 1047, 0&, True)
                    A_InsertMenuItem hMenuRichEdit, 1047, 0, mnuItem
                    
                    mnuItem = CreateMenuItem("SelectAll", 1048, 0&)
                    A_InsertMenuItem hMenuRichEdit, 1048, 0, mnuItem
                    
                    mnuItem = CreateMenuItem("-", 1049, 0&, True)
                    A_InsertMenuItem hMenuRichEdit, 1049, 0, mnuItem
                                        
                    lSubMenu = CreatePopupMenu
                    If lSubMenu > 0 Then
                        mnuItem = CreateMenuItem("Bold", 1050, 0&)
                        A_InsertMenuItem lSubMenu, 1050, 0, mnuItem
                        mnuItem = CreateMenuItem("Italic", 1051, 0&)
                        A_InsertMenuItem lSubMenu, 1051, 0, mnuItem
                        mnuItem = CreateMenuItem("Strike Through", 1052, 0&)
                        A_InsertMenuItem lSubMenu, 1052, 0, mnuItem
                        mnuItem = CreateMenuItem("Underline", 1053, 0&)
                        A_InsertMenuItem lSubMenu, 1053, 0, mnuItem
                        mnuItem = CreateMenuItem("-", 1054, 0&, True)
                        A_InsertMenuItem lSubMenu, 1054, 0, mnuItem
                        mnuItem = CreateMenuItem("Replace with Combo Text", 1055, 0&)
                        A_InsertMenuItem lSubMenu, 1055, 0, mnuItem
                    End If
                    mnuItem = CreateMenuItem("Selection/Current Word", 1056, lSubMenu)
                    A_InsertMenuItem hMenuRichEdit, 1056, 0, mnuItem
                End If
            End If 'End Rebar
            
            ShowWndForFirstTime hFrmToolBar
        End If
    Else
        ShowWindow hFrmToolBar, SW_SHOW
    End If
End Sub

Public Function ToolBarWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Dim nmh As NMHDR
    Dim dwBtnSize As Long
    
    Select Case uMsg
        Case WM_SIZE
            If wParam <> SIZE_MINIMIZED Then
                Dim lRect As RECT
                'Get form's current dimensions.
                GetWindowRect hFrmToolBar, lRect
                dwBtnSize = A_SendMessage(hRebar, RB_GETBARHEIGHT, 0&, 0&) + 5
                SetWindowPos hRebar, 0, 0, 0, lRect.Right, dwBtnSize, SWP_NOZORDER Or SWP_SHOWWINDOW
                SetWindowPos hRichEdit, 0, 0, dwBtnSize, lRect.Right, lRect.Bottom - dwBtnSize, SWP_NOZORDER Or SWP_SHOWWINDOW
            End If
'        'wParam
'        '   The high-order word specifies the notification code if the message is from a control.
'        '   If the message is from an accelerator, this value is 1.
'        '   If the message is from a menu, this value is zero.
'        '   The low-order word specifies the identifier of the menu item, control, or accelerator.
'        'lParam
'        '   Handle to the control sending the message if the message is from a control. Otherwise, this parameter is NULL.
'        Case WM_COMMAND
'            lRet = HiWord(wParam) 'Notification Msg
'            If lRet = CBN_SELCHANGE Then
'                'Debug.Print "=" & ComboBox_GetSelectedText(hRebarCombo)
'            ElseIf lRet = BN_CLICKED Then
'                If lParam = hToolbar Then
'                    Debug.Print "BT_CLICK1, idCommand:" & LoWord(wParam)
'                    If LoWord(wParam) = 1004 Then
'                        'Debug.Print "BT_CLICK2, idCommand:" & LoWord(wParam)
'                    End If
'                End If
'            End If
        Case WM_NOTIFY
            CopyMemory nmh, ByVal lParam, Len(nmh)
            lRet = nmh.code
            Select Case lRet
                'TB_COMMANDTOINDEX idCommand to index
                Case TBN_GETINFOTIP
                    Dim nmtb As NMTBGETINFOTIP
                    sToolbarTip = "Toolbar idCommand: "
                    CopyMemory nmtb, ByVal lParam, Len(nmtb)
                    'Debug.Print "idCommand: " & nmtb.iItem '& " " & nmtb.hdr.hwndFrom
                    'Need to add a Null char at the end to signal the end of string
                    'the function stops at the first null char found
                    sToolbarTip = sToolbarTip & CStr(nmtb.iItem) & vbNullChar
                    nmtb.cchTextMax = Len(sToolbarTip)
                    'can not be larger than INFOTIPSIZE 1024 characters
                    'Copies only the characters, not embeded nulls, that's why we need a null char at the end
                    A_lstrcpynStrPtr ByVal nmtb.pszText, sToolbarTip, Len(sToolbarTip)
                    Exit Function
                Case TBN_DROPDOWN
                    'only the hdr and iItem members of this structure are valid
                    Dim nmtoolb As NMTOOLBAR
                    CopyMemory nmtoolb, ByVal lParam, Len(nmtoolb)
                    'ToolbarDropdownReturnValues
                    'TBDDRET_DEFAULT The drop-down was handled.
                    'TBDDRET_NODEFAULT The drop-down was not handled.
                    'TBDDRET_TREATPRESSED The drop-down was handled, but treat the button like a regular button.
                    'Debug.Print "idCommand: " & nmtoolb.iItem
                    If nmtoolb.iItem = 1000 Then
                        'Clears without the ability to undo
                        TextBox_Clear hRichEdit
                    End If
                    ToolBarWndProc = TBDDRET_DEFAULT
                    Exit Function
                Case RBN_HEIGHTCHANGE
                    'Adjust richedit pos
                    GetWindowRect hFrmToolBar, lRect
                    dwBtnSize = A_SendMessage(hRebar, RB_GETBARHEIGHT, 0&, 0&) + 5
                    SetWindowPos hRichEdit, 0, 0, dwBtnSize, lRect.Right, lRect.Bottom - dwBtnSize, SWP_NOZORDER Or SWP_SHOWWINDOW
                Case EN_MSGFILTER
                    'Mouse event notification of RichEdit, only interested in RButtonup for displaying contextmenu
                    Dim msgf As MSGFILTER
                    Dim lMenuID As Long, lByteLen As Long, lNumChars As Long
                    Dim pt As POINTAPI
                    Dim emst As SETTEXTEX
                    Dim chrRange As CHARRANGE
                    Dim ccf As A_CHARFORMAT
                    Dim sSelectedText As String

                    CopyMemory msgf, ByVal lParam, Len(msgf)
                    If msgf.MSG = WM_RBUTTONUP Then
                        'Set up menus - First disable all menu items
                        EnableMenuItem hMenuRichEdit, 1040, MF_BYCOMMAND Or MF_GRAYED 'disabled and grayed
                        EnableMenuItem hMenuRichEdit, 1041, MF_BYCOMMAND Or MF_GRAYED
                        EnableMenuItem hMenuRichEdit, 1043, MF_BYCOMMAND Or MF_GRAYED
                        EnableMenuItem hMenuRichEdit, 1044, MF_BYCOMMAND Or MF_GRAYED
                        EnableMenuItem hMenuRichEdit, 1045, MF_BYCOMMAND Or MF_GRAYED
                        EnableMenuItem hMenuRichEdit, 1046, MF_BYCOMMAND Or MF_GRAYED
                        EnableMenuItem hMenuRichEdit, 1048, MF_BYCOMMAND Or MF_GRAYED
                        'EnableMenuItem hMenuRichEdit, 1056, MF_BYCOMMAND Or MF_GRAYED
                        
                        If RichEdit_CanRedo(hRichEdit) Then
                            EnableMenuItem hMenuRichEdit, 1040, MF_BYCOMMAND Or MF_ENABLED
                        End If
                        If TextBox_CanUndo(hRichEdit) Then 'Common in both textbox and richedit
                            EnableMenuItem hMenuRichEdit, 1041, MF_BYCOMMAND Or MF_ENABLED
                        End If
                        
                        'Do we have any text
                        lByteLen = RichEdit_GetTextLengthEx(hRichEdit)
                        If lByteLen <> E_INVALIDARG Then
                            EnableMenuItem hMenuRichEdit, 1048, MF_BYCOMMAND Or MF_ENABLED
                        End If
                        
                        'Get selected range if any
                        RichEdit_GetSelRange hRichEdit, chrRange
                        
                        'If the cpMin and cpMax members are equal, the range is empty.
                        'The range includes everything if cpMin is 0 and cpMax is —1.
                        If chrRange.cpMax <> chrRange.cpMin Then
                            EnableMenuItem hMenuRichEdit, 1043, MF_BYCOMMAND Or MF_ENABLED
                            EnableMenuItem hMenuRichEdit, 1044, MF_BYCOMMAND Or MF_ENABLED
                            EnableMenuItem hMenuRichEdit, 1046, MF_BYCOMMAND Or MF_ENABLED
                            'EnableMenuItem hMenuRichEdit, 1056, MF_BYCOMMAND Or MF_ENABLED
                        End If
                    
                        'wParam = Specifies the Clipboard Formats to try.
                        '         To try any format currently on the clipboard,
                        '         set this parameter to zero.
                        If RichEdit_CanPaste(hRichEdit) Then
                            EnableMenuItem hMenuRichEdit, 1045, MF_BYCOMMAND Or MF_ENABLED
                        End If

                        GetCursorPos pt
                        'This call is recommanded by MSDN due to a bug with TrackPopupMenuEx
                        SetForegroundWindow hFrmToolBar
                        'Returns the ID of the menu clicked or 0
                        lMenuID = TrackPopupMenuEx(hMenuRichEdit, TPM_RETURNCMD Or TPM_LEFTALIGN, pt.X, pt.Y, hFrmToolBar, ByVal 0&)
                        'Continuing with bug related to TrackPopupMenuEx
                        A_PostMessage hFrmToolBar, WM_NULL, 0&, 0&
                        
                        'Go through user selection
                        If lMenuID <> 0 Then
                            ccf.cbSize = Len(ccf)
                            Select Case lMenuID
                                Case 1040 'Redo
                                    RichEdit_Redo hRichEdit
                                Case 1041 'Undo
                                    TextBox_Undo hRichEdit
                                Case 1043 'Cut
                                    TextBox_Cut hRichEdit
                                Case 1044 'Copy
                                    TextBox_Copy hRichEdit
                                Case 1045 'Paste
                                    TextBox_Paste hRichEdit
                                Case 1046 'Delete
                                    TextBox_Delete hRichEdit
                                Case 1048 'SelectAll
                                    TextBox_SetSel hRichEdit
                                Case 1050 'Bold
                                    'Sending this msg twice will undo the first formating (Like a toggle switch)
                                    'Also would be proper to get the formating first to make sure we don't
                                    'loose any formating we had from before unless otherwise
                                    ccf.dwMask = CFM_BOLD
                                    ccf.dwEffects = CFE_BOLD
                                    A_SendMessageAnyRef hRichEdit, EM_SETCHARFORMAT, SCF_WORD Or SCF_SELECTION, ccf
                                Case 1051 'Italic
                                    ccf.dwMask = CFM_ITALIC
                                    ccf.dwEffects = CFE_ITALIC
                                    A_SendMessageAnyRef hRichEdit, EM_SETCHARFORMAT, SCF_WORD Or SCF_SELECTION, ccf
                                Case 1052 'Strike Through
                                    ccf.dwMask = CFM_STRIKEOUT
                                    ccf.dwEffects = CFE_STRIKEOUT
                                    A_SendMessageAnyRef hRichEdit, EM_SETCHARFORMAT, SCF_WORD Or SCF_SELECTION, ccf
                                Case 1053 'Underline
                                    ccf.dwMask = CFM_UNDERLINE
                                    ccf.dwEffects = CFE_UNDERLINE
                                    A_SendMessageAnyRef hRichEdit, EM_SETCHARFORMAT, SCF_WORD Or SCF_SELECTION, ccf
                                Case 1055 'Replace
                                    'Can use
                                    emst.codepage = 1200 'Unicode
                                    emst.flags = ST_KEEPUNDO Or ST_SELECTION 'ST_DEFAULT
                                    sSelectedText = ComboBox_GetSelectedText(hRebarCombo)
                                    'The EM_SETTEXTEX message combines the functionality of WM_SETTEXT and EM_REPLACESEL
                                    'and adds the ability to set text using a code page and to use either
                                    'Rich Text Format (RTF) rich text or plain text.
                                    A_SendMessageAnyAnyRef hRichEdit, EM_SETTEXTEX, emst, ByVal StrPtr(sSelectedText)
                            End Select
                        End If
                    End If
'                    Case TBN_HOTITEMCHANGE 'Hits before TBN_GETINFOTIP
'                        Dim nmtbh As NMTBHOTITEM
'                        CopyMemory nmtbh, ByVal lParam, Len(nmtbh)
'                        Debug.Print "Old:" & nmtbh.idOld & " New:" & nmtbh.idNew
'                        'Return zero to allow the item to be highlighted, or
'                        'nonzero to prevent the item from being highlighted.
'                        Exit Function
            End Select
        Case WM_CLOSE
            ShowWindow hFrmToolBar, SW_HIDE
            'DestroyWindow hFrmToolBar
            Exit Function
    End Select
    ToolBarWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
