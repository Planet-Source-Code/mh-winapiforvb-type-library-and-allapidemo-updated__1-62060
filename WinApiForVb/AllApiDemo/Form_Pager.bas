Attribute VB_Name = "Form_Pager"
Option Explicit

'A pager control is a window container that is used with a window that
'does not have enough display area to show all of its content.
'The pager control allows the user to scroll to the area of the
'window that is not currently in view.
Public hFrmPager As Long
Public hPagerCtl As Long
Public hPagerToolbar As Long    'Toolbar contained within pager

Public Sub BtnShowPageScroller_Click()
    On Error GoTo BtnShowPageScroller_Click_Error

    If hFrmPager = 0 Then
        hFrmPager = CreateForm(300, 300, 300, 100, AddressOf PagerCtlWndProc, hMainForm, WS_SYSMENU Or WS_BORDER Or WS_CAPTION, WS_EX_TOOLWINDOW, "FormPagerControl", "Pager Control")
        If hFrmPager > 0 Then
            'Create pager
            hPagerCtl = CreatePageScroller(hFrmPager, 0, 0, 290, 30, , False)
            
            'Toolbar
            'Need to set bHostedByRebar parameter to true so the function
            'adds CCS_NORESIZE style to prevent the control from attempting
            'to resize itself to the pager control's size
            hPagerToolbar = Comctl32_Controls.CreateToolBar(hPagerCtl, True, 0, 0, 400, 30, , , , , , , False, , IDB_STD_SMALL_COLOR)

            'Add buttons
            Toolbar_AddButton hPagerToolbar, 2000, STD_FILENEW, "New", , BTNS_BUTTON Or BTNS_AUTOSIZE, , True, BTNS_WHOLEDROPDOWN     ', "New"
            Toolbar_AddButton hPagerToolbar, 2001, STD_FILEOPEN, "Open"
            Toolbar_AddButton hPagerToolbar, 2002, STD_FILESAVE, "Save"
            Toolbar_AddButton hPagerToolbar, 2003, STD_CUT, "Cut"
            Toolbar_AddButton hPagerToolbar, 2004, STD_COPY, "Copy"
            Toolbar_AddButton hPagerToolbar, 2005, STD_PASTE, "Paste"
            Toolbar_AddButton hPagerToolbar, 2006, , "", TBSTATE_ENABLED, BTNS_SEP
            'Create a group
            Toolbar_AddButton hPagerToolbar, 2007, STD_PRINT, "", TBSTATE_ENABLED Or TBSTATE_CHECKED, BTNS_CHECKGROUP
            Toolbar_AddButton hPagerToolbar, 2008, STD_DELETE, "", TBSTATE_ENABLED, BTNS_CHECKGROUP
            Toolbar_AddButton hPagerToolbar, 2009, STD_FIND, "", TBSTATE_ENABLED, BTNS_CHECKGROUP
            'A_SendMessage hToolbar, TB_SETSTATE, 1009, MAKELONG(TBSTATE_ENABLED Or TBSTATE_CHECKED, 0&)
            Toolbar_AddButton hPagerToolbar, 2010, , "", TBSTATE_ENABLED, BTNS_SEP
            Toolbar_AddButton hPagerToolbar, 2011, STD_HELP, "", TBSTATE_ENABLED Or TBSTATE_CHECKED, BTNS_CHECK

            'you must send the TB_AUTOSIZE message after all the items and strings have been inserted into the control to cause the toolbar to recalculate its size based on its content.
            A_SendMessage hPagerToolbar, TB_AUTOSIZE, 0&, 0&

            'Set toolbar as the pager child
            'this message does not actually change the parent window of the
            'contained window; it simply assigns the contained window.
            A_SendMessage hPagerCtl, PGM_SETCHILD, 0&, hPagerToolbar
            
            ShowWndForFirstTime hFrmPager
        End If
    Else
        ShowWindow hFrmPager, SW_SHOW
    End If

    Exit Sub
BtnShowPageScroller_Click_Error:
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure BtnShowPageScroller_Click of Module Form_Pager"
End Sub

Public Function PagerCtlWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Dim nmh As NMHDR
    'At a minimum, it is necessary to process the PGN_CALCSIZE notification.
    'If you don't process this notification and enter a value for the width or height,
    'the scroll arrows in the pager control will not be displayed. This is because
    'the pager control uses the width or height supplied in the PGN_CALCSIZE notification
    'to determine the "ideal" size of the contained window.
    If uMsg = WM_NOTIFY Then
        CopyMemory nmh, ByVal lParam, Len(nmh)
        lRet = nmh.code
        If lRet = PGN_CALCSIZE Then
            Dim nmpg As NMPGCALCSIZE
            Dim tSize As Size
            'Get the structure
            CopyMemory nmpg, ByVal lParam, Len(nmpg)
            'Get the toolbar size
            A_SendMessageAnyRef hPagerToolbar, TB_GETMAXSIZE, 0&, tSize
            If nmpg.dwFlag = PGF_CALCWIDT Then
                nmpg.iWidth = tSize.cx
            ElseIf nmpg.dwFlag = PGF_CALCHEIGHT Then
                nmpg.iHeight = tSize.cy
            End If
            'Put structure back in the lParam
            CopyMemory ByVal lParam, nmpg, Len(nmpg)
            'return value is ignored
            Exit Function
        End If
    ElseIf uMsg = WM_CLOSE Then
        'just hide, do not destroy
        ShowWindow hFrmPager, SW_HIDE
        'DestroyWindow hFrmPager
        Exit Function
    End If
    PagerCtlWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
