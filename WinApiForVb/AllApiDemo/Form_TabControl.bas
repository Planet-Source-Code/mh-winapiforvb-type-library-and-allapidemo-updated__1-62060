Attribute VB_Name = "Form_TabControl"
Option Explicit

'Form Tab control
Public hFrmTabControl As Long
Public hTabControl As Long
Public hTxtTabCaption As Long
Public hTxtTabImage As Long
Public hTxtTablParam As Long
Public hTxtTabInsertIndex As Long
Public hBtnAddTab As Long
Public hBtnRemoveTab As Long

Public Sub BtnShowTabControl_Click()
    Dim udtMsg As MSG
    If hFrmTabControl = 0 Then '
        hFrmTabControl = CreateForm(300, 300, 280, 250, AddressOf TabControlWndProc, hMainForm, , , "FormTabControl", "Tab Control")
        If hFrmTabControl > 0 Then
            'Create controls
            hTabControl = CreateTabCtl(hFrmTabControl, 2, 2, 270, 30)
            'Add tabs and create other controls if we have a tab control
            If hTabControl > 0 Then
                'Set up an image list
                If hTvImgListState16 = 0 Then SetupImageListState16
                TabCtl_SetImageList hTabControl, hTvImgListState16
                'Add two tabs
                TabCtl_InsertItem hTabControl, "Item One", 1, 10, 0
                TabCtl_InsertItem hTabControl, "Item with big caption more than tab", 2, 20, 1
                'Create the rest of controls
                CreateLabel hFrmTabControl, "Tab Caption:", 5, 30, 100, 17, SS_RIGHT
                hTxtTabCaption = CreateTextbox(hFrmTabControl, 105, 30, 150, 25, "Item One")
                CreateLabel hFrmTabControl, "Tab Image:", 5, 60, 100, 17, SS_RIGHT
                hTxtTabImage = CreateTextbox(hFrmTabControl, 105, 60, 50, 25, "1", , , , , , , , , , , , True)
                CreateLabel hFrmTabControl, "Tab lParam:", 5, 90, 100, 17, SS_RIGHT
                hTxtTablParam = CreateTextbox(hFrmTabControl, 105, 90, 50, 25, "10", , , , , , , , , , , , True)
                CreateLabel hFrmTabControl, "Tab Index:", 5, 120, 100, 17, SS_RIGHT
                hTxtTabInsertIndex = CreateTextbox(hFrmTabControl, 105, 120, 50, 25, "0")
                hBtnAddTab = CreateCmdButton(hFrmTabControl, "Add New Tab", 5, 150, 150, 30, , , , , , , , WS_EX_DLGMODALFRAME)
                hBtnRemoveTab = CreateCmdButton(hFrmTabControl, "Delete Selected Tab", 5, 185, 150, 30, , , , , , , , WS_EX_DLGMODALFRAME)
            End If
            ShowWndForFirstTime hFrmTabControl
        End If
    Else
        ShowWindow hFrmTabControl, SW_SHOW
    End If
End Sub

Public Function TabControlWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Dim nmh As NMHDR
    Dim curTab As Long
    Dim sTabText As String
    Dim lInsertIndex As Long

    Select Case uMsg
        Case WM_SIZE
            If wParam <> SIZE_MINIMIZED Then
                Dim lRect As RECT
                'Get form's current dimensions.
                GetWindowRect hFrmTabControl, lRect
                'Resize tabcontrol
                SetWindowPos hTabControl, 0, 2, 2, lRect.Right - (lRect.Left + 10), 25, SWP_NOZORDER Or SWP_SHOWWINDOW
                Exit Function
            End If
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                Select Case lParam
                    Case hBtnAddTab
                        'No error checking is done for image, or index
                        lInsertIndex = CLng(Window_GetText(hTxtTabInsertIndex))
                        curTab = TabCtl_InsertItem(hTabControl, Window_GetText(hTxtTabCaption), _
                                                        CLng(Window_GetText(hTxtTabImage)), _
                                                        CLng(Window_GetText(hTxtTablParam)), _
                                                        lInsertIndex)
                        'Select the new tab
                        'A tab control does not send a TCN_SELCHANGING or TCN_SELCHANGE notification message
                        'when a tab is selected using this message.
                        If curTab > -1 Then A_SendMessage hTabControl, TCM_SETCURSEL, curTab, 0&
                        If lInsertIndex > -1 Then TabCtl_ReassignTooltips hTabControl, lInsertIndex
                        Exit Function
                    Case hBtnRemoveTab
                        curTab = A_SendMessage(hTabControl, TCM_GETCURSEL, 0&, 0&)
                        If curTab > -1 Then
                            A_SendMessage hTabControl, TCM_DELETEITEM, curTab, 0&
                            A_SendMessage hTabControl, TCM_SETCURSEL, curTab, 0&
                            TabCtl_ReassignTooltips hTabControl, curTab
                            Exit Function
                        End If
                End Select
            End If
        Case WM_NOTIFY
            CopyMemory nmh, ByVal lParam, Len(nmh)
            lRet = nmh.code
            Select Case lRet
                'Selection changed
                Case TCN_SELCHANGE
                    'Returns the index of the selected tab if successful, or -1 if no tab is selected.
                    curTab = A_SendMessage(hTabControl, TCM_GETCURSEL, 0&, 0&)
                    If curTab > -1 Then
                        If TabCtl_GetTabText(hTabControl, curTab, sTabText) = True Then
                            A_SetWindowText hTxtTabCaption, sTabText
                        Else
                            A_SetWindowText hTxtTabCaption, "No Text"
                        End If
                        A_SetWindowText hTxtTabImage, CStr(TabCtl_GetTabImage(hTabControl, curTab))
                        A_SetWindowText hTxtTablParam, CStr(TabCtl_GetTablParam(hTabControl, curTab))
                        A_SetWindowText hTxtTabInsertIndex, CStr(curTab)
                        Exit Function
                    End If
                'To display tooltips, coming from tab's tooltip control
                'first time when it needs the tiptext. Hits only once
                Case TTN_GETDISPINFO
                    Dim nmtdi As NMTTDISPINFO
                    Dim tchi As TCHITTESTINFO
                    Dim hitPt As POINTAPI
                    'Copy structure to local var
                    CopyMemory nmtdi, ByVal lParam, Len(nmtdi)
                    'Get the tab under the mouse
                    GetCursorPos hitPt
                    ScreenToClient hTabControl, hitPt
                    tchi.pt = hitPt
                    'Get the current tab under the mouse
                    curTab = A_SendMessageAnyRef(hTabControl, TCM_HITTEST, 0&, tchi)
                    If curTab > -1 Then
                        'Get the tab text
                        If TabCtl_GetTabText(hTabControl, curTab, sTabText) = True Then
                            sTabText = sTabText & vbNullChar
                            A_lstrcpynStrByte nmtdi.szText(0), sTabText, Len(sTabText)
                            'Copy local structure back into lParam
                            CopyMemory ByVal lParam, nmtdi, Len(nmtdi)
                            'Can update tiptext manually also
                            'ToolTip_UpdateTipText nmtdi.hdr.hwndFrom, hTabControl, curTab, sTabText
                            Exit Function
                        End If
                    End If
'                Case NM_CLICK
            End Select
        Case WM_CLOSE
'            'just hide, do not destroy
            ShowWindow hFrmTabControl, SW_HIDE
            'DestroyWindow hFrmTabControl
            'hFrmTabControl = 0
            Exit Function
    End Select
    TabControlWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
