Attribute VB_Name = "Form_ListView"
Option Explicit

'Form Listview
Public hFrmListView As Long
Public hListView As Long
Public hLvStatusbar As Long
Public hLvSmallImgList As Long
Public hLvLargeImgList As Long
'Public hLvBkImgList As Long     'Offer a different BK image for each view?
Public hLvMenu As Long
Public sLvTip As String
Public hLvEdit As Long
Public origLvEditWndProc As Long

Public Sub BtnShowListView_Click()
    Dim mnuItem As MENUITEMINFO
    Dim hTmpItem As Long, mnuSubItem As Long, sIcon As Long, lIcon As Long
    Dim sB As String
    
    If hFrmListView = 0 Then
        hFrmListView = CreateForm(300, 300, 350, 300, AddressOf ListViewWndProc, hMainForm, WS_SYSMENU Or WS_BORDER Or WS_OVERLAPPED, WS_EX_CONTROLPARENT Or WS_EX_WINDOWEDGE, "FormMyListView", "Cache Viewr - ListView")
        If hFrmListView > 0 Then
            'Create listview
            hListView = CreateListView(hFrmListView, 2, 2, 340, 250)
            'Set Bk image
            'ListView_SetBkImage hListView, App.Path & "\images\lvbk.bmp"
            ListView_SetBkColor hListView, x_Aquamarine
            
            sIcon = GetSystemMetrics(SM_CXSMICON)
            lIcon = GetSystemMetrics(SM_CXICON)
            'Create and assign image lists
            'The current image list will be destroyed when the list-view control is destroyed unless
            'the LVS_SHAREIMAGELISTS style is set. If you use this message to replace one image list with another,
            'your application must explicitly destroy all image lists other than the current one.
            'Create the imagelists and two icons to them - 32 bit with alpha channels Or ILD_TRANSPARENT
            hLvSmallImgList = ImageList_Create(sIcon, sIcon, ILC_COLOR32 Or ILC_MASK, 2, 0)
            If hLvSmallImgList > 0 Then
                ImageList_SetBkColor hLvSmallImgList, CLR_NONE
                
                sB = App.Path & "\images\alpha1.ICO"
                hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, sIcon, sIcon, LR_LOADFROMFILE)
                'i = Index of the image to replace. If i is -1,
                'the function appends the image to the end of the list.
                ImageList_ReplaceIcon hLvSmallImgList, -1, hTmpItem
                DestroyIcon hTmpItem
                
                sB = App.Path & "\images\alpha2.ICO"
                hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, sIcon, sIcon, LR_LOADFROMFILE)
                ImageList_ReplaceIcon hLvSmallImgList, -1, hTmpItem
                DestroyIcon hTmpItem
                
                If ImageList_GetImageCount(hLvSmallImgList) > 0 Then
                    ListView_SetImageList hListView, hLvSmallImgList, LVSIL_SMALL
                End If
            End If
            hLvLargeImgList = ImageList_Create(lIcon, lIcon, ILC_COLOR32 Or ILC_MASK, 2, 0)
            If hLvLargeImgList > 0 Then
                ImageList_SetBkColor hLvLargeImgList, CLR_NONE
                
                sB = App.Path & "\images\alpha1.ICO"
                hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, lIcon, lIcon, LR_LOADFROMFILE)
                ImageList_ReplaceIcon hLvLargeImgList, -1, hTmpItem
                DestroyIcon hTmpItem
                
                sB = App.Path & "\images\alpha2.ICO"
                hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, lIcon, lIcon, LR_LOADFROMFILE)
                ImageList_ReplaceIcon hLvLargeImgList, -1, hTmpItem
                DestroyIcon hTmpItem
                
                If ImageList_GetImageCount(hLvLargeImgList) > 0 Then
                    ListView_SetImageList hListView, hLvLargeImgList, LVSIL_NORMAL
                End If
            End If
            
            'Add columns with various text alignment
            ListView_AddColumn hListView, 0, "URL", , , 80
            ListView_AddColumn hListView, 1, "Title", , LVCFMT_CENTER, 80
            ListView_AddColumn hListView, 2, "LastVisited", , , 80
            ListView_AddColumn hListView, 3, "Expires", , LVCFMT_RIGHT, 80
            
            
    Dim cHist As CURLHistory
    Dim cItem As URLHistoryItem

    Set cHist = New CURLHistory

    If Not cHist Is Nothing Then
        For Each cItem In cHist
            hTmpItem = ListView_AddItemv5(hListView, cItem.URL, 1)
            If hTmpItem > -1 Then
                ListView_AddSubItem hListView, cItem.Title, hTmpItem, 1
                ListView_AddSubItem hListView, CStr(cItem.LastVisited), hTmpItem, 2
                ListView_AddSubItem hListView, CStr(cItem.Expires), hTmpItem, 3
            End If
        Next
        Set cHist = Nothing
    End If
            
'            'Add list items and sub items
'            hTmpItem = ListView_AddItemv5(hListView, "DOG a friendly pet", 1)
'                If hTmpItem > -1 Then
'                    ListView_AddSubItem hListView, "Poodle", hTmpItem, 1
'                    ListView_AddSubItem hListView, "White", hTmpItem, 2
'                    ListView_AddSubItem hListView, "$300.00", hTmpItem, 3
'                End If
'            hTmpItem = ListView_AddItemv5(hListView, "Cat a playful pet", 0)
'                If hTmpItem > -1 Then
'                    ListView_AddSubItem hListView, "Siamese", hTmpItem, 1
'                    ListView_AddSubItem hListView, "Mixed", hTmpItem, 2
'                    ListView_AddSubItem hListView, "$200.00", hTmpItem, 3
'                End If
'            hTmpItem = ListView_AddItemv5(hListView, "Fish a silent companion", 1)
'                If hTmpItem > -1 Then
'                    ListView_AddSubItem hListView, "Gold Fish", hTmpItem, 1
'                    ListView_AddSubItem hListView, "Red + White", hTmpItem, 2
'                    ListView_AddSubItem hListView, "$10.00", hTmpItem, 3
'                End If
            'Create a simple statusbar
            hLvStatusbar = CreateStatusBarSimple(hFrmListView, "Righ click for options", , , , 30)

            'Set up context menu
            hLvMenu = CreatePopupMenu()
            If hLvMenu > 0 Then
                'Add
                mnuItem = CreateMenuItem("Add", 1020, 0&)
                A_InsertMenuItem hLvMenu, 1020, 0, mnuItem
                'Remove
                mnuItem = CreateMenuItem("Remove Selected", 1021, 0&)
                A_InsertMenuItem hLvMenu, 1021, 0, mnuItem
                'Remove all
                mnuItem = CreateMenuItem("Remove All", 1022, 0&)
                A_InsertMenuItem hLvMenu, 1022, 0, mnuItem
                'Sep
                mnuItem = CreateMenuItem("-", 1023, 0&, True)
                A_InsertMenuItem hLvMenu, 1023, 0, mnuItem
                'View
                mnuSubItem = CreatePopupMenu()
                mnuItem = CreateMenuItem("Report", 1010, 0&, , , True)
                A_InsertMenuItem mnuSubItem, 1010, 0, mnuItem
                
                mnuItem = CreateMenuItem("Icons", 1011, 0&)
                A_InsertMenuItem mnuSubItem, 1011, 0, mnuItem
                
                mnuItem = CreateMenuItem("List", 1012, 0&)
                A_InsertMenuItem mnuSubItem, 1012, 0, mnuItem
                
                mnuItem = CreateMenuItem("ThumbNails", 1013, 0&)
                A_InsertMenuItem mnuSubItem, 1013, 0, mnuItem
                
                mnuItem = CreateMenuItem("Tiles", 1014, 0&)
                A_InsertMenuItem mnuSubItem, 1014, 0, mnuItem
                
                mnuItem = CreateMenuItem("View", 1015, mnuSubItem)
                A_InsertMenuItem hLvMenu, 1015, 0, mnuItem
                
                mnuItem = CreateMenuItem("-", 1016, 0&, True)
                A_InsertMenuItem hLvMenu, 1016, 0, mnuItem
                
                mnuItem = CreateMenuItem("Edit", 1017, 0&)
                A_InsertMenuItem hLvMenu, 1017, 0, mnuItem
            End If
            
            'Create our edit control and hide it
            hLvEdit = CreateTextbox(hListView, 0, 0, 30, 20, "", 0, , , True)
            ShowWindow hLvEdit, SW_HIDE
            
            'Subclass
            If hLvEdit > 0 Then origLvEditWndProc = A_SetWindowLong(hLvEdit, GWL_WNDPROC, AddressOf LvEditWndProc)
            
            ShowWndForFirstTime hFrmListView
        End If
    Else
        ShowWindow hFrmListView, SW_SHOW
    End If
    
End Sub

Public Function ListViewWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Dim lItem As Long
    Dim lMenuID As Long
    Dim nmh As NMHDR
    Dim sText As String
    Dim pt As POINTAPI
    Dim nmcd As NMLVCUSTOMDRAW
    Dim rcItem As RECT
    Dim rcBounds As RECT, rcLabel As RECT, rcIcon As RECT, rcCol As RECT, rcHighlight As RECT, rcClient As RECT
    Dim idp As IMAGELISTDRAWPARAMSv6
    Dim bHighlight As Boolean
    Dim pitem As LVITEM

    Select Case uMsg
        Case WM_NOTIFY ' lParam NMHDR
                CopyMemory nmh, ByVal lParam, Len(nmh)
                lRet = nmh.code
                Select Case lRet
''''''''''''''''
'''Custom Draw
''''''''''''''''
'For report mode:
'   The first NM_CUSTOMDRAW notification will have the dwDrawStage member of the associated NMCUSTOMDRAW
'   structure set to CDDS_PREPAINT. Return CDRF_NOTIFYITEMDRAW.
'   You will then receive an NM_CUSTOMDRAW notification with dwDrawStage set to CDDS_ITEMPREPAINT.
'   If you specify new fonts or colors and return CDRF_NEWFONT, all subitems of the item will be changed.
'   If you want instead to handle each subitem separately, return CDRF_NOTIFYSUBITEMDRAW.
'   If you returned CDRF_NOTIFYSUBITEMDRAW in the previous step, you will then receive an NM_CUSTOMDRAW
'   notification for each subitem with dwDrawStage set to CDDS_SUBITEM | CDDS_ITEMPREPAINT.
'   To change the font or color for that subitem, specify a new font or color and return CDRF_NEWFONT.

'For the large icon, small icon, and list modes:
'   The first NM_CUSTOMDRAW notification will have the dwDrawStage member of the associated NMCUSTOMDRAW
'   structure set to CDDS_PREPAINT. Return CDRF_NOTIFYITEMDRAW.
'   You will then receive an NM_CUSTOMDRAW notification with dwDrawStage set to CDDS_ITEMPREPAINT.
'   You can change the fonts or colors of an item by specifying new fonts and colors and returning
'   CDRF_NEWFONT. Because these modes do not have subitems, you will not receive any additional
'   NM_CUSTOMDRAW notifications.

                    Case NM_CUSTOMDRAW
                        'Set to default for now
                        ListViewWndProc = CDRF_DODEFAULT
                        'Copy structure
                        CopyMemory nmcd, ByVal lParam, Len(nmcd)
                        
                        'First drawing notification
                        If nmcd.nmcd.dwDrawStage = CDDS_PREPAINT Then
                            ListViewWndProc = CDRF_NOTIFYITEMDRAW
                            Exit Function
                        'Second drawing notification
                        ElseIf nmcd.nmcd.dwDrawStage = CDDS_ITEMPREPAINT Then
                            'Get view
                            lItem = ListView_GetView(hListView)
                            'if report we continue, else we do drawing here
                            If lItem = LVS_ICON Or lItem = LVS_SMALLICON Or lItem = LVS_LIST Then
                                    'contains garbage: nmcd.nmcd.uItemState
                                    'Get all states and images for this item
                                    pitem.iItem = nmcd.nmcd.dwItemSpec
                                    pitem.mask = LVIF_STATE Or LVIF_IMAGE
                                    pitem.iSubItem = 0
                                    pitem.stateMask = LVIS_SELECTED Or LVIS_DROPHILITED '&HFFFF 'Get all states
                                    'Could not get the item, just do default
                                    If ListView_GetItem(hListView, pitem) = False Then Exit Function
                                    'Set the flag
                                    bHighlight = (pitem.state And LVIS_SELECTED) Or (pitem.state And LVIS_DROPHILITED) And _
                                                    (GetFocus = hListView) Or (A_GetWindowLong(hListView, GWL_STYLE) And LVS_SHOWSELALWAYS)
                                    
                                    'Get rects
                                    ListView_GetItemRect hListView, nmcd.nmcd.dwItemSpec, rcBounds, LVIR_BOUNDS
                                    ListView_GetItemRect hListView, nmcd.nmcd.dwItemSpec, rcLabel, LVIR_LABEL
                                    ListView_GetItemRect hListView, nmcd.nmcd.dwItemSpec, rcIcon, LVIR_ICON
                                    
                                    'Get item text
                                    ListView_GetItemText hListView, nmcd.nmcd.dwItemSpec, sText

                                    'Draw stateimage + image using imagelist
                                    'idp.cbSize = Len(idp)
                                    'idp.hdcDst = nmcd.nmcd.hdc
                                    'idp.rgbBk = CLR_NONE
                                    'idp.rgbFg = CLR_NONE
                                    
                                    
                                    If bHighlight Then
                                        'Set text bk+fc
                                        SetTextColor nmcd.nmcd.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
                                        SetBkColor nmcd.nmcd.hdc, GetSysColor(COLOR_HIGHLIGHT)
                                        BoxSolidDC nmcd.nmcd.hdc, rcLabel, GetSysColor(COLOR_HIGHLIGHT)
                                        'Draw icon
                                        If lItem = LVS_ICON Then
                                            'idp.himl = hLvLargeImgList
                                            
                                            ImageList_Draw hLvLargeImgList, pitem.iImage, nmcd.nmcd.hdc, rcIcon.Left, rcIcon.Top, ILD_BLEND25
                                        Else
                                            ImageList_Draw hLvSmallImgList, pitem.iImage, nmcd.nmcd.hdc, rcIcon.Left, rcIcon.Top, ILD_BLEND25
                                        End If
                                    Else
                                        'If we have no bk image, just erase bk and get it ready
                                        'BoxSolidDC nmcd.nmcd.hdc, rcBounds, GetSysColor(COLOR_WINDOW)
                                        SetTextColor nmcd.nmcd.hdc, x_Yellow
                                        If lItem = LVS_ICON Then
                                            ImageList_Draw hLvLargeImgList, pitem.iImage, nmcd.nmcd.hdc, rcIcon.Left, rcIcon.Top, ILD_TRANSPARENT
                                        Else
                                            ImageList_Draw hLvSmallImgList, pitem.iImage, nmcd.nmcd.hdc, rcIcon.Left, rcIcon.Top, ILD_TRANSPARENT
                                        End If
                                    End If
                                    'Draw text
                                    A_DrawText nmcd.nmcd.hdc, sText, -1, rcLabel, DT_LEFT Or DT_END_ELLIPSIS
                                    If bHighlight Then DrawFocusRect nmcd.nmcd.hdc, rcLabel
                                    'Done!
                                    ListViewWndProc = CDRF_SKIPDEFAULT
                                    Exit Function
                            Else
                                'Report view, continue
                                ListViewWndProc = CDRF_NOTIFYSUBITEMDRAW
                            End If
                            Exit Function
'                       'Third drawing notification is received only if you are in report mode and passed
                        'CDRF_NOTIFYSUBITEMDRAW flage in the last drawing notification "CDDS_ITEMPREPAINT"
                        ElseIf nmcd.nmcd.dwDrawStage = CDDS_SUBITEMPOSPAINT Then
                            'nmcd.nmcd.dwItemSpec contains the 0 based index of each item
                            'nmcd.iSubItem contains 0 based index of each subitem
                            'change the selected color
                            If ListView_GetItemState(hListView, nmcd.nmcd.dwItemSpec, LVIS_SELECTED) And LVIS_SELECTED Then
                                nmcd.clrText = x_Black 'green
                                'nmcd.clrTextBk = x_Gold
                            Else
                                nmcd.clrText = x_Blue ' x_Yellow
                                'is ignored since we set the textbk to CLR_NONE in SetBkImage
                                'nmcd.clrTextBk = x_Gold
                                
                                'nmcd.clrText = GetSysColor(COLOR_WINDOWTEXT)
                                'nmcd.clrTextBk = GetSysColor(COLOR_WINDOW)
                            End If
                            'Copy structure back after modification
                            CopyMemory ByVal lParam, nmcd, Len(nmcd)
                            'Notify LV that things have changed
                            ListViewWndProc = CDRF_NEWFONT
                            Exit Function
                        End If
                        Exit Function
                    'Sent only for sub-item 0 to set it's tooltip
                    Case LVN_GETINFOTIP
                        'To add tips to the subitems:
                        'Create own tooltip, create listview without default toopltip, enable hottracking,
                        'use LVN_HOTTRACK msg(which Is similar To NM_CLICK, but uses NMLISTVIEW type)
                        'If iSubItem member is > -1 then get the text, update tooltip text and display the tooltip
                        Dim nmif As NMLVGETINFOTIP
                        CopyMemory nmif, ByVal lParam, Len(nmif)
                        'Get the mouse point, see if it is in
                        sLvTip = CHAR_ZERO_LENGTH_STRING
                        ListView_GetItemText hListView, nmif.iItem, sLvTip
                        If LenB(sLvTip) > NUM_ZERO Then
                            sLvTip = "Custome tip for " & sLvTip & vbNullChar
                            nmif.cchTextMax = Len(sLvTip)
                            A_lstrcpynStrPtr ByVal nmif.pszText, sLvTip, Len(sLvTip)
                        End If
                    Case NM_RCLICK   ' lParam = lp NMHDR
                        'Could use GetSelected also
                        lItem = ListView_HitTest(hListView)
                        If lItem > -1 Then
                            'Select the item as this event does not cause the item being selected
                            ListView_SetSelectedItem hListView, lItem
                            'Update status
                            ListView_GetItemText hListView, lItem, sText
                            StatusBar_SetText hLvStatusbar, ">> " & sText & " <<", SBT_POPOUT
                            'Display context menu
                            'This call is recommanded by MSDN due to a bug with TrackPopupMenuEx
                            SetForegroundWindow hFrmListView
                            Call GetCursorPos(pt)
                            'Returns the ID of the menu clicked or 0
                            lMenuID = TrackPopupMenuEx(hLvMenu, TPM_RETURNCMD Or TPM_LEFTALIGN, pt.X, pt.Y, hFrmListView, ByVal 0&)
                            'Continuing with bug related to TrackPopupMenuEx
                            A_PostMessage hFrmListView, WM_NULL, 0&, 0&
                            If lMenuID <> 0 Then
                                'Uncheck all view related menus
                                CheckMenuItem hLvMenu, 1010, MF_BYCOMMAND Or MF_UNCHECKED
                                CheckMenuItem hLvMenu, 1011, MF_BYCOMMAND Or MF_UNCHECKED
                                CheckMenuItem hLvMenu, 1012, MF_BYCOMMAND Or MF_UNCHECKED
                                CheckMenuItem hLvMenu, 1013, MF_BYCOMMAND Or MF_UNCHECKED
                                CheckMenuItem hLvMenu, 1014, MF_BYCOMMAND Or MF_UNCHECKED
                                Select Case lMenuID
                                    Case 1010 'Report
                                        ListView_SetView hListView, LVS_REPORT
                                        CheckMenuItem hLvMenu, 1010, MF_BYCOMMAND Or MF_CHECKED
                                    Case 1011 'Icons
                                        ListView_SetView hListView, LVS_ICON
                                        CheckMenuItem hLvMenu, 1011, MF_BYCOMMAND Or MF_CHECKED
                                    Case 1012 'List
                                        ListView_SetView hListView, LVS_LIST
                                        CheckMenuItem hLvMenu, 1012, MF_BYCOMMAND Or MF_CHECKED
                                    Case 1013 'ThumbNails
                                        ListView_SetView hListView, LVS_SMALLICON
                                        CheckMenuItem hLvMenu, 1013, MF_BYCOMMAND Or MF_CHECKED
                                    Case 1014 'Tiles
                                        'only for XP
                                        ListView_SetView hListView, LVS_TILE
                                        CheckMenuItem hLvMenu, 1014, MF_BYCOMMAND Or MF_CHECKED
                                    'Case 1020 'Add
                                        
                                    Case 1021 'Remove
                                        If A_MessageBox(hFrmListView, "Proceed to delete " & sText & " ?", "Remove Item", MB_YESNO Or MB_ICONQUESTION) = IDYES Then
                                            ListView_DeleteItem hListView, lItem
                                        End If
                                    Case 1022 'Remove all
                                        If A_MessageBox(hFrmListView, "Proceed to remove all items?", "Remove All Items", MB_YESNO Or MB_ICONQUESTION) = IDYES Then
                                            ListView_DeleteAllItems hListView
                                        End If
                                    Case 1017 'Edit
                                    
                                End Select
                            End If
                        Else
                            Dim lCount As Long, lHeader As Long, lHeaderCount As Long
                            Dim rtColumn As RECT
                            
                            'returns the parent item index or -1, zero based
                            lItem = ListView_SubItemHitTest(hListView)
                            'Debug.Print "litem= " & lItem

                            Call GetCursorPos(pt)
                            
                            SetForegroundWindow hFrmListView
                            'Returns the ID of the menu clicked or 0
                            lMenuID = TrackPopupMenuEx(hLvMenu, TPM_RETURNCMD Or TPM_LEFTALIGN, pt.X, pt.Y, hFrmListView, ByVal 0&)
                            'Continuing with bug related to TrackPopupMenuEx
                            A_PostMessage hFrmListView, WM_NULL, 0&, 0&
                            
                            ScreenToClient hListView, pt
                            'Get the column by looping through each header and comparing x and y
                            lHeader = ListView_GetHeader(hListView)
                            If lHeader > 0 Then
                                'Get header count, zero based
                                lHeaderCount = A_SendMessage(lHeader, HDM_GETITEMCOUNT, 0&, 0&)
                                If lHeaderCount > -1 Then
                                    'Loop through all columns
                                    For lCount = 0 To lHeaderCount - 1
                                        'Get column rect
                                        If A_SendMessageAnyRef(lHeader, HDM_GETITEMRECT, lCount, rtColumn) <> 0 Then
                                            'Find the subitem by comparing the pt.x and pt.y to the rect coordinates
                                            If pt.X > rtColumn.Left And pt.X < rtColumn.Right Then
                                                If lItem > -1 Then
                                                    'Select item
                                                    'ListView_SetSelectedItem hListView, lItem
                                                    'Get subitem text
                                                    ListView_GetItemText hListView, lItem, sText, lCount
                                                    'Reset rect
                                                    SetRectEmpty rtColumn
                                                    'Get subitem rect
                                                    ListView_GetSubItemRect hListView, lItem, lCount, LVIR_BOUNDS, rtColumn
                                                    'Adjust edit
                                                    MoveWindow hLvEdit, rtColumn.Left, rtColumn.Top, rtColumn.Right - rtColumn.Left, rtColumn.Bottom - rtColumn.Top, 0&
                                                    'Set edit text
                                                    A_SetWindowText hLvEdit, sText
                                                    'Show edit
                                                    ShowWindow hLvEdit, SW_SHOW
                                                    'set focus to edit
                                                    SetFocusApi hLvEdit
                                                    'Select the edit text
                                                    TextBox_SetSel hLvEdit
                                                End If
                                                'Debug.Print "lItem:" & lItem & " ColIndex:" & lCount & "=" & rtColumn.Left & "=" & rtColumn.Top & "=" & rtColumn.Right & "=" & rtColumn.Bottom
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If 'End >> If lHeaderCount > -1 Then
                            End If 'End >> If lHeader > 0 Then

                        End If 'End >> If lItem > -1 Then
                    Case NM_CLICK 'lParam = NMITEMACTIVATE
                        Dim nmia As NMITEMACTIVATE
                        CopyMemory nmia, ByVal lParam, Len(nmia)
                        'The iItem member of lpnmitem will only be valid if the icon or first-column label
                        'has been clicked. To determine which item is selected when a click takes place
                        'elsewhere in a row, send an LVM_SUBITEMHITTEST message.
                        If nmia.iItem > -1 Then
                            ListView_GetItemText hListView, nmia.iItem, sText
                            'Changes the text and reverts back to the original style
                            'A_SetWindowText hLvStatusbar, sText
                            'Same as  above in a simple status bar except it also allows you to specify a drawing style
                            StatusBar_SetText hLvStatusbar, "ITEM >> " & sText & " <<", SBT_POPOUT
                        ElseIf nmia.iSubItem > 0 Then 'over other columns not the parent
                            'returns the parent item index or -1
                            lItem = ListView_SubItemHitTest(hListView)
                            If lItem > -1 Then
                                'Get the subitem text
                                ListView_GetItemText hListView, lItem, sText, nmia.iSubItem
                                StatusBar_SetText hLvStatusbar, "SUBITEM >> " & sText & " <<", SBT_POPOUT
                                'select the parent
                                ListView_SetSelectedItem hListView, lItem
                            End If
                        End If
'                    Case NM_DBLCLK 'lParam = NMITEMACTIVATE
'                        'Handling is similar to NM_CLICK msg, uses same type
                    Case LVN_COLUMNCLICK 'lParam = NMLISTVIEW
                        'The iItem member is -1, and the iSubItem member identifies the column.
                        'All other members are zero.
                        Dim nmlv As NMLISTVIEW
                        CopyMemory nmlv, ByVal lParam, Len(nmlv)
                        StatusBar_SetText hLvStatusbar, "COLUMN >> " & CStr(nmlv.iSubItem) & " <<", SBT_POPOUT
                    Case LVN_KEYDOWN 'lParam = NMLVKEYDOWN
                        Dim nmkd As NMLVKEYDOWN
                        CopyMemory nmkd, ByVal lParam, Len(nmkd)
                        StatusBar_SetText hLvStatusbar, "KeyDown >> " & CStr(nmkd.wVKey) & " <<", SBT_POPOUT
'                    Case NM_KILLFOCUS 'lParam NMHDR
'                     Case NM_SETFOCUS 'lParam NMHDR
                        ''If IsWindowVisible(hLvEdit) Then ShowWindow hLvEdit, SW_HIDE
'                    Case NM_RETURN 'lParam NMHDR
'                        'Notifies a list-view control's parent window that the control has the input focus
'                        'and that the user has pressed the ENTER key.
'                    Case LVN_ITEMCHANGED 'lParam NMLISTVIEW
'                        'Pointer to an NMLISTVIEW structure that identifies the item and specifies which of
'                        'its attributes have changed. If the iItem member of the structure
'                        'pointed to by pnmv is -1, the change has been applied to all items in the list view.
                End Select 'End case lRet
        Case WM_CLOSE
            'just hide, do not destroy
            ShowWindow hFrmListView, SW_HIDE
            BringWindowToTop hMainForm
            'DestroyWindow hFrmListView
            Exit Function
    End Select
    ListViewWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function

'Wnd procedures for the LvEdit textbox to catch return, escape and killfocus msgs
Public Function LvEditWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_CONTEXTMENU
            'Debug.Print "NoContext"
            Exit Function
        Case WM_KILLFOCUS
            ShowWindow hLvEdit, SW_HIDE
        Case WM_KEYDOWN
            'Can limit the key input here if wParam = VK_F,...
            'If enter we update the subitem
            'if escape then we cancel
            If wParam = VK_RETURN Or wParam = VK_ESCAPE Then '13 27
                ShowWindow hLvEdit, SW_HIDE
                'Return 0 to indicate that msg was handeled
                Exit Function
            End If
    End Select
    LvEditWndProc = A_CallWindowProc(origLvEditWndProc, hWnd, uMsg, wParam, lParam)
End Function

