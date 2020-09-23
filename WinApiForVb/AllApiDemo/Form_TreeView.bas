Attribute VB_Name = "Form_TreeView"
Option Explicit

'Form Treeview
Public hFrmTreeView As Long
'Tree Context Menu
Public hMenuTree As Long
Public hTreeLabel As Long
Public hTree As Long
Public hBtnCloseTreeView  As Long

'ImageLists 16x16, normal+state
Public hTvImgList16 As Long
Public hTvImgListState16 As Long

'Treeview Contextmenu ID's
Public Const MNUTV_EMAIL_ID As Long = 1000
Public Const MNUTV_IMPORTANT_ID As Long = 1001
Public Const MNUTV_LOCK_ID As Long = 1002
Public Const MNUTV_PRINT_ID As Long = 1003
Public Const MNUTV_UPLOAD_ID As Long = 1004
'Sep
Public Const MNUTV_SEP_ID1 As Long = 100
'edit expand collapse node
Public Const MNUTV_EDITLABEL_ID As Long = 1005
Public Const MNUTV_EXPAND_COLLAPSE_ID As Long = 1006
'Sep
Public Const MNUTV_SEP_ID2 As Long = 101
'Add remove
Public Const MNUTV_ADDNEW_ID As Long = 1007
Public Const MNUTV_REMOVESELECTED_ID As Long = 1008
Public Const MNUTV_REMOVEALL_ID As Long = 1009
'Sep
Public Const MNUTV_SEP_ID3 As Long = 102
Public Const MNUTV_SAVEWITHTAB_ID As Long = 1010
'Sep
Public Const MNUTV_SEP_ID4 As Long = 103
Public Const MNUTV_CHANGEFONT_ID As Long = 1020

'Treeview drag and drop
Public bDragTree As Boolean
Public lDragImg As Long
Public lDragItem As Long
Public lDropItem As Long
'Timer and timer id
Public lTreeTimer As Long
Public Const TREE_TIMER_ID As Long = NUM_ONE

'To display custom tips in WM_NOTIFY -> TVN_GETINFOTIP
Public sTreeTip As String

'To display individual fonts and colors
Public hMnuFont As Long
'Temp var used in enumfontproc as id for the new submenus
Private lFontMenuID As Long
'A simple class to manage nodes data, just a sample, easily can be modified/adopted
Public cNodesD As cNodeData


'===============Wnd procedure

Public Function TreeViewWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim nmh As NMHDR
    Dim nmtv As NMTREEVIEW
    Dim nmkd As TVKEYDOWN
    Dim tvdis As NMTVDISPINFO
    Dim tVi As TVITEM
    Dim nmcd As NMTVCUSTOMDRAW
    Dim pt As POINTAPI
    Dim rcNode As RECT
    Dim lMenuID As Long
    Dim lLen As Long, lItem As Long
    Dim sText As String
    'Back and forecolor
    Dim lBC As Long, lFC As Long, lLevel As Long, lFontStyle As Long, lTemp As Long
    Dim lOldBC As Long, lOldFC As Long
    
    Dim lRet As Long
    Select Case uMsg
'''''''''''''''''''''''''''
''''Tree DragDrop handling
'''''''''''''''''''''''''''
    Case WM_MOUSEMOVE
        If bDragTree Then
            'Get screen coordinates
            GetCursorPos pt
            ImageList_DragMove pt.X, pt.Y
            lDropItem = TreeView_GetFromPoint(hTree, pt)
            If lDropItem > 0 Then
                ImageList_DragShowNolock 0&
                TreeView_SelectDropTarget hTree, lDropItem
                ImageList_DragShowNolock 1&
            Else
                'We can change the cursor to no drop here
                lDropItem = 0
            End If
        End If
        'An application should return zero if it processes this message.
        Exit Function
    Case WM_LBUTTONUP
        If bDragTree Then
            Dim sDrag As String, sDrop As String
            KillTimer hFrmTreeView, TREE_TIMER_ID
            ImageList_DragLeave 0& 'hTree
            ImageList_EndDrag
            ReleaseCapture
            If lDragImg > 0 Then ImageList_Destroy lDragImg
            bDragTree = False
            'unhighlight the target
            TreeView_SelectDropTarget hTree, 0
            If lDropItem <> 0 Then
                TreeView_SetSelected hTree, lDropItem
                TreeView_GetText hTree, lDropItem, sDrop
            End If
            lDropItem = 0
            RefreshApi hTree
            TreeView_GetText hTree, lDragItem, sDrag
            sDrag = "lDragitem: <" & sDrag & "> lDropItem: <" & sDrop & ">"
            A_SetWindowText hTreeLabel, sDrag
            'An application should return zero if it processes this message.
            Exit Function
        End If
'''''''''''''''''''''''''''
'''''Tree DragDrop Timer Handling, for scrolling up/down while dragging
'''''Internal timers, make me almost cry
'''''''''''''''''''''''''''
        Case WM_TIMER
            'Timer to scroll tree up or down while dragging
            If wParam = TREE_TIMER_ID Then
                Dim dRect As RECT
                'get the tree rect
                GetClientRect hTree, dRect
                'Get the cursor position relative to screen
                GetCursorPos pt
                'Convert cursor pos relative to tree rect (client)
                ScreenToClient hTree, pt
                'scroll up or down?
                If pt.Y < NUM_ZERO Then 'Up
                    A_SendMessage hTree, WM_VSCROLL, 0&, vbNull
                ElseIf pt.Y > dRect.Bottom Then 'Down
                    A_SendMessage hTree, WM_VSCROLL, 1&, vbNull
                End If
                'An application should return zero if it processes this message.
                Exit Function
            End If
'''''''''''''''''''''''''''
''''''TREEVIEW notification handling
'''''''''''''''''''''''''''
        Case WM_NOTIFY ' lParam NMHDR
                CopyMemory nmh, ByVal lParam, Len(nmh)
                lRet = nmh.code
                Select Case lRet
                    'Ownerdrawn
                    Case NM_CUSTOMDRAW
                        CopyMemory nmcd, ByVal lParam, Len(nmcd)
                        
                        'In prepaint tell the tree that we only want to paint the item(text) part only
                        'So notify us for each item being prepainted. We can paint everything here ourselves
                        If nmcd.nmcd.dwDrawStage = CDDS_PREPAINT Then
                            TreeViewWndProc = CDRF_NOTIFYITEMDRAW Or CDRF_NOTIFYPOSTPAINT
                        ElseIf nmcd.nmcd.dwDrawStage = CDDS_ITEMPOSTPAINT Then
                            If nmcd.nmcd.uItemState And CDIS_FOCUS Then
                                TreeView_GetItemRect hTree, nmcd.nmcd.dwItemSpec, rcNode
                                If cNodesD.GetNodeData(sText, lLevel, lBC, lFC, lItem, lFontStyle, , lLen) > -1 Then
                                    'Set colors
                                    rcNode.Top = rcNode.Top + 1
                                    'Get the text
                                    TreeView_GetText hTree, nmcd.nmcd.dwItemSpec, sText
                                    'Save b + f colors
                                    lOldFC = SetTextColor(nmcd.nmcd.hdc, lFC)
                                    lOldBC = SetBkColor(nmcd.nmcd.hdc, lBC)
                                    'draw text
                                    A_DrawText nmcd.nmcd.hdc, " " & sText, -1, rcNode, DT_LEFT
                                    'reset b + f colors
                                    SetTextColor nmcd.nmcd.hdc, lOldFC
                                    SetBkColor nmcd.nmcd.hdc, lOldBC
                                'Else
                                    'Erase the Color Text drawn previously
                                    'FillRect nmcd.nmcd.hdc, rcNode, COLOR_WINDOW + 1
                                End If
                            End If
                            TreeViewWndProc = CDRF_DODEFAULT
                        ElseIf nmcd.nmcd.dwDrawStage = CDDS_ITEMPREPAINT Then
                            'See if we have a nodedata item by this nmcd.nmcd.dwItemSpec id
                            lLen = cNodesD.FindNodeIndex(nmcd.nmcd.dwItemSpec)
                            If lLen > -1 Then
                                'Get the FONT
                                lItem = cNodesD.GetNodeFont(, lLen)
                                'Debug.Print "Index: " & lLen & " Node: " & nmcd.nmcd.dwItemSpec & " Font=" & lItem
                                'Set the new font
                                If lItem <> 0 Then
                                    SelectObject nmcd.nmcd.hdc, lItem
                                    'Let tree know things have changed
                                    TreeViewWndProc = CDRF_NEWFONT Or CDRF_NOTIFYPOSTPAINT
                                End If
                            End If
                        End If
                        'let the ctl do it's default painting
                        'TreeViewWndProc = CDRF_DODEFAULT '0
                        Exit Function
                  ''To customize label editing, implement a handler for TVN_BEGINLABELEDIT and have it
                  ''send a TVM_GETEDITCONTROL message to the tree-view control. If a label is being edited,
                  ''the return value will be a handle to the edit control. Use this handle to customize
                  ''the edit control by sending the usual EM_XXX messages.
    '              Case TVN_BEGINLABELEDIT 'lParam TVDISPINFO
    '                CopyMemory tvdis, ByVal lParam, Len(tvdis)
    '                Debug.Print "TVN_BEGINLABELEDIT-NodeID " & tvdis.item.hItem
                  ''If label editing was canceled, the pszText member of the TVITEM structure is NULL; otherwise, pszText is the address of the edited text.
                  ''If the pszText member is non-NULL, return TRUE to set the item's label to the edited text. Return FALSE to reject the edited text and revert to the original label.
    '              Case TVN_ENDLABELEDIT 'lParam TVDISPINFO
    '                CopyMemory tvdis, ByVal lParam, Len(tvdis)
    '                Debug.Print "TVN_ENDLABELEDIT-NodeID " & tvdis.item.hItem
    '              Case TVN_KEYDOWN   ' lParam = lp TVKEYDOWN
    '                ''If the wVKey member of ptvkd is a character key code, the character will be used as part of an incremental search. Return nonzero to exclude the character from the incremental search, or zero to include the character in the search.
    '                CopyMemory nmkd, ByVal lParam, Len(nmkd)
    '                Debug.Print "TVN_KEYDOWN: " & Chr(nmkd.wVKey)
                  Case TVN_GETINFOTIP
                    'To display custom tips by setting TVS_INFOTIP flag
                    'during tree creation
                    Dim tvginfo As TVGETINFOTIP
                    CopyMemory tvginfo, ByVal lParam, Len(tvginfo)
                    sTreeTip = CHAR_ZERO_LENGTH_STRING
                    TreeView_GetText hTree, tvginfo.hItem, sTreeTip
                    'tvginfo.cchTextMax is normall 1024 but MS states that there are no guaranties
                    'of what the actual size of buffer is?
                    If LenB(sTreeTip) > NUM_ZERO Then
                        sTreeTip = "Custome tip for " & RTrim(sTreeTip) & vbNullChar
                        tvginfo.cchTextMax = Len(sTreeTip)
                        A_lstrcpynStrPtr ByVal tvginfo.pszText, sTreeTip, Len(sTreeTip)
                        'CopyMemoryFromStr ByVal tvginfo.pszText, sTreeTip, Len(sTreeTip)
                    End If
'The next commented block causes major problems with XP when TVS_SINGLEEXPAND flag is set
'                  ''Notifies a tree-view control's parent window that a parent item's list of child items has expanded or collapsed.
'                  ''The action member indicates whether the list expanded or collapsed. For a list of possible values
'                  Case TVN_ITEMEXPANDING   ' lParam = lp NMTREEVIEW
'                    'Debug.Print "TVN_ITEMEXPANDING"
'                    CopyMemory nmtv, ByVal lParam, Len(nmtv)
'                    'Have any children
'                    If TreeView_GetChild(hTree, nmtv.itemNew.hItem) <= NUM_ZERO Then
'                        TreeView_SetNumberOfChildren hTree, nmtv.itemNew.hItem
'                    Else
'                        'To have a plus button, this does not change the actual number of nodes
'                        'chldren, just a way to force display of plus buttons while adding nodes
'                        TreeView_SetNumberOfChildren hTree, nmtv.itemNew.hItem, 1
'                    End If
'                    'Has to be visible otherwise won't work
'                    TreeView_EnsureVisible hTree, nmtv.itemNew.hItem
'                    TreeView_SetSelected hTree, nmtv.itemNew.hItem
'                  Case TVN_GETDISPINFO    ' lParam = NMTVDISPINFO
'                    'Receiving this msg in response to setting TVITEM flag(s) to
'                    'I_CHILDRENCALLBACK - I_IMAGECALLBACK - LPSTR_TEXTCALLBACK when adding nodes
'                    'Here we force the tree to add a plus button to every node
'                    'then in selction changed or item expanding we reset the children value
'                    Dim nmdi As NMTVDISPINFO
'                    CopyMemory nmdi, ByVal lParam, Len(nmdi)
'                    If nmdi.Item.mask = TVIF_CHILDREN Then nmdi.Item.cChildren = 1
                  Case TVN_BEGINDRAG       ' lParam = lp NMTREEVIEW
                    Dim iif As IMAGEINFO
                    CopyMemory nmtv, ByVal lParam, Len(nmtv)
                    lDragItem = nmtv.itemNew.hItem
                    lDropItem = 0
                    TreeView_SelectDropTarget hTree, 0
                    lDragImg = TreeView_CreateDragImage(hTree, lDragItem)
                    'Debug.Print "lDargImg: " & lDragImg & " " & lDragItem '& " =" & GetParent(hTree) & "=" & hFrmTreeView
                    'Get the hotspot set up. Center mouse pointer on the top center of drag image
                    ImageList_GetImageInfo lDragImg, 0, iif
                    pt.X = (iif.rcImage.Right - iif.rcImage.Left) / 2
                    pt.Y = (iif.rcImage.Bottom - iif.rcImage.Top) / 2
                    'begin drag
                    ImageList_BeginDrag lDragImg, 0, -15, -15 'pt.X, pt.Y
                    'Get drag coordiates
                    pt = nmtv.ptDrag
                    'Covert to screen coordinates
                    ClientToScreen hTree, pt
                    'Enter drag, pass 0 as hwnd so the user cab drag the cursor outside tree
                    ImageList_DragEnter 0&, pt.X, pt.Y
                    'set mouse capture to parent so tree will send us the drag msgs back to us in this proc
                    SetCapture GetParent(hTree)
                    bDragTree = True
                    'Start the drag timer, for scrolling up or down while dragging
                    lTreeTimer = SetTimer(hFrmTreeView, TREE_TIMER_ID, 200, ByVal 0&)
                    'sends a WM_LBUTTONUP when done
'                  Case TVN_SELCHANGED   ' lParam = lp NMTREEVIEW
'                    CopyMemory nmtv, ByVal lParam, Len(nmtv)
    '                 Select Case nmtv.Action
    '                    Case TVC_BYMOUSE
    '                        Debug.Print "ACTION-TVC_BYMOUSE"
    '                    Case TVC_BYKEYBOARD
    '                        Debug.Print "ACTIONT-VC_BYKEYBOARD"
    '                    Case TVC_UNKNOWN
    '                        Debug.Print "ACTION-TVC_UNKNOWN"
    '                 End Select
'Causes problems with TVS_SINGLEEXPAND style in XP
'                    'This does not actually sets the number of children, merely adds/removes
'                    'the plus button
'                    If TreeView_GetChild(hTree, nmtv.itemNew.hItem) <= NUM_ZERO Then
'                        TreeView_SetNumberOfChildren hTree, nmtv.itemNew.hItem
'                    Else
'                        TreeView_SetNumberOfChildren hTree, nmtv.itemNew.hItem, 1
'                    End If
'                    If (TVIF_TEXT Or nmtv.itemNew.mask) And (nmtv.itemNew.hItem > 0) Then
'                        If TreeView_GetText(hTree, nmtv.itemNew.hItem, sText) = True Then
'                            Debug.Print "Node " & sText & " was clicked: " & nmtv.itemNew.hItem
'                        End If
'                    End If
    '                Case TVN_SELCHANGEDW   ' lParam = lp NMTREEVIEW
                    'NM_RCLICK hits before TVN_SELCHANGEDA hits
                    Case NM_RCLICK   ' lParam = lp NMHDR
                      'How can I select items with a right-click of the mouse without the highlight going back to the last selected?
                      lItem = TreeView_GetDropHilite(hTree)
                      If lItem <> 0 Then
                        'hFrmTreeView
                        TreeView_SetSelected hTree, lItem
                        'This call is recommanded by MSDN due to a bug with TrackPopupMenuEx
                        SetForegroundWindow hFrmTreeView
                        Call GetCursorPos(pt)
                        'Returns the ID of the menu clicked or 0
                        lMenuID = TrackPopupMenuEx(hMenuTree, TPM_RETURNCMD Or TPM_LEFTALIGN, pt.X, pt.Y, hFrmTreeView, ByVal 0&)
                        'Continuing with bug related to TrackPopupMenuEx
                        A_PostMessage hFrmTreeView, WM_NULL, 0&, 0&
                        'Process selected menu
                        If lMenuID <> 0 Then ProcessContextMenu lItem, lMenuID
                      'Try the selected one, just in case they RClicked on the
                      'same node twice
                      Else
                          lItem = TreeView_GetSelected(hTree)
                          If lItem <> 0 Then
                            TreeView_SetSelected hTree, lItem
                            SetForegroundWindow hFrmTreeView
                            Call GetCursorPos(pt)
                            'Returns the ID of the menu clicked or 0
                            lMenuID = TrackPopupMenuEx(hMenuTree, TPM_RETURNCMD Or TPM_LEFTALIGN, pt.X, pt.Y, hFrmTreeView, ByVal 0&)
                            A_PostMessage hFrmTreeView, WM_NULL, 0&, 0&
                            'Process selected menu
                            If lMenuID <> 0 Then ProcessContextMenu lItem, lMenuID
                          End If
                      End If
                    'NM_CLICK hits before TVN_SELCHANGEDA hits
                    Case NM_CLICK ' lParam = lp NMHDR ' Tree_NodeClick
                        'Get the highlighted treeview item (instead of just TVHT_ONITEM)
                        lItem = TreeView_GetHitTest(hTree)
                        If lItem > 0 Then
                            TreeView_GetText hTree, lItem, sText
                            sText = "Selected Node >> " & sText & " <<"
                            A_SetWindowText hTreeLabel, sText
                        End If
                End Select
'''''''''''''''''''''''''''
''''''Btn, menu, keyboard actions
'''''''''''''''''''''''''''
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                If lParam = hBtnCloseTreeView Then
                    ShowWindow hFrmTreeView, SW_HIDE
                    BringWindowToTop hMainForm
                    Exit Function
                End If
            End If
        Case WM_CLOSE
            ShowWindow hFrmTreeView, SW_HIDE
            BringWindowToTop hMainForm
            'DestroyWindow hFrmTreeView
            Exit Function
    End Select
    TreeViewWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Public Sub ProcessContextMenu(lItem As Long, lMnuId As Long)
    On Error GoTo ProcessContextMenu_Error
    
    Dim sText As String
    Dim lLen As Long
    Dim tvFont As Long
    Dim hTmpFont As Long
    Dim lfTreeView As A_LOGFONT
    Dim tmpLF As A_LOGFONT

    Select Case lMnuId
        Case MNUTV_SAVEWITHTAB_ID
            'TreeView_LoadTreeFormFileWithTab hTree, App.Path & "\treesaved.txt"
            TreeView_SaveToFileWithTabs hTree, App.Path & "\treesaved.txt"
            A_MessageBox hFrmTreeView, "Saved to:" & vbCrLf & App.Path & "\treesaved.txt", "Saved Tree", MB_OK
        Case MNUTV_EDITLABEL_ID
            SetFocusApi hTree
            'Returns handle to edit control, temporary , will be destroyed after editting is doen
            Debug.Print "EditLabel=" & TreeView_EditLabel(hTree, lItem)
        Case MNUTV_ADDNEW_ID
            If InputBoxApi(hFrmTreeView, "New Item", "Add Item", "Please enter a text for new item:", sText) = IDOK Then
                lLen = TreeView_Add(hTree, lItem, 0, sText, , 0, 1, 4)
                If lLen > 0 Then
                    TreeView_EnsureVisible hTree, lLen
                    TreeView_SetSelected hTree, lLen
                End If
            End If
        Case MNUTV_REMOVESELECTED_ID
            If TreeView_GetText(hTree, lItem, sText) = True Then
                If A_MessageBox(hFrmTreeView, "Proceed to remove: " & sText, "Delete Node", MB_YESNO Or MB_ICONQUESTION) = IDYES Then
                    TreeView_DeleteItem hTree, lItem
                End If
            End If
        Case MNUTV_REMOVEALL_ID
            TreeView_DeleteAllItems hTree
        Case MNUTV_EXPAND_COLLAPSE_ID
            TreeView_Expand hTree, lItem, TVE_TOGGLE
        Case MNUTV_EMAIL_ID '1
            TreeView_SetStateImage hTree, lItem, NUM_ONE
        Case MNUTV_IMPORTANT_ID '2
            TreeView_SetStateImage hTree, lItem, NUM_TWO
        Case MNUTV_PRINT_ID '3
            TreeView_SetStateImage hTree, lItem, NUM_THREE
        Case MNUTV_UPLOAD_ID '6
            TreeView_SetStateImage hTree, lItem, NUM_SIX
        Case MNUTV_LOCK_ID '5
            TreeView_SetStateImage hTree, lItem, NUM_FIVE
        Case Else
            'Get the menu text
            sText = GetSubMenuText(hMenuTree, lMnuId)
            'Debug.Print "Clicked22=" & sText & "="
            If LenB(sText) > 0 Then
                'Get tree font
                tvFont = A_SendMessage(hTree, WM_GETFONT, 0&, 0&)
                'Store current treeview font in a LOGFONT struct
                'Returns the number of bytes copied
                A_GetObject tvFont, Len(lfTreeView), lfTreeView
                'Create a new font based on the treeview default font
                tmpLF.lfCharSet = lfTreeView.lfCharSet
                tmpLF.lfHeight = lfTreeView.lfHeight
                tmpLF.lfWidth = lfTreeView.lfWidth
                'Bold it so it stands out
                tmpLF.lfWeight = FW_BOLD 'lfTreeView.lfWeight
                
                'Need to add vbNullChar for the ANSI version of lstrcpyn to work
                sText = sText & vbNullChar
                'Copy font name into lfFaceName byte array
                A_lstrcpynStrByte tmpLF.lfFaceName(0), sText, Len(sText)
                hTmpFont = A_CreateFontIndirect(tmpLF)
                lLen = cNodesD.FindNodeIndex(lItem)
                'Debug.Print "=" & lfTreeView.lfCharSet & "=" & lfTreeView.lfHeight & "=" & lfTreeView.lfWidth & "=" & lfTreeView.lfWeight
                'Debug.Print "lItem: " & lItem & " hTmpFont=" & hTmpFont & " Index=" & lLen & " Name=" & StripNulls(StrConv(tmpLF.lfFaceName, vbUnicode)) '& " sText=" & sText
                'Update or Add
                If lLen > -1 Then
                    cNodesD.SetNodeFont hTmpFont, , lLen
                Else
                    cNodesD.AddNodeData lItem, , x_Blue, x_Yellow, hTmpFont
                End If
            End If
    End Select

    Exit Sub
ProcessContextMenu_Error:
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure ProcessContextMenu of Module Form_TreeView"
End Sub

Public Sub BtnShowTreeView_Click()
    If hFrmTreeView = 0 Then
        hFrmTreeView = CreateForm(300, 300, 315, 280, AddressOf TreeViewWndProc, hMainForm, WS_CAPTION Or WS_SYSMENU Or WS_OVERLAPPED, WS_EX_CONTROLPARENT Or WS_EX_WINDOWEDGE, "FormTreeView", "TreeView")
        If hFrmTreeView > 0 Then
            SetupTreeViewControl 0, 0, 310, 180
            hTreeLabel = CreateLabel(hFrmTreeView, "Select an item:", 5, 190, 300, 17)
            hBtnCloseTreeView = CreateCmdButton(hFrmTreeView, "Close", 5, 215, 80, 30, , , , False)
            ShowWndForFirstTime hFrmTreeView
        End If
    Else
        ShowWindow hFrmTreeView, SW_SHOW
    End If
End Sub

Public Sub SetupImageList16()
    Dim hBitmap As Long, hMask As Long
    Dim lImg As Long
    Dim sB As String
    
    If hTvImgListState16 <> 0 Then Exit Sub
    'Create imagelist normal
    'NOTE:
    'Using ILC_COLOR8 flag causes a problem in XP, when CreateDragImage is called, it does not display any images
    hTvImgList16 = ImageList_Create(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CXSMICON), ILC_COLOR16, 2, 10)
    'Load bitmap
    sB = App.Path & "\images\folder-close.bmp"
    hBitmap = A_LoadImage(App.hInstance, sB, IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    lImg = ImageList_Add(hTvImgList16, hBitmap, 0)
DeleteObject hBitmap
    sB = App.Path & "\images\folder-open.bmp"
    hBitmap = A_LoadImage(App.hInstance, sB, IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    lImg = ImageList_Add(hTvImgList16, hBitmap, 0)
    'Debug.Print "Count=" & ImageList_GetImageCount(hTvImgList16)
DeleteObject hBitmap
End Sub

Public Sub SetupImageListState16()
    Dim hBitmap As Long, hMask As Long
    Dim lImg As Long
    Dim sB As String
    
    If hTvImgListState16 <> 0 Then Exit Sub

    'Create imagelist state
    hTvImgListState16 = ImageList_Create(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CXSMICON), ILC_COLOR8, 7, 10)
    
    'passing 0 index to treeview_setimagelist(htree,0,state) will unset the tree state imagelist
    'The ImageList_AddMasked function also adds bitmapped images to a masked image list.
    'This function is similar to ImageList_Add, except that you do not specify a mask bitmap.
    'Instead, you specify a color that the system combines with the image bitmap to
    'automatically generate the masks.
    sB = App.Path & "\images\blank.bmp"
    hBitmap = A_LoadImage(App.hInstance, sB, IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    If hBitmap <> 0 Then lImg = ImageList_Add(hTvImgListState16, hBitmap, 0)
DeleteObject hBitmap
    sB = App.Path & "\images\email.bmp"
    hBitmap = A_LoadImage(App.hInstance, sB, IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    If hBitmap <> 0 Then lImg = ImageList_Add(hTvImgListState16, hBitmap, 0)
DeleteObject hBitmap
    sB = App.Path & "\images\important.bmp"
    hBitmap = A_LoadImage(App.hInstance, sB, IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    If hBitmap <> 0 Then lImg = ImageList_Add(hTvImgListState16, hBitmap, 0)
DeleteObject hBitmap
    sB = App.Path & "\images\print.bmp"
    hBitmap = A_LoadImage(App.hInstance, sB, IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    If hBitmap <> 0 Then lImg = ImageList_Add(hTvImgListState16, hBitmap, 0)
DeleteObject hBitmap
    sB = App.Path & "\images\unknown.bmp"
    hBitmap = A_LoadImage(App.hInstance, sB, IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    If hBitmap <> 0 Then lImg = ImageList_Add(hTvImgListState16, hBitmap, 0)
DeleteObject hBitmap
    sB = App.Path & "\images\lock.bmp"
    hBitmap = A_LoadImage(App.hInstance, sB, IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    If hBitmap <> 0 Then lImg = ImageList_Add(hTvImgListState16, hBitmap, 0)
DeleteObject hBitmap
    sB = App.Path & "\images\upload.bmp"
    hBitmap = A_LoadImage(App.hInstance, sB, IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    If hBitmap <> 0 Then lImg = ImageList_Add(hTvImgListState16, hBitmap, 0)

    'Debug.Print "Count=" & ImageList_GetImageCount(hTvImgListState16)
    DeleteObject hBitmap
    
End Sub

Public Sub SetupTreeViewControl(cx As Long, cy As Long, cW As Long, cH As Long)
    Dim hNode As Long, hNode1 As Long, hNode2 As Long
    Dim tNode As TVITEM
    Dim tStruct As TVINSERTSTRUCT
    Dim sName As String
    Dim sPad As String
    
    Dim lMask As Long
    Dim hBitmap As Long
    Dim hTmpNode As Long, hTmpNode1 As Long
    Dim lIndent As Long


   'Set up image list
    SetupImageList16
    SetupImageListState16
    
    'Set up context menu
    SetupTreeContextMenu
    
    'Create tree
    hTree = CreateTreeView(hFrmTreeView, cx, cy, cW, cH, , , , , , , , , True)
    If hTree <> 0 Then
        'When setting varuaty of fonts for nodes, it is recommanded by MSDN that
        'before adding any items set the CCM_SETVERSION to ensure
        'that the tree does not clip text due to backward compatability
        'when we are changing fonts per item
        'A_SendMessage hTree, CCM_SETVERSION, 5, 0&
    
        'Initialize the nodedata manager
        Set cNodesD = New cNodeData
        cNodesD.SetArraySize 50
            
        'assign imagelist normal
        TreeView_SetImageList hTree, hTvImgList16
        'assign imagelist state
        TreeView_SetImageList hTree, hTvImgListState16, False
        
        lIndent = TreeView_GetIndent(hTree)
        'Debug.Print "lIndent: " & lIndent
        If lIndent < 19 Then TreeView_SetIndent hTree, 19
        
        'This is necessary to stop text from being clipped when a font is changed
        sPad = String$(10, CHAR_SPACE)
        
        hNode = TreeView_Add(hTree, , , "ROOT" & sPad, 0, 0, 1, 1)
        hNode1 = TreeView_Add(hTree, hNode, 0, "Root child" & sPad, , 0, 1, 2)
        hNode2 = TreeView_Add(hTree, hNode, 0, "AHILD 2" & sPad, , 0, 1, 3)
        hTmpNode = TreeView_Add(hTree, hNode, 0, "bome text" & sPad, , 0, 1, 4)
        hTmpNode = TreeView_Add(hTree, hNode, 0, "aome text" & sPad, , 0, 1, 1)
            hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "Sub1 aome" & sPad, , 0, 1, 1)
            hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "Sub2 aome" & sPad, , 0, 1, 1)
            hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "Sub3 aome" & sPad, , 0, 1, 1)
        hTmpNode = TreeView_Add(hTree, hNode, 0, "come text" & sPad, , 0, 1, 1)
            hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "Sub1 come" & sPad, , 0, 1, 1)
            hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "Sub2 come" & sPad, , 0, 1, 1)
                hTmpNode = TreeView_Add(hTree, hTmpNode1, 0, "Sub1 Sub2 come" & sPad, , 0, 1, 1)
                    hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "Sub1 Sub1 Sub2 come" & sPad, , 0, 1, 1)
        hTmpNode = TreeView_Add(hTree, hNode, 0, "fome text" & sPad, , 0, 1, 1)
        hTmpNode = TreeView_Add(hTree, hNode, 0, "eome text" & sPad, , 0, 1, 1)
        hTmpNode = TreeView_Add(hTree, hNode, 0, "home text" & sPad, , 0, 1, 1)
        hTmpNode = TreeView_Add(hTree, hNode, 0, "gome text" & sPad, , 0, 1, 1)
        hTmpNode = TreeView_Add(hTree, hNode, 0, "iome text" & sPad, , 0, 1, 1)
        'Debug.Print "StateSet " & TreeView_SetStateImage(hTree, hNode, 0)
        hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "mome text" & sPad, , 0, 1, 1)
            hTmpNode = TreeView_Add(hTree, hTmpNode1, 0, "oome text" & sPad, , 0, 1, 1)
                hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "rome text" & sPad, , 0, 1, 1)
                    hTmpNode = TreeView_Add(hTree, hTmpNode1, 0, "pome text" & sPad, , 0, 1, 1)
                        hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "wome text" & sPad, , 0, 1, 1)
                            hTmpNode = TreeView_Add(hTree, hTmpNode1, 0, "qome text" & sPad, , 0, 1, 1)
                                hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "yome text" & sPad, , 0, 1, 1)
                                    hTmpNode = TreeView_Add(hTree, hTmpNode1, 0, "vome text" & sPad, , 0, 1, 1)
                                        hTmpNode1 = TreeView_Add(hTree, hTmpNode, 0, "tome text" & sPad, , 0, 1, 1)
        
        TreeView_EnsureVisible hTree, hNode2 'hTmpNode1
        
        'To restore to default
        'TreeView_SetLineColor hTree, CLR_DEFAULT
        'COLORREF = RGB(xxx,xxx,xxx)
        TreeView_SetBkColor hTree, vbWhite
        TreeView_SetTextColor hTree, vbBlue
        TreeView_SetLineColor hTree, vbRed
        'Sorts only the immediate nodes under the given node
        'does not work recursive, so we go with root
        TreeView_SortChildren hTree, hNode
        'Bold a node
        'TreeView_SetState hTree, hNode2, TVIS_BOLD, True
        'Get it's new state
        'If TreeView_GetState(hTree, hNode2, lMask) Then If lMask = TVIS_BOLD Then Debug.Print "state bold"
        'Set some node
        TreeView_SetSelected hTree, hNode2
        UpdateWindow hTree
        'Debug.Print "hTree= " & hTree & " =hNode= " & hNode & " =hNode1= " & hNode1 & " =hNod2= " & hNode2
    End If
End Sub

Public Sub SetupTreeContextMenu()
    Dim mnuItem As MENUITEMINFO
    Dim mInfo As MENUINFO
    
    hMenuTree = CreatePopupMenu()
    If hMenuTree = 0 Then
        Debug.Print "hMenuTree 0: " & hMenuTree
        Exit Sub
    End If
    mnuItem = CreateMenuItem("Change state to email", MNUTV_EMAIL_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_EMAIL_ID, 0, mnuItem
    mnuItem = CreateMenuItem("Change state to Important", MNUTV_IMPORTANT_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_IMPORTANT_ID, 0, mnuItem
    mnuItem = CreateMenuItem("Change state to lock", MNUTV_LOCK_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_LOCK_ID, 0, mnuItem
    mnuItem = CreateMenuItem("Change state to print", MNUTV_PRINT_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_PRINT_ID, 0, mnuItem
    mnuItem = CreateMenuItem("Change state to upload", MNUTV_UPLOAD_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_UPLOAD_ID, 0, mnuItem
    
    mnuItem = CreateMenuItem("-", MNUTV_SEP_ID1, 0&, True)
    A_InsertMenuItem hMenuTree, MNUTV_SEP_ID1, 0, mnuItem

    mnuItem = CreateMenuItem("Edit label", MNUTV_EDITLABEL_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_EDITLABEL_ID, 0, mnuItem
    mnuItem = CreateMenuItem("Expand/Collapse", MNUTV_EXPAND_COLLAPSE_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_EXPAND_COLLAPSE_ID, 0, mnuItem

    mnuItem = CreateMenuItem("-", MNUTV_SEP_ID2, 0&, True)
    A_InsertMenuItem hMenuTree, MNUTV_SEP_ID2, 0, mnuItem

    mnuItem = CreateMenuItem("Insert new", MNUTV_ADDNEW_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_ADDNEW_ID, 0, mnuItem
    mnuItem = CreateMenuItem("Remove selected", MNUTV_REMOVESELECTED_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_REMOVESELECTED_ID, 0, mnuItem
    mnuItem = CreateMenuItem("Remove all", MNUTV_REMOVEALL_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_REMOVEALL_ID, 0, mnuItem

    mnuItem = CreateMenuItem("-", MNUTV_SEP_ID3, 0&, True)
    A_InsertMenuItem hMenuTree, MNUTV_SEP_ID3, 0, mnuItem

    mnuItem = CreateMenuItem("Save with Tab", MNUTV_SAVEWITHTAB_ID, 0&)
    A_InsertMenuItem hMenuTree, MNUTV_SAVEWITHTAB_ID, 0, mnuItem

    mnuItem = CreateMenuItem("-", MNUTV_SEP_ID4, 0&, True)
    A_InsertMenuItem hMenuTree, MNUTV_SEP_ID4, 0, mnuItem

    hMnuFont = CreatePopupMenu
    lFontMenuID = 1100
    'Fill submenus with available system fonts 1100
    A_EnumFonts GetDC(hTree), vbNullString, AddressOf EnumFontProc, 0
    
    mnuItem = CreateMenuItem("Change Node Font to...", MNUTV_CHANGEFONT_ID, hMnuFont)
    A_InsertMenuItem hMenuTree, MNUTV_CHANGEFONT_ID, 0, mnuItem
    
    'Debug.Print "SetMenu: " & SetMenu(Me.hwnd, hMenuTree)

End Sub

''''FONT Enums procs

'Sample call
'A_EnumFonts GetDC(0), vbNullString, AddressOf EnumFontProc, 0
Public Function EnumFontProc(ByVal lplf As Long, ByVal lptm As Long, ByVal dwType As Long, ByVal lpData As Long) As Long
    Dim LF As A_LOGFONT
    Dim FontName As String
    Dim ZeroPos As Long
    Dim hAdded As Long
    Dim mnuItem As MENUITEMINFO
    
    CopyMemory LF, ByVal lplf, LenB(LF)
    FontName = StrConv(LF.lfFaceName, vbUnicode)
    ZeroPos = InStr(1, FontName, Chr$(0))
    If ZeroPos > 0 Then FontName = Left$(FontName, ZeroPos - 1)
    
    lFontMenuID = lFontMenuID + 1
    'Debug.Print "=" & lFontMenuID
    mnuItem = CreateMenuItem(FontName, lFontMenuID, 0&)
    A_InsertMenuItem hMnuFont, lFontMenuID, 0, mnuItem

    EnumFontProc = 1
End Function

'Sample call
'LF = A_LOGFONT
'A_EnumFontFamiliesEx GetDC(0), LF, AddressOf EnumFontFamProc, ByVal 0&, 0
Public Function EnumFontFamProc(lpNLF As A_LOGFONT, lpNTM As A_NEWTEXTMETRIC, ByVal FontType As Long, lParam As Long) As Long
    Dim FaceName As String
    
    'convert the returned string to Unicode
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    'print
    'debug.Print Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
        
    'continue enumeration
    EnumFontFamProc = 1
End Function
