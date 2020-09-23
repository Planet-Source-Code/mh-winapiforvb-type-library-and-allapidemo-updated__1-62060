Attribute VB_Name = "Comctl32_Macros"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Comctl32 and general use macros
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Put together by MH
'Main reference:
'   Comctl header file
'   Brad's VB-32 Programs & Samples for most of ListView and TreeView macros (http://www.mvps.org/btmtz/)
'   VBspeed for the general macros (http://www.xbeat.net/vbspeed/)
'Other macros will be added to this file in the future as time permits

'Comments starting with // (C++ start single line comment) are copied directly from header file

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''IP Address Field macros
'''''''''''''''''''''''''''''''''''''''''''

'// And this is a useful macro for making the IP Address to be passed
'// as a LPARAM.
'b1-b4 Field 0 address to Field 4 address, in byte, returns long
'#define MAKEIPADDRESS(b1,b2,b3,b4)  ((LPARAM)(((DWORD)(b1)<<24)+((DWORD)(b2)<<16)+((DWORD)(b3)<<8)+((DWORD)(b4))))
Public Function IPADDRESS_MAKEIPADDRESS(b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte) As Long
    IPADDRESS_MAKEIPADDRESS = CLng(b1 * (2 ^ 24)) + CLng(b2 * (2 ^ 16)) + CLng(b3 * (2 ^ 8)) + CLng(b4)
End Function

'// The following is a useful macro for passing the range values in the
'// IPM_SETRANGE message.
'#define MAKEIPRANGE(low, high)    ((LPARAM)(WORD)(((BYTE)(high) << 8) + (BYTE)(low)))
Public Function IPADDRESS_MAKEIPRANGE(high As Byte, low As Byte) As Long
    IPADDRESS_MAKEIPRANGE = CLng(CInt((high * (2 ^ 8)) + low))
End Function

'// Get individual number
'#define FIRST_IPADDRESS(x)  ((x>>24) & 0xff) //
Public Function IPADDRESS_FIRST_IPADDRESS(X As Long) As Byte
    IPADDRESS_FIRST_IPADDRESS = (X / (2 ^ 24)) And MAXBYTE
End Function

'#define SECOND_IPADDRESS(x) ((x>>16) & 0xff) // (x / (2 ^ 16)) and MAXBYTE
Public Function IPADDRESS_SECOND_IPADDRESS(X As Long) As Byte
    IPADDRESS_SECOND_IPADDRESS = (X / (2 ^ 16)) And MAXBYTE
End Function

'#define THIRD_IPADDRESS(x)  ((x>>8) & 0xff)  // (x / (2 ^ 8)) and MAXBYTE //2 ^ 8 = 256
Public Function IPADDRESS_THIRD_IPADDRESS(X As Long) As Byte
    IPADDRESS_THIRD_IPADDRESS = (X / (2 ^ 8)) And MAXBYTE
End Function

'#define FOURTH_IPADDRESS(x) (x & 0xff)       // x and MAXBYTE
Public Function IPADDRESS_FOURTH_IPADDRESS(X As Long) As Byte
    IPADDRESS_FOURTH_IPADDRESS = X And MAXBYTE
End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Trackbar macros
'''''''''''''''''''''''''''''''''''''''''''

Public Function TrackBar_SetPageSize(hWnd As Long, lNewSize As Long) As Long
    TrackBar_SetPageSize = A_SendMessage(hWnd, TBM_SETPAGESIZE, 0&, lNewSize)
End Function

Public Function TrackBar_SetLineSize(hWnd As Long, lNewSize As Long) As Long
    TrackBar_SetLineSize = A_SendMessage(hWnd, TBM_SETLINESIZE, 0&, lNewSize)
End Function

Public Function TrackBar_SetRange(hWnd As Long, lMin As Long, lMax As Long, Optional bRedraw As Boolean = False) As Long
    TrackBar_SetRange = A_SendMessage(hWnd, TBM_SETRANGE, CLng(bRedraw), MAKELONG(lMin, lMax))
End Function

Public Sub TrackBar_SetPos(hWnd As Long, lPos As Long, Optional bRedraw As Boolean = False)
    A_SendMessage hWnd, TBM_SETPOS, CLng(bRedraw), lPos
End Sub

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''TreeView macros
'''''''''''''''''''''''''''''''''''''''''''

Public Function TreeView_GetCount(hwndTree As Long) As Long
    TreeView_GetCount = A_SendMessageAnyRef(hwndTree, TVM_GETCOUNT, 0, 0&)
End Function

'Retrieves the immediate parent of this node, not the root
Public Function TreeView_GetParent(hwndTree As Long, lItem As Long) As Long
    TreeView_GetParent = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_PARENT, ByVal lItem)
End Function

'Gets the next node in the same hirearchy, returns NULL/0 if no more nodes are in the hire...
Public Function TreeView_GetNextSibling(hwndTree As Long, lItem As Long) As Long
    TreeView_GetNextSibling = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal lItem)
End Function

'Gets the last node in the same hirearchy starting from a parent node lItem
Public Function TreeView_GetLastSibling(hwndTree As Long, lItem As Long) As Long
    Dim hTemp As Long
    'Get the first child
    hTemp = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_CHILD, ByVal lItem)
    Do While hTemp > 0
        'Keep the last hnode
        TreeView_GetLastSibling = hTemp
        'Get the next sibling
        hTemp = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hTemp)
    Loop
End Function

Public Function TreeView_GetPrevSibling(hwndTree As Long, lItem As Long) As Long
    TreeView_GetPrevSibling = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_PREVIOUS, ByVal lItem)
End Function

Public Function TreeView_IsItemVisible(hwndTree As Long, lItem As Long) As Boolean
    Dim lRect As RECT
    Dim cRect As RECT
    If TreeView_GetItemRect(hTree, lItem, lRect) = True Then
        GetClientRect hTree, cRect
        If lRect.Bottom <= cRect.Top Then Exit Function
        If lRect.Top >= cRect.Bottom Then Exit Function
        If lRect.Right < cRect.Left Then Exit Function
        If lRect.Left > cRect.Right Then Exit Function
        TreeView_IsItemVisible = True
     End If
End Function

Public Function TreeView_GetLastVisible(hwndTree As Long) As Long
    Dim lMask As Long
    Dim pt As POINTAPI
    Dim tvhti As TVHITTESTINFO
    Dim lRect As RECT
    
    'Use approximation
    'TVHT_ONITEM = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
    lMask = TVHT_ONITEM Or TVHT_ONITEMINDENT Or TVHT_ONITEMBUTTON
    GetClientRect hwndTree, lRect
    pt.X = lRect.Left + 10
    pt.Y = lRect.Bottom - 10
    tvhti.pt = pt
    Call A_SendMessageAnyRef(hwndTree, TVM_HITTEST, 0, tvhti)
    If (tvhti.flags And lMask) Then
        TreeView_GetLastVisible = tvhti.hItem
    End If
End Function
'
'Version 4.71. Retrieves the last expanded item in the tree.
'This does not retrieve the last item visible in the tree-view window.
'
Public Function TreeView_GetLastExpanded(hwndTree As Long) As Long
    TreeView_GetLastExpanded = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_LASTVISIBLE, 0)
End Function

'Action flag. This parameter can be one or more of the following values:
'TVE_COLLAPSE       = Collapses the list.
'TVE_COLLAPSERESET  = Collapses the list and removes the child items.
'                       The TVIS_EXPANDEDONCE state flag is reset.
'                       This flag must be used with the TVE_COLLAPSE flag.
'TVE_EXPAND         = Expands the list.
'TVE_EXPANDPARTIAL  = Version 4.70. Partially expands the list.
'                       In this state the child items are visible and the parent item's plus sign (+),
'                       indicating that it can be expanded, is displayed.
'                       This flag must be used in combination with the TVE_EXPAND flag.
'TVE_TOGGLE         = Collapses the list if it is expanded or expands it if it is collapsed.
'Attempting to explicitly set TVIS_EXPANDEDONCE will result in unpredictable behavior.
'
Public Function TreeView_Expand(hwndTree As Long, lItem As Long, TVE_Flag As TreeViewExpandFlags) As Boolean
    TreeView_Expand = A_SendMessageAnyRef(hwndTree, TVM_EXPAND, TVE_Flag, ByVal lItem)
End Function

Public Function TreeView_GetFirstVisible(hwndTree As Long) As Long
    TreeView_GetFirstVisible = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, 0)
End Function

'Retrieves the first visible item that precedes the specified item.
'The specified item must be visible.
'Use the TVM_GETITEMRECT message to determine whether an item is visible.
Public Function TreeView_GetPrevVisible(hwndTree As Long, lItem As Long) As Long
    TreeView_GetPrevVisible = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_PREVIOUSVISIBLE, ByVal lItem)
End Function

'Retrieves the next visible item that follows the specified item.
'The specified item must be visible.
'Use the TVM_GETITEMRECT message to determine whether an item is visible.
Public Function TreeView_GetNextVisible(hwndTree As Long, lItem As Long) As Long
    TreeView_GetNextVisible = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, ByVal lItem)
End Function

Public Function TreeView_GetRoot(hwndTree As Long) As Long
    TreeView_GetRoot = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_ROOT, 0)
End Function

Public Function TreeView_GetChild(hwndTree As Long, lItem As Long) As Long
    TreeView_GetChild = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_CHILD, ByVal lItem)
End Function

'=========

Public Function TreeView_SetToolTips(hwndTree As Long, hwndTT As Long) As Long   ' IE3
    TreeView_SetToolTips = A_SendMessageAnyRef(hwndTree, TVM_SETTOOLTIPS, hwndTT, 0&)
End Function

Public Function TreeView_GetToolTips(hwndTree As Long) As Long   ' IE3
    TreeView_GetToolTips = A_SendMessageAnyRef(hwndTree, TVM_GETTOOLTIPS, 0&, 0&)
End Function

Public Function TreeView_SetScrollTime(hwndTree As Long, uTime As Long) As Long   ' IE4
    TreeView_SetScrollTime = A_SendMessageAnyRef(hwndTree, TVM_SETSCROLLTIME, uTime, 0&)
End Function

Public Function TreeView_GetScrollTime(hwndTree As Long) As Long   ' IE4
    TreeView_GetScrollTime = A_SendMessageAnyRef(hwndTree, TVM_GETSCROLLTIME, 0, 0)
End Function

Public Function TreeView_SetInsertMark(hwndTree As Long, hItem As Long, fAfter As Long) As Boolean   ' IE4
    TreeView_SetInsertMark = A_SendMessageAnyRef(hwndTree, TVM_SETINSERTMARK, fAfter, ByVal hItem)
End Function

Public Function TreeView_DeleteItem(hwndTree As Long, lItem As Long) As Long
    TreeView_DeleteItem = A_SendMessageAnyRef(hwndTree, TVM_DELETEITEM, 0, ByVal lItem)
End Function

Public Function TreeView_DeleteAllItems(hwndTree As Long) As Long
    TreeView_DeleteAllItems = A_SendMessageAnyRef(hwndTree, TVM_DELETEITEM, 0, ByVal 0&)
End Function

Public Function TreeView_EditLabel(hwndTree As Long, lItem As Long) As Long
    TreeView_EditLabel = A_SendMessageAnyRef(hwndTree, TVM_EDITLABEL, 0, ByVal lItem)
End Function

' Ends the editing of a tree-view item's label.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_EndEditLabelNow(hwndTree As Long, Optional lEndSave As Long = 1) As Boolean
  TreeView_EndEditLabelNow = A_SendMessageAnyRef(hwndTree, TVM_ENDEDITLABELNOW, lEndSave, 0)
End Function

'' Retrieves the bounding rectangle for a tree-view item and indicates whether the item is visible.
'' If the item is visible and retrieves the bounding rectangle, the return value is TRUE.
'' Otherwise, the TVM_GETITEMRECT message returns FALSE and does not retrieve
'' the bounding rectangle.
'
Public Function TreeView_GetItemRect(hwndTree As Long, lItem As Long, prc As RECT, Optional OnlyTextRect As Long = 1) As Boolean
    prc.Left = lItem
    TreeView_GetItemRect = A_SendMessageAnyRef(hwndTree, TVM_GETITEMRECT, OnlyTextRect, prc)
End Function

Public Function TreeView_GetItemHeight(hwndTree As Long) As Long
    TreeView_GetItemHeight = A_SendMessageAnyRef(hwndTree, TVM_GETITEMHEIGHT, 0, 0)
End Function

'Returns prev H
Public Function TreeView_SetItemHeight(hwndTree As Long, lHeight As Long) As Long
    TreeView_SetItemHeight = A_SendMessageAnyRef(hwndTree, TVM_SETITEMHEIGHT, ByVal lHeight, 0)
End Function

Public Function TreeView_SetIndent(hwndTree As Long, lIndent As Long) As Long
    TreeView_SetIndent = A_SendMessageAnyRef(hwndTree, TVM_SETINDENT, ByVal lIndent, 0)
End Function

'Use the CLR_DEFAULT value to restore the system default colors.
Public Function TreeView_SetLineColor(hwndTree As Long, lColor As Long) As Long
    TreeView_SetLineColor = A_SendMessageAnyRef(hwndTree, TVM_SETLINECOLOR, 0&, ByVal lColor)
End Function

Public Function TreeView_SetTextColor(hwndTree As Long, lColor As Long) As Long
    TreeView_SetTextColor = A_SendMessageAnyRef(hwndTree, TVM_SETTEXTCOLOR, 0&, ByVal lColor)
End Function

Public Function TreeView_SetBkColor(hwndTree As Long, lColor As Long) As Long
    TreeView_SetBkColor = A_SendMessageAnyRef(hwndTree, TVM_SETBKCOLOR, 0&, ByVal lColor)
End Function

Public Function TreeView_GetBkColor(hwndTree As Long) As Long
    TreeView_GetBkColor = A_SendMessageAnyRef(hwndTree, TVM_GETBKCOLOR, 0&, 0&)
End Function

Public Function TreeView_GetIndent(hwndTree As Long) As Long
    TreeView_GetIndent = A_SendMessageAnyRef(hwndTree, TVM_GETINDENT, 0, 0)
End Function

Public Function TreeView_GetTextColor(hwndTree As Long) As Long
    TreeView_GetTextColor = A_SendMessageAnyRef(hwndTree, TVM_GETTEXTCOLOR, 0&, 0&)
End Function

Public Function TreeView_GetLineColor(hwndTree As Long) As Long
    TreeView_GetLineColor = A_SendMessageAnyRef(hwndTree, TVM_GETLINECOLOR, 0&, 0&)
End Function

Public Function TreeView_EnsureVisible(hwndTree As Long, lItem As Long) As Long
    TreeView_EnsureVisible = A_SendMessageAnyRef(hwndTree, TVM_ENSUREVISIBLE, 0, ByVal lItem)
End Function

Public Function TreeView_VisibleCount(hwndTree As Long) As Long
    TreeView_VisibleCount = A_SendMessageAnyRef(hwndTree, TVM_GETVISIBLECOUNT, 0, 0)
End Function

Public Function TreeView_Count(hwndTree As Long) As Long
    TreeView_Count = A_SendMessageAnyRef(hwndTree, TVM_GETCOUNT, 0, 0)
End Function

Public Function TreeView_GetImageList(hwndTree As Long, Optional bNormal As Boolean = True) As Long
    If bNormal = True Then
        TreeView_GetImageList = A_SendMessageAnyRef(hwndTree, TVM_GETIMAGELIST, TVSIL_NORMAL, 0)
    Else
        TreeView_GetImageList = A_SendMessageAnyRef(hwndTree, TVM_GETIMAGELIST, TVSIL_STATE, 0)
    End If
End Function

'pass 0 for lImageList to set image list to nothing
Public Function TreeView_SetImageList(hwndTree As Long, lImageList As Long, Optional bNormal As Boolean = True) As Long
    If bNormal = True Then
        TreeView_SetImageList = A_SendMessageAnyRef(hwndTree, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal lImageList)
    Else
        TreeView_SetImageList = A_SendMessageAnyRef(hwndTree, TVM_SETIMAGELIST, TVSIL_STATE, ByVal lImageList)
    End If
End Function

Public Function TreeView_Add(hwndTree As Long, _
                            Optional hParent As Long = -1, Optional tviInsert As Long = -1, _
                            Optional sText As String = "", Optional lParam As Long = -1, _
                            Optional lImage As Long = -1, Optional lSelectedImage As Long = -1, _
                            Optional lState As Long = -1, Optional bForcePlusButton As Boolean = False) As Long
    Dim tvis As TVINSERTSTRUCT
    Dim tNode As TVITEM
    Dim lMask As Long

    If LenB(sText) > NUM_ZERO Then
        lMask = TVIF_TEXT
        tNode.cchTextMax = Len(sText)
        tNode.pszText = StrPtr(StrConv(sText, vbFromUnicode))
    End If
    If lParam > NUM_ZERO Then
        lMask = lMask Or TVIF_PARAM
        tNode.lParam = lParam
    End If
    If lImage > NUM_ZERO Then
        lMask = lMask Or TVIF_IMAGE
        tNode.iImage = lImage
    End If
    If lSelectedImage > NUM_ZERO Then
        lMask = lMask Or TVIF_SELECTEDIMAGE
        tNode.iSelectedImage = lSelectedImage
    End If
    If lState > NUM_ZERO Then
        lMask = lMask Or TVIF_HANDLE Or TVIF_STATE
        tNode.state = INDEXTOSTATEIMAGEMASK(lState)
        tNode.stateMask = TVIS_STATEIMAGEMASK
    End If
    If bForcePlusButton = True Then
'The parent window keeps track of whether the item has child items.
'In this case, when the tree-view control needs to display the item,
'the control sends the parent a TVN_GETDISPINFO notification message
'to determine whether the item has child items.
'If the tree-view control has the TVS_HASBUTTONS style, it uses
'this member to determine whether to display the button indicating the
'presence of child items. You can use this member to force the control
'to display the button even though the item does not have any child items
'inserted. This allows you to display the button while minimizing
'the control's memory usage by inserting child items only when the item
'is visible or expanded.
'''This causes errors under XP with themes enabled
        lMask = lMask Or TVIF_CHILDREN
        tNode.cChildren = I_CHILDRENCALLBACK
    End If
    tNode.mask = lMask
    If hParent > NUM_ZERO Then
        tvis.hParent = hParent
    Else
        tvis.hParent = TVI_ROOT
    End If
    If tviInsert > NUM_ZERO Then
        tvis.hInsertAfter = tviInsert
    Else
        tvis.hInsertAfter = TVI_ROOT
    End If
    tvis.Item = tNode
    TreeView_Add = A_SendMessageAnyRef(hwndTree, TVM_INSERTITEM, ByVal 0&, tvis)
End Function

'flags
'Variable that receives information about the results of a hit test. This member can be one or more of the following values:
'TVHT_ABOVE
'   Above the client area.
'TVHT_BELOW
'   Below the client area.
'TVHT_NOWHERE
'   In the client area, but below the last item.
'TVHT_ONITEMBUTTON
'   On the button associated with an item.
'TVHT_ONITEMICON
'   On the bitmap associated with an item.
'TVHT_ONITEMINDENT
'   In the indentation associated with an item.
'TVHT_ONITEMLABEL
'   On the label (string) associated with an item.
'TVHT_ONITEMRIGHT
'   In the area to the right of an item.
'TVHT_ONITEMSTATEICON
'   On the state icon for a tree-view item that is in a user-defined state.
'TVHT_TOLEFT
'   To the left of the client area.
'TVHT_TORIGHT
'   To the right of the client area.

'TVHT_ONITEM
'   On the bitmap or label associated with an item.
Public Function TreeView_GetFromPoint(hwndTree As Long, pt As POINTAPI, Optional TVHT_flags As TreeViewHitTest = TVHT_ONITEM) As Long
    Dim tvhti As TVHITTESTINFO
    
    Call ScreenToClient(hwndTree, pt)
    tvhti.pt = pt
    Call A_SendMessageAnyRef(hwndTree, TVM_HITTEST, 0, tvhti)
    If (tvhti.flags And TVHT_flags) Then
        TreeView_GetFromPoint = tvhti.hItem
    End If
End Function

'TVHT_ONITEM
'   On the bitmap or label associated with an item.
Public Function TreeView_GetHitTest(hwndTree As Long) As Long
    Dim pt As POINTAPI
    Dim tvhti As TVHITTESTINFO
    
    Call GetCursorPos(pt)
    Call ScreenToClient(hwndTree, pt)
    tvhti.pt = pt
    Call A_SendMessageAnyRef(hwndTree, TVM_HITTEST, 0, tvhti)
    If (tvhti.flags And TVHT_ONITEM) Then
        TreeView_GetHitTest = tvhti.hItem
    End If
End Function

Public Function TreeView_GetDropHilite(hwndTree As Long) As Long
    TreeView_GetDropHilite = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_DROPHILITE, 0)
End Function

'Pass 0 to unhighlight the target
Public Function TreeView_SelectDropTarget(hwndTree As Long, lItem As Long) As Long
    TreeView_SelectDropTarget = A_SendMessageAnyRef(hwndTree, TVM_SELECTITEM, TVGN_DROPHILITE, ByVal lItem)
End Function

Public Function TreeView_GetSelected(hwndTree As Long) As Long
    TreeView_GetSelected = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_CARET, 0)
End Function

'' Selects the specified tree-view item, scrolls the item into view, or redraws the item
'' in the style used to indicate the target of a drag-and-drop operation.
'' If hitem is NULL, the selection is removed from the currently selected item, if any.
'' Returns TRUE if successful or FALSE otherwise.
'
Public Function TreeView_SetSelected(hwndTree As Long, lItemToSelect As Long) As Long
    TreeView_SetSelected = A_SendMessageAnyRef(hwndTree, TVM_SELECTITEM, TVGN_CARET, ByVal lItemToSelect)
    SetFocusApi hwndTree
End Function

'-1 = error
'zero = The item has no child items.
'one = The item has one or more child items.
Public Function TreeView_GetNumberOfChildren(hwndTree As Long, lItem As Long) As Long
    Dim lChild As Long
    
    TreeView_GetNumberOfChildren = NUM_ZERO
    lChild = TreeView_GetChild(hwndTree, lItem)
    Do While lChild > NUM_ZERO
        TreeView_GetNumberOfChildren = TreeView_GetNumberOfChildren + NUM_ONE
        lChild = TreeView_GetNextSibling(hwndTree, lChild)
    Loop
End Function

Public Function TreeView_GetcChildren(hwndTree As Long, lItem As Long) As Long
    Dim tVi As TVITEM

    tVi.hItem = lItem
    'Set which property we want
    tVi.mask = TVIF_CHILDREN
    'Get the item with requested properties
    If CBool(A_SendMessageAnyRef(hwndTree, TVM_SETITEM, 0, tVi)) = True Then
        TreeView_GetcChildren = tVi.cChildren
    End If
End Function

Public Function TreeView_SetNumberOfChildren(hwndTree As Long, lItem As Long, Optional lChildern As Long = NUM_ZERO) As Boolean
    Dim tNode As TVITEM
    
    tNode.hItem = lItem
    tNode.mask = TVIF_CHILDREN
    tNode.cChildren = lChildern
    TreeView_SetNumberOfChildren = A_SendMessageAnyRef(hwndTree, TVM_SETITEM, 0, tNode)
End Function

Public Function TreeView_Getlparam(hwndTree As Long, lItem As Long, lParam As Long) As Boolean
    Dim tVi As TVITEM

    tVi.hItem = lItem
    'Set which property we want
    tVi.mask = TVIF_PARAM
    'Get the item with requested properties
    TreeView_Getlparam = A_SendMessageAnyRef(hwndTree, TVM_GETITEM, 0, tVi)
    If TreeView_Getlparam Then lParam = tVi.lParam
End Function

Public Function TreeView_Setlparam(hwndTree As Long, lItem As Long, lParam As Long) As Boolean
    Dim tVi As TVITEM
    
    tVi.hItem = lItem
    tVi.mask = TVIF_PARAM
    tVi.lParam = lParam
    TreeView_Setlparam = A_SendMessageAnyRef(hwndTree, TVM_SETITEM, 0, tVi)
End Function

Public Function TreeView_GetText(hwndTree As Long, lItem As Long, sText As String) As Boolean
    Dim tVi As TVITEM
    Dim bBuf() As Byte
                
    ReDim bBuf(MAX_PATH - NUM_ONE)
    'Set which item we want
    tVi.hItem = lItem
    'Set whuch property we want
    tVi.mask = TVIF_TEXT
    tVi.cchTextMax = MAX_PATH
    'Set pre allocate buffer address
    'In Unicode, we just pass strptr
    tVi.pszText = VarPtr(bBuf(0))
    'Get the item with requested properties
    TreeView_GetText = A_SendMessageAnyRef(hwndTree, TVM_GETITEM, 0, tVi)
    If TreeView_GetText Then sText = StripNulls(String_Macros.String_ByteArrayToString(bBuf))
End Function

Public Function TreeView_SetText(hwndTree As Long, lItem As Long, sText As String) As Boolean
    Dim tVi As TVITEM
    
    tVi.hItem = lItem
    tVi.mask = TVIF_TEXT
    tVi.cchTextMax = Len(sText)
    tVi.pszText = StrPtr(StrConv(sText, vbFromUnicode))
    TreeView_SetText = A_SendMessageAnyRef(hwndTree, TVM_SETITEM, 0, tVi)
End Function

Public Function TreeView_GetImage(hwndTree As Long, lItem As Long, lImage As Long) As Boolean
    Dim tVi As TVITEM

    tVi.hItem = lItem
    tVi.mask = TVIF_IMAGE
    TreeView_GetImage = A_SendMessageAnyRef(hwndTree, TVM_GETITEM, 0, tVi)
    If TreeView_GetImage Then lImage = tVi.iImage
End Function

Public Function TreeView_SetImage(hwndTree As Long, lItem As Long, lImage As Long) As Boolean
    Dim tVi As TVITEM

    tVi.hItem = lItem
    tVi.mask = TVIF_IMAGE
    tVi.iImage = lImage
    TreeView_SetImage = A_SendMessageAnyRef(hwndTree, TVM_SETITEM, 0, tVi)
End Function

Public Function TreeView_GetSelectedImage(hwndTree As Long, lItem As Long, lSelectedImage As Long) As Boolean
    Dim tVi As TVITEM

    tVi.hItem = lItem
    tVi.mask = TVIF_SELECTEDIMAGE
    TreeView_GetSelectedImage = A_SendMessageAnyRef(hwndTree, TVM_GETITEM, 0, tVi)
    If TreeView_GetSelectedImage Then lSelectedImage = tVi.iSelectedImage
End Function

Public Function TreeView_SetSelectedImage(hwndTree As Long, lItem As Long, lSelectedImage As Long) As Boolean
    Dim tVi As TVITEM

    tVi.hItem = lItem
    tVi.mask = TVIF_SELECTEDIMAGE
    tVi.iSelectedImage = lSelectedImage
    TreeView_SetSelectedImage = A_SendMessageAnyRef(hwndTree, TVM_SETITEM, 0, tVi)
End Function

' returns prior color as (COLORREF)
Public Function TreeView_SetInsertMarkColor(hwndTree As Long, clr As Long) As Long   ' IE4
    TreeView_SetInsertMarkColor = A_SendMessageAnyRef(hwndTree, TVM_SETINSERTMARKCOLOR, 0, ByVal clr)
End Function

' returns (COLORREF - RGB)
Public Function TreeView_GetInsertMarkColor(hwndTree As Long) As Long   ' IE4
    TreeView_GetInsertMarkColor = A_SendMessageAnyRef(hwndTree, TVM_GETINSERTMARKCOLOR, 0, 0)
End Function

Public Function TreeView_SetUnicodeFormat(hwndTree As Long, fUnicode As Long) As Boolean   ' IE4
    TreeView_SetUnicodeFormat = A_SendMessageAnyRef(hwndTree, TVM_SETUNICODEFORMAT, fUnicode, 0)
End Function

Public Function TreeView_GetUnicodeFormat(hwndTree As Long) As Boolean   ' IE4
    TreeView_GetUnicodeFormat = A_SendMessageAnyRef(hwndTree, TVM_GETUNICODEFORMAT, 0, 0)
End Function

'' Retrieves the incremental search string for a tree-view control. The tree-view control uses the
'' incremental search string to select an item based on characters typed by the user.
'' Returns the number of characters in the incremental search string.
'' If the tree-view control is not in incremental search mode, the return value is zero.
'
'Public Function TreeView_GetISearchString(hwndTree As Long, lpsz As String) As Boolean
'  TreeView_GetISearchString = SendMessage(hwndTree, TVM_GETISEARCHSTRING, 0, lpsz)
'End Function
'

'' Sorts tree-view items using an application-defined callback function that compares the items.
'' Returns TRUE if successful or FALSE otherwise.
'The callback function specified by lpfnCompare has the following form:
'       int CALLBACK CompareFunc(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort);
'The callback function must return
'   a negative value if the first item should precede the second,
'   a positive value if the first item should follow the second,
'   or zero if the two items are equivalent.
'The lParam1 and lParam2 parameters correspond to the lParam member of the TVITEM structure
'for the two items being compared. The lParamSort parameter corresponds to
'the lParam member of this structure.
'
Public Function TreeView_SortChildrenCB(hwndTree As Long, psort As TVSORTCB) As Boolean
    'fRecurse = Reserved. Must be zero.
    TreeView_SortChildrenCB = A_SendMessageAnyRef(hwndTree, TVM_SORTCHILDRENCB, 0, psort)
End Function

'If you create a tree-view control without an associated image list,
'you cannot use the TVM_CREATEDRAGIMAGE message to create the image to display
'during a drag operation. You must implement your own method of creating a drag cursor.
'
Public Function TreeView_CreateDragImage(hwndTree As Long, lItem As Long) As Long
    TreeView_CreateDragImage = A_SendMessageAnyRef(hwndTree, TVM_CREATEDRAGIMAGE, 0, ByVal lItem)
End Function

' Sorts the child items of the specified parent item in a tree-view control.
' Returns TRUE if successful or FALSE otherwise.
' fRecurse: Value that specifies whether the sorting is recursive.
'            Set fRecurse to TRUE to sort all levels of child items below the parent item.
'            Otherwise, only the parent's immediate children are sorted.
'
Public Function TreeView_SortChildren(hwndTree As Long, hItem As Long, Optional fRecurse As Long = 1) As Boolean
    TreeView_SortChildren = A_SendMessageAnyRef(hwndTree, TVM_SORTCHILDREN, fRecurse, ByVal hItem)
End Function

'When label editing begins, an edit control is created, but not positioned or displayed.
'Before it is displayed, the tree-view control sends its parent window
'an TVN_BEGINLABELEDIT notification message.
'To customize label editing, implement a handler for TVN_BEGINLABELEDIT and
'have it send a TVM_GETEDITCONTROL message to the tree-view control.
'If a label is being edited, the return value will be a handle to the edit control.
'Use this handle to customize the edit control by sending the usual EM_XXX messages.
'
Public Function TreeView_GetEditControl(hwndTree As Long) As Long
    TreeView_GetEditControl = A_SendMessageAnyRef(hwndTree, TVM_GETEDITCONTROL, 0, 0)
End Function

''=============STATE
'
' Returns the index of the specified treeview item's state image.
Public Function TreeView_GetStateImage(hwndTree As Long, lItem As Long) As Long
  Dim tVi As TVITEM

  ' Initialize the struct and get the item's state value.
  ' (TVIF_HANDLE does not need to be specified, it's use is implied...)
  tVi.mask = TVIF_HANDLE Or TVIF_STATE
  tVi.hItem = lItem
  tVi.stateMask = TVIS_STATEIMAGEMASK

  If A_SendMessageAnyRef(hwndTree, TVM_SETITEM, 0, tVi) Then
    TreeView_GetStateImage = STATEIMAGEMASKTOINDEX(tVi.state And TVIS_STATEIMAGEMASK) 'tVi.state 'And TVIS_STATEIMAGEMASK)
  End If
End Function
'
'' Sets the index of the specified treeview item's state image.
'' Returns True if successful, returns False otherwise.
'
Public Function TreeView_SetStateImage(hwndTree As Long, lItem As Long, iState As Long) As Boolean
  Dim tVi As TVITEM

  tVi.mask = TVIF_HANDLE Or TVIF_STATE
  tVi.hItem = lItem
  tVi.stateMask = TVIS_STATEIMAGEMASK
  tVi.state = INDEXTOSTATEIMAGEMASK(iState)

  TreeView_SetStateImage = A_SendMessageAnyRef(hwndTree, TVM_SETITEM, 0, tVi)

End Function

Public Function TreeView_GetState(hwndTree As Long, lItem As Long, lState As Long) As Boolean
    Dim tVi As TVITEM

    tVi.hItem = lItem
    tVi.mask = TVIF_HANDLE Or TVIF_STATE
    TreeView_GetState = A_SendMessageAnyRef(hwndTree, TVM_GETITEM, 0, tVi)
    If TreeView_GetState Then lState = tVi.state
End Function

Public Function TreeView_SetState(hwndTree As Long, lItem As Long, TVIS_NewState As TreeView_ItemStates, fAdd As Boolean) As Boolean
    Dim tVi As TVITEM

    tVi.hItem = lItem
    tVi.mask = TVIF_HANDLE Or TVIF_STATE
    tVi.state = fAdd And TVIS_NewState
    'Indicate what state bits we're changing
    tVi.stateMask = TVIS_NewState
    TreeView_SetState = A_SendMessageAnyRef(hwndTree, TVM_SETITEM, 0, tVi)
End Function

' Returns the one-based index of the specifed state image mask, shifted
' left twelve bits. A common control utility macro.

' Prepares the index of a state image so that a tree view control or list
' view control can use the index to retrieve the state image for an item.

' Allows up to 19 state image indices (the 2 ^ 12 - 2 ^ 30 indices)
'#define INDEXTOSTATEIMAGEMASK(i) ((i) << 12)
'UINT INDEXTOSTATEIMAGEMASK( UINT i);

Public Function INDEXTOSTATEIMAGEMASK(iIndex As Long) As Long
' #define INDEXTOSTATEIMAGEMASK(i) ((i) << 12)
  INDEXTOSTATEIMAGEMASK = iIndex * (2 ^ 12)
End Function

' Returns the state image index from the one-based index state image mask.
' The inverse of INDEXTOSTATEIMAGEMASK.

' A user-defined function (not in Commctrl.h)

Public Function STATEIMAGEMASKTOINDEX(iState As Long) As Long
  STATEIMAGEMASKTOINDEX = iState / (2 ^ 12)
End Function

' Find Routines, recursive
'
Public Function TreeView_FindNodeByText(sFind As String, hRetHNode As Long, _
                                        hwndTree As Long, Optional hStartup As Long = NUM_ZERO, _
                                        Optional vbComp As VbCompareMethod = vbBinaryCompare, _
                                        Optional bEnsureVisible As Boolean = True) As Boolean
    Dim sText As String
    Dim hTemp As Long, hTemp1 As Long, hTemp2 As Long
    
    'If item has been found in prior recursion
    If hRetHNode > NUM_ZERO Then
        TreeView_FindNodeByText = True
        Exit Function
    End If
    
    'Start from root?
    If hStartup = NUM_ZERO Then
        hTemp = TreeView_GetRoot(hwndTree)
    Else
        hTemp = hStartup
    End If
    If hTemp <= NUM_ZERO Then Exit Function
    Do
        'No text, leave
        If TreeView_GetText(hwndTree, hTemp, sText) = False Then
            Debug.Print "No Text=" & sText & "="
            Exit Function
        End If
        Debug.Print "stext=" & sText
        If StrComp(sText, sFind, vbComp) = NUM_ZERO Then
            hRetHNode = hTemp
            TreeView_EnsureVisible hwndTree, hTemp
            TreeView_SetSelected hwndTree, hTemp
            TreeView_FindNodeByText = True
            Exit Function
        End If
        
        'Debug.Print "Next=" & hTemp
        'hTemp1 = hTemp1 + 1
        'If hTemp1 > 20 Then Exit Function
        'Go after children
        hTemp2 = TreeView_GetChild(hwndTree, hTemp)
        If hTemp2 > NUM_ZERO Then
            'Recurse down
            If TreeView_FindNodeByText(sFind, hRetHNode, hwndTree, hTemp2) = True Then
                TreeView_FindNodeByText = True
                Exit Function
            End If
        End If
        'If hTemp1 > 20 Then Exit Function
        'Any siblings
        'TreeView_GetNextSibling
        hTemp = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hTemp)
    Loop While hTemp > 0
End Function

Public Function TreeView_GetFullPath(hwndTree As Long, lItem As Long, Optional sPathSep As String = CHAR_BACK_SLASH) As String
    Dim hParent As Long
    Dim sText As String
    
    If TreeView_GetText(hwndTree, lItem, sText) = True Then TreeView_GetFullPath = sText '& CHAR_BACK_SLASH
    hParent = TreeView_GetParent(hwndTree, lItem)

    Do While hParent > NUM_ZERO
        If TreeView_GetText(hwndTree, hParent, sText) = True Then TreeView_GetFullPath = sText & CHAR_BACK_SLASH & TreeView_GetFullPath
        hParent = TreeView_GetParent(hwndTree, hParent)
    Loop
    
End Function

'Can save entire or a part of tree
Public Sub TreeView_SaveToFileWithTabs(hwndTree As Long, sFileName As String, _
                                            Optional hStartup As Long = NUM_ZERO, _
                                            Optional bAppend As Boolean = False)
    
    Dim hTmpFile As Long
    On Error GoTo TreeView_SaveToFileWithTabs_Error
    If LenB(sFileName) = NUM_ZERO Then Exit Sub
    'File is created if does not exists
    If File_Macros.FileApi_OpenFile(sFileName, hTmpFile, , bAppend) = True Then
        'Start from root?
        If hStartup = NUM_ZERO Then hStartup = TreeView_GetRoot(hwndTree)
        'Recurse and write to file
        SaveTreeWithTab hwndTree, hStartup, hTmpFile
        'close file
        File_Macros.FileApi_CloseFile hTmpFile
    End If

    Exit Sub
TreeView_SaveToFileWithTabs_Error:
    If hTmpFile <> NUM_ZERO Then File_Macros.FileApi_CloseFile hTmpFile
    'Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure TreeView_SaveToFileWithTabs of Module modTVMacros"
End Sub

Private Sub SaveTreeWithTab(hwndTree As Long, lItem As Long, hSaveFile As Long)
    Dim sText As String
    Dim hTemp As Long, hTemp1 As Long, lBytWritten As Long
    
    hTemp = lItem
    Do
        'Write info to file
        If TreeView_GetText(hwndTree, hTemp, sText) Then
            'If no "\" character exists then UBound(Split(Node.FullPath, "\")) = 0
            sText = String$(UBound(Split(TreeView_GetFullPath(hwndTree, hTemp), "\")), vbTab) & sText & vbCrLf
            WriteFileStr hSaveFile, sText, Len(sText), lBytWritten, ByVal CLng(0)
        End If
        hTemp1 = TreeView_GetChild(hwndTree, hTemp)
        If hTemp1 > NUM_ZERO Then
            SaveTreeWithTab hwndTree, hTemp1, hSaveFile
        End If
        hTemp = A_SendMessageAnyRef(hwndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hTemp)
    Loop While hTemp > 0
End Sub

Public Sub TreeView_LoadTreeFormFileWithTab(hWnd As Long, sFile As String)
    Dim text_line As String
    Dim level As Integer
    Dim tree_nodes() As Long
    Dim tmpNode As Long, lineCount As Long, lCounter As Long
    Dim num_nodes As Integer
    Dim arrTmp() As String
    Dim i As Integer, i1 As Integer
    
    On Error GoTo TreeView_LoadTreeFormFileWithTab_Error

    If LenB(sFile) = NUM_ZERO Or FilePresent(sFile) = False Then Exit Sub
    arrTmp = Split(File_Macros.TextFromFileApi(sFile), vbCrLf)
    lineCount = UBound(arrTmp)
    
    TreeView_DeleteAllItems hWnd
    
    For lCounter = 0 To lineCount
        text_line = arrTmp(lCounter)
        If LenB(text_line) > 0 Then
            level = 1
            Do While Left$(text_line, 1) = vbTab
                level = level + 1
                text_line = Mid$(text_line, 2)
            Loop
            'Increase storage
            If level > num_nodes Then
                num_nodes = level
                ReDim Preserve tree_nodes(1 To num_nodes)
            End If
            
            If level = 1 Then
                tree_nodes(level) = TreeView_Add(hWnd, , , text_line, 0, 0, 1, 1)
            Else
                tree_nodes(level) = TreeView_Add(hWnd, tree_nodes(level - 1), , text_line, 0, 0, 1, 1)
            End If
        End If
    Next
    UpdateWindow hWnd

    Exit Sub
TreeView_LoadTreeFormFileWithTab_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TreeView_LoadTreeFormFileWithTab of Module Comctl32_Macros"
End Sub
'
''Loads a treview and rearranges in order of folder first, link last
'Public Sub OpenTreeViewFromFileWithTab(sFileName As String, tTree As TreeView)
'    On Error GoTo Err_Handle
'    Dim text_line As String
'    Dim level As Integer
'    Dim tree_nodes() As Node
'    Dim tmpNode As Node
'    Dim num_nodes As Integer
'    Dim arrTmp() As String
'    Dim i As Integer, i1 As Integer
'
'    'Checks
'    If lstrlen(sFileName) = 0 Then Exit Sub
'    If Dir(sFileName) = "" Then Exit Sub
'    If tTree Is Nothing Then Exit Sub
'
'    'Just for testing
'    'Form1.txtText.Text = GetFileText(Form1.cboOpenSave.Text)
'
'    iFreeFile = FreeFile
'    Open sFileName For Input As iFreeFile
'
'    tTree.Nodes.Clear
''    ResetUniqueID
'
'    Do While Not EOF(iFreeFile)
'        Line Input #iFreeFile, text_line
'        level = 1
'        Do While Left$(text_line, 1) = vbTab
'            level = level + 1
'            text_line = Mid$(text_line, 2)
'        Loop
'        If level > num_nodes Then
'            num_nodes = level
'            ReDim Preserve tree_nodes(1 To num_nodes)
''            Set tNode = Nothing
''        ElseIf level < num_nodes Then
''            Set tNode = Nothing
'        End If
'
'        arrTmp = Split(text_line, PIPE)
'
'        If level = 1 Then
'            If lstrlen(arrTmp(2)) > 0 And lstrlen(arrTmp(3)) > 0 Then
'                Set tree_nodes(level) = tTree.Nodes.Add(, , CreateUniqueID, arrTmp(0), CInt(arrTmp(2)), CInt(arrTmp(3)))
'            Else
'                Set tree_nodes(level) = tTree.Nodes.Add(, , CreateUniqueID, arrTmp(0), CInt(arrTmp(2)))
'            End If
'            tree_nodes(level).Tag = arrTmp(1) 'URL if link , "" if dir
'            'If UBound(arrTmp) > 3 Then
'            tree_nodes(level).Checked = CBool(arrTmp(4))
'        Else
'            If lstrlen(arrTmp(2)) > 0 And lstrlen(arrTmp(3)) > 0 Then
'                'Folder
'                If tTree.Nodes.Count = 1 Then
'                    Set tree_nodes(level) = tTree.Nodes.Add(tree_nodes(level - 1), tvwChild, CreateUniqueID, arrTmp(0), CInt(arrTmp(2)), CInt(arrTmp(3)))
'                Else
'                    i = getFirstChildFolder(tree_nodes(level - 1).Index)
'                    If i > -1 Then
'                        Set tree_nodes(level) = tTree.Nodes.Add(tTree.Nodes(i), tvwNext, CreateUniqueID, arrTmp(0), CInt(arrTmp(2)), CInt(arrTmp(3)))
'                    Else
'                        i1 = getFirstChildFile(tree_nodes(level - 1).Index)
'                        If i1 > -1 Then
'                            Set tree_nodes(level) = tTree.Nodes.Add(tTree.Nodes(i1), tvwPrevious, CreateUniqueID, arrTmp(0), CInt(arrTmp(2)), CInt(arrTmp(3)))
'                        Else
'                            Set tree_nodes(level) = tTree.Nodes.Add(tree_nodes(level - 1), tvwChild, CreateUniqueID, arrTmp(0), CInt(arrTmp(2)), CInt(arrTmp(3)))
'                        End If
'                    End If
'                End If
'                tree_nodes(level).Tag = arrTmp(1)
'                'If UBound(arrTmp) > 3 Then
'                tree_nodes(level).Checked = CBool(arrTmp(4))
'            Else 'Link
'                Set tmpNode = tTree.Nodes.Add(tree_nodes(level - 1), tvwChild, CreateUniqueID, arrTmp(0), CInt(arrTmp(2)))
'                tmpNode.Tag = arrTmp(1)
'                'If UBound(arrTmp) > 3 Then
'                tmpNode.Checked = CBool(arrTmp(4))
'            End If
'            'tree_nodes(level).EnsureVisible
'        End If
'    Loop
'    tTree.Nodes.item(1).Expanded = True
''    LoadNodesFromFileWithTab
'
'    Close iFreeFile
'    Erase tree_nodes
'    Set tmpNode = Nothing
'    'showErrors "Complete! Loading nodes from text file, structured with tabs", vbInformation, "Load TreeView"
'    Exit Sub
'Err_Handle:
'    Close iFreeFile
'    showErrors "OpenTreeViewFromFileWithTab " & Err.Description
'End Sub

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''ListView macros
'''''''''''''''''''''''''''''''''''''''''''

Public Sub ListView_SetView(hWnd As Long, dwView As Long)
    Dim dwStyle As Long
    
    'Get current style
    dwStyle = A_GetWindowLong(hWnd, GWL_STYLE)
    'Only set the window style if the view bits have changed
    If (dwStyle And LVS_TYPEMASK) <> dwView Then
        A_SetWindowLong hWnd, GWL_STYLE, (dwStyle And (Not LVS_TYPEMASK)) Or dwView
    End If
End Sub

Public Function ListView_GetView(hWnd As Long) As Long
    Dim lView As Long
    lView = A_GetWindowLong(hWnd, GWL_STYLE)
    If lView And ListViewStyles.LVS_ICON Then
        ListView_GetView = ListViewStyles.LVS_ICON
        Exit Function
    End If
    If lView And ListViewStyles.LVS_SMALLICON Then
        ListView_GetView = ListViewStyles.LVS_SMALLICON
        Exit Function
    End If
    If lView And ListViewStyles.LVS_REPORT Then
        ListView_GetView = ListViewStyles.LVS_REPORT
        Exit Function
    End If
    If lView And ListViewStyles.LVS_LIST Then
        ListView_GetView = ListViewStyles.LVS_LIST
        Exit Function
    End If
    
End Function

Public Function ListView_AddColumn(hWnd As Long, iCol As Long, sText As String, _
                                    Optional lMask As ListViewColumnMasks = LVCF_TEXT Or LVCF_FMT Or LVCF_SUBITEM, _
                                    Optional lvColumnFormat As ListViewColumnFormats = LVCFMT_LEFT, _
                                    Optional lWidth As Long = -1, _
                                    Optional lImage As Long = -1) As Long
    Dim pcol As LVCOLUMN
    
    pcol.cchTextMax = Len(sText)
    pcol.pszText = StrPtr(StrConv(sText, vbFromUnicode))
    If lWidth > -1 Then
        lMask = lMask Or LVCF_WIDTH
        pcol.cx = lWidth
    End If
    If lImage > -1 Then
        lMask = lMask Or LVCF_IMAGE
        pcol.iImage = lImage
    End If
    pcol.mask = lMask
    pcol.fmt = lvColumnFormat
    
    ListView_AddColumn = A_SendMessageAnyRef(hWnd, LVM_INSERTCOLUMN, ByVal iCol, pcol)
End Function

'Using LVITEM version prior to 6
'Returns the index of the new item if successful, or -1 otherwise.
Public Function ListView_AddItemv5(hWnd As Long, sText As String, _
                                    Optional lImage As Long = -1, _
                                    Optional lIndent As Long = -1, _
                                    Optional lParam As Long = -1) As Long
    Dim pitem As LVITEM
    Dim lMask As Long
    
    pitem.cchTextMax = Len(sText)
    pitem.pszText = StrPtr(StrConv(sText, vbFromUnicode))
    lMask = LVIF_TEXT
    If lImage > -1 Then
        lMask = lMask Or LVIF_IMAGE
        pitem.iImage = lImage
        
    End If
    If lIndent > -1 Then
        lMask = lMask Or LVIF_INDENT
        pitem.iIndent = lIndent
    End If
    If lParam > -1 Then
        lMask = lMask Or LVIF_PARAM
        pitem.lParam = lParam
    End If
    pitem.mask = lMask
    'Must be zero, can not set subitems here
    pitem.iSubItem = 0
    pitem.state = 0
    pitem.stateMask = 0
    
    ListView_AddItemv5 = A_SendMessageAnyRef(hWnd, LVM_INSERTITEM, 0&, pitem)
End Function

'Using LVITEM version 6 with Grouping style
'LVIF_COLUMNS flage and cColumns member = Version 6.0 Number of tile view columns to display for this item.
Public Function ListView_AddItemv6(hWnd As Long, sText As String, _
                                    Optional lImage As Long = -1, _
                                    Optional lIndent As Long = -1, _
                                    Optional lParam As Long = -1, _
                                    Optional lGroupID As Long = -1, _
                                    Optional cColumns As Long = -1) As Long
    Dim pitem As LVITEMv6
    Dim lMask As Long
    
    pitem.cchTextMax = Len(sText)
    pitem.pszText = StrPtr(StrConv(sText, vbFromUnicode))
    lMask = LVIF_TEXT
    If lImage > -1 Then
        lMask = lMask Or LVIF_IMAGE
        pitem.iImage = lImage
    End If
    If lIndent > -1 Then
        lMask = lMask Or LVIF_INDENT
        pitem.iIndent = lIndent
    End If
    If lParam > -1 Then
        lMask = lMask Or LVIF_PARAM
        pitem.lParam = lParam
    End If
    If cColumns > -1 Then
        lMask = lMask Or LVIF_COLUMNS
        pitem.cColumns = cColumns
    End If
    If lGroupID > -1 Then
        lMask = lMask Or LVIF_GROUPID
        pitem.iGroupId = lGroupID
    End If

    pitem.mask = lMask
    'Must be zero, can not set subitems here
    pitem.iSubItem = 0
    pitem.state = 0
    pitem.stateMask = 0
    
    ListView_AddItemv6 = A_SendMessageAnyRef(hWnd, LVM_INSERTITEM, 0&, pitem)
End Function

'Subitem index starts at one
Public Function ListView_AddSubItem(hWnd As Long, sText As String, lItemIndex As Long, lSubIndex As Long, _
                                    Optional lImage As Long = -1) As Boolean
    Dim pitem As LVITEM
    
    'The only valid flags for subitems
    pitem.mask = LVIF_TEXT 'LVIF_STATE
    If lImage > -1 Then
        pitem.mask = pitem.mask Or LVIF_IMAGE
        pitem.iImage = lImage
    End If
    pitem.cchTextMax = Len(sText)
    pitem.pszText = StrPtr(StrConv(sText, vbFromUnicode))
    pitem.iItem = lItemIndex
    pitem.iSubItem = lSubIndex
    
    ListView_AddSubItem = CBool(A_SendMessageAnyRef(hWnd, LVM_SETITEM, 0&, pitem))
End Function

Public Function ListView_GetBkColor(hWnd As Long) As Long
    ListView_GetBkColor = A_SendMessageAnyRef(hWnd, LVM_GETBKCOLOR, 0&, 0&)
End Function

Public Function ListView_SetBkColor(hWnd As Long, clrBk As Long) As Boolean
    ListView_SetBkColor = A_SendMessageAnyRef(hWnd, LVM_SETBKCOLOR, 0&, ByVal clrBk)
End Function

Public Function ListView_GetImageList(hWnd As Long, LVSIL_ImageList As ListViewImageListStyle) As Long
    ListView_GetImageList = A_SendMessageAnyRef(hWnd, LVM_GETIMAGELIST, LVSIL_ImageList, 0&)
End Function

Public Function ListView_SetImageList(hWnd As Long, himl As Long, LVSIL_ImageList As ListViewImageListStyle) As Long
    ListView_SetImageList = A_SendMessageAnyRef(hWnd, LVM_SETIMAGELIST, LVSIL_ImageList, ByVal himl)
End Function

Public Function ListView_GetItemCount(hWnd As Long) As Long
    ListView_GetItemCount = A_SendMessageAnyRef(hWnd, LVM_GETITEMCOUNT, 0&, 0&)
End Function

Public Function ListView_GetItem(hWnd As Long, pitem As LVITEM) As Boolean
    ListView_GetItem = A_SendMessageAnyRef(hWnd, LVM_GETITEM, 0&, pitem)
End Function

Public Function ListView_SetItem(hWnd As Long, pitem As LVITEM) As Boolean
    ListView_SetItem = A_SendMessageAnyRef(hWnd, LVM_SETITEM, 0&, pitem)
End Function

Public Function ListView_InsertItem(hWnd As Long, pitem As LVITEM) As Long
    ListView_InsertItem = A_SendMessageAnyRef(hWnd, LVM_INSERTITEM, 0&, pitem)
End Function

Public Function ListView_DeleteItem(hWnd As Long, i As Long) As Boolean
    ListView_DeleteItem = A_SendMessageAnyRef(hWnd, LVM_DELETEITEM, ByVal i, 0&)
End Function

Public Function ListView_DeleteAllItems(hWnd As Long) As Boolean
    ListView_DeleteAllItems = A_SendMessageAnyRef(hWnd, LVM_DELETEALLITEMS, 0&, 0&)
End Function

Public Function ListView_GetCallbackMask(hWnd As Long) As Long   ' LVStyles
    ListView_GetCallbackMask = A_SendMessageAnyRef(hWnd, LVM_GETCALLBACKMASK, 0&, 0&)
End Function

Public Function ListView_SetCallbackMask(hWnd As Long, mask As ListViewStyles) As Boolean
    ListView_SetCallbackMask = A_SendMessageAnyRef(hWnd, LVM_SETCALLBACKMASK, ByVal mask, 0&)
End Function

Public Function ListView_GetNextItem(hWnd As Long, i As Long, flags As ListViewNextItemFlags) As Long
    ListView_GetNextItem = A_SendMessageAnyRef(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal MAKELPARAM(flags, 0&))
End Function

Public Function ListView_FindItem(hWnd As Long, iStart, plvfi As LVFINDINFO) As Long
    ListView_FindItem = A_SendMessageAnyRef(hWnd, LVM_FINDITEM, ByVal iStart, plvfi)
End Function

Public Function ListView_GetItemRect(hWnd As Long, i As Long, prc As RECT, Optional code As ListViewItemRectFlags = LVIR_BOUNDS) As Boolean
'prc
'   [in, out] Pointer to a RECT structure that receives the bounding rectangle. When the message is sent, the left member of this structure is used to specify the portion of the list-view item from which to retrieve the bounding rectangle. It must be set to one of the following values:
'LVIR_BOUNDS
'   Returns the bounding rectangle of the entire item, including the icon and label.
'LVIR_ICON
'   Returns the bounding rectangle of the icon or small icon.
'LVIR_LABEL
'   Returns the bounding rectangle of the item text.
'LVIR_SELECTBOUNDS
'   Returns the union of the LVIR_ICON and LVIR_LABEL rectangles, but excludes columns in report view.
    prc.Left = code
    ListView_GetItemRect = A_SendMessageAnyRef(hWnd, LVM_GETITEMRECT, ByVal i, prc)
End Function

Public Function ListView_SetItemPosition(hwndLV As Long, iWorkArea As Long, X As Long, Y As Long) As Boolean
    ListView_SetItemPosition = A_SendMessageAnyRef(hwndLV, LVM_SETITEMPOSITION, ByVal iWorkArea, ByVal MAKELPARAM(X, Y))
End Function

Public Function ListView_GetItemPosition(hwndLV As Long, iWorkArea As Long, ppt As POINTAPI) As Boolean
    ListView_GetItemPosition = A_SendMessageAnyRef(hwndLV, LVM_GETITEMPOSITION, ByVal iWorkArea, ppt)
End Function

Public Function ListView_GetStringWidth(hwndLV As Long, psz As String) As Long
    ListView_GetStringWidth = A_SendMessageStr(hwndLV, LVM_GETSTRINGWIDTH, 0&, psz)
End Function

'Returns the index of the item at the specified position, if any, or -1 otherwise.
Public Function ListView_HitTest(hwndLV As Long) As Long
    Dim pt As POINTAPI
    Dim pinfo As LVHITTESTINFO
    
    Call GetCursorPos(pt)
    Call ScreenToClient(hwndLV, pt)
    pinfo.pt = pt
    Call A_SendMessageAnyRef(hwndLV, LVM_HITTEST, 0&, pinfo)
    If (pinfo.flags And LVHT_ONITEM) Then
        ListView_HitTest = pinfo.iItem
    End If
End Function

'Returns the index of the item or subitem tested, if any, or -1 otherwise. If an item or subitem is at the given coordinates, the fields of the LVHITTESTINFO structure will be filled with the applicable hit information.
Public Function ListView_SubItemHitTest(hWnd As Long) As Long
    Dim pt As POINTAPI
    Dim pinfo As LVHITTESTINFO
    
    Call GetCursorPos(pt)
    Call ScreenToClient(hWnd, pt)
    pinfo.pt = pt
    ListView_SubItemHitTest = A_SendMessageAnyRef(hWnd, LVM_SUBITEMHITTEST, 0&, pinfo)
End Function

Public Function ListView_EnsureVisible(hwndLV As Long, i As Long, fPartialOK As Boolean) As Boolean
    ListView_EnsureVisible = A_SendMessageAnyRef(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal MAKELPARAM(Abs(fPartialOK), 0&))
End Function

Public Function ListView_Scroll(hwndLV As Long, dx As Long, dy As Long) As Boolean
    ListView_Scroll = A_SendMessageAnyRef(hwndLV, LVM_SCROLL, ByVal dx, ByVal dy)
End Function

Public Function ListView_RedrawItems(hwndLV As Long, iFirst As Long, iLast As Long) As Boolean
    ListView_RedrawItems = A_SendMessageAnyRef(hwndLV, LVM_REDRAWITEMS, ByVal iFirst, ByVal iLast)
End Function

Public Function ListView_Arrange(hwndLV As Long, code As ListViewArrangeFlags) As Boolean
    ListView_Arrange = A_SendMessageAnyRef(hwndLV, LVM_ARRANGE, ByVal code, 0&)
End Function

Public Function ListView_EditLabel(hwndLV As Long, i As Long) As Long
    ListView_EditLabel = A_SendMessageAnyRef(hwndLV, LVM_EDITLABEL, ByVal i, 0&)
End Function

Public Function ListView_GetEditControl(hwndLV As Long) As Long
    ListView_GetEditControl = A_SendMessageAnyRef(hwndLV, LVM_GETEDITCONTROL, 0&, 0&)
End Function

Public Function ListView_GetColumn(hWnd As Long, iCol As Long, pcol As LVCOLUMN) As Boolean
    ListView_GetColumn = A_SendMessageAnyRef(hWnd, LVM_GETCOLUMN, ByVal iCol, pcol)
End Function

Public Function ListView_SetColumn(hWnd As Long, iCol As Long, pcol As LVCOLUMN) As Boolean
    ListView_SetColumn = A_SendMessageAnyRef(hWnd, LVM_SETCOLUMN, ByVal iCol, pcol)
End Function

Public Function ListView_InsertColumn(hWnd As Long, iCol As Long, pcol As LVCOLUMN) As Long
    ListView_InsertColumn = A_SendMessageAnyRef(hWnd, LVM_INSERTCOLUMN, ByVal iCol, pcol)
End Function

Public Function ListView_DeleteColumn(hWnd As Long, iCol As Long) As Boolean
    ListView_DeleteColumn = A_SendMessageAnyRef(hWnd, LVM_DELETECOLUMN, ByVal iCol, 0&)
End Function

Public Function ListView_GetColumnWidth(hWnd As Long, iCol As Long) As Long
    ListView_GetColumnWidth = A_SendMessageAnyRef(hWnd, LVM_GETCOLUMNWIDTH, ByVal iCol, 0&)
End Function

Public Function ListView_SetColumnWidth(hWnd As Long, iCol As Long, cx As Long) As Boolean
    ListView_SetColumnWidth = A_SendMessageAnyRef(hWnd, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal MAKELPARAM(cx, 0&))
End Function

Public Function ListView_GetHeader(hWnd As Long) As Long
    ListView_GetHeader = A_SendMessageAnyRef(hWnd, LVM_GETHEADER, 0&, 0&)
End Function

Public Function ListView_CreateDragImage(hWnd As Long, i As Long, lpptUpLeft As POINTAPI) As Long
    ListView_CreateDragImage = A_SendMessageAnyRef(hWnd, LVM_CREATEDRAGIMAGE, ByVal i, lpptUpLeft)
End Function

Public Function ListView_GetViewRect(hWnd As Long, prc As RECT) As Boolean
    ListView_GetViewRect = A_SendMessageAnyRef(hWnd, LVM_GETVIEWRECT, 0&, prc)
End Function

Public Function ListView_GetTextColor(hWnd As Long) As Long
    ListView_GetTextColor = A_SendMessageAnyRef(hWnd, LVM_GETTEXTCOLOR, 0&, 0&)
End Function

Public Function ListView_SetTextColor(hWnd As Long, clrText As Long) As Boolean
    ListView_SetTextColor = A_SendMessageAnyRef(hWnd, LVM_SETTEXTCOLOR, 0&, ByVal clrText)
End Function

Public Function ListView_GetTextBkColor(hWnd As Long) As Long
    ListView_GetTextBkColor = A_SendMessageAnyRef(hWnd, LVM_GETTEXTBKCOLOR, 0&, 0&)
End Function

Public Function ListView_SetTextBkColor(hWnd As Long, clrTextBk As Long) As Boolean
    ListView_SetTextBkColor = A_SendMessageAnyRef(hWnd, LVM_SETTEXTBKCOLOR, 0&, ByVal clrTextBk)
End Function

Public Function ListView_GetTopIndex(hwndLV As Long) As Long
    ListView_GetTopIndex = A_SendMessageAnyRef(hwndLV, LVM_GETTOPINDEX, 0&, 0&)
End Function

Public Function ListView_GetCountPerPage(hwndLV As Long) As Long
    ListView_GetCountPerPage = A_SendMessageAnyRef(hwndLV, LVM_GETCOUNTPERPAGE, 0&, 0&)
End Function

Public Function ListView_GetOrigin(hwndLV As Long, ppt As POINTAPI) As Boolean
    ListView_GetOrigin = A_SendMessageAnyRef(hwndLV, LVM_GETORIGIN, 0&, ppt)
End Function

Public Function ListView_Update(hwndLV As Long, i As Long) As Boolean
    ListView_Update = A_SendMessageAnyRef(hwndLV, LVM_UPDATE, ByVal i, 0&)
End Function

Public Function ListView_SetItemState(hwndLV As Long, i As Long, state As ListViewItemStates, mask As ListViewItemFlag) As Boolean
    Dim lvi As LVITEM
    lvi.state = state
    lvi.stateMask = mask
    ListView_SetItemState = A_SendMessageAnyRef(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
End Function

Public Function ListView_GetItemState(hwndLV As Long, i As Long, mask As ListViewItemStates) As Long   ' LVITEM_state
    ListView_GetItemState = A_SendMessageAnyRef(hwndLV, LVM_GETITEMSTATE, ByVal i, ByVal mask)
End Function

Public Function ListView_GetCheckState(hwndLV As Long, iIndex As Long) As Long   ' updated
    Dim dwState As Long
    dwState = A_SendMessageAnyRef(hwndLV, LVM_GETITEMSTATE, ByVal iIndex, ByVal LVIS_STATEIMAGEMASK)
    ListView_GetCheckState = (dwState \ 2 ^ 12) - 1
    '((((UINT)(a_SendMessage(hwndLV, LVM_GETITEMSTATE, ByVal i, LVIS_STATEIMAGEMASK))) >> 12) -1)
End Function

'Returns the number of characters in the pszText member of the LVITEM structure.
Public Function ListView_GetItemText(hwndLV As Long, lItemIndex As Long, _
                                pszText As String, Optional lSubItemIndex As Long = 0) As Long
    Dim lvi As LVITEM
    Dim bBuf() As Byte
    
    ReDim bBuf(MAX_PATH - NUM_ONE)
    lvi.iSubItem = lSubItemIndex
    lvi.cchTextMax = MAX_PATH
    lvi.pszText = VarPtr(bBuf(0))
    ListView_GetItemText = A_SendMessageAnyRef(hwndLV, LVM_GETITEMTEXT, ByVal lItemIndex, lvi)
    If ListView_GetItemText > 0 Then pszText = StripNulls(String_Macros.String_ByteArrayToString(bBuf))
End Function

Public Function ListView_SetItemText(hwndLV As Long, lItemIndex As Long, pszText As String, Optional lSubItemIndex As Long = 0) As Boolean
    Dim lvi As LVITEM
    
    lvi.iSubItem = lSubItemIndex
    lvi.cchTextMax = Len(pszText)
    lvi.pszText = StrPtr(StrConv(pszText, vbFromUnicode))
    ListView_SetItemText = CBool(A_SendMessageAnyRef(hwndLV, LVM_SETITEMTEXT, ByVal lItemIndex, lvi))
End Function

Public Sub ListView_SetItemCount(hwndLV As Long, cItems As Long)
    A_SendMessageAnyRef hwndLV, LVM_SETITEMCOUNT, ByVal cItems, 0
End Sub

Public Sub ListView_SetItemCountEx(hwndLV As Long, cItems As Long, dwFlags As Long)
    A_SendMessageAnyRef hwndLV, LVM_SETITEMCOUNT, ByVal cItems, ByVal dwFlags
End Sub

Public Function ListView_SortItems(hwndLV As Long, pfnCompare As Long, lParamSort As Long) As Boolean
    ListView_SortItems = A_SendMessageAnyRef(hwndLV, LVM_SORTITEMS, ByVal lParamSort, ByVal pfnCompare)
End Function

Public Sub ListView_SetItemPosition32(hwndLV As Long, i As Long, X As Long, Y As Long)
    Dim ptNewPos As POINTAPI
    ptNewPos.X = X
    ptNewPos.Y = Y
    A_SendMessageAnyRef hwndLV, LVM_SETITEMPOSITION32, ByVal i, ptNewPos
End Sub

Public Function ListView_GetSelectedCount(hwndLV As Long) As Long
    ListView_GetSelectedCount = A_SendMessageAnyRef(hwndLV, LVM_GETSELECTEDCOUNT, 0&, 0&)
End Function

Public Function ListView_GetItemSpacing(hwndLV As Long, fSmall As Boolean) As Long
    ListView_GetItemSpacing = A_SendMessageAnyRef(hwndLV, LVM_GETITEMSPACING, ByVal fSmall, 0&)
End Function

Public Function ListView_GetISearchString(hwndLV As Long, lpsz As String) As Boolean
    ListView_GetISearchString = A_SendMessageStr(hwndLV, LVM_GETISEARCHSTRING, 0&, lpsz)
End Function

' =============================================================
' the next three macros are user-defined

' Returns the index of the item that is selected and has the focus rectangle
Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
    ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
End Function

' Selects the specified item and gives it the focus rectangle.
' does not de-select any currently selected items
Public Function ListView_SetSelectedItem(hwndLV As Long, i As Long) As Boolean
    ListView_SetSelectedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, _
                                                    LVIS_FOCUSED Or LVIS_SELECTED)
End Function

' Selects all listview items. The item with the focus rectangle maintains it.
Public Function ListView_SelectAll(hwndLV As Long) As Boolean
    ListView_SelectAll = ListView_SetItemState(hwndLV, -1, LVIS_SELECTED, LVIS_SELECTED)
End Function

' // -1 for cx and cy means we'll use the default (system settings)
' // 0 for cx or cy means use the current setting (allows you to change just one param)
Public Function ListView_SetIconSpacing(hwndLV As Long, cx As Long, cy As Long) As Long
    ListView_SetIconSpacing = A_SendMessageAnyRef(hwndLV, LVM_SETICONSPACING, 0&, ByVal MAKELONG(cx, cy))
End Function

Public Function ListView_SetExtendedListViewStyle(hwndLV As Long, dw As Long) As Long
    ListView_SetExtendedListViewStyle = A_SendMessageAnyRef(hwndLV, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal dw)
End Function

Public Function ListView_GetExtendedListViewStyle(hwndLV As Long) As Long
    ListView_GetExtendedListViewStyle = A_SendMessageAnyRef(hwndLV, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
End Function

Public Function ListView_GetSubItemRect(hWnd As Long, iItem As Long, iSubItem As Long, _
                                        LVIR_flag As Long, prc As RECT) As Boolean
    prc.Top = iSubItem
    prc.Left = LVIR_flag
    ListView_GetSubItemRect = A_SendMessageAnyRef(hWnd, LVM_GETSUBITEMRECT, ByVal iItem, prc)
End Function

Public Function ListView_SetColumnOrderArray(hWnd As Long, iCount As Long, lpiArray As Long) As Boolean
    ListView_SetColumnOrderArray = A_SendMessageAnyRef(hWnd, LVM_SETCOLUMNORDERARRAY, ByVal iCount, lpiArray)
End Function

Public Function ListView_GetColumnOrderArray(hWnd As Long, iCount As Long, lpiArray As Long) As Boolean
    ListView_GetColumnOrderArray = A_SendMessageAnyRef(hWnd, LVM_GETCOLUMNORDERARRAY, ByVal iCount, lpiArray)
End Function

Public Function ListView_SetHotItem(hWnd As Long, i As Long) As Long
    ListView_SetHotItem = A_SendMessageAnyRef(hWnd, LVM_SETHOTITEM, ByVal i, 0&)
End Function

Public Function ListView_GetHotItem(hWnd As Long) As Long
    ListView_GetHotItem = A_SendMessageAnyRef(hWnd, LVM_GETHOTITEM, 0&, 0&)
End Function

Public Function ListView_SetHotCursor(hWnd As Long, hcur As Long) As Long
    ListView_SetHotCursor = A_SendMessageAnyRef(hWnd, LVM_SETHOTCURSOR, 0&, ByVal hcur)
End Function

Public Function ListView_GetHotCursor(hWnd As Long) As Long
    ListView_GetHotCursor = A_SendMessageAnyRef(hWnd, LVM_GETHOTCURSOR, 0&, 0&)
End Function

Public Function ListView_ApproximateViewRect(hWnd As Long, iWidth As Long, _
                                                iHeight As Long, iCount As Long) As Long
    ListView_ApproximateViewRect = A_SendMessageAnyRef(hWnd, _
                                                LVM_APPROXIMATEVIEWRECT, _
                                                ByVal iCount, _
                                                ByVal MAKELPARAM(iWidth, iHeight))
End Function

Public Function ListView_SetUnicodeFormat(hWnd As Long, fUnicode As Boolean) As Boolean
    ListView_SetUnicodeFormat = A_SendMessageAnyRef(hWnd, LVM_SETUNICODEFORMAT, ByVal fUnicode, 0&)
End Function

Public Function ListView_GetUnicodeFormat(hWnd As Long) As Boolean
    ListView_GetUnicodeFormat = A_SendMessageAnyRef(hWnd, LVM_GETUNICODEFORMAT, 0&, 0&)
End Function

Public Function ListView_SetExtendedListViewStyleEx(hwndLV As Long, dwMask As Long, dw As Long) As Long
    ListView_SetExtendedListViewStyleEx = A_SendMessageAnyRef(hwndLV, LVM_SETEXTENDEDLISTVIEWSTYLE, _
                                                                                    ByVal dwMask, ByVal dw)
End Function

Public Function ListView_SetWorkAreas(hWnd As Long, nWorkAreas As Long, prc() As RECT) As Boolean
    ListView_SetWorkAreas = A_SendMessageAnyRef(hWnd, LVM_SETWORKAREAS, ByVal nWorkAreas, prc(0&))
End Function

Public Function ListView_GetWorkAreas(hWnd As Long, nWorkAreas, prc() As RECT) As Boolean
    ListView_GetWorkAreas = A_SendMessageAnyRef(hWnd, LVM_GETWORKAREAS, ByVal nWorkAreas, prc(0&))
End Function

Public Function ListView_GetNumberOfWorkAreas(hWnd As Long, pnWorkAreas As Long) As Boolean
    ListView_GetNumberOfWorkAreas = A_SendMessageAnyRef(hWnd, LVM_GETNUMBEROFWORKAREAS, 0&, pnWorkAreas)
End Function

Public Function ListView_GetSelectionMark(hWnd As Long) As Long
    ListView_GetSelectionMark = A_SendMessageAnyRef(hWnd, LVM_GETSELECTIONMARK, 0&, 0&)
End Function

Public Function ListView_SetSelectionMark(hWnd As Long, i As Long) As Long
    ListView_SetSelectionMark = A_SendMessageAnyRef(hWnd, LVM_SETSELECTIONMARK, 0&, ByVal i)
End Function

Public Function ListView_SetHoverTime(hwndLV As Long, dwHoverTimeMs As Long) As Long
    ListView_SetHoverTime = A_SendMessageAnyRef(hwndLV, LVM_SETHOVERTIME, 0&, ByVal dwHoverTimeMs)
End Function

Public Function ListView_GetHoverTime(hwndLV As Long) As Long
    ListView_GetHoverTime = A_SendMessageAnyRef(hwndLV, LVM_GETHOVERTIME, 0&, 0&)
End Function

Public Function ListView_SetToolTips(hwndLV As Long, hwndNewHwnd As Long) As Long
    ListView_SetToolTips = A_SendMessageAnyRef(hwndLV, LVM_SETTOOLTIPS, ByVal hwndNewHwnd, 0&)
End Function

Public Function ListView_GetToolTips(hwndLV As Long) As Long
    ListView_GetToolTips = A_SendMessageAnyRef(hwndLV, LVM_GETTOOLTIPS, 0&, 0&)
End Function

Public Function ListView_SetBkImage(hWnd As Long, URL As String) As Boolean ' plvbki As LVBKIMAGE) As Boolean
  Dim uLBI As LVBKIMAGE
  Dim lRet As Long
        
        With uLBI
            .pszImage = StrPtr(StrConv(URL & vbNullChar, vbFromUnicode))
            .cchImageMax = Len(URL) + 1
            .ulFlags = LVBKIF_SOURCE_URL Or LVBKIF_STYLE_TILE
        End With
        lRet = A_SendMessageAnyRef(hWnd, LVM_SETBKIMAGE, 0&, uLBI)
        
        If (lRet) Then
            Call A_SendMessage(hWnd, LVM_SETTEXTBKCOLOR, 0&, CLR_NONE)
        End If

    ListView_SetBkImage = lRet 'A_SendMessageAnyRef(hWnd, LVM_SETBKIMAGE, 0&, plvbki)
End Function

Public Function ListView_GetBkImage(hWnd As Long, plvbki As LVBKIMAGE) As Boolean
    ListView_GetBkImage = A_SendMessageAnyRef(hWnd, LVM_GETBKIMAGE, 0&, plvbki)
End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''RichEdit (textbox) macros
'''''''''''''''''''''''''''''''''''''''''''

Public Function RichEdit_CanRedo(hWnd As Long) As Boolean
    If A_SendMessage(hWnd, EM_CANREDO, 0&, 0&) > 0 Then RichEdit_CanRedo = True
End Function

Public Function RichEdit_Redo(hWnd As Long) As Long
    RichEdit_Redo = A_SendMessage(hWnd, EM_REDO, 0&, 0&)
End Function

Public Function RichEdit_CanPaste(hWnd As Long) As Boolean
    If A_SendMessage(hWnd, EM_CANPASTE, 0&, 0&) > 0 Then RichEdit_CanPaste = True
End Function

Public Function RichEdit_GetTextLengthEx(hWnd As Long, _
                                        Optional GTL_flags As Long = GTL_DEFAULT, _
                                        Optional lCodePage As Long = 1200) As Long
    Dim emgtl As GETTEXTLENGTHEX
    
    emgtl.flags = GTL_flags 'GTL_DEFAULT (Num of chars) GTL_USECRLF GTL_NUMCHARS GTL_NUMBYTES
    emgtl.codepage = lCodePage '1200 UNICODE / CP_ACP for ANSI
    RichEdit_GetTextLengthEx = A_SendMessageAnyAnyRef(hWnd, EM_GETTEXTLENGTHEX, emgtl, 0&)
End Function

Public Function RichEdit_GetTextEx(hWnd As Long, _
                                    sText As String, _
                                    Optional lLen As Long = -1, _
                                    Optional GTL_flags As Long = GTL_DEFAULT, _
                                    Optional lCodePage As Long = 1200) As Long
    Dim emgt As GETTEXTEX
    If lLen = -1 Then
        lLen = RichEdit_GetTextLengthEx(hWnd)
    End If
    sText = String$(lLen, vbNullChar)
    'Account for terminating null
    lLen = lLen + 1
    emgt.cb = lLen * 2
    emgt.flags = GTL_flags
    emgt.codepage = lCodePage

   'Indicates number of bytes available in the buffer including terminating null
   RichEdit_GetTextEx = A_SendMessageAnyAnyRef(hWnd, EM_GETTEXTEX, emgt, ByVal StrPtr(sText))

End Function

'If the cpMin and cpMax members of CHARRANGE are equal, the range is empty.
'If cpMin is 0 and cpMax is —1, the range includes everything.
Public Function RichEdit_GetSelRange(hWnd As Long, chrRange As CHARRANGE) As Long
    RichEdit_GetSelRange = A_SendMessageAnyRef(hWnd, EM_EXGETSEL, 0&, chrRange)
End Function



'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Tab Control macros
'''''''''''''''''''''''''''''''''''''''''''

Public Function TabCtl_SetImageList(hWnd As Long, hImgList As Long) As Long
    TabCtl_SetImageList = A_SendMessageAnyRef(hWnd, TCM_SETIMAGELIST, 0&, ByVal hImgList)
End Function

Public Function TabCtl_GetTabText(hWnd As Long, lItemIndex As Long, sText As String) As Boolean
    Dim tie As TCITEM
    Dim bBuf() As Byte
                
    ReDim bBuf(MAX_PATH - NUM_ONE)
    'Set whuch property we want
    tie.mask = TCIF_TEXT
    tie.cchTextMax = MAX_PATH
    'Set pre allocate buffer address
    'In Unicode, we just pass strptr
    tie.pszText = VarPtr(bBuf(0))
    'Get the item with requested properties
    TabCtl_GetTabText = A_SendMessageAnyRef(hWnd, TCM_GETITEM, lItemIndex, tie)
    If TabCtl_GetTabText Then sText = StripNulls(String_Macros.String_ByteArrayToString(bBuf))
End Function

Public Function TabCtl_GetTabImage(hWnd As Long, lItemIndex As Long) As Long
    Dim tie As TCITEM
    'Set whuch property we want
    tie.mask = TCIF_IMAGE
    If CBool(A_SendMessageAnyRef(hWnd, TCM_GETITEM, lItemIndex, tie)) = True Then
        TabCtl_GetTabImage = tie.iImage
    Else
        TabCtl_GetTabImage = -1
    End If
End Function

Public Function TabCtl_GetTablParam(hWnd As Long, lItemIndex As Long) As Long
    Dim tie As TCITEM
    'Set whuch property we want
    tie.mask = TCIF_PARAM
    If CBool(A_SendMessageAnyRef(hWnd, TCM_GETITEM, lItemIndex, tie)) = True Then
        TabCtl_GetTablParam = tie.lParam
    Else
        TabCtl_GetTablParam = -1
    End If
End Function

'Returns the index of the new tab if successful, or -1 otherwise.
Public Function TabCtl_InsertItem(hWnd As Long, sText As String, _
                                    Optional lImage As Long = -1, _
                                    Optional lParam As Long = 0, _
                                    Optional lInsertIndex As Long = -1) As Long
    Dim tie As TCITEM
    Dim lTabCount As Long
    
    tie.mask = TCIF_TEXT Or TCIF_IMAGE Or TCIF_PARAM
    tie.cchTextMax = Len(sText)
    tie.pszText = StrPtr(StrConv(sText, vbFromUnicode))
    tie.iImage = lImage
    tie.lParam = lParam
    If lInsertIndex = -1 Then
        'Get the tab count and insert after last one
        lInsertIndex = A_SendMessage(hWnd, TCM_GETITEMCOUNT, 0&, 0&)
    End If
    TabCtl_InsertItem = A_SendMessageAnyRef(hWnd, TCM_INSERTITEM, lInsertIndex, tie)
End Function

Public Sub TabCtl_ReassignTooltips(hWnd As Long, lStartIndex As Long)
    Dim hTooltipWnd As Long
    Dim lCount As Long
    Dim lTabCount As Long
    Dim sTabText As String
    
'    On Error GoTo TabCtl_ReassignTooltips_Error
    
    'Get handle to tooltip
    hTooltipWnd = A_SendMessage(hWnd, TCM_GETTOOLTIPS, 0&, 0&)
    If hTooltipWnd <= 0 Then Exit Sub
    'Get tab count
    lTabCount = A_SendMessage(hWnd, TCM_GETITEMCOUNT, 0&, 0&)
    If lTabCount = lStartIndex Then Exit Sub 'Was inserted at the end
    
    If lTabCount = 1 Then
        'Get the tab text
        If TabCtl_GetTabText(hWnd, 0&, sTabText) = True Then
            'Update the tooltip for this item
            ToolTip_UpdateTipText hTooltipWnd, hWnd, 0&, sTabText
        End If
    ElseIf lTabCount > 1 Then
        For lCount = lStartIndex To lTabCount - 1
            If TabCtl_GetTabText(hWnd, lCount, sTabText) = True Then
                ToolTip_UpdateTipText hTooltipWnd, hWnd, lCount, sTabText
            End If
        Next
    End If

'    Exit Sub
'TabCtl_ReassignTooltips_Error:
    'A_MessageBox hWnd, "Error " & Err.Number & " (" & Err.Description & ") in procedure TabCtl_ReassignTooltips of Module modSubclass", "ERROR", MB_OK Or MB_ICONERROR
End Sub



'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Edit (textbox) macros
'''''''''''''''''''''''''''''''''''''''''''

'Returns number of bytes written or 0
Public Function TextBox_SaveToFile(hWnd As Long, sFileName As String, Optional bAppend As Boolean = False) As Long
    TextBox_SaveToFile = File_Macros.WriteToFileTextApi(sFileName, Window_GetText(hWnd), bAppend)
End Function

Public Function TextBox_LoadFromFile(hWnd As Long, sFileName As String) As Boolean
    Dim sStr As String
    sStr = File_Macros.TextFromFileApi(sFileName)
    If LenB(sStr) > 0 Then TextBox_LoadFromFile = CBool(A_SetWindowText(hWnd, sStr))
End Function

'This message does not return a value.
Public Sub TextBox_ReplaceSel(hWnd As Long, sText As String, Optional bCanUndo As Boolean = True)
    A_SendMessageStr hWnd, EM_REPLACESEL, CLng(bCanUndo), sText
End Sub

Public Function TextBox_SetText(hWnd As Long, sText As String) As Long
    TextBox_SetText = A_SendMessageStr(hWnd, WM_SETTEXT, 0&, sText)
End Function

Public Function TextBox_Clear(hWnd As Long) As Long
    TextBox_Clear = A_SendMessageStr(hWnd, WM_SETTEXT, 0&, "")
End Function

Public Function TextBox_Paste(hWnd As Long) As Long
    TextBox_Paste = A_SendMessage(hWnd, WM_PASTE, 0&, 0&)
End Function

Public Function TextBox_Cut(hWnd As Long) As Long
    TextBox_Cut = A_SendMessage(hWnd, WM_CUT, 0&, 0&)
End Function

Public Function TextBox_Copy(hWnd As Long) As Long
    TextBox_Copy = A_SendMessage(hWnd, WM_COPY, 0&, 0&)
End Function

Public Function TextBox_Delete(hWnd As Long) As Long
    TextBox_Delete = A_SendMessage(hWnd, WM_CLEAR, 0&, 0&)
End Function

Public Function TextBox_CanUndo(hWnd As Long) As Boolean
    If A_SendMessage(hWnd, EM_CANUNDO, 0&, 0&) > 0 Then TextBox_CanUndo = True
End Function

Public Function TextBox_Undo(hWnd As Long) As Boolean
    TextBox_Undo = CBool(A_SendMessage(hWnd, EM_UNDO, 0&, 0&))
End Function

Public Function TextBox_SetSel(hWnd As Long, Optional lStart As Long = 0, Optional lEnd As Long = -1) As Long
    'If the start is 0 and the end is –1, all the text in the edit control is selected.
    'If the start is –1, any current selection is deselected.
    TextBox_SetSel = A_SendMessage(hWnd, EM_SETSEL, lStart, lEnd)
End Function

Public Function TextBox_GetSelText(hWnd As Long, sSelText As String) As Long
    TextBox_GetSelText = A_SendMessageStr(hWnd, EM_GETSELTEXT, 0&, sSelText)
End Function

'indicates whether the contents of the edit control have been modified. used in both Edit+RichEdit
Public Function TextBox_GetModify(hWnd As Long) As Boolean
    TextBox_GetModify = A_SendMessage(hWnd, EM_GETMODIFY, 0&, 0&)
End Function

'A value of TRUE indicates the text has been modified, and a value of FALSE indicates it has not been modified.
Public Sub TextBox_SetModify(hWnd As Long, Optional bModified As Boolean = False)
    A_SendMessage hWnd, EM_SETMODIFY, CLng(bModified), 0&
End Sub

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''StatusBar macros
'''''''''''''''''''''''''''''''''''''''''''

Public Function StatusBar_SetText(hWnd As Long, sText As String, _
                                    SBT_TxtDrawingStyle As Long, _
                                    Optional lPartIndex As Long = 0) As Long
    If lPartIndex = 0 Then
        StatusBar_SetText = A_SendMessageStr(hWnd, SB_SETTEXT, SBT_TxtDrawingStyle, sText)
    Else
        StatusBar_SetText = A_SendMessageStr(hWnd, SB_SETTEXT, lPartIndex Or SBT_TxtDrawingStyle, sText)
    End If
End Function



'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Tooltip macros
'''''''''''''''''''''''''''''''''''''''''''

Public Sub ToolTip_UpdateTipText(hTooltipWnd As Long, hToolWnd As Long, luId As Long, sTip As String)
    Dim ti As TTTOOLINFO
    
    'Fill in the structure
    ti.cbSize = Len(ti)
    ti.hWnd = hToolWnd
    ti.uID = luId
    'Do we have a tool for this one?
    If CBool(A_SendMessageAnyRef(hTooltipWnd, TTM_GETTOOLINFO, 0&, ti)) = True Then
        ti.lpszText = StrPtr(StrConv(sTip, vbFromUnicode))
        'Has no return value
        A_SendMessageAnyRef hTooltipWnd, TTM_UPDATETIPTEXT, 0&, ti
    End If
End Sub

'Your application creates a multiline ToolTip by responding to a TTN_GETDISPINFO notification
'message. To force the ToolTip control to use multiple lines, send a TTM_SETMAXTIPWIDTH message,
'specifying the width of the display rectangle. Text that exceeds this width will wrap to
'the next line rather than widening the display region. The rectangle height will be
'increased as needed to accommodate the additional lines. The ToolTip control will wrap
'the lines automatically, or you can use a carriage return/line feed combination, \r\n,
'to force line breaks at particular locations.
Public Function ToolTip_AddTool(hTipWnd As Long, hWnd As Long, luId As Long, sTip As String, _
                                Optional bCenterTip As Boolean = False) As Boolean
    Dim ti As TTTOOLINFO
    Dim RECT As RECT
    Dim lStyle As Long
    
    GetClientRect hWnd, RECT
    
    ti.cbSize = Len(ti)
    
    lStyle = TTF_SUBCLASS Or TTF_ABSOLUTE
    If bCenterTip = True Then lStyle = lStyle Or TTF_CENTERTIP
    If hWnd = luId Then lStyle = lStyle Or TTF_IDISHWND
    
    ti.uFlags = lStyle
    ti.hWnd = hWnd
    ti.uID = luId
    'If lpszText is set to LPSTR_TEXTCALLBACK, the control sends the TTN_GETDISPINFO notification
    'message to the owner window to retrieve the text.
    If LenB(sTip) = 0 Then
        ti.lpszText = LPSTR_TEXTCALLBACK
    Else
        ti.lpszText = StrPtr(StrConv(sTip, vbFromUnicode))
    End If
'   ToolTip control will cover the whole window
    ti.RECT.Left = RECT.Left
    ti.RECT.Top = RECT.Top
    ti.RECT.Right = RECT.Right
    ti.RECT.Bottom = RECT.Bottom

'    SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW */
    ToolTip_AddTool = CBool(A_SendMessageAnyRef(hTipWnd, TTM_ADDTOOL, 0&, ti))
End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Rebar macros
'''''''''''''''''''''''''''''''''''''''''''

'iImage
'   Zero-based index of any image that should be displayed in the band.
'   The image list is set using the RB_SETBARINFO message.
'hwndChild
'   Handle to the child window contained in the band, if any.
'cxMinChild
'   Minimum width of the child window, in pixels. The band can't be sized smaller than this value.
'cyMinChild
'   Minimum height of the child window, in pixels. The band can't be sized smaller than this value.
'cX
'   Length of the band, in pixels.
'hbmBack
'   Handle to a bitmap that is used as the background for this band.
'RBBS_USECHEVRON
'   Version 5.80. Show a chevron button if the band is smaller than cxIdeal.
Public Function Rebar_InsertBand(hBar As Long, hChild As Long, lcxMinChild As Long, lcyMinChild As Long, lcX As Long, Optional bNewLine As Boolean = False, Optional lIndex As Long = -1, _
                                Optional RBBIM_fMask As ReBarBandInfoMask = RBBIM_STYLE Or RBBIM_CHILD Or RBBIM_CHILDSIZE Or RBBS_GRIPPERALWAYS Or RBBIM_SIZE, _
                                Optional RBBS_fStyle As ReBarBandStyles = RBBS_CHILDEDGE, _
                                Optional sText As String = "", Optional liImage As Long = -1, Optional hbmBack As Long = 0, _
                                Optional lBackColor As Long = 0, Optional lForeColor As Long = 0, Optional llParam As Long = -1) As Boolean
    Dim rbBand As REBARBANDINFO
'Or RBBS_GRIPPERALWAYS Or RBBS_USECHEVRON
'  Initialize structure members that both bands will share.
    rbBand.cbSize = Len(rbBand)  'Required
    rbBand.fMask = RBBIM_fMask 'Or RBBIM_IDEALSIZE

    If bNewLine = True Then
        rbBand.fStyle = RBBS_fStyle Or RBBS_BREAK
    Else
        rbBand.fStyle = RBBS_fStyle
    End If

''   Set values unique to the band with the toolbar.
    If LenB(sText) > 0 Then
        rbBand.fMask = rbBand.fMask Or RBBIM_TEXT
        rbBand.lpText = StrPtr(StrConv(sText, vbFromUnicode))
        rbBand.cch = Len(sText)
    End If
    rbBand.hwndChild = hChild
    rbBand.cxMinChild = lcxMinChild '0
    rbBand.cyMinChild = lcyMinChild 'HiWord(dwBtnSize)
    rbBand.cx = lcX '250
    'Can be set using RB_MAXIMIZEBAND msg
    'rbBand.cxIdeal = 400
    If lBackColor > 0 Or lBackColor > 0 Then
        rbBand.fMask = rbBand.fMask Or RBBIM_COLORS
        rbBand.clrBack = lBackColor
        rbBand.clrFore = lForeColor
    End If
    If hbmBack > 0 Then
        rbBand.fMask = rbBand.fMask Or RBBIM_BACKGROUND
        rbBand.hbmBack = hbmBack
    End If
    If liImage > -1 Then
        rbBand.fMask = rbBand.fMask Or RBBIM_IMAGE
        rbBand.iImage = liImage
    End If
    If llParam > -1 Then
        rbBand.fMask = rbBand.fMask Or RBBIM_LPARAM
        rbBand.lParam = llParam
    End If

'   Add the band that has the toolbar. 0 failed, nonzero success
   Rebar_InsertBand = CBool(A_SendMessageAnyRef(hBar, RB_INSERTBANDA, lIndex, rbBand))

End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Toolbar macros
'''''''''''''''''''''''''''''''''''''''''''

Public Function Toolbar_AddButton(hWnd As Long, lIdCommand As Long, _
                                    Optional lBitmap As Long = -1, _
                                    Optional sText As String = "", _
                                    Optional lState As ToolBarButtonStates = TBSTATE_ENABLED, _
                                    Optional lStyle As ToolbarButtonStyles = BTNS_BUTTON Or BTNS_AUTOSIZE Or BTNS_SHOWTEXT, _
                                    Optional lData As Long = 0, Optional bDropDown As Boolean = False, _
                                    Optional lDropDownStyle As ToolbarButtonStyles = BTNS_DROPDOWN) As Long
    Dim tBB As TBBUTTON
    
    tBB.iBitmap = lBitmap
    tBB.iString = StrPtr(StrConv(sText, vbFromUnicode))
    tBB.fsState = lState
    If bDropDown = True Then lStyle = lStyle Or lDropDownStyle  'BTNS_WHOLEDROPDOWN
    tBB.fsStyle = lStyle
    tBB.dwData = lData
    'idCommand:
    '   Command identifier associated with the button.
    '   This identifier is used in a WM_COMMAND message when the button is chosen. (LoWord(wParam))
    tBB.idCommand = lIdCommand
    
    Toolbar_AddButton = A_SendMessageAnyRef(hWnd, TB_ADDBUTTONS, 1&, tBB)
End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Window macros
'''''''''''''''''''''''''''''''''''''''''''

Public Function Window_GetText(hWnd As Long) As String
    Dim lLen As Long
    'Display selected Value for UpDown and DateTime Picker
    'Get the current text len
    lLen = A_GetWindowTextLength(hWnd)
    'initialize buffer with nulls, though we can use
    'SysAllocStringByteLen(lStrptr, lNChars) and pass 0& to get a string with len specified
    'but not initilized, faster but wiser, You decide?
    Window_GetText = String$(lLen, vbNullChar)
    A_GetWindowText hWnd, Window_GetText, lLen + 1
End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Combo Macros
'''''''''''''''''''''''''''''''''''''''''''

'Writes data from a combobox to a file
'Returns total bytes written
Public Function ComboBox_SaveToFile(hWnd As Long, sFileName As String, Optional bAppend As Boolean = False) As Long
    On Error GoTo ComboBox_SaveToFile_Error

    Dim lTmpFile As Long
    Dim iCount As Long
    Dim i As Long
    Dim lBytWritten As Long
    Dim lTotal As Long
    Dim sValue As String
    
    If LenB(sFileName) = NUM_ZERO Then Exit Function
    i = ComboBox_GetCount(hWnd)
    
    'No items to write
    If i = NUM_ZERO Then Exit Function
    i = i - 1
    
    'Open a filefor writing
    If File_Macros.FileApi_OpenFile(sFileName, lTmpFile, , bAppend) = True Then
        For iCount = NUM_ZERO To i
            sValue = ComboBox_GetLBText(hWnd, iCount)
            WriteFileStr lTmpFile, sValue, Len(sValue), lBytWritten, ByVal CLng(0)
            lTotal = lTotal + lBytWritten
        Next iCount
        'Close file
        File_Macros.FileApi_CloseFile lTmpFile
        ComboBox_SaveToFile = lTotal
    Else
        A_MessageBox GetDesktopWindow, "Unable to open/create " & sFileName, "Error Saving Data", MB_OK Or MB_ICONERROR
    End If

    Exit Function
ComboBox_SaveToFile_Error:
    File_Macros.FileApi_CloseFile lTmpFile
End Function

'Returns number of items loaded
Public Function ComboBox_LoadFromFile(hWnd As Long, sFileName As String, Optional bIncludeEmptyLines As Boolean = False) As Long
    Dim sTemp As String
    Dim sRet() As String
    Dim i As Long
    Dim iCount As Long

    sTemp = File_Macros.TextFromFileApi(sFileName)
    
    If LenB(sTemp) = NUM_ZERO Then Exit Function

    sRet = Split(sTemp, vbCrLf)
    iCount = UBound(sRet)
    For i = NUM_ZERO To iCount
        If LenB(sRet(i)) = NUM_ZERO Then
            If bIncludeEmptyLines Then A_SendMessageStr hWnd, CB_ADDSTRING, 0&, sRet(i)
        Else
            A_SendMessageStr hWnd, CB_ADDSTRING, 0&, sRet(i)
        End If
    Next
    ComboBox_LoadFromFile = iCount + 1
End Function

Public Function ComboBox_GetSelectedText(hWnd As Long) As String
    Dim lSel As Long, lSelLen As Long
    lSel = A_SendMessage(hWnd, CB_GETCURSEL, 0&, 0&)
    If lSel > -1 Then
        lSelLen = A_SendMessage(hWnd, CB_GETLBTEXTLEN, lSel, 0&)
        If lSelLen > 0 Then
            ComboBox_GetSelectedText = String(lSelLen, vbNullChar)
            A_SendMessageStr hWnd, CB_GETLBTEXT, lSel, ComboBox_GetSelectedText
        End If
    End If
End Function

'Returns the start and end pos of the selected text in edit part of combo
Public Sub ComboBox_GetEditSel(hWnd As Long, lStartPos As Long, lEndPos As Long)
    Dim lPos As Long
    'The return value is a zero-based DWORD value with the starting position of the selection
    'in the low-order word and with the ending position of the first character after
    'the last selected character in the high-order word.
    lPos = A_SendMessage(hWnd, CB_GETEDITSEL, 0&, 0&)
    lStartPos = LoWord(lPos)
    lEndPos = HiWord(lPos)
End Sub

Public Function ComboBox_GetLBText(hWnd As Long, lIndex As Long) As String
    Dim lSLen As Long
    
    'Get a buffer set up
    ComboBox_GetLBText = String$(MAX_PATH, vbNullChar)
    'Returns number of charcters copied int our buffer
    lSLen = A_SendMessageStr(hWnd, CB_GETLBTEXT, lIndex, ComboBox_GetLBText)
    If lSLen <> CB_ERR Then
        ComboBox_GetLBText = Left$(ComboBox_GetLBText, lSLen)
    Else
        ComboBox_GetLBText = CHAR_ZERO_LENGTH_STRING
    End If
End Function

Public Function ComboBox_GetCount(hWnd As Long) As Long
    ComboBox_GetCount = A_SendMessage(hWnd, CB_GETCOUNT, 0&, 0&)
End Function

'If the message is successful, the return value is the index of the item selected.
'If wParam is greater than the number of items in the list or
'if wParam is –1, the return value is CB_ERR and the selection is cleared.
Public Function ComboBox_SetCurSelection(hWnd As Long, lCurSel As Long) As Long
    ComboBox_SetCurSelection = A_SendMessage(hWnd, CB_SETCURSEL, lCurSel, 0&)
End Function

Public Function ComboBox_ClearSelection(hWnd As Long) As Long
    ComboBox_ClearSelection = A_SendMessage(hWnd, CB_SETCURSEL, -1&, 0&)
End Function

Public Function ComboBox_GetCurSelection(hWnd As Long) As Long
    ComboBox_GetCurSelection = A_SendMessage(hWnd, CB_GETCURSEL, 0&, 0&)
End Function

'The return value is the zero-based index to the string in the list box of the combo box.
'If an error occurs, the return value is CB_ERR.
'If insufficient space is available to store the new string, it is CB_ERRSPACE.
' if lInsertAt > -1 then item is inserted at the specific index
Public Function ComboBox_AddString(hWnd As Long, sText As String, Optional lInsertAt As Long = -1) As Long
    If lInsertAt > -1 Then
        ComboBox_AddString = A_SendMessageStr(hWnd, CB_INSERTSTRING, lInsertAt, sText)
    Else
        ComboBox_AddString = A_SendMessageStr(hWnd, CB_ADDSTRING, 0&, sText)
    End If
End Function

Public Function ComboBoxEx_SetImageList(hWnd As Long, hImgList As Long) As Long
    ComboBoxEx_SetImageList = A_SendMessageAnyRef(hWnd, CBEM_SETIMAGELIST, 0&, ByVal hImgList)
End Function

Public Function ComboBoxEx_InsertItem(hWnd As Long, sText As String, _
                                        Optional lInsertAt As Long = -1, _
                                        Optional lImage As Long = -1, _
                                        Optional lSelectedImage As Long = -1, _
                                        Optional lIndent As Long = -1, _
                                        Optional lParam As Long = 0) As Long
    Dim cmi As COMBOBOXEXITEM
    Dim lMask As Long
    
    If LenB(sText) > 0 Then
        lMask = CBEIF_TEXT
        cmi.cchTextMax = Len(sText)
        cmi.pszText = StrPtr(StrConv(sText, vbFromUnicode))
    End If
    If lImage > -1 Then
        lMask = iif(lMask > 0, lMask Or CBEIF_IMAGE, CBEIF_IMAGE)
        cmi.iImage = lImage
    End If
    If lSelectedImage > -1 Then
        lMask = iif(lMask > 0, lMask Or CBEIF_SELECTEDIMAGE, CBEIF_SELECTEDIMAGE)
        cmi.iSelectedImage = lSelectedImage
    End If
    If lIndent > -1 Then
        lMask = iif(lMask > 0, lMask Or CBEIF_INDENT, CBEIF_INDENT)
        cmi.iIndent = lIndent
    End If
    'To insert an item at the end of the list, set the iItem member to -1.
    cmi.iItem = lInsertAt '0
    cmi.lParam = lParam
    'mask and iOverlay are omitted,can be added easily here
    'Set the mask
    cmi.mask = lMask
    'Returns the index at which the new item was inserted if successful, or -1 otherwise. 0
    ComboBoxEx_InsertItem = A_SendMessageAnyRef(hWnd, CBEM_INSERTITEM, 0&, cmi)
End Function

Public Function ComboBox_DeleteString(hWnd As Long, lIndex As Long) As Long
    ComboBox_DeleteString = A_SendMessage(hWnd, CB_DELETESTRING, lIndex, 0&)
End Function

'This message always returns CB_OKAY.
Public Function ComboBox_Clear(hWnd As Long) As Long
    ComboBox_Clear = A_SendMessage(hWnd, CB_RESETCONTENT, 0&, 0&)
End Function

Public Function ComboBox_IsDropDownVisible(hWnd As Long) As Boolean
    ComboBox_IsDropDownVisible = CBool(A_SendMessage(hWnd, CB_GETDROPPEDSTATE, 0&, 0&))
End Function

'wParam = Specifies the zero-based index of the item preceding the first item to be searched.
'         When the search reaches the bottom of the list, it continues from the top of the list
'         back to the item specified by the wParam parameter.
'         If wParam is –1, the entire list is searched from the beginning.
'If the string is found, the return value is the index of the selected item.
'If a matching item is found, it is selected and copied to the edit control.
Public Function ComboBox_SelectString(hWnd As Long, sText As String, Optional lStartIndex As Long = -1) As Long
    ComboBox_SelectString = A_SendMessageStr(hWnd, CB_SELECTSTRING, lStartIndex, sText)
End Function

'User defined
Public Function ComboBox_GetChildHwnds(hWnd As Long, EditHwnd As Long, ListHwnd As Long) As Long
    Dim cbi As COMBOBOXINFO
    
    cbi.cbSize = Len(cbi)
    ComboBox_GetChildHwnds = GetComboBoxInfo(hWnd, cbi)
    If ComboBox_GetChildHwnds <> 0 Then
        EditHwnd = cbi.hwndItem
        ListHwnd = cbi.hwndList
    End If
End Function

Public Function ComboBox_GetEditHwnd(hWnd As Long, EditHwnd As Long) As Long
    Dim cbi As COMBOBOXINFO
    
    cbi.cbSize = Len(cbi)
    ComboBox_GetEditHwnd = GetComboBoxInfo(hWnd, cbi)
    If ComboBox_GetEditHwnd <> 0 Then
        EditHwnd = cbi.hwndItem
    End If
End Function

Public Function ComboBox_GetListHwnd(hWnd As Long, ListHwnd As Long) As Long
    Dim cbi As COMBOBOXINFO
    
    cbi.cbSize = Len(cbi)
    ComboBox_GetListHwnd = GetComboBoxInfo(hWnd, cbi)
    If ComboBox_GetListHwnd <> 0 Then
        ListHwnd = cbi.hwndList
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''
''''''''''''General Macros
'''''''''''''''''''''''''''''''''''''''''''

Public Function HiByte(ByVal Word As Integer) As Byte
    HiByte = (Word And &HFF00&) \ &H100
End Function

Public Function LoByte(ByVal Word As Integer) As Byte
    LoByte = Word And &HFF
End Function

Public Function LoWord(ByVal dw As Long) As Integer
    CopyMemory LoWord, ByVal VarPtr(dw), 2
End Function

Public Function HiWord(ByVal dw As Long) As Integer
    CopyMemory HiWord, ByVal VarPtr(dw) + 2, 2
End Function

'Faster version
Public Function LoWord01(lDWord As Long) As Integer
  If lDWord And &H8000& Then
    LoWord01 = lDWord Or &HFFFF0000
  Else
    LoWord01 = lDWord And &HFFFF&
  End If
End Function

'Faster version
Public Function HiWord01(lDWord As Long) As Integer
  HiWord01 = (lDWord And &HFFFF0000) \ &H10000
End Function

Public Function MAKEWORD(LoByte As Byte, HiByte As Byte) As Integer
    If HiByte And &H80 Then
        MAKEWORD = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
    Else
        MAKEWORD = (HiByte * &H100) Or LoByte
    End If
End Function

' Combines two integers into a long integer
Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
  MAKELONG = wLow
  CopyMemory ByVal VarPtr(MAKELONG) + 2, wHigh, 2
End Function

' Combines two integers into a DWord
Public Function MakeDWord(ByVal HiWord As Long, ByVal LoWord As Long) As Long
'2 ^ 16 = 65536
    MakeDWord = ((LoWord And 65536) Or ((HiWord And 65536) * 65536))
End Function

'Faster
Public Function MAKEDWORDInt(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
  ' High word is coerced to Long to allow it to
  ' overflow limits of multiplication which shifts
  ' it left.
  'H10000 = = 65535
  'HFFFF  = 65596 = 2 ^ 16
    MAKEDWORDInt = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function

' Combines two integers into a long integer
Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
  MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

Public Function BitToLong(bitexpr As String) As Long
    Static t%(31): Dim asc0%
    
    If Len(bitexpr) <> 32 Then Exit Function
    
    CopyMemory t(0), ByVal StrPtr(bitexpr), 64
    asc0 = KeyCodeConstants.vbKey0
 
    BitToLong = t(1) - asc0
    BitToLong = 2 * BitToLong + t(2) - asc0
    BitToLong = 2 * BitToLong + t(3) - asc0
    BitToLong = 2 * BitToLong + t(4) - asc0
    BitToLong = 2 * BitToLong + t(5) - asc0
    BitToLong = 2 * BitToLong + t(6) - asc0
    BitToLong = 2 * BitToLong + t(7) - asc0
    BitToLong = 2 * BitToLong + t(8) - asc0
    BitToLong = 2 * BitToLong + t(9) - asc0
    BitToLong = 2 * BitToLong + t(10) - asc0
    BitToLong = 2 * BitToLong + t(11) - asc0
    BitToLong = 2 * BitToLong + t(12) - asc0
    BitToLong = 2 * BitToLong + t(13) - asc0
    BitToLong = 2 * BitToLong + t(14) - asc0
    BitToLong = 2 * BitToLong + t(15) - asc0
    BitToLong = 2 * BitToLong + t(16) - asc0
    BitToLong = 2 * BitToLong + t(17) - asc0
    BitToLong = 2 * BitToLong + t(18) - asc0
    BitToLong = 2 * BitToLong + t(19) - asc0
    BitToLong = 2 * BitToLong + t(20) - asc0
    BitToLong = 2 * BitToLong + t(21) - asc0
    BitToLong = 2 * BitToLong + t(22) - asc0
    BitToLong = 2 * BitToLong + t(23) - asc0
    BitToLong = 2 * BitToLong + t(24) - asc0
    BitToLong = 2 * BitToLong + t(25) - asc0
    BitToLong = 2 * BitToLong + t(26) - asc0
    BitToLong = 2 * BitToLong + t(27) - asc0
    BitToLong = 2 * BitToLong + t(28) - asc0
    BitToLong = 2 * BitToLong + t(29) - asc0
    BitToLong = 2 * BitToLong + t(30) - asc0
    BitToLong = t(31) - asc0 + 2 * BitToLong
    If t(0) <> asc0 Then BitToLong = BitToLong Or &H80000000
End Function

'LongToBit(&HAAAAAAAA) --> "10101010101010101010101010101010"
'LongToBit(&H0) --> "00000000000000000000000000000000"
'LongToBit(&HFFFFFFFF) --> "11111111111111111111111111111111"
Public Static Function LongToBit(l As Long) As String
' by Peter Nierop, pnierop.pnc@inter.nl.net, 20001226
  Dim lDone&, sNibble(0 To 15) As String, sByte(0 To 255) As String
  If lDone = 0 Then
    sNibble(0) = "0000 "
    sNibble(1) = "0001 "
    sNibble(2) = "0010 "
    sNibble(3) = "0011 "
    sNibble(4) = "0100 "
    sNibble(5) = "0101 "
    sNibble(6) = "0110 "
    sNibble(7) = "0111 "
    sNibble(8) = "1000 "
    sNibble(9) = "1001 "
    sNibble(10) = "1010 "
    sNibble(11) = "1011 "
    sNibble(12) = "1100 "
    sNibble(13) = "1101 "
    sNibble(14) = "1110 "
    sNibble(15) = "1111 "
    For lDone = 0 To 255
      sByte(lDone) = sNibble(lDone \ &H10) & sNibble(lDone And &HF)
    Next
  End If

  If l < 0 Then
    LongToBit = sByte(128 + (l And &H7FFFFFFF) \ &H1000000 And &HFF) _
                & sByte((l And &H7FFFFFFF) \ &H10000 And &HFF) _
                & sByte((l And &H7FFFFFFF) \ &H100 And &HFF) _
                & sByte(l And &HFF)
  Else
    LongToBit = sByte(l \ &H1000000 And &HFF) _
                & sByte(l \ &H10000 And &HFF) _
                & sByte(l \ &H100 And &HFF) _
                & sByte(l And &HFF)
  End If

End Function

Public Sub RefreshApi(hWnd As Long)
    'we force a repaint by first marking ctl's entire region as invalidated (needs repainting)
    InvalidateRgn hWnd, 0&, 1&
    'The UpdateWindow function updates the client area of the specified window by sending a
    'WM_PAINT message to the window if the window's update region is not empty.
    'The function sends a WM_PAINT message directly to the window procedure of the specified window,
    'bypassing the application queue. If the update region is empty, no message is sent.
    UpdateWindow hWnd
End Sub

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Menu Macro
'''''''''''''''''''''''''''''''''''''''''''

Public Function GetSubMenuText(mnuMain As Long, mnuID As Long, Optional bByPosition As Boolean = False) As String
    Dim MII As MENUITEMINFO
    Dim bBuf() As Byte
    Dim lRet As Long
    
    ReDim bBuf(MAX_PATH - 1)
    MII.cbSize = LenB(MII)
    MII.fMask = MIIM_STRING  'set to MIIM_ID to use wId and so on
    MII.dwTypeData = VarPtr(bBuf(0))
    MII.cch = MAX_PATH
    'fByPosition = If this parameter is FALSE, uItem is a menu item identifier. Otherwise, it is a menu item position.
    lRet = A_GetMenuItemInfo(mnuMain, mnuID, CLng(bByPosition), MII)
    If lRet <> 0& Then
        'Contains the actual number of characters copied to our buffer
        If MII.cch > 0& Then
            GetSubMenuText = StripNulls(StrConv(bBuf, vbUnicode))
        'Else
        '    Debug.Print "Menu" & mnuID & ": Bitmap/Separator"
        End If
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Form Macro
'''''''''''''''''''''''''''''''''''''''''''

'Displays a for for the first time
Public Sub ShowWndForFirstTime(hWnd As Long)
    ShowWindow hWnd, SW_NORMAL
    UpdateWindow hWnd
    SetFocusApi hWnd
End Sub

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''ListBox Macros
'''''''''''''''''''''''''''''''''''''''''''

Public Function ListBox_GetCount(hWnd As Long) As Long
    ListBox_GetCount = A_SendMessage(hWnd, LB_GETCOUNT, 0&, 0&)
End Function

'If the message is successful, the return value is the index of the item selected.
'If wParam is greater than the number of items in the list or
'if wParam is –1, the return value is CB_ERR and the selection is cleared.
Public Function ListBox_SetCurSelection(hWnd As Long, lCurSel As Long) As Long
    ListBox_SetCurSelection = A_SendMessage(hWnd, LB_SETCURSEL, lCurSel, 0&)
End Function

Public Function ListBox_ClearSelection(hWnd As Long) As Long
    ListBox_ClearSelection = A_SendMessage(hWnd, LB_SETCURSEL, -1&, 0&)
End Function

Public Function ListBox_GetCurSelection(hWnd As Long) As Long
    ListBox_GetCurSelection = A_SendMessage(hWnd, LB_GETCURSEL, 0&, 0&)
End Function

'The return value is the number of items placed in the buffer.
'If the list box is a single-selection list box, the return value is LB_ERR.
'arrIndexs = Uninitialized array to receive the indexs of selected items
Public Function ListBox_GetSelItems(hWnd As Long, arrIndexs() As Long) As Long
    Dim lSelCount As Long
    lSelCount = A_SendMessage(hWnd, LB_GETSELCOUNT, 0&, 0&)
    If lSelCount = LB_ERR Then Exit Function
    
    ReDim arrIndexs(lSelCount - 1)
    ListBox_GetSelItems = A_SendMessageAnyRef(hWnd, LB_GETSELITEMS, lSelCount, arrIndexs(0))
End Function

'The return value is the count of selected items in the list box.
'If the list box is a single-selection list box, the return value is LB_ERR.
Public Function ListBox_GetSelCount(hWnd As Long) As Long
    ListBox_GetSelCount = A_SendMessage(hWnd, LB_GETSELCOUNT, 0&, 0&)
End Function

Public Sub ListBox_GetMultiSelIndexs(hWnd As Long, lStartIndex As Long, lEndIndex As Long)
    lStartIndex = A_SendMessage(hWnd, LB_GETANCHORINDEX, 0&, 0&)
    lEndIndex = A_SendMessage(hWnd, LB_GETCARETINDEX, 0&, 0&)
End Sub

'The return value is the zero-based index to the string in the list box of the combo box.
'If an error occurs, the return value is CB_ERR.
'If insufficient space is available to store the new string, it is LB_ERRSPACE.
' if lInsertAt > -1 then item is inserted at the specific index
Public Function ListBox_AddString(hWnd As Long, sText As String, Optional lInsertAt As Long = -1) As Long
    If lInsertAt > -1 Then
        ListBox_AddString = A_SendMessageStr(hWnd, LB_INSERTSTRING, lInsertAt, sText)
    Else
        ListBox_AddString = A_SendMessageStr(hWnd, LB_ADDSTRING, 0&, sText)
    End If
End Function

'The return value is a count of the strings remaining in the list.
'The return value is LB_ERR if the wParam parameter specifies an index
'greater than the number of items in the list.
Public Function ListBox_DeleteString(hWnd As Long, lItemIndex As Long) As Long
    ListBox_DeleteString = A_SendMessage(hWnd, LB_DELETESTRING, lItemIndex, 0&)
End Function

Public Function ListBox_GetSelText(hWnd As Long) As String
    Dim lSel As Long, lSelLen As Long
    
    lSel = A_SendMessage(hWnd, LB_GETCURSEL, 0&, 0&)
    If lSel > -1 Then
        lSelLen = A_SendMessage(hWnd, LB_GETTEXTLEN, lSel, 0&)
        If lSelLen > 0 Then
            ListBox_GetSelText = String$(lSelLen, vbNullChar)
            A_SendMessageStr hWnd, LB_GETTEXT, lSel, ListBox_GetSelText
        End If
    End If
End Function

Public Function ListBox_GetText(hWnd As Long, lIndex As Long) As String
    Dim lSelLen As Long

    'Get a buffer set up
    ListBox_GetText = String$(MAX_PATH, vbNullChar)
    'Returns number of charcters copied int our buffer
    lSelLen = A_SendMessageStr(hWnd, LB_GETTEXT, lIndex, ListBox_GetText)
    If lSelLen <> CB_ERR Then
        ListBox_GetText = Left$(ListBox_GetText, lSelLen)
    Else
        ListBox_GetText = CHAR_ZERO_LENGTH_STRING
    End If
End Function

'wParam = Specifies the zero-based index of the item preceding the first item to be searched.
'         When the search reaches the bottom of the list, it continues from the top of the list
'         back to the item specified by the wParam parameter.
'         If wParam is –1, the entire list is searched from the beginning.
'If the string is found, the return value is the index of the selected item.
'If a matching item is found, it is selected and copied to the edit control.
Public Function ListBox_SelectString(hWnd As Long, sText As String, Optional lStartIndex As Long = -1) As Long
    ListBox_SelectString = A_SendMessageStr(hWnd, LB_SELECTSTRING, lStartIndex, sText)
End Function

'Writes data from a listbox to a file
'Returns total bytes written
Public Function ListBox_SaveToFile(hWnd As Long, sFileName As String, Optional bAppend As Boolean = False) As Long
    On Error GoTo ListBox_SaveToFile_Error

    Dim lTmpFile As Long
    Dim iCount As Long
    Dim i As Long
    Dim lBytWritten As Long
    Dim lTotal As Long
    Dim sValue As String
    
    If LenB(sFileName) = NUM_ZERO Then Exit Function
    i = ListBox_GetCount(hWnd)

    'No items to write
    If i = NUM_ZERO Then Exit Function
    i = i - 1

    'Open a filefor writing
    If File_Macros.FileApi_OpenFile(sFileName, lTmpFile, , bAppend) = True Then
        For iCount = NUM_ZERO To i
            sValue = ListBox_GetText(hWnd, iCount)
            WriteFileStr lTmpFile, sValue, Len(sValue), lBytWritten, ByVal CLng(0)
            lTotal = lTotal + lBytWritten
        Next iCount
        'Close file
        File_Macros.FileApi_CloseFile lTmpFile
        ListBox_SaveToFile = lTotal
    Else
        A_MessageBox GetDesktopWindow, "Unable to open/create " & sFileName, "Error Saving Data", MB_OK Or MB_ICONERROR
    End If

    Exit Function
ListBox_SaveToFile_Error:
    File_Macros.FileApi_CloseFile lTmpFile
End Function

Public Function ListBox_LoadFromFile(hWnd As Long, sFileName As String, Optional bIncludeEmptyLines As Boolean = False) As Long
    Dim sTemp As String
    Dim sRet() As String
    Dim i As Long
    Dim iCount As Long

    sTemp = File_Macros.TextFromFileApi(sFileName)
    
    If LenB(sTemp) = NUM_ZERO Then Exit Function

    sRet = Split(sTemp, vbCrLf)
    iCount = UBound(sRet)
    For i = NUM_ZERO To iCount
        If LenB(sRet(i)) = NUM_ZERO Then
            If bIncludeEmptyLines Then A_SendMessageStr hWnd, LB_ADDSTRING, 0&, sRet(i)
        Else
            A_SendMessageStr hWnd, LB_ADDSTRING, 0&, sRet(i)
        End If
    Next
    ListBox_LoadFromFile = iCount + 1
End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''CreateGUID Macro
'''''''''''''''''''''''''''''''''''''''''''

'{AAE5C902-EBE7-4E9E-8102-933D2751E485}
Public Function CreateGUID(Optional RemoveParanthesis As Boolean = True) As String

   Dim g As UUID   'GUID type(struct)
   Dim ret As Long
   Dim sGuid As String
   
  'create unique GUID
   If CoCreateGuid(g) = NUM_ZERO Then
   
     'convert to a string
      sGuid = Space$(MAX_PATH)
      ret = StringFromGUID2(g, sGuid, MAX_PATH)
      
      If ret > NUM_ZERO Then
      
        'convert from unicode
         'sGuid = StrConv(sGuid, vbFromUnicode)
         CreateGUID = Left$(sGuid, ret - NUM_ONE)
         If RemoveParanthesis Then
            CreateGUID = Replace(CreateGUID, "{", CHAR_ZERO_LENGTH_STRING)
            CreateGUID = Replace(CreateGUID, "}", CHAR_ZERO_LENGTH_STRING)
         End If
      End If
      
   End If

End Function
