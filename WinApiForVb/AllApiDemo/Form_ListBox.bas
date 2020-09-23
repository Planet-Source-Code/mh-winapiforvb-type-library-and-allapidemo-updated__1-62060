Attribute VB_Name = "Form_ListBox"
Option Explicit

'Form listbox
Public hFrmListBox As Long
Public hListNormal As Long
Public hListOwnerDraw As Long
Public hBtnListAdd As Long
Public hBtnListRemove As Long
Public hTxtListAdd As Long
Public hBtnListClose As Long
Public hLblListStatus As Long
Public arrFonts() As Long       'List of fonts used for drawing list items
Public hListImgList As Long     'Some images to draw
Private lCounter As Long
Public hBrushListBox As Long   'To draw bk of listbox
Private lfListBox As A_LOGFONT
Private tmpLF As A_LOGFONT

Public Sub BtnShowListBox_Click()
    Dim sB As String
    Dim hTmpItem As Long
    Dim lsFont As Long

    
    If hFrmListBox = 0 Then
        'Create form
        hFrmListBox = CreateForm(300, 300, 400, 300, AddressOf ListBoxWndProc, hMainForm, WS_SYSMENU Or WS_BORDER, WS_EX_TOOLWINDOW, "FormMyListBox", "Listbox")
        If hFrmListBox > 0 Then
            'Get form's font to use in CreateFontIndirect
            lsFont = A_SendMessage(hFrmListBox, WM_GETFONT, 0&, 0&)
            'Store current form's font in a LOGFONT struct
            'Returns the number of bytes copied
            A_GetObject lsFont, Len(lfListBox), lfListBox
            
            'Create list box normal
            CreateLabel hFrmListBox, "Normal ListBox Sorted:", 5, 5, 150, 17
            hListNormal = CreateListBox(hFrmListBox, 5, 30, 150, 190)
            'OD
            'Create a brush for the bk
            hBrushListBox = CreateSolidBrush(x_Yellow)
            CreateLabel hFrmListBox, "Ownerdraw ListBox:", 165, 5, 150, 17
            'Notify, want keyboard input, not sorted, scrollbars, border
            hListOwnerDraw = CreateListBox(hFrmListBox, 165, 30, 225, 180, False, True, , True, True, , , , , , True, , , , True)
            'Add some items
            If hListNormal > 0 Then
                For lCounter = 0 To 20
                    ListBox_AddString hListNormal, "ITEM: 0" & CStr(lCounter)
                Next
            End If
            
            'Set flag and add one item to font array
            If hListOwnerDraw > 0 Then
                'Set initial storage to add faster 260 items, each 260 bytes (130 charcters for each item)
                A_SendMessage hListOwnerDraw, LB_INITSTORAGE, MAX_PATH, MAX_PATH
                lCounter = 0
                ReDim arrFonts(lCounter)
                'Fill our array with fonts
                A_EnumFonts GetDC(0), vbNullString, AddressOf EnumFontProcListBox, 0&
                'set hscrollbar to an arbitary number
                A_SendMessage hListOwnerDraw, LB_SETHORIZONTALEXTENT, 300, 0&
            End If
            'Create image list
            hListImgList = ImageList_Create(24, 24, ILC_COLOR8, 2, 1)
            If hListImgList > 0 Then
                sB = App.Path & "\images\uncheck_square.ico"
                hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, 24, 24, LR_LOADFROMFILE)
                'i = Index of the image to replace. If i is -1,
                'the function appends the image to the end of the list.
                ImageList_ReplaceIcon hListImgList, -1, hTmpItem
                DestroyIcon hTmpItem
        
                sB = App.Path & "\images\check_square.ico"
                hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, 24, 24, LR_LOADFROMFILE)
                'i = Index of the image to replace. If i is -1,
                'the function appends the image to the end of the list.
                ImageList_ReplaceIcon hListImgList, -1, hTmpItem
                DestroyIcon hTmpItem
                'Debug.Print "Count=" & ImageList_GetImageCount(hListImgList)
            End If
            'Create the rest of controls
            hBtnListAdd = CreateCmdButton(hFrmListBox, " + ", 5, 210, 30, 30, , , , , , , , WS_EX_DLGMODALFRAME)
            hBtnListRemove = CreateCmdButton(hFrmListBox, " - ", 40, 210, 30, 30, , , , , , , , WS_EX_DLGMODALFRAME)
            hTxtListAdd = CreateTextbox(hFrmListBox, 75, 210, 80, 25)
            hBtnListClose = CreateCmdButton(hFrmListBox, "Close", 165, 210, 50, 30, , , , , , , , WS_EX_DLGMODALFRAME)
            'Status label
            hLblListStatus = CreateLabel(hFrmListBox, "Ready...", 5, 250, 385, 20, , SS_SUNKEN)

            ShowWndForFirstTime hFrmListBox
        End If
    Else
        ShowWndForFirstTime hFrmListBox
    End If
End Sub

Public Function ListBoxWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Dim sText As String, sDraw As String
    Dim lSel As Long
    Dim tm As A_TEXTMETRIC
    Dim X As Long, Y As Long
    Dim rcText As RECT
    
    Select Case uMsg
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                Select Case lParam
                    Case hBtnListAdd
                        sText = Window_GetText(hTxtListAdd)
                        lSel = ListBox_AddString(hListNormal, sText)
                        'Set the current selection
                        If lSel > -1 Then ListBox_SetCurSelection hListNormal, lSel
                    Case hBtnListRemove
                        lSel = -1
                        lSel = ListBox_GetCurSelection(hListNormal)
                        If lSel > -1 Then
                            sText = ListBox_GetSelText(hListNormal)
                            If A_MessageBox(hFrmListBox, "Proceed to remove: " & sText, "Confirmation", MB_YESNO Or MB_ICONQUESTION) = IDYES Then
                                ListBox_DeleteString hListNormal, lSel
                            End If
                        End If
                    Case hBtnListClose
                        ShowWindow hFrmListBox, SW_HIDE
                        BringWindowToTop hMainForm
                End Select
                Exit Function
            ElseIf lRet = LBN_SELCHANGE Then 'Selection changed
                If lParam = hListNormal Then 'Normal list
                    sText = "ListNormal: " & ListBox_GetSelText(hListNormal)
                    A_SendMessageStr hLblListStatus, WM_SETTEXT, 0&, sText
                ElseIf lParam = hListOwnerDraw Then 'Ownerdraw
                    'Keyup+Down have the same effect
                    sText = "ListOwnerDraw: " & ListBox_GetSelText(hListOwnerDraw)
                    A_SendMessageStr hLblListStatus, WM_SETTEXT, 0&, sText
                    If GetAsyncKeyState(VK_DOWN) <> 0 Or GetAsyncKeyState(VK_UP) <> 0 Then
                        'do nothing
                    Else
                        'Check/uncheck
                        lSel = ListBox_GetCurSelection(hListOwnerDraw)
                        If lSel > -1 Then
                            If A_SendMessage(hListOwnerDraw, LB_GETITEMDATA, lSel, 0&) = 0 Then
                                A_SendMessage hListOwnerDraw, LB_SETITEMDATA, lSel, 1&
                            Else
                                A_SendMessage hListOwnerDraw, LB_SETITEMDATA, lSel, 0&
                            End If
                        End If
                        'RefreshApi hListOwnerDraw
                    End If
                    RefreshApi hListOwnerDraw
                End If
                Exit Function
            End If
'Does not apply to ownerdraw or a listbox without LBS_HASSTRINGS flag set
'        Case WM_VKEYTOITEM
'            If LoWord(wParam) = VK_DOWN Or LoWord(wParam) = VK_UP Then
'                '0 or greater specifies the index of an item in the list box and indicates
'                'that the list box should perform the default action for the keystroke on the specified item.
'                '–2 indicates that the application handled all aspects of selecting the item
'                '–1 indicates that the list box should perform the default action
'                ListBoxWndProc = -1
'                Exit Function
'            End If
'OwnerDraw
        Case WM_MEASUREITEM
            Dim mis As MEASUREITEMSTRUCT
            CopyMemory mis, ByVal lParam, Len(mis)
            mis.itemHeight = 24 + 2 'for top and bottom gaps
            CopyMemory ByVal lParam, mis, Len(mis)
            Exit Function
        Case WM_DRAWITEM
            Dim lpdis As DRAWITEMSTRUCT
            Dim lOldBC As Long, lOldFC As Long
            Dim hOldFont As Long
            
            CopyMemory lpdis, ByVal lParam, Len(lpdis)
            'Check to see if an item exists
            If lpdis.itemID = -1 Then Exit Function

            Select Case lpdis.itemAction
                Case ODA_DRAWENTIRE, ODA_SELECT
                    'Get text
                    sDraw = ListBox_GetText(hListOwnerDraw, lpdis.itemID)
                    
                    'Calculate the vertical and horizontal position for the text
                    A_GetTextMetrics lpdis.hdc, tm
                    Y = ((lpdis.rcItem.Bottom + lpdis.rcItem.Top) - tm.tmHeight) / 2
                    X = LoWord(GetDialogBaseUnits) / 4
                    
                    If lpdis.itemState And ODS_SELECTED Then
                        'FillRect with blue
                        CopyRect rcText, lpdis.rcItem
                        'After the image, fill the rest
                        rcText.Left = rcText.Left + 26
                        BoxSolidDC lpdis.hdc, rcText, x_Blue
                        'Save b + f colors and set the new text and bk color for selected text
                        'Look in Form_ComboExWndProc to see how to use system highlight colors instead
                        'to give it a standard look
                        lOldFC = SetTextColor(lpdis.hdc, x_Yellow)
                        lOldBC = SetBkColor(lpdis.hdc, x_Blue)
                    End If
                    
                    'If we have a font for this item then use it
                    If arrFonts(lpdis.itemID) <> 0 Then hOldFont = SelectObject(lpdis.hdc, arrFonts(lpdis.itemID))
                    'Draw text
                    A_TextOut lpdis.hdc, X + 34, Y, sDraw, Len(sDraw)
                    
                    If lpdis.itemState And ODS_SELECTED Then
                        'reset b + f colors
                        SetTextColor lpdis.hdc, lOldFC
                        SetBkColor lpdis.hdc, lOldBC
                    End If
                    
                    'Draw check/uncheck image
                    lSel = lpdis.ItemData
                    If lSel > -1 Then ImageList_Draw hListImgList, lSel, lpdis.hdc, X, lpdis.rcItem.Top + 1, ILD_NORMAL
                    
                    Exit Function
                Case ODA_FOCUS
                    Exit Function
            End Select
'SET BK color
'        Case WM_CTLCOLORLISTBOX
'            If lParam = hListOwnerDraw Then
'                SelectObject wParam, hBrushListBox
'                'Return brush
'                ListBoxWndProc = hBrushListBox
'                Exit Function
'            End If
'Close
        Case WM_CLOSE
            'just hide, do not destroy
            ShowWindow hFrmListBox, SW_HIDE
            BringWindowToTop hMainForm
            'DestroyWindow hFrmListBox
            Exit Function
    End Select
    ListBoxWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function

'Enum all fonts for a given font family
Public Function EnumFontProcListBox(ByVal lplf As Long, ByVal lptm As Long, ByVal dwType As Long, ByVal lpData As Long) As Long
    Dim LF As A_LOGFONT
    Dim FontName As String
    Dim ZeroPos As Long
    
    'Copy to local structure
    CopyMemory LF, ByVal lplf, LenB(LF)
    'LF.lfFaceName is a byte array, simply convert to Unicode
    FontName = StrConv(LF.lfFaceName, vbUnicode)
    ZeroPos = InStr(1, FontName, vbNullChar)
    If ZeroPos > 0 Then FontName = Left$(FontName, ZeroPos - 1)
    
    'First element was added before calling this function
    If lCounter > 0 Then ReDim Preserve arrFonts(lCounter)
    'Add font name to list
    ListBox_AddString hListOwnerDraw, FontName
    'Create a new font based on the form's default font
    tmpLF.lfCharSet = lfListBox.lfCharSet
    tmpLF.lfHeight = lfListBox.lfHeight
    tmpLF.lfWidth = lfListBox.lfWidth
    'Need to add vbNullChar for the ANSI version of lstrcpyn to work
    FontName = FontName & vbNullChar
    'Copy font name into lfFaceName byte array
    A_lstrcpynStrByte tmpLF.lfFaceName(0), FontName, Len(FontName)
    'Store font for each entry
    arrFonts(lCounter) = A_CreateFontIndirect(tmpLF)
    lCounter = lCounter + 1
    
    'Continue
    EnumFontProcListBox = 1
End Function

'Fills a box with a given color
Public Function BoxSolidDC(hdc As Long, rcDraw As RECT, Fill As Long)
    Dim hTmpBrush As Long
    
    hTmpBrush = CreateSolidBrush(Fill)
    FillRect hdc, rcDraw, hTmpBrush
    DeleteObject hTmpBrush
End Function
