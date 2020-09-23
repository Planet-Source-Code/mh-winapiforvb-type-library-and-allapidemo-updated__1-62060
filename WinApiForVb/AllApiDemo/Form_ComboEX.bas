Attribute VB_Name = "Form_ComboEX"
Option Explicit


'Form Combo+ComboEx
Public hFrmComboEx As Long
Public hComboEx As Long
Public hLblComboEx As Long
Public hBtnHideComboEx As Long
'Image list used for both combos
Public hComboImgList16 As Long
'Owner drawn
Public hComboOwnerDrawn As Long
'Drive
Public hComboDrive As Long

Public Sub BtnShowComboEx_Click()
    Dim hTmpItem As Long, cbStyle As Long
    Dim sB As String
    Dim lEdit As Long
    Dim cbi As COMBOBOXINFO
    
    If hFrmComboEx = 0 Then
        hFrmComboEx = CreateForm(300, 300, 250, 245, AddressOf ComboExWndProc, hMainForm, WS_CAPTION Or WS_SYSMENU Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_OVERLAPPED, WS_EX_CONTROLPARENT Or WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW, "FormComboEx", "ComboEx")
        If hFrmComboEx > 0 Then
            hLblComboEx = CreateLabel(hFrmComboEx, "Please select an item from the list", 5, 5, 240, 17)
            
            CreateLabel hFrmComboEx, "COMBOEX DROPDOWN:", 5, 25, 200, 17
            hComboEx = CreateComboEx(hFrmComboEx, "MyCombo", 5, 45, 200, 250)
            
            If hComboEx > 0 Then
                'Initialize the image list
                hComboImgList16 = ImageList_Create(16, 16, ILC_COLOR8, 4, 1)
                If hComboImgList16 > 0 Then
                    sB = App.Path & "\images\floppy232.ico"
                    hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
                    'i = Index of the image to replace. If i is -1,
                    'the function appends the image to the end of the list.
                    ImageList_ReplaceIcon hComboImgList16, -1, hTmpItem
                    DestroyIcon hTmpItem
    
                    sB = App.Path & "\images\local32.ico"
                    hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
                    ImageList_ReplaceIcon hComboImgList16, -1, hTmpItem
                    DestroyIcon hTmpItem
    
                    sB = App.Path & "\images\cdrom32.ico"
                    hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
                    ImageList_ReplaceIcon hComboImgList16, -1, hTmpItem
                    DestroyIcon hTmpItem
                    
                    sB = App.Path & "\images\network32.ico"
                    hTmpItem = A_LoadImage(App.hInstance, sB, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
                    ImageList_ReplaceIcon hComboImgList16, -1, hTmpItem
                    DestroyIcon hTmpItem
    
                    If ImageList_GetImageCount(hComboImgList16) > 0 Then
                        'Add image list
                        ComboBoxEx_SetImageList hComboEx, hComboImgList16
                    End If
                End If
                ComboBoxEx_InsertItem hComboEx, "Sample Item One", 0, 0, 0
                ComboBoxEx_InsertItem hComboEx, "Sample Item Two", 1, 1, 1
                ComboBoxEx_InsertItem hComboEx, "Sample Item Three", 2, 2, 2
                ComboBoxEx_InsertItem hComboEx, "Sample Item Four", 3, 3, 3
                
                ComboBox_SetCurSelection hComboEx, 0&
            End If
            'Can add CBS_SORT
            'Because of the CBS_OWNERDRAWFIXED style, the system sends the WM_MEASUREITEM message only once
            cbStyle = CBS_NOINTEGRALHEIGHT Or CBS_DROPDOWNLIST Or CBS_OWNERDRAWFIXED Or CBS_HASSTRINGS Or WS_VSCROLL Or WS_TABSTOP
            CreateLabel hFrmComboEx, "OWNERDRAWN DROPDOWNLIST:", 5, 80, 200, 17
            hComboOwnerDrawn = CreateCombo(hFrmComboEx, "", 5, 100, 200, 250, , , cbStyle)
            
            If hComboOwnerDrawn > 0 Then
'                'Adjust Edit window pos, not for owner drawn
'                cbi.cbSize = Len(cbi)
'                GetComboBoxInfo hComboOwnerDrawn, cbi
'                If cbi.hwndItem > 0 Then SetWindowPos cbi.hwndItem, 0, 20, cbi.rcItem.Top, cbi.rcItem.Right - 20, cbi.rcItem.Bottom - cbi.rcItem.Top, SW_SHOWNOACTIVATE
'
                'Add some items for ownerdrawn
                ComboBox_AddString hComboOwnerDrawn, "A:)"
                ComboBox_AddString hComboOwnerDrawn, "C:)"
                ComboBox_AddString hComboOwnerDrawn, "D:)"
                ComboBox_AddString hComboOwnerDrawn, "E:)"
                ComboBox_SetCurSelection hComboOwnerDrawn, 0&
            End If
            CreateLabel hFrmComboEx, "DRIVE COMBOBOX:", 5, 135, 200, 17
            hComboDrive = CreateCombo(hFrmComboEx, "", 5, 155, 200, 250)
            If hComboDrive > 0 Then
                'Let's load up drive list
                A_SendMessageStr hComboDrive, CB_DIR, DDL_DRIVES, ""
                If ComboBox_GetCount(hComboDrive) > 0 Then
                    If ComboBox_SelectString(hComboDrive, "[-c-]") = CB_ERR Then
                        ComboBox_SetCurSelection hComboDrive, 0&
                    End If
                End If
            End If
            
            hBtnHideComboEx = CreateCmdButton(hFrmComboEx, "Close", 5, 185, 80, 30)
            ShowWndForFirstTime hFrmComboEx
            'SetFocusApi hComboEx
        End If
    Else
        ShowWindow hFrmComboEx, SW_SHOW
    End If
End Sub


Public Function ComboExWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Select Case uMsg
        Case WM_COMMAND
            lRet = HiWord(wParam) 'Notification Msg
            Select Case lRet
                'Sent when the user selects a list item, or selects an item and then closes the list.
                'It indicates that the user's selection is to be processed.
                Case CBN_SELENDOK
                    Dim sSel As String
                    If lParam = hComboOwnerDrawn Then
                        sSel = ComboBox_GetSelectedText(hComboOwnerDrawn)
                        sSel = "Ownerdrawn Combo Selection <" & sSel & ">"
                    ElseIf lParam = hComboEx Then
                        sSel = ComboBox_GetSelectedText(hComboEx)
                        sSel = "ComboEx Selection <" & sSel & ">"
                    ElseIf lParam = hComboDrive Then
                        sSel = ComboBox_GetSelectedText(hComboDrive)
                        sSel = "Drive combo Selection <" & sSel & ">"
                    End If
                    A_SendMessageStr hLblComboEx, WM_SETTEXT, 0&, sSel
                    Exit Function
                Case BN_CLICKED
                    If lParam = hBtnHideComboEx Then
                        ShowWindow hFrmComboEx, SW_HIDE
                    End If
                    Exit Function
            End Select
'Owner drawn combo Msgs
        Case WM_MEASUREITEM
            Dim mist As MEASUREITEMSTRUCT
            CopyMemory mist, ByVal lParam, Len(mist)
            'Make sure we have enough room for icon+2
            If mist.itemHeight < 18 Then mist.itemHeight = 18
            'Put struc back after modifications
            CopyMemory ByVal lParam, mist, Len(mist)
            Exit Function
        Case WM_DRAWITEM
            Dim lpdis As DRAWITEMSTRUCT
            Dim tm As A_TEXTMETRIC
            Dim clrForeground As Long, clrBackground As Long
            Dim X As Long, Y As Long
            Dim achTemp As String
            Dim hdc As Long, hIcon As Long
            
            CopyMemory lpdis, ByVal lParam, Len(lpdis)
            'Check for empty item
            If lpdis.itemID > -1 Then
                'The colors depend on whether the item is selected.
                clrForeground = SetTextColor(lpdis.hdc, GetSysColor(iif(lpdis.itemState And ODS_SELECTED, COLOR_HIGHLIGHTTEXT, COLOR_WINDOWTEXT)))
                clrBackground = SetBkColor(lpdis.hdc, GetSysColor(iif(lpdis.itemState And ODS_SELECTED, COLOR_HIGHLIGHT, COLOR_WINDOW)))
    
                'Calculate the vertical and horizontal position.
                A_GetTextMetrics lpdis.hdc, tm
                Y = ((lpdis.rcItem.Bottom + lpdis.rcItem.Top) - tm.tmHeight) / 2
                X = LoWord(GetDialogBaseUnits) / 4
                
                'Get and display the text for the list item.lpdis.rcItem.Top + 1
                achTemp = String$(MAX_PATH, vbNullChar)
                A_SendMessageStr lpdis.hwndItem, CB_GETLBTEXT, lpdis.itemID, achTemp
                achTemp = StripNulls(achTemp)
                
                'A_ExtTextOut lpdis.hdc, X + 18, Y, ETO_CLIPPED Or ETO_OPAQUE, lpdis.rcItem, achTemp, Len(achTemp), 0&
                A_TextOut lpdis.hdc, X + 18, Y, achTemp, Len(achTemp)
                
                'Restore the previous colors.
                SetTextColor lpdis.hdc, clrForeground
                SetBkColor lpdis.hdc, clrBackground
                
                'Draw icon using ImageList_Draw function
                If lpdis.itemState And ODS_FOCUS Then
                    ImageList_Draw hComboImgList16, lpdis.itemID, lpdis.hdc, X, lpdis.rcItem.Top + 1, ILD_BLEND25
                Else
                    ImageList_Draw hComboImgList16, lpdis.itemID, lpdis.hdc, X, lpdis.rcItem.Top + 1, ILD_NORMAL
                End If
                
                'If the item has the focus, draw focus rectangle.
                'Does not look good and causes repainting problems when mouse moves over the dropdown list
                'If lpdis.itemState And ODS_FOCUS Then DrawFocusRect lpdis.hdc, lpdis.rcItem

            End If
            Exit Function
        Case WM_CLOSE
            ShowWindow hFrmComboEx, SW_HIDE
            'DestroyWindow hFrmComboEx
            Exit Function
    End Select
    ComboExWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
