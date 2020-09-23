Attribute VB_Name = "Form_Main"
Option Explicit

'Form main
Public hMainForm As Long
Public hLabel As Long           'Simple frame
Public hMainToolTip As Long     'Baloon+multiline for all btns
Public hHotkey As Long          'Hotkey Control
Public hBtnshowIp As Long
Public hBtnShowAnim As Long
Public hBtnShowEllipseLabels As Long
Public hBtnShowToolbar As Long
Public hBtnShowUDDTP As Long
Public hBtnShowComboEx As Long
Public hBtnShowTreeView As Long
Public hBtnShowTGCR As Long
Public hBtnProgressStatusBar As Long
Public hBtnShowHotkey As Long
Public hBtnShowTabControl As Long
Public hBtnShowListView As Long
Public hBtnShowImages As Long   'Loads/Cycles images from a imagelist control
Public hIeImgList As Long       'Holds # images loaded from one bitmap (~47) with
Public hImageBox As Long        'Displays images loaded in hIeImgList (Static control)
Public lCurImg As Long          'Counter for current image index
Public hBtnShowPageScroller As Long
Public hBtnShowListBox As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''Main Win Msg Handler
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''Common Control Notification Messages (From MSDN)
'Common controls are child windows that send notification messages to the parent window when events,
'such as input from the user, occur in the control. The application relies on these notification messages
'to determine what action the user wants it to take. Except for trackbars, which use the WM_HSCROLL and
'WM_VSCROLL messages to notify their parent of changes, common controls send notification messages as
'WM_NOTIFY messages. The lParam parameter of WM_NOTIFY is either the address of an NMHDR structure or
'the address of a larger structure that includes NMHDR as its first member. The structure contains
'the notification code and identifies the common control that sent the notification message.
'The meaning of the remaining structure members, if any, varies depending on the notification code.
'Note:
'Not all controls will send WM_NOTIFY messages. In particular, the standard Windows controls
'(edit controls, combo boxes, list boxes, buttons, scroll bars, and static controls) do not
'send WM_NOTIFY messages. Consult the documentation for the control to determine if it will
'send any WM_NOTIFY messages and, if it does, which notification codes it will send.
'Each type of common control has a corresponding set of notification codes. The common control
'library also provides notification codes that can be sent by more than one type of common control.
'See the documentation for the control of interest to determine which notification codes it will send
'and what format they take.

Public Function MainWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    
    Select Case uMsg
        'Menus and btns, check, radio msgs
'        Case WM_CREATE
'            Debug.Print "CREATE"
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                Select Case lParam
                    Case hBtnshowIp
                        BtnshowIp_Click
                    Case hBtnShowAnim
                        BtnShowAnim_Click
                    Case hBtnShowEllipseLabels
                        BtnShowEllipseLabels_Click
                    Case hBtnShowUDDTP
                        BtnShowUDDTP_Click
                    Case hBtnShowComboEx
                        BtnShowComboEx_Click
                    Case hBtnShowTreeView
                        BtnShowTreeView_Click
                    Case hBtnShowTGCR
                        BtnShowTGCR_Click
                    Case hBtnShowToolbar
                        BtnShowToolbar_Click
                    Case hBtnProgressStatusBar
                        BtnProgressStatusBar_Click
                    Case hBtnShowHotkey
                        BtnShowHotkey_Click
                    Case hBtnShowTabControl
                        BtnShowTabControl_Click
                    Case hBtnShowListView
                        BtnShowListView_Click
                    Case hBtnShowImages
                        BtnShowImages_Click
                    Case hBtnShowPageScroller
                        BtnShowPageScroller_Click
                    Case hBtnShowListBox
                        BtnShowListBox_Click
                End Select 'End Case lParam
                'If an application processes this message, it should return zero.
                Exit Function
            End If 'End lRet = EN_BTNCLICK
'        Case WM_DRAWITEM
'            Dim sdis As DRAWITEMSTRUCT
'            CopyMemory sdis, ByVal lParam, Len(sdis)
'            If sdis.CtlType = ODT_STATIC And sdis.hwndItem = hImageBox Then
'                'Use sdis.hDC and sdis.rcItem to paint whatever you want
'            End If
        'This message is sent first to the window being destroyed and then to the child windows (if any)
        'as they are destroyed. During the processing of the message,
        'it can be assumed that all child windows still exist.
        Case WM_DESTROY
            'It causes the GetMessage function to return zero and terminates the message loop. END
            PostQuitMessage 0&
            Exit Function
        Case WM_CLOSE
            'If the specified window is a parent or owner window,
            'DestroyWindow automatically destroys the associated child
            'or owned windows when it destroys the parent or owner window.
            'The function first destroys child or owned windows, and then
            'it destroys the parent or owner window.
            'Sends WM_DESTROY and WM_NCDESTROY
            If A_MessageBox(hMainForm, "Proceed to exit demo app", "API", MB_YESNOCANCEL Or MB_ICONINFORMATION) = IDYES Then
                'Restore original msg handlings, can be done in WM_DESTROY msg of related forms (ListView, TGCR,...)
                If hLvEdit > 0 And origLvEditWndProc > 0 Then A_SetWindowLong hLvEdit, GWL_WNDPROC, origLvEditWndProc
                If htxtTextOnly > 0 And orgtxtTextOnly > 0 Then A_SetWindowLong htxtTextOnly, GWL_WNDPROC, orgtxtTextOnly
                If hTxtNumbersOnly > 0 And orgTxtNumbersOnly > 0 Then A_SetWindowLong hTxtNumbersOnly, GWL_WNDPROC, orgTxtNumbersOnly
                'Clean up
                Comctl32_Controls.Comctl32_Terminate
                DestroyWindow hMainForm
            Else
                Exit Function
            End If
    End Select
    MainWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function


'================================Sort of like event handlers=======================

Private Sub BtnShowImages_Click()
    On Error GoTo BtnShowImages_Click_Error
    
    Dim lBytes As Long
    Dim hBitmap As Long
    Dim lImg As Long
    Dim hImgFile As Long
    Dim lTmpImg As Long
    Dim rcImages As RECT
    
    If hIeImgList = 0 Then
        'Create Label to show
        'Create imagelist 24x24, adding mask flag
        'hIeImgList = ImageList_Create(24, 24, ILC_COLOR8 Or ILC_MASK, 50, 1)
        hIeImgList = ImageList_Create(16, 16, ILC_COLOR32 Or ILC_MASK, 50, 1)
        If hIeImgList > 0 Then
            'Load bitmap, 47 images
            'hBitmap = A_LoadImage(App.hInstance, App.Path & "\images\xpbrowser24-over.bmp", IMAGE_BITMAP, 1128&, 24, LR_LOADFROMFILE)
            'Load 32-bit (alpha channel) images from a bitmap to an api created imagelist
            hBitmap = A_LoadImage(App.hInstance, App.Path & "\images\toolbar.bmp", IMAGE_BITMAP, 0&, 0&, LR_CREATEDIBSECTION Or LR_DEFAULTSIZE Or LR_LOADFROMFILE)
            If hBitmap Then
                'Adds an image or images to an image list, generating a mask from the specified bitmap.
                'Bitmaps with color depth greater than 8bpp are not supported
                lImg = ImageList_AddMasked(hIeImgList, hBitmap, x_Black)
                DeleteObject hBitmap
                If lImg > -1 Then
                    lTmpImg = ImageList_GetIcon(hIeImgList, 0&, ILD_NORMAL)
                    'Create Static(Label)
                    hImageBox = CreateImageBox(hMainForm, 165, 215, , True, , lTmpImg)
                    A_SetWindowText hBtnShowImages, "Image Num: " & CStr(lCurImg)
                    'If you want to do the drawing yourself, uncomment below and comment out
                    'above lines. Then comment out WM_DRAWITEM part in MainWndProc and Draw
                    'hImageBox = CreateImageBox(hMainForm, 10, 330, True)
                End If
            End If
        End If
    Else
        If hIeImgList > 0 And ImageList_GetImageCount(hIeImgList) > 0 And hImageBox > 0 Then
            lCurImg = lCurImg + 1
            If lCurImg >= ImageList_GetImageCount(hIeImgList) Then lCurImg = 0
            
            'if static is ownerdrawn then comment out the rest of this code
            'and just refresh (forces a repaint) the static and do the drawing
            'in WM_DRAWITEM part of MainWndProc
            
            'Get next icon
            lTmpImg = ImageList_GetIcon(hIeImgList, lCurImg, ILD_NORMAL)
            'Replace icon with the new one
            lImg = A_SendMessage(hImageBox, STM_SETICON, lTmpImg, 0&)
            'Debug.Print "LastIcon=" & lImg & " lTmpImg: " & lTmpImg
            RefreshApi hImageBox
            A_SetWindowText hBtnShowImages, "Image Num: " & CStr(lCurImg)
        End If
    End If

    Exit Sub
BtnShowImages_Click_Error:
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure BtnShowImages_Click of Module Form_Main"
End Sub

Private Sub BtnShowHotkey_Click()
    If hHotkey > 0 Then
        'Trying it again
        If IsWindowVisible(hHotkey) Then
            Dim lHotKey As Long
            Dim lHotKeySetRet As Long
            'set focus back to Hotkey window
            SetFocusApi hHotkey
            'Get hot key combo
            lHotKey = A_SendMessage(hHotkey, HKM_GETHOTKEY, 0&, 0&)
            If lHotKey > 0 Then 'If valid
                'Set global hotkey for this window
                lHotKeySetRet = A_SendMessage(hMainForm, WM_SETHOTKEY, lHotKey, 0&)
                'Debug.Print "Hotkey" & lHotKey & " Setting=" & lHotKeySetRet
                Select Case lHotKeySetRet
                    Case -1
                        A_MessageBox hMainForm, "The hot key is invalid.", "HOTKEY SETTING RESULT", MB_OK Or MB_ICONERROR
                        SetFocusApi hHotkey
                    Case 0
                        A_MessageBox hMainForm, "The window is invalid.", "HOTKEY SETTING RESULT", MB_OK Or MB_ICONERROR
                    Case 1
                        A_MessageBox hMainForm, "Successfully set the hot key combo.", "HOTKEY SETTING RESULT", MB_OK Or MB_ICONINFORMATION
                        ShowWindow hHotkey, SW_HIDE
                    Case 2
                        A_MessageBox hMainForm, "Another window already has the same hot key?", "HOTKEY SETTING RESULT", MB_OK Or MB_ICONWARNING
                        SetFocusApi hHotkey
                End Select
            Else
                A_MessageBox hMainForm, "Please select a valid hotkey combo." & vbCrLf & "Ctlr+Alt+T", "HOTKEY SETTING RESULT", MB_OK Or MB_ICONERROR
                SetFocusApi hHotkey
            End If 'End of If lHotKey > 0
        Else
            ShowWindow hHotkey, SW_SHOW
            SetFocusApi hHotkey
        End If 'End of isWindowVisible
    Else
        'add a hotkey window, so we can set up a global hotkey using hotkey field
        hHotkey = CreateHotKey(hMainForm, 10, 10, 310, 20)
    End If 'End if hHotkey > 0
End Sub




