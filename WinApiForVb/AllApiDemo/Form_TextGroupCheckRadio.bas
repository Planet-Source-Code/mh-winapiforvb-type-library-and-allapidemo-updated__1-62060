Attribute VB_Name = "Form_TextGroupCheckRadio"
Option Explicit

'Form Textbox, Group, Checkbox, and Radio
Public hFrmTGCR As Long
Public strTxtBuffer As String
Public htxtTextOnly As Long
Public hTxtNumbersOnly As Long
Public hTxtMultiLine As Long
Public orgtxtTextOnly As Long
Public orgTxtNumbersOnly As Long

Public hTxtUpperCase As Long
Public hTxtLowerCase As Long

Public hGrpRadio As Long

Public hBrush As Long

Public hRadioColorDefault As Long
Public hRadioColorGreen As Long
Public hRadioColorCyan As Long
Public hRadioColorYellow As Long

Public lRadButValue As Long
Public hChkEnable As Long

'Textboxs, Group, Check, Radio
Public Sub BtnShowTGCR_Click()
    If hFrmTGCR = 0 Then
        hFrmTGCR = CreateForm(300, 300, 350, 420, AddressOf TGCRWndProc, hMainForm, WS_SYSMENU Or WS_BORDER Or WS_TABSTOP, WS_EX_TOOLWINDOW, "FormTGCR", "Text Group Check Radio")
        If hFrmTGCR > 0 Then
            CreateLabel hFrmTGCR, "Text Only:", 5, 5, 100, 17
            htxtTextOnly = CreateTextbox(hFrmTGCR, 110, 5, 70, 25)
            strTxtBuffer = "Text"
            'Set text limit to max allowable
            'A_SendMessage htxtTextOnly, EM_SETLIMITTEXT, 200, 0&
            
            CreateLabel hFrmTGCR, "UpperCase Only:", 5, 35, 100, 17
            hTxtUpperCase = CreateTextbox(hFrmTGCR, 110, 35, 70, 25, , , , , , , , , , , True)
            'A_SendMessage hTxtUpperCase, EM_SETLIMITTEXT, 500, 0&
            
            CreateLabel hFrmTGCR, "LowerCase Only:", 5, 65, 100, 17
            hTxtLowerCase = CreateTextbox(hFrmTGCR, 110, 65, 70, 25, , , , , , , , , , , , True)
            'A_SendMessage hTxtLowerCase, EM_SETLIMITTEXT, 500, 0&
    
            CreateLabel hFrmTGCR, "Numbers Only:", 5, 95, 100, 17
            hTxtNumbersOnly = CreateTextbox(hFrmTGCR, 110, 95, 70, 25, "1234", , , , , , , , , , , , True)
            'A_SendMessage hTxtNumbersOnly, EM_SETLIMITTEXT, 0&, 0&
            
            'Create a group box
            hGrpRadio = CreateGroupBox(hFrmTGCR, "Group Colors", 190, 5, 150, 150)
            
            hRadioColorDefault = CreateRadioButton(hFrmTGCR, "Default Color", 195, 20, 100, 25, , True)
            hRadioColorCyan = CreateRadioButton(hFrmTGCR, "Cyan Color", 195, 50, 100, 25)
            hRadioColorGreen = CreateRadioButton(hFrmTGCR, "Yellow Color", 195, 80, 100, 25)
            hRadioColorYellow = CreateRadioButton(hFrmTGCR, "Green Color", 195, 110, 100, 25)
            A_SendMessage hRadioColorDefault, BM_SETCHECK, BST_CHECKED, 0& 'Check the first one
            
            hChkEnable = CreateCheckbox(hFrmTGCR, "Disable Text Only", 5, 130, 140, 25)
            A_SendMessage hChkEnable, BM_SETCHECK, BST_CHECKED, 0& 'Textbox is enabled, check the checkbox
            CreateRadioButton hFrmTGCR, "Outside the group", 5, 160, 150, 25, , True
            
            hTxtMultiLine = CreateTextbox(hFrmTGCR, 5, 190, 330, 200, "Multiline textbox" & vbCrLf & "Second line", , , , , , True, True)
            
            'Subclass two textbox for context menu and keyboard restrictions
            If htxtTextOnly > 0 Then orgtxtTextOnly = A_SetWindowLong(htxtTextOnly, GWL_WNDPROC, AddressOf TextboxWndProc)
            If hTxtNumbersOnly > 0 Then orgTxtNumbersOnly = A_SetWindowLong(hTxtNumbersOnly, GWL_WNDPROC, AddressOf TextboxWndProc)
            
            ShowWndForFirstTime hFrmTGCR
        End If
    Else
        ShowWindow hFrmTGCR, SW_SHOW
    End If
End Sub

'Text Group Checkbox Radio
Public Function TGCRWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
   
    Select Case uMsg
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                Select Case lParam
                    Case hRadioColorDefault
                        lRadButValue = 1
                        'Tell Windows to repaint the textbox
                        InvalidateRgn htxtTextOnly, 0&, True
                    Case hRadioColorCyan
                        lRadButValue = 2
                        InvalidateRgn htxtTextOnly, 0&, True
                    Case hRadioColorGreen
                        lRadButValue = 4
                        InvalidateRgn htxtTextOnly, 0&, True
                    Case hRadioColorYellow
                        lRadButValue = 3
                        InvalidateRgn htxtTextOnly, 0&, True
                    Case hChkEnable
                        If A_SendMessage(hChkEnable, BM_GETCHECK, 0&, 0&) = BST_CHECKED Then
                            EnableWindow htxtTextOnly, True
                            A_SetWindowText hChkEnable, "Disable Text Only"
                        Else
                            EnableWindow htxtTextOnly, False
                            A_SetWindowText hChkEnable, "Enable Text Only"
                        End If
                End Select
                Exit Function
            'The EN_CHANGE notification message is sent when the user has taken an action that
            'may have altered text in an edit control. Unlike the EN_UPDATE notification message,
            'this notification message is sent after the system updates the screen.
            ElseIf lRet = EN_CHANGE Then
                If lParam = hTxtUpperCase Then
                    A_SetWindowText hTxtLowerCase, Window_GetText(hTxtUpperCase)
                    Exit Function
'                ElseIf lParam = htxtTextOnly Then
'                    Dim strTmp As String, strRet As String
'                    'Get the new text
'                    strTmp = Window_GetText(htxtTextOnly)
'                    strRet = strTmp
'                    If LenB(strTmp) > LenB(strTxtBuffer) Then
'                        'Get the added part
'                        strTmp = Right(strTmp, Len(strTmp) - Len(strTxtBuffer))
'                        'Debug.Print "strTmp:" & strTmp & "="
'                        'look for numeric value, if found replace it
'                        'Need to check every character in aloop
'                        If IsNumeric(strTmp) Then
'                            'A_SendMessage htxtTextOnly, EM_sette, 0&, 0&
'                            A_SetWindowText htxtTextOnly, strTxtBuffer
'                            'If the start is 0 and the end is –1, all the text in the edit control is selected. If the start is –1, any current selection is deselected.
'                            A_SendMessage htxtTextOnly, EM_SETSEL, Len(strTxtBuffer), Len(strTxtBuffer)
'                        Else
'                            strTxtBuffer = strRet
'                        End If
'                        Exit Function
'                    End If
                End If
            End If
        Case WM_CTLCOLOREDIT 'wParam HDC
            If lParam = htxtTextOnly Then
                If hBrush <> 0 Then DeleteObject hBrush
                If lRadButValue = 1 Then
                    hBrush = CreateSolidBrush(RGB(255, 255, 255))
                    SetBkColor wParam, x_White 'deault white
                    TGCRWndProc = hBrush
                    Exit Function
                ElseIf lRadButValue = 2 Then 'Cyan
                    hBrush = CreateSolidBrush(x_Cyan)
                    SetBkColor wParam, x_Cyan
                    TGCRWndProc = hBrush
                    Exit Function
                ElseIf lRadButValue = 3 Then 'Yellow
                    hBrush = CreateSolidBrush(RGB(255, 255, 0))
                    SetBkColor wParam, RGB(255, 255, 0)
                    TGCRWndProc = hBrush
                    Exit Function
                ElseIf lRadButValue = 4 Then 'Green
                    hBrush = CreateSolidBrush(RGB(0, 255, 0))
                    SetBkColor wParam, RGB(0, 255, 0)
                    TGCRWndProc = hBrush
                    Exit Function
                End If
            End If
        Case WM_CLOSE
            'just hide
            ShowWindow hFrmTGCR, SW_HIDE
            Exit Function
    End Select
    TGCRWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function

'Text only and numbers only wnd procedures
'Disabling context menu and Ctrl+c/v/x combos
Public Function TextboxWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_CONTEXTMENU
            'No Context for both edit controls
            Exit Function
        Case WM_CHAR
            'Only text in this one
            If hWnd = htxtTextOnly Then
                If String_IsCharNumber(CInt(wParam)) = True Then Exit Function
            End If
            'or use GetKeyState(VK_CONTROL) <> 1
            If GetAsyncKeyState(VK_CONTROL) <> 0 Then
                If GetAsyncKeyState(VK_X) <> 0 Or GetAsyncKeyState(VK_C) <> 0 Or GetAsyncKeyState(VK_V) <> 0 Then Exit Function
            End If
    End Select
    If hWnd = htxtTextOnly Then
        TextboxWndProc = A_CallWindowProc(orgtxtTextOnly, hWnd, uMsg, wParam, lParam)
    Else
        TextboxWndProc = A_CallWindowProc(orgTxtNumbersOnly, hWnd, uMsg, wParam, lParam)
    End If
End Function
