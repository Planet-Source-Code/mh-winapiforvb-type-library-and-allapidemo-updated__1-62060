Attribute VB_Name = "Form_IpField"
Option Explicit

'Form IP
Public hFrmIP As Long
Public hIPAddr As Long
Public hLabelIP As Long
Public hBtnIPOk As Long
Public hBtnIPCancel As Long

Public Sub BtnshowIp_Click()
    If hFrmIP = 0 Then
        hFrmIP = CreateForm(300, 300, 250, 100, AddressOf IpFieldWndProc, hMainForm, WS_SYSMENU Or WS_BORDER, WS_EX_TOOLWINDOW, "FormMyIp", "IP Toolwindow style")
        If hFrmIP > 0 Then
            hLabelIP = CreateLabel(hFrmIP, "Enter IP Address:", 5, 15, 100, 17, SS_RIGHT)
            hIPAddr = CreateIPField(hFrmIP, 110, 10, 120, 25, 127, 0, 0, 1)
            hBtnIPOk = CreateCmdButton(hFrmIP, "Accept", 10, 45, 70, 25, , , , False, True)
            hBtnIPCancel = CreateCmdButton(hFrmIP, "Cancel", 85, 45, 70, 25, , , , False, True)
            'Show the form for the first time
            ShowWndForFirstTime hFrmIP
        End If
    Else
        ShowWindow hFrmIP, SW_SHOW
    End If
End Sub

Public Function IpFieldWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Select Case uMsg
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                Select Case lParam
                    Case hBtnIPOk
                        Dim lIp As Long
                        Dim sRet As String
                        'lIp to a DWORD value that receives the address.
                        'The field 3 value will be contained in bits 0 through 7.
                        'The field 2 value will be contained in bits 8 through 15.
                        'The field 1 value will be contained in bits 16 through 23.
                        'The field 0 value will be contained in bits 24 through 31.
                        A_SendMessageAnyRef hIPAddr, IPM_GETADDRESS, 0, lIp
                        If lIp > 0 Then
                            sRet = IPADDRESS_FIRST_IPADDRESS(lIp) & "-" & IPADDRESS_SECOND_IPADDRESS(lIp) & "-" & IPADDRESS_THIRD_IPADDRESS(lIp) & "-" & IPADDRESS_FOURTH_IPADDRESS(lIp)
                            A_MessageBox hFrmIP, "IP: " & sRet, "IP Field", MB_OK Or MB_ICONINFORMATION
                        End If
                        ShowWindow hFrmIP, SW_HIDE
                        'bring the main form to top of the z order
                        BringWindowToTop hMainForm
                    Case hBtnIPCancel
                        ShowWindow hFrmIP, SW_HIDE
                        BringWindowToTop hMainForm
                End Select
                Exit Function
            End If
        Case WM_CLOSE
            'just hide, do not destroy
            ShowWindow hFrmIP, SW_HIDE
            'DestroyWindow hFrmIP
            Exit Function
    End Select
    IpFieldWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
