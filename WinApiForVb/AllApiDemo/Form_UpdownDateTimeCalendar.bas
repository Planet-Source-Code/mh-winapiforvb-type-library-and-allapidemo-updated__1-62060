Attribute VB_Name = "Form_UpdownDateTimeCalendar"
Option Explicit

'Form UpDown + DateTime Picker
'UpDownDateTimePicker = UDDTP
Public hFrmUDDTP As Long
Public hUpDown As Long
Public hTxtBuddy As Long
Public hDateTime As Long
Public hCalendar As Long
Public hBtnUDDTPOK As Long
Public hBtnUDDTPCancel As Long
Public strDateTime As String

'UpDown dateTime Picker Calendar
Public Sub BtnShowUDDTP_Click()
    If hFrmUDDTP = 0 Then
        hFrmUDDTP = CreateForm(300, 300, 250, 340, AddressOf UDDTPWndProc, hMainForm, WS_OVERLAPPED, , "FormUDDTP", "UpDown and DateTime")
        If hFrmUDDTP > 0 Then
            'Create buddy
            hTxtBuddy = CreateTextbox(hFrmUDDTP, 10, 10, 60, 25, "", , , , , ES_RIGHT, , , , , , , True)
            'Create updown
            hUpDown = CreateUpDown(hFrmUDDTP, 110, 10, 10, 25, False, , False, False)
            'Limit range
            A_SendMessageAnyRef hUpDown, UDM_SETRANGE, 0, ByVal MAKEDWORDInt(0, 300)
            'Set buddy
            A_SendMessageAnyRef hUpDown, UDM_SETBUDDY, hTxtBuddy, 0
            'Set initial pos
            A_SendMessageAnyRef hUpDown, UDM_SETPOS, 0, 0
            'Create DateTime Picker, updown style
            hDateTime = CreateDateTimePicker(hFrmUDDTP, "Date Time", 10, 40, 200, 25, , , , , True)
            'Limit the range of selection
            If hDateTime <> 0 Then
                Dim sysTime(1) As SYSTEMTIME
                sysTime(0).wYear = 2000
                sysTime(0).wMonth = 11
                sysTime(0).wDay = 5
                sysTime(0).wHour = 11
                sysTime(0).wMinute = 23
                sysTime(0).wSecond = 33
                GetSystemTime sysTime(1)
                A_SendMessageAnyRef hDateTime, DTM_SETRANGE, GDTR_MIN, sysTime(0)
            End If
            'create a month calendar
            hCalendar = CreateCalendar(hFrmUDDTP, 10, 70, 200, 200)
            'Create btns
            hBtnUDDTPOK = CreateCmdButton(hFrmUDDTP, "Ok", 10, 275, 80, 30, , , , False, , , , WS_EX_DLGMODALFRAME)
            hBtnUDDTPCancel = CreateCmdButton(hFrmUDDTP, "Cancel", 95, 275, 80, 30, , , , False, , , , WS_EX_DLGMODALFRAME)
            'Show form
            ShowWndForFirstTime hFrmUDDTP
        End If
    Else
        ShowWindow hFrmUDDTP, SW_SHOW
    End If
End Sub

Public Function UDDTPWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    
    Select Case uMsg
        'lParam contains NMHDR type pretty much for all controls that I have seen so far
        'lParam usallypoints to some type, in this case NMDATETIMECHANGE
        'and this type has a memeber called .NMHDR, this is what we are after first
        'to determine the notification code so we can deal with the indivisual notification msg
        Case WM_NOTIFY
            Dim nmh As NMHDR
            CopyMemory nmh, ByVal lParam, Len(nmh)
            '''''DATETIME PICKER notification
            If nmh.code = DTN_DATETIMECHANGE Then
                Dim lpChanged As NMDATETIMECHANGE
                CopyMemory lpChanged, ByVal lParam, Len(lpChanged)
                strDateTime = String_Macros.String_GetMonthToString(lpChanged.st.wMonth) & " " & lpChanged.st.wDay & " " & lpChanged.st.wYear
            End If
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                Select Case lParam
                    Case hBtnUDDTPOK
                        Dim scRet As String
                        Dim sysTime As SYSTEMTIME
                        'Returns nonzero if successful, or zero otherwise.
                        'This message will always fail when applied to month calendar controls
                        'set to the MCS_MULTISELECT style.
                        A_SendMessageAnyRef hCalendar, MCM_GETCURSEL, 0&, sysTime
                        strDateTime = " <DATETIME>>> " & strDateTime & " <CALENDAR>>> " & _
                                        String_Macros.String_GetMonthToString(sysTime.wMonth) & " " & _
                                        sysTime.wDay & " " & sysTime.wYear
                        'Display
                        A_MessageBox hFrmUDDTP, "<UPDOWN>>> " & Window_GetText(hTxtBuddy) & strDateTime, "UpDown DateTime", MB_OK Or MB_ICONINFORMATION
                        ShowWindow hFrmUDDTP, SW_HIDE
                        BringWindowToTop hMainForm
                    Case hBtnUDDTPCancel
                        ShowWindow hFrmUDDTP, SW_HIDE
                        BringWindowToTop hMainForm
                End Select
                Exit Function
            End If
        Case WM_CLOSE
            ShowWindow hFrmUDDTP, SW_HIDE
            'DestroyWindow hFrmUDDTP
            Exit Function
    End Select
    UDDTPWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function

