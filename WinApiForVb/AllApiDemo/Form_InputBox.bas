Attribute VB_Name = "Form_InputBox"
Option Explicit

'Adapted based on MSDN recommandation on how to set up a modal dialog

'Very simple implementation of a Modal dialog/window
'Create window and controls as usual, disable calling form
'start our own msg loop, when user is done then we Close+Destroy this form,
'exit inputbox msg loop and return exe to the calling form along with return values

'Put together by MH
'mehr13@hotmail.com
'Last updated: Sept 2004

'Form InputBox
Public hFrmInputbox As Long
Public hLblInputbox As Long
Public hTxtInputbox As Long
Public hBtnOkInputbox As Long
Public hBtnCancelInputbox As Long
Public lInputBoxRet As Long
Public sInputBoxRet As String

'Returns IDOK = 1, IDCANCEL = 2, or Error = -1
'Sample call:
'If InputBoxApi(hMainForm, "Username", "Enter Credential", "Please enter your name:", sInput) = IDOK Then
'If InputBoxApi(hMainForm, "Password", "Enter Credential", "Please enter Password:", sInput, True) = IDOK Then
Public Function InputBoxApi(hParent As Long, sText As String, sTitle As String, sLabel As String, sOutPut As String, Optional bPassword As Boolean = False) As Long
    If hParent = 0 Then Exit Function

    Dim udtMsg As MSG
    Dim lMsgRet As Long

    lInputBoxRet = NUM_MINUS_ONE
    sInputBoxRet = CHAR_ZERO_LENGTH_STRING
    
    'Create a form as usual
    hFrmInputbox = CreateForm(200, 200, 300, 115, AddressOf InputBoxApiWndProc, hParent, WS_SYSMENU Or WS_CAPTION Or WS_BORDER, , "FrminputBoxApi", sTitle)
    If hFrmInputbox > 0 Then
        'Create controls
        hLblInputbox = CreateLabel(hFrmInputbox, sLabel, 5, 5, 210, 17)
        hTxtInputbox = CreateTextbox(hFrmInputbox, 5, 25, 285, 25, sText, , , , , , , , , bPassword)
        'make Ok btn default
        hBtnOkInputbox = CreateCmdButton(hFrmInputbox, "Ok", 5, 55, 100, 25, , , , , , True)
        hBtnCancelInputbox = CreateCmdButton(hFrmInputbox, "Cancel", 110, 55, 100, 25)
        'Show the inputbox
        ShowWndForFirstTime hFrmInputbox
        'Set focus to textbox
        SetFocusApi hTxtInputbox
        'highlight entire text, passing no value for params
        TextBox_SetSel hTxtInputbox
        'Disbale calling form
        EnableWindow hParent, False
        'start our own message loop here
        'We could have used GetAndDispatch method in modMain
        lMsgRet = A_GetMessage(udtMsg, 0, 0, 0)
        Do While lMsgRet <> 0
            If lMsgRet = -1 Then
                'Debug.Print "Error Error Error"
                'Exit Do
            Else
                If IsDialogMessage(hFrmInputbox, udtMsg) = 0 Then
                    TranslateMessage udtMsg
                    A_DispatchMessage udtMsg
                End If
            End If
            lMsgRet = A_GetMessage(udtMsg, 0, 0, 0)
        Loop
        
        'User is done, either Ok/Close/X btn clicked
        'Enable callin form
        EnableWindow hParent, True
        sOutPut = sInputBoxRet
    End If
    InputBoxApi = lInputBoxRet
End Function

Public Function InputBoxApiWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Select Case uMsg
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                Select Case lParam
                    'If ENTER is hit while textbox has focus, hLblInputbox is used by system?
                    Case hBtnOkInputbox, hLblInputbox
                        lInputBoxRet = IDOK
                        sInputBoxRet = Window_GetText(hTxtInputbox)
                    Case hBtnCancelInputbox
                        lInputBoxRet = IDCANCEL
                End Select
                'Just get rid of the form
                DestroyWindow hFrmInputbox
                Exit Function
            End If
        Case WM_CLOSE 'If X btn was clicked
            lInputBoxRet = IDCANCEL
            DestroyWindow hFrmInputbox
            Exit Function
        Case WM_DESTROY 'Received in response to DestroyWindow call
            'It causes the GetMessage function to return zero and terminates the message loop
            PostQuitMessage 0&
            Exit Function
    End Select
    InputBoxApiWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
