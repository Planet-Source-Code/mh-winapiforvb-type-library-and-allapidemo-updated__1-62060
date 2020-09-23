Attribute VB_Name = "Form_Animation"
Option Explicit

'Form animation
Public hFrmAnimate As Long
Public hwndAnim As Long
Public hBtnStart As Long
Public hBtnStop As Long
Public hBtnHideAnim As Long

Public Sub BtnShowAnim_Click()
    If hFrmAnimate = 0 Then
        hFrmAnimate = CreateForm(300, 300, 400, 165, AddressOf AnimWndProc, hMainForm, WS_OVERLAPPED, , "FormAnim", "Sample Animation")
        If hFrmAnimate > 0 Then
            hBtnStart = CreateCmdButton(hFrmAnimate, "Start Animation", 10, 10, 100, 25, , , , False, True)
            hBtnStop = CreateCmdButton(hFrmAnimate, "Stop Animation", 115, 10, 100, 25, , , , False, True)
            EnableWindow hBtnStop, False
            hBtnHideAnim = CreateCmdButton(hFrmAnimate, "Hide Form", 230, 10, 110, 25, , , , False, True)
            hwndAnim = CreateAnimation(hFrmAnimate, 10, 45)
            'Load and show first frame
            A_SendMessageStr hwndAnim, ACM_OPENA, 0&, CStr(App.Path & "\images\WebToFolder.avi")
            ShowWindow hwndAnim, SW_SHOW
            ShowWndForFirstTime hFrmAnimate
        End If
    Else
        ShowWindow hFrmAnimate, SW_SHOW
    End If
End Sub

Public Function AnimWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Select Case uMsg
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                Select Case lParam
                    Case hBtnHideAnim
                        ShowWindow hFrmAnimate, SW_HIDE
                    Case hBtnStart
                        'cRepeat = Number of times to replay the AVI clip. A value of -1 means replay the clip indefinitely.
                        'wFrom = Zero-based index of the frame where playing begins. The value must be less than 65,536. A value of zero means begin with the first frame in the AVI clip.
                        'wTo = Zero-based index of the frame where playing ends. The value must be less than 65,536. A value of -1 means end with the last frame in the AVI clip.
                        A_SendMessage hwndAnim, ACM_PLAY, -1&, MAKELONG(0, -1)
                        EnableWindow hBtnStart, False
                        EnableWindow hBtnStop, True
                    Case hBtnStop
                        A_SendMessage hwndAnim, ACM_STOP, 0&, 0&
                        EnableWindow hBtnStart, True
                        EnableWindow hBtnStop, False
                End Select
                Exit Function
            End If
        Case WM_CLOSE
            ShowWindow hFrmAnimate, SW_HIDE
            'DestroyWindow hFrmAnimate
            Exit Function
    End Select
    AnimWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
