Attribute VB_Name = "Form_ProgressStatusBar"
Option Explicit

'Form Progressbar + Statusbar
Public hFrmPBSB As Long
Public hBtnStartPb As Long
Public hBtnStopPb As Long
Public hChkLoop As Long
Public hPbDefault As Long
Public hPbEx As Long
Public hTrackbar As Long
Public hTxtTrackbar As Long 'Trackbar buddy
Public hPbStatus As Long
Public hPbTimer As Long
Public hStatusbar As Long
Public Const PB_TIMER_ID As Long = NUM_TWO

Public Sub BtnProgressStatusBar_Click()
    If hFrmPBSB = 0 Then
        hFrmPBSB = CreateForm(300, 300, 400, 250, AddressOf ProgressStatusWndProc, hMainForm, WS_SYSMENU Or WS_BORDER Or WS_MINIMIZEBOX, , "FormPBSB", "Progress+Status Bar")
        If hFrmPBSB > 0 Then
            Dim sPartsText(2) As String
            Dim lPartsDrawStyle(2) As Long
            Dim lPartsRightCoord(2) As Long
            Dim hStatusIcon As Long
            Dim sStatusTip As String
            Dim PBRect As RECT
            
            hBtnStartPb = CreateCmdButton(hFrmPBSB, "Start", 5, 5, 100, 30, , , , , , , , WS_EX_DLGMODALFRAME)
            hBtnStopPb = CreateCmdButton(hFrmPBSB, "Stop", 110, 5, 100, 30, , , , , , , , WS_EX_DLGMODALFRAME)
            hChkLoop = CreateCheckbox(hFrmPBSB, "Loop Continous", 5, 40, 150, 25)
            'A_SendMessage hChkLoop, BM_SETCHECK, BST_CHECKED, 0&
            'Create progress bars
            hPbEx = CreateProgressBar(hFrmPBSB, 5, 70, 300, 20, False, 5, , , x_Blue, x_Yellow)
            hPbDefault = CreateProgressBar(hFrmPBSB, 5, 100, 300, 20, , 5, , , , , False)
            'Create Trackbar buddy
            hTxtTrackbar = CreateTextbox(hFrmPBSB, 310, 125, 50, 25, "0", , , , , , , , , , , , True)
            'Create trackbar
            hTrackbar = CreateTrackBar(hFrmPBSB, 5, 140, 300, 30, , , , , , , , 5, hTxtTrackbar, False)
            'Create status bar (simlpe)
            'hStatusbar = CreateStatusBarSimple(hFrmPBSB,"Ready")
            'Just to give the part a no border look
            sPartsText(0) = " "
            sPartsText(2) = "Last One"
            lPartsDrawStyle(0) = SBT_NOBORDERS 'StatusBar_Ts.SBT_POPOUT
            lPartsDrawStyle(1) = 0
            lPartsDrawStyle(2) = StatusBar_Ts.SBT_POPOUT
            lPartsRightCoord(0) = 100
            lPartsRightCoord(1) = 132
            'Fill in the rest
            lPartsRightCoord(2) = -1
            'Create status bar multi part
            hStatusbar = CreateStatusBar(hFrmPBSB, 3, sPartsText, lPartsDrawStyle, lPartsRightCoord, , , , , , , , True)
            If hStatusbar > 0 Then
                'Set the small window icon, we can set the large icon used by Alt+Tab combo also
                hStatusIcon = A_LoadImage(App.hInstance, App.Path & "\images\smiley.ico", IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
                'Set icon
                A_SendMessage hStatusbar, SB_SETICON, 1&, hStatusIcon
                'set tooltip
                'This ToolTip text is displayed in two situations:
                'When the corresponding pane in the status bar contains only an icon.
                'When the corresponding pane in the status bar contains text that is truncated due to the size of the pane.
                sStatusTip = "ToolTip for this one"
                A_SendMessageStr hStatusbar, SB_SETTIPTEXT, 1&, sStatusTip
                'To get statusbar height
                '   GetWindowRect(hStatus, &rcStatus)
                '   iStatusHeight = rcStatus.Bottom - rcStatus.Top
                'To insert a PB in a status part
                '   Use SB_GETRECT Message to get bounding rect of a part in a status window.
                '   Set PB parent to status and pos the PB acording to rect
                
                'Get the first part rect
                A_SendMessageAnyRef hStatusbar, SB_GETRECT, 0&, PBRect
                'Create PB to be hosted in the first part of status bar
                hPbStatus = CreateProgressBar(hStatusbar, PBRect.Left, PBRect.Top, PBRect.Right - 2, PBRect.Bottom - 2, False, 5, , , x_Blue, x_Yellow)
                'Hide it for now
                ShowWindow hPbStatus, SW_HIDE
            End If
            ShowWndForFirstTime hFrmPBSB
        End If
    Else
        ShowWindow hFrmPBSB, SW_SHOW
    End If
End Sub

Public Function ProgressStatusWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Dim nmh As NMHDR
    
    Select Case uMsg
        Case WM_TIMER
            If wParam = PB_TIMER_ID Then 'Timer hit
                'set the pos of progress bars based on checkbox value
                If A_SendMessage(hChkLoop, BM_GETCHECK, 0&, 0&) = BST_CHECKED Then
                    'Use PBM_STEPIT
                    'When the position exceeds the maximum range value,
                    'this message resets the current position so that the
                    'progress indicator starts over again from the beginning.
                    A_SendMessage hPbDefault, PBM_STEPIT, 0&, 0&
                    A_SendMessage hPbEx, PBM_STEPIT, 0&, 0&
                    ShowWindow hPbStatus, SW_SHOW
                    A_SendMessage hPbStatus, PBM_STEPIT, 0&, 0&
                    StatusBar_SetText hStatusbar, CStr(A_SendMessage(hPbDefault, PBM_GETPOS, 0&, 0&)) & "% completed", SBT_POPOUT, 2
                Else
                    'Get current pos, if at the end then stop
                    If A_SendMessage(hPbDefault, PBM_GETPOS, 0&, 0&) < 100 Then
                        'Use PBM_DELTAPOS
                        'Advances the current position of a progress bar by a specified increment
                        'and redraws the bar to reflect the new position.
                        A_SendMessage hPbDefault, PBM_DELTAPOS, 5, 0&
                        A_SendMessage hPbEx, PBM_DELTAPOS, 5, 0&
                        ShowWindow hPbStatus, SW_SHOW
                        A_SendMessage hPbStatus, PBM_DELTAPOS, 5, 0&
                        StatusBar_SetText hStatusbar, CStr(A_SendMessage(hPbDefault, PBM_GETPOS, 0&, 0&)) & "% completed", SBT_POPOUT, 2
                    Else
                        'Stop timer and reset
                        KillTimer hFrmPBSB, PB_TIMER_ID
                        EnableWindow hBtnStartPb, 1&
                        EnableWindow hBtnStopPb, 0&
                        A_SendMessage hPbDefault, PBM_SETPOS, 0&, 0&
                        A_SendMessage hPbEx, PBM_SETPOS, 0&, 0&
                        ShowWindow hPbStatus, SW_HIDE
                        A_SendMessage hPbStatus, PBM_SETPOS, 0&, 0&
                        StatusBar_SetText hStatusbar, "100% Completed", SBT_POPOUT, 2
                    End If
                End If
                Exit Function
            End If
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                If lParam = hBtnStartPb Then
                    EnableWindow hBtnStartPb, 0&
                    EnableWindow hBtnStopPb, 1&
                    'Start timer
                    'Passing 0& to the TimerFunc callback param causes the timer to send
                    'WM_TIMER msg back to us here, no need for timer procedure!!!
                    hPbTimer = SetTimer(hFrmPBSB, PB_TIMER_ID, 250, ByVal 0&)
                ElseIf lParam = hBtnStopPb Then
                    'Stop timer and reset
                    KillTimer hFrmPBSB, PB_TIMER_ID
                    EnableWindow hBtnStartPb, 1&
                    EnableWindow hBtnStopPb, 0&
                'ElseIf lParam = hChkLoop Then
                End If
                Exit Function
            End If
        Case WM_HSCROLL 'Deal with trackbar notification mssages
'TB_BOTTOM
'   VK_END
'TB_ENDTRACK
'   WM_KEYUP(the user released a key that sent a relevant virtual key code)
'TB_LINEDOWN
'   VK_RIGHT Or VK_DOWN
'TB_LINEUP
'   VK_LEFT Or VK_UP
'TB_PAGEDOWN
'   VK_NEXT (the user clicked the channel below or to the right of the slider)
'TB_PAGEUP
'   VK_PRIOR (the user clicked the channel above or to the left of the slider)
'TB_THUMBPOSITION
'   WM_LBUTTONUP following a TB_THUMBTRACK notification message
'TB_THUMBTRACK
'   Slider movement (the user dragged the slider)
'TB_TOP
'   VK_HOME
            'Get notification code
            lRet = LoWord(wParam)
            Select Case lRet
                Case TB_ENDTRACK, TB_LINEUP, TB_LINEDOWN
                    A_SetWindowText hTxtTrackbar, CStr(A_SendMessage(hTrackbar, TBM_GETPOS, 0&, 0&))
                Case TB_THUMBTRACK, TB_THUMBPOSITION
                    A_SetWindowText hTxtTrackbar, CStr(HiWord(wParam))
            End Select
            A_SendMessage hPbDefault, PBM_SETPOS, A_SendMessage(hTrackbar, TBM_GETPOS, 0&, 0&), 0&
        Case WM_CLOSE
            'just hide, do not destroy
            ShowWindow hFrmPBSB, SW_HIDE
            'DestroyWindow hFrmPBSB
            Exit Function
    End Select

    ProgressStatusWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
