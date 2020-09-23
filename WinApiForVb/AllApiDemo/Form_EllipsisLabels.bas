Attribute VB_Name = "Form_EllipsisLabels"
Option Explicit

'Form ellipses labels (3 styles)
'SS_ENDELLIPSIS
'SS_PATHELLIPSIS
'SS_WORDELLIPSIS
Public hFrmEllipsis As Long
Public hLblEnd As Long
Public hLblEndEllipsis As Long
Public hLblPath As Long
Public hLblPathEllipsis As Long
Public hLblWordEllipsis As Long
Public hLblWord As Long
Public hBtnHideEllipsis As Long

Public Sub BtnShowEllipseLabels_Click()
    If hFrmEllipsis = 0 Then
        hFrmEllipsis = CreateForm(300, 300, 200, 190, AddressOf EllipsisWndProc, hMainForm, WS_OVERLAPPED, , "FormEllipsis", "Sample Ellipsis Labels")
        If hFrmEllipsis > 0 Then
            hLblEnd = CreateLabel(hFrmEllipsis, "ENDELLIPSIS style:", 5, 5, 180, 17)
            hLblEndEllipsis = CreateLabel(hFrmEllipsis, "A very very very very very very very very very long text to display", 5, 25, 170, 17, , SS_ENDELLIPSIS Or SS_SUNKEN)
            hLblPath = CreateLabel(hFrmEllipsis, "PATHELLIPSIS style:", 5, 45, 200, 17)
            hLblPathEllipsis = CreateLabel(hFrmEllipsis, "C:\dir one\dir two\dir three one\dir four\dir five\dir six one\dir seven\", 5, 65, 170, 17, , SS_PATHELLIPSIS Or SS_SUNKEN)
            hLblWord = CreateLabel(hFrmEllipsis, "WORDELLIPSIS style:", 5, 85, 200, 17)
            hLblWordEllipsis = CreateLabel(hFrmEllipsis, "Draws the frame of the static control using the EDGE_ETCHED edge stylerequires some extra work whichis really not necessary for a label", 5, 105, 170, 17, , SS_WORDELLIPSIS Or SS_SUNKEN)
            hBtnHideEllipsis = CreateCmdButton(hFrmEllipsis, "Close", 5, 130, 80, 25, , , , False, True)
            ShowWndForFirstTime hFrmEllipsis
        End If
    Else
        ShowWindow hFrmEllipsis, SW_SHOW
    End If
End Sub

Public Function EllipsisWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRet As Long
    Select Case uMsg
        Case WM_COMMAND
            lRet = HiWord(wParam)
            If lRet = BN_CLICKED Then
                Select Case lParam
                    Case hBtnHideEllipsis
                        ShowWindow hFrmEllipsis, SW_HIDE
                End Select
                Exit Function
            End If
        Case WM_CLOSE
            ShowWindow hFrmEllipsis, SW_HIDE
            'DestroyWindow hFrmEllipsis
            Exit Function
    End Select
    EllipsisWndProc = A_DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function
