Attribute VB_Name = "modMain"
Option Explicit

Public inIDE As Boolean

''''''''''''''''''''''''''''''''''''''
''''''''''APP STARTING POINT
''''''''''''''''''''''''''''''''''''''

'Important: Run the code either in full compile or create an exe to avoid any possible errors

'All regular forms are in modules prefixed as "Form_".
'All macros (helper functions) are in modules suffixed as "_Macros"
'I tried to use a naming convension that (hopefully) will make the code a bit more readable.
 
'In each Form_, you will find the declaration of controls (as long vars),
'a wndproc function to handle it's msgs, and other related functions.

'All forms are created and initilized by calls from MainWndProc in Form_Main module
'Not much error handling or logic checking has been implemented. This is a demo (reference)
'to see how to create and use these controls.

'If you do not see a control in runtime, please check the styles of the control
'Most of the time the style is not compatible with the Comctl version on the OS.

'Put together by MH
'Contact: mehr13@hotmail.com
'Created: Sept 2004
'Last updated: June 2005


''''''''''''''''''''''''''''''''''
''Main App Entry point
''''''''''''''''''''''''''''''''''

Sub Main()
    On Error GoTo Main_Error

'ICC_LINK_CLASS 'XP only
'   Load a hyperlink control class.
'ICC_NATIVEFNTCTL_CLASS
'   Load a native font control class.
'ICC_PAGESCROLLER_CLASS
'   Load pager control class.
'ICC_STANDARD_CLASSES
'   Load one of the intrinsic User32 control classes. The user controls include
'   button, edit, static, listbox, combobox, and scrollbar.
'ICC_USEREX_CLASSES
'   Load ComboBoxEx class.
'ICC_WIN95_CLASSES
'   Load animate control, header, hot key, list-view, progress bar, status bar, tab,
'   ToolTip, toolbar, trackbar, tree-view, and up-down control classes.

    inIDE = iif(App.LogMode = 0, True, False)
    
    'Initialize needed common control classes
    Dim iccex As INITCOMMONCONTROLSEX
    With iccex
        .dwSize = LenB(iccex)
        .dwICC = ICC_WIN95_CLASSES Or ICC_INTERNET_CLASSES Or ICC_USEREX_CLASSES Or _
                ICC_COOL_CLASSES Or ICC_DATE_CLASSES Or ICC_PAGESCROLLER_CLASS
    End With
    INITCOMMONCONTROLSEX iccex
    
    'To get scrollbars on a form combine these flags:
    '   WS_OVERLAPPEDWINDOW Or WS_VSCROLL Or WS_HSCROLL
    hMainForm = CreateForm(200, 200, 340, 390, AddressOf MainWndProc, , , , "MyCustomClass", "App Main Form")
     
    If hMainForm > 0 Then
        Dim hIcon As Long
        'Set the small window icon, we can set the large icon used by Alt+Tab combo also
        hIcon = A_LoadImage(App.hInstance, App.Path & "\images\smiley.ico", IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        A_SendMessage hMainForm, WM_SETICON, ICON_SMALL, hIcon
        
        'Creates a simple frame, notice no var to hold the hwnd for this label
        CreateLabel hMainForm, "", 5, 5, 320, 320, , SS_ETCHEDFRAME
        'Create sunken static (label)
        hLabel = CreateLabel(hMainForm, "ComCtl32 controls made by API. Do not hit END in IDE...", 10, 10, 310, 20, SS_CENTER, SS_SUNKEN)
        
        'Set up buttons various styles
        hBtnshowIp = CreateCmdButton(hMainForm, "Show IP Field", 10, 40, 150, 30, , , , False, True)
        hBtnProgressStatusBar = CreateCmdButton(hMainForm, "Show Progress+Status bars", 165, 40, 150, 30, , , , False, True)
        
        hBtnShowAnim = CreateCmdButton(hMainForm, "Show Animation", 10, 75, 150, 30, , , , , , , , WS_EX_DLGMODALFRAME)
        hBtnShowHotkey = CreateCmdButton(hMainForm, "Set Hotkey Combo", 165, 75, 150, 30, , , , , , , , WS_EX_DLGMODALFRAME)
        
        hBtnShowToolbar = CreateCmdButton(hMainForm, "Show Toolbar+RichEdit", 10, 110, 150, 30, , , , , , , , WS_EX_CLIENTEDGE)
        hBtnShowTabControl = CreateCmdButton(hMainForm, "Show Tab Control", 165, 110, 150, 30, , , , , , , , WS_EX_CLIENTEDGE)
        
        hBtnShowEllipseLabels = CreateCmdButton(hMainForm, "Show Ellipsis Labels", 10, 145, 150, 30, , , , , , , , WS_EX_STATICEDGE)
        hBtnShowListView = CreateCmdButton(hMainForm, "Show ListView", 165, 145, 150, 30, , , , , , , , WS_EX_STATICEDGE)
        
        hBtnShowUDDTP = CreateCmdButton(hMainForm, "Show UpDown DateTime", 10, 180, 150, 30)
        hBtnShowImages = CreateCmdButton(hMainForm, "Show Images", 165, 180, 150, 30)
        
        hBtnShowComboEx = CreateCmdButton(hMainForm, "Show ComboEx", 10, 215, 150, 30)
        
        
        hBtnShowTreeView = CreateCmdButton(hMainForm, "Show TreeView", 10, 250, 150, 30)
        hBtnShowPageScroller = CreateCmdButton(hMainForm, "Show PageScroller", 165, 250, 150, 30)
        
        hBtnShowTGCR = CreateCmdButton(hMainForm, "Show Text-Radio-Check", 10, 285, 150, 30, , , , False, , , , WS_EX_DLGMODALFRAME)
        hBtnShowListBox = CreateCmdButton(hMainForm, "Show ListBox", 165, 285, 150, 30)
        
        'Create tooltip with no parent so we can add tools to it as we go
        hMainToolTip = CreateToolTip(, True, True, "Tooltip Help", 1, x_Blue, x_Aquamarine)
        If hMainToolTip > 0 Then
            'Add tip to someof the buttons
            ToolTip_AddTool hMainToolTip, hBtnshowIp, hBtnshowIp, "Click to view a sample of IP address field." & vbCrLf & "And here is the second line for this tooltip"
            ToolTip_AddTool hMainToolTip, hBtnShowAnim, hBtnShowAnim, "Click to view a sample of animation control"
            ToolTip_AddTool hMainToolTip, hBtnShowToolbar, hBtnShowToolbar, "Click to view a sample Coolbar, toolbar, controls"
            ToolTip_AddTool hMainToolTip, hBtnShowEllipseLabels, hBtnShowEllipseLabels, "Click to view a sample of labels with Ellipsis three styles"
            ToolTip_AddTool hMainToolTip, hBtnShowHotkey, hBtnShowHotkey, "Click to view hotkey control." & vbCrLf & "Enter desired Hotkey combo in the Hotkey field. Ctlr+Alt+A" & vbCrLf & "Then click this button again to set the global hotkey"
        End If
        
        'Show and set focus to the form
        ShowWndForFirstTime hMainForm
                
        'Start message loop
        'We are not using VB msg handling, be very carefull!!!
        GetAndDispatch hMainForm
    Else
        'Debug.Print "hMainForm: " & hMainForm
        A_MessageBox GetDesktopWindow, "Unable to create the main entry point.", "Unable to load", MB_OK Or MB_ICONERROR
    End If

    Exit Sub
Main_Error:
    A_MessageBox GetDesktopWindow, "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module modMain", "FATAL ERROR", MB_OK Or MB_ICONERROR
    DestroyWindow hMainForm
End Sub

''''''''''''''''''''''''''''''''''
''Main App Message Loop
''Idle handling enabled
''''''''''''''''''''''''''''''''''

'To dispatch msgs of a form after creating all elements
'Also, if a form other than the main form requires tabs + other system keyboard and mouse msgs enabled
'None of the forms created except for Inputbox has a msg loop of it's own, so the tabbing will do nothing.
'sample call => GetAndDispatch hForm
Public Sub GetAndDispatch(hParent As Long) ', sClass As String)
    Dim udtMsg As MSG
    Dim lMsgRet As Long
    Dim bDoIdle As Boolean
    Dim nIdleCount As Long

    If hParent = 0 Then Exit Sub
    
    'If the function retrieves a message other than WM_QUIT, the return value is nonzero.
    'If the function retrieves the WM_QUIT message, the return value is zero.
    'Return value of -1 indicates error, should be handled
    'lMsgRet = A_GetMessage(udtMsg, 0, 0, 0)
    bDoIdle = True
    Do While True 'lMsgRet <> 0
        Do While bDoIdle And (A_PeekMessage(udtMsg, 0&, 0, 0, PM_NOREMOVE) = 0)
            'Call any function, ie OnIdle to perform GUI updates and ...
            nIdleCount = nIdleCount + 1
            If OnIdle(nIdleCount) = False Then
                bDoIdle = False
            End If
        Loop
        lMsgRet = A_GetMessage(udtMsg, 0, 0, 0)
        'Check return values
        If lMsgRet = -1 Then
            'Do not process this and continue
        ElseIf lMsgRet = 0 Then
            'WM_QUIT, exit message loop
            Exit Do
        Else
            'MSDN states that the udtMsg passed to IsDialogMessage should not be modified!
            'Checks for keyboard messages and converts them into selections for the corresponding dialog box.
            'For example, the TAB key, when pressed, selects the next control
            'Because the IsDialogMessage function performs all necessary translating and dispatching of messages,
            'a message processed by IsDialogMessage must not be passed to the TranslateMessage or
            'DispatchMessage function.
            If IsDialogMessage(hParent, udtMsg) = 0 Then  'Can not process then pass it to TranslateMessage
                TranslateMessage udtMsg 'translates virtual-key messages into character messages.
                A_DispatchMessage udtMsg 'Dispatch Msg to window
            End If
        End If
        'Handle idle messages
        If IsIdleMessage(udtMsg) = True Then
            bDoIdle = True
            nIdleCount = 0
        End If
    Loop

'This call is optional, just to comply with win apps that send an exit code
'to indicate failure or success
' Call ExitProcess as the last action before closing
' otherwise it prevents proper clean up
' Attention: ExitProcess will also quit the VB IDE!
'    If inIDE = False Then ExitProcess 0&

    'Before calling this function, an application must destroy all windows created with the specified class.
    'All window classes that an application registers are unregistered when it terminates.
    'A_UnregisterClass sClass, App.hInstance

End Sub

'Determines if a message is an idle one
Public Function IsIdleMessage(pMsg As MSG) As Boolean
    'These messages should NOT cause idle processing
    Select Case pMsg.message
        Case WM_MOUSEMOVE:
'Conditional for CE
'#ifndef _WIN32_WCE
'      case WM_NCMOUSEMOVE:
'#endif 'Not _WIN32_WCE
        Case WM_PAINT:
        Case WM_SYSTIMER:    '(caret blink)
            IsIdleMessage = False
    End Select
    IsIdleMessage = True
End Function

'Handle idle moments here, update GUI, dl updates, ...
Public Function OnIdle(IdleCount As Long) As Boolean
    OnIdle = False
End Function
