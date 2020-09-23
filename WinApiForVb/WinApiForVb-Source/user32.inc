// ********** USER32 **********

[
	//uuid(64674046-3A82-101B-8181-00AA003743D3),
	helpstring("USER32.DLL - Windows User Functions and Constants"),
	dllname("USER32.DLL")
]
module User {

    [
    
    entry("DrawFrameControl")    ]
    long DrawFrameControl([in] long hdc,[in] RECT * lprc,[in] long uType,[in] long uState);

    [
    
    entry("DrawEdge")    ]
    long DrawEdge([in] long hdc,[in, out] RECT * lprc, [in] long edge,[in] long grfFlags);

    [
    
    entry("DrawCaption")    ]
    long DrawCaption([in] long hwnd,[in] long hdc,[in] RECT * lprc,[in] long uFlags);
	
    [
    
    entry("DrawAnimatedRects")    ]
    long DrawAnimatedRects([in] long hwnd,[in] long idAni, [in] RECT * lprcFrom, [in] RECT * lprcTo);
    
	[
	
	entry("RegisterClipboardFormatA")	]
	long A_RegisterClipboardFormat([in] LPSTR lpszFormat);

	[
	
	entry("RegisterClipboardFormatW")	]
	long W_RegisterClipboardFormat([in] LPWSTR lpszFormat);

    [
    
    entry("OpenClipboard")    ]
    long OpenClipboard([in] long hWndNewOwner);

    [
    
    entry("CloseClipboard")    ]
    long CloseClipboard();

    [
    
    entry("EmptyClipboard")    ]
    long EmptyClipboard();


    [
    
    entry("GetClipboardData")    ]
    long GetClipboardData([in] long uFormat);

    [
    
    entry("IsClipboardFormatAvailable")    ]
    long IsClipboardFormatAvailable([in] long format);

    [
    
    entry("SetClipboardData")    ]
    long SetClipboardData([in] long uFormat, [in] long hMem);


//End clipboard

    [
    
    entry("GetWindowDC")    ]
    long GetWindowDC([in] long hWnd);

    [
    
    entry("LockWindowUpdate")    ]
    long LockWindowUpdate([in] long hWndLock);

    [
    
    entry("WindowFromDC")    ]
    long WindowFromDC([in] long hdc);

	// ****** SendMessage support ANSI UNICODE *******

	[
	
	entry("SendMessageA")	]
	long A_SendMessage([in] long hWnd, [in] long wMsg, [in] long wParam, [in] long lParam);

	[
	
	entry("SendMessageA")	]
	long A_SendMessageAnyRef([in] long hWnd, [in] long wMsg, [in] long wParam, [in] void * lParam);

	[
	
	entry("SendMessageA")	]
	long A_SendMessageAnyVal([in] long hWnd, [in] long wMsg, [in] long wParam, [in] void lParam);

	[
	
	entry("SendMessageA")	]
	long A_SendMessageStr([in] long hWnd, [in] long wMsg, [in] long wParam, [in] LPSTR lParam);

	[
	
	entry("SendMessageA")	]
	long A_SendMessageAnyAnyRef([in] long hWnd, [in] long wMsg, [in] void * wParam, [in] void * lParam);

	[
	
	entry("SendMessageA")	]
	long A_SendMessageAnyAnyVal([in] long hWnd, [in] long wMsg, [in] void wParam, [in] void lParam);

	[
	
	entry("SendMessageW")	]
	long W_SendMessage([in] long hWnd, [in] long wMsg, [in] long wParam, [in] long lParam);

	[
	
	entry("SendMessageW")	]
	long W_SendMessageAnyRef([in] long hWnd, [in] long wMsg, [in] long wParam, [in] void * lParam);

	[
	
	entry("SendMessageW")	]
	long W_SendMessageAnyVal([in] long hWnd, [in] long wMsg, [in] long wParam, [in] void lParam);

	[
	
	entry("SendMessageW")	]
	long W_SendMessageStr([in] long hWnd, [in] long wMsg, [in] long wParam, [in] LPWSTR lParam);

	[
	
	entry("SendMessageW")	]
	long W_SendMessageAnyAnyRef([in] long hWnd, [in] long wMsg, [in] void * wParam, [in] void * lParam);

	[
	
	entry("SendMessageW")	]
	long W_SendMessageAnyAnyVal([in] long hWnd, [in] long wMsg, [in] void wParam, [in] void lParam);

	[
	
	entry("SendMessageTimeoutA")	]
	long A_SendMessageTimeout([in] long hWnd,
							[in] long wMsg,
							[in] long wParam,
							[in] long lParam,
							[in] long fuFlags,
							[in] long uTimeout,
							[in] long *lpdwResult);
	[
	
	entry("SendMessageTimeoutW")	]
	long W_SendMessageTimeout([in] long hWnd,
							[in] long wMsg,
							[in] long wParam,
							[in] long lParam,
							[in] long fuFlags,
							[in] long uTimeout,
							[in] long *lpdwResult);

	[
	
	entry("BroadcastSystemMessageA")	]
	long A_BroadcastSystemMessage([in] long dwFlags, [in] long *lpdwRecipients, [in] long uiMessage, [in] long wParam, [in] long lParam);
 
	[
	
	entry("BroadcastSystemMessageW")	]
	long W_BroadcastSystemMessage([in] long dwFlags, [in] long *lpdwRecipients, [in] long uiMessage, [in] long wParam, [in] long lParam);

/*
Windows 2000/XP: This thread must either belong to the same desktop as the calling thread
or to a process with the same locally unique identifier (LUID). Otherwise,
the function fails and returns ERROR_INVALID_THREAD_ID. 
*/
    [
    
    entry("SendNotifyMessageA")    ]
    long A_SendNotifyMessage([in] long hWnd,[in] long Msg, [in] long wParam,[in] long lParam);

    [
    
    entry("SendNotifyMessageW")    ]
    long W_SendNotifyMessage([in] long hWnd,[in] long Msg, [in] long wParam,[in] long lParam);

    [    
    entry("DispatchMessageA")    ]
    long A_DispatchMessage([in] LPVOID lpMsg);

    [
    entry("DispatchMessageW")    ]
    long W_DispatchMessage([in] LPVOID lpMsg);

    [
    entry("TranslateMessage")    ]
    long TranslateMessage([in] LPVOID lpMsg);

    [
    entry("GetMessageA")    ]
    long A_GetMessage([out] LPVOID lpMsg,[in] long hWnd, [in] long wMsgFilterMin,[in] long wMsgFilterMax);

    [ entry("GetMessageW")    ]
    long W_GetMessage([out] LPVOID lpMsg,[in] long hWnd, [in] long wMsgFilterMin,[in] long wMsgFilterMax);

    [entry("PeekMessageA")    ]
    long A_PeekMessage([out] LPVOID lpMsg,[in] long hWnd, [in] long wMsgFilterMin,[in] long wMsgFilterMax,[in] long wRemoveMsg);

    [
    entry("PeekMessageW")    ]
    long W_PeekMessage([out] LPVOID lpMsg,[in] long hWnd, [in] long wMsgFilterMin,[in] long wMsgFilterMax,[in] long wRemoveMsg);

    [
    entry("GetInputState")    ]
    long GetInputState(void);

    [
    entry("WaitMessage")    ]
    long WaitMessage(void);

    [
    entry("GetMessageTime")    ]
    long GetMessageTime(void);

    [
    
    entry("GetMessagePos")    ]
    long GetMessagePos(void);

    [
    
    entry("SetMessageExtraInfo")    ]
    long SetMessageExtraInfo([in] long lParam);

    [
    
    entry("GetMessageExtraInfo")    ]
    long GetMessageExtraInfo(void);
	
	[entry("IsDialogMessage")]
	BOOL IsDialogMessage([in] HWND hwnd,[in] MSG *lpMsg);

    [    entry("SendMessageCallbackA")    ]
    long A_SendMessageCallback([in] long hWnd,
							[in] long Msg,
							[in] long wParam,
							[in] long lParam,
							[in] long lpResultCallBack,
							[in] long dwData);

    [
    
    entry("SendMessageCallbackW")    ]
    long W_SendMessageCallback([in] long hWnd,
							[in] long Msg,
							[in] long wParam,
							[in] long lParam,
							[in] long lpResultCallBack, //Call back function (addressof xx)
							[in] long dwData);

    [
    
    entry("RegisterWindowMessageA")    ]
    long A_RegisterWindowMessage([in] LPSTR lpsz);
	//long RegisterWindowMessage(LPCTSTR lpsz);

    [
    
    entry("RegisterWindowMessageW")    ]
    long W_RegisterWindowMessage([in] LPWSTR lpsz);

	[
	
	entry("PostMessageA")	]
	long A_PostMessage([in] long hWnd, [in] long wMsg, [in] long wParam, [in] long lParam);

	[
	
	entry("PostMessageA")	]
	long A_PostMessageAnyVal([in] long hWnd, [in] long wMsg, [in] long wParam, [in] void lParam);

	[
	
	entry("PostMessageA")	]
	long A_PostMessageAnyRef([in] long hWnd, [in] long wMsg, [in] long wParam, [in] void * lParam);

	[
	
	entry("PostMessageA")	]
	long A_PostMessageAnyAnyVal([in] long hWnd, [in] long wMsg, [in] void wParam, [in] void lParam);

	[
	
	entry("PostMessageA")	]
	long A_PostMessageAnyAnyRef([in] long hWnd, [in] long wMsg, [in] void * wParam, [in] void * lParam);

	[
	
	entry("PostMessageA")	]
	long A_PostMessageStr([in] long hWnd, [in] long wMsg, [in] long wParam, [in] LPSTR lParam);

	[
	
	entry("PostMessageW")	]
	long W_PostMessage([in] long hWnd, [in] long wMsg, [in] long wParam, [in] long lParam);

	[
	
	entry("PostMessageW")	]
	long W_PostMessageAnyVal([in] long hWnd, [in] long wMsg, [in] long wParam, [in] void lParam);

	[
	
	entry("PostMessageW")	]
	long W_PostMessageAnyRef([in] long hWnd, [in] long wMsg, [in] long wParam, [in] void * lParam);

	[
	
	entry("PostMessageW")	]
	long W_PostMessageAnyAnyVal([in] long hWnd, [in] long wMsg, [in] void wParam, [in] void lParam);

	[
	
	entry("PostMessageW")	]
	long W_PostMessageAnyAnyRef([in] long hWnd, [in] long wMsg, [in] void * wParam, [in] void * lParam);

	[
	
	entry("PostMessageW")	]
	long W_PostMessageStr([in] long hWnd, [in] long wMsg, [in] long wParam, [in] LPWSTR lParam);

	[
	
	entry("PostThreadMessageA")	]
	long A_PostThreadMessage([in] long idThread, [in] long wMsg, [in] long wParam, [in] long lParam);

	[
	
	entry("PostThreadMessageA")	]
	long A_PostThreadMessageAnyVal([in] long idThread, [in] long wMsg, [in] long wParam, [in] void lParam);

	[
	
	entry("PostThreadMessageA")	]
	long A_PostThreadMessageAnyRef([in] long idThread, [in] long wMsg, [in] long wParam, [in] void * lParam);

	[
	
	entry("PostThreadMessageA")	]
	long A_PostThreadMessageAnyAnyVal([in] long idThread, [in] long wMsg, [in] void wParam, [in] void lParam);

	[
	
	entry("PostThreadMessageA")	]
	long A_PostThreadMessageAnyAnyRef([in] long idThread, [in] long wMsg, [in] void * wParam, [in] void * lParam);

	[
	
	entry("PostThreadMessageW")	]
	long W_PostThreadMessage([in] long idThread, [in] long wMsg, [in] long wParam, [in] long lParam);

	[
	
	entry("PostThreadMessageW")	]
	long W_PostThreadMessageAnyVal([in] long idThread, [in] long wMsg, [in] long wParam, [in] void lParam);

	[
	
	entry("PostThreadMessageW")	]
	long W_PostThreadMessageAnyRef([in] long idThread, [in] long wMsg, [in] long wParam, [in] void * lParam);

	[
	
	entry("PostThreadMessageW")	]
	long W_PostThreadMessageAnyAnyVal([in] long idThread, [in] long wMsg, [in] void wParam, [in] void lParam);

	[
	
	entry("PostThreadMessageW")	]
	long W_PostThreadMessageAnyAnyRef([in] long idThread, [in] long wMsg, [in] void * wParam, [in] void * lParam);

    [
    
    entry("PostQuitMessage")    ]
    void PostQuitMessage([in] long nExitCode);

	[
	
	entry("AttachThreadInput")    ]
	long AttachThreadInput([in] long idAttach,[in] long idAttachTo,[in] long fAttach);

/*
// RegisterDeviceNotification

#if(WINVER >= 0x0500)
typedef  PVOID           HDEVNOTIFY;
typedef  HDEVNOTIFY     *PHDEVNOTIFY;

	const long DEVICE_NOTIFY_WINDOW_HANDLE    = 0x00000000;

WINUSERAPI
HDEVNOTIFY
WINAPI
RegisterDeviceNotificationA(
    IN HANDLE hRecipient,
    IN LPVOID NotificationFilter,
    IN DWORD Flags
    );
WINUSERAPI
HDEVNOTIFY
WINAPI
RegisterDeviceNotificationW(
    IN HANDLE hRecipient,
    IN LPVOID NotificationFilter,
    IN DWORD Flags
    );
#ifdef UNICODE
#define RegisterDeviceNotification  RegisterDeviceNotificationW
#else
#define RegisterDeviceNotification  RegisterDeviceNotificationA
#endif // !UNICODE

WINUSERAPI
BOOL
WINAPI
UnregisterDeviceNotification(
    IN HDEVNOTIFY Handle
    );
#endif // WINVER >= 0x0500 
*/

	// ********* HOTKEY ********

    [
    
    entry("RegisterHotKey")    ]
    long RegisterHotKey([in] long hWnd, [in] long id,[in] long fsModifiers,[in] long vk);

    [
    
    entry("UnregisterHotKey")    ]
    long UnregisterHotKey([in] long hWnd, [in] long id);

    // ****** Mouse support *******

    [
    
    entry("BlockInput")    ]
    long BlockInput([in] long fBlockIt);

    [
    
    entry("SwapMouseButton")    ]
    long SwapMouseButton([in] long fSwap);

    [
    
    entry("GetCapture")    ]
    long GetCapture(void);

    [
    
    entry("SetCapture")    ]
    long SetCapture([in] long hWnd);

    [
    
    entry("ReleaseCapture")    ]
    long ReleaseCapture(void);

    [
    
    entry("SetDoubleClickTime")    ]
    long SetDoubleClickTime([in] long uInterval);

    [
    
    entry("GetDoubleClickTime")    ]
    long GetDoubleClickTime(void);



    // ****** Exit Windows support *******

	//define ExitWindows(dwReserved, Code) ExitWindowsEx(EWX_LOGOFF, 0xFFFFFFFF)

    [
    
    entry("ExitWindowsEx")    ]
    long ExitWindowsEx([in] long uFlags,[in] long dwReserved);

	// ********** Window Creation Support **********

	[
    
    entry("CreateWindowExA")    ]
    long A_CreateWindowEx([in] long dwExStyle,
						[in] LPSTR lpClassName,
						[in] LPSTR lpWindowName,
                        [in] long dwStyle,
						[in] long x,
						[in] long y,
						[in] long nWidth,
						[in] long nHeight,
						[in] long hWndParent,
						[in] long hMenu,
						[in] long hInstance,
						[in] void * lpParam); //Pass like byval 0& or CS (As CREATESTRUCT)
	[
    
    entry("CreateWindowExW")    ]
    long W_CreateWindowEx([in] long dwExStyle,
						[in] LPWSTR lpClassName,
						[in] LPWSTR lpWindowName,
                        [in] long dwStyle,
						[in] long x,
						[in] long y,
						[in] long nWidth,
						[in] long nHeight,
						[in] long hWndParent,
						[in] long hMenu,
						[in] long hInstance,
						[in] void * lpParam);

    [
    
    entry("RegisterClassA")    ]
    long A_RegisterClass([in] LPVOID lpWndClass);

    [
    
    entry("RegisterClassW")    ]
    long W_RegisterClass([in] LPVOID lpWndClass);

    [
    
    entry("RegisterClassExA")    ]
    long A_RegisterClassEx([in] LPVOID lpWndClass);

    [
    
    entry("RegisterClassExW")    ]
    long W_RegisterClassEx([in] LPVOID lpWndClass);

    [
    
    entry("UnregisterClassA")    ]
    long A_UnregisterClass([in] LPSTR lpClassName, [in] long hInstance);

    [
    
    entry("UnregisterClassW")    ]
    long W_UnregisterClass([in] LPWSTR lpClassName, [in] long hInstance);

    [
    
    entry("GetClassInfoA")    ]
    long A_GetClassInfo([in] long hInstance, [in] LPSTR lpClassName,[out] LPVOID lpWndClass);

    [
    
    entry("GetClassInfoW")    ]
    long W_GetClassInfo([in] long hInstance, [in] LPWSTR lpClassName,[out] LPVOID lpWndClass);

    [
    
    entry("GetClassInfoExA")    ]
    long A_GetClassInfoEx([in] long hInstance, [in] LPSTR lpClassName,[out] LPVOID lpWndClass);

    [
    
    entry("GetClassInfoExW")    ]
    long W_GetClassInfoEx([in] long hInstance, [in] LPWSTR lpClassName,[out] LPVOID lpWndClass);

    [
    
    entry("GetClassNameA")    ]
    long A_GetClassName([in] long hwnd, [out] LPSTR lpszClassName,[in] long cchClassName);

    [
    
    entry("GetClassNameW")    ]
    long W_GetClassName([in] long hwnd, [out] LPWSTR lpszClassName,[in] long cchClassName);


	// ****** Windows Subclassing support *******

	[
    
    entry("CallWindowProcA")    ]
    long A_CallWindowProc([in] long lpPrevWndFunc,
                        [in] long hwnd,
						[in] long Msg,
                        [in] long wParam,
						[in] long lParam);
	[
    
    entry("CallWindowProcW")    ]
    long W_CallWindowProc([in] long lpPrevWndFunc,
                        [in] long hwnd,
						[in] long Msg,
                        [in] long wParam,
						[in] long lParam);

	[
    
    entry("DefWindowProcA")    ]
    long A_DefWindowProc([in] long hwnd,
						[in] long Msg,
                        [in] long wParam,
						[in] long lParam);
	[
    
    entry("DefWindowProcW")    ]
    long W_DefWindowProc([in] long hwnd,
						[in] long Msg,
                        [in] long wParam,
						[in] long lParam);


	// ********* General Windows
    [
    
    entry("IsWindow")    ]
    long IsWindow([in] long hwnd);

    [
    
    entry("IsChild")    ]
    long IsChild([in] long hwndParent,[in] long hwndChild);

    [
    
    entry("DestroyWindow")    ]
    long DestroyWindow([in] long hwnd);

    [
    
    entry("ShowWindow")    ]
    long ShowWindow([in] long hwnd,[in] long nCmdShow);

    [
    
    entry("ShowWindowAsync")    ]
	long ShowWindowAsync([in] long hWnd,[in] long nCmdShow);

    [
    
    entry("FlashWindow")    ]
    long FlashWindow([in] long hwnd,[in] long fInvert);

    [
    
    entry("FlashWindowEx")    ]
    long FlashWindowEx([in] FLASHWINFO pfwi); 

    [
    
    entry("ShowOwnedPopups")    ]
    long ShowOwnedPopups([in] long hwnd,[in] long fShow);

    [
    
    entry("SetWindowPos")    ]
    long SetWindowPos([in] long hwnd, 
					  [in] long hwndInsertAfter,
                      [in] long x,
					  [in] long y,
                      [in] long cx,
					  [in] long cy,
                      [in] long fuFlags);

    [
    
    entry("GetWindowPlacement")    ]
    long GetWindowPlacement([in] long hwnd,[out] LPVOID lpwndpl);

    [
    
    entry("SetWindowPlacement")    ]
    long SetWindowPlacement([in] long hwnd,[in] LPVOID lpwndpl);

    [
    
    entry("MoveWindow")    ]
    long MoveWindow([in] long hWnd, 
					[in] long x,
					[in] long y, 
					[in] long nWidth,
					[in] long nHeight, 
					[in] long bRepaint);

    [
    
    entry("IsWindowVisible")    ]
    long IsWindowVisible([in] long hwnd);
    
    [
    
    entry("IsIconic")    ]
    long IsIconic([in] long hwnd);

	[
    
    entry("BringWindowToTop")    ]
    long BringWindowToTop([in] long hWnd);

    [
    
    entry("IsZoomed")    ]
    long IsZoomed([in] long hwnd);

    [
    
    entry("CloseWindow")    ]
    long CloseWindow([in] long hwnd);

    [
    
    entry("SetFocus")    ]
    long SetFocusApi([in] long hwnd);

    [
    
    entry("GetActiveWindow")    ]
    long GetActiveWindow(void);

    [
    
    entry("GetFocus")    ]
    long GetFocus(void);

    [
    
    entry("UpdateWindow")    ]
    long UpdateWindow([in] long hWnd);

    [
    
    entry("SetActiveWindow")    ]
    long SetActiveWindow([in] long hwnd);

    [
    
    entry("GetForegroundWindow")    ]
    long GetForegroundWindow(void);

	[
    
    entry("PaintDesktop")    ]
	long PaintDesktop([in] long hdc);

    [
    
    entry("SetForegroundWindow")    ]
    long SetForegroundWindow([in] long hWnd);

    [
    
    entry("GetDC")    ]
    long GetDC([in] long hwnd);

    [
    
    entry("ReleaseDC")    ]
    long ReleaseDC([in] long hwnd,[in] long hdc);


// ********* Keyboard and Mouse ************
    [
    
    entry("GetKeyState")    ]
    short GetKeyState([in] long nKeyState);

    [
    
    entry("GetAsyncKeyState")    ]
    short GetAsyncKeyState([in] long nKeyState);

    [
    
    entry("GetKeyboardState")    ]
    long GetKeyboardState([in] long lpbKeyState);

    [
    
    entry("SetKeyboardState")    ]
    long SetKeyboardState([in] long lpbKeyState);
	
	[
    
    entry("keybd_event")	]
	void keybd_event([in] long bVk,[in] long bScan,[in] long dwFlags,[in] long dwExtraInfo);

	[
    
    entry("mouse_event")	]
	void mouse_event([in] long dwFlags,[in] long dx,[in] long dy,[in] long dwData,[in] long dwExtraInfo);

	// ****** Timer support ********

    [
    
    entry("SetTimer")    ]
    long SetTimer([in] long hWnd, [in] long nIDEvent, 
							[in] long uElapse, [in] long lpTimerFunc);

    [
    
    entry("KillTimer")    ]
    long KillTimer([in] long hWnd, [in] long uIDEvent);

	// ********** Is window unicode
    [
    
    entry("IsWindowUnicode")    ]
	long IsWindowUnicode([in] long hWnd);
	

    // ********** Enabled state

    [
    
    entry("EnableWindow")    ]
    long EnableWindow([in] long hwnd, [in] long fEnable);

    [
    
    entry("IsWindowEnabled")    ]
    long IsWindowEnabled([in] long hwnd);

    // ****** System Metrics ********

    [
    
    entry("GetSystemMetrics")    ]
    long GetSystemMetrics([in] long nIndex);

	// ********** Menus *********

    [
    
    entry("GetMenu")    ]
    long GetMenu([in] long hwnd);

    [
    
    entry("SetMenu")    ]
    long SetMenu([in] long hwnd,[in] long hMenu);


    [
    
    entry("HiliteMenuItem")    ]
    long HiliteMenuItem([in] long hwnd,
						[in] long hMenu,
                        [in] long idHiliteItem,
                        [in] long fuHilite);

    [
    
    entry("GetMenuStringA")    ]
    long A_GetMenuString([in] long hMenu,
						[in] long idItem,
                        [out] LPSTR lpsz,
						[in] long cbMax,
                        [in] long fuFlags);
    [
    
    entry("GetMenuStringW")    ]
    long W_GetMenuString([in] long hMenu,
						[in] long idItem,
                        [out] LPWSTR lpsz,
						[in] long cbMax,
                        [in] long fuFlags);
    [
    
    entry("GetMenuItemInfoA")    ]
    long A_GetMenuItemInfo([in] long hMenu,
						  [in] long uItem,
                          [in] long fByPosition,
                          [in,out] LPVOID  lpmii);

    [
    
    entry("GetMenuItemInfoW")    ]
    long W_GetMenuItemInfo([in] long hMenu,
						  [in] long uItem,
                          [in] long fByPosition,
                          [in,out] LPVOID  lpmii);

    [
    
    entry("DrawMenuBar")    ]
    long DrawMenuBar([in] long hwnd);

    [
    
    entry("GetSystemMenu")    ]
    long GetSystemMenu([in] long hwnd, [in] long fRevert);

    [
    
    entry("CreateMenu")    ]
    long CreateMenu(void);

    [
    
    entry("CreatePopupMenu")    ]
    long CreatePopupMenu(void);

    [
    
    entry("DestroyMenu")    ]
    long DestroyMenu([in] long hMenu);

    [
    
    entry("CheckMenuItem")    ]
    long CheckMenuItem([in] long hMenu, [in] long idCheckItem,[in] long uCheck);


    [
    
    entry("EnableMenuItem")    ]
    long EnableMenuItem([in] long hMenu,[in] long idEnableItem,[in] long uEnable);

    [
    
    entry("GetSubMenu")    ]
    long GetSubMenu([in] long hMenu, [in] long nPos);

    [
    
    entry("GetMenuItemID")    ]
    long GetMenuItemID([in] long hMenu, [in] long pos);

    [
    
    entry("GetMenuItemCount")    ]
    long GetMenuItemCount([in] long hMenu);

    [entry("IsMenu")    ]
    long IsMenu([in] long hmenu);

    [entry("InsertMenuA")    ]
    long A_InsertMenu([in] long hMenu, [in] long idItem,
                              [in] long fuFlags, [in] long idNewItem,
                              [in] long lpNewItem);

    [entry("InsertMenuA")    ]
    long A_InsertMenuStr([in] long hMenu, [in] long idItem,
                              [in] long fuFlags, [in] long idNewItem,
                              [in] LPSTR lpNewItem);

    [entry("InsertMenuW")    ]
    long W_InsertMenu([in] long hMenu, [in] long idItem,
                              [in] long fuFlags, [in] long idNewItem,
                              [in] long lpNewItem);

    [    entry("InsertMenuW")    ]
    long W_InsertMenuStr([in] long hMenu, [in] long idItem,
                              [in] long fuFlags, [in] long idNewItem,
                              [in] LPWSTR lpNewItem);
	
//fByPosition
//[in] Specifies the meaning of uItem. If this parameter is FALSE,
//uItem is a menu item identifier. Otherwise, it is a menu item position. 
	[entry("InsertMenuItemA")]
	long A_InsertMenuItem([in] long hMenu,
						[in] UINT uItem,
						[in] long fByPosition,
						[in] MENUITEMINFO *lpmii);

	[entry("InsertMenuItemW")]
	long W_InsertMenuItem([in] long hMenu,
						[in] UINT uItem,
						[in] long fByPosition,
						[in] MENUITEMINFO *lpmii);

	[entry("SetMenuInfo")]
	BOOL SetMenuInfo([in] long hMenu,
						[in] MENUINFO *lpcmi);
	

	[entry("SetMenuItemInfoA")]
	BOOL A_SetMenuItemInfo([in] long hMenu,
						[in] UINT uItem,
						[in] long fByPosition,
						[in,out] MENUITEMINFO *lpmii);

	[entry("SetMenuItemInfoW")]
	BOOL W_SetMenuItemInfo([in] long hMenu,
						[in] UINT uItem,
						[in] long fByPosition,
						[in,out] MENUITEMINFO *lpmii);


	const DWORD GMDI_USEDISABLED    = 0x0001;
	const DWORD GMDI_GOINTOPOPUPS   = 0x0002;

	[entry("SetMenuDefaultItem")]
	BOOL SetMenuDefaultItem([in] long hMenu,
						[in] UINT uItem,
						[in] UINT fByPos);
	
	[entry("GetMenuDefaultItem")]
	UINT GetMenuDefaultItem([in] long hMenu,
						[in] UINT fByPos,
						[in] UINT gmdiFlags);

	[entry("GetMenuItemRect")]
	BOOL GetMenuItemRect([in] long hWnd,
						    [in] long hMenu,
							[in] UINT uItem,
							[out] RECT *lprcItem);
	
	[entry("MenuItemFromPoint")    ]
	long MenuItemFromPointAPI([in] long hWnd,
							[in] long hMenu,
							[in] POINTAPI ptScreen);

	[entry("MenuItemFromPoint")    ]
	long MenuItemFromPointVB([in] long hWnd,
							[in] long hMenu,
							[in] long X, [in] long Y);

    [entry("AppendMenuA")    ]
    long A_AppendMenu([in] long hMenu, [in] long fuFlags,
                              [in] long idNewItem, [in] long lpNewItem);

    [
    
    entry("AppendMenuA")    ]
    long A_AppendMenuStr([in] long hMenu, [in] long fuFlags,
                              [in] long idNewItem, [in] LPSTR lpNewItem);

    [
    
    entry("AppendMenuW")    ]
    long W_AppendMenu([in] long hMenu, [in] long fuFlags,
                              [in] long idNewItem, [in] long lpNewItem);

    [
    
    entry("AppendMenuW")    ]
    long W_AppendMenuStr([in] long hMenu, [in] long fuFlags,
                              [in] long idNewItem, [in] LPWSTR lpNewItem);

    [
    
    entry("ModifyMenuA")    ]
    long A_ModifyMenu([in] long hMenu, [in] long idItem,
                              [in] long fuFlags, [in] long idNewItem,
                              [in] long lpNewItem);
    [
    
    entry("ModifyMenuA")    ]
    long A_ModifyMenuStr([in] long hMenu, [in] long idItem,
                              [in] long fuFlags, [in] long idNewItem,
                              [in] LPSTR lpNewItem);

    [
    
    entry("ModifyMenuW")    ]
    long W_ModifyMenu([in] long hMenu, [in] long idItem,
                              [in] long fuFlags, [in] long idNewItem,
                              [in] long lpNewItem);

    [
    
    entry("ModifyMenuW")    ]
    long W_ModifyMenuStr([in] long hMenu, [in] long idItem,
                              [in] long fuFlags, [in] long idNewItem,
                              [in] LPWSTR lpNewItem);

    [
    
    entry("RemoveMenu")    ]
    long RemoveMenu([in] long hMenu, [in] long idItem, [in] long fuFlags);

    [
    
    entry("DeleteMenu")    ]
    long DeleteMenu([in] long hMenu, [in] long idItem, [in] long fuFlags);

    [
    
    entry("SetMenuItemBitmaps")    ]
    long SetMenuItemBitmaps([in] long hMenu, [in] long idItem,
                                      [in] long fuFlags,
                                      [in] long hbmUnchecked,
                                      [in] long hbmChecked);
    [
    
    entry("TrackPopupMenu")    ]
    long TrackPopupMenu([in] long hMenu, [in] long fuFlags,
                               [in] long x, [in] long y,
                               [in] long nReserved, [in] long hwnd,
                               [in] RECT *lprc);

    [
    
    entry("TrackPopupMenuEx")    ]
	long TrackPopupMenuEx([in] long hMenu, [in] long fuFlags, [in] long x, [in] long y, 
								 [in] long hWnd, [in] void *lptpm);

    [
    
    entry("DragDetect")    ]
	long DragDetect([in] long hWnd,[in] POINTAPI *pt);

    // ****** Text drawing support ***

    [
    
    entry("DrawTextA")	]
	long A_DrawText([in] long hDC, [in] LPSTR lpString, [in] long nCount,
						[in,out] RECT *lpRect, [in] long uFormat);

    [
    
    entry("DrawTextW")	]
	long W_DrawText([in] long hDC, [in] LPWSTR lpString, [in] long nCount,
						[in,out] RECT *lpRect, [in] long uFormat);

    [
    
    entry("DrawTextExA")    ]
	long A_DrawTextEx([in] long hDC, [in] LPSTR lpchText, [in] long cchText, 
						  [in,out] RECT *lprc, [in] long dwDTFormat, 
						  [in] void * lpDTParams);

    [
    
    entry("DrawTextExW")    ]
	long W_DrawTextEx([in] long hDC, [in] LPWSTR lpchText, [in] long cchText, 
						  [in,out] RECT *lprc, [in] long dwDTFormat, 
						  [in]  void * lpDTParams); //DRAWTEXTPARAMS

   [
    
    entry("GrayStringA")    ]
	long A_GrayString([in] long hDC, [in] long hBrush, [in] long lpOutputFunc, //OutputProc callback 
						   [in] long lpData, [in] long nCount, [in] long X, [in] long Y,
						   [in] long nWidth, [in] long nHeight);
   [
    
    entry("GrayStringA")    ]
	long A_GrayStringStr([in] long hDC, [in] long hBrush, [in] long lpOutputFunc, //OutputProc callback 
						   [in] LPSTR lpData, [in] long nCount, [in] long X, [in] long Y,
						   [in] long nWidth, [in] long nHeight);

   [
    
    entry("GrayStringW")    ]
	long W_GrayString([in] long hDC, [in] long hBrush, [in] long lpOutputFunc, //OutputProc callback 
						   [in] long lpData, [in] long nCount, [in] long X, [in] long Y,
						   [in] long nWidth, [in] long nHeight);

   [
    
    entry("GrayStringW")    ]
	long W_GrayStringStr([in] long hDC, [in] long hBrush, [in] long lpOutputFunc, //OutputProc callback 
						   [in] LPWSTR lpData, [in] long nCount, [in] long X, [in] long Y,
						   [in] long nWidth, [in] long nHeight);

	// Monolithic state-drawing routine 

	[
    
    entry("DrawStateA")    ]
	long A_DrawState([in] long hDC, [in] long hbr, 
						  [in] long lpOutputFunc, 
						  [in] long lData, [in] long wData, 
						  [in] long x, [in] long y, [in] long cx, [in] long cy, [in] long fuFlags);

	[
    
    entry("DrawStateW")    ]
	long W_DrawState([in] long hDC, [in] long hbr, 
						  [in] long lpOutputFunc, 
						  [in] long lData, [in] long wData, 
						  [in] long x, [in] long y, [in] long cx, [in] long cy, [in] long fuFlags);

	[
    
    entry("TabbedTextOutA")    ]
	long A_TabbedTextOut([in] long hDC, [in] long X, [in] long Y,[in] LPSTR lpString, 
							  [in] long nCount, [in] long nTabPositions, 
							  [in] long * lpnTabStopPositions, [in] long nTabOrigin);

	[
    
    entry("TabbedTextOutW")    ]
	long W_TabbedTextOut([in] long hDC, [in] long X, [in] long Y,[in] LPWSTR lpString, 
							  [in] long nCount, [in] long nTabPositions, 
							  [in] long * lpnTabStopPositions, [in] long nTabOrigin);

	[
    
    entry("GetTabbedTextExtentA")    ]
	long A_GetTabbedTextExtent([in] long hDC,[in] LPSTR lpString, 
								     [in] long nCount, [in] long nTabPositions, 
									 [in] long * lpnTabStopPositions);

	[
    
    entry("GetTabbedTextExtentW")    ]
	long W_GetTabbedTextExtent([in] long hDC,[in] LPWSTR lpString, 
								     [in] long nCount, [in] long nTabPositions, 
									 [in] long * lpnTabStopPositions);

	// ****** Window scrolling *****

    [
    
    entry("ScrollWindow")    ]
    long ScrollWindow([in] long hwnd, [in] long XAmount, [in] long YAmount, 
								[in] RECT *lpRect,[in] RECT *lpClipRect);

    [
    
    entry("ScrollDC")    ]
	long ScrollDC([in] long hdc, [in] long dx, [in] long dy,[in] RECT *lprcScroll, 
							[in] RECT *lprcClip, [in] long hrgnUpdate, 
							[out] RECT *lprcUpdate);

    [
    
    entry("ScrollWindowEx")    ]
    long ScrollWindowEx([in] long hwnd, [in] long dx, [in] long dy,
								  [in] RECT *prcScroll,[in] RECT *prcClip,
								  [in] long hrgnUpdate,[out] RECT *prcUpdate, 
								  [in] long flags);

    [
    
    entry("GetScrollInfo")    ]
    long GetScrollInfo([in] long hwnd, [in] long fnBar, [in,out] LPVOID lpsi);

    [
    
    entry("SetScrollInfo")    ]
    long SetScrollInfo([in] long hwnd, [in] long fnBar, [in] LPVOID lpsi, [in] long fRedraw);

    [
    
    entry("ShowScrollBar")    ]
    long ShowScrollBar([in] long hwnd, [in] long fnBar, [in] long fShow);

    [
    
    entry("EnableScrollBar")    ]
    long EnableScrollBar([in] long hwnd, [in] long fnSBFlags, [in] long fuArrowFlags);

    // ****** Window property support *******
    [
    
    entry("SetPropA")    ]
    long A_SetProp([in] long hWnd, [in] LPSTR lpString, [in] long hData);

    [
    
    entry("SetPropW")    ]
    long W_SetProp([in] long hWnd, [in] LPWSTR lpString, [in] long hData);

    [
    
    entry("GetPropA")    ]
    long A_GetProp([in] long hWnd, [in] LPSTR lpString);

    [
    
    entry("GetPropW")    ]
    long W_GetProp([in] long hWnd, [in] LPWSTR lpString);

    [
    
    entry("RemovePropA")    ]
    long A_RemoveProp([in] long hWnd, [in] LPSTR lpString);

    [
    
    entry("RemovePropW")    ]
    long W_RemoveProp([in] long hWnd, [in] LPWSTR lpString);

    [
    
    entry("EnumPropsExA")    ]
    long A_EnumPropsEx([in] long hWnd, [in] long lpEnumFunc, [in] long lParam);

    [
    
    entry("EnumPropsExW")    ]
    long W_EnumPropsEx([in] long hWnd, [in] long lpEnumFunc, [in] long lParam);

    [
    
    entry("EnumPropsA")    ]
    long A_EnumProps([in] long hWnd, [in] long lpEnumProc);

    [
    
    entry("EnumPropsW")    ]
    long W_EnumProps([in] long hWnd, [in] long lpEnumProc);

	// ********* Window text ********

    [
    
    entry("SetWindowTextA")    ]
    long A_SetWindowText([in] long hwnd, [in] LPSTR lpsz); //LPCTSTR lpsz

    [
    
    entry("SetWindowTextW")    ]
    long W_SetWindowText([in] long hwnd, [in] LPWSTR lpsz); //LPWSTR lpsz


    [
    
    entry("GetWindowTextA")    ]
    long A_GetWindowText([in] long hwnd, [out] LPSTR lpsz, [in] long cbMax);

    [
    
    entry("GetWindowTextW")    ]
    long W_GetWindowText([in] long hwnd, [out] LPWSTR lpsz, [in] long cbMax);

    [
    
    entry("GetWindowTextLengthA")    ]
    long A_GetWindowTextLength([in] long hwnd);

    [
    
    entry("GetWindowTextLengthW")    ]
    long W_GetWindowTextLength([in] long hwnd);

    [
    
    entry("GetClientRect")    ]
    long GetClientRect([in] long hwnd,[out] RECT * lpRect);

    [
    
    entry("GetWindowRect"),
    helpstring("Retrieves coordinates of a window's bounding rectangle")
    ]
    long GetWindowRect([in] long hwnd,[out] RECT * lpRect);

    [
    
    entry("AdjustWindowRect")    ]
    long AdjustWindowRect([in,out] RECT * lpRect,[in] long dwStyle, [in] long bMenu);

    [
    
    entry("AdjustWindowRectEx")    ]
    long AdjustWindowRectEx([in,out] RECT * lpRect,[in] long dwStyle, [in] long bMenu, [in] long dwExStyle);

    // ****** Message Box support ******

    [
    
    entry("MessageBoxA")    ]
    long A_MessageBox([in] long hwnd, [in] LPSTR lpText,
                              [in] LPSTR lpCaption, [in] long uType);

    [
    
    entry("MessageBoxW")    ]
    long W_MessageBox([in] long hwnd, [in] LPWSTR lpText,
                              [in] LPWSTR lpCaption, [in] long uType);

    [
    
    entry("MessageBoxExA")    ]
    long A_MessageBoxEx([in] long hwnd, [in] LPSTR lpText,
                                [in] LPSTR lpCaption, [in] long uType, 
								[in] short wLanduageId);

    [
    
    entry("MessageBoxExW")    ]
    long W_MessageBoxEx([in] long hwnd, [in] LPWSTR lpText,
                                [in] LPWSTR lpCaption, [in] long uType, 
								[in] short wLanduageId);

    [
    
    entry("MessageBoxIndirectA")	]
	long A_MessageBoxIndirect([in] LPVOID lpMsgBoxParams);

    [
    
    entry("MessageBoxIndirectW")	]
	long W_MessageBoxIndirect([in] LPVOID lpMsgBoxParams);

    [
    
    entry("MessageBeep")    ]
    long MessageBeep([in] long uType);

    // ****** Cursor support **********

    [
    
    entry("ShowCursor")    ]
    long ShowCursor([in] long bShow);

    [
    
    entry("SetCursorPos")    ]
    long SetCursorPos([in] long x, [in] long y);

    [
    
    entry("SetCursor")    ]
    long SetCursor([in] long hcur);

    [
    
    entry("GetCursorPos")    ]
	long GetCursorPos([out] POINTAPI * lpPoint);

    [
    
    entry("ClipCursor")    ]
    long ClipCursor([in] LPVOID lpRect); //RECT or NULL(0&)

    [
    
    entry("GetClipCursor")    ]
    long GetClipCursor([out] RECT * lpRect);

    [
    
    entry("GetCursor")    ]
    long GetCursor(void);

	// ****** Carat Support *******

    [
    
    entry("CreateCaret")    ]
    long CreateCaret([in] long hwnd, [in] long hBitmap, [in] long nWidth, [in] long nHeight);

    [
    
    entry("GetCaretBlinkTime")    ]
    long GetCaretBlinkTime(void);

    [
    
    entry("SetCaretBlinkTime")    ]
    long SetCaretBlinkTime([in] long uMSeconds);

    [
    
    entry("DestroyCaret")    ]
    long DestroyCaret(void);

    [
    
    entry("HideCaret")    ]
    long HideCaret([in] long hWnd);

    [
    
    entry("ShowCaret")    ]
    long ShowCaret([in] long hWnd);


    [
    
    entry("SetCaretPos")    ]
    long SetCaretPos([in] long X, [in] long Y);

    [
    
    entry("GetCaretPos")    ]
    long GetCaretPos([out] POINTAPI * lpPoint);


    // ****** Window coordinate mapping and hit-testing **************************

    [
    
    entry("ClientToScreen")    ]
    long ClientToScreen([in] long hwnd,[in,out] POINTAPI * lpPoint);

    [
    
    entry("ScreenToClient")    ]
    long ScreenToClient([in] long hwnd,[in] POINTAPI * lpPoint);

    [
    
    entry("MapWindowPoints")    ]
    long MapWindowPoints([in] long hwndFrom, [in] long hwndTo,[in,out] void * lpPoints, [in] long cPoints);

    [
    
    entry("WindowFromPoint")    ]
    long WindowFromPoint([in] POINTAPI *Point);

    [
    
    entry("ChildWindowFromPoint")    ]
    long ChildWindowFromPoint([in] long hwnd, [in] POINTAPI  *Point);

    [
    
    entry("ChildWindowFromPointEx")    ]
	long ChildWindowFromPointEx([in] long hWnd, [in] POINTAPI *pt, [in] long uFlags);

    // ****** Color support *******

    [
    
    entry("GetSysColor")    ]
    long GetSysColor([in] long nIndex);

    [
    
    entry("GetSysColorBrush")    ]
	long GetSysColorBrush([in] long nIndex);

    [
    
    entry("SetSysColors")    ]
    long SetSysColors([in] long cElements,
                                [in] long * lpaElements,
                                [in] long * lpaRgbValues);

    // ****** Rectangle support *****

    [
    
    entry("DrawFocusRect")    ]
	long DrawFocusRect([in] long hDC,[in] RECT *lprc);

    [
    
    entry("FillRect")    ]
	long FillRect([in] long hDC,[in] RECT *lprc, [in] long hbr);

    [
    
    entry("FrameRect")    ]
	long FrameRect([in] long hDC,[in] RECT *lprc, [in] long hbr);

    [
    
    entry("InvertRect")    ]
	long InvertRect([in] long hDC,[in] RECT *lprc);

    [
    
    entry("SetRect")    ]
	long SetRect([out] RECT * lprc, [in] long xLeft, [in] long yTop, [in] long xRight, [in] long yBottom);

    [
    
    entry("SetRectEmpty")    ]
	long SetRectEmpty([out] RECT * lprc);

    [
    
    entry("CopyRect")    ]
	long CopyRect([out] RECT * lprcDst,[in] RECT *lprcSrc); 

    [
    
    entry("InflateRect")    ]
	long InflateRect([in,out] RECT * lprc, [in] long dx, [in] long dy);

    [
    
    entry("IntersectRect")    ]
	long IntersectRect([out] RECT * lprcDst,[in] RECT *lprcSrc1,[in] RECT *lprcSrc2);

    [
    
    entry("UnionRect")    ]
	long UnionRect([out] RECT * lprcDst,[in] RECT *lprcSrc1, [in] RECT *lprcSrc2);

    [
    
    entry("SubtractRect")    ]
	long SubtractRect([out] RECT * lprcDst,[in] RECT *lprcSrc1, [in] RECT *lprcSrc2);

    [
    
    entry("OffsetRect")    ]
	long OffsetRect([in,out] RECT * lprc, [in] long dx, [in] long dy);

    [
    
    entry("IsRectEmpty")    ]
	long IsRectEmpty([in] RECT *lprc);

    [
    
    entry("EqualRect")    ]
	long EqualRect([in] RECT *lprc1,[in] RECT *lprc2);

    [
    
    entry("PtInRect")    ]
    long PtInRect([in] RECT * lprc, [in] long x, [in] long y);


    // ******* Window words *******

    [
    
    entry("GetWindowWord")    ]
    short GetWindowWord([in] long hwnd, [in] long nIndex);

    [
    
    entry("SetWindowWord")    ]
    short SetWindowWord([in] long hwnd, [in] long nIndex, [in] short nValue);

    [
    
    entry("GetWindowLongA")    ]
    long A_GetWindowLong([in] long hwnd, [in] long nIndex);

    [
    
    entry("GetWindowLongW")    ]
    long W_GetWindowLong([in] long hwnd, [in] long nIndex);

    [
    
    entry("SetWindowLongA")    ]
    long A_SetWindowLong([in] long hwnd, [in] long nIndex, [in] long nValue);

    [
    
    entry("SetWindowLongW")    ]
    long W_SetWindowLong([in] long hwnd, [in] long nIndex, [in] long nValue);

    // ******** Class words *********

    [
    
    entry("GetClassWord")    ]
    short GetClassWord([in] long hwnd, [in] long nIndex);

    [
    
    entry("SetClassWord")    ]
    short SetClassWord([in] long hwnd, [in] long nIndex, [in] short nValue);

    [
    
    entry("GetClassLongA")    ]
    long A_GetClassLong([in] long hwnd, [in] long nIndex);

    [
    
    entry("GetClassLongW")    ]
    long W_GetClassLong([in] long hwnd, [in] long nIndex);

    [
    
    entry("SetClassLongA")    ]
    long A_SetClassLong([in] long hwnd, [in] long nIndex, [in] long dwNewLong);

    [
    
    entry("SetClassLongW")    ]
    long W_SetClassLong([in] long hwnd, [in] long nIndex, [in] long dwNewLong);

	// ****** Window query and enumeration ******

    [
    
    entry("GetDesktopWindow")    ]
    long GetDesktopWindow(void);

    [
    
    entry("GetParent")    ]
    long GetParent([in] long hWnd);

    [
    
    entry("GetAncestor")    ]
    long GetAncestor([in] long hwnd, [in] long gaFlags);

    [
    
    entry("SetParent")    ]
    long SetParent([in] long hwndChild, [in] long hwndNewParent);

    [
    
    entry("EnumChildWindows")    ]
    long EnumChildWindows([in] long hWndParent, [in] long lpEnumFunction, [in] long lParam);

    [
    
    entry("FindWindowA")    ]
    long A_FindWindow([in] LPSTR lpszClassName, [in] LPSTR lpszWindow);

    [
    
    entry("FindWindowW")    ]
    long W_FindWindow([in] LPWSTR lpszClassName, [in] LPWSTR lpszWindow);

    [
    
    entry("FindWindowExA")    ]
	long A_FindWindowEx([in] long hwndParent, [in] long hwndChildAfter, [in] LPSTR lpszClass,[in] LPSTR lpszWindow);

    [
    
    entry("FindWindowExW")    ]
	long W_FindWindowEx([in] long hwndParent, [in] long hwndChildAfter, [in] LPWSTR lpszClass,[in] LPWSTR lpszWindow);

    [
    
    entry("EnumWindows")    ]
    long EnumWindows([in] long lpEnumFunction, [in] long lParam);

    [
    
    entry("EnumThreadWindows")    ]
	long EnumThreadWindows([in] long dwThreadId, [in] long lpfn, [in] long lParam);
	/*
	//define EnumTaskWindows(hTask, lpfn, lParam) 
	//       EnumThreadWindows((DWORD)hTask, lpfn, lParam)
	*/

	[
    
    entry("GetTopWindow")    ]
    long GetTopWindow([in] long hwnd);

    [
    
    entry("GetWindowThreadProcessId")    ]
    long GetWindowThreadProcessId([in] long hwnd, [out] long * lpdwProcessId);

    [
    
    entry("GetLastActivePopup")    ]
    long GetLastActivePopup([in] long hwnd);

    [
    
    entry("GetWindow")    ]
    long GetWindow([in] long hwnd, [in] long uCmd);

    [
    
    entry("GetNextWindow")    ]
    long GetNextWindow([in] long hwnd, [in] long fuDirection);

    [
    
    entry("SetWindowsHookEx")    ]
	long SetWindowsHookEx([in] long idHook, [in] long lpfn, [in] long hInstance, [in] long dwThreadId);

    [
    
    entry("UnhookWindowsHookEx")    ]
	long UnhookWindowsHookEx([in] long hHook);


    [
    
    entry("CallNextHookEx")    ]
	long CallNextHookEx([in] long hhk, [in] long code, [in] long wParam, [in] long lParam);

    [
    
    entry("LoadBitmapA")    ]
    long A_LoadBitmap([in] long hInst, [in] LPSTR lpBitmapName);

    [
    
    entry("LoadBitmapW")    ]
    long W_LoadBitmap([in] long hInst, [in] LPWSTR lpBitmapName);

    [
    
    entry("LoadCursorA")    ]
    long A_LoadCursor([in] long hInst, [in] LPSTR lpCursorName);

    [
    
    entry("LoadCursorW")    ]
    long W_LoadCursor([in] long hInst, [in] LPWSTR lpCursorName);

/*
The following is a list of system cursor names and identifiers.

Cursor Name Cursor Identifier 
"AppStarting" OCR_APPSTARTING 
"Arrow" OCR_NORMAL 
"Crosshair" OCR_CROSS 
"Hand" Windows 2000/XP: OCR_HAND 
"Help" OCR_HELP 
"IBeam" OCR_IBEAM 
"Icon" OCR_ICON 
"No" OCR_NO 
"Size" OCR_SIZE 
"SizeAll" OCR_SIZEALL 
"SizeNESW" OCR_SIZENESW 
"SizeNS" OCR_SIZENS 
"SizeNWSE" OCR_SIZENWSE 
"SizeWE" OCR_SIZEWE 
"UpArrow" OCR_UP 
"Wait" OCR_WAIT 

For example, if the WIN.INI file contains the following
[Cursors] 
    Arrow = "arrow.ani" 

Then the following call causes LoadCursorFromFile to obtain cursor data from the file ARROW.ANI. 
LoadCursorFromFile((LPWSTR)OCR_NORMAL) 
*/
    [
    
    entry("LoadCursorFromFileA")    ]
	long A_LoadCursorFromFile([in] long hInst, [in] LPSTR lpFileName);

    [
    
    entry("LoadCursorFromFileW")    ]
	long W_LoadCursorFromFile([in] long hInst, [in] LPWSTR lpFileName);

	[
    
    entry("CreateCursor")	]
    long CreateCursor([in] long hInst,
								[in] long xHotSpot,
								[in] long yHotSpot,
								[in] long nWidth,
								[in] long nHeight,
								[in] void *pvANDPlane,
								[in] void *pvXORPlane);

	[
    
    entry("DestroyCursor")    ]
    long DestroyCursor([in] long hcur);

	[
    
    entry("SetSystemCursor")	]
	long SetSystemCursor([in] long hcur, [in] long id);



    // ****** Icon support **********

    [
    
    entry("LoadIconA")    ]
    long A_LoadIcon([in] long hInstance, [in] LPSTR lpIconName);

    [
    
    entry("LoadIconW")    ]
    long W_LoadIcon([in] long hInstance, [in] LPWSTR lpIconName);

    [
    
    entry("CreateIcon")    ]
    long CreateIcon([in] long hInstance,
                            [in] long nWidth,[in] long nHeight,
                            [in] BYTE cPlanes, [in] BYTE cBitsPixel,
                            [in] void  *lpbANDbits, [in] void *lpbXORbits);

    [
    
    entry("DestroyIcon")    ]
    long DestroyIcon([in] long hicon);

    [
    
    entry("LookupIconIdFromDirectory")    ]
    long LookupIconIdFromDirectory([in] void * presbits, [in] long fIcon);

    [
    
    entry("LookupIconIdFromDirectoryEx")    ]
    long LookupIconIdFromDirectoryEx([in] void  * presbits, [in] long  fIcon,
									       [in] long   cxDesired, [in] long cyDesired, 
										   [in] long  Flags);

    [
    
    entry("CreateIconFromResource")    ]
    long CreateIconFromResource([in] void * presbits,
						        [in] long dwResSize,
						        [in] long fIcon,
						        [in] long dwVer);

    [
    
    entry("CreateIconFromResourceEx")    ]
    long CreateIconFromResourceEx([in] void  * presbits,
							     [in] long   dwResSize,
						         [in] long   fIcon,
						         [in] long   dwVer,
						         [in] long   cxDesired,
								 [in] long   cyDesired,
								 [in] long   Flags);
    // ******* Image support *********

    [
    
    entry("LoadImageA")    ]
    long A_LoadImage([in] long hinst, [in] LPSTR lpszName, [in] long uType,
                    [in] long cxDesired, [in] long cyDesired, [in] long fuLoad);

    [
    
    entry("LoadImageW")    ]
    long W_LoadImage([in] long hinst, [in] LPWSTR lpszName, [in] long uType,
                    [in] long cxDesired, [in] long cyDesired, [in] long fuLoad);

    [
    
    entry("CopyImage")    ]
    long CopyImage([in] long hImage, [in] long uType,
                   [in] long cxDesired, [in] long cyDesired,
                   [in] long fuFlags);

	// ****** Icon support ***

    [
    
    entry("DrawIcon")    ]
    long DrawIcon([in] long hdc, [in] long x, [in] long y,[in] long hicon);

    [
    
    entry("DrawIconEx")    ]
    long DrawIconEx([in] long hdc,
					[in] long xLeft, [in] long yTop,
                    [in] long hIcon,
                    [in] long cxWidth, [in] long cyWidth,
                    [in] long istepIfAniCur,
                    [in] long hbrFlickerFreeDraw,
                    [in] long diFlags);

    [
    
    entry("CreateIconIndirect")    ]
    long CreateIconIndirect([in] LPVOID  piconinfo);


    [
    
    entry("CopyIcon")    ]
    long CopyIcon([in] long hIcon);

    [
    
    entry("GetIconInfo")    ]
    long GetIconInfo([in] long hIcon,[out] LPVOID  piconinfo);

    // ****** String resource **********

    [
    
    entry("LoadStringW")    ]
    long LoadStringW([in] long hInst, [in] long idResource,
                     [in] LPWSTR lpszBuffer, [in] long cbBuffer);

    [
    
    entry("LoadStringA")    ]
    long LoadStringA([in] long hInst, [in] long idResource,
                     [in] LPSTR lpszBuffer, [in] long cbBuffer);

	// ********* MDI support ********

    [
    
    entry("CreateMDIWindowA")    ]
	long A_CreateMDIWindow([in] LPSTR lpClassName, [in] LPSTR lpWindowName,
						  [in] long dwStyle,
						  [in] long X, [in] long Y, 
						  [in] long nWidth, [in] long nHeight,
						  [in] long hWndParent, [in] long hInstance,
						  [in] long lParam);

    [
    
    entry("CreateMDIWindowW")    ]
	long W_CreateMDIWindow([in] LPWSTR lpClassName, [in] LPWSTR lpWindowName,
						  [in] long dwStyle,
						  [in] long X, [in] long Y, 
						  [in] long nWidth, [in] long nHeight,
						  [in] long hWndParent, [in] long hInstance,
						  [in] long lParam);
    [
    
    entry("ArrangeIconicWindows")    ]
	long ArrangeIconicWindows([in] long hwnd);

    [
    
    entry("TileWindows")    ]
	short TileWindows([in] long hwndParent, [in] long wHow, 
							   [in] RECT *lpRect, [in] long cKids, [in] void *lpKids);

    [
    
    entry("CascadeWindows")    ]
	short CascadeWindows([in] long hwndParent, [in] long wHow, 
								  [in] RECT *lpRect, [in] long cKids, [in] void *lpKids);

    [
    
    entry("TranslateMDISysAccel")    ]
	long TranslateMDISysAccel([in] long hwnd,[in] LPVOID msg);

    [
    
    entry("WinHelpA")    ]
    long A_WinHelp([in] long hwndMain,
				 [in] LPSTR lpszHelp,
                 [in] long usCommand,
				 [in] long ulData);

    [
    
    entry("WinHelpW")    ]
    long W_WinHelp([in] long hwndMain,
				 [in] LPWSTR lpszHelp,
                 [in] long usCommand,
				 [in] long ulData);

    [
    
    entry("ChangeDisplaySettings")    ]
	long A_ChangeDisplaySettings([in] LPVOID lpDevMode, [in] long dwFlags);

    [
    
    entry("ChangeDisplaySettings")    ]
	long W_ChangeDisplaySettings([in] LPVOID lpDevMode, [in] long dwFlags);

    [
    
    entry("ChangeDisplaySettingsEx")    ]
	long A_ChangeDisplaySettingsEx([in] long lpszDeviceName, 
								[in] LPVOID lpDevMode, [in] long hwnd, 
								[in] long dwflags, [in] void lParam);

    [
    
    entry("ChangeDisplaySettingsEx")    ]
	long W_ChangeDisplaySettingsEx([in] long lpszDeviceName, 
								[in] LPVOID lpDevMode, [in] long hwnd, 
								[in] long dwflags, [in] void lParam);

//	long EnumDisplaySettingsA([in] long lpszDeviceName, [in] long iModeNum, [out] DEVMODE *lpDevMode);
//	long EnumDisplaySettingsW([in] long lpszDeviceName, [in] long iModeNum, [out] DEVMODE *lpDevMode);

    [
    
    entry("SystemParametersInfoA")    ]
    long A_SystemParametersInfo([in] long uiAction, 
										[in] long uiParam, 
										[in] void * pvParam, 
										[in] long fWinIni);

    [
    
    entry("SystemParametersInfoW")    ]
    long W_SystemParametersInfo([in] long uiAction, 
										[in] long uiParam, 
										[in] void * pvParam, 
										[in] long fWinIni);
    [
    
    entry("SystemParametersInfoA")    ]
    long A_SystemParametersInfoStr([in] long uiAction, 
										   [in] long uiParam, 
										   [in] LPSTR pvParam, 
										   [in] long fWinIni);

    [
    
    entry("SystemParametersInfoW")    ]
    long W_SystemParametersInfoStr([in] long uiAction, 
										   [in] long uiParam, 
										   [in] LPWSTR pvParam, 
										   [in] long fWinIni);

	// ******* Accessibility support *******

    [
    
    entry("SetLastErrorEx")    ]
	long SetLastErrorEx([in] long dwErrCode, [in] long dwType);

    [
    
    entry("DestroyAcceleratorTable")    ]
	long DestroyAcceleratorTable([in] long hAccel);

	// ******* Char support
    [
    
    entry("CharUpperA")    ]
	long A_ByteArrayUpper(
		[in,out] BYTE * lpsz); 

    [
    
    entry("CharUpperW")    ]
	long W_ByteArrayUpper(
		[in,out] BYTE * lpsz); 

    [
    
    entry("CharUpperA")    ]
	long A_StringUpper(
		[in,out] LPSTR lpsz); //for string

    [
    
    entry("CharUpperW")    ]
	long W_StringUpper(
		[in,out] LPWSTR lpsz); //for string

    [
    
    entry("CharUpperA")    ]
	long A_CharUpper(
		[in,out] long * lpsz);

    [
    
    entry("CharUpperW")    ]
	long W_CharUpper(
		[in,out] long * lpsz);

    [
    
    entry("CharUpperA")    ]
	long A_ByteArrayLower(
		[in,out] BYTE * lpsz); 

    [
    
    entry("CharLowerW")    ]
	long W_ByteArrayLower(
		[in,out] BYTE * lpsz); 

    [
    
    entry("CharLowerA")    ]
	long A_StringLower(
		[in,out] LPSTR lpsz); //for string

    [
    
    entry("CharLowerW")    ]
	long W_StringLower(
		[in,out] LPWSTR lpsz); //for string

    [
    
    entry("CharLowerA")    ]
	long A_CharLower(
		[in,out] long * lpsz); //for string

    [
    
    entry("CharLowerW")    ]
	long W_CharLower(
		[in,out] long * lpsz); //for string

	[
    
    entry("CharUpperBuffA")    ]
	DWORD A_StringUpperByLen([in,out] LPSTR lpsz,[in] DWORD cchLength);

	[
    
    entry("CharUpperBuffW")    ]
	DWORD W_StringUpperByLen([in,out] LPWSTR lpsz,[in] DWORD cchLength);

	[
    
    entry("CharUpperBuffA")    ]
	DWORD A_ByteArrayUpperByLen([in,out] BYTE * lpsz,[in] DWORD cchLength);

	[
    
    entry("CharUpperBuffW")    ]
	DWORD W_ByteArrayUpperByLen([in,out] BYTE * lpsz,[in] DWORD cchLength);

	[
    
    entry("CharLowerBuffA")    ]
	DWORD A_StringLowerByLen([in,out] LPSTR lpsz,[in] DWORD cchLength);

	[
    
    entry("CharLowerBuffW")    ]
	DWORD W_StringLowerByLen([in,out] LPWSTR lpsz,[in] DWORD cchLength);

	[
    
    entry("CharLowerBuffA")    ]
	DWORD A_ByteArrayLowerByLen([in,out] BYTE * lpsz,[in] DWORD cchLength);

	[
    
    entry("CharLowerBuffW")    ]
	DWORD W_ByteArrayLowerByLen([in,out] BYTE * lpsz,[in] DWORD cchLength);

	[
    
    entry("IsCharAlphaA")    ]
	BOOL A_IsCharAlpha([in] short ch);

	[
    
    entry("IsCharAlphaW")    ]
	BOOL W_IsCharAlpha([in] short ch);

	[
    
    entry("IsCharAlphaA")   ]
	BOOL A_IsByteAlpha([in] BYTE ch);

	[
    
    entry("IsCharAlphaW")    ]
	BOOL W_IsByteAlpha([in] BYTE ch);

	[
    
    entry("IsCharAlphaNumericA")    ]
	BOOL A_IsCharAlphaNumeric([in] short ch);

	[
    
    entry("IsCharAlphaNumericW")    ]
	BOOL W_IsCharAlphaNumeric([in] short ch);

	[
    
    entry("IsCharAlphaNumericA")    ]
	BOOL A_IsByteAlphaNumeric([in] BYTE ch);

	[
    
    entry("IsCharAlphaNumericW")    ]
	BOOL W_IsByteAlphaNumeric([in] BYTE ch);

	[
    
    entry("IsCharUpperA")    ]
	BOOL A_IsCharUpper([in] short ch);

	[
    
    entry("IsCharUpperW")    ]
	BOOL W_IsCharUpper([in] short ch);

	[
    
    entry("IsCharUpperA")    ]
	BOOL A_IsByteUpper([in] BYTE ch);

	[
    
    entry("IsCharUpperW")    ]
	BOOL W_IsByteUpper([in] BYTE ch);

	[
    
    entry("IsCharLowerA")    ]
	BOOL A_IsCharLower([in] short ch);

	[
    
    entry("IsCharLowerW")    ]
	BOOL W_IsCharLower([in] short ch);

	[
    
    entry("IsCharLowerA")    ]
	BOOL A_IsByteLower([in] BYTE ch);

	[ entry("IsCharLowerW")    ]
	BOOL W_IsByteLower([in] BYTE ch);
	
	[ entry("InvalidateRect")    ]
    void InvalidateRect([in] HWND hwnd,[in] RECT *lpRect,[in] BOOL bErase);
	
	[ entry("ValidateRect")    ]
    void ValidateRect([in] HWND hwnd,[in] RECT *lpRect);
	
	[ entry("InvalidateRgn")    ]
    void InvalidateRgn([in] HWND hwnd,[in] long hrgn,[in] BOOL bErase );

	[entry("ValidateRgn")]
    void ValidateRgn([in] HWND hwnd, [in] long hrgn);

	[entry("BeginPaint")]
    long BeginPaint([in] HWND hwnd,[out] PAINTSTRUCT * Hdc);

    [entry("EndPaint")]
	BOOL EndPaint([in] HWND hwnd, [in] PAINTSTRUCT * Hdc);
	
	[entry("RedrawWindow")]
	BOOL RedrawWindow(
	  [in] HWND hWnd,               // handle to window
	  [in] RECT *lprcUpdate,			// update rectangle
	  [in] HRGN hrgnUpdate,         // handle to update region
	  [in] UINT flags               // array of redraw flags
	);

//    BOOL    WINAPI GetUpdateRect(HWND hwnd, RECT *, BOOL);
//    int     WINAPI GetUpdateRgn(HWND hwnd, HRGN, BOOL);

//	int WINAPI SetWindowRgn(HWND hWnd, HRGN hRgn, BOOL bRedraw);
//	int WINAPI GetWindowRgn(HWND hWnd, HRGN hRgn);

//    int     WINAPI ExcludeUpdateRgn(HDC hdc, HWND hwnd);

	[entry("GetDialogBaseUnits")]
	LONG GetDialogBaseUnits(void);

	[entry("GetComboBoxInfo")]
	BOOL GetComboBoxInfo([in] HWND hwndCombo, [out] COMBOBOXINFO *pcbi);

}; //End USER32
