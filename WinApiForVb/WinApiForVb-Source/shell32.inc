// Shell function entries and constants
[
	helpstring("Shell32 functins and consts"),
	dllname("SHELL32.DLL")
]
module Shell32 {

    // Regular WinExec() codes
    const long SE_ERR_FNF                = 2;       // file not found
    const long SE_ERR_PNF                = 3;       // path not found
    const long SE_ERR_ACCESSDENIED       = 5;       // access denied
    const long SE_ERR_OOM                = 8;       // out of memory
    const long SE_ERR_DLLNOTFOUND       = 32;


    // Error values for ShellExecute() beyond the regular WinExec() codes
    const long SE_ERR_SHARE                  = 26;
    const long SE_ERR_ASSOCINCOMPLETE        = 27;
    const long SE_ERR_DDETIMEOUT             = 28;
    const long SE_ERR_DDEFAIL                = 29;
    const long SE_ERR_DDEBUSY                = 30;
    const long SE_ERR_NOASSOC                = 31;

    // Note CLASSKEY overrides CLASSNAME
    const DWORD SEE_MASK_CLASSNAME        = 0x00000001;
    const DWORD SEE_MASK_CLASSKEY         = 0x00000003;
    // Note INVOKEIDLIST overrides IDLIST
    const DWORD SEE_MASK_IDLIST           = 0x00000004;
    const DWORD SEE_MASK_INVOKEIDLIST     = 0x0000000c;
    const DWORD SEE_MASK_ICON             = 0x00000010;
    const DWORD SEE_MASK_HOTKEY           = 0x00000020;
    const DWORD SEE_MASK_NOCLOSEPROCESS   = 0x00000040;
    const DWORD SEE_MASK_CONNECTNETDRV    = 0x00000080;
    const DWORD SEE_MASK_FLAG_DDEWAIT     = 0x00000100;
    const DWORD SEE_MASK_DOENVSUBST       = 0x00000200;
    const DWORD SEE_MASK_FLAG_NO_UI       = 0x00000400;
    const DWORD SEE_MASK_UNICODE          = 0x00010000;

    [
    
    entry("ShellExecuteA")    ]
    long A_ShellExecute([in] long hwnd,
                                  [in] LPSTR lpOperation,
                                  [in] LPSTR lpFile,
                                  [in] LPSTR lpParameters,
                                  [in] LPSTR lpDirectory,
                                  [in] long nShowCmd); // SW_NORMAL
    [
    
    entry("ShellExecuteW")    ]
    long W_ShellExecute([in] long hwnd,
                                  [in] LPWSTR lpOperation,
                                  [in] LPWSTR lpFile,
                                  [in] LPWSTR lpParameters,
                                  [in] LPWSTR lpDirectory,
                                  [in] long nShowCmd); // SW_NORMAL
	
    [
    
    entry("ShellExecuteExA")    ]
	BOOL A_ShellExecuteEx(
		[in] LPVOID lpExecInfo);

    [
    
    entry("ShellExecuteExW")    ]
	BOOL W_ShellExecuteEx(
		[in] LPVOID lpExecInfo);

    [
    
    entry("SHFileOperationA")    ]
    long A_SHFileOperation([in] LPVOID lpFileOp);


    [
    
    entry("SHFileOperationW")    ]
    long W_SHFileOperation([in] LPVOID  lpFileOp);

    [
    
    entry("FindExecutableA")    ]
    long A_FindExecutable([in] LPSTR lpFile,
                          [in] LPSTR lpDirectory,
                          [out] LPSTR lpResult);

    [
    
    entry("FindExecutableW")    ]
    long W_FindExecutable([in] LPWSTR lpFile,
                          [in] LPWSTR lpDirectory,
                          [out] LPWSTR lpResult);

    [
    
    entry("ShellAboutA")    ]
    long A_ShellAbout([in] long hWnd,
						[in] LPSTR szApp,
                        [in] LPSTR szOtherStuff,
                        [in] long hIcon);

    [
    
    entry("ShellAboutW")    ]
    long W_ShellAbout([in] long hWnd,
						[in] LPWSTR szApp,
                        [in] LPWSTR szOtherStuff,
                        [in] long hIcon);

    [
    
    entry("ExtractAssociatedIconA")    ]
    long A_ExtractAssociatedIcon([in] long hInst,
                               [in] LPSTR lpIconPath,
                               [in,out] WORD * lpiIcon);

    [
    
    entry("ExtractAssociatedIconW")    ]
    long W_ExtractAssociatedIcon([in] long hInst,
                               [in] LPWSTR lpIconPath,
                               [in,out] WORD * lpiIcon);

    [
    
    entry("DuplicateIcon")    ]
    long DuplicateIcon(
			[in] long hInst,
			[in] long hIcon);
	
	[
    
    entry("ExtractIconA")    ]
    long A_ExtractIcon([in] long hInst,
                     [in] LPSTR lpszExeFileName,
                     [in] UINT nIconIndex);

	[
    
    entry("ExtractIconW")    ]
    long W_ExtractIcon([in] long hInst,
                     [in] LPWSTR lpszExeFileName,
                     [in] UINT nIconIndex);


    [
    
    entry("ExtractIconExA")    ]
    UINT A_ExtractIconEx([in] LPSTR lpszFile, 
							  [in] long nIconIndex,
                              [out] LPVOID phiconLarge,
                              [out] LPVOID phiconSmall,
                              [in] UINT nIcons);

    [
    
    entry("ExtractIconExW")    ]
    UINT W_ExtractIconEx([in] LPWSTR lpszFile, 
							  [in] long nIconIndex,
                              [out] LPVOID phiconLarge,
                              [out] LPVOID phiconSmall,
                              [in] UINT nIcons);

    //
    // SHAddToRecentDocs
    //
	// A pointer to either a null-terminated string with the path and file name of the document, or a PIDL that identifies the document's file object. Set this parameter to NULL to clear all documents from the list. 
    [
    
    entry("SHAddToRecentDocs")    ]
    void A_SHAddToRecentDocsByStr(
			[in] long uFlags,
			[in] LPSTR pv);

    [
    
    entry("SHAddToRecentDocs")    ]
    void W_SHAddToRecentDocsByStr(
			[in] long uFlags,
			[in] LPWSTR pv);


    [
    
    entry("SHAddToRecentDocs")    ]
    void SHAddToRecentDocsByPidl(
			[in] long uFlags,
			[in] DWORD pv);

    //----------
    //
    // SHGetPathFromIDList
    //
    //  This function assumes the size of the buffer (MAX_PATH). The pidl
    // should point to a file system object.
    //
    //----------

    [
    
    entry("SHGetPathFromIDListA")    ]
    BOOL A_SHGetPathFromIDList(
			[in] DWORD pidl,
			[in] LPSTR pszPath);

    [
    
    entry("SHGetPathFromIDListW")    ]
    BOOL W_SHGetPathFromIDList(
			[in] DWORD pidl,
			[in] LPWSTR pszPath);


    // Tray notification definitions

    const DWORD NIM_ADD       = 0x00000000;
    const DWORD NIM_MODIFY    = 0x00000001;
    const DWORD NIM_DELETE    = 0x00000002;

    const DWORD NIF_MESSAGE   = 0x00000001;
    const DWORD NIF_ICON      = 0x00000002;
    const DWORD NIF_TIP       = 0x00000004;

    [
    
    entry("Shell_NotifyIconA")    ]
    BOOL A_Shell_NotifyIcon([in] DWORD dwMessage,
                                 [in] LPVOID  lpData);

    [
    
    entry("Shell_NotifyIconW")    ]
    BOOL W_Shell_NotifyIcon([in] DWORD dwMessage,
                                 [in] LPVOID  lpData);

    [
    
    entry("SHGetFileInfoA")    ]
    DWORD A_SHGetFileInfoByStr([in] LPSTR pszPath,
                               [in] DWORD dwFileAttributes,
                               [out] LPVOID  psfi,
                               [in] UINT cbFileInfo,
                               [in] UINT uFlags);

    [
    
    entry("SHGetFileInfoW")    ]
    DWORD W_SHGetFileInfoByStr([in] LPWSTR pszPath,
                               [in] DWORD dwFileAttributes,
                               [out] LPVOID  psfi,
                               [in] UINT cbFileInfo,
                               [in] UINT uFlags);

    [
    
    entry("SHGetFileInfoA")    ]
    DWORD A_SHGetFileInfoByPidl([in] LONG itemid,
                               [in] DWORD dwFileAttributes,
                               [out] LPVOID  psfi,
                               [in] UINT cbFileInfo,
                               [in] UINT uFlags);

    [
    
    entry("SHGetFileInfoW")    ]
    DWORD W_SHGetFileInfoByPidl([in] LONG itemid,
                               [in] DWORD dwFileAttributes,
                               [out] LPVOID  psfi,
                               [in] UINT cbFileInfo,
                               [in] UINT uFlags);

    [
    entry("SHGetSpecialFolderLocation")    ]
    long SHGetSpecialFolderLocation([in] HWND hwndOwner,
                                              [in] long nFolder,
                                              [out] LPVOID  ppidl); //ITEMIDLIST * ppidl



    //-------------------------------------------------------------------------
    //
    // SHBrowseForFolder API
    //
    //-------------------------------------------------------------------------

    // message from browser
    const short BFFM_INITIALIZED        = 1;
    const short BFFM_SELCHANGED         = 2;
    const short BFFM_VALIDATEFAILEDW    = 4;   // lParam:wzPath ret:1(cont),0(EndDialog)
    const short BFFM_VALIDATEFAILEDA    = 3;   // lParam:szPath ret:1(cont),0(EndDialog)

    // messages to browser
    const short BFFM_ENABLEOK           = 1125;
    const short BFFM_SETSTATUSTEXTA     = 1128;
    const short BFFM_SETSELECTIONA      = 1127;
    const short BFFM_SETSTATUSTEXTW     = 1124;
    const short BFFM_SETSELECTIONW      = 1126;

	const short BIF_RETURNONLYFSDIRS	= 1;

    [
    
    entry("SHBrowseForFolderA")    ]
    long A_SHBrowseForFolder([in,out] LPVOID  lpbi);

    [
    
    entry("SHBrowseForFolderW")    ]
    long W_SHBrowseForFolder([in,out] LPVOID  lpbi);


    //-------------------------------------------------------------------------
    //
    // SHLoadInProc
    //
    //   When this function is called, the shell calls CoCreateInstance
    //  (or equivalent) with CLSCTX_INPROC_SERVER and the specified CLSID
    //  from within the shell's process and release it immediately.
    //
    //-------------------------------------------------------------------------

    [
    entry("CommandLineToArgvW")    ]
	long CommandLineToArgvW(
					[in] LPWSTR lpCmdLine,
					[out] long * pNumArgs);

    [
    entry("SHLoadInProc")    ]
    long SHLoadInProc([in] REFCLSID rclsid); //UUID * rclsid

    //
    //  Helper function which returns a IShellFolder interface to the desktop
    // folder. This is equivalent to call CoCreateInstance with CLSID_ShellDesktop.
    //
    //  CoCreateInstance(CLSID_Desktop, NULL,
    //                   CLSCTX_INPROC, IID_IShellFolder, &pshf);
    //
    [
    entry("SHGetDesktopFolder")    ]
    long SHGetDesktopFolder(
			[in] LPVOID ppshf);


//
//  File System Notification flags
//



	const long SHCNE_RENAMEITEM          = 0x00000001;
	const long SHCNE_CREATE              = 0x00000002;
	const long SHCNE_DELETE              = 0x00000004;
	const long SHCNE_MKDIR               = 0x00000008;
	const long SHCNE_RMDIR               = 0x00000010;
	const long SHCNE_MEDIAINSERTED       = 0x00000020;
	const long SHCNE_MEDIAREMOVED        = 0x00000040;
	const long SHCNE_DRIVEREMOVED        = 0x00000080;
	const long SHCNE_DRIVEADD            = 0x00000100;
	const long SHCNE_NETSHARE            = 0x00000200;
	const long SHCNE_NETUNSHARE          = 0x00000400;
	const long SHCNE_ATTRIBUTES          = 0x00000800;
	const long SHCNE_UPDATEDIR           = 0x00001000;
	const long SHCNE_UPDATEITEM          = 0x00002000;
	const long SHCNE_SERVERDISCONNECT    = 0x00004000;
	const long SHCNE_UPDATEIMAGE         = 0x00008000;
	const long SHCNE_DRIVEADDGUI         = 0x00010000;
	const long SHCNE_RENAMEFOLDER        = 0x00020000;
	const long SHCNE_FREESPACE           = 0x00040000;
	const long SHCNE_EXTENDED_EVENT      = 0x00080000; // Extended Event.

	const long SHCNE_ASSOCCHANGED        = 0x08000000;

	const long SHCNE_DISKEVENTS          = 0x0002381F;
	const long SHCNE_GLOBALEVENTS        = 0x0C0D81E0; // Events that dont match pidls first
	const long SHCNE_ALLEVENTS           = 0x7FFFFFFF;
	const long SHCNE_INTERRUPT           = 0x80000000; // The presence of this flag indicates
                                            // that the event was generated by an
                                            // interrupt.  It is stripped out before
                                            // the clients of SHCNNotify_ see it.

	const long SHCNEE_THEMECHANGED       = 0x00000001;  // The theme changed


// Flags
// uFlags & SHCNF_TYPE is an ID which indicates what dwItem1 and dwItem2 mean
	const long SHCNF_IDLIST      = 0x0000;        // LPITEMIDLIST
	const long SHCNF_PATHA       = 0x0001;        // path name
	const long SHCNF_PRINTERA    = 0x0002;        // printer friendly name
	const long SHCNF_DWORD       = 0x0003;        // DWORD
	const long SHCNF_PATHW       = 0x0005;        // path name
	const long SHCNF_PRINTERW    = 0x0006;        // printer friendly name
	const long SHCNF_TYPE        = 0x00FF;
	const long SHCNF_FLUSH       = 0x1000;
	const long SHCNF_FLUSHNOWAIT = 0x2000;


    [
    
    entry("SHChangeNotify")    ]
    void SHChangeNotify([in] LONG wEventId,[in]  UINT uFlags,
                               [in] LPCVOID dwItem1, [in] LPCVOID dwItem2);


    [
    entry("SHGetInstanceExplorer")    ]
    HRESULT SHGetInstanceExplorer(
			[out] LPVOID ppunk);


//
// SHGetDataFromIDListA/W
//
	const long SHGDFIL_FINDDATA				= 1;
	const long SHGDFIL_NETRESOURCE			= 2;
	const long SHGDFIL_DESCRIPTIONID		= 3;

	const long SHDID_ROOT_REGITEM			= 1;
	const long SHDID_FS_FILE				= 2;
	const long SHDID_FS_DIRECTORY			= 3;
	const long SHDID_FS_OTHER				= 4;
	const long SHDID_COMPUTER_DRIVE35		= 5;
	const long SHDID_COMPUTER_DRIVE525		= 6;
	const long SHDID_COMPUTER_REMOVABLE		= 7;
	const long SHDID_COMPUTER_FIXED			= 8;
	const long SHDID_COMPUTER_NETDRIVE		= 9;
	const long SHDID_COMPUTER_CDROM			= 10;
	const long SHDID_COMPUTER_RAMDISK		= 11;
	const long SHDID_COMPUTER_OTHER			= 12;
	const long SHDID_NET_DOMAIN				= 13;
	const long SHDID_NET_SERVER				= 14;
	const long SHDID_NET_SHARE				= 15;
	const long SHDID_NET_RESTOFNET			= 16;
	const long SHDID_NET_OTHER				= 17;

    //
    // SHGetDataFromIDListA/W
    //
    [
    entry("SHGetDataFromIDListA")    ]
    HRESULT A_SHGetDataFromIDList([in] LPVOID psf,
                                       [in] LPVOID  ppidl,
                                       [in] long nFormat,
                                       [in] LPVOID pv,
                                       [in] long cb);

    [
    entry("SHGetDataFromIDListW")    ]
    HRESULT W_SHGetDataFromIDList([in] LPVOID psf,
                                       [in] LPVOID  ppidl,
                                       [in] long nFormat,
                                       [in] LPVOID pv,
                                       [in] long cb);
// flags for SHEmptyRecycleBin
//
	const long SHERB_NOCONFIRMATION    = 0x00000001;
	const long SHERB_NOPROGRESSUI      = 0x00000002;
	const long SHERB_NOSOUND           = 0x00000004;

    [
    entry("SHEmptyRecycleBinA")    ]
	HRESULT A_SHEmptyRecycleBin(
			[in] HWND hwnd,
			[in] LPSTR pszRootPath,
			[in] DWORD dwFlags);

    [
    entry("SHEmptyRecycleBinW")    ]
	HRESULT W_SHEmptyRecycleBin(
			[in] HWND hwnd,
			[in] LPWSTR pszRootPath,
			[in] DWORD dwFlags);

    [
    entry("SHQueryRecycleBinA")    ]
	HRESULT A_SHQueryRecycleBin(
			[in] LPSTR pszRootPath,
			[in] LPVOID  pSHQueryRBInfo);

    [
    entry("SHQueryRecycleBinW")    ]
	HRESULT W_SHQueryRecycleBin(
			[in] LPWSTR pszRootPath,
			[in] LPVOID  pSHQueryRBInfo);



}; //END of SHELL32