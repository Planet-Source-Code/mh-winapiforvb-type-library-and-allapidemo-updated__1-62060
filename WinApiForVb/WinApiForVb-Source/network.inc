// *********** Network Related
[
	helpstring("Networks"),
	dllname("MPR.DLL")
]
module Network {
    
	// Errors
	const UINT WN_SUCCESS           = 0;
	const UINT WN_NO_ERROR			= 0;
	const UINT WN_NOT_SUPPORTED     = 50;
	const UINT WN_CANCEL            = 1223;
	const UINT WN_CONTINUE          = 1237;
	const UINT WN_NET_ERROR         = 59;
	const UINT WN_MORE_DATA         = 234;
	const UINT WN_BAD_POINTER       = 487;
	const UINT WN_BAD_VALUE         = 87;
	const UINT WN_BAD_PASSWORD      = 86;
	const UINT WN_BAD_USER          = 2202;
	const UINT WN_ACCESS_DENIED     = 5;
	const UINT WN_FUNCTION_BUSY     = 142;
	const UINT WN_WINDOWS_ERROR     = 59;
	const UINT WN_OUT_OF_MEMORY     = 8;
    const long WN_NO_NETWORK        = 1222;
    const long WN_EXTENDED_ERROR	= 1208;
    const long WN_BAD_LEVEL         = 124;
    const long WN_BAD_HANDLE        = 6;
    const long WN_NOT_INITIALIZING	= 1247;
    const long WN_NO_MORE_DEVICES	= 1248;

	// Connection errors
	const UINT WN_NOT_CONNECTED     = 2250;
	const UINT WN_OPEN_FILES        = 2401;
    const long WN_DEVICE_IN_USE     = 2404;
	const UINT WN_BAD_NETNAME       = 67;
	const UINT WN_BAD_LOCALNAME     = 1200;
	const UINT WN_ALREADY_CONNECTED = 85;
	const UINT WN_DEVICE_ERROR      = 31;
	const UINT WN_CONNECTION_CLOSED = 1201;
    const long WN_NO_NET_OR_BAD_PATH    = 1203;
    const long WN_CANNOT_OPEN_PROFILE	= 1205;
	const long WN_BAD_PROVIDER		= 1204;
    const long WN_BAD_PROFILE		= 1206;
    const long WN_BAD_DEV_TYPE		= 66;
    const long WN_DEVICE_ALREADY_REMEMBERED  = 1202;

	// Enumeration
    const long WN_NO_MORE_ENTRIES	= 259;
    const long WN_NOT_CONTAINER		= 1207;

	// Authentication
    const long WN_NOT_AUTHENTICATED	= 1244;
    const long WN_NOT_LOGGED_ON		= 1245;
    const long WN_NOT_VALIDATED		= 1311;


    // ****** Network support *******

    [
    
    entry("WNetAddConnectionA")    ]
    UINT A_WNetAddConnection([in] LPSTR lpszNetPath,
                                  [in] LPSTR lpszPassword,
                                  [in] LPSTR lpszLocalName);

    [
    
    entry("WNetAddConnectionW")    ]
    UINT W_WNetAddConnection([in] LPWSTR lpszNetPath,
                                  [in] LPWSTR lpszPassword,
                                  [in] LPWSTR lpszLocalName);


    [
    
    entry("WNetGetUserA")    ]
	DWORD  A_WNetGetUser([in] LPSTR lpName, [out] LPSTR lpUserName, 
						     [in,out] DWORD *lpnLength);

    [
    
    entry("WNetGetUserW")    ]
	DWORD  W_WNetGetUser([in] LPWSTR lpName, [out] LPWSTR lpUserName, 
						     [in,out] DWORD *lpnLength);

    [
    
    entry("WNetGetConnectionA")    ]
    UINT  A_WNetGetConnection([in] LPSTR lpszLocalName,
                                  [in] LPSTR lpszRemoteName,
                                  [in,out] UINT *cbRemoteName);
    [
    
    entry("WNetGetConnectionW")    ]
    UINT  W_WNetGetConnection([in] LPWSTR lpszLocalName,
                                  [in] LPWSTR lpszRemoteName,
                                  [in,out] UINT *cbRemoteName);

    [
    
    entry("WNetCancelConnectionA")    ]
    UINT  A_WNetCancelConnection([in] LPSTR lpszName,
                                     [in] BOOL fForce);

    [
    
    entry("WNetCancelConnectionW")    ]
    UINT  W_WNetCancelConnection([in] LPWSTR lpszName,
                                     [in] BOOL fForce);

    [
    
    entry("WNetConnectionDialog"),
    ]
    UINT  WNetConnectionDialog([in] UINT hWnd,
                                     [in] UINT fdwResourceType);


    [
    
    entry("WNetDisconnectDialog"),
    ]
    DWORD  WNetDisconnectDialog([in] HWND  hwnd,
                                      [in] DWORD dwType);

    const long RESOURCETYPE_ANY         = 0x00000000;
    const long RESOURCETYPE_DISK        = 0x00000001;
    const long RESOURCETYPE_PRINT       = 0x00000002;
    const long RESOURCETYPE_UNKNOWN     = 0xFFFFFFFF;
    const long RESOURCETYPE_RESERVED    = 0x00000008;

    const DWORD RESOURCE_CONNECTED     = 0x00000001;
    const DWORD RESOURCE_GLOBALNET     = 0x00000002;
    const DWORD RESOURCE_REMEMBERED    = 0x00000003;
    const DWORD RESOURCE_RECENT        = 0x00000004;
    const DWORD RESOURCE_CONTEXT       = 0x00000005;

    const DWORD RESOURCEUSAGE_CONNECTABLE  = 0x00000001;
    const DWORD RESOURCEUSAGE_CONTAINER    = 0x00000002;
    const DWORD RESOURCEUSAGE_NOLOCALDEVICE = 0x00000004;
    const DWORD RESOURCEUSAGE_SIBLING      = 0x00000008;
    const DWORD RESOURCEUSAGE_ALL          = 0x00000003; // RESOURCEUSAGE_CONNECTABLE | RESOURCEUSAGE_CONTAINER
    const DWORD RESOURCEUSAGE_RESERVED     = 0x80000000;

    const DWORD RESOURCEDISPLAYTYPE_GENERIC       = 0x00000000;
    const DWORD RESOURCEDISPLAYTYPE_DOMAIN        = 0x00000001;
    const DWORD RESOURCEDISPLAYTYPE_SERVER        = 0x00000002;
    const DWORD RESOURCEDISPLAYTYPE_SHARE         = 0x00000003;
    const DWORD RESOURCEDISPLAYTYPE_FILE          = 0x00000004;
    const DWORD RESOURCEDISPLAYTYPE_GROUP         = 0x00000005;
    const DWORD RESOURCEDISPLAYTYPE_NETWORK       = 0x00000006;
    const DWORD RESOURCEDISPLAYTYPE_ROOT          = 0x00000007;
    const DWORD RESOURCEDISPLAYTYPE_SHAREADMIN    = 0x00000008;
    const DWORD RESOURCEDISPLAYTYPE_DIRECTORY     = 0x00000009;
    const DWORD RESOURCEDISPLAYTYPE_TREE          = 0x0000000A;

    //  Network Connections

    const long NETPROPERTY_PERSISTENT      = 1;

    const DWORD CONNECT_UPDATE_PROFILE     = 0x00000001;
    const DWORD CONNECT_UPDATE_RECENT      = 0x00000002;
    const DWORD CONNECT_TEMPORARY          = 0x00000004;
    const DWORD CONNECT_INTERACTIVE        = 0x00000008;
    const DWORD CONNECT_PROMPT             = 0x00000010;
    const DWORD CONNECT_NEED_DRIVE         = 0x00000020;
    const DWORD CONNECT_REFCOUNT           = 0x00000040;
    const DWORD CONNECT_REDIRECT           = 0x00000080;
    const DWORD CONNECT_LOCALDRIVE         = 0x00000100;
    const DWORD CONNECT_CURRENT_MEDIA      = 0x00000200;

    [
    
    entry("WNetAddConnection2A")    ]
    DWORD  A_WNetAddConnection2([in] LPVOID lpNetResource,
                                    [in] LPSTR lpPassword,
                                    [in] LPSTR lpUserName,
                                    [in] DWORD dwFlags);

    [
    
    entry("WNetAddConnection2W")    ]
    DWORD  W_WNetAddConnection2([in] LPVOID lpNetResource,
                                    [in] LPWSTR lpPassword,
                                    [in] LPWSTR lpUserName,
                                    [in] DWORD dwFlags);

    [
    
    entry("WNetAddConnection3A")    ]
    DWORD  A_WNetAddConnection3([in] HWND hwndOwner,
                                    [in] LPVOID lpNetResource,
                                    [in] LPSTR lpPassword,
                                    [in] LPSTR lpUserName,
                                    [in] DWORD dwFlags);

    [
    
    entry("WNetAddConnection3W")    ]
    DWORD  W_WNetAddConnection3([in] HWND hwndOwner,
                                    [in] LPVOID lpNetResource,
                                    [in] LPWSTR lpPassword,
                                    [in] LPWSTR lpUserName,
                                    [in] DWORD dwFlags);

    [
    
    entry("WNetCancelConnection2A")    ]
    DWORD  A_WNetCancelConnection2([in] LPSTR lpName,
                                       [in] DWORD dwFlags,
                                       [in] BOOL fForce);

    [
    
    entry("WNetCancelConnection2W")    ]
    DWORD  W_WNetCancelConnection2([in] LPWSTR lpName,
                                       [in] DWORD dwFlags,
                                       [in] BOOL fForce);


    [
    
    entry("WNetUseConnectionA")    ]
    DWORD  A_WNetUseConnection([in] HWND hwndOwner,
                                   [in] LPVOID lpNetResource,
                                   [in] LPSTR lpUserID,
                                   [in] LPSTR lpPassword,
                                   [in] DWORD dwFlags,
                                   [out] LPSTR lpAccessName,
                                   [in,out] DWORD * lpBufferSize,
                                   [out] DWORD * lpResult);

    [
    
    entry("WNetUseConnectionW")    ]
    DWORD  W_WNetUseConnection([in] HWND hwndOwner,
                                   [in] LPVOID lpNetResource,
                                   [in] LPWSTR lpUserID,
                                   [in] LPWSTR lpPassword,
                                   [in] DWORD dwFlags,
                                   [out] LPWSTR lpAccessName,
                                   [in,out] DWORD * lpBufferSize,
                                   [out] DWORD * lpResult);

    //  Network Connection Dialogs.

    const DWORD CONNDLG_RO_PATH    = 0x00000001; // Resource path should be read-only
    const DWORD CONNDLG_CONN_POINT = 0x00000002; // Netware -style movable connection point enabled
    const DWORD CONNDLG_USE_MRU    = 0x00000004; // Use MRU combobox
    const DWORD CONNDLG_HIDE_BOX   = 0x00000008; // Hide persistent connect checkbox

    const DWORD CONNDLG_PERSIST    = 0x00000010;  // Force persistent connection
    const DWORD CONNDLG_NOT_PERSIST = 0x00000020; // Force connection NOT persistent

    [
    
    entry("WNetConnectionDialog1A")    ]
    DWORD A_WNetConnectionDialog1(
			[in,out] LPVOID lpConnDlgStruct);

    [
    
    entry("WNetConnectionDialog1W")    ]
    DWORD W_WNetConnectionDialog1(
			[in,out] LPVOID lpConnDlgStruct);

    const DWORD DISC_UPDATE_PROFILE        = 0x00000001;
    const DWORD DISC_NO_FORCE              = 0x00000040;

    [
    
    entry("WNetDisconnectDialog1")    ]
    DWORD  WNetDisconnectDialog1(
				[in] LPVOID lpConnDlgStruct);

    //  Network Browsing.

    [
    
    entry("WNetOpenEnumA")    ]
    DWORD  A_WNetOpenEnum(
         [in] DWORD dwScope,
         [in] DWORD dwType,
         [in] DWORD dwUsage,
         [in] LPVOID lpNetResource,
         [out] HANDLE * lphEnum);

    [
    
    entry("WNetOpenEnumW")    ]
    DWORD  W_WNetOpenEnum(
         [in] DWORD dwScope,
         [in] DWORD dwType,
         [in] DWORD dwUsage,
         [in] LPVOID lpNetResource,
         [out] HANDLE * lphEnum);

   [
   
   entry("WNetEnumResourceA")   ]
   DWORD  A_WNetEnumResource(
        [in] HANDLE  hEnum,
        [in,out] DWORD * lpcCount,
        [out] void * lpBuffer,
        [in,out] DWORD * lpBufferSize);

   [
   
   entry("WNetEnumResourceW")   ]
   DWORD  W_WNetEnumResource(
        [in] HANDLE  hEnum,
        [in,out] DWORD * lpcCount,
        [out] void * lpBuffer,
        [in,out] DWORD * lpBufferSize);

   [
   
   entry("WNetCloseEnum")   ]
   DWORD  WNetCloseEnum(
		[in] HANDLE  hEnum);

   //  Universal Naming.

   const DWORD UNIVERSAL_NAME_INFO_LEVEL  = 0x00000001;
   const DWORD REMOTE_NAME_INFO_LEVEL     = 0x00000002;
    
   [
   
   entry("WNetGetUniversalNameA")   ]
   DWORD  A_WNetGetUniversalName(
        [in] LPSTR lpLocalPath,
        [in] DWORD dwInfoLevel,
        [out] LPVOID lpBuffer,
        [in,out] DWORD *lpBufferSize);

   [
   
   entry("WNetGetUniversalNameW")   ]
   DWORD  W_WNetGetUniversalName(
        [in] LPWSTR lpLocalPath,
        [in] DWORD dwInfoLevel,
        [out] LPVOID lpBuffer,
        [in,out] DWORD *lpBufferSize);

   // Other

   const WORD WNFMT_MULTILINE        = 0x01;
   const WORD WNFMT_ABBREVIATED      = 0x02;
   const WORD WNFMT_INENUM           = 0x10;
   const WORD WNFMT_CONNECTION       = 0x20;

    [
    
    entry("WNetGetProviderNameA")    ]
    DWORD  A_WNetGetProviderName(
       [in] DWORD   dwNetType,
       [out] LPSTR lpProviderName,
       [in,out] DWORD * lpBufferSize);

    [
    
    entry("WNetGetProviderNameW")    ]
    DWORD  W_WNetGetProviderName(
       [in] DWORD   dwNetType,
       [out] LPWSTR lpProviderName,
       [in,out] DWORD * lpBufferSize);


    const DWORD NETINFO_DLL16      = 0x00000001;  // Provider running as 16 bit Winnet Driver
    const DWORD NETINFO_DISKRED    = 0x00000004;  // Provider requires disk redirections to connect
    const DWORD NETINFO_PRINTERRED = 0x00000008;  // Provider requires printer redirections to connect

    [
    
    entry("WNetGetNetworkInformationA")    ]
    DWORD  A_WNetGetNetworkInformation(
       [in] LPSTR lpProvider,
       [out] LPVOID lpNetInfoStruct);

    [
    
    entry("WNetGetNetworkInformationW")    ]
    DWORD  W_WNetGetNetworkInformation(
       [in] LPWSTR lpProvider,
       [out] LPVOID lpNetInfoStruct);


   //  User Profiles

   const WORD RP_LOGON   = 0x01;  // if set, do for logon, else for logoff
   const WORD RP_INIFILE = 0x02;  // if set, reconcile .INI file, else reg. hive

   //  Policies.

   const WORD PP_DISPLAYERRORS   = 0x01;    // if set, display error messages, else fail silently if error


    //  Error handling.

    [
    
    entry("WNetGetLastErrorA")    ]
    DWORD  A_WNetGetLastError(
        [out] DWORD * lpError,
        [out] LPSTR lpErrorBuf,
        [in] DWORD nErrorBufSize,
        [out] LPSTR lpNameBuf,
        [in] DWORD nNameBufSize);

    [
    
    entry("WNetGetLastErrorW")    ]
    DWORD  W_WNetGetLastError(
        [out] DWORD * lpError,
        [out] LPWSTR lpErrorBuf,
        [in] DWORD nErrorBufSize,
        [out] LPWSTR lpNameBuf,
        [in] DWORD nNameBufSize);

    //  For Shell

    const DWORD WNCON_FORNETCARD       = 0x00000001;
    const DWORD WNCON_NOTROUTED        = 0x00000002;
    const DWORD WNCON_SLOWLINK         = 0x00000004;
    const DWORD WNCON_DYNAMIC          = 0x00000008;

    [
    
    entry("MultinetGetConnectionPerformanceA")    ]
    DWORD  A_MultinetGetConnectionPerformance(
        [in] LPVOID lpNetResource,
        [out] LPVOID lpNetConnectInfoStruct);

    [
    
    entry("MultinetGetConnectionPerformanceW")    ]
    DWORD  W_MultinetGetConnectionPerformance(
        [in] LPVOID lpNetResource,
        [out] LPVOID lpNetConnectInfoStruct);

    [
    
    entry("WNetGetResourceParentA")    ]	
	DWORD A_WNetGetResourceParent(
		[in] LPVOID lpNetResource,
		[out] LPVOID lpBuffer,
		[in,out] DWORD * lpcbBuffer);

    [
    
    entry("WNetGetResourceParentW")    ]	
	DWORD W_WNetGetResourceParent(
		[in] LPVOID lpNetResource,
		[out] LPVOID lpBuffer,
		[in,out] DWORD * lpcbBuffer);

    [
    
    entry("WNetGetResourceInformationA")    ]
	DWORD A_WNetGetResourceInformation(
		[in] LPVOID lpNetResource,
		[out] LPVOID lpBuffer,
		[in,out] DWORD * lpcbBuffer,
		[out] long * lplpSystem); //LPSTR

    [
    
    entry("WNetGetResourceInformationW")    ]
	DWORD W_WNetGetResourceInformation(
		[in] LPVOID lpNetResource,
		[out] LPVOID lpBuffer,
		[in,out] DWORD * lpcbBuffer,
		[out] long * lplpSystem); // LPWSTR


}; //End Network
