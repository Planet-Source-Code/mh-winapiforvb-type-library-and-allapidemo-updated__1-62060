

typedef enum HRESULTS {
    S_OK = 0,
    S_FALSE = 1,

// Codes 0x4000-0x40ff are reserved for OLE
	E_UNEXPECTED = 0x8000FFFF,

    E_NOTIMPL = 0x80004001,
    E_OUTOFMEMORY = 0x8007000E,
    E_INVALIDARG = 0x80070057,
    E_NOINTERFACE = 0x80004002,
    E_POINTER = 0x80004003,
    E_HANDLE = 0x80070006,
    E_ABORT = 0x80004004,
    E_FAIL = 0x80004005,
    E_ACCESSDENIED = 0x80070005,

    E_NOTIMPL_SEC = 0x80000001,
    E_OUTOFMEMORY_SEC = 0x80000002,
    E_INVALIDARG_SEC = 0x80000003,
    E_NOINTERFACE_SEC = 0x80000004,
    E_POINTER_SEC = 0x80000005,
    E_HANDLE_SEC = 0x80000006,
    E_ABORT_SEC = 0x80000007,
    E_FAIL_SEC = 0x80000008,
    E_ACCESSDENIED_SEC = 0x80000009,
	
	E_PENDING = 0x8000000A,
    // IConnectionPoint errors
        
    CONNECT_E_FIRST = 0x80040200,
    CONNECT_E_NOCONNECTION,  // there is no connection for this connection id
    CONNECT_E_ADVISELIMIT,   // this implementation's limit for advisory connections has been reached
    CONNECT_E_CANNOTCONNECT, // connection attempt failed
    CONNECT_E_OVERRIDDEN,    // must use a derived interface to connect
    

    // INET errors
    
    INET_E_INVALID_URL = 0x800C0002,
    INET_E_NO_SESSION = 0x800C0003,
    INET_E_CANNOT_CONNECT = 0x800C0004,
    INET_E_RESOURCE_NOT_FOUND = 0x800C0005,
    INET_E_OBJECT_NOT_FOUND = 0x800C0006,
    INET_E_DATA_NOT_AVAILABLE = 0x800C0007,
    INET_E_DOWNLOAD_FAILURE = 0x800C0008,
    INET_E_AUTHENTICATION_REQUIRED = 0x800C0009,
    INET_E_NO_VALID_MEDIA = 0x800C000A,
    INET_E_CONNECTION_TIMEOUT = 0x800C000B,
    INET_E_INVALID_REQUEST = 0x800C000C,
    INET_E_UNKNOWN_PROTOCOL = 0x800C000D,
    INET_E_SECURITY_PROBLEM = 0x800C000E,
    INET_E_CANNOT_LOAD_DATA = 0x800C000F,
    INET_E_CANNOT_INSTANTIATE_OBJECT = 0x800C0010,
    INET_E_USE_DEFAULT_PROTOCOLHANDLER = 0x800C0011,
    INET_E_DEFAULT_ACTION = 0x800C0011,
    INET_E_USE_DEFAULT_SETTING = 0x800C0012,
    INET_E_QUERYOPTION_UNKNOWN = 0x800C0013,
    INET_E_REDIRECTING = 0x800C0014,
    INET_E_REDIRECT_FAILED = 0x800C0014,
    INET_E_REDIRECT_TO_DIR = 0x800C0015,
    INET_E_CANNOT_LOCK_REQUEST = 0x800C0016,
    INET_E_USE_EXTEND_BINDING = 0x800C0017,
    INET_E_ERROR_FIRST = 0x800C0002,
    INET_E_ERROR_LAST = 0x800C0017,
    INET_E_CODE_DOWNLOAD_DECLINED = 0x800C0100,
    INET_E_RESULT_DISPATCHED = 0x800C0200,
    INET_E_CANNOT_REPLACE_SFP_FILE = 0x800C0300,

} HRESULTS;

typedef enum ROOTKEY {
    [ helpstring("Registry keys: COM information") ]
    HKEY_CLASSES_ROOT            = 0x80000000,
    [ helpstring("Registry keys: User-specific information") ]
    HKEY_CURRENT_USER            = 0x80000001,
    [ helpstring("Registry keys: Machine-specific information") ]
    HKEY_LOCAL_MACHINE           = 0x80000002,
    [ helpstring("Registry keys: User-specific information for all users") ]
    HKEY_USERS                   = 0x80000003,
    [ helpstring("Registry keys: Performance information") ]
    HKEY_PERFORMANCE_DATA        = 0x80000004,
    [ helpstring("Registry keys: Performance information") ]
    HKEY_CURRENT_CONFIG          = 0x80000005,
    [ helpstring("Registry keys: Configuration information") ]
    HKEY_DYN_DATA                = 0x80000006
} ROOTKEY;

typedef enum REGTYPE {
    [ helpstring("Registry types: No value type") ]
    REG_NONE                     = 0,
    [ helpstring("Registry types: Null-terminated string (Unicode in 32-bit)") ]
    REG_SZ                       = 1,
    [ helpstring("Registry types: Null-terminated string with environment variable references") ]
    REG_EXPAND_SZ                = 2,
    [ helpstring("Registry types: Free form binary") ]
    REG_BINARY                   = 3,
    [ helpstring("Registry types: 32-bit number") ]
    REG_DWORD                    = 4,
    [ helpstring("Registry types: 32-bit number") ]
    REG_DWORD_LITTLE_ENDIAN      = 4,
    [ helpstring("Registry types: 32-bit number") ]
    REG_DWORD_BIG_ENDIAN         = 5,
    [ helpstring("Registry types: Symbolic Link (Unicode)") ]
    REG_LINK                     = 6,
    [ helpstring("Registry types: Multiple Unicode strings") ]
    REG_MULTI_SZ                 = 7,
    [ helpstring("Registry types: Resource list in the resource map") ]
    REG_RESOURCE_LIST            = 8,
    [ helpstring("Registry types: Resource list in the hardware description") ]
    REG_FULL_RESOURCE_DESCRIPTOR = 9,
    REG_RESOURCE_REQUIREMENTS_LIST   = 10
} REGTYPE;

typedef enum REGACCESS {

    [ helpstring("Registry access: Allow value query") ]
    KEY_QUERY_VALUE         = 0x0001,
    [ helpstring("Registry access: Allow set subkey data") ]
    KEY_SET_VALUE           = 0x0002,
    [ helpstring("Registry access: Allow create subkey") ]
    KEY_CREATE_SUB_KEY      = 0x0004,
    [ helpstring("Registry access: Allow subkey enumeration") ]
    KEY_ENUMERATE_SUB_KEYS  = 0x0008,
    [ helpstring("Registry access: Allow change notification") ]
    KEY_NOTIFY              = 0x0010,
    [ helpstring("Registry access: Allow symbolic link creation") ]
    KEY_CREATE_LINK         = 0x0020,
    [ helpstring("Registry access: Allow query, enumerate, and notify") ]
    KEY_READ                = 0x00020009,
    [ helpstring("Registry access: Allow subkey create and set value") ]
    KEY_WRITE               = 0x00020006,
    [ helpstring("Registry access: Allow read") ]
    KEY_EXECUTE             = 0x00020009,
    [ helpstring("Registry access: Allow anything") ]
    KEY_ALL_ACCESS          = 0x000F003F

} REGACCESS;


typedef enum FILE_NOTIFY {
    FILE_NOTIFY_CHANGE_FILE_NAME  = 0x00000001,
    FILE_NOTIFY_CHANGE_DIR_NAME   = 0x00000002,
    FILE_NOTIFY_CHANGE_ATTRIBUTES = 0x00000004,
    FILE_NOTIFY_CHANGE_SIZE       = 0x00000008,
    FILE_NOTIFY_CHANGE_LAST_WRITE = 0x00000010,
    FILE_NOTIFY_CHANGE_SECURITY   = 0x00000100
} FILE_NOTIFY;

// Rights for OpenProcess
typedef enum PROCESSACCESS {
    PROCESS_TERMINATE        = 0x0001,
    PROCESS_CREATE_THREAD    = 0x0002,
    PROCESS_VM_OPERATION     = 0x0008,
    PROCESS_VM_READ          = 0x0010,
    PROCESS_VM_WRITE         = 0x0020,
    PROCESS_DUP_HANDLE       = 0x0040,
    PROCESS_CREATE_PROCESS   = 0x0080,
    PROCESS_SET_QUOTA        = 0x0100,
    PROCESS_SET_INFORMATION  = 0x0200,
    PROCESS_QUERY_INFORMATION = 0x0400,
    PROCESS_ALL_ACCESS       = 0x001F0FFF
} PROCESSACCESS;

typedef enum FORMAT_MESSAGE {
    FORMAT_MESSAGE_NONE             = 0x00000000,
    FORMAT_MESSAGE_ALLOCATE_BUFFER  = 0x00000100,
    FORMAT_MESSAGE_IGNORE_INSERTS   = 0x00000200,
    FORMAT_MESSAGE_FROM_STRING      = 0x00000400,
    FORMAT_MESSAGE_FROM_HMODULE     = 0x00000800,
    FORMAT_MESSAGE_FROM_SYSTEM      = 0x00001000,
    FORMAT_MESSAGE_ARGUMENT_ARRAY   = 0x00002000,
    FORMAT_MESSAGE_MAX_WIDTH_MASK   = 0x000000FF
} FORMAT_MESSAGE;

// Global Memory Flags
typedef enum GMEM {
     GMEM_FIXED          = 0x0000,
     GMEM_MOVEABLE       = 0x0002,
     GMEM_NOCOMPACT      = 0x0010,
     GMEM_NODISCARD      = 0x0020,
     GMEM_ZEROINIT       = 0x0040,
     GMEM_MODIFY         = 0x0080,
     GMEM_DISCARDABLE    = 0x0100,
     GMEM_NOT_BANKED     = 0x1000,
     GMEM_SHARE          = 0x2000,
     GMEM_DDESHARE       = 0x2000,
     GMEM_NOTIFY         = 0x4000,
     GMEM_LOWER          = 0x1000,
     GMEM_VALID_FLAGS    = 0x7F72,
     GMEM_INVALID_HANDLE = 0x8000,
     GHND = 0x42,
     GPTR = 0x40,

     // Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE)
     GMEM_DISCARDED      = 0x4000,
     GMEM_LOCKCOUNT      = 0x00FF,
} GMEM;

typedef enum GET_FILEEX_INFO_LEVELS {
    GetFileExInfoStandard,
    GetFileExMaxInfoLevel
} GET_FILEEX_INFO_LEVELS;

typedef [helpstring("IShellFolder.GetDisplayNameOf/SetNameOf uFlags")]
enum SHGNO_Flags {
    SHGDN_NORMAL = 0,
    SHGDN_INFOLDER = 1,
    SHGDN_FORADDRESSBAR = 0x4000,
    SHGDN_FORPARSING = 0x8000,
} SHGNO_Flags;

typedef [helpstring("IShellFolder.EnumObjects Flags")]
enum SHCONTF {
    SHCONTF_FOLDERS             = 0x0020,
    SHCONTF_NONFOLDERS          = 0x0040,
    SHCONTF_INCLUDEHIDDEN       = 0x0080,
    SHCONTF_INIT_ON_FIRST_NEXT  = 0x0100,
    SHCONTF_NETPRINTERSRCH      = 0x0200,
    SHCONTF_SHAREABLE           = 0x0400,
    SHCONTF_STORAGE             = 0x0800,
} SHCONTF;

typedef [helpstring("IShellFolder.GetAttributesOf flags")]
enum SFGAO_Flags {
    SFGAO_CANCOPY                    = 0x00000001,
    SFGAO_CANMOVE                    = 0x00000002,
    SFGAO_CANLINK                    = 0x00000004,
    SFGAO_CANRENAME                  = 0x00000010,
    SFGAO_CANDELETE                  = 0x00000020,
    SFGAO_HASPROPSHEET               = 0x00000040,
    SFGAO_DROPTARGET                 = 0x00000100,
    SFGAO_CAPABILITYMASK             = 0x00000177,

	SFGAO_ENCRYPTED                  = 0x00002000,
	SFGAO_ISSLOW                     = 0x00004000,
    SFGAO_GHOSTED                    = 0x00008000,
    SFGAO_LINK                       = 0x00010000,
    SFGAO_SHARE                      = 0x00020000,
    SFGAO_READONLY                   = 0x00040000,
    SFGAO_HIDDEN                     = 0x00080000,
    SFGAO_DISPLAYATTRMASK            = 0x000FC000,

    SFGAO_FILESYSANCESTOR            = 0x10000000,
    SFGAO_FOLDER                     = 0x20000000,
    SFGAO_FILESYSTEM                = 0x40000000,
    SFGAO_HASSUBFOLDER              = 0x80000000,
    SFGAO_CONTENTSMASK              = 0x80000000,

    SFGAO_VALIDATE                  = 0x01000000,
    SFGAO_REMOVABLE                 = 0x02000000,
    SFGAO_COMPRESSED                = 0x04000000,
    SFGAO_BROWSABLE                 = 0x08000000,
    SFGAO_NONENUMERATED             = 0x00100000,
    SFGAO_NEWCONTENT                = 0x00200000,

    SFGAO_CANMONIKER                = 0x00400000,
    SFGAO_HASSTORAGE                = 0x00400000,
    SFGAO_STREAM                    = 0x00400000,     
    SFGAO_STORAGEANCESTOR           = 0x00800000,
    SFGAO_STORAGECAPMASK            = 0x70C50008,
} SFGAO_Flags;

typedef enum DATADIR {
    DATADIR_GET = 1,
    DATADIR_SET = 2
} DATADIR;

typedef enum USERCLASSTYPE {
    USERCLASSTYPE_FULL = 1,
    USERCLASSTYPE_SHORT = 2,
    USERCLASSTYPE_APPNAME = 3
} USERCLASSTYPE;

typedef enum DVASPECT {
    DVASPECT_CONTENT = 1,
    DVASPECT_THUMBNAIL = 2,
    DVASPECT_ICON = 4,
    DVASPECT_DOCPRINT = 8,
    DVASPECT_OPAQUE = 16,
    DVASPECT_TRANSPARENT = 32
} DVASPECT;

typedef enum TYMED {
    TYMED_HGLOBAL = 1,
    TYMED_FILE = 2,
    TYMED_ISTREAM = 4,
    TYMED_ISTORAGE = 8,
    TYMED_GDI = 16,
    TYMED_MFPICT = 32,
    TYMED_ENHMF = 64,
    TYMED_NULL = 0
} TYMED;

typedef enum OLEMISC {
    OLEMISC_RECOMPOSEONRESIZE = 1,
    OLEMISC_ONLYICONIC = 2,
    OLEMISC_INSERTNOTREPLACE = 4,
    OLEMISC_STATIC = 8,
    OLEMISC_CANTLINKINSIDE = 16,
    OLEMISC_CANLINKBYOLE1 = 32,
    OLEMISC_ISLINKOBJECT = 64,
    OLEMISC_INSIDEOUT = 128,
    OLEMISC_ACTIVATEWHENVISIBLE = 256,
    OLEMISC_RENDERINGISDEVICEINDEPENDENT = 512,
    OLEMISC_INVISIBLEATRUNTIME = 1024,
    OLEMISC_ALWAYSRUN = 2048,
    OLEMISC_ACTSLIKEBUTTON = 4096,
    OLEMISC_ACTSLIKELABEL = 8192,
    OLEMISC_NOUIACTIVATE = 16384,
    OLEMISC_ALIGNABLE = 32768,
    OLEMISC_SIMPLEFRAME = 0x00010000,
    OLEMISC_SETCLIENTSITEFIRST = 0x00020000,
    OLEMISC_IMEMODE = 0x00040000,
    OLEMISC_IGNOREACTIVATEWHENVISIBLE = 0x00080000,
    OLEMISC_WANTSTOMENUMERGE = 0x00100000,
    OLEMISC_SUPPORTSMULTILEVELUNDO = 0x00200000
} OLEMISC;

typedef enum DVASPECTINFOFLAG {
    DVASPECTINFOFLAG_CANOPTIMIZE = 1
} DVASPECTINFOFLAG;

typedef enum ADVF {
    ADVF_NODATA = 1,
    ADVF_PRIMEFIRST = 2,
    ADVF_ONLYONCE = 4,
    ADVF_DATAONSTOP = 64,
    ADVFCACHE_NOHANDLER = 8,
    ADVFCACHE_FORCEBUILTIN = 16,
    ADVFCACHE_ONSAVE = 32
} ADVF;

typedef enum ESTRRET {
    STRRET_WSTR = 0,
    STRRET_OFFSET = 1,
    STRRET_CSTR = 2
} ESTRRET;

typedef enum STGMOVE {
    STGMOVE_MOVE = 0,
    STGMOVE_COPY = 1,
    STGMOVE_SHALLOWCOPY = 2
} STGMOVE;

    typedef enum NMREBARMasks {
        RBNM_ID = 0x00000001,
        RBNM_STYLE = 0x00000002,
        RBNM_LPARAM = 0x00000004,
    } NMREBARMasks;

typedef enum OLEVERBATTRIB {
    OLEVERBATTRIB_NEVERDIRTIES = 1,
    OLEVERBATTRIB_ONCONTAINERMENU = 2
} OLEVERBATTRIB;

typedef enum OLECONTF {
    OLECONTF_EMBEDDINGS = 1,
    OLECONTF_LINKS = 2,
    OLECONTF_OTHERS = 4,
    OLECONTF_ONLYUSER = 8,
    OLECONTF_ONLYIFRUNNING = 16
} OLECONTF;

typedef enum OLEGETMONIKER {
    OLEGETMONIKER_ONLYIFTHERE = 1,
    OLEGETMONIKER_FORCEASSIGN = 2,
    OLEGETMONIKER_UNASSIGN = 3,
    OLEGETMONIKER_TEMPFORUSER = 4
} OLEGETMONIKER;

typedef enum OLEWHICHMK {
    OLEWHICHMK_CONTAINER = 1,
    OLEWHICHMK_OBJREL = 2,
    OLEWHICHMK_OBJFULL = 3
} OLEWHICHMK;

typedef enum OLECLOSE {
    OLECLOSE_SAVEIFDIRTY = 0,
    OLECLOSE_NOSAVE = 1,
    OLECLOSE_PROMPTSAVE = 2
} OLECLOSE;

typedef enum STGM {
    STGM_DIRECT = 0x00000000,
    STGM_TRANSACTED = 0x00010000,
    STGM_SIMPLE = 0x08000000,

    STGM_READ = 0x00000000,
    STGM_WRITE = 0x00000001,
    STGM_READWRITE = 0x00000002,

    STGM_SHARE_DENY_NONE = 0x00000040,
    STGM_SHARE_DENY_READ = 0x00000030,
    STGM_SHARE_DENY_WRITE = 0x00000020,
    STGM_SHARE_EXCLUSIVE = 0x00000010,

    STGM_PRIORITY = 0x00040000,
    STGM_DELETEONRELEASE = 0x04000000,
    STGM_NOSCRATCH = 0x00100000,

    STGM_CREATE = 0x00001000,
    STGM_CONVERT = 0x00020000,
    STGM_FAILIFTHERE = 0x00000000,

    STGM_NOSNAPSHOT = 0x00200000,
    STGM_DIRECT_SWMR = 0x00400000,
} STGM;

typedef enum STGTY {
    STGTY_STORAGE = 1,
    STGTY_STREAM = 2,
    STGTY_LOCKBYTES = 3,
    STGTY_PROPERTY = 4
} STGTY;

typedef enum LOCKTYPE {
    LOCK_WRITE = 1,
    LOCK_EXCLUSIVE = 2,
    LOCK_ONLYONCE = 4
} LOCKTYPE;

typedef enum STGC {
     STGC_DEFAULT = 0,
     STGC_OVERWRITE = 1,
     STGC_ONLYIFCURRENT = 2,
     STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
     STGC_CONSOLIDATE = 8
} STGC;

typedef enum STATFLAG {
    STATFLAG_DEFAULT = 0,
    STATFLAG_NONAME = 1,
    STATFLAG_NOOPEN = 2
} STATFLAG;

    typedef enum BIND_FLAGS {
        BIND_MAYBOTHERUSER = 1,
        BIND_JUSTTESTEXISTENCE = 2
    } BIND_FLAGS;

typedef enum ROTFLAGS {
    ROTFLAGS_REGISTRATIONKEEPSALIVE = 0x1,
    ROTFLAGS_ALLOWANYCLIENT = 0x2
} ROTFLAGS;

typedef enum INVOKEKIND {
    INVOKE_FUNC = 1,
    INVOKE_PROPERTYGET = 2,
    INVOKE_PROPERTYPUT = 4,
    INVOKE_PROPERTYPUTREF = 8
} INVOKEKIND;

typedef [helpstring("Standard DISPIDs")] enum STANDARD_DISPIDS {
    DISPID_AUTOSIZE = (-500),
    DISPID_BACKCOLOR = (-501),
    DISPID_BACKSTYLE = (-502),
    DISPID_BORDERCOLOR = (-503),
    DISPID_BORDERSTYLE = (-504),
    DISPID_BORDERWIDTH = (-505),
    DISPID_DRAWMODE = (-507),
    DISPID_DRAWSTYLE = (-508),
    DISPID_DRAWWIDTH = (-509),
    DISPID_FILLCOLOR = (-510),
    DISPID_FILLSTYLE = (-511),
    DISPID_FONT = (-512),
    DISPID_FORECOLOR = (-513),
    DISPID_ENABLED = (-514),
    DISPID_HWND = (-515),
    DISPID_TABSTOP = (-516),
    DISPID_TEXT = (-517),
    DISPID_CAPTION = (-518),
    DISPID_BORDERVISIBLE = (-519),
    DISPID_APPEARANCE = (-520),
    DISPID_MOUSEPOINTER = (-521),
    DISPID_MOUSEICON = (-522),
    DISPID_PICTURE = (-523),
    DISPID_VALID = (-524),
    DISPID_READYSTATE = (-525),
    DISPID_LISTINDEX = (-526),
    DISPID_SELECTED = (-527),
    DISPID_LIST = (-528),
    DISPID_COLUMN = (-529),
    DISPID_LISTCOUNT = (-531),
    DISPID_MULTISELECT = (-532),
    DISPID_MAXLENGTH = (-533),
    DISPID_PASSWORDCHAR = (-534),
    DISPID_SCROLLBARS = (-535),
    DISPID_WORDWRAP = (-536),
    DISPID_MULTILINE = (-537),
    DISPID_NUMBEROFROWS = (-538),
    DISPID_NUMBEROFCOLUMNS = (-539),
    DISPID_DISPLAYSTYLE = (-540),
    DISPID_GROUPNAME = (-541),
    DISPID_IMEMODE = (-542),
    DISPID_ACCELERATOR = (-543),
    DISPID_ENTERKEYBEHAVIOR = (-544),
    DISPID_TABKEYBEHAVIOR = (-545),
    DISPID_SELTEXT = (-546),
    DISPID_SELSTART = (-547),
    DISPID_SELLENGTH = (-548),

    DISPID_REFRESH = (-550),
    DISPID_DOCLICK = (-551),
    DISPID_ABOUTBOX = (-552),
    DISPID_ADDITEM = (-553),
    DISPID_CLEAR = (-554),
    DISPID_REMOVEITEM = (-555),

    DISPID_CLICK = (-600),
    DISPID_DBLCLICK = (-601),
    DISPID_KEYDOWN = (-602),
    DISPID_KEYPRESS = (-603),
    DISPID_KEYUP = (-604),
    DISPID_MOUSEDOWN = (-605),
    DISPID_MOUSEMOVE = (-606),
    DISPID_MOUSEUP = (-607),
    DISPID_ERROREVENT = (-608),
    DISPID_READYSTATECHANGE = (-609),
    DISPID_CLICK_VALUE = (-610),
    DISPID_RIGHTTOLEFT = (-611),
    DISPID_TOPTOBOTTOM = (-612),

    DISPID_NAME = (-800),
    DISPID_DELETE = (-801),
    DISPID_OBJECT = (-802),
    DISPID_PARENT = (-803)
} STANDARD_DISPIDS ;

typedef enum AMBIENT_DISPIDS {
    DISPID_AMBIENT_BACKCOLOR = (-701),
    DISPID_AMBIENT_DISPLAYNAME = (-702),
    DISPID_AMBIENT_FONT = (-703),
    DISPID_AMBIENT_FORECOLOR = (-704),
    DISPID_AMBIENT_LOCALEID = (-705),
    DISPID_AMBIENT_MESSAGEREFLECT = (-706),
    DISPID_AMBIENT_SCALEUNITS = (-707),
    DISPID_AMBIENT_TEXTALIGN = (-708),
    DISPID_AMBIENT_USERMODE = (-709),
    DISPID_AMBIENT_UIDEAD = (-710),
    DISPID_AMBIENT_SHOWGRABHANDLES = (-711),
    DISPID_AMBIENT_SHOWHATCHING = (-712),
    DISPID_AMBIENT_DISPLAYASDEFAULT = (-713),
    DISPID_AMBIENT_SUPPORTSMNEMONICS = (-714),
    DISPID_AMBIENT_AUTOCLIP = (-715),
    DISPID_AMBIENT_APPEARANCE = (-716),

    DISPID_AMBIENT_PALETTE = (-726),
    DISPID_AMBIENT_TRANSFERPRIORITY = (-728),

    DISPID_AMBIENT_RIGHTTOLEFT = (-732),
    DISPID_AMBIENT_TOPTOBOTTOM = (-733)

} AMBIENT_DISPIDS;

typedef enum OLEOBJECTS_DISPIDS {
    DISPID_FONT_NAME = 0,
    DISPID_FONT_SIZE = 2,
    DISPID_FONT_BOLD = 3,
    DISPID_FONT_ITALIC = 4,
    DISPID_FONT_UNDER = 5,
    DISPID_FONT_STRIKE = 6,
    DISPID_FONT_WEIGHT = 7,
    DISPID_FONT_CHARSET = 8,

    DISPID_PICT_HANDLE = 0,
    DISPID_PICT_HPAL = 2,
    DISPID_PICT_TYPE = 3,
    DISPID_PICT_WIDTH = 4,
    DISPID_PICT_HEIGHT = 5,
    DISPID_PICT_RENDER = 6
} OLEOBJECTS_DISPIDS ;

typedef enum TYPEFLAGS {
    TYPEFLAG_FAPPOBJECT = 0x01,
    TYPEFLAG_FCANCREATE = 0x02,
    TYPEFLAG_FLICENSED = 0x04,
    TYPEFLAG_FPREDECLID = 0x08,
    TYPEFLAG_FHIDDEN = 0x10,
    TYPEFLAG_FCONTROL = 0x20,
    TYPEFLAG_FDUAL = 0x40,
    TYPEFLAG_FNONEXTENSIBLE = 0x80,
    TYPEFLAG_FOLEAUTOMATION = 0x100,
    TYPEFLAG_FRESTRICTED = 0x200,
    TYPEFLAG_FAGGREGATABLE = 0x400,
    TYPEFLAG_FREPLACEABLE = 0x800,
    TYPEFLAG_FDISPATCHABLE = 0x1000,
    TYPEFLAG_FREVERSEBIND = 0x2000
} TYPEFLAGS;

    typedef enum TYPEKIND {
        TKIND_ENUM = 0,
        TKIND_RECORD,
        TKIND_MODULE,
        TKIND_INTERFACE,
        TKIND_DISPATCH,
        TKIND_COCLASS,
        TKIND_ALIAS,
        TKIND_UNION,
        TKIND_MAX
    } TYPEKIND;

    typedef enum SYSKIND {
        SYS_WIN16 = 0,
        SYS_WIN32,
        SYS_MAC
    } SYSKIND;

typedef enum VARKIND {
    VAR_PERINSTANCE,
    VAR_STATIC,
    VAR_CONST,
    VAR_DISPATCH
} VARKIND;

typedef enum IMPLTYPE {
    IMPLTYPEFLAG_FDEFAULT = 0x1,
    IMPLTYPEFLAG_FSOURCE = 0x2,
    IMPLTYPEFLAG_FRESTRICTED = 0x4,
    IMPLTYPEFLAG_FDEFAULTVTABLE= 0x8,
} IMPLTYPE;

typedef enum CALLCONV {
    CC_FASTCALL = 0,
    CC_CDECL = 1,
    CC_MSCPASCAL,
    CC_PASCAL = 2,
    CC_MACPASCAL,
    CC_STDCALL,
    CC_FPFASTCALL,
    CC_SYSCALL,
    CC_MPWCDECL,
    CC_MPWPASCAL,
    CC_MAX
} CALLCONV;

typedef enum FUNCKIND {
    FUNC_VIRTUAL,
    FUNC_PUREVIRTUAL,
    FUNC_NONVIRTUAL,
    FUNC_STATIC,
    FUNC_DISPATCH
} FUNCKIND;

typedef enum REGKIND {
    REGKIND_DEFAULT,
    REGKIND_REGISTER,
    REGKIND_NONE
} REGKIND;


typedef enum StrToIntExFlags {
    STIF_DEFAULT = 0x00000000,
    STIF_SUPPORT_HEX = 0x00000001
} StrToIntExFlags;

// Return flags for PathGetCharType
typedef enum GCTFlags {
    GCT_INVALID = 0x0000,
    GCT_LFNCHAR = 0x0001,
    GCT_SHORTCHAR = 0x0002,
    GCT_WILD = 0x0004,
    GCT_SEPARATOR = 0x0008
} GCTFlags;

typedef enum URL_SCHEME {
    URL_SCHEME_INVALID = -1,
    URL_SCHEME_UNKNOWN = 0,
    URL_SCHEME_FTP,
    URL_SCHEME_HTTP,
    URL_SCHEME_GOPHER,
    URL_SCHEME_MAILTO,
    URL_SCHEME_NEWS,
    URL_SCHEME_NNTP,
    URL_SCHEME_TELNET,
    URL_SCHEME_WAIS,
    URL_SCHEME_FILE,
    URL_SCHEME_MK,
    URL_SCHEME_HTTPS,
    URL_SCHEME_SHELL,
    URL_SCHEME_SNEWS,
    URL_SCHEME_LOCAL,
    URL_SCHEME_JAVASCRIPT,
    URL_SCHEME_VBSCRIPT,
    URL_SCHEME_ABOUT,
    URL_SCHEME_RES,
    URL_SCHEME_MAXVALUE
} URL_SCHEME;

typedef enum URL_PART {
    URL_PART_NONE = 0,
    URL_PART_SCHEME = 1,
    URL_PART_HOSTNAME,
    URL_PART_USERNAME,
    URL_PART_PASSshort,
    URL_PART_PORT,
    URL_PART_QUERY,
} URL_PART;

typedef enum URLIS {
    URLIS_URL,
    URLIS_OPAQUE,
    URLIS_NOHISTORY,
    URLIS_FILEURL,
    URLIS_APPLIABLE,
    URLIS_DIRECTORY,
    URLIS_HASQUERY,
} URLIS;

typedef enum URL {
    URL_UNESCAPE = 0x10000000,
    URL_ESCAPE_UNSAFE = 0x20000000,
    URL_PLUGGABLE_PROTOCOL = 0x40000000,
    URL_WININET_COMPATIBILITY = 0x80000000,
    URL_DONT_ESCAPE_EXTRA_INFO = 0x02000000,
    URL_DONT_UNESCAPE_EXTRA_INFO = 0x02000000,
    URL_BROWSER_MODE = 0x02000000,
    URL_ESCAPE_SPACES_ONLY = 0x04000000,
    URL_DONT_SIMPLIFY = 0x08000000,
    URL_NO_META = 0x08000000,
    URL_UNESCAPE_INPLACE = 0x00100000,
    URL_CONVERT_IF_DOSPATH = 0x00200000,
    URL_UNESCAPE_HIGH_ANSI_ONLY = 0x00400000,
    URL_INTERNAL_PATH = 0x00800000, // Will escape #'s in paths
    URL_FILE_USE_PATHURL = 0x00010000,
    URL_ESCAPE_PERCENT = 0x00001000,
    URL_ESCAPE_SEGMENT_ONLY = 0x00002000, // Treat the entire URL param as one URL segment
} URL;

typedef enum {
    URL_PARTFLAG_KEEPSCHEME = 0x00000001
} URL_PARTFLAG;

typedef enum {
    URL_APPLY_DEFAULT = 0x00000001,
    URL_APPLY_GUESSSCHEME = 0x00000002,
    URL_APPLY_GUESSFILE = 0x00000004,
    URL_APPLY_FORCEAPPLY = 0x00000008,
} URL_APPLY;

typedef enum {
    SHREGDEL_DEFAULT = 0x00000000, // Delete's HKCU, or HKLM if HKCU is not found.
    SHREGDEL_HKCU = 0x00000001, // Delete HKCU only
    SHREGDEL_HKLM = 0x00000010, // Delete HKLM only.
    SHREGDEL_BOTH = 0x00000011, // Delete both HKCU and HKLM.
} SHREGDEL_FLAGS;

typedef enum {
   SHREGENUM_DEFAULT = 0x00000000, // Enumerates HKCU or HKLM if not found.
   SHREGENUM_HKCU = 0x00000001, // Enumerates HKCU only
   SHREGENUM_HKLM = 0x00000010, // Enumerates HKLM only.
   SHREGENUM_BOTH = 0x00000011, // Enumerates both HKCU and HKLM without duplicates. This option is NYI.
} SHREGENUM_FLAGS;

typedef enum {
    SHREGSET_HKCU = 0x00000001, // Write to HKCU if empty.
    SHREGSET_FORCE_HKCU = 0x00000002, // Write to HKCU.
    SHREGSET_HKLM = 0x00000004, // Write to HKLM if empty.
    SHREGSET_FORCE_HKLM = 0x00000008, // Write to HKLM.
    SHREGSET_DEFAULT = 0x00000006 // Default is SHREGSET_FORCE_HKCU | SHREGSET_HKLM.
} SHREGSET;


typedef enum {
	KEY_QUERY_VALUE        = 0x0001,
	KEY_SET_VALUE          = 0x0002,
	KEY_CREATE_SUB_KEY     = 0x0004,
	KEY_ENUMERATE_SUB_KEYS = 0x0008,
	KEY_NOTIFY             = 0x0010,
	KEY_CREATE_LINK        = 0x0020,
	KEY_READ               = 0x00020019,
	KEY_WRITE              = 0x00020006,
	KEY_EXECUTE            = 0x00020019,
	KEY_ALL_ACCESS         = 0x000F003F
} REGSAM;

typedef enum {
    ASSOCF_INIT_NOREMAPCLSID = 0x00000001, // do not remap clsids to progids
    ASSOCF_INIT_BYEXENAME = 0x00000002, // executable is being passed in
    ASSOCF_OPEN_BYEXENAME = 0x00000002, // executable is being passed in
    ASSOCF_INIT_DEFAULTTOSTAR = 0x00000004, // treat "*" as the BaseClass
    ASSOCF_INIT_DEFAULTTOFOLDER = 0x00000008, // treat "Folder" as the BaseClass
    ASSOCF_NOUSERSETTINGS = 0x00000010, // dont use HKCU
    ASSOCF_NOTRUNCATE = 0x00000020, // dont truncate the return string
    ASSOCF_VERIFY = 0x00000040, // verify data is accurate (DISK HITS)
    ASSOCF_REMAPRUNDLL = 0x00000080, // actually gets info about rundlls target if applicable
    ASSOCF_NOFIXUPS = 0x00000100, // attempt to fix errors if found
    ASSOCF_IGNOREBASECLASS = 0x00000200,  // dont recurse into the baseclass
} ASSOCF;

typedef enum {
    ASSOCSTR_COMMAND = 1, // shell\verb\command string
    ASSOCSTR_EXECUTABLE, // the executable part of command string
    ASSOCSTR_FRIENDLYDOCNAME, // friendly name of the document type
    ASSOCSTR_FRIENDLYAPPNAME, // friendly name of executable
    ASSOCSTR_NOOPEN, // noopen value
    ASSOCSTR_SHELLNEWVALUE, // query values under the shellnew key
    ASSOCSTR_DDECOMMAND, // template for DDE commands
    ASSOCSTR_DDEIFEXEC, // DDECOMMAND to use if just create a process
    ASSOCSTR_DDEAPPLICATION, // Application name in DDE broadcast
    ASSOCSTR_DDETOPIC, // Topic Name in DDE broadcast
    ASSOCSTR_INFOTIP, // info tip for an item, or list of properties to create info tip from
    ASSOCSTR_QUICKTIP, // same as ASSOCSTR_INFOTIP, except, this list contains only quickly retrievable properties
    ASSOCSTR_MAX // last item in enum...
} ASSOCSTR;

typedef enum {
    ASSOCKEY_SHELLEXECCLASS = 1, // the key that should be passed to ShellExec(hkeyClass)
    ASSOCKEY_APP, // the "Application" key for the association
    ASSOCKEY_CLASS, // the progid or class key
    ASSOCKEY_BASECLASS, // the BaseClass key
    ASSOCKEY_MAX // last item in enum...
} ASSOCKEY;

typedef enum {
    ASSOCDATA_MSIDESCRIPTOR = 1, // Component Descriptor to pass to MSI APIs
    ASSOCDATA_NOACTIVATEHANDLER, // restrict attempts to activate window
    ASSOCDATA_QUERYCLASSSTORE, // should check with the NT Class Store
    ASSOCDATA_HASPERUSERASSOC, // defaults to user specified association
    ASSOCDATA_MAX
} ASSOCDATA;

typedef enum {
    ASSOCENUM_NONE
} ASSOCENUM;

typedef enum {
	SHACF_DEFAULT = 0x00000000, // Currently (SHACF_FILESYSTEM | SHACF_URLALL)
	SHACF_FILESYSTEM = 0x00000001, // This includes the File System as well as the rest of the shell (Desktop\My Computer\Control Panel\)
	SHACF_URLHISTORY = 0x00000002, // URLs in the User's History
	SHACF_URLMRU = 0x00000004, // URLs in the User's Recently Used list.
	SHACF_URLALL = 0x00000006, // (SHACF_URLHISTORY | SHACF_URLMRU)
	SHACF_USETAB = 0x00000008, // Use the tab to move thru the autocomplete possibilities instead of to the next dialog/window control.
	SHACF_FILESYS_ONLY = 0x00000010, // Don't AutoComplete non-File System items.
	SHACF_AUTOSUGGEST_FORCE_ON = 0x10000000, // Ignore the registry default and force the feature on.
	SHACF_AUTOSUGGEST_FORCE_OFF = 0x20000000, // Ignore the registry default and force the feature off.
	SHACF_AUTOAPPEND_FORCE_ON = 0x40000000, // Ignore the registry default and force the feature on. (Also know as AutoComplete)
	SHACF_AUTOAPPEND_FORCE_OFF = 0x80000000, // Ignore the registry default and force the feature off. (Also know as AutoComplete)
} SHACF;

typedef enum {
	CTF_INSIST = 0x00000001, // SHCreateThread() dwFlags - call pfnThreadProc synchronously if CreateThread() fails
	CTF_THREAD_REF = 0x00000002, // hold a reference to the creating thread
	CTF_PROCESS_REF = 0x00000004, // hold a reference to the creating process
	CTF_COINIT = 0x00000008, // init COM for the created thread
	CTF_FREELIBANDEXIT = 0x00000010, // hold a ref to the DLL and call FreeLibraryAndExitThread() when done
} CTF;

// Platform IDs for DLLVERSIONINFO
typedef enum {
	DLLVER_PLATFORM_WINDOWS = 0x00000001, // Windows 95
	DLLVER_PLATFORM_NT = 0x00000002, // Windows NT
} DLLVER;

// Browsing for directory.
//EBROWSEOPTIONS
typedef enum BROWSEOPTIONS {

	[ helpstring("Only return directories--grey out OK button for other selections")]    
    BIF_RETURNONLYFSDIRS   = 0x0001,  // Find folder to start document search
	[ helpstring("Don't include network folders below the domain level")]    
    BIF_DONTGOBELOWDOMAIN  = 0x0002,  
	[ helpstring("Include a status area with text coming from messages sent by the callback function")]    
    BIF_STATUSTEXT         = 0x0004,
	[ helpstring("Only return file system ancestors--grey out OK button for other selections")]    
    BIF_RETURNFSANCESTORS  = 0x0008,
	[ helpstring("Include edit control for name of item")]    
	BIF_EDITBOX            = 0x0010,
	[ helpstring("Insist on valid result (or CANCEL)--requires will BIF_EDITBOX and callback procedure")]    
	BIF_VALIDATE           = 0x0020,  
	[ helpstring("Only return computers--grey out OK button for other selections")]  
    BIF_BROWSEFORCOMPUTER  = 0x1000,  
	[ helpstring("Only return printers--grey out OK button for other selections")]    
    BIF_BROWSEFORPRINTER   = 0x2000,  
	[ helpstring("Display files as well as folders")]    
    BIF_BROWSEINCLUDEFILES = 0x4000   
	
} BROWSEOPTIONS;

    //// Shell File Operations
//EFILEOPS
typedef enum FILEOPS {
    FO_MOVE         = 0x0001,
    FO_COPY         = 0x0002,
    FO_DELETE       = 0x0003,
    FO_RENAME       = 0x0004,

    FOF_MULTIDESTFILES      = 0x0001,
    FOF_CONFIRMMOUSE        = 0x0002,
    FOF_SILENT              = 0x0004, // don't create progress/report
    FOF_RENAMEONCOLLISION   = 0x0008,
    FOF_NOCONFIRMATION      = 0x0010, // Don't prompt the user.
    FOF_WANTMAPPINGHANDLE   = 0x0020, // Fill in SHFILEOPSTRUCT.hNameMappings
                                      // Must be freed using SHFreeNameMappings
	FOF_SILEN_NOCONFIRMATION = 20,	  //For conveniece
    FOF_ALLOWUNDO           = 0x0040,
    FOF_FILESONLY           = 0x0080, // on *.*, do only files
    FOF_SIMPLEPROGRESS      = 0x0100, // means don't show names of files
    FOF_NOCONFIRMMKDIR      = 0x0200, // don't confirm making any needed dirs

    PO_DELETE               = 0x0013, // printer is being deleted
    PO_RENAME               = 0x0014, // printer is being renamed
    PO_PORTCHANGE           = 0x0020, // port this printer connected to is being changed
                                      // if this id is set, the strings received by
                                      // the copyhook are a doubly-null terminated
                                      // list of strings.  The first is the printer
                                      // name and the second is the printer port.
    PO_REN_PORT             = 0x0034, // PO_RENAME and PO_PORTCHANGE at same time.

} FILEOPS;

//ESHARD
typedef enum SHARD {
    [ helpstring("SHAddToRecentDocs: Add item identifier list") ]
    SHARD_PIDL   = 0x00000001,
    [ helpstring("SHAddToRecentDocs: Add path string ANSI") ]
    SHARD_PATHA  = 0x00000002,
	[ helpstring("SHAddToRecentDocs: Add path string Unicode") ]
    SHARD_PATHW  = 0x00000003
} SHARD;

//EESHGFI
typedef enum SHGFI {
    SHGFI_ICON               = 0x00000100,
    SHGFI_DISPLAYNAME        = 0x00000200,
    SHGFI_TYPENAME           = 0x00000400,
    SHGFI_ATTRIBUTES         = 0x00000800,
    SHGFI_ICONLOCATION       = 0x00001000,
    SHGFI_EXETYPE            = 0x00002000,
    SHGFI_SYSICONINDEX       = 0x00004000,
    SHGFI_LINKOVERLAY        = 0x00008000,
    SHGFI_SELECTED           = 0x00010000,
    SHGFI_LARGEICON          = 0x00000000,
    SHGFI_SMALLICON          = 0x00000001,
    SHGFI_OPENICON           = 0x00000002,
    SHGFI_SHELLICONSIZE      = 0x00000004,
    SHGFI_PIDL               = 0x00000008,
    SHGFI_USEFILEATTRIBUTES  = 0x00000010,
	BASIC_SHGFI_FLAGS		 = 26116
} SHGFI;

//ECSIDL
typedef enum CSIDL {
	CSIDL_DESKTOP                 = 0x0000, // <desktop>
	CSIDL_INTERNET                = 0x0001, 
	CSIDL_PROGRAMS                = 0x0002,
	CSIDL_CONTROLS                = 0x0003,
	CSIDL_PRINTERS                = 0x0004,
	CSIDL_PERSONAL                = 0x0005, // My Documents
	CSIDL_FAVORITES               = 0x0006,
	CSIDL_STARTUP                 = 0x0007,
	CSIDL_RECENT                  = 0x0008,
	CSIDL_SENDTO                  = 0x0009,
	CSIDL_BITBUCKET               = 0x000a,
	CSIDL_STARTMENU               = 0x000b,
	CSIDL_MYDOCUMENTS             = 0x000c,        // logical "My Documents" desktop icon
	CSIDL_MYMUSIC                 = 0x000d,        // "My Music" folder
	CSIDL_MYVIDEO                 = 0x000e,        // "My Videos" folder
	CSIDL_DESKTOPDIRECTORY        = 0x0010,
	CSIDL_DRIVES                  = 0x0011,
	CSIDL_NETWORK                 = 0x0012,
	CSIDL_NETHOOD                 = 0x0013,
	CSIDL_FONTS                   = 0x0014,
	CSIDL_TEMPLATES               = 0x0015,
	CSIDL_COMMON_STARTMENU        = 0x0016,
	CSIDL_COMMON_PROGRAMS         = 0x0017,
	CSIDL_COMMON_STARTUP          = 0x0018,
	CSIDL_COMMON_DESKTOPDIRECTORY = 0x0019,
	CSIDL_APPDATA                 = 0x001a,
	CSIDL_PRINTHOOD               = 0x001b,
	CSIDL_LOCAL_APPDATA           = 0x001c,        // <user name>\Local Settings\Applicaiton Data (non roaming)
	CSIDL_ALTSTARTUP              = 0x001d,        // DBCS
	CSIDL_COMMON_ALTSTARTUP       = 0x001e,        // DBCS
	CSIDL_COMMON_FAVORITES        = 0x001f,
	CSIDL_INTERNET_CACHE          = 0x0020,
	CSIDL_COOKIES                 = 0x0021,
	CSIDL_HISTORY                 = 0x0022,
	CSIDL_COMMON_APPDATA          = 0x0023,        // All Users\Application Data
	CSIDL_WINDOWS                 = 0x0024,        // GetWindowsDirectory()
	CSIDL_SYSTEM                  = 0x0025,        // GetSystemDirectory()
	CSIDL_PROGRAM_FILES           = 0x0026,        // C:\Program Files
	CSIDL_MYPICTURES              = 0x0027,        // C:\Program Files\My Pictures
	CSIDL_PROFILE                 = 0x0028,        // USERPROFILE
	CSIDL_SYSTEMX86               = 0x0029,        // x86 system directory on RISC
	CSIDL_PROGRAM_FILESX86        = 0x002a,        // x86 C:\Program Files on RISC
	CSIDL_PROGRAM_FILES_COMMON    = 0x002b,        // C:\Program Files\Common
	CSIDL_PROGRAM_FILES_COMMONX86 = 0x002c,        // x86 Program Files\Common on RISC
	CSIDL_COMMON_TEMPLATES        = 0x002d,        // All Users\Templates
	CSIDL_COMMON_DOCUMENTS        = 0x002e,        // All Users\Documents
	CSIDL_COMMON_ADMINTOOLS       = 0x002f,        // All Users\Start Menu\Programs\Administrative Tools
	CSIDL_ADMINTOOLS              = 0x0030,        // <user name>\Start Menu\Programs\Administrative Tools
	CSIDL_CONNECTIONS             = 0x0031,        // Network and Dial-up Connections
	CSIDL_COMMON_MUSIC            = 0x0035,        // All Users\My Music
	CSIDL_COMMON_PICTURES         = 0x0036,        // All Users\My Pictures
	CSIDL_COMMON_VIDEO            = 0x0037,        // All Users\My Video
	CSIDL_RESOURCES               = 0x0038,        // %windir%\Resources\, For theme and other windows resources.
	CSIDL_RESOURCES_LOCALIZED     = 0x0039,        // %windir%\Resources\<LangID>, for theme and other windows specific resources.
	CSIDL_COMMON_OEM_LINKS        = 0x003a,        // Links to All Users OEM specific apps
	CSIDL_CDBURN_AREA             = 0x003b,        // USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
	CSIDL_COMPUTERSNEARME         = 0x003d,        // Computers Near Me (computered from Workgroup membership)

	CSIDL_FLAG_PER_USER_INIT      = 0x0800,        // combine with CSIDL_ value to indicate per-user init (eg. upgrade)
	CSIDL_FLAG_NO_ALIAS           = 0x1000,        // combine with CSIDL_ value to insure non-alias versions of the pidl
	CSIDL_FLAG_CREATE             = 0x8000,        // combine with CSIDL_ value to force folder creation in SHGetFolderPath()
	CSIDL_FLAG_DONT_VERIFY        = 0x4000,        // combine with CSIDL_ value to return an unverified folder path
	CSIDL_FLAG_MASK               = 0xFF00,        // mask for all possible flag values

	CSIDL_TEMPORARY_INTERNET_FILES_DIR	= 32,
	CSIDL_TEMPORARY_DIR					= 42,
} CSIDL;

typedef enum SHGDFIL {
    SHGDFIL_FINDDATA     = 1,
    SHGDFIL_NETRESOURCE  = 2
} SHGDFIL;


typedef enum {
    INTERNET_SCHEME_PARTIAL = -2,
    INTERNET_SCHEME_UNKNOWN = -1,
    INTERNET_SCHEME_DEFAULT = 0,
    INTERNET_SCHEME_FTP,
    INTERNET_SCHEME_GOPHER,
    INTERNET_SCHEME_HTTP,
    INTERNET_SCHEME_HTTPS,
    INTERNET_SCHEME_FILE,
    INTERNET_SCHEME_NEWS,
    INTERNET_SCHEME_MAILTO,
    INTERNET_SCHEME_SOCKS,
    INTERNET_SCHEME_FIRST = 1,
    INTERNET_SCHEME_LAST = 8
} INTERNET_SCHEME;

// flags for InternetCanonicalizeUrl() and InternetCombineUrl()
typedef enum CanonicalizingStyles {
	ICU_NO_ENCODE   = 0x20000000,  // Don't convert unsafe characters to escape sequence
	ICU_DECODE      = 0x10000000,  // Convert %XX escape sequences to characters
	ICU_NO_META     = 0x08000000,  // Don't convert .. etc. meta path sequences
	ICU_ENCODE_SPACES_ONLY = 0x04000000,  // Encode spaces only
	ICU_BROWSER_MODE = 0x02000000, // Special encode/decode rules for browser
} CanonicalizingStyles;

// *********** COMCTL32

    typedef enum ICC_CLASSES {
        [helpstring("listview, header")]
        ICC_LISTVIEW_CLASSES = 0x00000001,
        [helpstring("treeview, tooltips ")]
        ICC_TREEVIEW_CLASSES = 0x00000002,
        [helpstring("toolbar, statusbar, trackbar, tooltips ")]
        ICC_BAR_CLASSES = 0x00000004,
        [helpstring("tab, tooltips")]
        ICC_TAB_CLASSES = 0x00000008,
        [helpstring("updown")]
        ICC_UPDOWN_CLASS = 0x00000010,
        [helpstring("progress ")]
        ICC_PROGRESS_CLASS = 0x00000020,
        [helpstring("hotkey ")]
        ICC_HOTKEY_CLASS = 0x00000040,
        [helpstring("animate")]
        ICC_ANIMATE_CLASS = 0x00000080,
        ICC_WIN95_CLASSES = 0x000000FF,
        [helpstring("month picker, date picker, time picker, updown")]
        ICC_DATE_CLASSES = 0x00000100,
        [helpstring("comboex")]
        ICC_USEREX_CLASSES = 0x00000200,
        [helpstring("rebar (coolbar) control")]
        ICC_COOL_CLASSES = 0x00000400,
        ICC_INTERNET_CLASSES = 0x00000800,
        [helpstring("page scroller ")]
        ICC_PAGESCROLLER_CLASS = 0x00001000,
        [helpstring("native font control")]
        ICC_NATIVEFNTCTL_CLASS = 0x00002000,
    } ICC_CLASSES;

    typedef enum ODT {
        ODT_HEADER = 100,
        ODT_TAB = 101,
        ODT_LISTVIEW = 102,
    } ODT;

    typedef enum CommonStyles {
        CCS_TOP = 0x00000001,
        CCS_NOMOVEY = 0x00000002,
        CCS_BOTTOM = 0x00000003,
        CCS_NORESIZE = 0x00000004,
        CCS_NOPARENTALIGN = 0x00000008,
        CCS_ADJUSTABLE = 0x00000020,
        CCS_NODIVIDER = 0x00000040,
        CCS_VERT = 0x00000080,
        CCS_LEFT = 129, //(CCS_VERT | CCS_TOP),
        CCS_RIGHT = 131, //(CCS_VERT | CCS_BOTTOM),
        CCS_NOMOVEX = 130, //(CCS_VERT | CCS_NOMOVEY),
    } ComCtlStyles;

    typedef enum CommonMessages {
        CCM_FIRST = 0x2000, // Common control shared messages 8192
        CCM_SETBKCOLOR = 0x2001, // lParam is bkColor
        CCM_SETCOLORSCHEME = 0x2002, // lParam is color scheme
        CCM_GETCOLORSCHEME = (8195), // fills in COLORSCHEME pointed to by lParam
        CCM_GETDROPTARGET = (8196),
        CCM_SETUNICODEFORMAT = (8197),
        CCM_GETUNICODEFORMAT = (8198),
        CCM_SETVERSION = (8199),
        CCM_GETVERSION = (8200),
        CCM_SETNOTIFYWINDOW = (8201), // wParam = hwndParent.
    } CommonMessages;

    typedef enum NotificationCodes {
        NM_FIRST = 0, // generic to all controls
        NM_OUTOFMEMORY = (-1),
        NM_CLICK = (-2),
        NM_DBLCLK = (-3),
        NM_RETURN = (-4),
        NM_RCLICK = (-5),
        NM_RDBLCLK = (-6),
        NM_SETFOCUS = (-7),
        NM_KILLFOCUS = (-8),
        NM_CUSTOMDRAW = (-12),
        NM_HOVER = (-13),
        NM_NCHITTEST = (-14), // uses NMMOUSE struct
        NM_KEYDOWN = (-15),
        NM_RELEASEDCAPTURE = (-16),
        NM_SETCURSOR = (-17), // uses NMMOUSE struct
        NM_CHAR = (-18), // uses NMCHAR struct
        NM_TOOLTIPSCREATED = (-19), // notify of when the tooltips window is create
        NM_LDOWN = (-20),
        NM_RDOWN = (-21),
        NM_LAST = (-99)
    } NotificationCodes;

    typedef enum MSGFs {
        MSGF_COMMCTRL_BEGINDRAG = 0x4200,
        MSGF_COMMCTRL_SIZEHEADER = 0x4201,
        MSGF_COMMCTRL_DRAGSELECT = 0x4202,
        MSGF_COMMCTRL_TOOLBARCUST = 0x4203,
    } MSGFs;

    typedef enum CustomDrawReturnFlags {
        CDRF_DODEFAULT = 0x00000000,
        CDRF_NEWFONT = 0x00000002,
        CDRF_SKIPDEFAULT = 0x00000004,
        CDRF_NOTIFYPOSTPAINT = 0x00000010,
        CDRF_NOTIFYITEMDRAW = 0x00000020,
        CDRF_NOTIFYSUBITEMDRAW = 0x00000020, // flags are the same, we can distinguish by context
        CDRF_NOTIFYPOSTERASE = 0x00000040,
    } CustomDrawReturnFlags;

    typedef enum DrawStageFlags {
        CDDS_PREPAINT = 0x00000001,
        CDDS_POSTPAINT = 0x00000002,
        CDDS_PREERASE = 0x00000003,
        CDDS_POSTERASE = 0x00000004,
        CDDS_ITEM = 0x00010000,
		CDDS_ITEMPREPAINT  = (65536 + 1),	//(CDDS_ITEM | CDDS_PREPAINT)
		CDDS_ITEMPOSTPAINT = (65536 + 2),	//(CDDS_ITEM | CDDS_POSTPAINT)
		CDDS_ITEMPREERASE  = (65536 + 3),	//(CDDS_ITEM | CDDS_PREERASE)
		CDDS_ITEMPOSTERASE = (65536 + 4),	//(CDDS_ITEM | CDDS_POSTERASE)
        CDDS_SUBITEM = 0x00020000,			//131072 
		CDDS_SUBITEMPOSPAINT = 196609,		//(CDDS_SUBITEM | CDDS_ITEMPOSTPAINT) ????
    } DrawStageFlags;

    typedef enum ItemStateFlags {
        CDIS_SELECTED = 0x0001,
        CDIS_GRAYED = 0x0002,
        CDIS_DISABLED = 0x0004,
        CDIS_CHECKED = 0x0008,
        CDIS_FOCUS = 0x0010,
        CDIS_DEFAULT = 0x0020,
        CDIS_HOT = 0x0040,
        CDIS_MARKED = 0x0080,
        CDIS_INDETERMINATE = 0x0100,
    } ItemStateFlags;

    typedef enum ColorConstants {
        CLR_NONE = 0xFFFFFFFF,
        CLR_DEFAULT = 0xFF000000,
    } ColorConstants;


    typedef enum ImageList_Flags {
        ILC_MASK = 0x0001,
        ILC_COLOR = 0x0000,
        ILC_COLORDDB = 0x00FE,
        ILC_COLOR4 = 0x0004,
        ILC_COLOR8 = 0x0008,
        ILC_COLOR16 = 0x0010,
        ILC_COLOR24 = 0x0018,
        ILC_COLOR32 = 0x0020,
        ILC_PALETTE = 0x0800,
    } ImageList_Flags;

    typedef enum ImageList_DrawFlags {
        ILD_NORMAL = 0x0000,
        ILD_TRANSPARENT = 0x0001,
        ILD_MASK = 0x0010,
        ILD_IMAGE = 0x0020,
        ILD_ROP = 0x0040,
        ILD_BLEND25 = 0x0002,
        ILD_BLEND50 = 0x0004,
        ILD_OVERLAYMASK = 0x0F00,
    } ImageList_DrawFlags;

    typedef enum ImageList_CopyFlags {
        ILCF_MOVE = (0x00000000),
        ILCF_SWAP = (0x00000001),
    } ImageList_CopyFlags;

    typedef enum Header_Styles {
        HDS_HORZ = 0x0000,
        HDS_BUTTONS = 0x0002,
        HDS_HOTTRACK = 0x0004,
        HDS_HIDDEN = 0x0008,
        HDS_DRAGDROP = 0x0040,
        HDS_FULLDRAG = 0x0080,
        HDS_FILTERBAR = 0x0100,
    } HeaderStyles;

    typedef enum HeaderFilterFlags {
        HDFT_ISSTRING = 0x0000, // HD_ITEM.pvFilter points to a HD_TEXTFILTER
        HDFT_ISNUMBER = 0x0001, // HD_ITEM.pvFilter points to a INT
        HDFT_HASNOVALUE = 0x8000, // clear the filter, by setting this bit
    } HeaderFilterFlags;

    typedef enum Header_ItemMask {
        HDI_WIDTH = 0x0001,
        HDI_HEIGHT = 0x0001,
        HDI_TEXT = 0x0002,
        HDI_FORMAT = 0x0004,
        HDI_LPARAM = 0x0008,
        HDI_BITMAP = 0x0010,
        HDI_IMAGE = 0x0020,
        HDI_DI_SETITEM = 0x0040,
        HDI_ORDER = 0x0080,
        HDI_FILTER = 0x0100,
    } HeaderItemMask;

    typedef enum Header_Formats {
        HDF_LEFT = 0,
        HDF_RIGHT = 1,
        HDF_CENTER = 2,
        HDF_JUSTIFYMASK = 3,
        HDF_RTLREADING = 4,
        HDF_BITMAP_ON_RIGHT = 0x1000,
        HDF_BITMAP = 0x2000,
        HDF_STRING = 0x4000,
        HDF_OWNERDRAW = 0x8000,
        HDF_IMAGE = 0x0800,
    } HeaderFormats;

    typedef enum Header_Messages {
		HDM_FIRST = (4608),
        HDM_GETITEMCOUNT = (4608),
        HDM_INSERTITEMA = (4609),
        HDM_DELETEITEM = (4610),
        HDM_GETITEMA = (4611),
        HDM_SETITEMA = (4612),
        HDM_LAYOUT = (4613),
        HDM_HITTEST = (4614),
        HDM_GETITEMRECT = (4615),
        HDM_SETIMAGELIST = (4616),
        HDM_GETIMAGELIST = (4617),
        HDM_INSERTITEMW = (4618),
        HDM_GETITEMW = (4619),
        HDM_SETITEMW = (4620),
        HDM_ORDERTOINDEX = (4623),
        HDM_CREATEDRAGIMAGE = (4624),
        HDM_GETORDERARRAY = (4625),
        HDM_SETORDERARRAY = (4626),
        HDM_SETHOTDIVIDER = (4627),
        HDM_SETBITMAPMARGIN = (4628),
        HDM_GETBITMAPMARGIN = (4629),
        HDM_SETFILTERCHANGETIMEOUT = (4630),
        HDM_EDITFILTER = (4631),
        HDM_CLEARFILTER = (4632),
		HDM_SETUNICODEFORMAT = (8197),
    } HeaderMessages;

    typedef enum Header_HitTestValues {
        HHT_NOWHERE = 0x0001,
        HHT_ONHEADER = 0x0002,
        HHT_ONDIVIDER = 0x0004,
        HHT_ONDIVOPEN = 0x0008,
        HHT_ONFILTER = 0x0010,
        HHT_ONFILTERBUTTON = 0x0020,
        HHT_ABOVE = 0x0100,
        HHT_BELOW = 0x0200,
        HHT_TORIGHT = 0x0400,
        HHT_TOLEFT = 0x0800,
    } HeaderHitTestValues;

    typedef enum Header_NotifyMessages {
        HDN_ITEMCHANGINGA = (-300),
        HDN_ITEMCHANGINGW = (-320),
        HDN_ITEMCHANGEDA = (-301),
        HDN_ITEMCHANGEDW = (-321),
        HDN_ITEMCLICKA = (-302),
        HDN_ITEMCLICKW = (-322),
        HDN_ITEMDBLCLICKA = (-303),
        HDN_ITEMDBLCLICKW = (-323),
        HDN_DIVIDERDBLCLICKA = (-305),
        HDN_DIVIDERDBLCLICKW = (-325),
        HDN_BEGINTRACKA = (-306),
        HDN_BEGINTRACKW = (-326),
        HDN_ENDTRACKA = (-307),
        HDN_ENDTRACKW = (-327),
        HDN_TRACKA = (-308),
        HDN_TRACKW = (-328),
        HDN_GETDISPINFOA = (-309),
        HDN_GETDISPINFOW = (-329),
        HDN_BEGINDRAG = (-310),
        HDN_ENDDRAG = (-311),
        HDN_FILTERCHANGE = (-312),
        HDN_FILTERBTNCLICK = (-313),
    } HeaderNotifyMessages;

    typedef enum Toolbar_ButtonStates {
        TBSTATE_CHECKED = 0x01,
        TBSTATE_PRESSED = 0x02,
        TBSTATE_ENABLED = 0x04,
        TBSTATE_HIDDEN = 0x08,
        TBSTATE_INDETERMINATE = 0x10,
        TBSTATE_WRAP = 0x20,
        TBSTATE_ELLIPSES = 0x40,
        TBSTATE_MARKED = 0x80,
    } ToolBarButtonStates;

    typedef enum ToolbarStyles {
		TBSTYLE_CHECK = 0x0002,
		TBSTYLE_GROUP = 0x0004,
        TBSTYLE_TOOLTIPS = 0x0100,
        TBSTYLE_WRAPABLE = 0x0200,
        TBSTYLE_ALTDRAG = 0x0400,
        TBSTYLE_FLAT = 0x0800,
        TBSTYLE_LIST = 0x1000,
        TBSTYLE_CUSTOMERASE = 0x2000,
        TBSTYLE_REGISTERDROP = 0x4000,
        TBSTYLE_TRANSPARENT = 0x8000,
        TBSTYLE_EX_DRAWDDARROWS = 0x00000001,
        TBSTYLE_EX_MIXEDBUTTONS = 0x00000008,
        TBSTYLE_EX_HIDECLIPPEDBUTTONS = 0x00000010, // don't show partially obscured buttons
		TBSTYLE_CHECKGROUP = 0x0006, //(TBSTYLE_GROUP Or TBSTYLE_CHECK)
    } ToolbarStyles;

    typedef enum ToolbarButtonStyles {
        BTNS_BUTTON = 0x0000,
        BTNS_SEP = 0x0001,
        BTNS_CHECK = 0x0002,
        BTNS_GROUP = 0x0004,
        BTNS_CHECKGROUP = 0x0006, //6 (BTNS_GROUP | BTNS_CHECK),
        BTNS_DROPDOWN = 0x0008,
        BTNS_AUTOSIZE = 0x0010, // automatically calculate the cx of the button
        BTNS_NOPREFIX = 0x0020, // this button should not have accel prefix
        BTNS_SHOWTEXT = 0x0040, // ignored unless TBSTYLE_EX_MIXEDBUTTONS is set
        BTNS_WHOLEDROPDOWN = 0x0080, // draw drop-down arrow, but without split arrow section
    } ToolbarButtonStyles;

    typedef enum ToolbarCustomDrawReturnFlags {
        TBCDRF_NOEDGES = 0x00010000, // Don't draw button edges
        TBCDRF_HILITEHOTTRACK = 0x00020000, // Use color of the button bk when hottracked
        TBCDRF_NOOFFSET = 0x00040000, // Don't offset button if pressed
        TBCDRF_NOMARK = 0x00080000, // Don't draw default highlight of image/text for TBSTATE_MARKED
        TBCDRF_NOETCHEDEFFECT = 0x00100000, // Don't draw etched effect for disabled items
        TBCDRF_BLENDICON = 0x00200000, // Use ILD_BLEND50 on the icon image 
    } ToolbarCustomDrawReturnFlags;

    typedef enum CreateMappedBitmapFlags {
        CMB_MASKED = 0x02,
    } CreateMappedBitmapFlags;

   typedef enum Toolbar_Messages {
        TB_ENABLEBUTTON = (1025), //(WM_USER + 1) and so on
        TB_CHECKBUTTON = (1026),
        TB_PRESSBUTTON = (1027),
        TB_HIDEBUTTON = (1028),
        TB_INDETERMINATE = (1029),
        TB_MARKBUTTON = (1030),
        TB_ISBUTTONENABLED = (1033),
        TB_ISBUTTONCHECKED = (1034),
        TB_ISBUTTONPRESSED = (1035),
        TB_ISBUTTONHIDDEN = (1036),
        TB_ISBUTTONINDETERMINATE = (1037),
        TB_ISBUTTONHIGHLIGHTED = (1038),
        TB_SETSTATE = (1041),
        TB_GETSTATE = (1042),
        TB_ADDBITMAP = (1043),
        TB_ADDBUTTONS = (1044),
        TB_INSERTBUTTON = (1045),
        TB_DELETEBUTTON = (1046),
        TB_GETBUTTON = (1047),
        TB_BUTTONCOUNT = (1048),
        TB_COMMANDTOINDEX = (1049),
        TB_SAVERESTORE = (1050),
        TB_SAVERESTOREW = (1100),
        TB_CUSTOMIZE = (1051),
        TB_ADDSTRINGA = (1052),
        TB_ADDSTRINGW = (1101),
        TB_GETITEMRECT = (1053),
        TB_BUTTONSTRUCTSIZE = (1054),
        TB_SETBUTTONSIZE = (1055),
        TB_SETBITMAPSIZE = (1056),
        TB_AUTOSIZE = (1057),
        TB_GETTOOLTIPS = (1059),
        TB_SETTOOLTIPS = (1060),
        TB_SETPARENT = (1061),
        TB_SETROWS = (1063),
        TB_GETROWS = (1064),
        TB_SETCMDID = (1066),
        TB_CHANGEBITMAP = (1067),
        TB_GETBITMAP = (1068),
        TB_GETBUTTONTEXT = (1069),
        TB_GETBUTTONTEXTW = (1099),
        TB_REPLACEBITMAP = (1070),
        TB_SETINDENT = (1071),
        TB_SETIMAGELIST = (1072),
        TB_GETIMAGELIST = (1073),
        TB_LOADIMAGES = (1074),
        TB_GETRECT = (1075), // wParam is the Cmd instead of index
        TB_SETHOTIMAGELIST = (1076),
        TB_GETHOTIMAGELIST = (1077),
        TB_SETDISABLEDIMAGELIST = (1078),
        TB_GETDISABLEDIMAGELIST = (1079),
        TB_SETSTYLE = (1080),
        TB_GETSTYLE = (1081),
        TB_GETBUTTONSIZE = (1082),
        TB_SETBUTTONWIDTH = (1083),
        TB_SETMAXTEXTROWS = (1084),
        TB_GETTEXTROWS = (1085),
        TB_SETCOLORSCHEME = 0x2002, // lParam is color scheme
        TB_GETOBJECT = (1086), // wParam = = IID, lParam void **ppv
        TB_GETHOTITEM = (1095),
        TB_SETHOTITEM = (1096), // wParam = = iHotItem
        TB_SETANCHORHIGHLIGHT = (1097), // wParam = = TRUE/FALSE
        TB_GETANCHORHIGHLIGHT = (1098),
        TB_MAPACCELERATOR = (1102), // wParam = = ch, lParam LONG * pidBtn
        TB_GETINSERTMARK = (1103), // lParam = = LPTBINSERTMARK
        TB_SETINSERTMARK = (1104), // lParam = = LPTBINSERTMARK
        TB_INSERTMARKHITTEST = (1105), // wParam = = LPPOINT lParam = = LPTBINSERTMARK
        TB_MOVEBUTTON = (1106),
        TB_GETMAXSIZE = (1107), // lParam = = LPSIZE
        TB_SETEXTENDEDSTYLE = (1108), // For TBSTYLE_EX_*
        TB_GETEXTENDEDSTYLE = (1109), // For TBSTYLE_EX_*
        TB_GETBITMAPFLAGS = (1065),
        TB_GETBUTTONINFOW = (1087),
        TB_SETBUTTONINFOW = (1088),
        TB_GETBUTTONINFO = (1089),
        TB_SETBUTTONINFO = (1090),
        TB_INSERTBUTTONW = (1091),
        TB_ADDBUTTONSW = (1092),
        TB_HITTEST = (1093),
        TB_SETDRAWTEXTFLAGS = (1094), // wParam = = mask lParam = = bit values
        TB_GETPADDING = (1110),
        TB_SETPADDING = (1111),
        TB_SETINSERTMARKCOLOR = (1112),
        TB_GETINSERTMARKCOLOR = (1113),
        TB_MAPACCELERATORW = (1114), // wParam = = ch, lParam int * pidBtn
    } ToolbarMessages;

    typedef enum Toolbar_InsertMarkFlags {
        TBIMHT_AFTER = 0x00000001, // TRUE = insert After iButton, otherwise before
        TBIMHT_BACKGROUND = 0x00000002, // TRUE iff missed buttons completely
    } ToolbarInsertMarkFlags;

    typedef enum Toolbar_AddBitmapValues {
        HINST_COMMCTRL = (-1),
        IDB_STD_SMALL_COLOR = 0,
        IDB_STD_LARGE_COLOR = 1,
        IDB_VIEW_SMALL_COLOR = 4,
        IDB_VIEW_LARGE_COLOR = 5,
        IDB_HIST_SMALL_COLOR = 8,
        IDB_HIST_LARGE_COLOR = 9,
    } ToolbarAddBitmapValues;

 // icon indexes for standard bitmap
    typedef enum Toolbar_StandardBitmaps {
        STD_CUT = 0,
        STD_COPY = 1,
        STD_PASTE = 2,
        STD_UNDO = 3,
        STD_REDOW = 4,
        STD_DELETE = 5,
        STD_FILENEW = 6,
        STD_FILEOPEN = 7,
        STD_FILESAVE = 8,
        STD_PRINTPRE = 9,
        STD_PROPERTIES = 10,
        STD_HELP = 11,
        STD_FIND = 12,
        STD_REPLACE = 13,
        STD_PRINT = 14,
        VIEW_LARGEICONS = 0,
        VIEW_SMALLICONS = 1,
        VIEW_LIST = 2,
        VIEW_DETAILS = 3,
        VIEW_SORTNAME = 4,
        VIEW_SORTSIZE = 5,
        VIEW_SORTDATE = 6,
        VIEW_SORTTYPE = 7,
        VIEW_PARENTFOLDER = 8,
        VIEW_NETCONNECT = 9,
        VIEW_NETDISCONNECT = 10,
        VIEW_NEWFOLDER = 11,
        VIEW_VIEWMENU = 12,
        HIST_BACK = 0,
        HIST_FORWARD = 1,
        HIST_FAVORITES = 2,
        HIST_ADDTOFAVORITES = 3,
        HIST_VIEWTREE = 4,
    } ToolbarStandardBitmaps;

    typedef enum Toolbar_BitmapFlags {
        TBBF_LARGE = 0x0001,
    } ToolbarBitmapFlags;

    typedef enum Toolbar_ButtonInfoMask {
        TBIF_IMAGE = 0x00000001,
        TBIF_TEXT = 0x00000002,
        TBIF_STATE = 0x00000004,
        TBIF_STYLE = 0x00000008,
        TBIF_LPARAM = 0x00000010,
        TBIF_COMMAND = 0x00000020,
        TBIF_SIZE = 0x00000040,
    } ToolbarButtonInfoMask;

    typedef enum Toolbar_HotItemFlags {
        HICF_OTHER = 0x00000000,
        HICF_MOUSE = 0x00000001, // Triggered by mouse
        HICF_ARROWKEYS = 0x00000002, // Triggered by arrow keys
        HICF_ACCELERATOR = 0x00000004, // Triggered by accelerator
        HICF_DROPDOWN = 0x00000008, // Button dropped down
        HICF_ENTERING = 0x00000010, // idOld is invalid
        HICF_LEAVING = 0x00000020, // idNew is invalid
        HICF_RESELECT = 0x00000040, // hot item reselected
        HICF_LMOUSE = 0x00000080, // left mouse button selected
        HICF_TOGGLEDROPDOWN = 0x00000100, // Toggle button's dropdown state
    } ToolbarHotItemFlags;


    typedef enum Toolbar_NotifyCodes {
        TBN_GETBUTTONINFO = (-700),
        TBN_GETBUTTONINFOW = (-720),
        TBN_BEGINDRAG = (-701),
        TBN_ENDDRAG = (-702),
        TBN_BEGINADJUST = (-703),
        TBN_ENDADJUST = (-704),
        TBN_RESET = (-705),
        TBN_QUERYINSERT = (-706),
        TBN_QUERYDELETE = (-707),
        TBN_TOOLBARCHANGE = (-708),
        TBN_CUSTHELP = (-709),
        TBN_DROPDOWN = (-710),
        TBN_GETOBJECT = (-712),
        TBN_HOTITEMCHANGE = (-713),
        TBN_DRAGOUT = (-714), // this is sent when the user clicks down on a button then drags off the button
        TBN_DELETINGBUTTON = (-715), // uses TBNOTIFY
        TBN_GETDISPINFO = (-716), // This is sent when the  toolbar needs  some display information
        TBN_GETDISPINFOW = (-717), // This is sent when the  toolbar needs  some display information
        TBN_GETINFOTIP = (-718),
        TBN_GETINFOTIPW = (-719),
        TBN_RESTORE = (-721),
        TBN_SAVE = (-722),
        TBN_INITCUSTOMIZE = (-723) //or (-700 - 23),
    } ToolbarNotifyCodes;

    typedef enum ToolbarCustomizeReturnFlags {
        TBNRF_HIDEHELP = 0x00000001,
        TBNRF_ENDCUSTOMIZE = 0x00000002,
    } ToolbarCustomizeReturnFlags;

    typedef enum Toolbar_NotifyFlags {
        TBNF_IMAGE = 0x00000001,
        TBNF_TEXT = 0x00000002,
        TBNF_DI_SETITEM = 0x10000000,
    } ToolbarNotifyFlags;

    typedef enum ToolbarDropdownReturnValues {
        TBDDRET_DEFAULT = 0,
        TBDDRET_NODEFAULT = 1,
        TBDDRET_TREATPRESSED = 2, // Treat as a standard press button
    } ToolbarDropdownReturnValues;

    typedef enum ReBar_IM {
        RBIM_IMAGELIST = 0x00000001,
    } ReBar_IM;

    typedef enum ReBar_Styles {
        RBS_TOOLTIPS = 0x0100,
        RBS_VARHEIGHT = 0x0200,
        RBS_BANDBORDERS = 0x0400,
        RBS_FIXEDORDER = 0x0800,
        RBS_REGISTERDROP = 0x1000,
        RBS_AUTOSIZE = 0x2000,
        RBS_VERTICALGRIPPER = 0x4000, // this always has the vertical gripper (default for horizontal mode)
        RBS_DBLCLKTOGGLE = 0x8000,
    } ReBar_Styles;

    typedef enum ReBar_BandStyles {
        RBBS_BREAK = 0x00000001, // break to new line
        RBBS_FIXEDSIZE = 0x00000002, // band can't be sized
        RBBS_CHILDEDGE = 0x00000004, // edge around top & bottom of child window
        RBBS_HIDDEN = 0x00000008, // don't show
        RBBS_NOVERT = 0x00000010, // don't show when vertical
        RBBS_FIXEDBMP = 0x00000020, // bitmap doesn't move during band resize
        RBBS_VARIABLEHEIGHT = 0x00000040, // allow autosizing of this child vertically
        RBBS_GRIPPERALWAYS = 0x00000080, // always show the gripper
        RBBS_NOGRIPPER = 0x00000100, // never show the gripper
        RBBS_USECHEVRON = 0x00000200,
        RBBS_HIDETITLE = 0x00000400,
    } ReBarBandStyles;

    typedef enum ReBar_BandInfoMask {
        RBBIM_STYLE = 0x00000001,
        RBBIM_COLORS = 0x00000002,
        RBBIM_TEXT = 0x00000004,
        RBBIM_IMAGE = 0x00000008,
        RBBIM_CHILD = 0x00000010,
        RBBIM_CHILDSIZE = 0x00000020,
        RBBIM_SIZE = 0x00000040,
        RBBIM_BACKGROUND = 0x00000080,
        RBBIM_ID = 0x00000100,
        RBBIM_IDEALSIZE = 0x00000200,
        RBBIM_LPARAM = 0x00000400,
        RBBIM_HEADERSIZE = 0x00000800, // control the size of the header
    } ReBarBandInfoMask;

    typedef enum ReBar_Messages {
        RB_INSERTBANDA = (1024 + 1), //(WM_USER + 1) and so on 1024
        RB_DELETEBAND = (1024 +  2),
        RB_GETBARINFO = (1024 +  3),
        RB_SETBARINFO = (1024 +  4),
        RB_GETBANDINFO = (1024 +  5),
        RB_SETBANDINFOA = (1024 +  6),
        RB_SETPARENT = (1024 +  7),
        RB_HITTEST = (1024 +  8),
        RB_GETRECT = (1024 +  9),
        RB_INSERTBANDW = (1024 +  10),
        RB_SETBANDINFOW = (1024 +  11),
        RB_GETBANDCOUNT = (1024 +  12),
        RB_GETROWCOUNT = (1024 +  13),
        RB_GETROWHEIGHT = (1024 +  14),
        RB_IDTOINDEX = (1024 +  16), // wParam = = id
        RB_GETTOOLTIPS = (1024 +  17),
        RB_SETTOOLTIPS = (1024 +  18),
        RB_SETBKCOLOR = (1024 +  19), // sets the default BK color
        RB_GETBKCOLOR = (1024 +  20), // defaults to CLR_NONE
        RB_SETTEXTCOLOR = (1024 +  21),
        RB_GETTEXTCOLOR = (1024 +  22), // defaults to 0x00000000
        RB_SIZETORECT = (1024 +  23), // resize the rebar/break bands and such to this rect (lparam)
        RB_SETCOLORSCHEME = 0x2002, // lParama is color scheme
        RB_BEGINDRAG = (1024 + 24),
        RB_ENDDRAG = (1024 + 25),
        RB_DRAGMOVE = (1024 + 26),
        RB_GETBARHEIGHT = (1024 + 27),
        RB_GETBANDINFOW = (1024 + 28),
        RB_GETBANDINFOA = (1024 + 29),
        RB_MINIMIZEBAND = (1024 + 30),
        RB_MAXIMIZEBAND = (1024 + 31),
        RB_GETDROPTARGET = (1024 + 32),
        RB_GETBANDBORDERS = (1024 + 34), // returns in lparam = lprc the amount of edges added to band wparam
        RB_SHOWBAND = (1024 + 35), // show/hide band
        RB_SETPALETTE = (1024 + 37),
        RB_GETPALETTE = (1024 + 38),
        RB_MOVEBAND = (1024 + 39),
        RB_PUSHCHEVRON = (1024 + 43),
    } ReBarMessages;


    typedef enum ReBar_NotificationCodes {
        RBN_HEIGHTCHANGE = (-831 - 0),
        RBN_GETOBJECT = (-831 - 1),
        RBN_LAYOUTCHANGED = (-831 - 2),
        RBN_AUTOSIZE = (-831 - 3),
        RBN_BEGINDRAG = (-831 - 4),
        RBN_ENDDRAG = (-831 - 5),
        RBN_DELETINGBAND = (-831 - 6), // Uses NMREBAR
        RBN_DELETEDBAND = (-831 - 7), // Uses NMREBAR
        RBN_CHILDSIZE = (-831 - 8),
        RBN_CHEVRONPUSHED = (-831 - 10),
        RBN_MINMAX = (-831 - 21),
    } ReBar_NotificationCodes;


    typedef enum ReBar_HitTestFlags {
        RBHT_NOWHERE = 0x0001,
        RBHT_CAPTION = 0x0002,
        RBHT_CLIENT = 0x0003,
        RBHT_GRABBER = 0x0004,
        RBHT_CHEVRON = 0x0008,
    } ReBarHitTestFlags;

    typedef enum Tooltip_Styles {
        TTS_ALWAYSTIP = 0x01,
        TTS_NOPREFIX = 0x02,
        TTS_NOANIMATE = 0x10,
        TTS_NOFADE = 0x20,
        TTS_BALLOON = 0x40,
    } Tooltip_Styles;


    typedef enum Tooltip_Flags {
        TTF_IDISHWND = 0x0001,

// Use this to center around trackpoint in trackmode
// -OR- to center around tool in normal mode.
// Use TTF_ABSOLUTE to place the tip exactly at the track coords when
// in tracking mode.  TTF_ABSOLUTE can be used in conjunction with TTF_CENTERTIP
// to center the tip absolutely about the track point.

        TTF_CENTERTIP = 0x0002,
        TTF_RTLREADING = 0x0004,
        TTF_SUBCLASS = 0x0010,
        TTF_TRACK = 0x0020,
        TTF_ABSOLUTE = 0x0080,
        TTF_TRANSPARENT = 0x0100,
        TTF_DI_SETITEM = 0x8000, // valid only on the TTN_NEEDTEXT callback
    } Tooltip_Flags;

    typedef enum Tooltip_DelayTimeFlags {
        TTDT_AUTOMATIC = 0,
        TTDT_RESHOW = 1,
        TTDT_AUTOPOP = 2,
        TTDT_INITIAL = 3,
    } Tooltip_DelayTimeFlags;

    typedef enum Tooltip_Messages {
        TTM_ACTIVATE = (1024 + 1),
        TTM_SETDELAYTIME = (1024 + 3),
        TTM_ADDTOOL = (1024 + 4),
        TTM_ADDTOOLW = (1024 + 50),
        TTM_DELTOOL = (1024 + 5),
        TTM_DELTOOLW = (1024 + 51),
        TTM_NEWTOOLRECT = (1024 + 6),
        TTM_NEWTOOLRECTW = (1024 + 52),
        TTM_RELAYEVENT = (1024 + 7),
        TTM_GETTOOLINFO = (1024 + 8),
        TTM_GETTOOLINFOW = (1024 + 53),
        TTM_SETTOOLINFO = (1024 + 9),
        TTM_SETTOOLINFOW = (1024 + 54),
        TTM_HITTEST = (1024 +10),
        TTM_HITTESTW = (1024 +55),
        TTM_GETTEXT = (1024 +11),
        TTM_GETTEXTW = (1024 +56),
        TTM_UPDATETIPTEXT = (1024 +12),
        TTM_UPDATETIPTEXTW = (1024 +57),
        TTM_GETTOOLCOUNT = (1024 +13),
        TTM_ENUMTOOLS = (1024 +14),
        TTM_ENUMTOOLSW = (1024 +58),
        TTM_GETCURRENTTOOL = (1024 + 15),
        TTM_GETCURRENTTOOLW = (1024 + 59),
        TTM_WINDOWFROMPOINT = (1024 + 16),
        TTM_TRACKACTIVATE = (1024 + 17), // wParam = TRUE/FALSE start end  lparam = LPTOOLINFO
        TTM_TRACKPOSITION = (1024 + 18), // lParam = dwPos
        TTM_SETTIPBKCOLOR = (1024 + 19),
        TTM_SETTIPTEXTCOLOR = (1024 + 20),
        TTM_GETDELAYTIME = (1024 + 21),
        TTM_GETTIPBKCOLOR = (1024 + 22),
        TTM_GETTIPTEXTCOLOR = (1024 + 23),
        TTM_SETMAXTIPWIDTH = (1024 + 24),
        TTM_GETMAXTIPWIDTH = (1024 + 25),
        TTM_SETMARGIN = (1024 + 26), // lParam = lprc
        TTM_GETMARGIN = (1024 + 27), // lParam = lprc
        TTM_POP = (1024 + 28),
        TTM_UPDATE = (1024 + 29),
        TTM_GETBUBBLESIZE = (1024 + 30),
        TTM_ADJUSTRECT = (1024 + 31),
        TTM_SETTITLE = (1024 + 32), // wParam = TTI_*, lParam = char* szTitle
        TTM_SETTITLEW = (1024 + 33), // wParam = TTI_*, lParam = wchar* szTitle
    } Tooltip_Messages;

    typedef enum Tooltip_Icon {
        TTI_NONE = 0,
        TTI_INFO = 1,
        TTI_WARNING = 2,
        TTI_ERROR = 3,
    } Tooltip_Icon;

// ********* Statusbar Control
    typedef enum StatusBar_Messages {
        SB_SETTEXT = (1024+1),
        SB_SETTEXTW = (1024+11),
        SB_GETTEXT = (1024+2),
        SB_GETTEXTW = (1024+13),
        SB_GETTEXTLENGTH = (1024+3),
        SB_GETTEXTLENGTHW = (1024+12),
        SB_SETPARTS = (1024+4),
        SB_GETPARTS = (1024+6),
        SB_GETBORDERS = (1024+7),
        SB_SETMINHEIGHT = (1024+8),
        SB_SIMPLE = (1024+9),
        SB_GETRECT = (1024+10),
        SB_ISSIMPLE = (1024+14),
        SB_SETICON = (1024+15),
        SB_SETTIPTEXT = (1024+16),
        SB_SETTIPTEXTW = (1024+17),
        SB_GETTIPTEXT = (1024+18),
        SB_GETTIPTEXTW = (1024+19),
        SB_SETBKCOLOR = 0x2001,
    } StatusBar_Messages;

    typedef enum StatusBar_Ts {
        SBT_OWNERDRAW = 0x1000,
        SBT_NOBORDERS = 0x0100,
        SBT_POPOUT = 0x0200,
        SBT_RTLREADING = 0x0400,
        SBT_TOOLTIPS = 0x0800,
        SBT_NOTABPARSING = 0x0800,
    } StatusBar_Ts;


    typedef enum StatusBar_NotificationCodes {
        SBN_SIMPLEMODECHANGE = (-880 - 0),
    } StatusBar_NotificationCodes;

// ******* Trackbar COntrol

    typedef enum TrackBarStyles {
        TBS_AUTOTICKS = 0x0001,
        TBS_VERT = 0x0002,
        TBS_HORZ = 0x0000,
        TBS_TOP = 0x0004,
        TBS_BOTTOM = 0x0000,
        TBS_LEFT = 0x0004,
        TBS_RIGHT = 0x0000,
        TBS_BOTH = 0x0008,
        TBS_NOTICKS = 0x0010,
        TBS_ENABLESELRANGE = 0x0020,
        TBS_FIXEDLENGTH = 0x0040,
        TBS_NOTHUMB = 0x0080,
        TBS_TOOLTIPS = 0x0100,
        TBS_REVERSED = 0x0200,
    } TrackBarStyles;

    typedef enum TrackbarMessages {
        TBM_GETPOS = (1024),
        TBM_GETRANGEMIN = (1024+1),
        TBM_GETRANGEMAX = (1024+2),
        TBM_GETTIC = (1024+3),
        TBM_SETTIC = (1024+4),
        TBM_SETPOS = (1024+5),
        TBM_SETRANGE = (1024+6),
        TBM_SETRANGEMIN = (1024+7),
        TBM_SETRANGEMAX = (1024+8),
        TBM_CLEARTICS = (1024+9),
        TBM_SETSEL = (1024+10),
        TBM_SETSELSTART = (1024+11),
        TBM_SETSELEND = (1024+12),
        TBM_GETPTICS = (1024+14),
        TBM_GETTICPOS = (1024+15),
        TBM_GETNUMTICS = (1024+16),
        TBM_GETSELSTART = (1024+17),
        TBM_GETSELEND = (1024+18),
        TBM_CLEARSEL = (1024+19),
        TBM_SETTICFREQ = (1024+20),
        TBM_SETPAGESIZE = (1024+21),
        TBM_GETPAGESIZE = (1024+22),
        TBM_SETLINESIZE = (1024+23),
        TBM_GETLINESIZE = (1024+24),
        TBM_GETTHUMBRECT = (1024+25),
        TBM_GETCHANNELRECT = (1024+26),
        TBM_SETTHUMBLENGTH = (1024+27),
        TBM_GETTHUMBLENGTH = (1024+28),
        TBM_SETTOOLTIPS = (1024+29),
        TBM_GETTOOLTIPS = (1024+30),
        TBM_SETTIPSIDE = (1024+31),
        TBM_SETBUDDY = (1024+32), // wparam = BOOL fLeft; (or right)
        TBM_GETBUDDY = (1024+33), // wparam = BOOL fLeft; (or right)
    } TrackbarMessages;

    typedef enum TrackBarTipSideFlags {
        TBTS_TOP = 0,
        TBTS_LEFT = 1,
        TBTS_BOTTOM = 2,
        TBTS_RIGHT = 3,
    } TrackBarTipSideFlags;

    typedef enum TrackBarNotifications {
        TB_LINEUP = 0,
        TB_LINEDOWN = 1,
        TB_PAGEUP = 2,
        TB_PAGEDOWN = 3,
        TB_THUMBPOSITION = 4,
        TB_THUMBTRACK = 5,
        TB_TOP = 6,
        TB_BOTTOM = 7,
        TB_ENDTRACK = 8,
    } TrackBarNotifications;

    typedef enum TrackBarItemSpecs {
        TBCD_TICS = 0x0001,
        TBCD_THUMB = 0x0002,
        TBCD_CHANNEL = 0x0003,
    } TrackBarItemSpecs;

    typedef enum DragListMessages {
        DL_BEGINDRAG = (1024+133),
        DL_DRAGGING = (1024+134),
        DL_DROPPED = (1024+135),
        DL_CANCELDRAG = (1024+136),
    } DragListMessages;

    typedef enum DragListCursor {
        DL_CURSORSET = 0,
        DL_STOPCURSOR = 1,
        DL_COPYCURSOR = 2,
        DL_MOVECURSOR = 3,
    } DragListCursor;


    typedef enum UpDownStyles {
        UDS_WRAP = 0x0001,
        UDS_SETBUDDYINT = 0x0002,
        UDS_ALIGNRIGHT = 0x0004,
        UDS_ALIGNLEFT = 0x0008,
        UDS_AUTOBUDDY = 0x0010,
        UDS_ARROWKEYS = 0x0020,
        UDS_HORZ = 0x0040,
        UDS_NOTHOUSANDS = 0x0080,
        UDS_HOTTRACK = 0x0100,
    } UpDownStyles;

    typedef enum UpDownMessages {
        UDM_SETRANGE = (1024+101),
        UDM_GETRANGE = (1024+102),
        UDM_SETPOS = (1024+103),
        UDM_GETPOS = (1024+104),
        UDM_SETBUDDY = (1024+105),
        UDM_GETBUDDY = (1024+106),
        UDM_SETACCEL = (1024+107),
        UDM_GETACCEL = (1024+108),
        UDM_SETBASE = (1024+109),
        UDM_GETBASE = (1024+110),
        UDM_SETRANGE32 = (1024+111),
        UDM_GETRANGE32 = (1024+112), // wParam & lParam are LPINT
        UDM_SETPOS32 = (1024+113),
        UDM_GETPOS32 = (1024+114),
    } UpDownMessages;


    typedef enum ProgressStyles {
        PBS_SMOOTH		= 0x01,
        PBS_VERTICAL	= 0x04,
		PBS_MARQUEE		= 0x08,
    } ProgressStyles;

    typedef enum ProgressMessages {
        PBM_SETRANGE = (1024+1),
        PBM_SETPOS = (1024+2),
        PBM_DELTAPOS = (1024+3),
        PBM_SETSTEP = (1024+4),
        PBM_STEPIT = (1024+5),
		PBM_SETBARCOLOR = (1024+9),
		PBM_SETBKCOLOR  = (8192 + 1), // 0x2001
        PBM_SETRANGE32 = (1024+6), // lParam = high, wParam = low
        PBM_GETRANGE = (1024+7), // wParam = return (TRUE ? low : high).
        PBM_GETPOS = (1024+8),
    } ProgressMessages;

// ********** HOTKEY CONTROL


    typedef enum HotKeyFlags {
        HOTKEYF_SHIFT = 0x01,
        HOTKEYF_CONTROL = 0x02,
        HOTKEYF_ALT = 0x04,
        HOTKEYF_EXT = 0x08,
    } HotKeyFlags;

    typedef enum HotKeyRulesCombinations {
        HKCOMB_NONE = 0x0001,
        HKCOMB_S = 0x0002,
        HKCOMB_C = 0x0004,
        HKCOMB_A = 0x0008,
        HKCOMB_SC = 0x0010,
        HKCOMB_SA = 0x0020,
        HKCOMB_CA = 0x0040,
        HKCOMB_SCA = 0x0080,
    } HotKeyRulesCombinations;

    typedef enum HotKeyMessages {
        HKM_SETHOTKEY = (1024+1),
        HKM_GETHOTKEY = (1024+2),
        HKM_SETRULES = (1024+3),
    } HotKeyMessages;


// = = ====LISTVIEW CONTROL = = =====


    typedef enum ListViewStyles {
        LVS_ICON = 0x0000,
        LVS_REPORT = 0x0001,
        LVS_SMALLICON = 0x0002,
        LVS_LIST = 0x0003,
		LVS_TILE = 0x0004, //New XP style
        LVS_TYPEMASK = 0x0003,
        LVS_SINGLESEL = 0x0004,
        LVS_SHOWSELALWAYS = 0x0008,
        LVS_SORTASCENDING = 0x0010,
        LVS_SORTDESCENDING = 0x0020,
        LVS_SHAREIMAGELISTS = 0x0040,
        LVS_NOLABELWRAP = 0x0080,
        LVS_AUTOARRANGE = 0x0100,
        LVS_EDITLABELS = 0x0200,
        LVS_OWNERDATA = 0x1000,
        LVS_NOSCROLL = 0x2000,
        LVS_TYPESTYLEMASK = 0xfc00,
        LVS_ALIGNTOP = 0x0000,
        LVS_ALIGNLEFT = 0x0800,
        LVS_ALIGNMASK = 0x0c00,
        LVS_OWNERDRAWFIXED = 0x0400,
        LVS_NOCOLUMNHEADER = 0x4000,
        LVS_NOSORTHEADER = 0x8000,
    } ListViewStyles;

typedef enum ListViewImageListStyle {
    LVSIL_NORMAL = 0,
    LVSIL_SMALL = 1,
    LVSIL_STATE = 2,
} ListViewImageListStyle;

typedef enum ListViewItemFlags {
    LVIF_TEXT = 0x0001,
    LVIF_IMAGE = 0x0002,
    LVIF_PARAM = 0x0004,
    LVIF_STATE = 0x0008,
    LVIF_INDENT = 0x0010,
    LVIF_NORECOMPUTE = 0x0800,
    LVIF_DI_SETITEM = 0x1000,
	LVIF_COLUMNS = 0x200, //520
	LVIF_GROUPID = 0x100,
} ListViewItemFlag;

typedef enum ListViewItemStates {
    LVIS_FOCUSED = 0x0001,
    LVIS_SELECTED = 0x0002,
    LVIS_CUT = 0x0004,
    LVIS_DROPHILITED = 0x0008,
    LVIS_OVERLAYMASK = 0x0F00,
    LVIS_STATEIMAGEMASK = 0xF000,
} ListViewItemStates;


typedef enum ListViewImages {
    I_IMAGECALLBACK = (-1),
    I_IMAGENONE = (-2) ,
} ListViewCallbacks;

typedef enum ListViewNextItemFlags {
    LVNI_ALL = 0x0000,
    LVNI_FOCUSED = 0x0001,
    LVNI_SELECTED = 0x0002,
    LVNI_CUT = 0x0004,
    LVNI_DROPHILITED = 0x0008,
    LVNI_ABOVE = 0x0100,
    LVNI_BELOW = 0x0200,
    LVNI_TOLEFT = 0x0400,
    LVNI_TORIGHT = 0x0800,
} ListViewNextItemFlags;

typedef enum ListViewFindIndoFlags {
    LVFI_PARAM = 0x0001,
    LVFI_STRING = 0x0002,
    LVFI_PARTIAL = 0x0008,
    LVFI_WRAP = 0x0020,
    LVFI_NEARESTXY = 0x0040,
} ListViewFindIndoFlags;


typedef enum ListViewItemRectFlags {
    LVIR_BOUNDS = 0,
    LVIR_ICON = 1,
    LVIR_LABEL = 2,
    LVIR_SELECTBOUNDS = 3,
} ListViewItemRectFlags;

typedef enum ListViewHitTest {
    LVHT_NOWHERE = 0x0001,
    LVHT_ONITEMICON = 0x0002,
    LVHT_ONITEMLABEL = 0x0004,
    LVHT_ONITEMSTATEICON = 0x0008,
    LVHT_ONITEM = 0x000F, //(LVHT_ONITEMICON | LVHT_ONITEMLABEL | LVHT_ONITEMSTATEICON)
	LVHT_ABOVE = 0x0008,
	LVHT_BELOW = 0x0010,
	LVHT_TORIGHT = 0x0020,
	LVHT_TOLEFT = 0x0040,
} ListViewHitTest;


typedef enum ListViewArrangeFlags {
    LVA_DEFAULT = 0x0000,
    LVA_ALIGNLEFT = 0x0001,
    LVA_ALIGNTOP = 0x0002,
    LVA_SNAPTOGRID = 0x0005,
} ListViewArrangeFlags;

typedef enum ListViewColumnMasks {
    LVCF_FMT = 0x0001,
    LVCF_WIDTH = 0x0002,
    LVCF_TEXT = 0x0004,
    LVCF_SUBITEM = 0x0008,
    LVCF_IMAGE = 0x0010,
    LVCF_ORDER = 0x0020,
} ListViewColumnMasks;

typedef enum ListViewColumnFormats {
    LVCFMT_LEFT = 0x0000,
    LVCFMT_RIGHT = 0x0001,
    LVCFMT_CENTER = 0x0002,
    LVCFMT_JUSTIFYMASK = 0x0003,
    LVCFMT_IMAGE = 0x0800,
    LVCFMT_BITMAP_ON_RIGHT = 0x1000,
    LVCFMT_COL_HAS_IMAGES = 0x8000,
} ListViewColumnFormats;


typedef enum ListViewColumnWidth {
    LVSCW_AUTOSIZE = -1,
    LVSCW_AUTOSIZE_USEHEADER = -2,
} ListViewColumnWidth;

typedef enum ListViewItemCountFlags {
    LVSICF_NOINVALIDATEALL = 0x00000001,
    LVSICF_NOSCROLL = 0x00000002,
} ListViewItemCountFlags;

typedef enum ListViewExtendedStyles {
    LVS_EX_GRIDLINES = 0x00000001,
    LVS_EX_SUBITEMIMAGES = 0x00000002,
    LVS_EX_CHECKBOXES = 0x00000004,
    LVS_EX_TRACKSELECT = 0x00000008,
    LVS_EX_HEADERDRAGDROP = 0x00000010,
    LVS_EX_FULLROWSELECT = 0x00000020, // applies to report mode only
    LVS_EX_ONECLICKACTIVATE = 0x00000040,
    LVS_EX_TWOCLICKACTIVATE = 0x00000080,
    LVS_EX_FLATSB = 0x00000100, // cannot be cleared
    LVS_EX_REGIONAL = 0x00000200,
    LVS_EX_INFOTIP = 0x00000400, // listview does InfoTips for you
    LVS_EX_UNDERLINEHOT = 0x00000800,
    LVS_EX_UNDERLINECOLD = 0x00001000,
    LVS_EX_MULTIWORKAREAS = 0x00002000,
    LVS_EX_LABELTIP = 0x00004000,
    LVS_EX_BORDERSELECT = 0x8000,
    LVS_EX_DOUBLEBUFFER = 0x10000, //&H10000
    LVS_EX_HIDELABELS = 0x20000,
    LVS_EX_SINGLEROW = 0x40000,
    LVS_EX_SNAPTOGRID = 0x80000,
} ListViewExtendedStyles;

typedef enum ListViewBkImageFlags {
    LVBKIF_SOURCE_NONE = 0x00000000,
    LVBKIF_SOURCE_LONG = 0x00000001,
    LVBKIF_SOURCE_URL = 0x00000002,
    LVBKIF_SOURCE_MASK = 0x00000003,
    LVBKIF_STYLE_NORMAL = 0x00000000,
    LVBKIF_STYLE_TILE = 0x00000010,
    LVBKIF_STYLE_MASK = 0x00000010,
} ListViewBkImageFlags;


typedef enum ListViewMessages {
	LVM_FIRST = (0x1000), //4096
    LVM_GETBKCOLOR = (4096 + 0),
    LVM_SETBKCOLOR = (4096 + 1),
    LVM_GETIMAGELIST = (4096 + 2),
    LVM_SETIMAGELIST = (4096 + 3),
    LVM_GETITEMCOUNT = (4096 + 4),
    LVM_GETITEM = (4096 + 5),
    LVM_GETITEMW = (4096 + 75),
    LVM_SETITEM = (4096 + 6),
    LVM_SETITEMW = (4096 + 76),
    LVM_INSERTITEM = (4096 + 7),
    LVM_INSERTITEMW = (4096 + 77),
    LVM_DELETEITEM = (4096 + 8),
    LVM_DELETEALLITEMS = (4096 + 9),
    LVM_GETCALLBACKMASK = (4096 + 10),
    LVM_SETCALLBACKMASK = (4096 + 11),
    LVM_GETNEXTITEM = (4096 + 12),
    LVM_FINDITEM = (4096 + 13),

    LVM_FINDITEMW = (4096 + 83),
    LVM_GETITEMRECT = (4096 + 14),
    LVM_SETITEMPOSITION = (4096 + 15),
    LVM_GETITEMPOSITION = (4096 + 16),
    LVM_GETSTRINGWIDTH = (4096 + 17),
    LVM_GETSTRINGWIDTHW = (4096 + 87),
    LVM_HITTEST = (4096 + 18),
    LVM_ENSUREVISIBLE = (4096 + 19),
    LVM_SCROLL = (4096 + 20),
    LVM_REDRAWITEMS = (4096 + 21),
    LVM_ARRANGE = (4096 + 22),
    LVM_EDITLABEL = (4096 + 23),
    LVM_EDITLABELW = (4096 + 118),
    LVM_GETEDITCONTROL = (4096 + 24),
    LVM_GETCOLUMN = (4096 + 25),
    LVM_GETCOLUMNW = (4096 + 95),
    LVM_SETCOLUMN = (4096 + 26),
    LVM_SETCOLUMNW = (4096 + 96),
    LVM_INSERTCOLUMN = (4096 + 27),
    LVM_INSERTCOLUMNW = (4096 + 97),
    LVM_DELETECOLUMN = (4096 + 28),
    LVM_GETCOLUMNWIDTH = (4096 + 29),
    LVM_SETCOLUMNWIDTH = (4096 + 30),
    LVM_GETHEADER = (4096 + 31),
    LVM_CREATEDRAGIMAGE = (4096 + 33),
    LVM_GETVIEWRECT = (4096 + 34),
    LVM_GETTEXTCOLOR = (4096 + 35),
    LVM_SETTEXTCOLOR = (4096 + 36),
    LVM_GETTEXTBKCOLOR = (4096 + 37),
    LVM_SETTEXTBKCOLOR = (4096 + 38),
    LVM_GETTOPINDEX = (4096 + 39),
    LVM_GETCOUNTPERPAGE = (4096 + 40),
    LVM_GETORIGIN = (4096 + 41),
    LVM_UPDATE = (4096 + 42),
    LVM_SETITEMSTATE = (4096 + 43),
    LVM_GETITEMSTATE = (4096 + 44),
    LVM_GETITEMTEXT = (4096 + 45),
    LVM_GETITEMTEXTW = (4096 + 115),
    LVM_SETITEMTEXT = (4096 + 46),
    LVM_SETITEMTEXTW = (4096 + 116),
    LVM_SETITEMCOUNT = (4096 + 47),
    LVM_SORTITEMS = (4096 + 48),
    LVM_SETITEMPOSITION32 = (4096 + 49),
    LVM_GETSELECTEDCOUNT = (4096 + 50),
    LVM_GETITEMSPACING = (4096 + 51),
    LVM_GETISEARCHSTRING = (4096 + 52),
    LVM_GETISEARCHSTRINGW = (4096 + 117),
    LVM_SETICONSPACING = (4096 + 53),
    LVM_SETEXTENDEDLISTVIEWSTYLE = (4096 + 54),
    LVM_GETEXTENDEDLISTVIEWSTYLE = (4096 + 55),
    LVM_GETSUBITEMRECT = (4096 + 56),
    LVM_SUBITEMHITTEST = (4096 + 57),
    LVM_SETCOLUMNORDERARRAY = (4096 + 58),
    LVM_GETCOLUMNORDERARRAY = (4096 + 59),
    LVM_SETHOTITEM = (4096 + 60),
    LVM_GETHOTITEM = (4096 + 61),
    LVM_SETHOTCURSOR = (4096 + 62),
    LVM_GETHOTCURSOR = (4096 + 63),
    LVM_APPROXIMATEVIEWRECT = (4096 + 64),
    LVM_SETWORKAREAS = (4096 + 65),
    LVM_GETWORKAREAS = (4096 + 70),
    LVM_GETNUMBEROFWORKAREAS = (4096 + 73),
    LVM_GETSELECTIONMARK = (4096 + 66),
    LVM_SETSELECTIONMARK = (4096 + 67),
    LVM_SETHOVERTIME = (4096 + 71),
    LVM_GETHOVERTIME = (4096 + 72),
    LVM_SETBKIMAGE = (4096 + 68),
    LVM_SETBKIMAGEW = (4096 + 138),
    LVM_GETBKIMAGE = (4096 + 69),
    LVM_GETBKIMAGEW = (4096 + 139),
	LVM_SETUNICODEFORMAT = (8197),
	LVM_GETUNICODEFORMAT = (8198),
	LVM_SETTOOLTIPS      = (4096 + 74),
	LVM_GETTOOLTIPS      = (4096 + 78),
} ListViewMessages;

typedef enum ListViewKeyFlags {
    LVKF_ALT = 0x0001,
    LVKF_CONTROL = 0x0002,
    LVKF_SHIFT = 0x0004,
} ListViewKeyFlags;

typedef enum ListViewNotificationCodes {
    LVN_ITEMCHANGING = (-100-0),
    LVN_ITEMCHANGED = (-100-1),
    LVN_INSERTITEM = (-100-2),
    LVN_DELETEITEM = (-100-3),
    LVN_DELETEALLITEMS = (-100-4),
    LVN_BEGINLABELEDIT = (-100-5),
    LVN_BEGINLABELEDITW = (-100-75),
    LVN_ENDLABELEDIT = (-100-6),
    LVN_ENDLABELEDITW = (-100-76),
    LVN_COLUMNCLICK = (-100-8),
    LVN_BEGINDRAG = (-100-9),
    LVN_BEGINRDRAG = (-100-11),
    LVN_ODCACHEHINT = (-100-13),
    LVN_ODFINDITEM = (-100-52),
    LVN_ODFINDITEMW = (-100-79),
    LVN_ITEMACTIVATE = (-100-14),
    LVN_ODSTATECHANGED = (-100-15),
    LVN_HOTTRACK = (-100-21),
    LVN_GETDISPINFO = (-100-50),
    LVN_GETDISPINFOW = (-100-77),
    LVN_SETDISPINFO = (-100-51),
    LVN_SETDISPINFOW = (-100-78),
    LVN_KEYDOWN = (-100-55),
    LVN_MARQUEEBEGIN = (-100-56),
    LVN_GETINFOTIP = (-100-57),
    LVN_GETINFOTIPW = (-100-58),
} ListViewNotificationCodes;

typedef enum ListViewGetInfoTipFlags {
    LVGIT_UNFOLDED = 0x0001,
} ListViewGetInfoTipFlags;

typedef enum TreeView_Styles {
    TVS_HASBUTTONS = 0x0001,
    TVS_HASLINES = 0x0002,
    TVS_LINESATROOT = 0x0004,
    TVS_EDITLABELS = 0x0008,
    TVS_DISABLEDRAGDROP = 0x0010,
    TVS_SHOWSELALWAYS = 0x0020,
    TVS_RTLREADING = 0x0040,
    TVS_NOTOOLTIPS = 0x0080,
    TVS_CHECKBOXES = 0x0100,
    TVS_TRACKSELECT = 0x0200,
    TVS_SINGLEEXPAND = 0x0400,
    TVS_INFOTIP = 0x0800,
    TVS_FULLROWSELECT = 0x1000,
    TVS_NOSCROLL = 0x2000,
    TVS_NONEVENHEIGHT = 0x4000,
    TVS_NOHSCROLL = 0x8000,
} TreeView_Styles;

typedef enum TreeView_ItemFlags {
    TVIF_TEXT = 0x0001,
    TVIF_IMAGE = 0x0002,
    TVIF_PARAM = 0x0004,
    TVIF_STATE = 0x0008,
    TVIF_HANDLE = 0x0010,
    TVIF_SELECTEDIMAGE = 0x0020,
    TVIF_CHILDREN = 0x0040,
    TVIF_INTEGRAL = 0x0080,
    TVIF_DI_SETITEM = 0x1000,
} TreeView_ItemFlags;

typedef enum TreeView_ItemStates {
    TVIS_SELECTED = 0x0002,
    TVIS_CUT = 0x0004,
    TVIS_DROPHILITED = 0x0008,
    TVIS_BOLD = 0x0010,
    TVIS_EXPANDED = 0x0020,
    TVIS_EXPANDEDONCE = 0x0040,
    TVIS_EXPANDPARTIAL = 0x0080,
    TVIS_OVERLAYMASK = 0x0F00,
    TVIS_STATEIMAGEMASK = 0xF000,
    TVIS_USERMASK = 0xF000,
} TreeView_ItemStates;

typedef enum TreeViewSpecialItems {
    TVI_ROOT = 0xFFFF0000,
    TVI_FIRST = 0xFFFF0001,
    TVI_LAST = 0xFFFF0002,
    TVI_SORT = 0xFFFF0003,
} TreeViewSpecialItems;

typedef enum TreeViewExpandFlags {
    TVE_COLLAPSE = 0x0001,
    TVE_EXPAND = 0x0002,
    TVE_TOGGLE = 0x0003,
    TVE_EXPANDPARTIAL = 0x4000,
    TVE_COLLAPSERESET = 0x8000,
} TreeViewExpandFlags;

typedef enum TreeViewImageLists {
    TVSIL_NORMAL = 0,
    TVSIL_STATE = 2,
} TreeViewImageLists;

typedef enum TreeViewGetNextFlags {
    TVGN_ROOT = 0x0000,
    TVGN_NEXT = 0x0001,
    TVGN_PREVIOUS = 0x0002,
    TVGN_PARENT = 0x0003,
    TVGN_CHILD = 0x0004,
    TVGN_FIRSTVISIBLE = 0x0005,
    TVGN_NEXTVISIBLE = 0x0006,
    TVGN_PREVIOUSVISIBLE = 0x0007,
    TVGN_DROPHILITE = 0x0008,
    TVGN_CARET = 0x0009,
    TVGN_LASTVISIBLE = 0x000A,
} TreeViewGetNextFlags;

typedef enum TreeView_HitTest {
    TVHT_NOWHERE = 0x0001,
    TVHT_ONITEMICON = 0x0002,
    TVHT_ONITEMLABEL = 0x0004,
	[helpstring("TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON")]
    TVHT_ONITEM = 0x0007,
    TVHT_ONITEMINDENT = 0x0008,
    TVHT_ONITEMBUTTON = 0x0010,
    TVHT_ONITEMRIGHT = 0x0020,
    TVHT_ONITEMSTATEICON = 0x0040,
    TVHT_ABOVE = 0x0100,
    TVHT_BELOW = 0x0200,
    TVHT_TORIGHT = 0x0400,
    TVHT_TOLEFT = 0x0800,
} TreeViewHitTest;

typedef enum TreeViewSetGetBorderFlags {
    TVSBF_XBORDER = 0x00000001,
    TVSBF_YBORDER = 0x00000002,
} TreeViewSetGetBorderFlags;

//TV_FIRST = 4352 &H1100 0x1100
typedef enum TreeView_Messages {
    TVM_INSERTITEM = (4352 + 0),
    TVM_INSERTITEMW = (4352 + 50),
    TVM_DELETEITEM = (4352 + 1),
    TVM_EXPAND = (4352 + 2),
    TVM_GETITEMRECT = (4352 + 4),
    TVM_GETCOUNT = (4352 + 5),
    TVM_GETINDENT = (4352 + 6),
    TVM_SETINDENT = (4352 + 7),
    TVM_GETIMAGELIST = (4352 + 8),
    TVM_SETIMAGELIST = (4352 + 9),
    TVM_GETNEXTITEM = (4352 + 10),
    TVM_SELECTITEM = (4352 + 11),
    TVM_GETITEM = (4352 + 12),
    TVM_GETITEMW = (4352 + 62),
    TVM_SETITEM = (4352 + 13),
    TVM_SETITEMW = (4352 + 63),
    TVM_EDITLABEL = (4352 + 14),
    TVM_EDITLABELW = (4352 + 65),
    TVM_GETEDITCONTROL = (4352 + 15),
    TVM_GETVISIBLECOUNT = (4352 + 16),
    TVM_HITTEST = (4352 + 17),
    TVM_CREATEDRAGIMAGE = (4352 + 18),
    TVM_SORTCHILDREN = (4352 + 19),
    TVM_ENSUREVISIBLE = (4352 + 20),
    TVM_SORTCHILDRENCB = (4352 + 21),
    TVM_ENDEDITLABELNOW = (4352 + 22),
    TVM_GETISEARCHSTRING = (4352 + 23),
    TVM_GETISEARCHSTRINGW = (4352 + 64),
    TVM_SETTOOLTIPS = (4352 + 24),
    TVM_GETTOOLTIPS = (4352 + 25),
    TVM_SETINSERTMARK = (4352 + 26),
    TVM_SETITEMHEIGHT = (4352 + 27),
    TVM_GETITEMHEIGHT = (4352 + 28),
    TVM_SETBKCOLOR = (4352 + 29),
    TVM_SETTEXTCOLOR = (4352 + 30),
    TVM_GETBKCOLOR = (4352 + 31),
    TVM_GETTEXTCOLOR = (4352 + 32),
    TVM_SETSCROLLTIME = (4352 + 33),
    TVM_GETSCROLLTIME = (4352 + 34),
    TVM_SETBORDER = (4352 + 35),
    TVM_GETBORDER = (4352 + 36),
    TVM_GETITEMSTATE = (4352 + 39),
    TVM_SETLINECOLOR = (4352 + 40),
    TVM_GETLINECOLOR = (4352 + 41),
	TVM_SETINSERTMARKCOLOR = (4352 + 37),
	TVM_GETINSERTMARKCOLOR = (4352 + 38),
    TVM_SETUNICODEFORMAT = (8197),
    TVM_GETUNICODEFORMAT = (8198),
} TreeViewMessages;

typedef enum TreeView_ChangedFlags {
    TVC_UNKNOWN = 0x0000,
    TVC_BYMOUSE = 0x0001,
    TVC_BYKEYBOARD = 0x0002,
} TreeViewChangedFlags;

typedef enum TreeViewNotificationCodes {

    TVN_SELCHANGING = (-400-1),
    TVN_SELCHANGINGW = (-400-50),
    TVN_SELCHANGED = (-400-2),
    TVN_SELCHANGEDW = (-400-51),
    TVN_GETDISPINFO = (-400-3),
    TVN_GETDISPINFOW = (-400-52),
    TVN_SETDISPINFO = (-400-4),
    TVN_SETDISPINFOW = (-400-53),
    TVN_ITEMEXPANDING = (-400-5),
    TVN_ITEMEXPANDINGW = (-400-54),
    TVN_ITEMEXPANDED = (-400-6),
    TVN_ITEMEXPANDEDW = (-400-55),
    TVN_BEGINDRAG = (-400-7),
    TVN_BEGINDRAGW = (-400-56),
    TVN_BEGINRDRAG = (-400-8),
    TVN_BEGINRDRAGW = (-400-57),
    TVN_DELETEITEM = (-400-9),
    TVN_DELETEITEMW = (-400-58),
    TVN_BEGINLABELEDIT = (-400-10),
    TVN_BEGINLABELEDITW = (-400-59),
    TVN_ENDLABELEDIT = (-400-11),
    TVN_ENDLABELEDITW = (-400-60),
    TVN_KEYDOWN = (-400-12),
    TVN_GETINFOTIP = (-400-13),
    TVN_GETINFOTIPW = (-400-14),
    TVN_SINGLEEXPAND = (-400-15),
} TreeViewNotificationCodes;

typedef enum TreeViewNotificationReturnValues {
    TVNRET_DEFAULT = 0,
    TVNRET_SKIPOLD = 1,
    TVNRET_SKIPNEW = 2,
    TVCDRF_NOIMAGES = 0x00010000,
 } TreeViewNotificationReturnValues;

//WM_USER 1024

    typedef enum ComboBoxExItemFlags {
        CBEIF_TEXT = 0x00000001,
        CBEIF_IMAGE = 0x00000002,
        CBEIF_SELECTEDIMAGE = 0x00000004,
        CBEIF_OVERLAY = 0x00000008,
        CBEIF_INDENT = 0x00000010,
        CBEIF_LPARAM = 0x00000020,
        CBEIF_DI_SETITEM = 0x10000000,
    } ComboBoxExItemFlags;

    typedef enum ComboBoxExMessages {
        CBEM_INSERTITEM = (1024 + 1),
        CBEM_SETIMAGELIST = (1024 + 2),
        CBEM_GETIMAGELIST = (1024 + 3),
        CBEM_GETITEM = (1024 + 4),
        CBEM_SETITEM = (1024 + 5),
        CBEM_DELETEITEM = 0x144,
        CBEM_GETCOMBOCONTROL = (1024 + 6),
        CBEM_GETEDITCONTROL = (1024 + 7),
        CBEM_SETEXSTYLE = (1024 + 8),
        CBEM_SETEXTENDEDSTYLE = (1024 + 14),
        CBEM_GETEXTENDEDSTYLE = (1024 + 9),
        CBEM_HASEDITCHANGED = (1024 + 10),
        CBEM_INSERTITEMW = (1024 + 11),
        CBEM_SETITEMW = (1024 + 12),
        CBEM_GETITEMW = (1024 + 13),
    } ComboBoxExMessages;

    typedef enum ComboBoxExExtendedStyles {
        CBES_EX_NOEDITIMAGE = 0x00000001,
        CBES_EX_NOEDITIMAGEINDENT = 0x00000002,
        CBES_EX_PATHWORDBREAKPROC = 0x00000004,
        CBES_EX_NOSIZELIMIT = 0x00000008,
        CBES_EX_CASESENSITIVE = 0x00000010,
    } ComboBoxExExtendedStyles;

    typedef enum ComboBoxExNotificationCodes {
        CBEN_GETDISPINFO = (-800 - 0),
        CBEN_GETDISPINFOW = (-800 - 7),
        CBEN_INSERTITEM = (-800 - 1),
        CBEN_DELETEITEM = (-800 - 2),
        CBEN_BEGINEDIT = (-800 - 4),
        CBEN_ENDEDIT = (-800 - 5),
        CBEN_ENDEDITW = (-800 - 6),
        CBEN_DRAGBEGIN = (-800 - 8),
        CBEN_DRAGBEGINW = (-800 - 9),
    } ComboBoxExNotificationCodes;

    typedef enum ComboBoxExEndEditFlags {
        CBENF_KILLFOCUS = 1,
        CBENF_RETURN = 2,
        CBENF_ESCAPE = 3,
        CBENF_DROPDOWN = 4,
    } ComboBoxExEndEditFlags;

   typedef enum TabStyles {
        TCS_SCROLLOPPOSITE = 0x0001, // assumes multiline tab
        TCS_BOTTOM = 0x0002,
        TCS_RIGHT = 0x0002,
        TCS_MULTISELECT = 0x0004, // allow multi-select in button mode
        TCS_FLATBUTTONS = 0x0008,
        TCS_FORCEICONLEFT = 0x0010,
        TCS_FORCELABELLEFT = 0x0020,
        TCS_HOTTRACK = 0x0040,
        TCS_VERTICAL = 0x0080,
        TCS_TABS = 0x0000,
        TCS_BUTTONS = 0x0100,
        TCS_SINGLELINE = 0x0000,
        TCS_MULTILINE = 0x0200,
        TCS_RIGHTJUSTIFY = 0x0000,
        TCS_FIXEDWIDTH = 0x0400,
        TCS_RAGGEDRIGHT = 0x0800,
        TCS_FOCUSONBUTTONDOWN = 0x1000,
        TCS_OWNERDRAWFIXED = 0x2000,
        TCS_TOOLTIPS = 0x4000,
        TCS_FOCUSNEVER = 0x8000,
    } TabStyles;

    typedef enum TabExStyles {
 // EX styles for use with TCM_SETEXTENDEDSTYLE
        TCS_EX_FLATSEPARATORS = 0x00000001,
        TCS_EX_REGISTERDROP = 0x00000002,
    } TabExStyles;

    typedef enum TabItemState {
        TCIS_BUTTONPRESSED = 0x0001,
        TCIS_HIGHLIGHTED = 0x0002,
    } TabItemState;

    typedef enum TabItemFlags {
        TCIF_TEXT = 0x0001,
        TCIF_IMAGE = 0x0002,
        TCIF_RTLREADING = 0x0004,
        TCIF_PARAM = 0x0008,
        TCIF_STATE = 0x0010,
    } TabItemFlags;

//TCM_FIRST 0x1300 4864
    typedef enum TabMessages {
        [helpstring("SNDMSG(hwnd, TCM_GETIMAGELIST, 0, 0)")]
        TCM_GETIMAGELIST = (4864 + 2),
        [helpstring("SNDMSG(hwnd, TCM_SETIMAGELIST, 0, himl)")]
        TCM_SETIMAGELIST = (4864 + 3),
        [helpstring("SNDMSG(hwnd, TCM_GETITEMCOUNT, 0, 0L)")]
        TCM_GETITEMCOUNT = (4864 + 4),
        [helpstring("SNDMSG(hwnd, TCM_GETITEM, iItem, TC_ITEM *pitem)")]
        TCM_GETITEM = (4864 + 5),
        [helpstring("SNDMSG(hwnd, TCM_GETITEM, iItem, TC_ITEM *pitem)")]
        TCM_GETITEMW = (4864 + 60),
        [helpstring("SNDMSG(hwnd, TCM_SETITEM, iItem, TC_ITEM *pitem)")]
        TCM_SETITEM = (4864 + 6),
        [helpstring("SNDMSG(hwnd, TCM_SETITEM, iItem, TC_ITEM *pitem)")]
        TCM_SETITEMW = (4864 + 61),
        [helpstring("SNDMSG(hwnd, TCM_INSERTITEM, iItem, TC_ITEM *pitem)")]
        TCM_INSERTITEM = (4864 + 7),
        [helpstring("SNDMSG(hwnd, TCM_INSERTITEM, iItem, TC_ITEM *pitem)")]
        TCM_INSERTITEMW = (4864 + 62),
        [helpstring("SNDMSG((hwnd), TCM_DELETEITEM, iItem, 0)")]
        TCM_DELETEITEM = (4864 + 8),
        TCM_DELETEALLITEMS = (4864 + 9),
        [helpstring("SNDMSG((hwnd), TCM_GETITEMRECT, iItem, RECT *prc)")]
        TCM_GETITEMRECT = (4864 + 10),
        TCM_GETCURSEL = (4864 + 11),
        TCM_SETCURSEL = (4864 + 12),
        [helpstring("SNDMSG(hwnd, TCM_HITTEST, 0, TC_HITTESTINFO *pinfo)")]
        TCM_HITTEST = (4864 + 13),
        TCM_SETITEMEXTRA = (4864 + 14),
        TCM_ADJUSTRECT = (4864 + 40),
        TCM_SETITEMSIZE = (4864 + 41),
        TCM_REMOVEIMAGE = (4864 + 42),
        TCM_SETPADDING = (4864 + 43),
        TCM_GETROWCOUNT = (4864 + 44),
        TCM_GETTOOLTIPS = (4864 + 45),
        TCM_SETTOOLTIPS = (4864 + 46),
        TCM_GETCURFOCUS = (4864 + 47),
        TCM_SETCURFOCUS = (4864 + 48),
        TCM_SETMINTABWIDTH = (4864 + 49),
        TCM_DESELECTALL = (4864 + 50),
        TCM_HIGHLIGHTITEM = (4864 + 51),
        TCM_SETEXTENDEDSTYLE = (4864 + 52), // optional wParam = = mask
        TCM_GETEXTENDEDSTYLE = (4864 + 53),
    } TabMessages;

    typedef enum TabNotifications {
        TCN_KEYDOWN = (-550 - 0),
        TCN_SELCHANGE = (-550 - 1),
        TCN_SELCHANGING = (-550 - 2),
        TCN_GETOBJECT = (-550 - 3),
    } TabNotifications;

 // = ===== ANIMATE CONTROL = =============

    typedef enum AnimateStyles {
        ACS_CENTER = 0x0001,
        ACS_TRANSPARENT = 0x0002,
        ACS_AUTOPLAY = 0x0004,
        ACS_TIMER = 0x0008, // don't use threads... use timers
    } AnimateStyles;

    typedef enum AnimateMessages {
        ACM_OPENA = (0x400+100),
        ACM_OPENW = (0x400+103),
        ACM_PLAY = (0x400+101),
        ACM_STOP = (0x400+102),
    } AnimateMessages;

    typedef enum AnimateNotificationCodes {
        ACN_START = 1,
        ACN_STOP = 2,
    } AnimateNotificationCodes;

// = ===== MONTHCAL CONTROL = =========
    typedef enum MCSC_Flags {
        MCSC_BACKGROUND = 0, // the background color (between months)
        MCSC_TEXT = 1, // the dates
        MCSC_TITLEBK = 2, // background of the title
        MCSC_TITLETEXT = 3,
        MCSC_MONTHBK = 4, // background within the month cal
        MCSC_TRAILINGTEXT = 5, // the text color of header & trailing days
    } MonthCalSetColorFlags;

    typedef enum MCHT_Flags {
        MCHT_TITLE = 0x00010000,
        MCHT_CALENDAR = 0x00020000,
        MCHT_TODAYLINK = 0x00030000,

        MCHT_NEXT = 0x01000000, // these indicate that hitting
        MCHT_PREV = 0x02000000, // here will go to the next/prev month

        MCHT_NOWHERE = 0x00000000,

        MCHT_TITLEBK = (0x00010000),
        MCHT_TITLEMONTH = (65537), //(MCHT_TITLE | 0x0001),
        MCHT_TITLEYEAR = (65538), //(MCHT_TITLE | 0x0002),
        MCHT_TITLEBTNNEXT = (16842755), //(MCHT_TITLE | MCHT_NEXT | 0x0003),
        MCHT_TITLEBTNPREV = (33619971), //(MCHT_TITLE | MCHT_PREV | 0x0003),

        MCHT_CALENDARBK = (0x00020000), //(MCHT_CALENDAR),
        MCHT_CALENDARDATE = (131073), //(MCHT_CALENDAR | 0x0001),
        MCHT_CALENDARDATENEXT = (16908289), //(MCHT_CALENDARDATE | MCHT_NEXT),
        MCHT_CALENDARDATEPREV = (33685505), //(MCHT_CALENDARDATE | MCHT_PREV),
        MCHT_CALENDARDAY = (131074), //(MCHT_CALENDAR | 0x0002),
        MCHT_CALENDARWEEKNUM = (131075), //(MCHT_CALENDAR | 0x0003),
    } MonthCalHitTestFlags;

    typedef enum MonthCalMessages {

        MCM_FIRST = 0x1000,
        MCM_GETCURSEL = (4096 + 1),
        MCM_SETCURSEL = (4096 + 2),
        MCM_GETMAXSELCOUNT = (4096 + 3),
        MCM_SETMAXSELCOUNT = (4096 + 4),
        MCM_GETSELRANGE = (4096 + 5),
        MCM_SETSELRANGE = (4096 + 6),
        MCM_GETMONTHRANGE = (4096 + 7),
        MCM_SETDAYSTATE = (4096 + 8),
        MCM_GETMINREQRECT = (4096 + 9),
        MCM_SETCOLOR = (4096 + 10),
        MCM_GETCOLOR = (4096 + 11),
        MCM_SETTODAY = (4096 + 12),
        MCM_GETTODAY = (4096 + 13),
        MCM_HITTEST = (4096 + 14),
        MCM_SETFIRSTDAYOFWEEK = (4096 + 15),
        MCM_GETFIRSTDAYOFWEEK = (4096 + 16),
        MCM_GETRANGE = (4096 + 17),
        MCM_SETRANGE = (4096 + 18),
        MCM_GETMONTHDELTA = (4096 + 19),
        MCM_SETMONTHDELTA = (4096 + 20),
        MCM_GETMAXTODAYWIDTH = (4096 + 21),
    } MonthCalMessages;


    typedef enum MonthCallNotifications {

        MCN_SELCHANGE = (-750 + 1),
        MCN_GETDAYSTATE = (-750 + 3),
        MCN_SELECT = (-750 + 4),

    } MonthCallNotifications;

    typedef enum MonthCalStyles {
        MCS_DAYSTATE = 0x0001,
        MCS_MULTISELECT = 0x0002,
        MCS_WEEKNUMBERS = 0x0004,
        MCS_NOTODAYCIRCLE = 0x0008,
        MCS_NOTODAY = 0x0010,
    } MonthCalStyles;

    typedef enum MonthCalGetMonthRangeFlags {
        GMR_VISIBLE = 0, // visible portion of display
        GMR_DAYSTATE = 1 // above plus the grayed out parts of
 // partially displayed months
    } MonthCalGetMonthRangeFlags;

// ******* DATETIMEPICK CONTROL
//Private Const DTM_FIRST = 4096 &H1000

    typedef enum DateTimePickMessages {
        DTM_FIRST = 0x1000,
        DTM_GETSYSTEMTIME = (4096 + 1),
        DTM_SETSYSTEMTIME = (4096 + 2),
        DTM_GETRANGE = (4096 + 3),
        DTM_SETRANGE = (4096 + 4),
        DTM_SETFORMAT = (4096 + 5),
        DTM_SETFORMATW = (4096 + 50),
        DTM_SETMCCOLOR = (4096 + 6),
        DTM_GETMCCOLOR = (4096 + 7),
        DTM_GETMONTHCAL = (4096 + 8),
        DTM_SETMCFONT = (4096 + 9),
        DTM_GETMCFONT = (4096 + 10),
    } DateTimePickMessages;

    typedef enum DateTimePickStyles {
        DTS_UPDOWN = 0x0001 , // use UPDOWN instead of MONTHCAL
        DTS_SHOWNONE = 0x0002 , // allow a NONE selection
        DTS_SHORTDATEFORMAT = 0x0000 , // use the short date format (app must forward WM_WININICHANGE messages)
        DTS_LONGDATEFORMAT = 0x0004 , // use the long date format (app must forward WM_WININICHANGE messages)
        DTS_TIMEFORMAT = 0x0009 , // use the time format (app must forward WM_WININICHANGE messages)
        DTS_SHORTDATECENTURYFORMAT = 0x000C, // short date format with century (app must forward WM_WININICHANGE        DTS_APPCANPARSE = 0x0010 , // allow user entered strings (app MUST respond to DTN_USERSTRING)
        DTS_RIGHTALIGN = 0x0020 , // right-align popup instead of left-align it
    } DateTimePickStyles;

    typedef enum DateTimePickNotifications {

        DTN_DATETIMECHANGE = (-760 + 1), // the systemtime has changed
        DTN_USERSTRING = (-760 + 2), // the user has entered a string
        DTN_USERSTRINGW = (-760 + 15),
        DTN_WMKEYDOWN = (-760 + 3), // modify keydown on app format field (X)
        DTN_WMKEYDOWNW = (-760 + 16),
        DTN_FORMATA = (-760 + 4), // query display for app format field (X)
        DTN_FORMATW = (-760 + 17),
        DTN_FORMATQUERYA = (-760 + 5), // query formatting info for app format field (X)
        DTN_FORMATQUERYW = (-760 + 18),
        DTN_DROPDOWN = (-760 + 6), // MonthCal has dropped down
        DTN_CLOSEUP = (-760 + 7), // MonthCal is popping up
    } DateTimePickNotifications;

    typedef enum DateTimePickDateTimeChangeFlags {
        GDT_ERROR = -1,
        GDT_VALID = 0,
        GDT_NONE = 1,
    } DateTimePickDateTimeChangeFlags;

    typedef enum DateTimePickGetDateTimeRangeFlags {
        GDTR_MIN = 0x0001,
        GDTR_MAX = 0x0002,
    } DateTimePickGetDateTimeRangeFlags;


///    IP Address edit control
//IPM_CLEARADDRESS = (WM_USER+100)

    typedef enum IPAddressMessages {
        IPM_CLEARADDRESS = (1024+100),// no parameters
        IPM_SETADDRESS = (1024+101),// lparam = TCP/IP address
        IPM_GETADDRESS = (1024+102),// lresult = # of non black fields.
 // lparam = LPLONG for TCP/IP address
        IPM_SETRANGE = (1024+103), // wparam = field, lparam = range
        IPM_SETFOCUS = (1024+104), // wparam = field
        IPM_ISBLANK = (1024+105), // no parameters
    } IPAddressMessages;

    typedef enum IPAddressNotificationCodes {
        IPN_FIELDCHANGED = (-860 - 0),
    } IPAddressNotificationCodes;
//---------------------------------------------------------------------------------------
// Pager Control Styles
//---------------------------------------------------------------------------------------
    typedef enum PagerStyles {
        PGS_VERT = 0x0000,
        PGS_HORZ = 0x0001,
        PGS_AUTOSCROLL = 0x00000002,
        PGS_DRAGNDROP = 0x00000004,
    } PagerStyles;

//---------------------------------------------------------------------------------------
// Pager Button State
//---------------------------------------------------------------------------------------
//The scroll can be in one of the following control State
    typedef enum PagerButtonStates {
        PGF_INVISIBLE = 0, // Scroll button is not visible
        PGF_NORMAL = 1, // Scroll button is in normal state
        PGF_GRAYED = 2, // Scroll button is in grayed state
        PGF_DEPRESSED = 4, // Scroll button is in depressed state
        PGF_HOT = 8, // Scroll button is in hot state
    } PagerButtonStates;

    typedef enum PagerButton {
        PGB_TOPORLEFT = 0,
        PGB_BOTTOMORRIGHT = 1,
    } PagerButton ;

//---------------------------------------------------------------------------------------
// Pager Control  Messages
//---------------------------------------------------------------------------------------
    typedef enum PageMessages {
//PGM_FIRST = 5120 &H1400 0x1400
        PGM_SETCHILD = (5120 + 1), // lParam = = hwnd
        PGM_RECALCSIZE = (5120 + 2),
        PGM_FORWARDMOUSE = (5120 + 3),
        PGM_SETBKCOLOR = (5120 + 4),
        PGM_GETBKCOLOR = (5120 + 5),
        PGM_SETBORDER = (5120 + 6),
        PGM_GETBORDER = (5120 + 7),
        PGM_SETPOS = (5120 + 8),
        PGM_GETPOS = (5120 + 9),
        PGM_SETBUTTONSIZE = (5120 + 10),
        PGM_GETBUTTONSIZE = (5120 + 11),
        PGM_GETBUTTONSTATE = (5120 + 12),
    } PageMessages;

//---------------------------------------------------------------------------------------
//Pager Control Notification Messages
//---------------------------------------------------------------------------------------

    typedef enum PagerNotifications {
        PGN_SCROLL = (-900-1),
        PGN_CALCSIZE = (-900-2),
    } PagerNotifications;

    typedef enum PagerScrollFlags {
        PGF_SCROLLUP = 1,
        PGF_SCROLLDOWN = 2,
        PGF_SCROLLLEFT = 4,
        PGF_SCROLLRIGHT = 8,
    } PagerScrollFlags;

    typedef enum PagerCalcSizeFlags {
        PGF_CALCWIDT = 1,
        PGF_CALCHEIGHT = 2,
    } PagerCalcSizeFlags;

    typedef enum NativeFontStyles {
        NFS_EDIT                =0x0001,
        NFS_STATIC              =0x0002,
        NFS_LISTCOMBO           =0x0004,
        NFS_BUTTON              =0x0008,
        NFS_ALL                 =0x0010,
        NFS_USEFONTASSOC        =0x0020,
    } NativeFontStyles;

    typedef enum TrackMouseEventMessages {
        WM_MOUSEHOVER = 0x02A1,
        WM_MOUSELEAVE = 0x02A3,
    } TrackMouseEventMessages;

    typedef enum TrackMouseEventFlags {
        TME_HOVER = 0x00000001,
        TME_LEAVE = 0x00000002,
        TME_NONCLIENT = 0x00000010,
        TME_QUERY = 0x40000000,
        TME_CANCEL = 0x80000000,
    } TrackMouseEventFlags;

    typedef enum  TrackMouseHoverTime {
        HOVER_DEFAULT = 0xFFFFFFFF,
    } TrackMouseHoverTime;

    typedef enum FlatScrollBarProps {
        WSB_PROP_CYVSCROLL = 0x00000001,
        WSB_PROP_CXHSCROLL = 0x00000002,
        WSB_PROP_CYHSCROLL = 0x00000004,
        WSB_PROP_CXVSCROLL = 0x00000008,
        WSB_PROP_CXHTHUMB = 0x00000010,
        WSB_PROP_CYVTHUMB = 0x00000020,
        WSB_PROP_VBKGCOLOR = 0x00000040,
        WSB_PROP_HBKGCOLOR = 0x00000080,
        WSB_PROP_VSTYLE = 0x00000100,
        WSB_PROP_HSTYLE = 0x00000200,
        WSB_PROP_WINSTYLE = 0x00000400,
        WSB_PROP_PALETTE = 0x00000800,
        WSB_PROP_MASK = 0x00000FFF,
    } FlatScrollBarProps;

    typedef enum FlatScrollBarModes {
        FSB_FLAT_MODE = 2,
        FSB_ENCARTA_MODE = 1,
        FSB_REGULAR_MODE = 0,
    } FlatScrollBarModes;

typedef enum TOKEN_INFORMATION_CLASS {
    TokenUser = 1,
    TokenGroups,
    TokenPrivileges,
    TokenOwner,
    TokenPrimaryGroup,
    TokenDefaultDacl,
    TokenSource,
    TokenType,
    TokenImpersonationLevel,
    TokenStatistics,
    TokenRestrictedSids,
    TokenSessionId
} TOKEN_INFORMATION_CLASS;

// ******* FTP stuff

    typedef enum ICON_SIZES {
        Icon_16x16 = 16,
        Icon_32x32 = 32,
        Icon_48x48 = 48,
		Icon_72x72 = 72,
    } ICON_SIZES;

// ******* XColors

typedef enum XColors{
    x_White                 = 0xFFFFFF,
    x_Red                   = 0xFF,
    x_Green                 = 0xFF00,
    x_Blue                  = 0xFF0000,
    x_Magenta               = 0xFF00FF,
    x_Cyan                  = 0xFFFF00,
    x_Yellow                = 0xFFFF,
    x_Black                 = 0x0,
    x_Aquamarine            = 0x93DB70,
    x_Bakers_Chocolate      = 0x17335C,
    x_Blue_Violet           = 0x9F5F9F,
    x_Brass                 = 0x42A6B5,
    x_Bright_Gold           = 0x19D9D9,
    x_Brown                 = 0x2A2AA6,
    x_Bronze                = 0x53788C,
    x_Bronze_II             = 0x3D7DA6,
    x_Cadet_Blue            = 0x9F9F5F,
    x_Cool_Copper           = 0x1987D9,
    x_Copper                = 0x3373B8,
    x_Coral                 = 0x7FFF,
    x_Corn_Flower_Blue      = 0x6F4242,
    x_Dark_Brown            = 0x33405C,
    x_Dark_Green            = 0x2F4F2F,
    x_Dark_Green_Copper     = 0x6E764A,
    x_Dark_Olive_Green      = 0x2F4F4F,
    x_Dark_Orchid           = 0xCD3299,
    x_Dark_Purple           = 0x781F87,
    x_Dark_Slate_Blue       = 0x8E236B,
    x_Dark_Slate_Grey       = 0x4F4F2F,
    x_Dark_Tan              = 0x4F6997,
    x_Dark_Turquoise        = 0xDB9370,
    x_Dark_Wood             = 0x425E85,
    x_Dim_Grey              = 0x545454,
    x_Dusty_Rose            = 0x636385,
    x_Feldspar              = 0x7592D1,
    x_Firebrick             = 0x23238E,
    x_Forest_Green          = 0x238E23,
    x_Gold                  = 0x327FCD,
    x_Goldenrod             = 0x70DBDB,
    x_Grey                  = 0xC0C0C0,
    x_Green_Copper          = 0x767F52,
    x_Green_Yellow          = 0x70DB93,
    x_Hunter_Green          = 0x215E21,
    x_Indian_Red            = 0x2F2F4E,
    x_Khaki                 = 0x5F9F9F,
    x_Light_Blue            = 0xD9D9C0,
    x_Light_Grey            = 0xA8A8A8,
    x_Light_Steel_Blue      = 0xBD8F8F,
    x_Light_Wood            = 0xA6C2E9,
    x_Lime_Green            = 0x32CD32,
    x_Mandarian_Orange      = 0x3378E4,
    x_Maroon                = 0x6B238E,
    x_Medium_Aquamarine     = 0x99CD32,
    x_Medium_Blue           = 0xCD3232,
    x_Medium_Forest_Green   = 0x238E6B,
    x_Medium_Goldenrod      = 0xAEEAEA,
    x_Medium_Orchid         = 0xDB7093,
    x_Medium_Sea_Green      = 0x426F42,
    x_Medium_Slate_Blue     = 0xFF007F,
    x_Medium_Spring_Green   = 0xFF7F,
    x_Medium_Turquoise      = 0xDBDB70,
    x_Medium_Violet_Red     = 0x9370DB,
    x_Medium_Wood           = 0x6480A6,
    x_Midnight_Blue         = 0x4F2F2F,
    x_Navy_Blue             = 0x8E2323,
    x_Neon_Blue             = 0xFF4D4D,
    x_Neon_Pink             = 0xC76EFF,
    x_New_Midnight_Blue     = 0x9C0000,
    x_New_Tan               = 0x9EC7EB,
    x_Old_Gold              = 0x3BB5CF,
    x_Orange                = 0x7FFF,
    x_Orange_Red            = 0x24FF,
    x_Orchid                = 0xDB70DB,
    x_Pale_Green            = 0x8FBC8F,
    x_Pink                  = 0x8F8FBC,
    x_Plum                  = 0xEAADEA,
    x_Quartz                = 0xF3D9D9,
    x_Rich_Blue             = 0xAB5959,
    x_Salmon                = 0x42426F,
    x_Scarlet               = 0x17178C,
    x_Sea_Green             = 0x688E23,
    x_Semi_Sweet_Chocolate  = 0x26426B,
    x_Sienna                = 0x236B8E,
    x_Silver                = 0xFAE8E6,
    x_Sky_Blue              = 0xCC9932,
    x_Slate_Blue            = 0xFF7F00,
    x_Spicy_Pink            = 0xAE1CFF,
    x_Spring_Green          = 0x7FFF00,
    x_Steel_Blue            = 0x8E6B23,
    x_Summer_Sky            = 0xDEB038,
    x_Tan                   = 0x7093DB,
    x_Thistle               = 0xD8BFD8,
    x_Turquoise             = 0xEAEAAD,
    x_Very_Dark_Brown       = 0x33405C,
    x_Very_Light_Grey       = 0xCDCDCD,
    x_Violet                = 0x4F2F4F,
    x_Violet_Red            = 0x9932CC,
    x_Wheat                 = 0xBFD8D8,
    x_Yellow_Green          = 0x32CC99
} XColors;

// ******* WebBrowser

typedef enum DOCHOSTUITYPE
{
	DOCHOSTUITYPE_BROWSE	= 0,
	DOCHOSTUITYPE_AUTHOR	= 1
}DOCHOSTUITYPE;

typedef enum DOCHOSTUIDBLCLK
{
	DOCHOSTUIDBLCLK_DEFAULT	= 0,
	DOCHOSTUIDBLCLK_SHOWPROPERTIES	= 1,
	DOCHOSTUIDBLCLK_SHOWCODE	= 2
}DOCHOSTUIDBLCLK;

typedef enum DOCDOWNLOADCTLFLAGS
{
	DLCTL_DLIMAGES                          = 0x00000010,
	DLCTL_VIDEOS                            = 0x00000020,
	DLCTL_BGSOUNDS                          = 0x00000040,
	DLCTL_NO_SCRIPTS                        = 0x00000080,
	DLCTL_NO_JAVA                           = 0x00000100,
	DLCTL_NO_RUNACTIVEXCTLS                 = 0x00000200,
	DLCTL_NO_DLACTIVEXCTLS                  = 0x00000400,
	DLCTL_DOWNLOADONLY                      = 0x00000800,
	DLCTL_NO_FRAMEDOWNLOAD                  = 0x00001000,
	DLCTL_RESYNCHRONIZE                     = 0x00002000,
	DLCTL_PRAGMA_NO_CACHE                   = 0x00004000,
	DLCTL_NO_BEHAVIORS                      = 0x00008000,
	DLCTL_NO_METACHARSET                    = 0x00010000,
	DLCTL_URL_ENCODING_DISABLE_UTF8         = 0x00020000,
	DLCTL_URL_ENCODING_ENABLE_UTF8          = 0x00040000,
	DLCTL_NOFRAMES                          = 0x00080000,
	DLCTL_FORCEOFFLINE                      = 0x10000000,
	DLCTL_NO_CLIENTPULL                     = 0x20000000,
	DLCTL_SILENT                            = 0x40000000,
	DLCTL_OFFLINEIFNOTCONNECTED             = 0x80000000,
	DLCTL_OFFLINE                           = 0x80000000
}DOCDOWNLOADCTLFLAGS;

typedef enum DOCHOSTUIFLAG
{	
	DOCHOSTUIFLAG_DIALOG	= 0x1,
	DOCHOSTUIFLAG_DISABLE_HELP_MENU	= 0x2,
	DOCHOSTUIFLAG_NO3DBORDER	= 0x4,
	DOCHOSTUIFLAG_SCROLL_NO	= 0x8,
	DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE	= 0x10,
	DOCHOSTUIFLAG_OPENNEWWIN	= 0x20,
	DOCHOSTUIFLAG_DISABLE_OFFSCREEN	= 0x40,
	DOCHOSTUIFLAG_FLAT_SCROLLBAR	= 0x80,
	DOCHOSTUIFLAG_DIV_BLOCKDEFAULT	= 0x100,
	DOCHOSTUIFLAG_ACTIVATE_CLIENTHIT_ONLY	= 0x200,
	DOCHOSTUIFLAG_OVERRIDEBEHAVIORFACTORY	= 0x400,
	DOCHOSTUIFLAG_CODEPAGELINKEDFONTS	= 0x800,
	DOCHOSTUIFLAG_URL_ENCODING_DISABLE_UTF8	= 0x1000,
	DOCHOSTUIFLAG_URL_ENCODING_ENABLE_UTF8	= 0x2000,
	DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE	= 0x4000,
	DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION	= 0x10000,
	DOCHOSTUIFLAG_IME_ENABLE_RECONVERSION	= 0x20000,
	DOCHOSTUIFLAG_THEME	= 0x40000,
	DOCHOSTUIFLAG_NOTHEME	= 0x80000,
	DOCHOSTUIFLAG_NOPICS	= 0x100000,
	DOCHOSTUIFLAG_NO3DOUTERBORDER	= 0x200000,
	DOCHOSTUIFLAG_DISABLE_EDIT_NS_FIXUP	= 0x400000,
	DOCHOSTUIFLAG_LOCAL_MACHINE_ACCESS_CHECK	= 0x800000,
	DOCHOSTUIFLAG_DISABLE_UNTRUSTEDPROTOCOL	= 0x1000000
}DOCHOSTUIFLAG;

//URL History
typedef enum STATURL_QUERYFLAGS {
    [helpstring("Return cached items")]
    STATURL_QUERYFLAG_ISCACHED = 0x00010000,
    [helpstring("Don't return URL")]
    STATURL_QUERYFLAG_NOURL    = 0x00020000,
    [helpstring("Don't return URL title")]
    STATURL_QUERYFLAG_NOTITLE  = 0x00040000
} STATURL_QUERYFLAGS;

typedef enum STATURLFLAGS {
    [helpstring("The page is cached.")]
    STATURLFLAG_ISCACHED	= 0x00000001
} STATURLFLAGS;