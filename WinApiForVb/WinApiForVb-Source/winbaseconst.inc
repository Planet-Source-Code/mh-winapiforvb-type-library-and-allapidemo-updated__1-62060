// ******* Winbase (Kernel) consts
[
	helpstring("Win32 Base Constants"),
	dllname("NOSUCH.DLL")
]
module WinBaseConst {

    // constants for the NLS component here.

    const long LANG_NEUTRAL                  =   0x00;
    const long SUBLANG_NEUTRAL               =   0x00;    // language neutral
    const long SUBLANG_DEFAULT               =   0x01;    // user default
    const long SUBLANG_SYS_DEFAULT           =   0x02;    // system default
    const long SORT_DEFAULT                  =   0x00;    // system default
    
	const DWORD LANG_SYSTEM_DEFAULT     = 0x00000800;  // 1098765432109876543210
    const DWORD LANG_USER_DEFAULT       = 0x00000400;

    const DWORD LOCALE_SYSTEM_DEFAULT   = 0x00000800;
    const DWORD LOCALE_USER_DEFAULT     = 0x00000400;
    const DWORD LOCALE_NEUTRAL          = 0x00000000;

    // String Length Maximums.
    const DWORD MAX_LEADBYTES        = 12;             // 5 ranges, 2 bytes ea., 0 term.
    const DWORD MAX_DEFAULTCHAR      = 2;              // single or double byte

    const DWORD LOCALE_NOUSEROVERRIDE       = 0x80000000;   // do not use user overrides
    const DWORD LOCALE_USE_CP_ACP           = 0x40000000;   // use the system ACP
    const DWORD LOCALE_ILANGUAGE            = 0x00000001;   // language id
    const DWORD LOCALE_SLANGUAGE            = 0x00000002;   // localized name of language
    const DWORD LOCALE_SENGLANGUAGE         = 0x00001001;   // English name of language
    const DWORD LOCALE_SABBREVLANGNAME      = 0x00000003;   // abbreviated language name
    const DWORD LOCALE_SNATIVELANGNAME      = 0x00000004;   // native name of language
    const DWORD LOCALE_ICOUNTRY             = 0x00000005;   // country code
    const DWORD LOCALE_SCOUNTRY             = 0x00000006;   // localized name of country
    const DWORD LOCALE_SENGCOUNTRY          = 0x00001002;   // English name of country
    const DWORD LOCALE_SABBREVCTRYNAME      = 0x00000007;   // abbreviated country name
    const DWORD LOCALE_SNATIVECTRYNAME      = 0x00000008;   // native name of country
    const DWORD LOCALE_IDEFAULTLANGUAGE     = 0x00000009;   // default language id
    const DWORD LOCALE_IDEFAULTCOUNTRY      = 0x0000000A;   // default country code
    const DWORD LOCALE_IDEFAULTCODEPAGE     = 0x0000000B;   // default oem code page
    const DWORD LOCALE_IDEFAULTANSICODEPAGE = 0x00001004;   // default ansi code page

//
    const long CURRENCY_MULTIPLIER        = 10000;

    const long OWNER_SECURITY_INFORMATION  = 0x00000001;
    const long GROUP_SECURITY_INFORMATION  = 0x00000002;
    const long DACL_SECURITY_INFORMATION   = 0x00000004;
    const long SACL_SECURITY_INFORMATION   = 0x00000008;

    const long INVALID_HANDLE_VALUE  = -1;
    const long HFILE_INVALID = -1;
    const long hInvalid = -1;
    const long hNull = 0;
    const long INVALID_FILE_SIZE = 0xFFFFFFFF;
    const long INVALID_ATTRIBUTE = 0xFFFFFFFF;

    const long FILE_BEGIN          = 0;
    const long FILE_CURRENT        = 1;
    const long FILE_END            = 2;

    const long TIME_ZONE_ID_INVALID = 0xFFFFFFFF;

    const long WAIT_FAILED          = 0xFFFFFFFF;
    const long WAIT_OBJECT_0        = 0x00000000;

    const long WAIT_ABANDONED       = 0x00000080;
    const long WAIT_ABANDONED_0     = 0x00000080;

    const long WAIT_TIMEOUT         = 0x00000102;
    const long WAIT_IO_COMPLETION   = 0x000000C0;
    const long STILL_ACTIVE         = 0x00000103;

	
    const long DUPLICATE_CLOSE_SOURCE     = 0x00000001; 
    const long DUPLICATE_SAME_ACCESS      = 0x00000002;

    // File creation flags must start at the high end since they
    // are combined with the attributes

    const long FILE_FLAG_WRITE_THROUGH      = 0x80000000;
    const long FILE_FLAG_OVERLAPPED         = 0x40000000;
    const long FILE_FLAG_NO_BUFFERING       = 0x20000000;
    const long FILE_FLAG_RANDOM_ACCESS      = 0x10000000;
    const long FILE_FLAG_SEQUENTIAL_SCAN    = 0x08000000;
    const long FILE_FLAG_DELETE_ON_CLOSE    = 0x04000000;
    const long FILE_FLAG_BACKUP_SEMANTICS   = 0x02000000;
    const long FILE_FLAG_POSIX_SEMANTICS    = 0x01000000;

    const long CREATE_NEW           = 1;
    const long CREATE_ALWAYS        = 2;
    const long OPEN_EXISTING        = 3;
    const long OPEN_ALWAYS          = 4;
    const long TRUNCATE_EXISTING    = 5;

    // Define the NamedPipe definitions

    // Define the dwOpenMode values for CreateNamedPipe

    const long PIPE_ACCESS_INBOUND        = 0x00000001;
    const long PIPE_ACCESS_OUTBOUND       = 0x00000002;
    const long PIPE_ACCESS_DUPLEX         = 0x00000003;

    // Define the Named Pipe End flags for GetNamedPipeInfo

    const long PIPE_CLIENT_END            = 0x00000000;
    const long PIPE_SERVER_END            = 0x00000001;

    // Define the dwPipeMode values for CreateNamedPipe

    const long PIPE_WAIT                  = 0x00000000;
    const long PIPE_NOWAIT                = 0x00000001;
    const long PIPE_READMODE_BYTE         = 0x00000000;
    const long PIPE_READMODE_MESSAGE      = 0x00000002;
    const long PIPE_TYPE_BYTE             = 0x00000000;
    const long PIPE_TYPE_MESSAGE          = 0x00000004;

    // Define the well known values for CreateNamedPipe nMaxInstances

    const long PIPE_UNLIMITED_INSTANCES   = 255;

    const long SECURITY_CONTEXT_TRACKING  = 0x00040000;
    const long SECURITY_EFFECTIVE_ONLY    = 0x00080000;

    const long SECURITY_SQOS_PRESENT      = 0x00100000;
    const long SECURITY_VALID_SQOS_FLAGS  = 0x001F0000;

	const long PAGE_NOACCESS            = 0x00000001;
    const long PAGE_READONLY            = 0x00000002;
    const long PAGE_READWRITE           = 0x00000004;
    const long PAGE_WRITECOPY           = 0x00000008;
    const long PAGE_EXECUTE             = 0x00000010;
    const long PAGE_EXECUTE_READ        = 0x00000020;
    const long PAGE_EXECUTE_READWRITE   = 0x00000040;
    const long PAGE_EXECUTE_WRITECOPY   = 0x00000080;
    const long PAGE_GUARD               = 0x00000100;
    const long PAGE_NOCACHE             = 0x00000200;
    const long MEM_COMMIT               = 0x00001000;
    const long MEM_RESERVE              = 0x00002000;
    const long MEM_DECOMMIT             = 0x00004000;
    const long MEM_RELEASE              = 0x00008000;
    const long MEM_FREE                 = 0x00010000;
    const long MEM_PRIVATE              = 0x00020000;
    const long MEM_MAPPED               = 0x00040000;
    const long MEM_TOP_DOWN             = 0x00100000;
    const long SEC_FILE                 = 0x00800000;
    const long SEC_IMAGE                = 0x01000000;
    const long SEC_RESERVE              = 0x04000000;
    const long SEC_COMMIT               = 0x08000000;
    const long SEC_NOCACHE              = 0x10000000;
    const long MEM_IMAGE                = 0x01000000;

    // Global Memory Flags

    const long GMEM_FIXED       = 0x0000;
    const long GMEM_MOVEABLE        = 0x0002;
    const long GMEM_NOCOMPACT   = 0x0010;
    const long GMEM_NODISCARD   = 0x0020;
    const long GMEM_ZEROINIT        = 0x0040;
    const long GMEM_MODIFY      = 0x0080;
    const long GMEM_DISCARDABLE = 0x0100;
    const long GMEM_SHARE       = 0x2000;
    const long GMEM_NOTIFY      = 0x4000;
    const long GMEM_LOWER       = 0x1000;

    const long GMEM_NOT_BANKED    = 0x1000;
    const long GMEM_DDESHARE      = 0x2000;
    const long GMEM_VALID_FLAGS   = 0x7F72;
    const long GMEM_INVALID_HANDLE = 0x8000;

    const long GHND             =  0x0042; // GMEM_MOVEABLE | GMEM_ZEROINIT
    const long GPTR             =  0x0040; // GMEM_FIXED | GMEM_ZEROINIT

    // Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE)
    const long GMEM_DISCARDED   = 0x4000;
    const long GMEM_LOCKCOUNT   = 0x00FF;
    // dwCreationFlag values

    const long DEBUG_PROCESS              = 0x00000001;
    const long DEBUG_ONLY_THIS_PROCESS    = 0x00000002;

    const long CREATE_SUSPENDED           = 0x00000004;

    const long DETACHED_PROCESS           = 0x00000008;

    const long CREATE_NEW_CONSOLE         = 0x00000010;

    const long NORMAL_PRIORITY_CLASS      = 0x00000020;
    const long IDLE_PRIORITY_CLASS        = 0x00000040;
    const long HIGH_PRIORITY_CLASS        = 0x00000080;
    const long REALTIME_PRIORITY_CLASS    = 0x00000100;

    const long CREATE_NEW_PROCESS_GROUP   = 0x00000200;
    const long CREATE_UNICODE_ENVIRONMENT = 0x00000400;

    const long CREATE_SEPARATE_WOW_VDM    = 0x00000800;

    const long CREATE_DEFAULT_ERROR_MODE  = 0x04000000;
    const long CREATE_NO_WINDOW           = 0x08000000;

    // GetDriveType return values
    const long DRIVE_UNKNOWN     = 0;
    const long DRIVE_NO_ROOT_DIR = 1;
    const long DRIVE_REMOVABLE   = 2;
    const long DRIVE_FIXED       = 3;
    const long DRIVE_REMOTE      = 4;
	const long DRIVE_CDROM		 = 5;
    const long DRIVE_RAMDISK     = 6;

    //define GetFreeSpace(w)                 (0x100000L)

    const long FILE_TYPE_UNKNOWN  = 0x0000;
    const long FILE_TYPE_DISK     = 0x0001;
    const long FILE_TYPE_CHAR     = 0x0002;
    const long FILE_TYPE_PIPE     = 0x0003;
    const long FILE_TYPE_REMOTE   = 0x8000;


    const long STD_INPUT_HANDLE    = -10;
    const long STD_OUTPUT_HANDLE   = -11;
    const long STD_ERROR_HANDLE    = -12;

    const long IGNORE			= 0;			// Ignore signal
    const long INFINITE        = 0xFFFFFFFF;	// Infinite timeout

    const long NMPWAIT_WAIT_FOREVER           = 0xffffffff;
    const long NMPWAIT_NOWAIT                 = 0x00000001;
    const long NMPWAIT_USE_DEFAULT_WAIT       = 0x00000000;

    const long FILE_MAP_COPY        = 0x00000001; // SECTION_QUERY
    const long FILE_MAP_WRITE       = 0x00000002; // SECTION_MAP_WRITE
    const long FILE_MAP_READ        = 0x00000004; // SECTION_MAP_READ
    const long FILE_MAP_ALL_ACCESS  = 0x000F001F; // SECTION_ALL_ACCESS

    // OpenFile() Flags
    const long OF_READ              = 0x0000;
    const long OF_WRITE             = 0x0001;
    const long OF_READWRITE         = 0x0002;
    const long OF_SHARE_COMPAT      = 0x0000;
    const long OF_SHARE_EXCLUSIVE   = 0x0010;
    const long OF_SHARE_DENY_WRITE  = 0x0020;
    const long OF_SHARE_DENY_READ   = 0x0030;
    const long OF_SHARE_DENY_NONE   = 0x0040;
    const long OF_PARSE             = 0x0100;
    const long OF_DELETE            = 0x0200;
    const long OF_VERIFY            = 0x0400;
    const long OF_SEARCH            = 0x0400;
    const long OF_CANCEL            = 0x0800;
    const long OF_CREATE            = 0x1000;
    const long OF_PROMPT            = 0x2000;
    const long OF_EXIST             = 0x4000;
    const long OF_REOPEN            = 0x8000;

    const long OFS_MAXPATHNAME       = 128;
    //define PROCESS_HEAP_REGION             0x0001
    //define PROCESS_HEAP_UNCOMMITTED_RANGE  0x0002
    //define PROCESS_HEAP_ENTRY_BUSY         0x0004
    //define PROCESS_HEAP_ENTRY_MOVEABLE     0x0010
    //define PROCESS_HEAP_ENTRY_DDESHARE     0x0020

    const long SCS_32BIT_BINARY  = 0;
    const long SCS_DOS_BINARY    = 1;
    const long SCS_WOW_BINARY    = 2;
    const long SCS_PIF_BINARY    = 3;
    const long SCS_POSIX_BINARY  = 4;
    const long SCS_OS216_BINARY  = 5;

    const long LOCKFILE_FAIL_IMMEDIATELY  = 0x00000001;
    const long LOCKFILE_EXCLUSIVE_LOCK    = 0x00000002;

    const long HANDLE_FLAG_INHERIT        = 0x00000001;
    const long HANDLE_FLAG_AUDIT_ON_CLOSE = 0x00000002;

    const long HINSTANCE_ERROR = 32;

    /*
    //  Stream Ids

    const long BACKUP_INVALID         = 0x00000000;
    const long BACKUP_DATA            = 0x00000001;
    const long BACKUP_EA_DATA         = 0x00000002;
    const long BACKUP_SECURITY_DATA   = 0x00000003;
    const long BACKUP_ALTERNATE_DATA  = 0x00000004;
    const long BACKUP_LINK            = 0x00000005;
    const long BACKUP_PROPERTY_DATA   = 0x00000006;

    //  Stream Attributes

    const long STREAM_NORMAL_ATTRIBUTE    = 0x00000000;
    const long STREAM_MODIFIED_WHEN_READ  = 0x00000001;
    const long STREAM_CONTAINS_SECURITY   = 0x00000002;
    const long STREAM_CONTAINS_PROPERTIES = 0x00000004;
    */

    const long STARTF_USESHOWWINDOW        = 0x00000001;
    const long STARTF_USESIZE              = 0x00000002;
    const long STARTF_USEPOSITION          = 0x00000004;
    const long STARTF_USECOUNTCHARS        = 0x00000008;
    const long STARTF_USEFILLATTRIBUTE     = 0x00000010;
    const long STARTF_RUNFULLSCREEN        = 0x00000020;  // ignored for non-x86 platforms
    const long STARTF_FORCEONFEEDBACK      = 0x00000040;
    const long STARTF_FORCEOFFFEEDBACK     = 0x00000080;
    const long STARTF_USESTDHANDLES        = 0x00000100;

    const long SHUTDOWN_NORETRY               = 0x00000001;

    const long DONT_RESOLVE_DLL_REFERENCES = 0x00000001;
    const long LOAD_LIBRARY_AS_DATAFILE    = 0x00000002;
    const long LOAD_WITH_ALTERED_SEARCH_PATH   = 0x00000008;

    /*
    const long DDD_RAW_TARGET_PATH        = 0x00000001;
    const long DDD_REMOVE_DEFINITION      = 0x00000002;
    const long DDD_EXACT_MATCH_ON_REMOVE  = 0x00000004;
    */

    const long MOVEFILE_REPLACE_EXISTING       = 0x00000001;
    const long MOVEFILE_COPY_ALLOWED           = 0x00000002;
    const long MOVEFILE_DELAY_UNTIL_REBOOT     = 0x00000004;

	[ helpstring("GetComputerName: Maximum name length") ]
    const long MAX_COMPUTERNAME_LENGTH  = 15;

    const long VER_PLATFORM_WIN32s        = 0;
    const long VER_PLATFORM_WIN32_WINDOWS = 1;
    const long VER_PLATFORM_WIN32_NT      = 2;

    // Abnormal termination codes

    const long TC_NORMAL       = 0;
    const long TC_HARDERR      = 1;
    const long TC_GP_TRAP      = 2;
    const long TC_SIGNAL       = 3;

    const long AC_LINE_OFFLINE                 = 0x00;
    const long AC_LINE_ONLINE                  = 0x01;
    const long AC_LINE_BACKUP_POWER            = 0x02;
    const long AC_LINE_UNKNOWN                 = 0xFF;

    const long BATTERY_FLAG_HIGH               = 0x01;
    const long BATTERY_FLAG_LOW                = 0x02;
    const long BATTERY_FLAG_CRITICAL           = 0x04;
    const long BATTERY_FLAG_CHARGING           = 0x08;
    const long BATTERY_FLAG_NO_BATTERY         = 0x80;
    const long BATTERY_FLAG_UNKNOWN            = 0xFF;

    const long BATTERY_PERCENTAGE_UNKNOWN      = 0xFF;

    const long BATTERY_LIFE_UNKNOWN        = 0xFFFFFFFF;


    // Definitions from WINNT
    /*
    const long SEC_IMAGE            = 0x01000000;
    const long SEC_RESERVE          = 0x04000000;
    const long SEC_COMMIT           = 0x08000000;
    const long SEC_NOCACHE          = 0x10000000;
    */

	//  Masks for predefined standard access types

    const long DELETE				= 0x00010000;
    const long READ_CONTROL        = 0x00020000;
    const long WRITE_DAC           = 0x00040000;
    const long WRITE_OWNER         = 0x00080000;
    const long SYNCHRONIZE         = 0x00100000;

    const long STANDARD_RIGHTS_REQUIRED	= 0x000F0000;

	const long STANDARD_RIGHTS_READ	= 0x00020000; // READ_CONTROL
    const long STANDARD_RIGHTS_WRITE   = 0x00020000; // READ_CONTROL         
    const long STANDARD_RIGHTS_EXECUTE = 0x00020000; // READ_CONTROL         

    const long STANDARD_RIGHTS_ALL     = 0x001F0000;

    const long SPECIFIC_RIGHTS_ALL     = 0x0000FFFF;

	// AccessSystemAcl access type

    const long ACCESS_SYSTEM_SECURITY  = 0x01000000;

	// MaximumAllowed access type

    const long MAXIMUM_ALLOWED         = 0x02000000;

	// These are the generic rights

    const long GENERIC_READ            = 0x80000000;
    const long GENERIC_WRITE           = 0x40000000; 
    const long GENERIC_EXECUTE         = 0x20000000; 
    const long GENERIC_ALL             = 0x10000000;

    const long FILE_SHARE_READ          = 0x00000001;
    const long FILE_SHARE_WRITE         = 0x00000002;
    const long FILE_ATTRIBUTE_READONLY  = 0x00000001;
    const long FILE_ATTRIBUTE_HIDDEN    = 0x00000002;
    const long FILE_ATTRIBUTE_SYSTEM    = 0x00000004;
    const long FILE_ATTRIBUTE_DIRECTORY = 0x00000010;
    const long FILE_ATTRIBUTE_ARCHIVE   = 0x00000020;
    const long FILE_ATTRIBUTE_NORMAL    = 0x00000080;
    const long FILE_ATTRIBUTE_TEMPORARY = 0x00000100;
    const long FILE_ATTRIBUTE_COMPRESSED	= 0x00000800;
	const long FILE_ATTRIBUTE_OFFLINE		= 4096;
	const long FILE_ATTRIBUTE_SPARSE_FILE	= 512;
	const long FILE_ATTRIBUTE_REPARSE_POINT = 1024;
	const long FILE_ATTRIBUTE_ENCRYPTED		= 64;
	const long FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 8192;

	const long EXCEPTION_CONTINUE_EXECUTION = -1;
	const long EXCEPTION_EXECUTE_HANDLER    = 1;
	const long EXCEPTION_CONTINUE_SEARCH	= 0;
	const long EXCEPTION_MAXIMUM_PARAMETERS = 15;

	const long  STATUS_WAIT_0                    = 0x00000000;    
	const long  STATUS_ABANDONED_WAIT_0          = 0x00000080;    
	const long  STATUS_USER_APC                  = 0x000000C0;    
	const long  STATUS_TIMEOUT                   = 0x00000102;    
	const long  STATUS_PENDING                   = 0x00000103;    
	const long  STATUS_SEGMENT_NOTIFICATION      = 0x40000005;    
	const long  STATUS_GUARD_PAGE_VIOLATION      = 0x80000001;    
	const long  STATUS_DATATYPE_MISALIGNMENT     = 0x80000002;    
	const long  STATUS_BREAKPOINT                = 0x80000003;    
	const long  STATUS_SINGLE_STEP               = 0x80000004;    
	const long  STATUS_ACCESS_VIOLATION          = 0xC0000005;    
	const long  STATUS_IN_PAGE_ERROR             = 0xC0000006;    
	const long  STATUS_INVALID_HANDLE            = 0xC0000008;    
	const long  STATUS_NO_MEMORY                 = 0xC0000017;    
	const long  STATUS_ILLEGAL_INSTRUCTION       = 0xC000001D;    
	const long  STATUS_NONCONTINUABLE_EXCEPTION  = 0xC0000025;    
	const long  STATUS_INVALID_DISPOSITION       = 0xC0000026;    
	const long  STATUS_ARRAY_BOUNDS_EXCEEDED     = 0xC000008C;    
	const long  STATUS_FLOAT_DENORMAL_OPERAND    = 0xC000008D;    
	const long  STATUS_FLOAT_DIVIDE_BY_ZERO      = 0xC000008E;    
	const long  STATUS_FLOAT_INEXACT_RESULT      = 0xC000008F;    
	const long  STATUS_FLOAT_INVALID_OPERATION   = 0xC0000090;    
	const long  STATUS_FLOAT_OVERFLOW            = 0xC0000091;    
	const long  STATUS_FLOAT_STACK_CHECK         = 0xC0000092;    
	const long  STATUS_FLOAT_UNDERFLOW           = 0xC0000093;    
	const long  STATUS_INTEGER_DIVIDE_BY_ZERO    = 0xC0000094;    
	const long  STATUS_INTEGER_OVERFLOW          = 0xC0000095;    
	const long  STATUS_PRIVILEGED_INSTRUCTION    = 0xC0000096;    
	const long  STATUS_STACK_OVERFLOW            = 0xC00000FD;    
	const long  STATUS_CONTROL_C_EXIT            = 0xC000013A;    
	const long  STATUS_FLOAT_MULTIPLE_FAULTS     = 0xC00002B4;    
	const long  STATUS_FLOAT_MULTIPLE_TRAPS      = 0xC00002B5;    
	const long  STATUS_ILLEGAL_VLM_REFERENCE     = 0xC00002C0;

	const long EXCEPTION_ACCESS_VIOLATION          = 0xC0000005;
	const long EXCEPTION_DATATYPE_MISALIGNMENT     = 0x80000002;
	const long EXCEPTION_BREAKPOINT                = 0x80000003; 
	const long EXCEPTION_SINGLE_STEP               = 0x80000004;
	const long EXCEPTION_ARRAY_BOUNDS_EXCEEDED     = 0xC000008C; 
	const long EXCEPTION_FLT_DENORMAL_OPERAND      = 0xC000008D;
	const long EXCEPTION_FLT_DIVIDE_BY_ZERO        = 0xC000008E;
	const long EXCEPTION_FLT_INEXACT_RESULT        = 0xC000008F;
	const long EXCEPTION_FLT_INVALID_OPERATION     = 0xC0000090;
	const long EXCEPTION_FLT_OVERFLOW              = 0xC0000091;
	const long EXCEPTION_FLT_STACK_CHECK           = 0xC0000092;
	const long EXCEPTION_FLT_UNDERFLOW             = 0xC0000093;
	const long EXCEPTION_INT_DIVIDE_BY_ZERO        = 0xC0000094;
	const long EXCEPTION_INT_OVERFLOW              = 0xC0000095;
	const long EXCEPTION_PRIV_INSTRUCTION          = 0xC0000096;
	const long EXCEPTION_IN_PAGE_ERROR             = 0xC0000006;
	const long EXCEPTION_ILLEGAL_INSTRUCTION       = 0xC000001D;
	const long EXCEPTION_NONCONTINUABLE_EXCEPTION  = 0xC0000025;
	const long EXCEPTION_STACK_OVERFLOW            = 0xC00000FD;
	const long EXCEPTION_INVALID_DISPOSITION       = 0xC0000026;
	const long EXCEPTION_GUARD_PAGE                = 0x80000001;
	const long CONTROL_C_EXIT                      = 0xC000013A;

//
// Define possible return codes from the CopyFileEx callback routine
//

	const long PROGRESS_CONTINUE  = 0;
	const long PROGRESS_CANCEL    = 1;
	const long PROGRESS_STOP      = 2;
	const long PROGRESS_QUIET     = 3;

//
// Define CopyFileEx callback routine state change values
//

	const long CALLBACK_CHUNK_FINISHED        = 0x00000000;
	const long CALLBACK_STREAM_SWITCH         = 0x00000001;

//
// Define CopyFileEx option flags
//

	const long COPY_FILE_FAIL_IF_EXISTS        = 0x00000001;
	const long COPY_FILE_RESTARTABLE           = 0x00000002;
	const long COPY_FILE_OPEN_SOURCE_FOR_WRITE = 0x00000004;

	const LONG PROCESSOR_INTEL_386     = 386;
	const LONG PROCESSOR_INTEL_486     = 486;
	const LONG PROCESSOR_INTEL_PENTIUM = 586;
	const LONG PROCESSOR_MIPS_R4000    = 4000;
	const LONG PROCESSOR_ALPHA_21064   = 21064;

	const LONG PROCESSOR_ARCHITECTURE_INTEL = 0;
	const LONG PROCESSOR_ARCHITECTURE_MIPS  = 1;
	const LONG PROCESSOR_ARCHITECTURE_ALPHA = 2;
	const LONG PROCESSOR_ARCHITECTURE_PPC   = 3;
	const LONG PROCESSOR_ARCHITECTURE_UNKNOWN = 0xFFFF;

//
//  Code Page Default Values.
//
	const long CP_ACP                  =  0;           // default to ANSI code page
	const long CP_OEMCP                =  1;           // default to OEM  code page
	const long CP_MACCP                =  2;           // default to MAC  code page
	const long CP_THREAD_ACP           =  3;           // current thread's ANSI code page
	const long CP_SYMBOL               =  42;          // SYMBOL translations

	const long CP_UTF7                 =  65000;       // UTF-7 translation
	const LONG CP_UTF8                 =  65001;       // UTF-8 translation

//
//  MBCS and Unicode Translation Flags.
//
	const long MB_PRECOMPOSED          =  0x00000001;  // use precomposed chars
	const long MB_COMPOSITE            =  0x00000002;  // use composite chars
	const long MB_USEGLYPHCHARS        =  0x00000004;  // use glyph chars, not ctrl chars
	const long MB_ERR_INVALID_CHARS    =  0x00000008;  // error for invalid chars

	const long WC_COMPOSITECHECK       =  0x00000200;  // convert composite to precomposed
	const long WC_DISCARDNS            =  0x00000010;  // discard non-spacing chars
	const long WC_SEPCHARS             =  0x00000020;  // generate separate chars
	const long WC_DEFAULTCHAR          =  0x00000040;  // replace w/ default char

	const long WC_NO_BEST_FIT_CHARS    =  0x00000400;  // do not use best fit chars

//
//  Time Flags for GetTimeFormat.
//
	const long TIME_NOMINUTESORSECONDS  = 0x00000001;  // do not use minutes or seconds
	const long TIME_NOSECONDS           = 0x00000002;  // do not use seconds
	const long TIME_NOTIMEMARKER        = 0x00000004;  // do not use time marker
	const long TIME_FORCE24HOURFORMAT   = 0x00000008;  // always use 24 hour format

//
//  Date Flags for GetDateFormat.
//
	const long DATE_SHORTDATE           = 0x00000001;  // use short date picture
	const long DATE_LONGDATE            = 0x00000002;  // use long date picture
	const long DATE_USE_ALT_CALENDAR    = 0x00000004;  // use alternate calendar (if any)

	const long DATE_YEARMONTH           = 0x00000008;  // use year month picture
	const long DATE_LTRREADING          = 0x00000010;  // add marks for left to right reading order layout
	const long DATE_RTLREADING          = 0x00000020;  // add marks for right to left reading order layout


};