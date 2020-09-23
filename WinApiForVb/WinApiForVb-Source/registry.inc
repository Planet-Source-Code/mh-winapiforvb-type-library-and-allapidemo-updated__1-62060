
// ******* Registry functions
[
	helpstring("Registry"),
	dllname("ADVAPI32.DLL")
]
module Registry {

    [

    entry("GetUserNameA")    ]
    long A_GetUserName([in] LPSTR lpBuffer, [in,out] long * nSize);

    [

    entry("GetUserNameW")    ]
    long W_GetUserName([in] LPWSTR lpBuffer, [in,out] long * nSize);

    const long PROVIDER_KEEPS_VALUE_LENGTH = 0x1;


    //
    // Open/Create Options
    //

    // Parameter is reserved
    const long REG_OPTION_RESERVED         = 0x00000000;
    // Key is preserved when system is rebooted
    const long REG_OPTION_NON_VOLATILE     = 0x00000000;
    // Key is not preserved when system is rebooted
    const long REG_OPTION_VOLATILE         = 0x00000001;
    // Created key is a symbolic link
    const long REG_OPTION_CREATE_LINK      = 0x00000002;
    // open for backup or restore special access rules privilege required
    const long REG_OPTION_BACKUP_RESTORE   = 0x00000004;

    const long REG_LEGAL_OPTION            = 0x00000007;

    //
    // Key creation/open disposition
    //

    // New Registry Key created
    const long REG_CREATED_NEW_KEY         = 0x00000001;
    // Existing Key opened
    const long REG_OPENED_EXISTING_KEY     = 0x00000002;

    //
    // Key restore flags
    //

    // Restore whole hive volatile
    const long REG_WHOLE_HIVE_VOLATILE     = 0x00000001;
    // Unwind changes to last flush
    const long REG_REFRESH_HIVE            = 0x00000002;

    //
    // Notify filter values
    //
    // Create or delete (child)
    const long REG_NOTIFY_CHANGE_NAME          = 0x00000001;
    const long REG_NOTIFY_CHANGE_ATTRIBUTES    = 0x00000002;
    // time stamp
    const long REG_NOTIFY_CHANGE_LAST_SET      = 0x00000004;
    const long REG_NOTIFY_CHANGE_SECURITY      = 0x00000008;

    const long REG_LEGAL_CHANGE_FILTER         = 0x0000000F;

    //
    // Default values for parameters that do not exist in the Win 3.1
    // compatible APIs.
    //

    //         #define WIN31_CLASS              NULL

    //
    // API Prototypes.
    //

    [
    entry("RegOpenKeyExA")    ]
    LONG A_RegOpenKeyEx(
        [in] long hKey,
        [in] LPSTR lpSubKey,
        [in] DWORD ulOptions,
        [in] long samDesired,
        [out] long * phkResult);

    [
    entry("RegOpenKeyExW")    ]
    LONG W_RegOpenKeyEx(
        [in] long hKey,
        [in] LPWSTR lpSubKey,
        [in] DWORD ulOptions,
        [in] long samDesired,
        [out] long * phkResult);


    [
    entry("RegCloseKey")    ]
    long RegCloseKey([in] long hKey);

    [
    entry("RegConnectRegistryA")    ]
    long A_RegConnectRegistry(
		[in] LPSTR lpMachineName,
		[in] long hKey,
		[out] long * phkResult);

    [
    entry("RegConnectRegistryW")    ]
    long W_RegConnectRegistry(
		[in] LPWSTR lpMachineName,
		[in] long hKey,
		[out] long * phkResult);

    [
    entry("RegCreateKeyExA")    ]
    long A_RegCreateKeyEx(
        [in] long hKey,
        [in] LPSTR lpSubKey,
        [in] long Reserved,
        [in] LPSTR lpClass,
        [in] long dwOptions,
        [in] long samDesired,
		[in] void * lpSecurityAttributes, //SECURITY_ATTRIBUTES
        [out] long * phkResult,
        [out] long * lpdwDisposition
        );

    [
    entry("RegCreateKeyExW")    ]
    long W_RegCreateKeyEx(
        [in] long hKey,
        [in] LPWSTR lpSubKey,
        [in] long Reserved,
        [in] LPWSTR lpClass,
        [in] long dwOptions,
        [in] long samDesired,
		[in] void * lpSecurityAttributes, //SECURITY_ATTRIBUTES
        [out] long * phkResult,
        [out] long * lpdwDisposition
        );

    [
    entry("RegDeleteKeyA")    ]
    long A_RegDeleteKey([in] long hKey,[in] LPSTR lpSubKey);

    [
    entry("RegDeleteKeyW")    ]
    long W_RegDeleteKey([in] long hKey,[in] LPWSTR lpSubKey);

    [
    entry("RegDeleteValueA")    ]
    long A_RegDeleteValue([in] long hKey,[in] LPSTR lpValueName);

    [
    entry("RegDeleteValueW")    ]
    long W_RegDeleteValue([in] long hKey,[in] LPWSTR lpValueName);

    [
    entry("RegEnumKeyExA")    ]
    long A_RegEnumKeyEx(
        [in] long hKey,
        [in] long dwIndex,
        [out] LPSTR lpName,
        [in, out] long * lpcbName,
        [in] long * lpReserved, //Must be NULL, in VB: byVal 0&
        [in, out] LPSTR lpClass,
        [in, out] long * lpcbClass,
        [out] LPVOID lpftLastWriteTime);



    [
    entry("RegEnumKeyExW")    ]
    long W_RegEnumKeyEx(
        [in] long hKey,
        [in] long dwIndex,
        [out] LPWSTR lpName,
        [in, out] long * lpcbName,
        [in] long * lpReserved,
        [in, out] LPWSTR lpClass,
        [in, out] long * lpcbClass,
        [out] LPVOID lpftLastWriteTime);

    [
    entry("RegEnumValueA")	]
    LONG A_RegEnumValueByte(
        [in] long hKey,
        [in] DWORD dwIndex,
        [out] LPSTR lpValueName,
        [in, out] DWORD * lpcbValueName,
        [in] DWORD * lpReserved,
        [out] long * lpType,
        [out] BYTE * lpData,
        [in, out] DWORD * lpcbData);

    [
    entry("RegEnumValueW")	]
    LONG W_RegEnumValueByte(
        [in] long hKey,
        [in] DWORD dwIndex,
        [out] LPWSTR lpValueName,
        [in, out] DWORD * lpcbValueName,
        [in] DWORD * lpReserved,
        [out] long * lpType,
        [out] BYTE * lpData,
        [in, out] DWORD * lpcbData);

    [
    entry("RegEnumValueA")	]
    LONG A_RegEnumValueAny(
        [in] long hKey,
        [in] DWORD dwIndex,
        [out] LPSTR lpValueName,
        [in, out] DWORD * lpcbValueName,
        [in] DWORD * lpReserved,
        [out] long * lpType,
        [out] void * lpData,
        [in, out] DWORD * lpcbData);

    [
    entry("RegEnumValueA")	]
    LONG W_RegEnumValueAny(
        [in] long hKey,
        [in] DWORD dwIndex,
        [out] LPWSTR lpValueName,
        [in, out] DWORD * lpcbValueName,
        [in] DWORD * lpReserved,
        [out] long * lpType,
        [out] void * lpData,
        [in, out] DWORD * lpcbData);

    [
    entry("RegQueryValueExA")    ]
    LONG A_RegQueryValueExAny(
		[in] long hKey,
        [in] LPSTR lpValueName,
        [in] DWORD * lpReserved,
        [out] long * lpType,
        [out] void * lpData,
        [in, out] DWORD * lpcbData);

    [
    entry("RegQueryValueExW")    ]
    LONG W_RegQueryValueExAny(
		[in] long hKey,
        [in] LPWSTR lpValueName,
        [in] DWORD * lpReserved,
        [out] long * lpType,
        [out] void * lpData,
        [in, out] DWORD * lpcbData);

    [
    entry("RegSetValueExA")    ]
    LONG A_RegSetValueExAny(
        [in] long hKey,
        [in] LPSTR lpValueName,
        [in] DWORD Reserved,
        [in] long dwType,
        [in] void * lpData, //EREGTYPE
        [in] DWORD cbData);

    [
    entry("RegSetValueExA")    ]
    LONG W_RegSetValueExAny(
        [in] long hKey,
        [in] LPWSTR lpValueName,
        [in] DWORD Reserved,
        [in] long dwType, //EREGTYPE
        [in] void * lpData,
        [in] DWORD cbData);

    [
    entry("RegFlushKey")    ]
    LONG RegFlushKey([in] long hKey);

    [
    entry("RegQueryInfoKeyA")    ]
    LONG A_RegQueryInfoKey(
        [in] long hKey,
        [out] LPSTR lpClass,
        [in, out] DWORD * lpcbClass,
        [in] DWORD * lpReserved, //Must be NULL
        [out] DWORD * lpcSubKeys,
        [out] DWORD * lpcbMaxSubKeyLen,
        [out] DWORD * lpcbMaxClassLen,
        [out] DWORD * lpcValues,
        [out] DWORD * lpcbMaxValueNameLen,
        [out] DWORD * lpcbMaxValueLen,
        [out] DWORD * lpcbSecurityDescriptor,
        [out] LPVOID lpftLastWriteTime);

    [
    entry("RegQueryInfoKeyA")    ]
    LONG W_RegQueryInfoKey(
        [in] long hKey,
        [out] LPWSTR lpClass,
        [in, out] DWORD * lpcbClass,
        [in] DWORD * lpReserved, //Must be NULL
        [out] DWORD * lpcSubKeys,
        [out] DWORD * lpcbMaxSubKeyLen,
        [out] DWORD * lpcbMaxClassLen,
        [out] DWORD * lpcValues,
        [out] DWORD * lpcbMaxValueNameLen,
        [out] DWORD * lpcbMaxValueLen,
        [out] DWORD * lpcbSecurityDescriptor,
        [out] LPVOID lpftLastWriteTime);

// *********

	const LONG IS_TEXT_UNICODE_ASCII16               = 0x0001;
	const LONG IS_TEXT_UNICODE_REVERSE_ASCII16       = 0x0010;

	const LONG IS_TEXT_UNICODE_STATISTICS            = 0x0002;
	const LONG IS_TEXT_UNICODE_REVERSE_STATISTICS    = 0x0020;

	const LONG IS_TEXT_UNICODE_CONTROLS              = 0x0004;
	const LONG IS_TEXT_UNICODE_REVERSE_CONTROLS      = 0x0040;

	const LONG IS_TEXT_UNICODE_SIGNATURE             = 0x0008;
	const LONG IS_TEXT_UNICODE_REVERSE_SIGNATURE     = 0x0080;

	const LONG IS_TEXT_UNICODE_ILLEGAL_CHARS         = 0x0100;
	const LONG IS_TEXT_UNICODE_ODD_LENGTH            = 0x0200;
	const LONG IS_TEXT_UNICODE_DBCS_LEADBYTE         = 0x0400;
	const LONG IS_TEXT_UNICODE_NULL_BYTES            = 0x1000;

	const LONG IS_TEXT_UNICODE_UNICODE_MASK          = 0x000F;
	const LONG IS_TEXT_UNICODE_REVERSE_MASK          = 0x00F0;
	const LONG IS_TEXT_UNICODE_NOT_UNICODE_MASK      = 0x0F00;
	const LONG IS_TEXT_UNICODE_NOT_ASCII_MASK        = 0xF000;

    [

    entry("IsTextUnicode")    ]
	long A_IsStringUnicode([in] LPSTR lpBuffer, [in] long cb, [in,out] long * lpi);

    [

    entry("IsTextUnicode")    ]
	long W_IsStringUnicode([in] LPWSTR lpBuffer, [in] long cb, [in,out] long * lpi);

    [

    entry("IsTextUnicode")    ]
	long A_IsStringPtrUnicode([in] long * lpBuffer, [in] long cb, [in,out] long * lpi);

    [

    entry("IsTextUnicode")    ]
	long W_IsStringPtrUnicode([in] long * lpBuffer, [in] long cb, [in,out] long * lpi);

// ***********

////////////////////////////////////////////////////////////////////////
//                                                                    //
//               NT Defined Privileges                                //
//                                                                    //
////////////////////////////////////////////////////////////////////////

//
// Current security descriptor revision value
//

const long SECURITY_DESCRIPTOR_REVISION    = (1);
const long SECURITY_DESCRIPTOR_REVISION1   = (1);

const long OWNER_SECURITY_INFORMATION      = (0x00000001);
const long GROUP_SECURITY_INFORMATION      = (0x00000002);
const long DACL_SECURITY_INFORMATION       = (0x00000004);
const long SACL_SECURITY_INFORMATION       = (0x00000008);
const long PROCESS_TERMINATE        = (0x0001);  
const long PROCESS_CREATE_THREAD    = (0x0002);  
const long PROCESS_SET_SESSIONID    = (0x0004);  
const long PROCESS_VM_OPERATION     = (0x0008);  
const long PROCESS_VM_READ          = (0x0010);  
const long PROCESS_VM_WRITE         = (0x0020);  
const long PROCESS_DUP_HANDLE       = (0x0040);  
const long PROCESS_CREATE_PROCESS   = (0x0080);  
const long PROCESS_SET_QUOTA        = (0x0100);  
const long PROCESS_SET_INFORMATION  = (0x0200);  
const long PROCESS_QUERY_INFORMATION	= (0x0400);

const long MAXIMUM_PROCESSORS = 32;

const LPSTR SE_CREATE_TOKEN_NAME              = "SeCreateTokenPrivilege";
const LPSTR SE_ASSIGNPRIMARYTOKEN_NAME        = "SeAssignPrimaryTokenPrivilege";
const LPSTR SE_LOCK_MEMORY_NAME               = "SeLockMemoryPrivilege";
const LPSTR SE_INCREASE_QUOTA_NAME            = "SeIncreaseQuotaPrivilege";
const LPSTR SE_UNSOLICITED_INPUT_NAME         = "SeUnsolicitedInputPrivilege";
const LPSTR SE_MACHINE_ACCOUNT_NAME           = "SeMachineAccountPrivilege";
const LPSTR SE_TCB_NAME                       = "SeTcbPrivilege";
const LPSTR SE_SECURITY_NAME                  = "SeSecurityPrivilege";
const LPSTR SE_TAKE_OWNERSHIP_NAME            = "SeTakeOwnershipPrivilege";
const LPSTR SE_LOAD_DRIVER_NAME               = "SeLoadDriverPrivilege";
const LPSTR SE_SYSTEM_PROFILE_NAME            = "SeSystemProfilePrivilege";
const LPSTR SE_SYSTEMTIME_NAME                = "SeSystemtimePrivilege";
const LPSTR SE_PROF_SINGLE_PROCESS_NAME       = "SeProfileSingleProcessPrivilege";
const LPSTR SE_INC_BASE_PRIORITY_NAME         = "SeIncreaseBasePriorityPrivilege";
const LPSTR SE_CREATE_PAGEFILE_NAME           = "SeCreatePagefilePrivilege";
const LPSTR SE_CREATE_PERMANENT_NAME          = "SeCreatePermanentPrivilege";
const LPSTR SE_BACKUP_NAME                    = "SeBackupPrivilege";
const LPSTR SE_RESTORE_NAME                   = "SeRestorePrivilege";
const LPSTR SE_SHUTDOWN_NAME                  = "SeShutdownPrivilege";
const LPSTR SE_DEBUG_NAME                     = "SeDebugPrivilege";
const LPSTR SE_AUDIT_NAME                     = "SeAuditPrivilege";
const LPSTR SE_SYSTEM_ENVIRONMENT_NAME        = "SeSystemEnvironmentPrivilege";
const LPSTR SE_CHANGE_NOTIFY_NAME             = "SeChangeNotifyPrivilege";
const LPSTR SE_REMOTE_SHUTDOWN_NAME           = "SeRemoteShutdownPrivilege";

const long SE_OWNER_DEFAULTED              = (0x0001);
const long SE_GROUP_DEFAULTED              = (0x0002);
const long SE_DACL_PRESENT                 = (0x0004);
const long SE_DACL_DEFAULTED               = (0x0008);
const long SE_SACL_PRESENT                 = (0x0010);
const long SE_SACL_DEFAULTED               = (0x0020);
const long SE_DACL_AUTO_INHERIT_REQ        = (0x0100);
const long SE_SACL_AUTO_INHERIT_REQ        = (0x0200);
const long SE_DACL_AUTO_INHERITED          = (0x0400);
const long SE_SACL_AUTO_INHERITED          = (0x0800);
const long SE_DACL_PROTECTED               = (0x1000);
const long SE_SACL_PROTECTED               = (0x2000);
const long SE_SELF_RELATIVE                = (0x8000);

const long ANYSIZE_ARRAY = 1;
const long SE_PRIVILEGE_ENABLED_BY_DEFAULT = (0x00000001);
const long SE_PRIVILEGE_ENABLED            = (0x00000002);
const long SE_PRIVILEGE_USED_FOR_ACCESS    = (0x80000000);
const long TOKEN_ASSIGN_PRIMARY    = (0x0001);
const long TOKEN_DUPLICATE         = (0x0002);
const long TOKEN_IMPERSONATE       = (0x0004);
const long TOKEN_QUERY             = (0x0008);
const long TOKEN_QUERY_SOURCE      = (0x0010);
const long TOKEN_ADJUST_PRIVILEGES = (0x0020);
const long TOKEN_ADJUST_GROUPS     = (0x0040);
const long TOKEN_ADJUST_DEFAULT    = (0x0080);
const long TOKEN_ADJUST_SESSIONID  = (0x0100);

    [
    entry("OpenThreadToken")]
    BOOL OpenThreadToken (
        [in] HANDLE ThreadHandle,
        [in] DWORD  DesiredAccess,
        [in] BOOL OpenAsSelf,
        [out] long *TokenHandle);

	[entry("LookupPrivilegeValueA")]
    BOOL A_LookupPrivilegeValue(
        [in] LPSTR lpSystemName,
        [in] LPSTR lpName,
        [out] LPVOID lpLuid);

	[entry("LookupPrivilegeValueW")]
    BOOL W_LookupPrivilegeValue(
        [in] LPWSTR lpSystemName,
        [in] LPWSTR lpName,
        [out] LPVOID lpLuid);
	
	[entry("AdjustTokenPrivileges")]
	BOOL AdjustTokenPrivileges (
        [in] HANDLE TokenHandle,
        [in] BOOL DisableAllPrivileges,
        [in] LPVOID NewState,
        [in] DWORD BufferLength,
        [out] LPVOID PreviousState,
        [out] DWORD *ReturnLength);

	[entry("PrivilegeCheck")]
	BOOL PrivilegeCheck (
        [in] HANDLE ClientToken,
        [in] LPVOID PrivilegesSet,
        [out] LPBOOL pfResult);

	[entry("OpenProcessToken")]
	BOOL OpenProcessToken (
        [in] HANDLE ProcessHandle,
        [in] DWORD DesiredAccess,
        [out] HANDLE *TokenHandle);

	[entry("GetTokenInformation")]
	BOOL GetTokenInformation (
        [in] HANDLE TokenHandle,
        [in] long TokenInformationClass,
        [out] LPVOID TokenInformation,
        [in] DWORD TokenInformationLength,
        [out] DWORD *ReturnLength);

	[entry("SetTokenInformation")]
	BOOL SetTokenInformation (
        [in] HANDLE TokenHandle,
        [in] long TokenInformationClass,
        [in] LPVOID TokenInformation,
        [in] DWORD TokenInformationLength);
// ***************
    const LONG LOGON32_PROVIDER_DEFAULT  =  0;
    const LONG LOGON32_PROVIDER_WINNT35  =  1;
    const LONG LOGON32_PROVIDER_WINNT40  =  2;
    const LONG LOGON32_PROVIDER_WINNT50  =  3;
    const LONG LOGON32_LOGON_INTERACTIVE =  2;
    const LONG LOGON32_LOGON_BATCH       =  4;
	const LONG LOGON32_LOGON_SERVICE     =  5;
	
	[entry("LogonUserA")]
	BOOL A_LogonUser (
        [in] LPSTR lpszUsername,
        [in] LPSTR lpszDomain,
        [in] LPSTR lpszPassword,
        [in] DWORD dwLogonType,
        [in] DWORD dwLogonProvider,
		[out] HANDLE *phToken);
	
	[entry("LogonUserW")]
	BOOL W_LogonUser (
        [in] LPWSTR lpszUsername,
        [in] LPWSTR lpszDomain,
        [in] LPWSTR lpszPassword,
        [in] DWORD dwLogonType,
        [in] DWORD dwLogonProvider,
		[out] HANDLE *phToken);

}; //End ADVAPI32(registry)