// ******* Winbase Kernel functions
[
	helpstring("Win32 Base Functions"),
	dllname("KERNEL32.DLL")
]
module WinBase {

    [
    entry("RtlMoveMemory")    ]
    void CopyMemory([in] void *pvDst, [in] void *pvSrc, [in] long cbCopy);

    [
    entry("RtlMoveMemory")    ]
    void CopyMemoryStrToStr([in] LPSTR pvDst, [in] LPSTR pvSrc, [in] long cbCopy);

    [
    entry("RtlMoveMemory")    ]
    void CopyMemoryFromStr([in] void pvDst, [in] LPSTR pvSrc,[in] long cbCopy);

    [
    entry("RtlMoveMemory")    ]
    void CopyMemoryToStr([in] LPSTR pvDst, [in] void pvSrc, [in] long cbCopy);

    [
    entry("RtlMoveMemory")    ]
    void CopyMemoryByteToStr([in] LPSTR pvDst, [in] BYTE * pvSrc, [in] long cbCopy);

    [
    entry("RtlMoveMemory")    ]
    void CopyMemoryStrToByte([in] BYTE * pvDst, [in] LPSTR pvSrc, [in] long cbCopy);

    [
    entry("RtlMoveMemory")    ]
    void CopyMemoryByteToByte([in] BYTE * pvDst, [in] BYTE * pvSrc, [in] long cbCopy);

    [
    entry("RtlMoveMemory")    ]
    void CopyMemoryVal([in] void pvDest,[in] void pvSrc, [in] long cbCopy);

	// ****** Resource Management *******
    [

    entry("LoadResource")    ]
    long LoadResource([in] long hInst, [in] long hrsrc);

    [

    entry("SizeofResource")    ]
    long SizeofResource([in] long hInst, [in] long hrsrc);
/*
The FreeResource function is obsolete and is only supported for backward compatibility with 16-bit Microsoft® Windows®. It has been replced
Use one of:

Accelerator: DestroyAcceleratorTable
Bitmap: DeleteObject
Cursor: DestroyCursor
Icon: DestroyIcon
Menu: DestroyMenu

    [

    entry("FreeResource"),
    helpstring("Frees a loaded resource passed by the handle received from LoadResource"),
    ]
    long FreeResource([in] long hglbResource);
*/
    [

    entry("LockResource")    ]
    long LockResource([in] long hglbResource);

    [

    entry("FindResourceA")    ]
    long A_FindResource([in] long hInst, [in] LPSTR lpszName, [in] LPSTR lpszType);

    [

    entry("FindResourceW")    ]
    long W_FindResource([in] long hInst, [in] LPWSTR lpszName, [in] LPWSTR lpszType);

    [

    entry("FindResourceExA")    ]
	long A_FindResourceEx([in] long hModule,
						[in] LPSTR lpType,         
						[in] LPSTR lpName,
						[in] short wLanguage);
    [

    entry("FindResourceExW")    ]
	long W_FindResourceEx([in] long hModule,
						[in] LPWSTR lpType,         
						[in] LPWSTR lpName,
						[in] short wLanguage);
    [

    entry("EnumResourceTypesA")    ]
    long A_EnumResourceTypes([in] long hModule,
						  [in] long lpEnumFunc,
						  [in] long lParam);

    [

    entry("EnumResourceTypesW")    ]
    long W_EnumResourceTypes([in] long hModule,
						  [in] long lpEnumFunc,
						  [in] long lParam);
    [

    entry("EnumResourceNamesA")    ]
    long A_EnumResourceNamesStr([in] long hModule,
							  [in] LPSTR lpType,
							  [in] long lpEnumFunc,
							  [in] long lParam);
    [

    entry("EnumResourceNamesW")    ]
    long W_EnumResourceNamesStr([in] long hModule,
							  [in] LPWSTR lpType,
							  [in] long lpEnumFunc,
							  [in] long lParam);

    [

    entry("EnumResourceLanguagesA")    ]
    long A_EnumResourceLanguages([in] long hModule,
								[in] LPSTR lpType,
								[in] LPSTR lpName,
								[in] long lpEnumFunc,
								[in] long lParam);
    [

    entry("EnumResourceLanguagesW")    ]
    long W_EnumResourceLanguages([in] long hModule,
								[in] LPWSTR lpType,
								[in] LPWSTR lpName,
								[in] long lpEnumFunc,
								[in] long lParam);
    [

    entry("BeginUpdateResourceA")    ]
    long A_BeginUpdateResource([in] LPSTR pFileName, [in] long bDeleteExistingResources);

    [

    entry("BeginUpdateResourceW")    ]
    long W_BeginUpdateResource([in] LPWSTR pFileName, [in] long bDeleteExistingResources);

    [

    entry("UpdateResourceA")    ]
    long A_UpdateResource([in] long hUpdate,
						[in] LPSTR lpType,
						[in] LPSTR lpName,
						[in] short wLanguage,
						[in] long lpData, // LPVOID
						[in] long cbData);

    [

    entry("UpdateResourceW")    ]
    long W_UpdateResource([in] long hUpdate,
						[in] LPWSTR lpType,
						[in] LPWSTR lpName,
						[in] short wLanguage,
						[in] long lpData, // LPVOID
						[in] long cbData);

    [

    entry("EndUpdateResourceA")    ]
    long A_EndUpdateResource([in] long hUpdate, [in] long fDiscard);

    [

    entry("EndUpdateResourceW")    ]
    long W_EndUpdateResource([in] long hUpdate, [in] long fDiscard);


	// ****** Module Management *******

    [

    entry("LoadLibraryA")    ]
    long A_LoadLibrary([in] LPSTR lpszLibFileName);

    [

    entry("LoadLibraryW")    ]
    long W_LoadLibrary([in] LPWSTR lpszLibFileName);

    [

    entry("LoadLibraryExA")    ]
    long A_LoadLibraryEx([in] LPSTR lpLibFileName, [in] long hFile, [in] long dwFlags);

    [

    entry("LoadLibraryExW")    ]
    long W_LoadLibraryEx([in] LPWSTR lpLibFileName, [in] long hFile, [in] long dwFlags);

    [

    entry("FreeLibrary")    ]
    long FreeLibrary([in] long hModule);

    [

    entry("GetModuleFileNameA")    ]
    long A_GetModuleFileName([in] long hInst,[out] LPSTR lpszFileName, [in] long cbFileName);

    [

    entry("GetModuleFileNameW")    ]
    long W_GetModuleFileName([in] long hInst,[out] LPWSTR lpszFileName,[in] long cbFileName);

    [

    entry("GetModuleHandleA")    ]
    long A_GetModuleHandle([in] LPSTR lpszModuleName);

    [

    entry("GetModuleHandleW")    ]
    long W_GetModuleHandle([in] LPWSTR lpszModuleName);

    [

    entry("CreateProcessA")    ]
    long A_CreateProcess([in] LPSTR lpApplicationName,
						[in,out] LPSTR lpCommandLine,
						[in] void * lpProcessAttributes, // SECURITY_ATTRIBUTES
						[in] void * lpThreadAttributes,
                        [in] long bInheritHandles,
                        [in] long dwCreationFlags,
                        [in] long lpEnvironment,
                        [in] LPSTR lpCurrentDirectory,
                        [in] LPVOID  lpStartupInfo,
                        [in] LPVOID  lpProcessInformation);

    [

    entry("CreateProcessW")    ]
    long W_CreateProcess([in] LPWSTR lpApplicationName,
						[in,out] LPWSTR lpCommandLine,
						[in] void * lpProcessAttributes, // SECURITY_ATTRIBUTES
						[in] void * lpThreadAttributes,
                        [in] long bInheritHandles,
                        [in] long dwCreationFlags,
                        [in] long lpEnvironment,
                        [in] LPWSTR lpCurrentDirectory,
                        [in] LPVOID  lpStartupInfo,
                        [in] LPVOID  lpProcessInformation);

    [

    entry("CreateProcessWithLogonW")    ]
	long W_CreateProcessWithLogonW(
			[in] LPWSTR lpUsername,
			[in] LPWSTR lpDomain,
			[in] LPWSTR lpPassword,
			[in] long dwLogonFlags,
			[in] LPWSTR lpApplicationName,
			[in] LPWSTR lpCommandLine,
			[in] long dwCreationFlags,
			[in] long lpEnvironment,
			[in] LPWSTR lpCurrentDirectory,
			[in] LPVOID  lpStartupInfo,
			[out] LPVOID  lpProcessInfo);

    [

    entry("GetProcAddress")    ]
    long GetProcAddress([in] long hModule,[in] LPSTR lpProcName);

    [

    entry("OpenProcess")    ]
    long OpenProcess([in] long dwDesiredAccess,
                              [in] long bInheritHandle,
                              [in] long dwProcessId);
    [

    entry("GetCurrentProcess")    ]
    long GetCurrentProcess(void);

    [

    entry("GetCurrentProcessId")    ]
    long GetCurrentProcessId(void);

    [

    entry("ExitProcess")    ]
    void ExitProcess([in] long uExitCode);

    [

    entry("TerminateProcess")    ]
    long TerminateProcess([in] long hProcess, [in] long uExitCode);

    [

    entry("GetExitCodeProcess")    ]
    long GetExitCodeProcess([in] long hProcess, [out] long * lpExitCode);
   
    [

    entry("GetEnvironmentStringsA")    ]
    long A_GetEnvironmentStrings(void);

    [

    entry("GetEnvironmentStringsW")    ]
    long W_GetEnvironmentStrings(void);

    [

    entry("FreeEnvironmentStringsA")    ]
	long A_FreeEnvironmentStrings([in] LPSTR lpszEnvironmentBlock);

    [

    entry("FreeEnvironmentStringsW")    ]
	long W_FreeEnvironmentStrings([in] LPWSTR lpszEnvironmentBlock);

    [

    entry("GetEnvironmentVariableA")    ]
    long A_GetEnvironmentVariable([in] LPSTR lpName,[out] LPSTR lpBuffer,[in] long nSize);

    [

    entry("GetEnvironmentVariableW")    ]
    long W_GetEnvironmentVariable([in] LPWSTR lpName,[out] LPWSTR lpBuffer,[in] long nSize);

    [

    entry("SetEnvironmentVariableA")    ]
    long A_SetEnvironmentVariable([in] LPSTR lpName,[in] LPSTR lpValue);

    [

    entry("SetEnvironmentVariableW")    ]
    long W_SetEnvironmentVariable([in] LPSTR lpName,[in] LPSTR lpValue);

    [

    entry("ExpandEnvironmentStringsA")    ]
    long A_ExpandEnvironmentStrings([in] LPSTR lpSrc,[out] LPSTR lpDst,[in] long nSize);

    [

    entry("ExpandEnvironmentStringsW")    ]
    long W_ExpandEnvironmentStrings([in] LPWSTR lpSrc,[out] LPWSTR lpDst,[in] long nSize);

    [

    entry("GetStartupInfoA")    ]
    void A_GetStartupInfo([out] LPVOID  lpStartupInfo);

    [

    entry("GetStartupInfoW")    ]
    void W_GetStartupInfo([out] LPVOID  lpStartupInfo);

    [

    entry("GetCommandLineA")    ]
    long A_GetCommandLine(void);

    [

    entry("GetCommandLineW")    ]
    long W_GetCommandLine(void);

	// ******** Exception support
	[

    entry("RaiseException")    ]
    void RaiseException(
        [in] long dwExceptionCode,
        [in] long dwExceptionFlags,
        [in] long nNumberOfArguments,
        [in] long *lpArguments);

	[

    entry("SetUnhandledExceptionFilter")    ]
    long SetUnhandledExceptionFilter([in] long lpTopLevelExceptionFilter);

	[

    entry("GetExceptionCode")    ]
	long GetExceptionCode(void);

	[

    entry("GetExceptionInformation")    ]
	long GetExceptionInformation(void);

	[

    entry("FatalExit")    ]
    long FatalExit([in] long ExitCode);

    [

    entry("GetLastError")    ]
    long GetLastError(void);

    [

    entry("SetLastError")    ]
    void SetLastError([in] long dwErrCode);

    [

    entry("FormatMessageA")    ]
    long A_FormatMessage(
        [in] long dwFlags,
        [in] void * lpSource,
        [in] long dwMessageId,
        [in] long dwLanguageId,
        [in] LPSTR lpBuffer,
        [in] long nSize,
        [in] void * Arguments); //va_list *

    [

    entry("FormatMessageW")    ]
    long W_FormatMessage(
        [in] long dwFlags,
        [in] void * lpSource,
        [in] long dwMessageId,
        [in] long dwLanguageId,
        [in] LPWSTR lpBuffer,
        [in] long nSize,
        [in] void * Arguments); //va_list *


    [
	usesgetlasterror,
    entry("Beep")    ]
    BOOL Beep([in] DWORD dwFreq,[in] DWORD dwDuration);

	// ************ String Support **********

    [

    entry("lstrcmpA")    ]
    long A_lstrcmp([in] LPSTR lpszString1,[in] LPSTR lpszString2);

    [

    entry("lstrcmpW")    ]
    long W_lstrcmp([in] LPWSTR lpszString1,[in] LPWSTR lpszString2);

    [

    entry("lstrcmpA")    ]
    long A_lstrcmpStrPtr([in] LPSTR lpszString1,[in] void lpszString2);

    [

    entry("lstrcmpW")    ]
    long W_lstrcmpStrPtr([in] LPWSTR lpszString1,[in] void lpszString2);

    [

    entry("lstrcmpA")    ]
    long A_lstrcmpPtrStr([in] void lpszString1,[in] LPSTR lpszString2);

    [

    entry("lstrcmpW")    ]
    long W_lstrcmpPtrStr([in] void lpszString1,[in] LPWSTR lpszString2);

    [

    entry("lstrcmpA")    ]
    long A_lstrcmpPtrPtr([in] void lpszString1, [in] void lpszString2);

    [

    entry("lstrcmpW")    ]
    long W_lstrcmpPtrPtr([in] void lpszString1, [in] void lpszString2);

	[

    entry("lstrcmpiA")    ]
    long A_lstrcmpi([in] LPSTR lpszString1, [in] LPSTR lpszString2);

	[

    entry("lstrcmpiA")    ]
    long W_lstrcmpi([in] LPWSTR lpszString1, [in] LPWSTR lpszString2);

    [

    entry("lstrcmpiA")    ]
    long A_lstrcmpiPtrPtr([in] void lpszString1,[in] void lpszString2);

    [

    entry("lstrcmpiW")    ]
    long W_lstrcmpiPtrPtr([in] void lpszString1,[in] void lpszString2);

	[

    entry("lstrcmpiA")    ]
    long A_lstrcmpiPtrStr([in] void lpszString1,[in] LPSTR lpszString2);

	[

    entry("lstrcmpiW")    ]
    long W_lstrcmpiPtrStr([in] void lpszString1,[in] LPWSTR lpszString2);

    [

    entry("lstrcmpiA")    ]
    long A_lstrcmpiStrPtr([in] LPSTR lpszString1, [in] void lpszString2);

    [

    entry("lstrcmpiW")    ]
    long W_lstrcmpiStrPtr([in] LPWSTR lpszString1, [in] void lpszString2);

//
    [
    entry("lstrcpynA")    ]
    long A_lstrcpyn([out] LPSTR lpszDst, [in] LPSTR lpszSrc, [in] long cChars);

    [
    entry("lstrcpynW")    ]
    long W_lstrcpyn([out] LPWSTR lpszDst, [in] LPWSTR lpszSrc, [in] long cChars);

    [
    entry("lstrcpynA")    ]
    long A_lstrcpynPtrStr([out] LPSTR lpszDst,[in] void lpszSrc, [in] long cChars);

    [
    entry("lstrcpynW")    ]
    long W_lstrcpynPtrStr([in] LPWSTR lpszDst,[in] void lpszSrc, [in] long cChars);

    [
    entry("lstrcpynA")    ]
    long A_lstrcpynStrPtr([in] void lpszDst,[in] LPSTR lpszSrc, [in] long cChars);

    [
    entry("lstrcpynW")    ]
    long W_lstrcpynStrPtr([in] void lpszDst,[in] LPWSTR lpszSrc, [in] long cChars);
//
    [
    entry("lstrcpynA")    ]
    long A_lstrcpynPtrPtr([in] void lpszDst,[in] void lpszSrc, [in] long cChars);

    [
    entry("lstrcpynW")    ]
    long W_lstrcpynPtrPtr([in] void lpszDst,[in] void lpszSrc, [in] long cChars);

    [
    entry("lstrcpynA")    ]
    long A_lstrcpynByteStr([out] LPSTR lpszDst,[in] BYTE * lpszSrc, [in] long cChars);

    [
    entry("lstrcpynW")    ]
    long W_lstrcpynByteStr([out] LPWSTR lpszDst,[in] BYTE * lpszSrc, [in] long cChars);

    [
    entry("lstrcpynA")    ]
    long A_lstrcpynStrByte([out] BYTE * lpszDst,[in] LPSTR lpszSrc,[in] long cChars);

    [
    entry("lstrcpynW")    ]
    long W_lstrcpynStrByte([out] BYTE * lpszDst,[in] LPWSTR lpszSrc,[in] long cChars);

    [
    entry("lstrcpynA")    ]
    long A_lstrcpynByteByte([out] BYTE * lpszDst, [in] BYTE * lpszSrc,[in] long cChars);

    [
    entry("lstrcpynW")    ]
    long W_lstrcpynByteByte([out] BYTE * lpszDst, [in] BYTE * lpszSrc,[in] long cChars);

    [
    entry("lstrcpyA")    ]
    long A_lstrcpy([out] LPSTR lpszDst,[in] LPSTR lpszSrc);

    [

    entry("lstrcpyW")    ]
    long W_lstrcpy([out] LPWSTR lpszDst,[in] LPWSTR lpszSrc);

    [
    entry("lstrcpyA")    ]
    long A_lstrcpyStrPtr([in] void lpszDst,[in] LPSTR lpszSrc);

    [
    entry("lstrcpyW")    ]
    long W_lstrcpyStrPtr([in] void lpszDst,[in] LPWSTR lpszSrc);

    [
    entry("lstrcpyA")    ]
    long A_lstrcpyPtrStr([out] LPSTR lpszDst,[in] void lpszSrc);

    [
    entry("lstrcpyW")    ]
    long W_lstrcpyPtrStr([out] LPWSTR lpszDst,[in] void lpszSrc);

    [
    entry("lstrcpyA")    ]
    long A_lstrcpyPtrPtr([in] void lpszDst,[in] void lpszSrc);

    [
    entry("lstrcpyW")    ]
    long W_lstrcpyPtrPtr([in] void lpszDst,[in] void lpszSrc);

    [
    entry("lstrcpyA")    ]
    long A_lstrcpyStrByte([out] BYTE * lpszDst,[in] LPSTR lpszSrc);

    [
    entry("lstrcpyW")    ]
    long W_lstrcpyStrByte([out] BYTE * lpszDst,[in] LPWSTR lpszSrc);

    [
    entry("lstrcpyA")    ]
    long A_lstrcpyByteStr([out] LPSTR lpszDst, [in] BYTE * lpszSrc);

    [
    entry("lstrcpyW")    ]
    long W_lstrcpyByteStr([out] LPWSTR lpszDst, [in] BYTE * lpszSrc);

    [
    entry("lstrcpyA")    ]
    long A_lstrcpyByteByte([out] BYTE *  lpszDst,[in] BYTE * lpszSrc);

    [
    entry("lstrcpyW")    ]
    long W_lstrcpyByteByte([out] BYTE * lpszDst,[in] BYTE * lpszSrc);

    [
    entry("lstrcatA")    ]
    long A_lstrcat([in,out] LPSTR lpszDst,[in] LPSTR lpszSrc);

    [
    entry("lstrcatW")    ]
    long W_lstrcat([in,out] LPWSTR lpszDst,[in] LPWSTR lpszSrc);

    [
    entry("lstrcatA")    ]
    long A_lstrcatByteStr([in,out] LPSTR lpszDst, [in] BYTE * lpszSrc);

    [
    entry("lstrcatW")    ]
    long W_lstrcatByteStr([in,out] LPWSTR lpszDst, [in] BYTE * lpszSrc);

    [
	entry("lstrcatA")    ]
    long A_lstrcatStrByte([in,out] BYTE * lpszDst,[in] LPSTR lpszSrc);

    [
	entry("lstrcatW")    ]
    long W_lstrcatStrByte([in,out] BYTE * lpszDst,[in] LPWSTR lpszSrc);

    [
    entry("lstrcatA")    ]
    long A_lstrcatByteByte([in,out] BYTE * lpszDst,[in] BYTE * lpszSrc);

    [
    entry("lstrcatW")    ]
    long W_lstrcatByteByte([in,out] BYTE * lpszDst,[in] BYTE * lpszSrc);

	// String version only useful for Unicode
    [
    entry("lstrlenW")    ]
    long W_lstrlenStr([in] LPWSTR lpsz);

    [
    entry("lstrlenA")    ]
    long A_lstrlenPtr([in] long lpsz);

    [
    entry("lstrlenW")    ]
    long W_lstrlenPtr([in] long lpsz);

    [
    entry("lstrlenA")    ]
    long A_lstrlenByte([in] BYTE * lpsz);

    [
    entry("lstrlenW")    ]
    long W_lstrlenByte([in] BYTE * lpsz);

	// ****** File I/O **********
    [
    entry("CreateFileA")    ]
    long A_CreateFile([in] LPSTR lpFileName,
                             [in] long dwDesiredAccess,
                             [in] long dwShareMode,
							 [in] void * lpSecurityAttributes,
                             [in] long dwCreationDisposition,
                             [in] long dwFlagsAndAttributes,
                             [in] long hTemplateFile);
    [
    entry("CreateFileW")    ]
    long W_CreateFile([in] LPWSTR lpFileName,
                             [in] long dwDesiredAccess,
                             [in] long dwShareMode,
							 [in] void * lpSecurityAttributes,
                             [in] long dwCreationDisposition,
                             [in] long dwFlagsAndAttributes,
                             [in] long hTemplateFile);
    [
    entry("CreateFileA")    ]
    long A_CreateFileSec([in] LPSTR lpFileName,
                             [in] long dwDesiredAccess,
                             [in] long dwShareMode,
							 [in] LPVOID  lpSecurityAttributes,
                             [in] long dwCreationDisposition,
                             [in] long dwFlagsAndAttributes,
                             [in] long hTemplateFile);
    [
    entry("CreateFileW")    ]
    long W_CreateFileSec([in] LPWSTR lpFileName,
                             [in] long dwDesiredAccess,
                             [in] long dwShareMode,
							 [in] LPVOID  lpSecurityAttributes,
                             [in] long dwCreationDisposition,
                             [in] long dwFlagsAndAttributes,
                             [in] long hTemplateFile);

    [
    entry("GetFileInformationByHandle")    ]
    BOOL GetFileInformationByHandle([in] HANDLE hFile,[out] LPVOID lpFileInformation);

    [
    entry("GetFileType")    ]
    DWORD GetFileType([in] HANDLE hFile);

    [
    entry("GetFileSize")    ]
    DWORD GetFileSize([in] HANDLE hFile, [out] DWORD * lpFileSizeHigh);

    [
    entry("GetStdHandle")    ]
    HANDLE GetStdHandle([in] DWORD nStdHandle);

    [
    entry("SetStdHandle")    ]
    BOOL SetStdHandle([in] DWORD nStdHandle, [in] HANDLE hHandle);

    [
    entry("WriteFile")    ]
    BOOL WriteFile([in] HANDLE hFile,
                          [in] LPVOID lpBuffer,
                          [in] DWORD nNumberOfBytesToWrite,
                          [out] DWORD * lpNumberOfBytesWritten,
                          [in] LPVOID lpOverlapped); //Pointer to an OVERLAPPED structure

    [
    entry("ReadFile")    ]
    BOOL ReadFile([in] HANDLE hFile,
                         [out] LPVOID lpBuffer,
                         [in] DWORD nNumberOfBytesToRead,
                         [out] DWORD * lpNumberOfBytesRead,
                         [in] LPVOID lpOverlapped);

    [    entry("WriteFile")    ]
    BOOL WriteFileByte([in] HANDLE hFile,
                          [in] LPBYTE lpBuffer,
                          [in] DWORD nNumberOfBytesToWrite,
                          [out] LPDWORD lpNumberOfBytesWritten,
                          [in] LPVOID lpOverlapped); //Pointer to an OVERLAPPED structure

    [    entry("ReadFile")    ]
    BOOL ReadFileByte([in] HANDLE hFile,
                         [out] LPBYTE lpBuffer,
                         [in] DWORD nNumberOfBytesToRead,
                         [out] LPDWORD lpNumberOfBytesRead,
                         [in] LPVOID lpOverlapped);

    [
    entry("WriteFile")    ]
    BOOL WriteFileStr([in] HANDLE hFile,
                             [in] LPSTR lpBuffer,
                             [in] DWORD nNumberOfBytesToWrite,
                             [out] DWORD * lpNumberOfBytesWritten,
                             [in] LPVOID lpOverlapped);

    [
    entry("ReadFile")    ]
    BOOL ReadFileStr([in] HANDLE hFile,
                            [out] LPSTR lpBuffer,
                            [in] DWORD nNumberOfBytesToRead,
                            [out] DWORD * lpNumberOfBytesRead,
                            [in] LPVOID lpOverlapped);

    [
	usesgetlasterror,
    entry("FlushFileBuffers")    ]
    BOOL FlushFileBuffers([in] HANDLE hFile);

    [
    entry("SetEndOfFile")    ]
    BOOL SetEndOfFile([in] HANDLE hFile);

    [
    entry("SetFileValidData")    ]
    BOOL SetFileValidData([in] HANDLE hFile,[in] long ValidDataLength);

    [
    entry("SetFilePointer")    ]
    DWORD SetFilePointer([in] HANDLE hFile, 
								[in] LONG lDistanceToMove,
                                [in] LONG * lpDistanceToMoveHigh,
                                [in] DWORD dwMoveMethod);

    [
	entry("FindClose")    ]
    BOOL FindClose([in] HANDLE hFindFile);

    [
    entry("GetFileTime")    ]
    BOOL GetFileTime([in] HANDLE hFile,
                            [out] LPVOID  lpCreationTime,
                            [out] LPVOID  lpLastAccessTime,
                            [out] LPVOID  lpLastWriteTime);
    [
    entry("SetFileTime")    ]
    BOOL SetFileTime([in] HANDLE hFile,
                            [in] LPVOID  lpCreationTime,
                            [in] LPVOID lpLastAccessTime,
                            [in] LPVOID  lpLastWriteTime);

    [
    entry("CloseHandle")    ]
    BOOL CloseHandle([in] HANDLE hObject);

    [
    entry("DuplicateHandle")    ]
    BOOL DuplicateHandle([in] HANDLE hSourceProcessHandle,
                                [in] HANDLE hSourceHandle,
                                [in] HANDLE hTargetProcessHandle,
                                [out] HANDLE * lpTargetHandle,
                                [in] DWORD dwDesiredAccess,
                                [in] BOOL bInheritHandle, 
								[in] DWORD dwOptions);

    [
    entry("GetHandleInformation")    ]
    BOOL GetHandleInformation([in] HANDLE hObject, [out] DWORD * lpdwFlags);

    [
    entry("SetHandleInformation")    ]
    BOOL SetHandleInformation([in] HANDLE hObject, 
									 [in] DWORD dwMask,
                                     [in] DWORD dwFlags);

    [

    entry("SetFileAttributesA")    ]
    BOOL A_SetFileAttributes([in] LPSTR lpFileName, [in] DWORD dwFileAttributes);

    [

    entry("SetFileAttributesW")    ]
    BOOL W_SetFileAttributes([in] LPWSTR lpFileName, [in] DWORD dwFileAttributes);

    [

    entry("GetFileAttributesA")    ]
    DWORD A_GetFileAttributes([in] LPSTR lpFileName);

    [

    entry("GetFileAttributesW")    ]
    DWORD W_GetFileAttributes([in] LPSTR lpFileName);

    [

    entry("GetFileAttributesExA")    ]
	BOOL A_GetFileAttributesEx([in] LPSTR lpFileName,
							[in] GET_FILEEX_INFO_LEVELS fInfoLevelId,
							[out] LPVOID lpFileInformation);

    [

    entry("GetFileAttributesExW")    ]
	BOOL W_GetFileAttributesEx([in] LPWSTR lpFileName,
							[in] long fInfoLevelId,
							[out] LPVOID lpFileInformation);

    [

    entry("GetCompressedFileSizeA")    ]
    DWORD A_GetCompressedFileSize([in] LPSTR lpFileName, [out] DWORD * lpFileSizeHigh);

    [

    entry("GetCompressedFileSizeW")    ]
    DWORD W_GetCompressedFileSize([in] LPWSTR lpFileName, [out] DWORD * lpFileSizeHigh);

    [

    entry("DeleteFileA")    ]
    BOOL A_DeleteFile([in] LPSTR lpFileName);

    [

    entry("DeleteFileW")    ]
    BOOL W_DeleteFile([in] LPWSTR lpFileName);

    [
	usesgetlasterror,
    entry("FindFirstFileA")    ]
    HANDLE A_FindFirstFile([in] LPSTR lpFileName, [out] LPVOID lpFindFileData);

    [
	usesgetlasterror,
    entry("FindFirstFileW")    ]
    HANDLE W_FindFirstFile([in] LPWSTR lpFileName, [out] LPVOID lpFindFileData);

    [
	usesgetlasterror,
    entry("FindNextFileA")    ]
    BOOL A_FindNextFile([in] HANDLE hFindFile, [out] LPVOID lpFindFileData);

    [

    entry("FindNextFileW")    ]
    BOOL W_FindNextFile([in] HANDLE hFindFile, [out] LPVOID lpFindFileData);

    [

    entry("SearchPathA")    ]
    DWORD A_SearchPath([in] LPSTR lpPath, [in] LPSTR lpFileName,
                           [in] LPSTR lpExtension,
                           [in] DWORD nBufferLength,
                           [out] LPSTR lpBuffer,
                           [out] DWORD *lpFilePart);
    [

    entry("SearchPathW")    ]
    DWORD W_SearchPath([in] LPWSTR lpPath, [in] LPWSTR lpFileName,
                           [in] LPWSTR lpExtension,
                           [in] DWORD nBufferLength,
                           [out] LPWSTR lpBuffer,
                           [out] DWORD *lpFilePart);

    [

    entry("CopyFileA")    ]
    BOOL A_CopyFile([in] LPSTR lpExistingFileName,
                         [in] LPSTR lpNewFileName, 
						 [in] BOOL bFailIfExists);

    [

    entry("CopyFileW")    ]
    BOOL W_CopyFile([in] LPWSTR lpExistingFileName,
                         [in] LPWSTR lpNewFileName, 
						 [in] BOOL bFailIfExists);


//COPY_FILE_ALLOW_DECRYPTED_DESTINATION
    [

    entry("CopyFileExA")    ]
	BOOL A_CopyFileEx([in] LPSTR lpExistingFileName,
				[in] LPSTR lpNewFileName,
				[in] long * lpProgressRoutine, //CopyProgressRoutine callback
				[in] LPVOID lpData,
				[in] long * pbCancel, 
				[in] DWORD dwCopyFlags);

    [

    entry("CopyFileExW")    ]
	BOOL W_CopyFileEx([in] LPWSTR lpExistingFileName,
				[in] LPWSTR lpNewFileName,
				[in] long * lpProgressRoutine, //CopyProgressRoutine callback
				[in] LPVOID lpData,
				[in] long * pbCancel, 
				[in] DWORD dwCopyFlags);

    [

    entry("MoveFileA")    ]
    BOOL A_MoveFile([in] LPSTR lpExistingFileName, [in] LPSTR lpNewFileName);

    [

    entry("MoveFileW")    ]
    BOOL W_MoveFile([in] LPWSTR lpExistingFileName, [in] LPWSTR lpNewFileName);

    [

    entry("MoveFileExA")    ]
    BOOL A_MoveFileEx([in] LPSTR lpExistingFileName,
                          [in] LPSTR lpNewFileName, 
						  [in] DWORD dwFlags);

    [

    entry("MoveFileExW")    ]
    BOOL W_MoveFileEx([in] LPWSTR lpExistingFileName,
                          [in] LPWSTR lpNewFileName, 
						  [in] DWORD dwFlags);

    [

    entry("MoveFileWithProgressA")    ]
	BOOL A_MoveFileWithProgress([in] LPSTR lpExistingFileName,
							[in] LPSTR lpNewFileName,
							[in] long * lpProgressRoutine, //CopyProgressRoutine callback
							[in] LPVOID lpData,
							[in] DWORD dwFlags);
    [

    entry("MoveFileWithProgressW")    ]
	BOOL W_MoveFileWithProgress([in] LPWSTR lpExistingFileName,
							[in] LPWSTR lpNewFileName,
							[in] long * lpProgressRoutine, //CopyProgressRoutine callback
							[in] LPVOID lpData,
							[in] DWORD dwFlags);

    [

    entry("CreatePipe")    ]
    BOOL CreatePipe([out] HANDLE * hReadPipe,
                           [out] HANDLE * hWritePipe,
					       [in] LPVOID lpPipeAttributes,
                           [in] DWORD nSize);

    [

    entry("CreateNamedPipeA")    ]
    HANDLE A_CreateNamedPipe([in] LPSTR lpName,
                                 [in] DWORD  dwOpenMode,
                                 [in] DWORD  dwPipeMode,
                                 [in] DWORD  nMaxInstances,
                                 [in] DWORD  nOutBufferSize,
                                 [in] DWORD  nInBufferSize,
                                 [in] DWORD  nDefaultTimeOut,
								 [in] LPVOID lpSecurityAttributes); // SECURITY_ATTRIBUTES
    [

    entry("CreateNamedPipeW")    ]
    HANDLE W_CreateNamedPipe([in] LPWSTR lpName,
                                 [in] DWORD  dwOpenMode,
                                 [in] DWORD  dwPipeMode,
                                 [in] DWORD  nMaxInstances,
                                 [in] DWORD  nOutBufferSize,
                                 [in] DWORD  nInBufferSize,
                                 [in] DWORD  nDefaultTimeOut,
								 [in] LPVOID lpSecurityAttributes); // SECURITY_ATTRIBUTES


    [

    entry("DisconnectNamedPipe")    ]
    BOOL DisconnectNamedPipe([in] HANDLE hNamedPipe);

    [

    entry("SetNamedPipeHandleState")    ]
    BOOL SetNamedPipeHandleState([in] HANDLE hNamedPipe,
                                        [in] DWORD *lpMode,
                                        [in] DWORD * lpMaxCollectionCount,
                                        [in] DWORD * lpCollectDataTimeout);

    [

    entry("GetNamedPipeInfo")    ]
    BOOL GetNamedPipeInfo([in] HANDLE hNamedPipe,
                                [in] DWORD * lpFlags,
                                [out] DWORD * lpOutBufferSize,
                                [out] DWORD * lpInBufferSize,
                                [out] DWORD * lpMaxInstances);
    [

    entry("PeekNamedPipe")    ]
    BOOL PeekNamedPipeStr([in] HANDLE hNamedPipe,
                                 [out] LPSTR lpBuffer,
                                 [in] DWORD  nBufferSize,
                                 [out] DWORD * lpBytesRead,
                                 [out] DWORD * lpTotalBytesAvail,
                                 [out] DWORD * lpBytesLeftThisMessage);
    [

    entry("PeekNamedPipe")    ]
    BOOL PeekNamedPipe([in] HANDLE hNamedPipe,
                              [out] LPVOID lpBuffer,
                              [in] DWORD  nBufferSize,
                              [out] DWORD * lpBytesRead,
                              [out] DWORD * lpTotalBytesAvail,
                              [out] DWORD * lpBytesLeftThisMessage);

    [

    entry("TransactNamedPipe")    ]
    BOOL TransactNamedPipeStr([in] HANDLE hNamedPipe,
                                     [in] LPSTR lpInBuffer,
                                     [in] DWORD  nInBufferSize,
                                     [out] LPSTR lpOutBuffer,
                                     [in] DWORD  nOutBufferSize,
                                     [out] DWORD * lpBytesRead,
                                     [in] DWORD  lpOverlapped); //OVERLAPPED structure.
    [

    entry("CreateMailslotA")    ]
    HANDLE A_CreateMailslot([in] LPSTR lpName,
                                 [in] DWORD  nMaxMessageSize,
                                 [in] DWORD  lReadTimeout,
                                 [in] LPVOID lpSecurityAttributes);
    [

    entry("CreateMailslotW")    ]
    HANDLE W_CreateMailslot([in] LPWSTR lpName,
                                 [in] DWORD  nMaxMessageSize,
                                 [in] DWORD  lReadTimeout,
                                 [in] LPVOID lpSecurityAttributes);


    [

    entry("CreateMailslotA")    ]
    HANDLE A_CreateMailslotSec([in] LPSTR lpName,
                                    [in] DWORD  nMaxMessageSize,
                                    [in] DWORD  lReadTimeout,
                                    [in] SECURITY_ATTRIBUTES lpSecurityAttributes);

    [

    entry("CreateMailslotW")    ]
    HANDLE W_CreateMailslotSec([in] LPWSTR lpName,
                                    [in] DWORD  nMaxMessageSize,
                                    [in] DWORD  lReadTimeout,
                                    [in] SECURITY_ATTRIBUTES lpSecurityAttributes);

    [

    entry("GetMailslotInfo")    ]
    BOOL GetMailslotInfo(
        [in] HANDLE hMailslot,
        [in] DWORD * lpMaxMessageSize,
        [in] DWORD * lpNextSize,
        [in] DWORD * lpMessageCount,
        [in] DWORD * lpReadTimeout);

    [

    entry("SetMailslotInfo")    ]
    BOOL SetMailslotInfo(
        [in] HANDLE hMailslot,
        [in] DWORD  lReadTimeout); //0 or MAILSLOT_WAIT_FOREVER

    [

    entry("CreateMutexA")    ]
    long A_CreateMutex(
		[in] LPVOID lpMutexAttributes,
        [in] BOOL bInitialOwner,
        [in] LPSTR lpName);

    [

    entry("CreateMutexW")    ]
    long W_CreateMutex(
		[in] LPVOID lpMutexAttributes,
        [in] BOOL bInitialOwner,
        [in] LPWSTR lpName);


    [

    entry("CreateMutexA")    ]
    long A_CreateMutexSec(
        [in] LPVOID  lpMutexAttributes,
        [in] BOOL bInitialOwner,
        [in] LPSTR lpName);

    [

    entry("CreateMutexW")    ]
    long W_CreateMutexSec(
        [in] LPVOID  lpMutexAttributes,
        [in] BOOL bInitialOwner,
        [in] LPWSTR lpName);


    [

    entry("OpenMutexA")    ]
    long A_OpenMutex(
        [in] DWORD  dwDesiredAccess,
        [in] BOOL bInheritHandle,
        [in] LPSTR lpName);

    [

    entry("OpenMutexW")    ]
    long W_OpenMutex(
        [in] DWORD  dwDesiredAccess,
        [in] BOOL bInheritHandle,
        [in] LPWSTR lpName);

    [

    entry("FlushViewOfFile")    ]
    BOOL FlushViewOfFile(
        [in] LPVOID lpBaseAddress,
        [in] DWORD dwNumberOfBytesToFlush);

    [

    entry("MapViewOfFile")    ]
    LONG MapViewOfFile(
        [in] HANDLE hFileMappingObject,
        [in] DWORD dwDesiredAccess, //FILE_MAP
        [in] DWORD dwFileOffsetHigh,
        [in] DWORD dwFileOffsetLow,
        [in] DWORD dwNumberOfBytesToMap);

    [

    entry("MapViewOfFileEx")    ]
    LONG MapViewOfFileEx(
        [in] HANDLE hFileMappingObject,
        [in] DWORD dwDesiredAccess,
        [in] DWORD dwFileOffsetHigh,
        [in] DWORD dwFileOffsetLow,
        [in] DWORD dwNumberOfBytesToMap,
        [in] LPVOID lpBaseAddress);

    [

    entry("UnmapViewOfFile")    ]
    BOOL UnmapViewOfFile([in] LPVOID lpBaseAddress);

    [

    entry("GetNamedPipeHandleStateA")    ]
    BOOL A_GetNamedPipeHandleState([in] HANDLE hNamedPipe,
                                        [out] DWORD * lpState, //PIPE_NOWAIT PIPE_READMODE_MESSAGE
                                        [out] DWORD * lpCurInstances,
                                        [out] DWORD * lpMaxCollectionCount,
                                        [out] DWORD * lpCollectDataTimeout,
                                        [out] LPSTR lpUserName,
                                        [in] DWORD nMaxUserNameSize);
    [

    entry("GetNamedPipeHandleStateW")    ]
    BOOL W_GetNamedPipeHandleState([in] HANDLE hNamedPipe,
                                        [out] DWORD * lpState, //PIPE_NOWAIT PIPE_READMODE_MESSAGE
                                        [out] DWORD * lpCurInstances,
                                        [out] DWORD * lpMaxCollectionCount,
                                        [out] DWORD * lpCollectDataTimeout,
                                        [out] LPSTR lpUserName,
                                        [in] DWORD nMaxUserNameSize);

    [

    entry("CallNamedPipeA")    ]
    BOOL A_CallNamedPipe([in] LPSTR lpNamedPipeName,
                              [in] LPVOID lpInBuffer,
                              [in] DWORD nInBufferSize,
                              [in] LPVOID lpOutBuffer,
                              [in] DWORD nOutBufferSize,
                              [out] DWORD * lpBytesRead,
                              [in] DWORD nTimeOut); //NMPWAIT_NOWAIT NMPWAIT_WAIT_FOREVER NMPWAIT_USE_DEFAULT_WAIT

    [

    entry("CallNamedPipeW")    ]
    BOOL W_CallNamedPipe([in] LPWSTR lpNamedPipeName,
                              [in] LPVOID lpInBuffer,
                              [in] DWORD nInBufferSize,
                              [in] LPVOID lpOutBuffer,
                              [in] DWORD nOutBufferSize,
                              [out] DWORD * lpBytesRead,
                              [in] DWORD nTimeOut); //NMPWAIT_NOWAIT NMPWAIT_WAIT_FOREVER NMPWAIT_USE_DEFAULT_WAIT


    [

    entry("WaitNamedPipeA")    ]
    BOOL A_WaitNamedPipe([in] LPSTR lpNamedPipeName,[in] DWORD nTimeOut);

    [

    entry("WaitNamedPipeW")    ]
    BOOL W_WaitNamedPipe([in] LPWSTR lpNamedPipeName,[in] DWORD nTimeOut);

    [

    entry("FindFirstChangeNotificationA")    ]
    HANDLE A_FindFirstChangeNotification(
        [in] LPSTR lpPathName,
        [in] BOOL bWatchSubtree,
        [in] long dwNotifyFilter);

    [

    entry("FindFirstChangeNotificationW")    ]
    HANDLE W_FindFirstChangeNotification(
        [in] LPWSTR lpPathName,
        [in] BOOL bWatchSubtree,
        [in] long dwNotifyFilter);


    [

    entry("FindNextChangeNotification")    ]
    BOOL FindNextChangeNotification([in] HANDLE hChangeHandle);

    [

    entry("FindCloseChangeNotification")    ]
    BOOL FindCloseChangeNotification([in] HANDLE hChangeHandle);
	
	[entry("ReadDirectoryChangesW")]
	BOOL ReadDirectoryChangesW(
        [in] HANDLE hDirectory,
        [in,out] LPVOID lpBuffer,
        [in] DWORD nBufferLength,
        [in] BOOL bWatchSubtree,
        [in] DWORD dwNotifyFilter,
        [out] DWORD *lpBytesReturned,
        [in] LPVOID lpOverlapped,
        [in] LPVOID lpCompletionRoutine); //Using FileIOCompletionRoutine calback

// *********

	[entry("Sleep")]
	VOID Sleep(
        [in] DWORD dwMilliseconds);

	[entry("SleepEx")]
	DWORD SleepEx(
        [in] DWORD dwMilliseconds,
        [in] BOOL bAlertable);

// *********

	[entry("WaitForSingleObject")]
	DWORD WaitForSingleObject(
        [in] HANDLE hHandle,
        [in] DWORD dwMilliseconds);

	[entry("WaitForSingleObjectEx")]
	DWORD WaitForSingleObjectEx(
        [in]  HANDLE hHandle,
        [in]  DWORD dwMilliseconds,
        [in]  BOOL bAlertable);

	[entry("WaitForMultipleObjects")]
	DWORD WaitForMultipleObjects(
        [in]  DWORD nCount,
        [in]  LPVOID lpHandles,
        [in]  BOOL bWaitAll,
        [in]  DWORD dwMilliseconds);

	[entry("WaitForMultipleObjectsEx")]
	DWORD WaitForMultipleObjectsEx(
        [in]  DWORD nCount,
        [in]  LPVOID lpHandles,
        [in]  BOOL bWaitAll,
        [in]  DWORD dwMilliseconds,
        [in]  BOOL bAlertable);

	// ****** Drives and Directories
   [

    entry("SetVolumeLabelA")    ]
    BOOL A_SetVolumeLabel([in] LPSTR lpRootPathName,[in] LPSTR lpVolumeName);

    [

    entry("SetVolumeLabelW")    ]
    BOOL W_SetVolumeLabel([in] LPWSTR lpRootPathName,[in] LPWSTR lpVolumeName);


    [

    entry("GetVolumeInformationA")    ]
    BOOL A_GetVolumeInformation([in] LPSTR lpRootPathName,
                                     [out] LPSTR lpVolumeNameBuffer,
                                     [in] DWORD nVolumeNameSize,
                                     [out] DWORD *lpVolumeSerialNumber,
                                     [out] DWORD *lpMaximumComponentLength,
                                     [out] DWORD *lpFileSystemFlags,
                                     [out] LPSTR lpFileSystemNameBuffer,
                                     [in] DWORD nFileSystemNameSize);
    [

    entry("GetVolumeInformationW")    ]
    BOOL W_GetVolumeInformation([in] LPWSTR lpRootPathName,
                                     [out] LPWSTR lpVolumeNameBuffer,
                                     [in] DWORD nVolumeNameSize,
                                     [out] DWORD *lpVolumeSerialNumber,
                                     [out] DWORD *lpMaximumComponentLength,
                                     [out] DWORD *lpFileSystemFlags,
                                     [out] LPWSTR lpFileSystemNameBuffer,
                                     [in] DWORD nFileSystemNameSize);

    [

    entry("GetLogicalDrives")    ]
    DWORD GetLogicalDrives(void);

    [

    entry("GetDriveTypeA")    ]
	UINT A_GetDriveType([in] LPSTR lpRootPathName);

    [

    entry("GetDriveTypeW")    ]
	UINT W_GetDriveType([in] LPWSTR lpRootPathName);


    [

    entry("GetLogicalDriveStringsA")    ]
	DWORD A_GetLogicalDriveStrings([in] DWORD nBufferLength,[out] LPSTR lpBuffer);

    [

    entry("GetLogicalDriveStringsW")    ]
	DWORD W_GetLogicalDriveStrings([in] DWORD nBufferLength,[out] LPWSTR lpBuffer);

    [

    entry("GetSystemDirectoryA")    ]
    UINT A_GetSystemDirectory([out] LPSTR lpszSysPath, [in] UINT cbSysPath);

    [

    entry("GetSystemDirectoryW")    ]
    UINT W_GetSystemDirectory([out] LPWSTR lpszSysPath, [in] UINT cbSysPath);

    [

    entry("GetSystemWindowsDirectoryA")    ]
	UINT A_GetSystemWindowsDirectory([out] LPSTR lpBuffer,[in] UINT uSize);

    [

    entry("GetSystemWindowsDirectoryW")    ]
	UINT W_GetSystemWindowsDirectory([out] LPWSTR lpBuffer,[in] UINT uSize);
/*
	[

    entry("GetSystemWow64DirectoryA"),
	helpstring("s the path of the system directory used by WOW64. This directory is not present on 32-bit Windows. XP only"),
	]
	UINT A_GetSystemWow64Directory([out] LPSTR lpBuffer,[in] UINT uSize);
*/
    [

    entry("GetTempPathA")    ]
    DWORD A_GetTempPath([in] DWORD nBufferLength,[out] LPSTR lpBuffer);

    [

    entry("GetTempPathW")    ]
    DWORD W_GetTempPath([in] DWORD nBufferLength,[out] LPWSTR lpBuffer);


    [

    entry("GetTempFileNameA")    ]
    UINT A_GetTempFileName([in] LPSTR lpPathName,
                                [in] LPSTR lpPrefixString,
                                [in] UINT uUnique,
                                [out] LPSTR lpTempFileName);

    [

    entry("GetTempFileNameW")    ]
    UINT W_GetTempFileName([in] LPWSTR lpPathName,
                                [in] LPWSTR lpPrefixString,
                                [in] UINT uUnique,
                                [out] LPWSTR lpTempFileName);

    [

    entry("GetWindowsDirectoryA")    ]
    UINT A_GetWindowsDirectory([out] LPSTR lpszSysPath, [in] UINT cbSysPath);

    [

    entry("GetWindowsDirectoryW")    ]
    UINT W_GetWindowsDirectory([out] LPWSTR lpszSysPath, [in] UINT cbSysPath);

    [

    entry("SetCurrentDirectoryA")    ]
    BOOL A_SetCurrentDirectory([in] LPSTR lpPathName);

    [

    entry("SetCurrentDirectoryW")    ]
    BOOL W_SetCurrentDirectory([in] LPWSTR lpPathName);

    [

    entry("GetCurrentDirectoryA")    ]
    DWORD A_GetCurrentDirectory([in] DWORD nBufferLength,[out] LPSTR lpBuffer);

    [

    entry("GetCurrentDirectoryW")    ]
    DWORD W_GetCurrentDirectory([in] DWORD nBufferLength,[out] LPWSTR lpBuffer);


    [

    entry("GetDiskFreeSpaceA")    ]
    BOOL A_GetDiskFreeSpace([in] LPSTR lpRootPathName,
                                 [out] DWORD *lpSectorsPerCluster,
                                 [out] DWORD *lpBytesPerSector,
                                 [out] DWORD *lpNumberOfFreeClusters,
                                 [out] DWORD *lpTotalNumberOfClusters);

    [

    entry("GetDiskFreeSpaceW")    ]
    BOOL W_GetDiskFreeSpace([in] LPWSTR lpRootPathName,
                                 [out] DWORD *lpSectorsPerCluster,
                                 [out] DWORD *lpBytesPerSector,
                                 [out] DWORD *lpNumberOfFreeClusters,
                                 [out] DWORD *lpTotalNumberOfClusters);

    [

    entry("GetDiskFreeSpaceExA")    ]
    BOOL A_GetDiskFreeSpaceEx([in] LPSTR lpDirectoryName,
                                 [out] CURRENCY *lpFreeBytesAvailable,
                                 [out] CURRENCY *lpTotalNumberOfBytes,
                                 [out] CURRENCY *lpTotalNumberOfFreeBytes);

    [

    entry("GetDiskFreeSpaceExW")    ]
    BOOL W_GetDiskFreeSpaceEx([in] LPWSTR lpDirectoryName,
                                 [out] CURRENCY *lpFreeBytesAvailable,
                                 [out] CURRENCY *lpTotalNumberOfBytes,
                                 [out] CURRENCY *lpTotalNumberOfFreeBytes);

    [

    entry("GetDiskFreeSpaceExA")    ]
    BOOL A_GetDiskFreeSpaceExAny([in] LPSTR lpDirectoryName,
                                 [out] LPVOID lpFreeBytesAvailable,
                                 [out] LPVOID lpTotalNumberOfBytes,
                                 [out] LPVOID lpTotalNumberOfFreeBytes);

    [

    entry("GetDiskFreeSpaceExW")    ]
    BOOL W_GetDiskFreeSpaceExAny([in] LPWSTR lpDirectoryName,
                                 [out] LPVOID lpFreeBytesAvailable,
                                 [out] LPVOID lpTotalNumberOfBytes,
                                 [out] LPVOID lpTotalNumberOfFreeBytes);
    [

    entry("CreateDirectoryA")    ]
    BOOL A_CreateDirectory([in] LPSTR lpPathName,[in] LPVOID lpSecurityAttributes);

    [

    entry("CreateDirectoryW")    ]
    BOOL W_CreateDirectory([in] LPWSTR lpPathName,[in] LPVOID lpSecurityAttributes);


    [

    entry("CreateDirectoryExA")    ]
    BOOL A_CreateDirectoryEx([in] LPSTR lpTemplateDirectory,
								[in] LPSTR lpNewDirectory,
								[in] LPVOID lpSecurityAttributes);

    [

    entry("CreateDirectoryExW")    ]
    BOOL W_CreateDirectoryEx([in] LPWSTR lpTemplateDirectory,
								[in] LPWSTR lpNewDirectory,
								[in] LPVOID lpSecurityAttributes);

    [

    entry("RemoveDirectoryA")    ]
    BOOL A_RemoveDirectory([in] LPSTR lpPathName);

    [

    entry("RemoveDirectoryW")    ]
    BOOL W_RemoveDirectory([in] LPWSTR lpPathName);

    [

    entry("GetFullPathNameA")    ]
    DWORD A_GetFullPathName([in] LPSTR lpFileName,
                                 [in] DWORD nBufferLength,
                                 [out] LPSTR lpBuffer,
                                 [out] LPSTR lpFilePart);
    [

    entry("GetFullPathNameW")    ]
    DWORD W_GetFullPathName([in] LPWSTR lpFileName,
                                 [in] DWORD nBufferLength,
                                 [out] LPWSTR lpBuffer,
                                 [out] LPWSTR lpFilePart);

    [

    entry("GetLongPathNameA")    ]
	DWORD A_GetLongPathName([in] LPSTR lpszShortPath,[out] LPSTR lpszLongPath,[in] DWORD cchBuffer);

    [

    entry("GetLongPathNameW")    ]
	DWORD W_GetLongPathName([in] LPWSTR lpszShortPath,[out] LPWSTR lpszLongPath,[in] DWORD cchBuffer);


    [

    entry("GetShortPathNameA")    ]
	DWORD A_GetShortPathName([in] LPSTR lpszLongPath,[out] LPSTR lpszShortPath,[in] DWORD cchBuffer);

    [

    entry("GetShortPathNameW")    ]
	DWORD W_GetShortPathName([in] LPWSTR lpszLongPath,[out] LPWSTR lpszShortPath,[in] DWORD cchBuffer);


    // ****** WIN.INI Support ******

    [

    entry("GetProfileIntA")    ]
    UINT A_GetProfileInt([in] LPSTR lpszSection,[in] LPSTR lpszEntry,[in] long iDef);

    [

    entry("GetProfileIntW")    ]
    UINT W_GetProfileInt([in] LPWSTR lpszSection,[in] LPWSTR lpszEntry,[in] long iDef);

    [

    entry("GetProfileStringA")    ]
    long A_GetProfileString([in] LPSTR lpszSection,
                            [in] LPSTR lpszEntry,
                            [in] LPSTR lpszDefault,
                            [out] LPSTR lpszReturnBuffer,
                            [in] long cbReturnBuffer);
    [

    entry("GetProfileStringW")    ]
    long W_GetProfileString([in] LPWSTR lpszSection,
                            [in] LPWSTR lpszEntry,
                            [in] LPWSTR lpszDefault,
                            [out] LPWSTR lpszReturnBuffer,
                            [in] long cbReturnBuffer);

    [

    entry("WriteProfileStringA")    ]
    BOOL A_WriteProfileString([in] LPSTR lpszSection,
                                      [in] LPSTR lpszEntry,
                                      [in] LPSTR lpszString);
    [

    entry("WriteProfileStringW")    ]
    BOOL W_WriteProfileString([in] LPWSTR lpszSection,
                                      [in] LPWSTR lpszEntry,
                                      [in] LPWSTR lpszString);

    [

    entry("GetPrivateProfileIntA")    ]
    UINT A_GetPrivateProfileInt([in] LPSTR lpszSection,
                                [in] LPSTR lpszEntry,
                                [in] long iDef,
                                [in] LPSTR lpszFilename);
    [

    entry("GetPrivateProfileIntW")    ]
    UINT W_GetPrivateProfileInt([in] LPWSTR lpszSection,
                                [in] LPWSTR lpszEntry,
                                [in] long iDef,
                                [in] LPWSTR lpszFilename);

    [

    entry("GetPrivateProfileStringA")    ]
    long A_GetPrivateProfileString([in] LPSTR lpszSection,
                                           [in] LPSTR lpszEntry,
                                           [in] LPSTR lpszDefault,
                                           [out] LPSTR lpszReturnBuffer,
                                           [in] long cbReturnBuffer,
                                           [in] LPSTR lpszFilename);

    [

    entry("GetPrivateProfileStringW")    ]
    long W_GetPrivateProfileString([in] LPWSTR lpszSection,
                                           [in] LPWSTR lpszEntry,
                                           [in] LPWSTR lpszDefault,
                                           [out] LPWSTR lpszReturnBuffer,
                                           [in] long cbReturnBuffer,
                                           [in] LPWSTR lpszFilename);

    [

    entry("WritePrivateProfileStringA")    ]
    BOOL A_WritePrivateProfileString([in] LPSTR lpszSection,
                                             [in] LPSTR lpszEntry,
                                             [in] LPSTR lpszString,
                                             [in] LPSTR lpszFilename);

    [

    entry("WritePrivateProfileStringW")    ]
    BOOL W_WritePrivateProfileString([in] LPWSTR lpszSection,
                                             [in] LPWSTR lpszEntry,
                                             [in] LPWSTR lpszString,
                                             [in] LPWSTR lpszFilename);

	//==========

    [

    entry("GetComputerNameA")    ]
    BOOL A_GetComputerName([out] LPSTR lpBuffer,[in,out] DWORD * nSize);

    [

    entry("GetComputerNameW")    ]
    BOOL W_GetComputerName([out] LPWSTR lpBuffer,[in,out] DWORD * nSize);

    [

    entry("SetComputerNameA")    ]
    BOOL A_SetComputerName([in] LPSTR lpComputerName);
	
    [

    entry("SetComputerNameW")    ]
    BOOL W_SetComputerName([in] LPWSTR lpComputerName);

	// ****** Time support
    [

    entry("GetSystemTime")    ]
    void GetSystemTime([out] LPVOID lpSystemTime);

    [

    entry("SetSystemTime")    ]
    BOOL SetSystemTime([in] LPVOID lpSystemTime);

    [

    entry("GetLocalTime")    ]
    void GetLocalTime([out] LPVOID lpSystemTime);

    [

    entry("SetLocalTime")    ]
    BOOL SetLocalTime([in] LPVOID lpSystemTime);

    [

    entry("GetSystemInfo")    ]
    void GetSystemInfo([out] LPVOID lpSystemInfo);

    // ***** Routines to convert back and forth between system time and file time

    [

    entry("SystemTimeToFileTime")    ]
    BOOL SystemTimeToFileTime([in] LPVOID lpSystemTime,[out] LPVOID lpFileTime);

    [

    entry("FileTimeToLocalFileTime")    ]
    BOOL FileTimeToLocalFileTime([in] LPVOID  lpFileTime,[out] LPVOID  lpLocalFileTime);

    [

    entry("LocalFileTimeToFileTime")    ]
    BOOL LocalFileTimeToFileTime([in] LPVOID  lpLocalFileTime,[out] LPVOID  lpFileTime);

    [

    entry("FileTimeToSystemTime")    ]
    BOOL FileTimeToSystemTime([in] LPVOID  lpFileTime,[out] LPVOID lpSystemTime);

    [

    entry("CompareFileTime")    ]
    LONG CompareFileTime([in] LPVOID  lpFileTime1,[in] LPVOID  lpFileTime2);

    [

    entry("FileTimeToDosDateTime")    ]
    BOOL FileTimeToDosDateTime([in] LPVOID  lpFileTime,
                               [out] WORD * lpFatDate,
                               [out] WORD * lpFatTime);

    [

    entry("DosDateTimeToFileTime")    ]
    BOOL DosDateTimeToFileTime([in] WORD wFatDate,
                               [in] WORD wFatTime,
                               [out] LPVOID  lpFileTime);

    [

    entry("GetTickCount")    ]
    DWORD GetTickCount(void);

    [

    entry("GetTickCount")    ]
    DWORD GetCurrentTime(void);

    [

    entry("SetSystemTimeAdjustment")    ]
    BOOL SetSystemTimeAdjustment([in] DWORD dwTimeAdjustment,[in] BOOL bTimeAdjustmentDisabled);

    [

    entry("GetSystemTimeAdjustment")    ]
    BOOL GetSystemTimeAdjustment([out] DWORD * lpTimeAdjustment,
                                 [out] DWORD * lpTimeIncrement,
                                 [out] BOOL  * lpTimeAdjustmentDisabled);

    [

    entry("GetSystemTimeAsFileTime")    ]
	void GetSystemTimeAsFileTime([out] LPVOID lpSystemTimeAsFileTime);

    [

    entry("GetTimeZoneInformation")    ]
    DWORD A_GetTimeZoneInformation([out] LPVOID lpTimeZoneInformation);

    [

    entry("GetTimeZoneInformation")    ]
    DWORD W_GetTimeZoneInformation([out] LPVOID lpTimeZoneInformation);

    [

    entry("SetTimeZoneInformation")    ]
    BOOL A_SetTimeZoneInformation([in] LPVOID lpTimeZoneInformation);

    [

    entry("SetTimeZoneInformation")    ]
    BOOL W_SetTimeZoneInformation([in] LPVOID lpTimeZoneInformation);

    // *********** Performance counter 
    [

    entry("QueryPerformanceCounter")    ]
    BOOL QueryPerformanceCounterAny([out] LPVOID lpPerformanceCount);

    [

    entry("QueryPerformanceCounter")    ]
    BOOL QueryPerformanceCounter([out] CURRENCY *lpPerformanceCount);

    [

    entry("QueryPerformanceFrequency")    ]
    BOOL QueryPerformanceFrequencyAny([out] LPVOID lpFrequency);

    [

    entry("QueryPerformanceFrequency")    ]
    BOOL QueryPerformanceFrequency([out] CURRENCY *lpFrequency);

    // ****** Global memory management ******

    [

    entry("GlobalAlloc")    ]
    long GlobalAlloc([in] long fuAlloc, [in] long cbAlloc);

    [

    entry("GlobalReAlloc")    ]
    long GlobalReAlloc([in] long hglb, [in] long cbNewSize, [in] long fuAlloc);

    [

    entry("GlobalSize")    ]
    long GlobalSize([in] long hglb);

    [

    entry("GlobalFlags")    ]
    long GlobalFlags([in] long hglb);

    [

    entry("GlobalLock")    ]
    long GlobalLock([in] long hglb);

    [

    entry("GlobalHandle")    ]
    long GlobalHandle([in] LPVOID pMem);

    [

    entry("GlobalUnlock")    ]
    long GlobalUnlock([in] long hglb);

    [

    entry("GlobalFree")    ]
    long GlobalFree([in] long hglb);

    [

    entry("GlobalMemoryStatus")    ]
    void GlobalMemoryStatus([out] LPVOID  lpBuffer);

    [

    entry("GlobalMemoryStatusEx")    ]
    long GlobalMemoryStatusEx([in,out] LPVOID  lpBuffer);

	// ******* OS Version
    [

    entry("GetVersion")    ]
    long GetVersion();

    [

    entry("GetVersionExA")    ]
    long A_GetVersionEx([in,out] LPVOID lpVersionInformation);

    [

    entry("GetVersionExW")    ]
    long W_GetVersionEx([in,out] LPVOID lpVersionInformation);

	// ***********
    [

    entry("GetDateFormatA")    ]
	long A_GetDateFormat(
			[in] long Locale,               // locale , pass: byval 0&
			[in] DWORD dwFlags,             // options
			[in] LPVOID lpDate,		// date
			[in] LPSTR lpFormat,			// date format
			[out] LPSTR lpDateStr,          // formatted string buffer
			[in] long cchDate);            // size of buffer

    [

    entry("GetDateFormatW")    ]
	long W_GetDateFormat(
			[in] long Locale,               // locale , pass: byval 0&
			[in] DWORD dwFlags,             // options
			[in] LPVOID lpDate,		// date
			[in] LPWSTR lpFormat,			// date format
			[out] LPWSTR lpDateStr,          // formatted string buffer
			[in] long cchDate);               // size of buffer

    [

    entry("GetTimeFormatA")    ]
	long A_GetTimeFormat(
				[in] long Locale,				// locale
				[in] DWORD dwFlags,				// options
				[in] LPVOID lpTime,		// time
				[in] LPSTR lpFormat,			// time format string
				[out] LPSTR lpTimeStr,			// formatted string buffer
				[in] long cchTime);               // size of string buffer

    [

    entry("GetTimeFormatW")    ]
	long W_GetTimeFormat(
				[in] long Locale,				// locale
				[in] DWORD dwFlags,				// options
				[in] LPVOID lpTime,		// time
				[in] LPWSTR lpFormat,			// time format string
				[out] LPWSTR lpTimeStr,			// formatted string buffer
				[in] long cchTime);               // size of string buffer);

    [

    entry("GetNumberFormatA")    ]
	long A_GetNumberFormat(
			[in] long Locale,               // locale
			[in] DWORD dwFlags,             // options
			[in] LPSTR lpValue,				// input number string
			[in] LPVOID lpFormat,		// formatting information
			[out] LPSTR lpNumberStr,         // output buffer
			[in] long cchNumber);           // size of output buffer

    [

    entry("GetNumberFormatW")    ]
	long W_GetNumberFormat(
			[in] long Locale,               // locale
			[in] DWORD dwFlags,             // options
			[in] LPWSTR lpValue,			// input number string
			[in] LPVOID lpFormat,		// formatting information
			[out] LPWSTR lpNumberStr,        // output buffer
			[in] long cchNumber);           // size of output buffer

    [

    entry("GetCurrencyFormatA")    ]
	long A_GetCurrencyFormat(
			[in] long Locale,					// locale
			[in] DWORD dwFlags,					// operation option
			[in] LPSTR lpValue,					// input number string
			[in] LPVOID lpFormat,		// formatting information
			[out] LPSTR lpCurrencyStr,			// output buffer
			[in] long cchCurrency);             // size of output buffer


    [

    entry("GetCurrencyFormatW")    ]
	long W_GetCurrencyFormat(
			[in] long Locale,					// locale
			[in] DWORD dwFlags,					// operation option
			[in] LPWSTR lpValue,				// input number string
			[in] LPVOID lpFormat,		// formatting information
			[out] LPWSTR lpCurrencyStr,         // output buffer
			[in] long cchCurrency);             // size of output buffer
	
    [

    entry("GetSystemDefaultLangID")    ]
	short GetSystemDefaultLangID(void); //LANGID

    [

    entry("GetSystemDefaultLCID")    ]
	long GetSystemDefaultLCID(void); //LCID

    [

    entry("GetSystemDefaultUILanguage")    ]
	short GetSystemDefaultUILanguage(void);

    [

    entry("GetUserDefaultLangID")    ]
	short GetUserDefaultLangID(void);

    [

    entry("GetUserDefaultLCID")    ]
	long GetUserDefaultLCID(void);

    [

    entry("GetUserDefaultUILanguage")    ]
	short GetUserDefaultUILanguage(void);

	// ***********

    [

    entry("MultiByteToWideChar")    ]
	long MultiByteToWideChar(
				[in] long CodePage,         // code page
				[in] long dwFlags,         // character-type options
				[in] void * lpMultiByteStr,	// string to map
				[in] long cbMultiByte,      // number of bytes in string
				[out] LPWSTR lpWideCharStr, // wide-character buffer
				[in] long cchWideChar);     // size of buffer

    [

    entry("WideCharToMultiByte")    ]
	long WideCharToMultiByte(
				[in] long CodePage,            // code page
				[in] long dwFlags,            // performance and mapping flags
				[in] LPWSTR lpWideCharStr,    // wide-character string
				[in] long cchWideChar,          // number of chars in string.
				[out] void * lpMultiByteStr,     // buffer for new string
				[in] long cbMultiByte,          // size of buffer
				[in] long lpDefaultChar,     // default for unmappable chars
				[in] long * lpUsedDefaultChar);  // set when default char used

    [

    entry("WideCharToMultiByte")    ]
	long WideCharToMultiByteAny(
				[in] long CodePage,            // code page
				[in] long dwFlags,            // performance and mapping flags
				[in] void * lpWideCharStr,    // wide-character string
				[in] long cchWideChar,          // number of chars in string.
				[out] void * lpMultiByteStr,     // buffer for new string
				[in] long cbMultiByte,          // size of buffer
				[in] long lpDefaultChar,     // default for unmappable chars
				[in] long * lpUsedDefaultChar);  // set when default char used

    [

    entry("IsBadStringPtrA")    ]
    BOOL A_IsBadStringPtr(
        [in] LPSTR lpsz,
        [in] UINT ucchMax);

    [

    entry("IsBadStringPtrW")    ]
    BOOL W_IsBadStringPtr(
        [in] LPWSTR lpsz,
        [in] UINT ucchMax);

    [

    entry("IsBadStringPtrA")    ]
    BOOL A_IsBadStringPtrByPtr(
        [in] long * lpsz,
        [in] UINT ucchMax);

    [

    entry("IsBadStringPtrW")    ]
    BOOL W_IsBadStringPtrByPtr(
        [in] long * lpsz,
        [in] UINT ucchMax);

    [

    entry("IsBadReadPtr")    ]
	BOOL IsBadReadPtr(
		[in] long * lp,
		[in] long ucb);

    [

    entry("IsBadWritePtr")    ]
	BOOL IsBadWritePtr(
		[in] long * lp,
		[in] long ucb);

    [

    entry("VirtualAlloc")    ]
	long VirtualAlloc(
		[in] long * lpAddress,
		[in] long dwSize,
		[in] DWORD flAllocationType,
		[in] DWORD flProtect);

    [

    entry("VirtualFree")    ]
	BOOL VirtualFree(
		[in] long * lpAddress,
		[in] long dwSize,
		[in] DWORD dwFreeType);

    [

    entry("VirtualLock")    ]
	BOOL VirtualLock(
		[in] long * lpAddress,
		[in] long dwSize);

    [

    entry("VirtualUnlock")    ]
	BOOL VirtualUnlock(
		[in] long * lpAddress,
		[in] long dwSize);

    [entry("MulDiv")]
	BOOL MulDiv(
		[in] long nNumber,
		[in] long nNumerator,
		[in] long nDenominator);

}; //End Kernel