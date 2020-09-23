
[
	helpstring("WININET Functions"),
	dllname("WININET.DLL")
]
module WinInet {

	[
	entry("InternetTimeFromSystemTimeA")
	]
	BOOL A_InternetTimeFromSystemTime(
		[in]  SYSTEMTIME *pst,  // input GMT time
		[in]  DWORD dwRFC,            // RFC format
		[out] LPSTR lpszTime,         // output string buffer
		[in]  DWORD cbTime);          // output buffer size

	[
	entry("InternetTimeFromSystemTimeW")
	]
	BOOL W_InternetTimeFromSystemTime(
		[in]  SYSTEMTIME *pst,  // input GMT time
		[in]  DWORD dwRFC,            // RFC format
		[out] LPWSTR lpszTime,         // output string buffer
		[in]  DWORD cbTime);          // output buffer size

	[
	entry("InternetTimeToSystemTimeA")
	]
	BOOL A_InternetTimeToSystemTime(
		[in]  LPSTR lpszTime,         // NULL terminated string
		[out] SYSTEMTIME *pst,         // output in GMT time
		[in]  DWORD dwReserved);

	[
	entry("InternetTimeToSystemTimeW")
	]
	BOOL W_InternetTimeToSystemTime(
		[in]  LPWSTR lpszTime,         // NULL terminated string
		[out] SYSTEMTIME *pst,         // output in GMT time
		[in]  DWORD dwReserved);

	[
	entry("InternetCrackUrlA")
	]
	BOOL A_InternetCrackUrl(
		[in] LPSTR lpszUrl,
		[in] DWORD dwUrlLength,
		[in] DWORD dwFlags,
		[in,out] LPVOID lpUrlComponents);

	[
	entry("InternetCrackUrlW")
	]
	BOOL W_InternetCrackUrl(
		[in] LPWSTR lpszUrl,
		[in] DWORD dwUrlLength,
		[in] DWORD dwFlags,
		[in,out] LPVOID lpUrlComponents);

	[entry("InternetCreateUrlA")]
	BOOL A_InternetCreateUrl(
    [in] LPVOID lpUrlComponents,
    [in] DWORD dwFlags,
    [out] LPSTR lpszUrl,
    [in,out] LPDWORD lpdwUrlLength
    );

	[entry("InternetCreateUrlW")]
	BOOL W_InternetCreateUrl(
    [in] LPVOID lpUrlComponents,
    [in] DWORD dwFlags,
    [out] LPWSTR lpszUrl,
    [in,out] LPDWORD lpdwUrlLength
    );

	[entry("InternetCanonicalizeUrlA")]
	BOOL A_InternetCanonicalizeUrl(
    [in] LPSTR lpszUrl,
    [out] LPSTR lpszBuffer,
    [in,out] LPDWORD lpdwBufferLength,
    [in] DWORD dwFlags
    );

	[entry("InternetCanonicalizeUrlW")]
	BOOL W_InternetCanonicalizeUrl(
    [in] LPWSTR lpszUrl,
    [out] LPWSTR lpszBuffer,
    [in,out] LPDWORD lpdwBufferLength,
    [in] DWORD dwFlags
    );

	[entry("InternetCombineUrlA")]
	BOOL A_InternetCombineUrl(
    [in] LPSTR lpszBaseUrl,
    [in] LPSTR lpszRelativeUrl,
    [out] LPSTR lpszBuffer,
    [in,out] LPDWORD lpdwBufferLength,
    [in] DWORD dwFlags
    );

	[entry("InternetCombineUrlW")]
	BOOL W_InternetCombineUrl(
    [in] LPWSTR lpszBaseUrl,
    [in] LPWSTR lpszRelativeUrl,
    [out] LPWSTR lpszBuffer,
    [in,out] LPDWORD lpdwBufferLength,
    [in] DWORD dwFlags
    );

	[entry("InternetOpenA")]
	long A_InternetOpen(
    [in] LPSTR lpszAgent,
    [in] DWORD dwAccessType,
    [in] LPSTR lpszProxy,
    [in] LPSTR lpszProxyBypass,
    [in] DWORD dwFlags
    );

	[entry("InternetOpenW")]
	long W_InternetOpen(
    [in] LPWSTR lpszAgent,
    [in] DWORD dwAccessType,
    [in] LPWSTR lpszProxy,
    [in] LPWSTR lpszProxyBypass,
    [in] DWORD dwFlags
    );
	
	[entry("InternetCloseHandle")]
	BOOL InternetCloseHandle(
    [in] HINTERNET hInternet
    );

	[entry("InternetConnectA")]
	long A_InternetConnect(
    [in] HINTERNET hInternet,
    [in] LPSTR lpszServerName,
    [in] INTERNET_PORT nServerPort,
    [in] LPSTR lpszUserName,
    [in] LPSTR lpszPassword,
    [in] DWORD dwService,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("InternetConnectW")]
	long W_InternetConnect(
    [in] HINTERNET hInternet,
    [in] LPWSTR lpszServerName,
    [in] INTERNET_PORT nServerPort,
    [in] LPWSTR lpszUserName,
    [in] LPWSTR lpszPassword,
    [in] DWORD dwService,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("InternetOpenUrlA")]
	long A_InternetOpenUrl(
    [in] HINTERNET hInternet,
    [in] LPSTR lpszUrl,
    [in] LPSTR lpszHeaders,
    [in] DWORD dwHeadersLength,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("InternetOpenUrlW")]
	long W_InternetOpenUrl(
    [in] HINTERNET hInternet,
    [in] LPWSTR lpszUrl,
    [in] LPWSTR lpszHeaders,
    [in] DWORD dwHeadersLength,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("InternetReadFile")]
	BOOL InternetReadFile(
    [in] HINTERNET hFile,
    [out] LPVOID lpBuffer,
    [in] DWORD dwNumberOfBytesToRead,
    [out] LPDWORD lpdwNumberOfBytesRead
    );
	[entry("InternetReadFile")]
	BOOL InternetReadFileStr(
    [in] HINTERNET hFile,
    [out] LPSTR lpBuffer,
    [in] DWORD dwNumberOfBytesToRead,
    [out] LPDWORD lpdwNumberOfBytesRead
    );

	[entry("InternetReadFileExA")]
	long A_InternetReadFileEx(
    [in] HINTERNET hFile,
    [out] LPVOID lpBuffersOut, //Array of internet buffers
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );
	[entry("InternetReadFileExW")]
	long W_InternetReadFileEx(
    [in] HINTERNET hFile,
    [out] LPVOID lpBuffersOut,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("InternetSetFilePointer")]
	long InternetSetFilePointer(
    [in] HINTERNET hFile,
    [in] LONG  lDistanceToMove,
    [in] LPVOID pReserved,
    [in] DWORD dwMoveMethod,
    [in] DWORD dwContext
    );

	[entry("InternetWriteFile")]
	BOOL InternetWriteFile(
    [in] HINTERNET hFile,
    [in] LPCVOID lpBuffer,
    [in] DWORD dwNumberOfBytesToWrite,
    [out] LPDWORD lpdwNumberOfBytesWritten
    );

	[entry("InternetQueryDataAvailable")]
	BOOL InternetQueryDataAvailable(
    [in] HINTERNET hFile,
    [out] LPDWORD lpdwNumberOfBytesAvailable,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("InternetFindNextFileA")]
	BOOL A_InternetFindNextFile(
    [in] HINTERNET hFind,
    [out] LPVOID lpvFindData
    );

	[entry("InternetFindNextFileW")]
	BOOL W_InternetFindNextFile(
    [in] HINTERNET hFind,
    [out] LPVOID lpvFindData
    );

	[entry("InternetQueryOptionA")]
	BOOL A_InternetQueryOption(
    [in] HINTERNET hInternet,
    [in] DWORD dwOption,
    [out] LPVOID lpBuffer,
    [in,out] LPDWORD lpdwBufferLength
    );

	[entry("InternetQueryOptionW")]
	BOOL W_InternetQueryOption(
    [in] HINTERNET hInternet,
    [in] DWORD dwOption,
    [out] LPVOID lpBuffer,
    [in,out] LPDWORD lpdwBufferLength
    );

	[entry("InternetSetOptionA")]
	BOOL A_InternetSetOption(
    [in] HINTERNET hInternet,
    [in] DWORD dwOption,
    [in] LPVOID lpBuffer,
    [in] DWORD dwBufferLength
    );

	[entry("InternetSetOptionW")]
	BOOL W_InternetSetOption(
    [in] HINTERNET hInternet,
    [in] DWORD dwOption,
    [in] LPVOID lpBuffer,
    [in] DWORD dwBufferLength
    );

	[entry("InternetSetOptionExA")]
	BOOL A_InternetSetOptionEx(
    [in] HINTERNET hInternet,
    [in] DWORD dwOption,
    [in] LPVOID lpBuffer,
    [in] DWORD dwBufferLength,
    [in] DWORD dwFlags
    );
	[entry("InternetSetOptionExW")]
	BOOL W_InternetSetOptionEx(
    [in] HINTERNET hInternet,
    [in] DWORD dwOption,
    [in] LPVOID lpBuffer,
    [in] DWORD dwBufferLength,
    [in] DWORD dwFlags
    );

	[entry("InternetLockRequestFile")]
	BOOL InternetLockRequestFile(
    [in]  HINTERNET hInternet,
    [out] HANDLE * lphLockRequestInfo
    );

	[entry("InternetUnlockRequestFile")]
	BOOL InternetUnlockRequestFile(
    [in] HANDLE hLockRequestInfo
    );

	[entry("InternetGetLastResponseInfoA")]
	BOOL A_InternetGetLastResponseInfo(
    [out] LPDWORD lpdwError,
    [out] LPSTR lpszBuffer,
    [in,out]  LPDWORD lpdwBufferLength
    );
	[entry("InternetGetLastResponseInfoW")]
	BOOL W_InternetGetLastResponseInfo(
    [out] LPDWORD lpdwError,
    [out] LPWSTR lpszBuffer,
    [in,out]  LPDWORD lpdwBufferLength
    );

	[entry("InternetSetStatusCallback")]
	long InternetSetStatusCallback(
    [in] HINTERNET hInternet,
    [in] long lpfnInternetCallback
    );

//
// prototypes
//

	[entry("FtpFindFirstFileA")]
	long A_FtpFindFirstFile(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszSearchFile,
    [out] LPVOID lpFindFileData, //WIN32_FIND_DATAA
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );
	[entry("FtpFindFirstFileW")]
	long W_FtpFindFirstFile(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszSearchFile,
    [out] LPVOID lpFindFileData, //WIN32_FIND_DATAW
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("FtpGetFileA")]
	BOOL A_FtpGetFile(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszRemoteFile,
    [in] LPSTR lpszNewFile,
    [in] BOOL fFailIfExists,
    [in] DWORD dwFlagsAndAttributes,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );
	[entry("FtpGetFileW")]
	BOOL W_FtpGetFile(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszRemoteFile,
    [in] LPWSTR lpszNewFile,
    [in] BOOL fFailIfExists,
    [in] DWORD dwFlagsAndAttributes,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("FtpPutFileA")]
	BOOL A_FtpPutFile(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszLocalFile,
    [in] LPSTR lpszNewRemoteFile,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );
	[entry("FtpPutFileW")]
	BOOL W_FtpPutFile(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszLocalFile,
    [in] LPWSTR lpszNewRemoteFile,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("FtpDeleteFileA")]
	BOOL A_FtpDeleteFile(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszFileName
    );
	[entry("FtpDeleteFileW")]
	BOOL W_FtpDeleteFile(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszFileName
    );

	[entry("FtpRenameFileA")]
	BOOL A_FtpRenameFile(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszExisting,
    [in] LPSTR lpszNew
    );
	[entry("FtpRenameFileW")]
	BOOL W_FtpRenameFile(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszExisting,
    [in] LPWSTR lpszNew
    );

	[entry("FtpOpenFileA")]
	long A_FtpOpenFile(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszFileName,
    [in] DWORD dwAccess,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("FtpOpenFileW")]
	long W_FtpOpenFile(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszFileName,
    [in] DWORD dwAccess,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("FtpCreateDirectoryA")]
	BOOL A_FtpCreateDirectory(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszDirectory
    );
	[entry("FtpCreateDirectoryW")]
	BOOL W_FtpCreateDirectory(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszDirectory
    );

	[entry("FtpRemoveDirectoryA")]
	BOOL A_FtpRemoveDirectory(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszDirectory
    );
	[entry("FtpRemoveDirectoryW")]
	BOOL W_FtpRemoveDirectory(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszDirectory
    );

	[entry("FtpSetCurrentDirectoryA")]
	BOOL A_FtpSetCurrentDirectory(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszDirectory
    );
	[entry("FtpSetCurrentDirectoryW")]
	BOOL W_FtpSetCurrentDirectory(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszDirectory
    );

	[entry("FtpGetCurrentDirectoryA")]
	BOOL A_FtpGetCurrentDirectory(
    [in] HINTERNET hConnect,
    [out] LPSTR lpszCurrentDirectory,
    [in,out]  LPDWORD lpdwCurrentDirectory
    );
	[entry("FtpGetCurrentDirectoryW")]
	BOOL W_FtpGetCurrentDirectory(
    [in] HINTERNET hConnect,
    [out] LPWSTR lpszCurrentDirectory,
    [in,out]  LPDWORD lpdwCurrentDirectory
    );

//
// HTTP
//
	[entry("HttpOpenRequestA")]
	long A_HttpOpenRequest(
    [in] HINTERNET hConnect,
    [in] LPSTR lpszVerb,
    [in] LPSTR lpszObjectName,
    [in] LPSTR lpszVersion,
    [in] LPSTR lpszReferrer,
    [in] long * lplpszAcceptTypes,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );
	[entry("HttpOpenRequestW")]
	long W_HttpOpenRequest(
    [in] HINTERNET hConnect,
    [in] LPWSTR lpszVerb,
    [in] LPWSTR lpszObjectName,
    [in] LPWSTR lpszVersion,
    [in] LPWSTR lpszReferrer,
    [in] long * lplpszAcceptTypes,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("HttpAddRequestHeadersA")]
	BOOL A_HttpAddRequestHeaders(
    [in] HINTERNET hRequest,
    [in] LPSTR lpszHeaders,
    [in] DWORD dwHeadersLength,
    [in] DWORD dwModifiers
    );
	[entry("HttpAddRequestHeadersW")]
	BOOL W_HttpAddRequestHeaders(
    [in] HINTERNET hRequest,
    [in] LPWSTR lpszHeaders,
    [in] DWORD dwHeadersLength,
    [in] DWORD dwModifiers
    );

	[entry("HttpSendRequestA")]
	BOOL A_HttpSendRequest(
    [in] HINTERNET hRequest,
    [in] LPSTR lpszHeaders,
    [in] DWORD dwHeadersLength,
    [in] LPVOID lpOptional,
    [in] DWORD dwOptionalLength
    );
	[entry("HttpSendRequestW")]
	BOOL W_HttpSendRequest(
    [in] HINTERNET hRequest,
    [in] LPWSTR lpszHeaders,
    [in] DWORD dwHeadersLength,
    [in] LPVOID lpOptional,
    [in] DWORD dwOptionalLength
    );

	[entry("HttpSendRequestExA")]
	BOOL A_HttpSendRequestEx(
    [in] HINTERNET hRequest,
    [in] LPVOID lpBuffersIn,
    [out] LPVOID lpBuffersOut,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );
	[entry("HttpSendRequestExW")]
	BOOL W_HttpSendRequestEx(
    [in] HINTERNET hRequest,
    [in] LPVOID lpBuffersIn,
    [out] LPVOID lpBuffersOut,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("HttpEndRequestA")]
	long A_HttpEndRequest(
    [in] HINTERNET hRequest,
    [out] LPVOID lpBuffersOut,
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );
	[entry("HttpEndRequestW")]
	long W_HttpEndRequest(
    [in] HINTERNET hRequest,
    [out] LPVOID lpBuffersOut, //Reserved. Must be NULL.
    [in] DWORD dwFlags,
    [in] DWORD dwContext
    );

	[entry("HttpQueryInfoA")]
	BOOL A_HttpQueryInfo(
    [in] HINTERNET hRequest,
    [in] DWORD dwInfoLevel,
    [in,out]  LPVOID lpBuffer,
    [in,out]  LPDWORD lpdwBufferLength,
    [in,out]  LPDWORD lpdwIndex
    );
	[entry("HttpQueryInfoW")]
	BOOL W_HttpQueryInfo(
    [in] HINTERNET hRequest,
    [in] DWORD dwInfoLevel,
    [in,out]  LPVOID lpBuffer,
    [in,out]  LPDWORD lpdwBufferLength,
    [in,out]  LPDWORD lpdwIndex
    );

//
// Cookie APIs
//

	[entry("InternetSetCookieA")]
	BOOL A_InternetSetCookie(
    [in] LPSTR lpszUrl,
    [in] LPSTR lpszCookieName,
    [in] LPSTR lpszCookieData
    );
	[entry("InternetSetCookieW")]
	BOOL W_InternetSetCookie(
    [in] LPSTR lpszUrl,
    [in] LPWSTR lpszCookieName,
    [in] LPWSTR lpszCookieData
    );

	[entry("InternetGetCookieA")]
	BOOL A_InternetGetCookie(
    [in] LPSTR lpszUrl,
    [in] LPSTR lpszCookieName,
    [out] LPSTR lpCookieData,
    [in,out]  LPDWORD lpdwSize
    );
	[entry("InternetGetCookieW")]
	BOOL W_InternetGetCookie(
    [in] LPWSTR lpszUrl,
    [in] LPWSTR lpszCookieName,
    [out] LPWSTR lpCookieData,
    [in,out]  LPDWORD lpdwSize
    );

//
// offline browsing
//
	[entry("InternetAttemptConnect")]
	long InternetAttemptConnect(
    [in] DWORD dwReserved
    );

	[entry("InternetCheckConnectionA")]
	BOOL A_InternetCheckConnection(
    [in] LPSTR lpszUrl,
    [in] DWORD dwFlags,
    [in] DWORD dwReserved
    );
	[entry("InternetCheckConnectionW")]
	BOOL W_InternetCheckConnection(
    [in] LPWSTR lpszUrl,
    [in] DWORD dwFlags,
    [in] DWORD dwReserved
    );

//
// Internet UI
//

	[entry("InternetErrorDlg")]
	long InternetErrorDlg(
    [in] HWND hWnd,
    [in]  HINTERNET hRequest,
    [in] DWORD dwError,
    [in] DWORD dwFlags,
    [in,out]  LPVOID * lppvData
    );

	[entry("InternetConfirmZoneCrossingA")]
	long A_InternetConfirmZoneCrossing(
    [in] HWND hWnd,
    [in] LPSTR szUrlPrev,
    [in] LPSTR szUrlNew,
    [in] BOOL bPost
    );

	[entry("InternetConfirmZoneCrossingW")]
	long W_InternetConfirmZoneCrossing(
    [in] HWND hWnd,
    [in] LPWSTR szUrlPrev,
    [in] LPWSTR szUrlNew,
    [in] BOOL bPost
    );


//
// Cache APIs
//
	[entry("CreateUrlCacheEntryA")]
	BOOL A_CreateUrlCacheEntry(
    [in] LPSTR lpszUrlName,
    [in] DWORD dwExpectedFileSize,
    [in] LPSTR lpszFileExtension,
    [out] LPSTR lpszFileName,
    [in] DWORD dwReserved
    );
	[entry("CreateUrlCacheEntryW")]
	BOOL W_CreateUrlCacheEntry(
    [in] LPWSTR lpszUrlName,
    [in] DWORD dwExpectedFileSize,
    [in] LPWSTR lpszFileExtension,
    [out] LPWSTR lpszFileName,
    [in] DWORD dwReserved
    );

	[entry("CommitUrlCacheEntryA")]
	BOOL A_CommitUrlCacheEntry(
    [in] LPSTR lpszUrlName,
    [in] LPSTR lpszLocalFileName,
    [in] FILETIME ExpireTime,
    [in] FILETIME LastModifiedTime,
    [in] DWORD CacheEntryType,
    [in] LPBYTE lpHeaderInfo,
    [in] DWORD dwHeaderSize,
    [in] LPSTR lpszFileExtension,
    [in] DWORD dwReserved
    );
	[entry("CommitUrlCacheEntryW")]
	BOOL W_CommitUrlCacheEntry(
    [in] LPWSTR lpszUrlName,
    [in] LPWSTR lpszLocalFileName,
    [in] FILETIME ExpireTime,
    [in] FILETIME LastModifiedTime,
    [in] DWORD CacheEntryType,
    [in] LPBYTE lpHeaderInfo,
    [in] DWORD dwHeaderSize,
    [in] LPWSTR lpszFileExtension,
    [in] DWORD dwReserved
    );

	[entry("RetrieveUrlCacheEntryFileA")]
	BOOL A_RetrieveUrlCacheEntryFile(
    [in] LPSTR  lpszUrlName,
    [out] LPVOID lpCacheEntryInfo,
    [in,out]  LPDWORD lpdwCacheEntryInfoBufferSize,
    [in] DWORD dwReserved
    );
	[entry("RetrieveUrlCacheEntryFileW")]
	BOOL W_RetrieveUrlCacheEntryFile(
    [in] LPWSTR  lpszUrlName,
    [out] LPVOID lpCacheEntryInfo,
    [in,out]  LPDWORD lpdwCacheEntryInfoBufferSize,
    [in] DWORD dwReserved
    );

	[entry("UnlockUrlCacheEntryFileA")]
	BOOL A_UnlockUrlCacheEntryFile(
    [in] LPSTR lpszUrlName,
    [in] DWORD dwReserved
    );
	[entry("UnlockUrlCacheEntryFileW")]
	BOOL W_UnlockUrlCacheEntryFile(
    [in] LPWSTR lpszUrlName,
    [in] DWORD dwReserved
    );

	[entry("RetrieveUrlCacheEntryStreamA")]
	long A_RetrieveUrlCacheEntryStream(
    [in] LPSTR  lpszUrlName,
    [out] LPVOID lpCacheEntryInfo,
    [in,out]  LPDWORD lpdwCacheEntryInfoBufferSize,
    [in] BOOL fRandomRead,
    [in] DWORD dwReserved
    );
	[entry("RetrieveUrlCacheEntryStreamW")]
	long W_RetrieveUrlCacheEntryStream(
    [in] LPWSTR  lpszUrlName,
    [out] LPVOID lpCacheEntryInfo,
    [in,out] LPDWORD lpdwCacheEntryInfoBufferSize,
    [in] BOOL fRandomRead,
    [in] DWORD dwReserved
    );

	[entry("ReadUrlCacheEntryStream")]
	BOOL ReadUrlCacheEntryStream(
    [in] HANDLE hUrlCacheStream,
    [in] DWORD dwLocation,
    [in,out]  LPVOID lpBuffer,
    [in,out]  LPDWORD lpdwLen,
    [in] DWORD Reserved
    );

	[entry("UnlockUrlCacheEntryStream")]
	BOOL UnlockUrlCacheEntryStream(
    [in] HANDLE hUrlCacheStream,
    [in] DWORD Reserved
    );

	[entry("GetUrlCacheEntryInfoA")]
	BOOL A_GetUrlCacheEntryInfo(
    [in] LPSTR lpszUrlName,
    [out] LPVOID lpCacheEntryInfo,
    [in,out] LPDWORD lpdwCacheEntryInfoBufferSize
    );
	[entry("GetUrlCacheEntryInfoW")]
	BOOL W_GetUrlCacheEntryInfo(
    [in] LPWSTR lpszUrlName,
    [out] LPVOID lpCacheEntryInfo,
    [in,out]  LPDWORD lpdwCacheEntryInfoBufferSize
    );

	[entry("GetUrlCacheEntryInfoExA")]
	BOOL A_GetUrlCacheEntryInfoEx(
    [in] LPSTR lpszUrl,
    [out] LPVOID lpCacheEntryInfo,
    [in,out] LPDWORD lpdwCacheEntryInfoBufSize,
    [out] LPSTR      lpszReserved,  // must pass null
    [in,out] LPDWORD lpdwReserved,  // must pass null
    [in] LPVOID         lpReserved,    // must pass null
    [in] DWORD          dwFlags        // reserved 0
    );
	[entry("GetUrlCacheEntryInfoExW")]
	BOOL W_GetUrlCacheEntryInfoEx(
    [in] LPWSTR lpszUrl,
    [out] LPVOID lpCacheEntryInfo,
    [in,out] LPDWORD lpdwCacheEntryInfoBufSize,
    [out] LPWSTR      lpszReserved,  // must pass null
    [in,out] LPDWORD lpdwReserved,  // must pass null
    [in] LPVOID         lpReserved,    // must pass null
    [in] DWORD          dwFlags        // reserved
    );

	[entry("SetUrlCacheEntryInfoA")]
	BOOL A_SetUrlCacheEntryInfo(
    [in] LPSTR lpszUrlName,
    [in] LPVOID lpCacheEntryInfo,
    [in] DWORD dwFieldControl
    );
	[entry("SetUrlCacheEntryInfoW")]
	BOOL W_SetUrlCacheEntryInfoW(
    [in] LPWSTR lpszUrlName,
    [in] LPVOID lpCacheEntryInfo,
    [in] DWORD dwFieldControl
    );


//
// Cache Group Functions
//

	[entry("CreateUrlCacheGroup")]
	GROUPID CreateUrlCacheGroup(
    [in] DWORD  dwFlags,
    [in] LPVOID lpReserved  // must pass NULL
    );

	[entry("DeleteUrlCacheGroup")]
	BOOL DeleteUrlCacheGroup(
    [in]  GROUPID    GroupId,
    [in]  DWORD   dwFlags,       // must pass 0
    [in]  LPVOID  lpReserved     // must pass NULL
    );

	[entry("SetUrlCacheEntryGroupA")]
	BOOL A_SetUrlCacheEntryGroup(
    [in] LPSTR    lpszUrlName,
    [in] DWORD    dwFlags,
    [in] GROUPID     GroupId,
    [in] LPVOID   pbGroupAttributes, // must pass NULL
    [in] DWORD    cbGroupAttributes, // must pass 0
    [in] LPVOID   lpReserved         // must pass NULL
    );
	[entry("SetUrlCacheEntryGroupW")]
	BOOL W_SetUrlCacheEntryGroup(
    [in] LPWSTR    lpszUrlName,
    [in] DWORD    dwFlags,
    [in] GROUPID     GroupId,
    [in] LPVOID   pbGroupAttributes, // must pass NULL
    [in] DWORD    cbGroupAttributes, // must pass 0
    [in] LPVOID   lpReserved);         // must pass NULL

	[entry("FindFirstUrlCacheEntryExA")]
	long A_FindFirstUrlCacheEntryEx(
    [in]     LPSTR		lpszUrlSearchPattern,
    [in]     DWORD		dwFlags,
    [in]     DWORD		dwFilter,
    [in]     GROUPID	GroupId,
    [out]    LPVOID		lpFirstCacheEntryInfo,
    [in,out]  LPDWORD   lpdwFirstCacheEntryInfoBufferSize,
    [out]    LPVOID		lpGroupAttributes,     // must pass NULL
    [in,out]  LPDWORD   pcbGroupAttributes,    // must pass NULL
    [in]     LPVOID		lpReserved);             // must pass NULL
    
	[entry("FindFirstUrlCacheEntryExW")]	
	long W_FindFirstUrlCacheEntryEx(
    [in]     LPWSTR		lpszUrlSearchPattern,
    [in]     DWORD		dwFlags,
    [in]     DWORD		dwFilter,
    [in]     GROUPID	GroupId,
    [out]    LPVOID		lpFirstCacheEntryInfo,
    [in,out]  LPDWORD   lpdwFirstCacheEntryInfoBufferSize,
    [out]    LPVOID		lpGroupAttributes,     // must pass NULL
    [in,out]  LPDWORD   pcbGroupAttributes,    // must pass NULL
    [in]     LPVOID		lpReserved             // must pass NULL
    );

	[entry("FindNextUrlCacheEntryExA")]
	BOOL A_FindNextUrlCacheEntryEx(
    [in]     HANDLE    hEnumHandle,
    [out]    LPVOID    lpFirstCacheEntryInfo,
    [in,out]  LPDWORD  lpdwFirstCacheEntryInfoBufferSize,
    [out]    LPVOID    lpGroupAttributes,     // must pass NULL
    [in,out]  LPDWORD  pcbGroupAttributes,    // must pass NULL
    [in]     LPVOID    lpReserved             // must pass NULL
    );
	[entry("FindNextUrlCacheEntryExW")]
	BOOL W_FindNextUrlCacheEntryEx(
    [in]     HANDLE		hEnumHandle,
    [out]    LPVOID		lpFirstCacheEntryInfo,
    [in,out]  LPDWORD   lpdwFirstCacheEntryInfoBufferSize,
    [out]    LPVOID		lpGroupAttributes,     // must pass NULL
    [in,out]  LPDWORD   pcbGroupAttributes,    // must pass NULL
    [in]     LPVOID		lpReserved             // must pass NULL
    );

	[entry("FindFirstUrlCacheEntryA")]
	long A_FindFirstUrlCacheEntry(
    [in] LPSTR lpszUrlSearchPattern,
    [out] LPVOID lpFirstCacheEntryInfo,
    [in,out]  LPDWORD lpdwFirstCacheEntryInfoBufferSize
    );
	[entry("FindFirstUrlCacheEntryW")]
	long W_FindFirstUrlCacheEntry(
    [in] LPWSTR lpszUrlSearchPattern,
    [out] LPVOID lpFirstCacheEntryInfo,
    [in,out]  LPDWORD lpdwFirstCacheEntryInfoBufferSize
    );

	[entry("FindNextUrlCacheEntryA")]
	BOOL A_FindNextUrlCacheEntry(
    [in] HANDLE hEnumHandle,
    [out] LPVOID lpNextCacheEntryInfo,
    [in,out] LPDWORD lpdwNextCacheEntryInfoBufferSize
    );
	[entry("FindNextUrlCacheEntryW")]
	BOOL W_FindNextUrlCacheEntry(
    [in] HANDLE hEnumHandle,
    [out] LPVOID lpNextCacheEntryInfo,
    [in,out]  LPDWORD lpdwNextCacheEntryInfoBufferSize
    );

	[entry("FindCloseUrlCache")]
	BOOL FindCloseUrlCache(
    [in] HANDLE hEnumHandle
    );

	[entry("DeleteUrlCacheEntryA")]
	BOOL A_DeleteUrlCacheEntry(
    [in] LPSTR lpszUrlName);

	[entry("DeleteUrlCacheEntryW")]
	BOOL W_DeleteUrlCacheEntry(
    [in] LPWSTR lpszUrlName);

//
// Autodial APIs
//
	[entry("InternetDialA")]
	DWORD A_InternetDial(
    [in] HWND     hwndParent,
    [in] LPSTR   lpszConnectoid,
    [in] DWORD    dwFlags,
    [out] LPDWORD lpdwConnection,
    [in] DWORD    dwReserved);

	[entry("InternetDialW")]
	DWORD W_InternetDial(
    [in] HWND     hwndParent,
    [in] LPWSTR   lpszConnectoid,
    [in] DWORD    dwFlags,
    [out] LPDWORD lpdwConnection,
    [in] DWORD    dwReserved);

	[entry("InternetHangUp")]
	DWORD InternetHangUp(
    [in] DWORD    dwConnection,
    [in] DWORD    dwReserved);

	[entry("InternetGoOnlineA")]
	BOOL A_InternetGoOnline(
    [in] LPSTR   lpszURL,
    [in] HWND     hwndParent,
    [in] DWORD    dwFlags);
	[entry("InternetGoOnlineW")]
	BOOL W_InternetGoOnline(
    [in] LPWSTR   lpszURL,
    [in] HWND     hwndParent,
    [in] DWORD    dwFlags);

	[entry("InternetAutodial")]
	BOOL InternetAutodial(
    [in] DWORD    dwFlags,
    [in] DWORD    dwReserved);

	[entry("InternetAutodialHangup")]
	BOOL InternetAutodialHangup(
    [in] DWORD    dwReserved);

	[entry("InternetGetConnectedState")]
	BOOL InternetGetConnectedState(
    [out] LPDWORD  lpdwFlags,
    [in] DWORD    dwReserved);


}; //END WININET
