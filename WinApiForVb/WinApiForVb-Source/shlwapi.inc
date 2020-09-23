[
	helpstring("Windows light-weight utility APIs"),
    dllname("SHLWAPI.DLL")
]
module shlwapi {

//
//=============== Path Routines ===================================
//
//For LPWSTR's we need to add Chr&(0) to the end

    [entry("PathAddBackslashA")]
    LONG A_PathAddBackslash(
        [in, out] LPSTR pszPath);

    [entry("PathAddBackslashW")]
    LONG W_PathAddBackslash(
        [in, out] LPWSTR pszPath);

    [entry("PathAddExtensionA")]
    BOOL A_PathAddExtension(
        [in, out] LPSTR pszPath,
        [in] LPSTR pszExt);

    [entry("PathAddExtensionW")]
    BOOL W_PathAddExtension(
        [in, out] LPWSTR pszPath,
        [in] LPWSTR pszExt);

    [entry("PathAppendA")]
    BOOL A_PathAppend(
        [in, out]  LPSTR pszPath,
        [in] LPSTR pMore);

    [entry("PathAppendW")]
    BOOL W_PathAppend(
        [in, out]  LPWSTR pszPath,
        [in] LPWSTR pMore);

    [entry("PathBuildRootA")]
    LONG A_PathBuildRoot(
        [out] LPSTR pszRoot,
        [in] long iDrive);

    [entry("PathBuildRootW")]
    LONG W_PathBuildRoot(
        [out] LPWSTR pszRoot,
        [in] long iDrive);

    [entry("PathCanonicalizeA")]
    BOOL A_PathCanonicalize(
        [out] LPSTR pszBuf,
        [in] LPSTR pszPath);

    [entry("PathCanonicalizeW")]
    BOOL W_PathCanonicalize(
        [out] LPWSTR pszBuf,
        [in] LPWSTR pszPath);

    [entry("PathCombineA")]
    LONG A_PathCombine(
        [out] LPSTR pszDest,
        [in] LPSTR pszDir,
        [in] LPSTR pszFile);

    [entry("PathCombineW")]
    LONG W_PathCombine(
        [out] LPWSTR pszDest,
        [in] LPWSTR pszDir,
        [in] LPWSTR pszFile);

    [entry("PathCompactPathA")]
    BOOL A_PathAddEllipse(
        [in] LONG hDC,
        [in, out] LPSTR pszPath,
        [in] LONG dx);

    [entry("PathCompactPathW")]
    BOOL W_PathAddEllipse(
        [in] LONG hDC,
        [in, out] LPWSTR pszPath,
        [in] LONG dx);

    [entry("PathCompactPathExA")]
    BOOL A_PathAddEllipseEx(
        [out] LPSTR pszOut,
        [in] LPSTR pszSrc,
        [in] LONG cchMax,
        [in] LONG dwFlags);

    [entry("PathCompactPathExW")]
    BOOL W_PathAddEllipseEx(
        [out] LPWSTR pszOut,
        [in] LPWSTR pszSrc,
        [in] LONG cchMax,
        [in] LONG dwFlags);

    [entry("PathCommonPrefixA")]
    LONG A_PathCommonPrefix(
        [in] LPSTR pszFile1,
        [in] LPSTR pszFile2,
        [out] LPSTR achPath);

    [entry("PathCommonPrefixW")]
    LONG W_PathCommonPrefix(
        [in] LPWSTR pszFile1,
        [in] LPWSTR pszFile2,
        [out] LPWSTR achPath);

    [entry("PathFileExistsA")]
    BOOL A_PathFileExists(
        [in] LPSTR pszPath);

    [entry("PathFileExistsW")]
    BOOL W_PathFileExists(
        [in] LPWSTR pszPath);

    [entry("PathFindExtensionA")]
    LONG A_PathFindExtension(
        [in] LPSTR pszPath);

    [entry("PathFindExtensionW")]
    LONG W_PathFindExtension(
        [in] LPWSTR pszPath);

    [entry("PathFindFileNameA")]
    LONG A_PathFindFileName(
        [in] LPSTR pszPath);

    [entry("PathFindFileNameW")]
    LONG W_PathFindFileName(
        [in] LPWSTR pszPath);

    [entry("PathFindNextComponentA")]
    LONG A_PathFindNextComponent(
        [in] LPSTR pszPath);

    [entry("PathFindNextComponentW")]
    LONG W_PathFindNextComponent(
        [in] LPWSTR pszPath);

    [entry("PathFindOnPathA")]
    BOOL A_PathFindOnPath(
        [in] LPSTR pszPath,
        [in] LPSTR ppszOtherDirs);

    [entry("PathFindOnPathW")]
    BOOL W_PathFindOnPath(
        [in] LPWSTR pszPath,
        [in] LPWSTR ppszOtherDirs);

    [entry("PathGetArgsA")]
    LONG A_PathGetArgs(
        [in] LPSTR pszPath);

    [entry("PathGetArgsW")]
    LONG W_PathGetArgs(
        [in] LPWSTR pszPath);

    [entry("PathFindSuffixArrayA")]
    LONG A_PathFindSuffixArray(
        [in] LPSTR pszPath,
        [in] LPSTR apszSuffix,
        [in] LONG iArraySize);

    [entry("PathFindSuffixArrayW")]
    LONG W_PathFindSuffixArray(
        [in] LPWSTR pszPath,
        [in] LPWSTR apszSuffix,
        [in] long iArraySize);

    [entry("PathIsLFNFileSpecA")]
    BOOL A_PathIsLFNFileSpec(
        [in] LPSTR lpName);

    [entry("PathIsLFNFileSpecW")]
    BOOL W_PathIsLFNFileSpec(
        [in] LPWSTR lpName);

    [entry("PathGetCharTypeA")]
    GCTFlags A_PathGetCharType(
        [in] long ch);

    [entry("PathGetCharTypeW")]
    GCTFlags W_PathGetCharType(
        [in] long ch);

    [entry("PathGetDriveNumberA")]
    LONG A_PathGetDriveNumber(
        [in] LPSTR pszPath);

    [entry("PathGetDriveNumberW")]
    LONG W_PathGetDriveNumber(
        [in] LPWSTR pszPath);

    [entry("PathIsDirectoryA")]
    BOOL A_PathIsDirectory(
        [in] LPSTR pszPath);

    [entry("PathIsDirectoryW")]
    BOOL W_PathIsDirectory(
        [in] LPWSTR pszPath);

    [entry("PathIsDirectoryEmptyA")]
    BOOL A_PathIsDirectoryEmpty(
        [in] LPSTR pszPath);

    [entry("PathIsDirectoryEmptyW")]
    BOOL W_PathIsDirectoryEmpty(
        [in] LPWSTR pszPath);

    [entry("PathIsFileSpecA")]
    BOOL A_PathIsFileSpec(
        [in] LPSTR pszPath);

    [entry("PathIsFileSpecW")]
    BOOL W_PathIsFileSpec(
        [in] LPWSTR pszPath);

    [entry("PathIsPrefixA")]
    BOOL A_PathIsPrefix(
        [in] LPSTR pszPrefix,
        [in] LPSTR pszPath);

    [entry("PathIsPrefixW")]
    BOOL W_PathIsPrefix(
        [in] LPWSTR pszPrefix,
        [in] LPWSTR pszPath);

    [entry("PathIsRelativeA")]
    BOOL A_PathIsRelative(
        [in] LPSTR pszPath);

    [entry("PathIsRelativeW")]
    BOOL W_PathIsRelative(
        [in] LPWSTR pszPath);

    [entry("PathIsRootA")]
    BOOL A_PathIsRoot(
        [in] LPSTR pszPath);

    [entry("PathIsRootW")]
    BOOL W_PathIsRoot(
        [in] LPWSTR pszPath);

    [entry("PathIsSameRootA")]
    BOOL A_PathIsSameRoot(
        [in] LPSTR pszPath1,
        [in] LPSTR pszPath2);

    [entry("PathIsSameRootW")]
    BOOL W_PathIsSameRoot(
        [in] LPWSTR pszPath1,
        [in] LPWSTR pszPath2);

    [entry("PathIsUNCA")]
    BOOL A_PathIsUNC(
        [in] LPSTR pszPath);

    [entry("PathIsUNCW")]
    BOOL W_PathIsUNC(
        [in] LPWSTR pszPath);

    [entry("PathIsNetworkPathA")]
    BOOL A_PathIsNetworkPath(
        [in] LPSTR pszPath);

    [entry("PathIsNetworkPathW")]
    BOOL W_PathIsNetworkPath(
        [in] LPWSTR pszPath);

    [entry("PathIsUNCServerA")]
    BOOL A_PathIsUNCServer(
        [in] LPSTR pszPath);

    [entry("PathIsUNCServerW")]
    BOOL W_PathIsUNCServer(
        [in] LPWSTR pszPath);

    [entry("PathIsUNCServerShareA")]
    BOOL A_PathIsUNCServerShare(
        [in] LPSTR pszPath);

    [entry("PathIsUNCServerShareW")]
    BOOL W_PathIsUNCServerShare(
        [in] LPWSTR pszPath);

    [entry("PathIsContentTypeA")]
    BOOL A_PathIsContentType(
        [in] LPSTR pszPath,
        [in] LPSTR pszContentType);

    [entry("PathIsContentTypeW")]
    BOOL W_PathIsContentType(
        [in] LPWSTR pszPath,
        [in] LPWSTR pszContentType);

    [entry("PathIsURLA")]
    BOOL A_PathIsURL(
        [in] LPSTR pszPath);

    [entry("PathIsURLW")]
    BOOL W_PathIsURL(
        [in] LPWSTR pszPath);

    [entry("PathMakePrettyA")]
    BOOL A_PathMakePretty(
        [in, out] LPSTR pszPath);

    [entry("PathMakePrettyW")]
    BOOL W_PathMakePretty(
        [in, out] LPWSTR pszPath);

    [entry("PathMatchSpecA")]
    BOOL A_PathMatchSpec(
        [in] LPSTR pszFile,
        [in] LPSTR pszSpec);

    [entry("PathMatchSpecW")]
    BOOL W_PathMatchSpec(
        [in] LPWSTR pszFile,
        [in] LPWSTR pszSpec);

    [entry("PathParseIconLocationA")]
    LONG A_PathParseIconLocation(
        [in, out] LPSTR pszIconFile);

    [entry("PathParseIconLocationW")]
    LONG W_PathParseIconLocation(
        [in, out] LPWSTR pszIconFile);

    [entry("PathQuoteSpacesA")]
    void A_PathQuoteSpaces(
        [in, out] LPSTR lpsz);

    [entry("PathQuoteSpacesW")]
    void W_PathQuoteSpaces(
        [in, out] LPWSTR lpsz);

    [entry("PathRelativePathToA")]
    BOOL A_PathRelativePathTo(
        [out] LPSTR pszPath,
        [in] LPSTR pszFrom,
        [in] LONG dwAttrFrom,
        [in] LPSTR pszTo,
        [in] LONG dwAttrTo);

    [entry("PathRelativePathToW")]
    BOOL W_PathRelativePathTo(
        [out] LPWSTR pszPath,
        [in] LPWSTR pszFrom,
        [in] LONG dwAttrFrom,
        [in] LPWSTR pszTo,
        [in] LONG dwAttrTo);

    [entry("PathRemoveArgsA")]
    void A_PathRemoveArgs(
        [in, out] LPSTR pszPath);

    [entry("PathRemoveArgsW")]
    void W_PathRemoveArgs(
        [in, out] LPWSTR pszPath);

    [entry("PathRemoveBackslashA")]
    LONG A_PathRemoveBackslash(
        [in, out] LPSTR pszPath);

    [entry("PathRemoveBackslashW")]
    LONG W_PathRemoveBackslash(
        [in, out] LPWSTR pszPath);

    [entry("PathRemoveBlanksA")]
    void A_PathRemoveBlanks(
        [in, out] LPSTR pszPath);

    [entry("PathRemoveBlanksW")]
    void W_PathRemoveBlanks(
        [in, out] LPWSTR pszPath);

    [entry("PathRemoveExtensionA")]
    void A_PathRemoveExtension(
        [in, out] LPSTR pszPath);

    [entry("PathRemoveExtensionW")]
    void W_PathRemoveExtension(
        [in, out] LPWSTR pszPath);

    [entry("PathRemoveFileSpecA")]
    BOOL A_PathRemoveFileSpec(
        [in, out] LPSTR pszPath);

    [entry("PathRemoveFileSpecW")]
    BOOL W_PathRemoveFileSpec(
        [in, out] LPWSTR pszPath);

    [entry("PathRenameExtensionA")]
    BOOL A_PathRenameExtension(
        [in, out] LPSTR pszPath,
        [in] LPSTR pszExt);

    [entry("PathRenameExtensionW")]
    BOOL W_PathRenameExtension(
        [in, out] LPWSTR pszPath,
        [in] LPWSTR pszExt);

    [entry("PathSearchAndQualifyA")]
    BOOL A_PathSearchAndQualify(
        [in] LPSTR pszPath,
        [out] LPSTR pszBuf,
        [in] LONG cchBuf);

    [entry("PathSearchAndQualifyW")]
    BOOL W_PathSearchAndQualify(
        [in] LPWSTR pszPath,
        [out] LPWSTR pszBuf,
        [in] LONG cchBuf);

    [entry("PathSetDlgItemPathA")]
    void A_PathSetDlgItemPath(
        [in] LONG hDlg,
        [in] long id,
        [in] LPSTR pszPath);

    [entry("PathSetDlgItemPathW")]
    void W_PathSetDlgItemPath(
        [in] LONG hDlg,
        [in] long id,
        [in] LPWSTR pszPath);

    [entry("PathSkipRootA")]
    LONG A_PathSkipRoot(
        [in, out] LPSTR pszPath);

    [entry("PathSkipRootW")]
    LONG W_PathSkipRoot(
        [in, out] LPWSTR pszPath);

    [entry("PathStripPathA")]
    void A_PathStripPath(
        [in, out] LPSTR pszPath);

    [entry("PathStripPathW")]
    void W_PathStripPath(
        [in, out] LPWSTR pszPath);

    [entry("PathStripToRootA")]
    BOOL A_PathStripToRoot(
        [in, out] LPSTR pszPath);

    [entry("PathStripToRootW")]
    BOOL W_PathStripToRoot(
        [in, out] LPWSTR pszPath);

    [entry("PathUnquoteSpacesA")]
    void A_PathUnquoteSpaces(
        [in, out] LPSTR lpsz);

    [entry("PathUnquoteSpacesW")]
    void W_PathUnquoteSpaces(
        [in, out] LPWSTR lpsz);

    [entry("PathMakeSystemFolderA")]
    BOOL A_PathMakeSystemFolder(
        [in] LPSTR pszPath);

    [entry("PathMakeSystemFolderW")]
    BOOL W_PathMakeSystemFolder(
        [in] LPWSTR pszPath);

    [entry("PathUnmakeSystemFolderA")]
    BOOL A_PathUnmakeSystemFolder(
        [in] LPSTR pszPath);

    [entry("PathUnmakeSystemFolderW")]
    BOOL W_PathUnmakeSystemFolder(
        [in] LPWSTR pszPath);

    [entry("PathIsSystemFolderA")]
    BOOL A_PathIsSystemFolder(
        [in] LPSTR pszPath,
        [in] LONG dwAttrb);

    [entry("PathIsSystemFolderW")]
    BOOL W_PathIsSystemFolder(
        [in] LPWSTR pszPath,
        [in] LONG dwAttrb);

    [entry("PathUndecorateA")]
    void A_PathUndecorate(
        [in] LPSTR pszPath);

    [entry("PathUndecorateW")]
    void W_PathUndecorate(
        [in] LPWSTR pszPath);

    [entry("PathUnExpandEnvStringsA")]
    BOOL A_PathUnExpandEnvStrings(
        [in] LPSTR pszPath,
        [out] LPSTR pszBuf,
        [in] LONG cchBuf);

    [entry("PathUnExpandEnvStringsW")]
    BOOL W_PathUnExpandEnvStrings(
        [in] LPWSTR pszPath,
        [out] LPWSTR pszBuf,
        [in] LONG cchBuf);

//
// URL functions
//

    [entry("UrlCompareA")]
    LONG A_UrlCompare(
        [in] LPSTR psz1,
        [in] LPSTR psz2,
        [in] BOOL fIgnoreSlash);

    [entry("UrlCompareW")]
    LONG W_UrlCompare(
        [in] LPWSTR psz1,
        [in] LPWSTR psz2,
        [in] BOOL fIgnoreSlash);

    [entry("UrlCombineA")]
    HRESULT A_UrlCombine(
        [in] LPSTR pszBase,
        [in] LPSTR pszRelative,
        [out] LPSTR pszCombined,
        [in, out] LONG *pcchCombined,
        [in] LONG dwFlags);

    [entry("UrlCombineW")]
    HRESULT W_UrlCombine(
        [in] LPWSTR pszBase,
        [in] LPWSTR pszRelative,
        [out] LPWSTR pszCombined,
        [in, out] LONG *pcchCombined,
        [in] LONG dwFlags);

    [entry("UrlCanonicalizeA")]
    HRESULT A_UrlCanonicalize(
        [in] LPSTR pszUrl,
        [out] LPSTR pszCanonicalized,
        [in, out] LONG *pcchCanonicalized,
        [in] LONG dwFlags);

    [entry("UrlCanonicalizeW")]
    HRESULT W_UrlCanonicalize(
        [in] LPWSTR pszUrl,
        [out] LPWSTR pszCanonicalized,
        [in, out] LONG *pcchCanonicalized,
        [in] LONG dwFlags);

    [entry("UrlIsOpaqueA")]
    BOOL A_UrlIsOpaque(
        [in] LPSTR pszURL);

    [entry("UrlIsOpaqueW")]
    BOOL W_UrlIsOpaque(
        [in] LPWSTR pszURL);

    [entry("UrlIsNoHistoryA")]
    BOOL A_UrlIsNoHistory(
        [in] LPSTR pszURL);

    [entry("UrlIsNoHistoryW")]
    BOOL W_UrlIsNoHistory(
        [in] LPWSTR pszURL);

    [entry("UrlIsA")]
    BOOL A_UrlIs(
        [in] LPSTR pszUrl,
        [in] URLIS UrlIs);

    [entry("UrlIsW")]
    BOOL W_UrlIs(
        [in] LPWSTR pszUrl,
        [in] URLIS UrlIs);

    [entry("UrlGetLocationA")]
    LONG A_UrlGetLocation(
        [in] LPSTR psz1);

    [entry("UrlGetLocationW")]
    LONG W_UrlGetLocation(
        [in] LPWSTR psz1);

    [entry("UrlUnescapeA")]
    HRESULT A_UrlUnescape(
        [in, out] LPSTR pszUrl,
        [out] LPSTR pszUnescaped,
        [in, out] LONG *pcchUnescaped,
        [in] LONG dwFlags);

    [entry("UrlUnescapeW")]
    HRESULT W_UrlUnescape(
        [in, out] LPWSTR pszUrl,
        [out] LPWSTR pszUnescaped,
        [in, out] LONG *pcchUnescaped,
        [in] LONG dwFlags);

    [entry("UrlEscapeA")]
    HRESULT A_UrlEscape(
        [in] LPSTR pszUrl,
        [out] LPSTR pszEscaped,
        [in, out] LONG *pcchEscaped,
        [in] LONG dwFlags);

    [entry("UrlEscapeW")]
    HRESULT W_UrlEscape(
        [in] LPWSTR pszUrl,
        [out] LPWSTR pszEscaped,
        [in, out] LONG *pcchEscaped,
        [in] LONG dwFlags);

    [entry("UrlCreateFromPathA")]
    HRESULT A_UrlCreateFromPath(
        [in] LPSTR pszPath,
        [out] LPSTR pszUrl,
        [in, out] LONG *pcchUrl,
        [in] LONG dwFlags);

    [entry("UrlCreateFromPathW")]
    HRESULT W_UrlCreateFromPath(
        [in] LPWSTR pszPath,
        [out] LPWSTR pszUrl,
        [in, out] LONG *pcchUrl,
        [in] LONG dwFlags);

    [entry("PathCreateFromUrlA")]
    HRESULT A_PathCreateFromUrl(
        [in] LPSTR pszUrl,
        [out] LPSTR pszPath,
        [in, out] LONG *pcchPath,
        [in] LONG dwFlags);

    [entry("PathCreateFromUrlW")]
    HRESULT W_PathCreateFromUrl(
        [in] LPWSTR pszUrl,
        [out] LPWSTR pszPath,
        [in, out] LONG *pcchPath,
        [in] LONG dwFlags);

    [entry("UrlHashA")]
    HRESULT A_UrlHash(
        [in] LPSTR pszUrl,
        [out] LPSTR pbHash,
        [in] LONG cbHash);

    [entry("UrlHashW")]
    HRESULT W_UrlHash(
        [in] LPWSTR pszUrl,
        [out] LPSTR pbHash,
        [in] LONG cbHash);

    [entry("UrlGetPartW")]
    HRESULT W_UrlGetPart(
        [in] LPWSTR pszIn,
        [out] LPWSTR pszOut,
        [in, out] LONG *pcchOut,
        [in] LONG dwPart,
        [in] LONG dwFlags);

    [entry("UrlGetPartA")]
    HRESULT A_UrlGetPart(
        [in] LPSTR pszIn,
        [out] LPSTR pszOut,
        [in, out] LONG *pcchOut,
        [in] LONG dwPart,
        [in] LONG dwFlags);

    [entry("UrlApplySchemeA")]
    HRESULT A_UrlApplyScheme(
        [in] LPSTR pszIn,
        [out] LPSTR pszOut,
        [in, out] LONG *pcchOut,
        [in] LONG dwFlags);

    [entry("UrlApplySchemeW")]
    HRESULT W_UrlApplyScheme(
        [in] LPWSTR pszIn,
        [out] LPWSTR pszOut,
        [in, out] LONG *pcchOut,
        [in] LONG dwFlags);

    [entry("HashData")]
    HRESULT A_HashData(
        [in] BYTE *pbData,
        [in] LONG cbData,
        [out] BYTE *pbHash,
        [in] LONG cbHash);

// Autocomplete

    [
	helpstring("hwndEdit: HWND of editbox, ComboBox or ComboBoxEx. dwFlags: see SHACF_ flags"),
	entry("SHAutoComplete")]
    HRESULT SHAutoComplete(
        [in] LONG hwndEdit,
        [in] SHACF dwFlags);

    [entry("SHSetThreadRef")]
    HRESULT SHSetThreadRef(
        [in] stdole.IUnknown *punk);

    [entry("SHGetThreadRef")]
    HRESULT SHGetThreadRef(
        [out] stdole.IUnknown *ppunk);

    [entry("SHSkipJunction")]
    BOOL SHSkipJunction(
        [in] IBindCtx *pbc,
        [in] UUID *pclsid);

    [entry("SHCreateThread")]
    BOOL SHCreateThread(
        [in] LONG pfnThreadProc,
        [in] void *pData,
        [in] CTF dwFlags,
        [in] LONG pfnCallback);

    [entry("SHCreateStreamOnFileA")]
    HRESULT A_SHCreateStreamOnFile(
        [in] LPSTR pszFile,
        [in] STGM grfMode,
        [out, retval] IStream **ppstm);

    [entry("SHCreateStreamOnFileW")]
    HRESULT W_SHCreateStreamOnFile(
        [in] LPWSTR pszFile,
        [in] STGM grfMode,
        [out, retval] IStream **ppstm);

	[helpstring("WinXP only"),
	entry("SHCreateStreamOnFileEx")]	
	HRESULT SHCreateStreamOnFileEx([in] LPWSTR pszFile,
		[in] DWORD grfMode,			//STGM
		[in] DWORD dwAttributes,	//File attributes
		[in] BOOL fCreate,
		[in] LONG pstmTemplate,	//Reserved. IStream
		[out, retval] IStream **ppstm);

	[helpstring("Win2K only. cchBuf, num of chars + 1"),
	entry("SHUnicodeToAnsi")]
	LONG SHUnicodeToAnsi([in] LPWSTR pwszSrc,
		[out] BYTE *pszDst,
		[in] long cchBuf);

	[helpstring("Win2K only. cchBuf, num of chars + 1"),
	entry("SHAnsiToUnicode")]
	LONG SHAnsiToUnicode([in] BYTE *pszSrc,
		[out] LPWSTR pwszDst,
		[in] long cwchBuf);

	[helpstring("Win2K only.pdwFlags; pdwFlags; 0 default"),
	entry("SHFormatDateTimeA")]
	LONG A_SHFormatDateTime([in] FILETIME puft,
		[in] DWORD pdwFlags,
		[out] LPSTR pszBuf,
		[in] UINT ucchBuf);

	[helpstring("Win2K only.pdwFlags; pdwFlags; 0 default"),
	entry("SHFormatDateTimeW")]
	LONG W_SHFormatDateTime([in] FILETIME puft,
		[in] DWORD pdwFlags,
		[out] LPWSTR pszBuf,
		[in] UINT ucchBuf);

//FDTF_DEFAULT = Equivalent to FDTF_SHORTDATE | FDTF_SHORTTIME.
//        2/27/2004 5:10 PM
//FDTF_LONGDATE | FDTF_SHORTTIME
//        Friday, February 27, 2004, 5:10 PM
//FDTF_LONGDATE | FDTF_LONGTIME | FDTF_RELATIVE
//        Today, February 27, 2004, 5:10:36 PM


//
//====== GDI helper functions ================================================
//

    [entry("SHCreateShellPalette")]
    LONG SHCreateShellPalette(
        [in] LONG hdc);

    [entry("ColorRGBToHLS")]
    void ColorRGBToHLS(
        [in] LONG clrRGB,
        [out] short *pwHue,
        [out] short *pwLuminance,
        [out] short *pwSaturation);

    [entry("ColorHLSToRGB")]
    LONG ColorHLSToRGB(
        [in] short wHue,
        [in] short wLuminance,
        [in] short wSaturation);

    [entry("ColorAdjustLuma")]
    LONG ColorAdjustLuma(
        [in] LONG clrRGB,
        [in] long n,
        [in] BOOL fScale);
	
	[entry("SHGetInverseCMAP")]
	HRESULT SHGetInverseCMAP(
		[out] BYTE *pbMap,
		[in] long cbMap);


}; //End SHLWAPI