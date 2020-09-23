// ******* Oleaut32
[
	helpstring("OLE Automation Functions"),
	dllname("OLEAUT32.DLL")
]
module OleAuto {

/*
	const WORD STDOLE_MAJORVERNUM = 0x1;
	const WORD STDOLE_MINORVERNUM = 0x0;
	const WORD STDOLE_LCID        = 0x0000;
*/
    //----------
    // BSTR API
    //----------

    [entry("SysAllocString")]
    BSTR SysAllocString([in] LONG *pStr);

    [
	entry("SysAllocString")	]
    LONG SysAllocStringPtr([in] LONG *pStr);

    [
	entry("SysAllocString")	]
    BSTR RTrimNull([in] LPWSTR pStr);

    [
	entry("SysReAllocString")	]
    LONG SysReAllocString(
        [in] BSTR pbstr,
        [in] LONG * sz);

    [
	entry("SysAllocStringLen")	]
    BSTR SysAllocStringLen(
        [in] LONG * pStr,
        [in] LONG cb);

    [
	entry("SysAllocStringLen")]
    LONG SysAllocStringLenPtr(
        [in] LONG * pStr,
        [in] LONG cb);

    [
	entry("SysReAllocStringLen")
	]
    LONG SysReAllocStringLen(
        [in] BSTR bstr,
        [in] LONG * pStr,
        [in] lONG cb);

    [
	entry("SysFreeString")]
    void SysFreeString(
        [in] BSTR bstr);
 
    [
	entry("SysStringLen")]
    LONG SysStringLen(
        [in] BSTR str);

    [
	entry("SysStringByteLen")]
    LONG SysStringByteLen(
        [in] BSTR bstr);

    [
	entry("SysAllocStringByteLen")	]
    BSTR SysAllocStringByteLen(
        [in] long *psz,
        [in] LONG len);


    [entry("SystemTimeToVariantTime")]
    LONG SystemTimeToVariantTime(
        [in] SYSTEMTIME *lpSystemTime,
        [out] DOUBLE *pvtime);

    [entry("VariantTimeToSystemTime")]
    LONG VariantTimeToSystemTime(
        [in] DOUBLE vtime,
        [in] SYSTEMTIME *lpSystemTime);

    [entry("DosDateTimeToVariantTime")]
    LONG DosDateTimeToVariantTime(
        [in] short wDosDate,
        [in] short wDosTime,
        [out] DOUBLE *pvtime);

    [entry("VariantTimeToDosDateTime")]
    LONG VariantTimeToDosDateTime(
        [in] DOUBLE vtime,
        [out] short *pwDosDate,
        [out] short *pwDosTime);


	//define MEMBERID_NIL DISPID_UNKNOWN
	const DWORD ID_DEFAULTINST  = -2;


	const DWORD IDLFLAG_NONE    = 0;
	const DWORD IDLFLAG_FIN     = 0x1;
	const DWORD IDLFLAG_FOUT    = 0x2;
	const DWORD IDLFLAG_FLCID   = 0x4;
	const DWORD IDLFLAG_FRETVAL = 0x8;


	// Flags for IDispatch::Invoke 
	const DWORD DISPATCH_METHOD         = 0x1;
	const DWORD DISPATCH_PROPERTYGET    = 0x2;
	const DWORD DISPATCH_PROPERTYPUT    = 0x4;
	const DWORD DISPATCH_PROPERTYPUTREF = 0x8;

    [entry("LHashValOfName")]
    LONG LHashValOfName(
        [in] LONG lcid,
        [in] LPWSTR szName);

    [entry("LHashValOfNameSys")]
    LONG LHashValOfNameSys(
        [in] long syskind, //SYSKIND
        [in] LONG  lcid,
        [in] LPWSTR szName);

    [entry("RegisterTypeLib")]
    long RegisterTypeLib(
        [in] void *ptlib, //ITypeLib
        [in] LPWSTR szFullPath,
        [in] LPWSTR szHelpDir);

    [entry("UnRegisterTypeLib")]
    long UnRegisterTypeLib(
        [in] UUID *libID,
        [in] short wVerMajor,
        [in] short wVerMinor,
        [in] LONG lcid,
        [in] long syskind); //

    [entry("CreateTypeLib")]
    long CreateTypeLib(
        [in] long syskind,
        [in] LPWSTR *szFile,
        [out] LPVOID ppctlib); //ICreateTypeLib *

    [entry("CreateTypeLib2")]
    long CreateTypeLib2(
        [in] long syskind,
        [in] LPWSTR szFile,
        [out] LPVOID ppctlib);

    [entry("LoadRegTypeLib")]
    long LoadRegTypeLib(
        [in] UUID *rguid,
        [in] short wVerMajor,
        [in] short wVerMinor,
        [in] LONG lcid,
        [out] LPVOID pptlib);

    [entry("LoadTypeLibEx")]
    long LoadTypeLibEx(
        [in] LPWSTR szFile,
        [in] REGKIND regkind,
        [out] LPVOID pptlib);

    [entry("LoadTypeLib")]
    long LoadTypeLib(
        [in] LPWSTR szFile,
        [out] LPVOID pptlib);

    [entry("QueryPathOfRegTypeLib")]
    long QueryPathOfRegTypeLib(
        [in] UUID *GUID,
        [in] short wVerMajor,
        [in] short wVerMinor,
        [in] LONG LCID,
        [in] BSTR lpbstrPathName);

    [entry("DispGetParam")]
    long DispGetParam(
        [in] LPVOID pdispparams,
        [in] LONG position,
        [in] short vtTarg,
        [out] LPVOID pvarResult,
        [out] LONG *puArgErr);

    [entry("DispGetIDsOfNames")]
    long DispGetIDsOfNames(
        [in] LPVOID ptinfo,
        [in] LPWSTR *rgszNames,
        [in] LONG cNames,
        [out] LONG *rgdispid);

    [entry("DispInvoke")]
    long DispInvoke(
        [in] LPVOID this,
        [in] LPVOID ptinfo,
        [in] LONG dispidMember,
        [in] short wFlags,
        [in, out] LPVOID pparams,
        [out] LPVOID pvarResult,
        [in, out] LPVOID pexcepinfo,
        [out] LONG *puArgErr);

    [entry("CreateStdDispatch")]
    long CreateStdDispatch(
        [in] stdole.IUnknown *punkOuter,
        [in] void *pvThis,
        [in] void *ptinfo, //ITypeInfo
        [out] stdole.IUnknown *ppunkStdDisp);


    [entry("OleLoadPicture")]
    long OleLoadPicture(
        [in] void *lpstream,
        [in] LONG lSize,
        [in] BOOL fRunmode,
        [in] UUID *riid,
        [out] void *lplpvObj); //IPicture

    [entry("OleLoadPictureEx")]
    long OleLoadPictureEx(
        [in] void *lpstream, //IStream
        [in] LONG lSize,
        [in] BOOL fRunmode,
        [in] UUID *riid,
        [in] LONG xSizeDesired,
        [in] LONG ySizeDesired,
        [in] stdole.LoadPictureConstants dwFlags,
        [out] void *lplpvObj); //IPicture

    [entry("OleLoadPicturePath")]
    long OleLoadPicturePath(
        [in] LPWSTR szURLorPath,
	    [in] void *punkCaller, //IUnknown
        [in] LONG dwReserved,
	    [in] OLE_COLOR clrReserved,
        [in] UUID *riid,
        [out] void *ppvRet); //IPicture

    [entry("OleLoadPictureFile")]
    long OleLoadPictureFile(
        [in] VARIANT varFileName,
        [out] void *lplpdispPicture); //IPictureDisp

    [entry("OleLoadPictureFileEx")]
    long OleLoadPictureFileEx(
        [in] VARIANT varFileName,
        [in] LONG xSizeDesired,
        [in] LONG ySizeDesired,
        [in] stdole.LoadPictureConstants dwFlags,
        [out] void *lplpdispPicture); //IPictureDisp

    [entry("OleSavePictureFile")]
    long OleSavePictureFile(
        [in] void *lpdispPicture, //IPictureDisp
        [in] BSTR bstrFileName);

	[entry("SafeArrayGetDim")]
	UINT SafeArrayGetStringDim([in] SAFEARRAY(BSTR)* psa);

	[entry("SafeArrayGetDim")]
	UINT SafeArrayGetLongDim([in] SAFEARRAY(long)* psa);

	[entry("SafeArrayGetDim")]
	UINT SafeArrayGetByteDim([in] SAFEARRAY(BYTE)* psa);

	[entry("SafeArrayGetDim")]
	UINT SafeArrayGetShortDim([in] SAFEARRAY(short)* psa);

	[entry("SafeArrayGetDim")]
	UINT SafeArrayGetDoubleDim([in] SAFEARRAY(DOUBLE)* psa);

	[entry("SafeArrayGetDim")]
	UINT SafeArrayGetVariantDim([in] SAFEARRAY(VARIANT)* psa);

}; //End Oleaut32