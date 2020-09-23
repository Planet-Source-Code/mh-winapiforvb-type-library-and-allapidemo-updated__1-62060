[
	helpstring("WININET Constants"),
	dllname("NOSUCHW.DLL")
]
module WININETConstant {

    const long FLAG_ICC_FORCE_CONNECTION       = 0x00000001;

    const long INTERNET_INVALID_PORT_NUMBER    = 0;           // use the protocol-specific default

    const long INTERNET_DEFAULT_FTP_PORT       = 21;          // default for FTP servers
    const long INTERNET_DEFAULT_GOPHER_PORT    = 70;          //    "     "  gopher "
    const long INTERNET_DEFAULT_HTTP_PORT      = 80;          //    "     "  HTTP   "
    const long INTERNET_DEFAULT_HTTPS_PORT     = 443;         //    "     "  HTTPS  "
    const long INTERNET_DEFAULT_SOCKS_PORT     = 1080;        // default for SOCKS firewall servers.

    const long INTERNET_MAX_HOST_NAME_LENGTH   = 256;
    const long INTERNET_MAX_USER_NAME_LENGTH   = 128;
    const long INTERNET_MAX_PASSWORD_LENGTH    = 128;
    const long INTERNET_MAX_PORT_NUMBER_LENGTH = 5;           // INTERNET_PORT is unsigned short
    const long INTERNET_MAX_PORT_NUMBER_VALUE  = 65535;       // maximum unsigned short value
    const long INTERNET_MAX_PATH_LENGTH        = 2048;
    const long INTERNET_MAX_SCHEME_LENGTH      = 32;          // longest protocol name length
//    const long INTERNET_MAX_URL_LENGTH         (INTERNET_MAX_SCHEME_LENGTH \
//                                        + sizeof("://") \
//                                        + INTERNET_MAX_PATH_LENGTH)

//    const long INTERNET_KEEP_ALIVE_UNKNOWN     ((DWORD)-1)
    const long INTERNET_KEEP_ALIVE_ENABLED     = 1;
    const long INTERNET_KEEP_ALIVE_DISABLED    = 0;

    const long INTERNET_REQFLAG_FROM_CACHE     = 0x00000001;  // response came from cache
    const long INTERNET_REQFLAG_ASYNC          = 0x00000002;  // request was made asynchronously
    const long INTERNET_REQFLAG_VIA_PROXY      = 0x00000004;  // request was made via a proxy
    const long INTERNET_REQFLAG_NO_HEADERS     = 0x00000008;  // orginal response contained no headers
    const long INTERNET_REQFLAG_PASSIVE        = 0x00000010;  // FTP: passive-mode connection
    const long INTERNET_REQFLAG_CACHE_WRITE_DISABLED = 0x00000040;  // HTTPS: this request not cacheable

    const long INTERNET_FLAG_RELOAD            = 0x80000000;  // retrieve the original item

    const long INTERNET_FLAG_RAW_DATA          = 0x40000000;  // FTP/gopher find: receive the item as raw (structured) data
    const long INTERNET_FLAG_EXISTING_CONNECT  = 0x20000000;  // FTP: use existing InternetConnect handle for server if possible

    const long INTERNET_FLAG_ASYNC             = 0x10000000;  // this request is asynchronous (where supported)

    const long INTERNET_FLAG_PASSIVE           = 0x08000000;  // used for FTP connections

    const long INTERNET_FLAG_NO_CACHE_WRITE    = 0x04000000;  // don't write this item to the cache
    const long INTERNET_FLAG_DONT_CACHE        = 0x04000000;
    const long INTERNET_FLAG_MAKE_PERSISTENT   = 0x02000000;  // make this item persistent in cache
    const long INTERNET_FLAG_FROM_CACHE        = 0x01000000;  // use offline semantics
    const long INTERNET_FLAG_OFFLINE           = 0x01000000;

    const long INTERNET_FLAG_SECURE            = 0x00800000;  // use PCT/SSL if applicable (HTTP)
    const long INTERNET_FLAG_KEEP_CONNECTION   = 0x00400000;  // use keep-alive semantics
    const long INTERNET_FLAG_NO_AUTO_REDIRECT  = 0x00200000;  // don't handle redirections automatically
    const long INTERNET_FLAG_READ_PREFETCH     = 0x00100000;  // do background read prefetch
    const long INTERNET_FLAG_NO_COOKIES        = 0x00080000;  // no automatic cookie handling
    const long INTERNET_FLAG_NO_AUTH           = 0x00040000;  // no automatic authentication handling
    const long INTERNET_FLAG_CACHE_IF_NET_FAIL = 0x00010000;  // return cache file if net request fails

    const long INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP  = 0x00008000; // ex: https:// to http://
    const long INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = 0x00004000; // ex: http:// to https://
    const long INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = 0x00002000; // expired X509 Cert.
    const long INTERNET_FLAG_IGNORE_CERT_CN_INVALID   = 0x00001000; // bad common name in X509 Cert.

    const long INTERNET_FLAG_RESYNCHRONIZE    = 0x00000800;  // asking wininet to update an item if it is newer
    const long INTERNET_FLAG_HYPERLINK        = 0x00000400;  // asking wininet to do hyperlinking semantic which works right for scripts
    const long INTERNET_FLAG_NO_UI            = 0x00000200;  // no cookie popup
    const long INTERNET_FLAG_PRAGMA_NOCACHE   = 0x00000100;  // asking wininet to add "pragma: no-cache"
    const long INTERNET_FLAG_CACHE_ASYNC      = 0x00000080;  // ok to perform lazy cache-write
    const long INTERNET_FLAG_FORMS_SUBMIT     = 0x00000040;  // this is a forms submit
    const long INTERNET_FLAG_NEED_FILE        = 0x00000010;  // need a file for this request
    const long INTERNET_FLAG_MUST_CACHE_REQUEST = 0x00000010;

    const long INTERNET_FLAG_TRANSFER_ASCII   = 0x00000001;
    const long INTERNET_FLAG_TRANSFER_BINARY  = 0x00000002;

    const long INTERNET_ERROR_MASK_INSERT_CDROM                   = 0x1;
    const long WININET_API_FLAG_ASYNC         = 0x00000001;  // force async operation
    const long WININET_API_FLAG_SYNC          = 0x00000004;  // force sync operation
    const long WININET_API_FLAG_USE_CONTEXT   = 0x00000008;  // use value supplied in dwContext (even if 0)

    const long INTERNET_NO_CALLBACK            = 0;
    const long ISO_FORCE_DISCONNECTED  = 0x00000001;

	[helpstring("Constant for InternetTimeFromSystemTime")]
    const long INTERNET_RFC1123_FORMAT    =  0;
	[helpstring("Constant for InternetTimeFromSystemTime")]
    const long INTERNET_RFC1123_BUFSIZE   = 30;

	[helpstring("Flag for InternetCrackUrl() and InternetCreateUrl()")]
	const long ICU_ESCAPE     = 0x80000000;  // (un)escape URL characters
	[helpstring("Flag for InternetCrackUrl() and InternetCreateUrl()")]
	const long ICU_USERNAME   = 0x40000000;  // use internal username & password

// flags for InternetCanonicalizeUrl() and InternetCombineUrl()
//	const long ICU_NO_ENCODE   = 0x20000000;  // Don't convert unsafe characters to escape sequence
//	const long ICU_DECODE      = 0x10000000;  // Convert %XX escape sequences to characters
//	const long ICU_NO_META     = 0x08000000;  // Don't convert .. etc. meta path sequences
//	const long ICU_ENCODE_SPACES_ONLY = 0x04000000;  // Encode spaces only
//	const long ICU_BROWSER_MODE = 0x02000000; // Special encode/decode rules for browser

// access types for InternetOpen()

    const long INTERNET_OPEN_TYPE_PRECONFIG                    = 0;   // use registry configuration
    const long INTERNET_OPEN_TYPE_DIRECT                       = 1;   // direct to net
    const long INTERNET_OPEN_TYPE_PROXY                        = 3;   // via named proxy
    const long INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY  = 4;   // prevent using java/script/INS

// old names for access types

    const long PRE_CONFIG_INTERNET_ACCESS  = 0;
    const long LOCAL_INTERNET_ACCESS       = 1;
    const long CERN_PROXY_INTERNET_ACCESS  = 3;

// flags for InternetReadFileEx()

    const long IRF_ASYNC       = 0x00000001;
    const long IRF_SYNC        = 0x00000004;
    const long IRF_USE_CONTEXT = 0x00000008;
    const long IRF_NO_WAIT     = 0x00000008;

// flags for InternetSetOptionEx()

    const long ISO_GLOBAL      = 0x00000001;  // modify option globally
    const long ISO_REGISTRY    = 0x00000002;  // write option to registry (where applicable)

//    const long ISO_VALID_FLAGS (ISO_GLOBAL | ISO_REGISTRY)

// options manifests for Internet{Query|Set}Option

    const long INTERNET_OPTION_CALLBACK                = 1;
    const long INTERNET_OPTION_CONNECT_TIMEOUT         = 2;
    const long INTERNET_OPTION_CONNECT_RETRIES         = 3;
    const long INTERNET_OPTION_CONNECT_BACKOFF         = 4;
    const long INTERNET_OPTION_SEND_TIMEOUT            = 5;
    const long INTERNET_OPTION_CONTROL_SEND_TIMEOUT    = 5;
    const long INTERNET_OPTION_RECEIVE_TIMEOUT         = 6;
    const long INTERNET_OPTION_CONTROL_RECEIVE_TIMEOUT = 6;
    const long INTERNET_OPTION_DATA_SEND_TIMEOUT       = 7;
    const long INTERNET_OPTION_DATA_RECEIVE_TIMEOUT    = 8;
    const long INTERNET_OPTION_HANDLE_TYPE             = 9;
    const long INTERNET_OPTION_LISTEN_TIMEOUT          = 11;
    const long INTERNET_OPTION_READ_BUFFER_SIZE        = 12;
    const long INTERNET_OPTION_WRITE_BUFFER_SIZE       = 13;

    const long INTERNET_OPTION_ASYNC_ID                = 15;
    const long INTERNET_OPTION_ASYNC_PRIORITY          = 16;

    const long INTERNET_OPTION_PARENT_HANDLE           = 21;
    const long INTERNET_OPTION_KEEP_CONNECTION         = 22;
    const long INTERNET_OPTION_REQUEST_FLAGS           = 23;
    const long INTERNET_OPTION_EXTENDED_ERROR          = 24;

    const long INTERNET_OPTION_OFFLINE_MODE            = 26;
    const long INTERNET_OPTION_CACHE_STREAM_HANDLE     = 27;
    const long INTERNET_OPTION_USERNAME                = 28;
    const long INTERNET_OPTION_PASSWORD                = 29;
    const long INTERNET_OPTION_ASYNC                   = 30;
    const long INTERNET_OPTION_SECURITY_FLAGS          = 31;
    const long INTERNET_OPTION_SECURITY_CERTIFICATE_STRUCT = 32;
    const long INTERNET_OPTION_DATAFILE_NAME           = 33;
    const long INTERNET_OPTION_URL                     = 34;
    const long INTERNET_OPTION_SECURITY_CERTIFICATE    = 35;
    const long INTERNET_OPTION_SECURITY_KEY_BITNESS    = 36;
    const long INTERNET_OPTION_REFRESH                 = 37;
    const long INTERNET_OPTION_PROXY                   = 38;
    const long INTERNET_OPTION_SETTINGS_CHANGED        = 39;
    const long INTERNET_OPTION_VERSION                 = 40;
    const long INTERNET_OPTION_USER_AGENT              = 41;
    const long INTERNET_OPTION_END_BROWSER_SESSION     = 42;
    const long INTERNET_OPTION_PROXY_USERNAME          = 43;
    const long INTERNET_OPTION_PROXY_PASSWORD          = 44;
    const long INTERNET_OPTION_CONTEXT_VALUE           = 45;
    const long INTERNET_OPTION_CONNECT_LIMIT           = 46;
    const long INTERNET_OPTION_SECURITY_SELECT_CLIENT_CERT = 47;
    const long INTERNET_OPTION_POLICY                  = 48;
    const long INTERNET_OPTION_DISCONNECTED_TIMEOUT    = 49;
    const long INTERNET_OPTION_CONNECTED_STATE         = 50;
    const long INTERNET_OPTION_IDLE_STATE              = 51;
    const long INTERNET_OPTION_OFFLINE_SEMANTICS       = 52;
    const long INTERNET_OPTION_SECONDARY_CACHE_KEY     = 53;
    const long INTERNET_OPTION_CALLBACK_FILTER         = 54;
    const long INTERNET_OPTION_CONNECT_TIME            = 55;
    const long INTERNET_OPTION_SEND_THROUGHPUT         = 56;
    const long INTERNET_OPTION_RECEIVE_THROUGHPUT      = 57;
    const long INTERNET_OPTION_REQUEST_PRIORITY        = 58;
    const long INTERNET_OPTION_HTTP_VERSION            = 59;
    const long INTERNET_OPTION_RESET_URLCACHE_SESSION  = 60;
    const long INTERNET_OPTION_ERROR_MASK              = 62;

    const long INTERNET_FIRST_OPTION                   = 1;
    const long INTERNET_LAST_OPTION                    = 62;


// values for INTERNET_OPTION_PRIORITY

    const long INTERNET_PRIORITY_FOREGROUND            = 1000;

// handle types

    const long INTERNET_HANDLE_TYPE_INTERNET           = 1;
    const long INTERNET_HANDLE_TYPE_CONNECT_FTP        = 2;
    const long INTERNET_HANDLE_TYPE_CONNECT_GOPHER     = 3;
    const long INTERNET_HANDLE_TYPE_CONNECT_HTTP       = 4;
    const long INTERNET_HANDLE_TYPE_FTP_FIND           = 5;
    const long INTERNET_HANDLE_TYPE_FTP_FIND_HTML      = 6;
    const long INTERNET_HANDLE_TYPE_FTP_FILE           = 7;
    const long INTERNET_HANDLE_TYPE_FTP_FILE_HTML      = 8;
    const long INTERNET_HANDLE_TYPE_GOPHER_FIND        = 9;
    const long INTERNET_HANDLE_TYPE_GOPHER_FIND_HTML   = 10;
    const long INTERNET_HANDLE_TYPE_GOPHER_FILE        = 11;
    const long INTERNET_HANDLE_TYPE_GOPHER_FILE_HTML   = 12;
    const long INTERNET_HANDLE_TYPE_HTTP_REQUEST       = 13;

// values for INTERNET_OPTION_SECURITY_FLAGS
// query only

    const long SECURITY_FLAG_SECURE                    = 0x00000001; // can query only
    const long SECURITY_FLAG_STRENGTH_WEAK             = 0x10000000;
    const long SECURITY_FLAG_STRENGTH_MEDIUM           = 0x40000000;
    const long SECURITY_FLAG_STRENGTH_STRONG           = 0x20000000;
    const long SECURITY_FLAG_UNKNOWNBIT                = 0x80000000;
    const long SECURITY_FLAG_NORMALBITNESS             = 0x10000000;

// The following are unused
    const long SECURITY_FLAG_SSL                       = 0x00000002;
    const long SECURITY_FLAG_SSL3                      = 0x00000004;
    const long SECURITY_FLAG_PCT                       = 0x00000008;
    const long SECURITY_FLAG_PCT4                      = 0x00000010;
    const long SECURITY_FLAG_IETFSSL4                  = 0x00000020;

// The following are for backwards compatability only.
    const long SECURITY_FLAG_40BIT                     = 0x10000000;
    const long SECURITY_FLAG_128BIT                    = 0x20000000;
    const long SECURITY_FLAG_56BIT                     = 0x40000000;


// setable flags
    const long SECURITY_FLAG_IGNORE_REVOCATION         = 0x00000080;
    const long SECURITY_FLAG_IGNORE_UNKNOWN_CA         = 0x00000100;
    const long SECURITY_FLAG_IGNORE_WRONG_USAGE        = 0x00000200;

    const long SECURITY_FLAG_IGNORE_CERT_CN_INVALID    = 0x00001000;
    const long SECURITY_FLAG_IGNORE_CERT_DATE_INVALID  = 0x00002000;


    const long SECURITY_FLAG_IGNORE_REDIRECT_TO_HTTPS  = 0x00004000;
    const long SECURITY_FLAG_IGNORE_REDIRECT_TO_HTTP   = 0x00008000;

// status manifests for Internet status callback

    const long INTERNET_STATUS_RESOLVING_NAME          = 10;
    const long INTERNET_STATUS_NAME_RESOLVED           = 11;
    const long INTERNET_STATUS_CONNECTING_TO_SERVER    = 20;
    const long INTERNET_STATUS_CONNECTED_TO_SERVER     = 21;
    const long INTERNET_STATUS_SENDING_REQUEST         = 30;
    const long INTERNET_STATUS_REQUEST_SENT            = 31;
    const long INTERNET_STATUS_RECEIVING_RESPONSE      = 40;
    const long INTERNET_STATUS_RESPONSE_RECEIVED       = 41;
    const long INTERNET_STATUS_CTL_RESPONSE_RECEIVED   = 42;
    const long INTERNET_STATUS_PREFETCH                = 43;
    const long INTERNET_STATUS_CLOSING_CONNECTION      = 50;
    const long INTERNET_STATUS_CONNECTION_CLOSED       = 51;
    const long INTERNET_STATUS_HANDLE_CREATED          = 60;
    const long INTERNET_STATUS_HANDLE_CLOSING          = 70;
    const long INTERNET_STATUS_REQUEST_COMPLETE        = 100;
    const long INTERNET_STATUS_REDIRECT                = 110;
    const long INTERNET_STATUS_INTERMEDIATE_RESPONSE   = 120;
    const long INTERNET_STATUS_STATE_CHANGE            = 200;

// the following can be indicated in a state change notification:

    const long INTERNET_STATE_CONNECTED                = 0x00000001;  // connected state (mutually exclusive with disconnected)
    const long INTERNET_STATE_DISCONNECTED             = 0x00000002;  // disconnected from network
    const long INTERNET_STATE_DISCONNECTED_BY_USER     = 0x00000010;  // disconnected by user request
    const long INTERNET_STATE_IDLE                     = 0x00000100;  // no network requests being made (by Wininet)
    const long INTERNET_STATE_BUSY                     = 0x00000200;  // network requests being made (by Wininet)

// manifests

    const long FTP_TRANSFER_TYPE_UNKNOWN   = 0x00000000;
    const long FTP_TRANSFER_TYPE_ASCII     = 0x00000001;
    const long FTP_TRANSFER_TYPE_BINARY    = 0x00000002;

//    const long FTP_TRANSFER_TYPE_MASK      (FTP_TRANSFER_TYPE_ASCII | FTP_TRANSFER_TYPE_BINARY)

// the default major/minor HTTP version numbers

    const long HTTP_MAJOR_VERSION      = 1;
    const long HTTP_MINOR_VERSION      = 0;

//
// HttpQueryInfo info levels. Generally, there is one info level
// for each potential RFC822/HTTP/MIME header that an HTTP server
// may send as part of a request response.
//
// The HTTP_QUERY_RAW_HEADERS info level is provided for clients
// that choose to perform their own header parsing.
//

    const long HTTP_QUERY_MIME_VERSION                 = 0;
    const long HTTP_QUERY_CONTENT_TYPE                 = 1;
    const long HTTP_QUERY_CONTENT_TRANSFER_ENCODING    = 2;
    const long HTTP_QUERY_CONTENT_ID                   = 3;
    const long HTTP_QUERY_CONTENT_DESCRIPTION          = 4;
    const long HTTP_QUERY_CONTENT_LENGTH               = 5;
    const long HTTP_QUERY_CONTENT_LANGUAGE             = 6;
    const long HTTP_QUERY_ALLOW                        = 7;
    const long HTTP_QUERY_PUBLIC                       = 8;
    const long HTTP_QUERY_DATE                         = 9;
    const long HTTP_QUERY_EXPIRES                      = 10;
    const long HTTP_QUERY_LAST_MODIFIED                = 11;
    const long HTTP_QUERY_MESSAGE_ID                   = 12;
    const long HTTP_QUERY_URI                          = 13;
    const long HTTP_QUERY_DERIVED_FROM                 = 14;
    const long HTTP_QUERY_COST                         = 15;
    const long HTTP_QUERY_LINK                         = 16;
    const long HTTP_QUERY_PRAGMA                       = 17;
    const long HTTP_QUERY_VERSION                      = 18;  // special: part of status line
    const long HTTP_QUERY_STATUS_CODE                  = 19;  // special: part of status line
    const long HTTP_QUERY_STATUS_TEXT                  = 20;  // special: part of status line
    const long HTTP_QUERY_RAW_HEADERS                  = 21;  // special: all headers as ASCIIZ
    const long HTTP_QUERY_RAW_HEADERS_CRLF             = 22; // special: all headers
    const long HTTP_QUERY_CONNECTION                   = 23;
    const long HTTP_QUERY_ACCEPT                       = 24;
    const long HTTP_QUERY_ACCEPT_CHARSET               = 25;
    const long HTTP_QUERY_ACCEPT_ENCODING              = 26;
    const long HTTP_QUERY_ACCEPT_LANGUAGE              = 27;
    const long HTTP_QUERY_AUTHORIZATION                = 28;
    const long HTTP_QUERY_CONTENT_ENCODING             = 29;
    const long HTTP_QUERY_FORWARDED                    = 30;
    const long HTTP_QUERY_FROM                         = 31;
    const long HTTP_QUERY_IF_MODIFIED_SINCE            = 32;
    const long HTTP_QUERY_LOCATION                     = 33;
    const long HTTP_QUERY_ORIG_URI                     = 34;
    const long HTTP_QUERY_REFERER                      = 35;
    const long HTTP_QUERY_RETRY_AFTER                  = 36;
    const long HTTP_QUERY_SERVER                       = 37;
    const long HTTP_QUERY_TITLE                        = 38;
    const long HTTP_QUERY_USER_AGENT                   = 39;
    const long HTTP_QUERY_WWW_AUTHENTICATE             = 40;
    const long HTTP_QUERY_PROXY_AUTHENTICATE           = 41;
    const long HTTP_QUERY_ACCEPT_RANGES                = 42;
    const long HTTP_QUERY_SET_COOKIE                   = 43;
    const long HTTP_QUERY_COOKIE                       = 44;
    const long HTTP_QUERY_REQUEST_METHOD               = 45;  // special: GET/POST etc.
    const long HTTP_QUERY_REFRESH                      = 46;
    const long HTTP_QUERY_CONTENT_DISPOSITION          = 47;

//
// HTTP 1.1 defined headers
//

    const long HTTP_QUERY_AGE                          = 48;
    const long HTTP_QUERY_CACHE_CONTROL                = 49;
    const long HTTP_QUERY_CONTENT_BASE                 = 50;
    const long HTTP_QUERY_CONTENT_LOCATION             = 51;
    const long HTTP_QUERY_CONTENT_MD5                  = 52;
    const long HTTP_QUERY_CONTENT_RANGE                = 53;
    const long HTTP_QUERY_ETAG                         = 54;
    const long HTTP_QUERY_HOST                         = 55;
    const long HTTP_QUERY_IF_MATCH                     = 56;
    const long HTTP_QUERY_IF_NONE_MATCH                = 57;
    const long HTTP_QUERY_IF_RANGE                     = 58;
    const long HTTP_QUERY_IF_UNMODIFIED_SINCE          = 59;
    const long HTTP_QUERY_MAX_FORWARDS                 = 60;
    const long HTTP_QUERY_PROXY_AUTHORIZATION          = 61;
    const long HTTP_QUERY_RANGE                        = 62;
    const long HTTP_QUERY_TRANSFER_ENCODING            = 63;
    const long HTTP_QUERY_UPGRADE                      = 64;
    const long HTTP_QUERY_VARY                         = 65;
    const long HTTP_QUERY_VIA                          = 66;
    const long HTTP_QUERY_WARNING                      = 67;

    const long HTTP_QUERY_MAX                          = 67;
    const long HTTP_QUERY_CUSTOM                       = 65535;
    const long HTTP_QUERY_FLAG_REQUEST_HEADERS         = 0x80000000;
    const long HTTP_QUERY_FLAG_SYSTEMTIME              = 0x40000000;
    const long HTTP_QUERY_FLAG_NUMBER                  = 0x20000000;
    const long HTTP_QUERY_FLAG_COALESCE                = 0x10000000;
    const long HTTP_STATUS_CONTINUE            = 100; // OK to continue with request
    const long HTTP_STATUS_SWITCH_PROTOCOLS    = 101; // server has switched protocols in upgrade header

    const long HTTP_STATUS_OK                  = 200; // request completed
    const long HTTP_STATUS_CREATED             = 201; // object created, reason = new URI
    const long HTTP_STATUS_ACCEPTED            = 202; // async completion (TBS)
    const long HTTP_STATUS_PARTIAL             = 203; // partial completion
    const long HTTP_STATUS_NO_CONTENT          = 204; // no info to return
    const long HTTP_STATUS_RESET_CONTENT       = 205; // request completed, but clear form
    const long HTTP_STATUS_PARTIAL_CONTENT     = 206; // partial GET furfilled

    const long HTTP_STATUS_AMBIGUOUS           = 300; // server couldn't decide what to return
    const long HTTP_STATUS_MOVED               = 301; // object permanently moved
    const long HTTP_STATUS_REDIRECT            = 302; // object temporarily moved
    const long HTTP_STATUS_REDIRECT_METHOD     = 303; // redirection w/ new access method
    const long HTTP_STATUS_NOT_MODIFIED        = 304; // if-modified-since was not modified
    const long HTTP_STATUS_USE_PROXY           = 305; // redirection to proxy, location header specifies proxy to use
    const long HTTP_STATUS_REDIRECT_KEEP_VERB  = 307; // HTTP/1.1: keep same verb

    const long HTTP_STATUS_BAD_REQUEST         = 400; // invalid syntax
    const long HTTP_STATUS_DENIED              = 401; // access denied
    const long HTTP_STATUS_PAYMENT_REQ         = 402; // payment required
    const long HTTP_STATUS_FORBIDDEN           = 403; // request forbidden
    const long HTTP_STATUS_NOT_FOUND           = 404; // object not found
    const long HTTP_STATUS_BAD_METHOD          = 405; // method is not allowed
    const long HTTP_STATUS_NONE_ACCEPTABLE     = 406; // no response acceptable to client found
    const long HTTP_STATUS_PROXY_AUTH_REQ      = 407; // proxy authentication required
    const long HTTP_STATUS_REQUEST_TIMEOUT     = 408; // server timed out waiting for request
    const long HTTP_STATUS_CONFLICT            = 409; // user should resubmit with more info
    const long HTTP_STATUS_GONE                = 410; // the resource is no longer available
    const long HTTP_STATUS_LENGTH_REQUIRED     = 411; // the server refused to accept request w/o a length
    const long HTTP_STATUS_PRECOND_FAILED      = 412; // precondition given in request failed
    const long HTTP_STATUS_REQUEST_TOO_LARGE   = 413; // request entity was too large
    const long HTTP_STATUS_URI_TOO_LONG        = 414; // request URI too long
    const long HTTP_STATUS_UNSUPPORTED_MEDIA   = 415; // unsupported media type

    const long HTTP_STATUS_SERVER_ERROR        = 500; // internal server error
    const long HTTP_STATUS_NOT_SUPPORTED       = 501; // required not supported
    const long HTTP_STATUS_BAD_GATEWAY         = 502; // error response received from gateway
    const long HTTP_STATUS_SERVICE_UNAVAIL     = 503; // temporarily overloaded
    const long HTTP_STATUS_GATEWAY_TIMEOUT     = 504; // timed out waiting for gateway
    const long HTTP_STATUS_VERSION_NOT_SUP     = 505; // HTTP version not supported

    const long HTTP_STATUS_FIRST               = 100;
    const long HTTP_STATUS_LAST                = 505;

//HERE
// values for dwModifiers parameter of HttpAddRequestHeaders()

    const long HTTP_ADDREQ_INDEX_MASK      = 0x0000FFFF;
    const long HTTP_ADDREQ_FLAGS_MASK      = 0xFFFF0000;

    const long HTTP_ADDREQ_FLAG_ADD_IF_NEW = 0x10000000;

    const long HTTP_ADDREQ_FLAG_ADD        = 0x20000000;

    const long HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA       = 0x40000000;
    const long HTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON   = 0x01000000;
    const long HTTP_ADDREQ_FLAG_COALESCE                  = 0x40000000;

    const long HTTP_ADDREQ_FLAG_REPLACE    = 0x80000000;

// flags for HttpSendRequestEx(), HttpEndRequest()

    const long HSR_ASYNC       = 0x00000001;          // force async
    const long HSR_SYNC        = 0x00000004;           // force sync
    const long HSR_USE_CONTEXT = 0x00000008;    // use dwContext value
    const long HSR_INITIATE    = 0x00000008;                      // iterative operation (completed by HttpEndRequest)
    const long HSR_DOWNLOAD    = 0x00000010;                      // download to file
    const long HSR_CHUNKED     = 0x00000020;                      // operation is send of chunked data


// InternetErrorDlg - Provides UI for certain Errors.

    const long FLAGS_ERROR_UI_FILTER_FOR_ERRORS        = 0x01;
    const long FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS     = 0x02;
    const long FLAGS_ERROR_UI_FLAGS_GENERATE_DATA      = 0x04;
    const long FLAGS_ERROR_UI_FLAGS_NO_UI              = 0x08;
    const long FLAGS_ERROR_UI_SERIALIZE_DIALOGS        = 0x10;

//
// Internet API error returns
//

    const long INTERNET_ERROR_BASE                    = 12000;

    const long ERROR_INTERNET_OUT_OF_HANDLES          = 12001;
    const long ERROR_INTERNET_TIMEOUT                 = 12002;
    const long ERROR_INTERNET_EXTENDED_ERROR          = 12003;
    const long ERROR_INTERNET_INTERNAL_ERROR          = 12004;
    const long ERROR_INTERNET_INVALID_URL             = 12005;
    const long ERROR_INTERNET_UNRECOGNIZED_SCHEME     = 12006;
    const long ERROR_INTERNET_NAME_NOT_RESOLVED       = 12007;
    const long ERROR_INTERNET_PROTOCOL_NOT_FOUND      = 12008;
    const long ERROR_INTERNET_INVALID_OPTION          = 120009;
    const long ERROR_INTERNET_BAD_OPTION_LENGTH       = 12010;
    const long ERROR_INTERNET_OPTION_NOT_SETTABLE     = 12011;
    const long ERROR_INTERNET_SHUTDOWN                = 12012;
    const long ERROR_INTERNET_INCORRECT_USER_NAME     = 12013;
    const long ERROR_INTERNET_INCORRECT_PASSWORD      = 12014;
    const long ERROR_INTERNET_LOGIN_FAILURE           = 12015;
    const long ERROR_INTERNET_INVALID_OPERATION       = 12016;
    const long ERROR_INTERNET_OPERATION_CANCELLED     = 12017;
    const long ERROR_INTERNET_INCORRECT_HANDLE_TYPE   = 12018;
    const long ERROR_INTERNET_INCORRECT_HANDLE_STATE  = 12019;
    const long ERROR_INTERNET_NOT_PROXY_REQUEST       = 12020;
    const long ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = 12021;
    const long ERROR_INTERNET_BAD_REGISTRY_PARAMETER  = 12022;
    const long ERROR_INTERNET_NO_DIRECT_ACCESS        = 12023;
    const long ERROR_INTERNET_NO_CONTEXT              = 12024;
    const long ERROR_INTERNET_NO_CALLBACK             = 12025;
    const long ERROR_INTERNET_REQUEST_PENDING         = 12026;
    const long ERROR_INTERNET_INCORRECT_FORMAT        = 12027;
    const long ERROR_INTERNET_ITEM_NOT_FOUND          = 12028;
    const long ERROR_INTERNET_CANNOT_CONNECT          = 12029;
    const long ERROR_INTERNET_CONNECTION_ABORTED      = 12030;
    const long ERROR_INTERNET_CONNECTION_RESET        = 12031;
    const long ERROR_INTERNET_FORCE_RETRY             = 12032;
    const long ERROR_INTERNET_INVALID_PROXY_REQUEST   = 12033;
    const long ERROR_INTERNET_NEED_UI                 = 12034;

    const long ERROR_INTERNET_HANDLE_EXISTS           = 12036;
    const long ERROR_INTERNET_SEC_CERT_DATE_INVALID   = 12037;
    const long ERROR_INTERNET_SEC_CERT_CN_INVALID     = 12038;
    const long ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR  = 12039;
    const long ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR  = 12040;
    const long ERROR_INTERNET_MIXED_SECURITY          = 12041;
    const long ERROR_INTERNET_CHG_POST_IS_NON_SECURE  = 12042;
    const long ERROR_INTERNET_POST_IS_NON_SECURE      = 12043;
    const long ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = 12044;
    const long ERROR_INTERNET_INVALID_CA              = 12045;
    const long ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP   = 12046;
    const long ERROR_INTERNET_ASYNC_THREAD_FAILED     = 12047;
    const long ERROR_INTERNET_REDIRECT_SCHEME_CHANGE  = 12048;
    const long ERROR_INTERNET_DIALOG_PENDING          = 12049;
    const long ERROR_INTERNET_RETRY_DIALOG            = 12050;
    const long ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = 12052;
    const long ERROR_INTERNET_INSERT_CDROM            = 12053;


//
// FTP API errors
//

    const long ERROR_FTP_TRANSFER_IN_PROGRESS         = 12110;
    const long ERROR_FTP_DROPPED                      = 12111;
    const long ERROR_FTP_NO_PASSIVE_MODE              = 12112;


//
// HTTP API errors
//

    const long ERROR_HTTP_HEADER_NOT_FOUND            = 12150;
    const long ERROR_HTTP_DOWNLEVEL_SERVER            = 12151;
    const long ERROR_HTTP_INVALID_SERVER_RESPONSE     = 12152;
    const long ERROR_HTTP_INVALID_HEADER              = 12153;
    const long ERROR_HTTP_INVALID_QUERY_REQUEST       = 12154;
    const long ERROR_HTTP_HEADER_ALREADY_EXISTS       = 12155;
    const long ERROR_HTTP_REDIRECT_FAILED             = 12156;
    const long ERROR_HTTP_NOT_REDIRECTED              = 12160;
    const long ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION   = 12161;
    const long ERROR_HTTP_COOKIE_DECLINED             = 12162;
    const long ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION = 12168;

//
// additional Internet API error codes
//

    const long ERROR_INTERNET_SECURITY_CHANNEL_ERROR  = 12157;
    const long ERROR_INTERNET_UNABLE_TO_CACHE_FILE    = 12158;
    const long ERROR_INTERNET_TCPIP_NOT_INSTALLED     = 12159;
    const long ERROR_INTERNET_DISCONNECTED            = 12163;
    const long ERROR_INTERNET_SERVER_UNREACHABLE      = 12164;
    const long ERROR_INTERNET_PROXY_SERVER_UNREACHABLE  = 12165;

    const long ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT     = 12166;
    const long ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = 12167;
    const long ERROR_INTERNET_SEC_INVALID_CERT   = 12169;
    const long ERROR_INTERNET_SEC_CERT_REVOKED   = 12170;

// InternetAutodial specific errors

    const long ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = 12171;
    const long INTERNET_ERROR_LAST = 12171;

//
// URLCACHE APIs
//

// cache entry type flags.

    const long NORMAL_CACHE_ENTRY              = 0x00000001;
    const long COOKIE_CACHE_ENTRY              = 0x00100000;
    const long URLHISTORY_CACHE_ENTRY          = 0x00200000;
    const long TRACK_OFFLINE_CACHE_ENTRY       = 0x00000010;
    const long TRACK_ONLINE_CACHE_ENTRY        = 0x00000020;
    const long STICKY_CACHE_ENTRY              = 0x00000004;
    const long SPARSE_CACHE_ENTRY              = 0x00010000;


    const long CACHE_ENTRY_ATTRIBUTE_FC    = 0x00000004;
    const long CACHE_ENTRY_HITRATE_FC      = 0x00000010;
    const long CACHE_ENTRY_MODTIME_FC      = 0x00000040;
    const long CACHE_ENTRY_EXPTIME_FC      = 0x00000080;
    const long CACHE_ENTRY_ACCTIME_FC      = 0x00000100;
    const long CACHE_ENTRY_SYNCTIME_FC     = 0x00000200;
    const long CACHE_ENTRY_HEADERINFO_FC   = 0x00000400;
    const long CACHE_ENTRY_EXEMPT_DELTA_FC = 0x00000800;

// Flags for SetUrlCacheEntryGroup
    const long INTERNET_CACHE_GROUP_ADD      = 0;
    const long INTERNET_CACHE_GROUP_REMOVE   = 1;


// Flags for InternetDial - must not conflict with InternetAutodial flags
//                          as they are valid here also.
    const long INTERNET_DIAL_UNATTENDED       = 0x8000;
    const long INTERENT_GOONLINE_REFRESH = 0x00000001;
    const long INTERENT_GOONLINE_MASK = 0x00000001;

// Flags for InternetAutodial
    const long INTERNET_AUTODIAL_FORCE_ONLINE          = 1;
    const long INTERNET_AUTODIAL_FORCE_UNATTENDED      = 2;
    const long INTERNET_AUTODIAL_FAILIFSECURITYCHECK   = 4;

// Flags for InternetGetConnectedState
    const long INTERNET_CONNECTION_MODEM           = 1;
    const long INTERNET_CONNECTION_LAN             = 2;
    const long INTERNET_CONNECTION_PROXY           = 4;
    const long INTERNET_CONNECTION_MODEM_BUSY      = 8;

// Flags for custom dial handler
    const long INTERNET_CUSTOMDIAL_CONNECT         = 0;
    const long INTERNET_CUSTOMDIAL_UNATTENDED      = 1;
    const long INTERNET_CUSTOMDIAL_DISCONNECT      = 2;
    const long INTERNET_CUSTOMDIAL_SHOWOFFLINE     = 4;

// Custom dial handler supported functionality flags
    const long INTERNET_CUSTOMDIAL_SAFE_FOR_UNATTENDED = 1;
    const long INTERNET_CUSTOMDIAL_WILL_SUPPLY_STATE   = 2;
    const long INTERNET_CUSTOMDIAL_CAN_HANGUP          = 4;

// States for InternetSetDialState
    const long INTERNET_DIALSTATE_DISCONNECTED     = 1;

};