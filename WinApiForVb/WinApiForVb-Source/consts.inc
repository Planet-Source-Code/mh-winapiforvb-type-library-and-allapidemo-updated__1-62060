[
	helpstring("String Constants"),
	dllname("NOSUCHS.DLL")
]
module MiscConsts {

    const LPSTR CHAR_ZERO_LENGTH_STRING = "";
	const LPSTR CHAR_SPACE = " ";
    
	const LPSTR CHAR_CRLF_CRLF = "\r\n\r\n";

	const LPSTR CHAR_BACK_SLASH			= "\\";
	[ helpstring("Double backslash (network directory indicator)") ]
	const LPSTR CHAR_DOUBLE_BACK_SLASH	= "\\\\";
    [ helpstring("Path separators") ]
    const LPSTR sPathSeparators			= "\:";
	const LPSTR CHAR_FORWARD_SLASH		= "/";
	
	[ helpstring("Colon :") ]
    const LPSTR CHAR_COLON				= ":";
    const LPSTR CHAR_DOT				= ".";
	[ helpstring("Semi colon ;") ]
	const LPSTR CHAR_SEMI_COLON			= ";";
    [ helpstring("Single quote") ]
    const LPSTR CHAR_SINGLE_QUATAION	= "'";
    [ helpstring("Double quote") ]
    const LPSTR CHAR_DOUBLE_QUATAION	= "\"";
	const LPSTR CHAR_PIPE				= "|";
	const LPSTR CHAR_RIGHT_ARROW		= ">";
	const LPSTR CHAR_LEFT_ARROW			= "<";
	
// 96 `
	const LPSTR CHAR_PLUS_SIGN			= "+";
	const LPSTR CHAR_PERCENT_SIGN		= "%";
	const LPSTR CHAR_STAR				= "*";
	const LPSTR CHAR_QUESTION			= "?";
	const LPSTR CHAR_EQUAL				= "=";
	const LPSTR CHAR_RIGHT_BRACKET		= "}";
	const LPSTR CHAR_LEFT_BRAKET		= "{";
	const LPSTR CHAR_RIGHT_PARANTHESIS	= ")";
	const LPSTR CHAR_LEFT_PARANTHESIS	= "(";
	const LPSTR CHAR_DOLLAR_SIGN		= "$";
	const LPSTR CHAR_POUND				= "#";
	const LPSTR CHAR_EMAIL_SIGN			= "@";
	const LPSTR CHAR_EXCLAMATION		= "!";
	const LPSTR CHAR_AMPERCENT			= "&";
	const LPSTR CHAR_MINUS				= "-";
	const LPSTR CHAR_UNDERSCORE			= "_";
	const LPSTR CHAR_UPARROW			= "^";
	const LPSTR CHAR_WIGILLIE_SIGN		= "~";
	
	const long ASCII_BELL				= 7;
	const long ASCII_BACKSPACE			= 8;
	const long ASCII_VERTICALTAB		= 11;
	const long ASCII_FORMFEED			= 14;
	const long ASCII_SEMI_COLON			= 59;
	const long ASCII_COLON				= 58;
	const long ASCII_SINGLE_QUATATION	= 34; 
	const long ASCII_DOUBLE_QUATATION	= 39;
	const long ASCII_STAR				= 42;
	const long ASCII_QUESTION			= 63;
	const long ASCII_PIPE				= 124;
	const long ASCII_RIGHT_ARROW		= 62;
	const long ASCII_LEFT_ARROW			= 60;
	const long ASCII_SPACE				= 32;
	const long ASCII_BACK_SLASH			= 92;
	const long ASCII_FORWARD_SLASH		= 47;
	const long ASCII_DOT				= 46;
	const long ASCII_EQUAL				= 61;
	const long ASCII_RIGHT_BRACKET		= 125;
	const long ASCII_LEFT_BRAKET		= 123;
	const long ASCII_RIGHT_PARANTHESIS	= 41;
	const long ASCII_LEFT_PARANTHESIS	= 40;
	const long ASCII_DOLLAR_SIGN		= 36;
	const long ASCII_POUND				= 35;
	const long ASCII_EMAIL_SIGN			= 64;
	const long ASCII_EXCLAMATION		= 33;
	const long ASCII_AMPERCENT			= 38;
	const long ASCII_MINUS				= 45;
	const long ASCII_UNDERSCORE			= 95;
	const long ASCII_UPARROW			= 94;
	const long ASCII_WIGILLIE_SIGN		= 126;


    // Keycodes that are missing or have confusing names in VB type library
    [ helpstring("Page Up key.") ]
    const int vbKeyPrior = 0x21;
    [ helpstring("Page Down key.") ]
    const int vbKeyNext = 0x22;

	const LPSTR S_ABOUT_BLANK	= "about:blank";
	const LPSTR S_FTP			= "ftp://";
	const LPSTR S_HTTP			= "http://";
	const LPSTR S_HTTPS			= "https://";

	const LPSTR FTP_GENERIC_USERNAME = "anonymous";
	const LPSTR FTP_GENERIC_PASSWORD = "user@nowhere.com";
	
	[ helpstring("For FindFiles filter *.*") ]
	const LPSTR CHAR_FIND_ALL_FILES_FILTER	= "*.*";
	const LPSTR CHAR_DOUBLE_DOT				= "..";
	const LPSTR CHAR_TRIPLE_DOT				= "...";

    // MIDL bug generates empty string rather than null character
    //[ helpstring("Null character (ASCII 0)") ]
    // const LPSTR sNullChar = "\0";
	const LPSTR CHAR_ZERO		= "0";
	const LPSTR CHAR_ONE		= "1";
	const LPSTR CHAR_TWO		= "2";
	const LPSTR CHAR_THREE		= "3";
	const LPSTR CHAR_FOUR		= "4";
	const LPSTR CHAR_FIVE		= "5";
	const LPSTR CHAR_SIX		= "6";
	const LPSTR CHAR_SEVEN		= "7";
	const LPSTR CHAR_EIGHT		= "8";
	const LPSTR CHAR_NINE		= "9";
	const LPSTR CHAR_TEN		= "10";

//A-Z 65-90 VK_A to VK_Z
	const LPSTR CHAR_U_A ="A";
	const LPSTR CHAR_U_B ="B";
	const LPSTR CHAR_U_C ="C";
	const LPSTR CHAR_U_D ="D";
	const LPSTR CHAR_U_E ="E";
	const LPSTR CHAR_U_F ="F";
	const LPSTR CHAR_U_G ="G";
	const LPSTR CHAR_U_H ="H";
	const LPSTR CHAR_U_I ="I";
	const LPSTR CHAR_U_J ="J";
	const LPSTR CHAR_U_K ="K";
	const LPSTR CHAR_U_L ="L";
	const LPSTR CHAR_U_M ="M";
	const LPSTR CHAR_U_N ="N";
	const LPSTR CHAR_U_O ="O";
	const LPSTR CHAR_U_P ="P";
	const LPSTR CHAR_U_Q ="Q";
	const LPSTR CHAR_U_R ="R";
	const LPSTR CHAR_U_S ="S";
	const LPSTR CHAR_U_T ="T";
	const LPSTR CHAR_U_U ="U";
	const LPSTR CHAR_U_V ="V";
	const LPSTR CHAR_U_W ="W";
	const LPSTR CHAR_U_X ="X";
	const LPSTR CHAR_U_Y ="Y";
	const LPSTR CHAR_U_Z ="Z";

	const LPSTR CHAR_a ="a";
	const long ASCII_a = 97;
	const LPSTR CHAR_b ="b";
	const long ASCII_b = 98;
	const LPSTR CHAR_c ="c";
	const long ASCII_c = 99;
	const LPSTR CHAR_d ="d";
	const long ASCII_d = 100;
	const LPSTR CHAR_e ="e";
	const long ASCII_e = 101;
	const LPSTR CHAR_f ="f";
	const long ASCII_f = 102;
	const LPSTR CHAR_g ="g";
	const long ASCII_g = 103;
	const LPSTR CHAR_h ="h";
	const long ASCII_h = 104;
	const LPSTR CHAR_i ="i";
	const long ASCII_i = 105;
	const LPSTR CHAR_j ="j";
	const long ASCII_j = 106;
	const LPSTR CHAR_k ="k";
	const long ASCII_k = 107;
	const LPSTR CHAR_l ="l";
	const long ASCII_l = 108;
	const LPSTR CHAR_m ="m";
	const long ASCII_m = 109;
	const LPSTR CHAR_n ="n";
	const long ASCII_n = 110;
	const LPSTR CHAR_o ="o";
	const long ASCII_o = 111;
	const LPSTR CHAR_p ="p";
	const long ASCII_p = 112;
	const LPSTR CHAR_q ="q";
	const long ASCII_q = 113;
	const LPSTR CHAR_r ="r";
	const long ASCII_r = 114;
	const LPSTR CHAR_s ="s";
	const long ASCII_s = 115;
	const LPSTR CHAR_t ="t";
	const long ASCII_t = 116;
	const LPSTR CHAR_u ="u";
	const long ASCII_u = 117;
	const LPSTR CHAR_v ="v";
	const long ASCII_v = 118;
	const LPSTR CHAR_w ="w";
	const long ASCII_w = 119;
	const LPSTR CHAR_x ="x";
	const long ASCII_x = 120;
	const LPSTR CHAR_y ="y";
	const long ASCII_y = 121;
	const LPSTR CHAR_z ="z";
	const long ASCII_z = 122;

    const LPSTR sExeExt = ".EXE";
    const LPSTR sPifExt = ".PIF";
    const LPSTR sCmdExt = ".CMD";
    const LPSTR sBatExt = ".BAT";
    const LPSTR sComExt = ".COM";
	const LPSTR sSysExt	= ".SYS";

	const LPSTR S_FILE_TIME_FORMAT	= "mm/dd/yyyy hh:nn";
	const LPSTR S_AM				= " AM";
	const LPSTR S_PM				= " PM";

	const long NUM_ONE_THOUSAND_TWENTY_FOUR = 1024;
	const long NUM_SIXTY					= 60;
	const long NUM_TWENTY 					= 20;
	const long NUM_NINTEEN 					= 19;
	const long NUM_EIGHTEEN 				= 18;
	const long NUM_SEVENTEEN 				= 17;
	const long NUM_SIXTEEN 					= 16;
	const long NUM_FIFTEEN 					= 15;
	const long NUM_FOURTEEN 				= 14;
	const long NUM_THIRTEEN 				= 13;
	const long NUM_TWELEVE 					= 12;
	const long NUM_ELEVEN 					= 11;
	const long NUM_TEN 						= 10;
	const long NUM_NINE 					= 9;
	const long NUM_EIGHT 					= 8;
	const long NUM_SEVEN 					= 7;
	const long NUM_SIX 						= 6;
	const long NUM_FIVE 					= 5;
	const long NUM_FOUR 					= 4;
	const long NUM_THREE 					= 3;
	const long NUM_TWO 						= 2;
	const long NUM_ONE 						= 1;
	const long NUM_ZERO 					= 0;
	const long NUM_MINUS_ONE 				= -1;

// ******

	const LPSTR WNDCTRL_HEADER				= "SysHeader32";
	const LPSTR WNDCTRL_STATIC				= "STATIC";
	const LPSTR WNDCTRL_BUTTON				= "BUTTON";
	const LPSTR WNDCTRL_EDIT				= "EDIT";
	const LPSTR WNDCTRL_LISTBOX				= "ListBox";
    const LPSTR WNDCTRL_TOOLBARCLASSNAME	= "ToolbarWindow32";
	const LPSTR WNDCTRL_REBARCLASSNAME		= "ReBarWindow32";
	const LPSTR WNDCTRL_TOOLTIPS_CLASS		= "tooltips_class32";
	const LPSTR WNDCTRL_STATUSCLASSNAMEW	= "msctls_statusbar32";
	const LPSTR WNDCTRL_TRACKBAR_CLASS		= "msctls_trackbar32";
	const LPSTR WNDCTRL_DRAGLISTMSGSTRING	= "commctrl_DragListMsg";
    const LPSTR WNDCTRL_UPDOWN_CLASS		= "msctls_updown32";
	const LPSTR WNDCTRL_PROGRESS_CLASS		= "msctls_progress32";
	const LPSTR WNDCTRL_HOTKEY_CLASS		= "msctls_hotkey32";
    const LPSTR WNDCTRL_LISTVIEW			= "SysListView32";
    const LPSTR WNDCTRL_TREEVIEW			= "SysTreeView32";
	const LPSTR WNDCTRL_COMBOBOX			= "COMBOBOX";
    const LPSTR WNDCTRL_COMBOBOXEX			= "ComboBoxEx32";
	const LPSTR WNDCTRL_TABCONTROLA			= "SysTabControl32";
	const LPSTR WNDCTRL_ANIMATE_CLASS		= "SysAnimate32";
	const LPSTR WNDCTRL_MONTHCAL_CLASS		= "SysMonthCal32";
	const LPSTR WNDCTRL_DATETIMEPICK_CLASS	= "SysDateTimePick32";
	const LPSTR WNDCTRL_IPADDRESS			= "SysIPAddress32";
	const LPSTR WNDCTRL_PAGESCROLLER		= "SysPager";
	const LPSTR WNDCTRL_NATIVEFONTCTL		= "NativeFontCtl";
	const LPSTR WNDCTRL_RICHEDIT_CLASS10A	= "RICHEDIT";
	const LPSTR WNDCTRL_RICHEDIT_CLASS20A	= "RichEdit20A";
	const LPSTR WNDCTRL_RICHEDIT_CLASS20W	= "RichEdit20W";

    // registry entries for special paths are kept in :
    const LPSTR REGSTR_PATH_SPECIAL_FOLDERS  = "\\Shell Folders";

	const LPSTR HTTP_VERSION			= "HTTP/1.0";
	const LPSTR UD_FLAGS_STR			= "/ud";
    

    // Mime types
    const LPSTR CFSTR_MIME_NULL			= 0;
    const LPSTR CFSTR_MIME_TEXT         = "text/plain";
    const LPSTR CFSTR_MIME_RICHTEXT     = "text/richtext";
    const LPSTR CFSTR_MIME_X_BITMAP     = "image/x-xbitmap";
    const LPSTR CFSTR_MIME_POSTSCRIPT   = "application/postscript";
    const LPSTR CFSTR_MIME_AIFF         = "audio/aiff";
    const LPSTR CFSTR_MIME_BASICAUDIO   = "audio/basic";
    const LPSTR CFSTR_MIME_WAV          = "audio/wav";
    const LPSTR CFSTR_MIME_X_WAV        = "audio/x-wav";
    const LPSTR CFSTR_MIME_GIF          = "image/gif";
    const LPSTR CFSTR_MIME_PJPEG        = "image/pjpeg";
    const LPSTR CFSTR_MIME_JPEG         = "image/jpeg";
    const LPSTR CFSTR_MIME_TIFF         = "image/tiff";
    const LPSTR CFSTR_MIME_X_PNG        = "image/x-png";
    const LPSTR CFSTR_MIME_BMP          = "image/bmp";
    const LPSTR CFSTR_MIME_X_ART        = "image/x-jg";
    const LPSTR CFSTR_MIME_X_EMF        = "image/x-emf";
    const LPSTR CFSTR_MIME_X_WMF        = "image/x-wmf";
    const LPSTR CFSTR_MIME_AVI          = "video/avi";
    const LPSTR CFSTR_MIME_MPEG         = "video/mpeg";
    const LPSTR CFSTR_MIME_FRACTALS     = "application/fractals";
    const LPSTR CFSTR_MIME_RAWDATA      = "application/octet-stream";
    const LPSTR CFSTR_MIME_RAWDATASTRM  = "application/octet-stream";
    const LPSTR CFSTR_MIME_PDF          = "application/pdf";
    const LPSTR CFSTR_MIME_HTA          = "application/hta";
	const LPSTR CFSTR_MIME_CDFA			= "application/x-cdf";
    const LPSTR CFSTR_MIME_X_AIFF       = "audio/x-aiff";
    const LPSTR CFSTR_MIME_X_REALAUDIO  = "audio/x-pn-realaudio";
    const LPSTR CFSTR_MIME_XBM          = "image/xbm";
    const LPSTR CFSTR_MIME_QUICKTIME    = "video/quicktime";
    const LPSTR CFSTR_MIME_X_MSVIDEO    = "video/x-msvideo";
    const LPSTR CFSTR_MIME_X_SGI_MOVIE  = "video/x-sgi-movie";
    const LPSTR CFSTR_MIME_HTML         = "text/html";
    const LPSTR CFSTR_MIME_XML          = "text/xml";

    // clipboard formats - use as parameter to RegisterClipboardFormat()
    const LPSTR  CF_RTF          = "Rich Text Format";
    const LPSTR  CF_RTFNOOBJS    = "Rich Text Format Without Objects";
    const LPSTR  CF_RETEXTOBJ    = "RichEdit Text and Objects";


/*
	//Not used for now
    // Clipboard Formats

    const LPSTR CFSTR_SHELLIDLIST         = "Shell IDList Array";
    const LPSTR CFSTR_SHELLIDLISTOFFSET   = "Shell Object Offsets";
    const LPSTR CFSTR_NETRESOURCES        = "Net Resource";
    const LPSTR CFSTR_FILEDESCRIPTORA     = "FileGroupDescriptor";
    const LPSTR CFSTR_FILEDESCRIPTORW     = "FileGroupDescriptorW";
    const LPSTR CFSTR_FILECONTENTS        = "FileContents";
    const LPSTR CFSTR_FILENAMEA           = "FileName";
    const LPSTR CFSTR_FILENAMEW           = "FileNameW";
    const LPSTR CFSTR_PRINTERGROUP        = "PrinterFriendlyName";
    const LPSTR CFSTR_FILENAMEMAPA        = "FileNameMap";
    const LPSTR CFSTR_FILENAMEMAPW        = "FileNameMapW";
    const LPSTR CFSTR_SHELLURL            = "UniformResourceLocator";
    const LPSTR CFSTR_INETURLA            = "UniformResourceLocator";
    const LPSTR CFSTR_INETURLW            = "UniformResourceLocatorW";
    const LPSTR CFSTR_PREFERREDDROPEFFECT = "Preferred DropEffect";
    const LPSTR CFSTR_PERFORMEDDROPEFFECT = "Performed DropEffect";
    const LPSTR CFSTR_PASTESUCCEEDED      = "Paste Succeeded";
    const LPSTR CFSTR_INDRAGLOOP          = "InShellDragLoop";
    const LPSTR CFSTR_DRAGCONTEXT         = "DragContext";
    const LPSTR CFSTR_MOUNTEDVOLUME       = "MountedVolume";
    const LPSTR CFSTR_PERSISTEDDATAOBJECT = "PersistedDataObject";
    const LPSTR CFSTR_TARGETCLSID         = "TargetCLSID";   // HGLOBAL with a CLSID of the drop target
    const LPSTR CFSTR_LOGICALPERFORMEDDROPEFFECT	= "Logical Performed DropEffect";
    const LPSTR CFSTR_OutlookExpressMessages		= "Outlook Express Messages";
    const LPSTR CFSTR_HYPERLINK						= "Hyperlink";

    // Command group IDs
    const LPSTR CGIDSTR_EditStateCommands     = "{3050f4b6-98b5-11cf-bb82-00aa00bdce0b}";
    const LPSTR CGIDSTR_DocHostCommandHandler = "{f38bc242-b950-11d1-8918-00c04fc2c836}";
    const LPSTR CGIDSTR_DeskBand              = "{EB0FE172-1A3A-11D0-89B3-00A0C90A90AC}";
    const LPSTR CGIDSTR_Explorer              = "{000214D0-0000-0000-C000-000000000046}";
    const LPSTR CGIDSTR_ShellDocView          = "{000214D1-0000-0000-C000-000000000046}";
    const LPSTR CGIDSTR_ShellServiceObject    = "{000214D2-0000-0000-C000-000000000046}";
    const LPSTR CGIDSTR_ExplorerBarDoc        = "{000214D3-0000-0000-C000-000000000046}";
    const LPSTR CGIDSTR_IDeskBand             = "{EB0FE172-1A3A-11D0-89B3-00A0C90A90AC}";
    const LPSTR CGIDSTR_ShortCut              = "{93a68750-951a-11d1-946f-000000000000}";
    const LPSTR CGIDSTR_HTML                  = "{DE4BA900-59CA-11CF-9592-444553540000}";

    // Interface IDs
    const LPSTR IIDSTR_IUnknown                 = "{00000000-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IDispatch                = "{00020400-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IConnectionPoint         = "{B196B286-BAB4-101A-B69C-00AA00341D07}";
    const LPSTR IIDSTR_IConnectionPointContainer= "{B196B284-BAB4-101A-B69C-00AA00341D07}";
    const LPSTR IIDSTR_IOleClientSite           = "{00000118-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IDocHostShowUI           = "{C4D244B0-D43E-11CF-893B-00AA00BDCE1A}";
    const LPSTR IIDSTR_IDocHostUIHandler        = "{BD3F23C0-D43E-11CF-893B-00AA00BDCE1A}";
    const LPSTR IIDSTR_IServiceProvider         = "{6D5140C1-7436-11CE-8034-00AA006009FA}";
    const LPSTR IIDSTR_IWebBrowserApp           = "{0002DF05-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IInternetSecurityManager = "{79eac9ee-baf9-11ce-8c82-00aa004ba90b}";
    const LPSTR IIDSTR_StandardFont             = "{7EBDAAE0-8120-11CF-899F-00AA00688B10}";
    const LPSTR IIDSTR_StandardColor            = "{7EBDAAE1-8120-11CF-899F-00AA00688B10}";
    const LPSTR IIDSTR_StandardPicture          = "{7EBDAAE2-8120-11CF-899F-00AA00688B10}";    
    const LPSTR IIDSTR_IPicture                 = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}";
    const LPSTR IIDSTR_IPictureDisp             = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}";
    const LPSTR IIDSTR_IFont                    = "{BEF6E002-A874-101A-8BBA-00AA00300CAB}";
    const LPSTR IIDSTR_IFontDisp                = "{BEF6E003-A874-101A-8BBA-00AA00300CAB}";
    const LPSTR IIDSTR_IQueryAssociations       = "{c46ca590-3c3f-11d2-bee6-0000f805ca57}";
    const LPSTR IIDSTR_IContextMenu             = "{000214E4-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IContextMenu2            = "{000214F4-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IQueryInfo               = "{000214E7-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IExtractIconA            = "{000214EB-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IExtractIconW            = "{000214FA-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IShellFolder             = "{000214E6-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IShellView               = "{000214E3-0000-0000-C000-000000000046}";
    const LPSTR IIDSTR_IShellView2              = "{88E39E80-3578-11CF-AE69-08002B2E1262}";
    const LPSTR IIDSTR_IRedbookDiscMaster       = "{E3BC42CD-4E5C-11D3-9144-00104BA11C5E)}";
    const LPSTR IIDSTR_IJolietDiscMaster        = "{E3BC42CE-4E5C-11D3-9144-00104BA11C5E}";

    // CLSIDs
    const LPSTR CLSIDSTR_InternetSecurityManager = "{7B8A2D94-0AC9-11D1-896C-00C04FB6BFC4}";
    const LPSTR CLSIDSTR_InternetZoneManager     = "{7B8A2D95-0AC9-11D1-896C-00C04FB6BFC4}";
    const LPSTR CLSIDSTR_QueryAssociations       = "{c46ca590-3c3f-11d2-bee6-0000f805ca57}";
    const LPSTR CLSIDSTR_ShellDesktop            = "{00021400-0000-0000-C000-000000000046}";
    const LPSTR CLSIDSTR_ShellLink               = "{00021401-0000-0000-C000-000000000046}";
    const LPSTR CLSIDSTR_NetworkPlaces           = "{208D2C60-3AEA-1069-A2D7-08002B30309D}";
    const LPSTR CLSIDSTR_NetworkDomain           = "{46e06680-4bf0-11d1-83ee-00a0c90dc849}";
    const LPSTR CLSIDSTR_NetworkServer           = "{c0542a90-4bf0-11d1-83ee-00a0c90dc849}";
    const LPSTR CLSIDSTR_NetworkShare            = "{54a754c0-4bf1-11d1-83ee-00a0c90dc849}";
    const LPSTR CLSIDSTR_MyComputer              = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
    const LPSTR CLSIDSTR_Internet                = "{871C5380-42A0-1069-A2EA-08002B30309D}";
    const LPSTR CLSIDSTR_ShellFSFolder           = "{F3364BA0-65B9-11CE-A9BA-00AA004AE837}";
    const LPSTR CLSIDSTR_RecycleBin              = "{645FF040-5081-101B-9F08-00AA002F954E}";
    const LPSTR CLSIDSTR_ControlPanel            = "{21EC2020-3AEA-1069-A2DD-08002B30309D}";
    const LPSTR CLSIDSTR_Printers                = "{2227A280-3AEA-1069-A2DE-08002B30309D}";
    const LPSTR CLSIDSTR_MyDocuments             = "{450D8FBA-AD25-11D0-98A8-0800361B1103}";
*/
};