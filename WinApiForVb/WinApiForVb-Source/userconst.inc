// ******* User32 consts
[
	helpstring("User32 Constants"),
	dllname("NOSUCHU.DLL")
]
module UserConst {
//WM_USER = 1024
	const long MAX_PATH = 260;
	const long MAX_PATH_WIDE = 520;
	const long MAX_FILE_PATH = 260;
	const long MAX_PATH_PLUS_ONE = 261;
       
	const long MAXCHAR    = 0x7f;
	const long MAXSHORT   = 0x7fff;
	const long MAXLONG    = 0x7fffffff;
	const long MAXBYTE    = 0xff;
	const long MAXWORD    = 0xffff;
	//0xffffffff returns -1 in VB, out of range
	const double MAXDWORD   = 4294967296; //0xffffffff 2 ^ 32; 

//Virtual keys
    const long VK_LBUTTON    = 0x1;
    const long VK_RBUTTON    = 0x2;
    const long VK_CANCEL     = 0x3;
    const long VK_MBUTTON    = 0x4;
    const long VK_BACK       = 0x8;
    const long VK_TAB        = 0x9;
    const long VK_CLEAR      = 0xC;
    const long VK_RETURN     = 0xD;
    const long VK_SHIFT      = 0x10;
    const long VK_CONTROL    = 0x11;
    const long VK_MENU       = 0x12;
    const long VK_PAUSE      = 0x13;
    const long VK_CAPITAL    = 0x14;
    const long VK_ESCAPE     = 0x1B;
    const long VK_SPACE      = 0x20;
    const long VK_PRIOR      = 0x21;
    const long VK_NEXT       = 0x22;
    const long VK_END        = 0x23;
    const long VK_HOME       = 0x24;
    const long VK_LEFT       = 0x25;
    const long VK_UP         = 0x26;
    const long VK_RIGHT      = 0x27;
    const long VK_DOWN       = 0x28;
    const long VK_SELECT     = 0x29;
    const long VK_PRINT      = 0x2A;
    const long VK_EXECUTE    = 0x2B;
    const long VK_SNAPSHOT   = 0x2C;
    const long VK_INSERT     = 0x2D;
    const long VK_DELETE     = 0x2E;
    const long VK_HELP       = 0x2F;
    const long VK_0          = 0x30;
    const long VK_1          = 0x31;
    const long VK_2          = 0x32;
    const long VK_3          = 0x33;
    const long VK_4          = 0x34;
    const long VK_5          = 0x35;
    const long VK_6          = 0x36;
    const long VK_7          = 0x37;
    const long VK_8          = 0x38;
    const long VK_9          = 0x39;
    const long VK_A          = 0x41;
    const long VK_B          = 0x42;
    const long VK_C          = 0x43;
    const long VK_D          = 0x44;
    const long VK_E          = 0x45;
    const long VK_F          = 0x46;
    const long VK_G          = 0x47;
    const long VK_H          = 0x48;
    const long VK_I          = 0x49;
    const long VK_J          = 0x4A;
    const long VK_K          = 0x4B;
    const long VK_L          = 0x4C;
    const long VK_M          = 0x4D;
    const long VK_N          = 0x4E;
    const long VK_O          = 0x4F;
    const long VK_P          = 0x50;
    const long VK_Q          = 0x51;
    const long VK_R          = 0x52;
    const long VK_S          = 0x53;
    const long VK_T          = 0x54;
    const long VK_U          = 0x55;
    const long VK_V          = 0x56;
    const long VK_W          = 0x57;
    const long VK_X          = 0x58;
    const long VK_Y          = 0x59;
    const long VK_Z          = 0x5A;
    const long VK_NUMPAD0    = 0x60;
    const long VK_NUMPAD1    = 0x61;
    const long VK_NUMPAD2    = 0x62;
    const long VK_NUMPAD3    = 0x63;
    const long VK_NUMPAD4    = 0x64;
    const long VK_NUMPAD5    = 0x65;
    const long VK_NUMPAD6    = 0x66;
    const long VK_NUMPAD7    = 0x67;
    const long VK_NUMPAD8    = 0x68;
    const long VK_NUMPAD9    = 0x69;
    const long VK_MULTIPLY   = 0x6A;
    const long VK_ADD        = 0x6B;
    const long VK_SEPARATOR  = 0x6C;
    const long VK_SUBTRACT   = 0x6D;
    const long VK_DECIMAL    = 0x6E;
    const long VK_DIVIDE     = 0x6F;
    const long VK_F1         = 0x70;
    const long VK_F2         = 0x71;
    const long VK_F3         = 0x72;
    const long VK_F4         = 0x73;
    const long VK_F5         = 0x74;
    const long VK_F6         = 0x75;
    const long VK_F7         = 0x76;
    const long VK_F8         = 0x77;
    const long VK_F9         = 0x78;
    const long VK_F10        = 0x79;
    const long VK_F11        = 0x7A;
    const long VK_F12        = 0x7B;
    const long VK_F13        = 0x7C;
    const long VK_F14        = 0x7D;
    const long VK_F15        = 0x7E;
    const long VK_F16        = 0x7F;
    const long VK_F17        = 0x80;
    const long VK_F18        = 0x81;
    const long VK_F19        = 0x82;
    const long VK_F20        = 0x83;
    const long VK_F21        = 0x84;
    const long VK_F22        = 0x85;
    const long VK_F23        = 0x86;
    const long VK_F24        = 0x87;
    const long VK_NUMLOCK    = 0x90;
    const long VK_SCROLL     = 0x91;

    const long VK_LWIN         = 0x5B; //Left Windows key (Microsoft® Natural® keyboard)
    const long VK_RWIN         = 0x5C; //Right Windows key (Natural keyboard)
    const long VK_APPS         = 0x5D; //Applications key (Natural keyboard)

    const long VK_KANA         =  0x15;
    const long VK_HANGEUL      =  0x15;  /* old name - should be here for compatibility */
    const long VK_HANGUL       =  0x15;
    const long VK_JUNJA        =  0x17;
    const long VK_FINAL        =  0x18;
    const long VK_HANJA        =  0x19;
    const long VK_KANJI        =  0x19;

    const long VK_CONVERT      =  0x1C;
    const long VK_NONCONVERT   =  0x1D;
    const long VK_ACCEPT       =  0x1E;
    const long VK_MODECHANGE   =  0x1F;

/*
 * VK_L* & VK_R* - left and right Alt, Ctrl and Shift virtual keys.
 * Used only as parameters to GetAsyncKeyState() and GetKeyState().
 * No other API or message will distinguish left and right keys in this way.
 */
const long VK_LSHIFT         =  0xA0;
const long VK_RSHIFT         =  0xA1;
const long VK_LCONTROL       =  0xA2;
const long VK_RCONTROL       =  0xA3;
const long VK_LMENU          =  0xA4;
const long VK_RMENU          =  0xA5;

const long VK_PROCESSKEY     =  0xE5;

const long VK_ATTN           =  0xF6;
const long VK_CRSEL          =  0xF7;
const long VK_EXSEL          =  0xF8;
const long VK_EREOF          =  0xF9;
const long VK_PLAY           =  0xFA;
const long VK_ZOOM           =  0xFB;
const long VK_NONAME         =  0xFC;
const long VK_PA1            =  0xFD;
const long VK_OEM_CLEAR      =  0xFE;

const long VK_BROWSER_BACK = 166;
const long VK_BROWSER_FAVORITES = 171;
const long VK_BROWSER_FORWARD = 167;
const long VK_BROWSER_HOME = 172;
const long VK_BROWSER_REFRESH = 168;
const long VK_BROWSER_SEARCH = 170;
const long VK_BROWSER_STOP = 169;

const long VK_ICO_00 = 228; 
const long VK_ICO_CLEAR = 230;
const long VK_ICO_HELP = 227;
const long VK_LAUNCH_APP1 = 182; //Windows 2000
const long VK_LAUNCH_APP2 = 183; //Windows 2000
const long VK_LAUNCH_MAIL = 180; //Windows 2000
const long VK_LAUNCH_MEDIA_SELECT = 181; //Windows 2000
const long VK_MEDIA_NEXT_TRACK = 176;
const long VK_MEDIA_PLAY_PAUSE = 179;
const long VK_MEDIA_PREV_TRACK = 177;
const long VK_MEDIA_STOP = 178;

const long VK_SLEEP = 95;
const long VK_VOLUME_DOWN = 174;
const long VK_VOLUME_MUTE = 173;
const long VK_VOLUME_UP = 175;

	// **** Window Messages ****

    // WM_SHOWWINDOW wParam codes
    const long SW_PARENTCLOSING =    1;
    const long SW_OTHERZOOM     =    2;
    const long SW_PARENTOPENING =    3;
    const long SW_OTHERUNZOOM   =    4; 

    // WM_KEYUP/DOWN/CHAR HIWORD(lParam) flags
    const long KF_EXTENDED     = 0x0100;
    const long KF_DLGMODE      = 0x0800;
    const long KF_MENUMODE     = 0x1000;
    const long KF_ALTDOWN      = 0x2000;
    const long KF_REPEAT       = 0x4000;
    const long KF_UP           = 0x8000;


    // NOTE: All messages below 0x0400 are RESERVED by Windows
    const long WM_USER			= 0x0400;

    const long WM_NULL          = 0x0000;
    const long WM_CREATE        = 0x0001;
    const long WM_DESTROY       = 0x0002;
    const long WM_MOVE          = 0x0003;
    const long WM_SIZE          = 0x0005;

    const long WM_ACTIVATE      = 0x0006;
    const long WM_SETFOCUS      = 0x0007;
    const long WM_KILLFOCUS     = 0x0008;
    const long WM_ENABLE        = 0x000A;
    const long WM_SETREDRAW     = 0x000B;
    const long WM_SETTEXT       = 0x000C;
    const long WM_GETTEXT       = 0x000D;
    const long WM_GETTEXTLENGTH = 0x000E;
    const long WM_PAINT         = 0x000F;
    const long WM_CLOSE			= 0x0010;

    const long WM_QUERYENDSESSION  = 0x0011;
    const long WM_QUIT             = 0x0012;
    const long WM_QUERYOPEN     = 0x0013;
    const long WM_ERASEBKGND    = 0x0014;
    const long WM_SYSCOLORCHANGE = 0x0015;
    const long WM_ENDSESSION       = 0x0016;
    const long WM_SHOWWINDOW    = 0x0018;
    const long WM_WININICHANGE  = 0x001A;
    const long WM_SETTINGCHANGE	= 0x001A;

    const long WM_DEVMODECHANGE	= 0x001B;
	const long WM_ACTIVATEAPP	= 0x001C;
    const long WM_FONTCHANGE		= 0x001D;
	const long WM_TIMECHANGE		= 0x001E;
    const long WM_CANCELMODE		= 0x001F;
    const long WM_SETCURSOR     = 0x0020;
    const long WM_MOUSEACTIVATE = 0x0021;
    const long WM_CHILDACTIVATE = 0x0022;
	const long WM_QUEUESYNC     = 0x0023;

    const long WM_GETMINMAXINFO = 0x0024;


	const long WM_PAINTICON     = 0x0026;
    const long WM_ICONERASEBKGND = 0x0027;
    const long WM_NEXTDLGCTL    = 0x0028;
    const long WM_SPOOLERSTATUS = 0x002A;
    const long WM_DRAWITEM      = 0x002B;
    const long WM_MEASUREITEM   = 0x002C;
    const long WM_DELETEITEM    = 0x002D;
    const long WM_VKEYTOITEM    = 0x002E;
    const long WM_CHARTOITEM    = 0x002F;
    const long WM_SETFONT       = 0x0030;
    const long WM_GETFONT       = 0x0031;
	const long WM_SETHOTKEY     = 0x0032;
	const long WM_GETHOTKEY     = 0x0033;
    const long WM_QUERYDRAGICON = 0x0037;
    const long WM_COMPAREITEM   = 0x0039;

	const long WM_COMPACTING    = 0x0041;
	const long WM_COMMNOTIFY    = 0x0044; // no longer suported +
    const long WM_WINDOWPOSCHANGING = 0x0046;
    const long WM_WINDOWPOSCHANGED  = 0x0047;

    const long WM_POWER     = 0x0048;

        // wParam for WM_POWER window message and DRV_POWER driver notification
    const long PWR_OK             =   1;
    const long PWR_FAIL           =  -1;
    const long PWR_SUSPENDREQUEST =   1;
    const long PWR_SUSPENDRESUME  =   2;
    const long PWR_CRITICALRESUME =   3;

	const long WM_COPYDATA                     = 0x004A;
	const long WM_CANCELJOURNAL                = 0x004B;

    const long WM_NOTIFY               = 0x004E;
	const long WM_INPUTLANGCHANGEREQUEST       = 0x0050;
	const long WM_INPUTLANGCHANGE              = 0x0051;
	const long WM_TCARD                        = 0x0052;
	const long WM_HELP                         = 0x0053;
	const long WM_USERCHANGED                  = 0x0054;
	const long WM_NOTIFYFORMAT                 = 0x0055;

	const long NFR_ANSI                             = 1;
	const long NFR_UNICODE                          = 2;
	const long NF_QUERY                             = 3;
	const long NF_REQUERY                           = 4;

    const long WM_CONTEXTMENU          = 0x007B;
	const long WM_STYLECHANGING                = 0x007C;
	const long WM_STYLECHANGED                 = 0x007D;
	const long WM_DISPLAYCHANGE                = 0x007E;
	const long WM_GETICON                      = 0x007F;
	const long WM_SETICON                      = 0x0080;

    const long WM_NCCREATE      = 0x0081;
    const long WM_NCDESTROY     = 0x0082;
    const long WM_NCCALCSIZE    = 0x0083;
    const long WM_NCHITTEST     = 0x0084;
    const long WM_NCPAINT       = 0x0085;
    const long WM_NCACTIVATE    = 0x0086;
    const long WM_GETDLGCODE    = 0x0087;

    const long WM_NCMOUSEMOVE      = 0x00A0;
    const long WM_NCLBUTTONDOWN    = 0x00A1;
    const long WM_NCLBUTTONUP      = 0x00A2;
    const long WM_NCLBUTTONDBLCLK  = 0x00A3;
    const long WM_NCRBUTTONDOWN    = 0x00A4;
    const long WM_NCRBUTTONUP      = 0x00A5;
    const long WM_NCRBUTTONDBLCLK  = 0x00A6;
    const long WM_NCMBUTTONDOWN    = 0x00A7;
    const long WM_NCMBUTTONUP      = 0x00A8;
    const long WM_NCMBUTTONDBLCLK  = 0x00A9;

    const long WM_KEYFIRST      = 0x0100;
    const long WM_KEYDOWN       = 0x0100;
    const long WM_KEYUP         = 0x0101;
    const long WM_CHAR          = 0x0102;
    const long WM_DEADCHAR      = 0x0103;
    const long WM_SYSKEYDOWN    = 0x0104;
    const long WM_SYSKEYUP      = 0x0105;
    const long WM_SYSCHAR       = 0x0106;
    const long WM_SYSDEADCHAR   = 0x0107;
    const long WM_KEYLAST       = 0x0108;

	const long WM_IME_STARTCOMPOSITION         = 0x010D;
	const long WM_IME_ENDCOMPOSITION           = 0x010E;
	const long WM_IME_COMPOSITION              = 0x010F;
	const long WM_IME_KEYLAST                  = 0x010F;

    const long WM_INITDIALOG    = 0x0110;

    const long WM_COMMAND       = 0x0111;
    const long WM_SYSCOMMAND = 0x0112;
    const long WM_TIMER     = 0x0113;
	const long WM_SYSTIMER  = 0x0118;
    const long WM_HSCROLL       = 0x0114;
    const long WM_VSCROLL       = 0x0115;
    const long WM_INITMENU      = 0x0116;
    const long WM_INITMENUPOPUP = 0x0117;
    const long WM_MENUSELECT    = 0x011F;
    const long WM_MENUCHAR      = 0x0120;
    const long WM_ENTERIDLE     = 0x0121;

    const long WM_MENURBUTTONUP               = 0x0122;
    const long WM_MENUDRAG                    = 0x0123;
    const long WM_MENUGETOBJECT               = 0x0124;
    const long WM_UNINITMENUPOPUP             = 0x0125;
    const long WM_MENUCOMMAND                 = 0x0126;

	const long WM_CTLCOLORMSGBOX               = 0x0132;
	const long WM_CTLCOLOREDIT                 = 0x0133;
	const long WM_CTLCOLORLISTBOX              = 0x0134;
	const long WM_CTLCOLORBTN                  = 0x0135;
	const long WM_CTLCOLORDLG                  = 0x0136;
	const long WM_CTLCOLORSCROLLBAR            = 0x0137;
	const long WM_CTLCOLORSTATIC               = 0x0138;

    const long WM_MOUSEFIRST    = 0x0200;
    const long WM_MOUSEMOVE     = 0x0200;
    const long WM_LBUTTONDOWN   = 0x0201;
    const long WM_LBUTTONUP     = 0x0202;
    const long WM_LBUTTONDBLCLK = 0x0203;
    const long WM_RBUTTONDOWN   = 0x0204;
    const long WM_RBUTTONUP     = 0x0205;
    const long WM_RBUTTONDBLCLK = 0x0206;
    const long WM_MBUTTONDOWN   = 0x0207;
    const long WM_MBUTTONUP     = 0x0208;
    const long WM_MBUTTONDBLCLK = 0x0209;
	const long WM_MOUSEWHEEL    = 0x020A;
    const long WM_MOUSELAST     = 0x020A;

	const long WHEEL_DELTA      = 120;        // Value for rolling one detent 
	const long WHEEL_PAGESCROLL = 0xffffffff; // (UINT_MAX) Scroll one page 

    const long WM_PARENTNOTIFY  = 0x0210;
	const long MENULOOP_WINDOW                 = 0;
	const long MENULOOP_POPUP                  = 1;
	const long WM_ENTERMENULOOP                = 0x0211;
	const long WM_EXITMENULOOP                 = 0x0212;

	const long WM_NEXTMENU                     = 0x0213;

    const long WM_SIZING                       = 0x0214;
	const long WM_CAPTURECHANGED               = 0x0215;
	const long WM_MOVING                       = 0x0216;
	const long WM_POWERBROADCAST               = 0x0218;
	const long WM_DEVICECHANGE                 = 0x0219;

	const long WM_IME_SETCONTEXT               = 0x0281;
	const long WM_IME_NOTIFY                   = 0x0282;
	const long WM_IME_CONTROL                  = 0x0283;
	const long WM_IME_COMPOSITIONFULL          = 0x0284;
	const long WM_IME_SELECT                   = 0x0285;
	const long WM_IME_CHAR                     = 0x0286;
	const long WM_IME_KEYDOWN                  = 0x0290;
	const long WM_IME_KEYUP                    = 0x0291;

    // MDI messages
    const long WM_MDICREATE      = 0x0220;
    const long WM_MDIDESTROY     = 0x0221;
    const long WM_MDIACTIVATE    = 0x0222;
    const long WM_MDIRESTORE     = 0x0223;
    const long WM_MDINEXT        = 0x0224;
    const long WM_MDIMAXIMIZE    = 0x0225;
    const long WM_MDITILE        = 0x0226;
    const long WM_MDICASCADE     = 0x0227;
    const long WM_MDIICONARRANGE = 0x0228;
    const long WM_MDIGETACTIVE   = 0x0229;
    
	const long WM_MDISETMENU     = 0x0230;
	const long WM_ENTERSIZEMOVE                = 0x0231;
	const long WM_EXITSIZEMOVE                 = 0x0232;
    const long WM_DROPFILES     = 0x0233;
	const long WM_MDIREFRESHMENU               = 0x0234;

	const long WM_MOUSEHOVER                   = 0x02A1;
	const long WM_MOUSELEAVE                   = 0x02A3;

    // Clipboard viewer messages
    const long WM_CUT           = 0x0300;
    const long WM_COPY          = 0x0301;
    const long WM_PASTE         = 0x0302;
    const long WM_CLEAR         = 0x0303;
    const long WM_UNDO          = 0x0304;
    const long WM_RENDERFORMAT     = 0x0305;
    const long WM_RENDERALLFORMATS = 0x0306;
    const long WM_DESTROYCLIPBOARD = 0x0307;
    const long WM_DRAWCLIPBOARD    = 0x0308;
    const long WM_PAINTCLIPBOARD   = 0x0309;
    const long WM_VSCROLLCLIPBOARD = 0x030A;
    const long WM_SIZECLIPBOARD    = 0x030B;
    const long WM_ASKCBFORMATNAME  = 0x030C;
    const long WM_CHANGECBCHAIN    = 0x030D;
    const long WM_HSCROLLCLIPBOARD = 0x030E;
    const long WM_QUERYNEWPALETTE = 0x030F;
    const long WM_PALETTEISCHANGING = 0x0310;
    const long WM_PALETTECHANGED    = 0x0311;
	const long WM_HOTKEY                       = 0x0312;

	const long WM_PRINT                        = 0x0317;
    const long WM_PRINTCLIENT          = 0x0318;

	const long WM_HANDHELDFIRST                = 0x0358;
	const long WM_HANDHELDLAST                 = 0x035F;

	const long WM_AFXFIRST                     = 0x0360;
	const long WM_AFXLAST                      = 0x037F;

    const long WM_PENWINFIRST   = 0x0380;
    const long WM_PENWINLAST    = 0x038F;

	const long WM_APP                          = 0x8000;

	// wParam for WM_SIZING message  
	const long WMSZ_LEFT           = 1;
	const long WMSZ_RIGHT          = 2;
	const long WMSZ_TOP            = 3;
	const long WMSZ_TOPLEFT        = 4;
	const long WMSZ_TOPRIGHT       = 5;
	const long WMSZ_BOTTOM         = 6;
	const long WMSZ_BOTTOMLEFT     = 7;
	const long WMSZ_BOTTOMRIGHT    = 8;

	// WM_SYNCTASK Commands
	const long ST_BEGINSWP			= 0;
	const long ST_ENDSWP			= 1;
	const long ST_ADVISE			= 2;
	const long ST_BLOCKED			= 8;
	const long ST_BLOCKNEXT			= 128; //&H80
	const long ST_CLIENT			= 16; //&H10
	const long ST_CONNECTED			= 1;
	const long ST_DEFAULT			= 0;
	const long ST_INLIST			= 64; //&H40
	const long ST_ISLOCAL			= 4;
	const long ST_ISSELF			= 256; //&H100
	const long ST_KEEPUNDO			= 1;
	const long ST_SELECTION			= 2;
	const long ST_TERMINATED		= 32; //&H20


	// WM_NCHITTEST and MOUSEHOOKSTRUCT Mouse Position Codes
    const long HTERROR       = -2;
    const long HTTRANSPARENT = -1;
    const long HTNOWHERE     = 0;
    const long HTCLIENT      = 1;
    const long HTCAPTION     = 2;
    const long HTSYSMENU     = 3;
    const long HTGROWBOX     = 4; // HTSIZE;
    const long HTSIZE        = 4;
    const long HTMENU        = 5;
    const long HTHSCROLL     = 6;
    const long HTVSCROLL     = 7;
    const long HTMINBUTTON   = 8;
    const long HTMAXBUTTON   = 9;
    const long HTLEFT        = 10;
    const long HTRIGHT       = 11;
    const long HTTOP         = 12;
    const long HTTOPLEFT     = 13;
    const long HTTOPRIGHT    = 14;
    const long HTBOTTOM      = 15;
    const long HTBOTTOMLEFT  = 16;
    const long HTBOTTOMRIGHT = 17;
    const long HTBORDER      = 18;
    const long HTREDUCE      = 8; // HTMINBUTTON;
    const long HTZOOM        = 9; // HTMAXBUTTON;
	const long HTSIZEFIRST         = 10; // HTLEFT
	const long HTSIZELAST          = 17; // HTBOTTOMRIGHT
	const long HTOBJECT            = 19;
	const long HTCLOSE             = 20;
	const long HTHELP              = 21;


	// WM_ACTIVATE state values
    const long WA_INACTIVE      = 0;
    const long WA_ACTIVE       = 1;
    const long WA_CLICKACTIVE   =       2;

    // WM_MOUSEACTIVATE return codes
    const long MA_ACTIVATE      = 1;
    const long MA_ACTIVATEANDEAT =   2;
    const long MA_NOACTIVATE    = 3;
    const long MA_NOACTIVATEANDEAT = 4;

	// WM_SETICON / WM_GETICON Type Codes
	const long ICON_SMALL          = 0;
	const long ICON_BIG            = 1;

    // WM_SIZE message wParam values
    const long SIZE_RESTORED    = 0;
    const long SIZE_MINIMIZED   = 1;
    const long SIZE_MAXIMIZED   = 2;
    const long SIZE_MAXSHOW     = 3;
    const long SIZE_MAXHIDE     = 4;

    // 3D border styles
    const long BDR_RAISEDOUTER  = 0x0001;
    const long BDR_SUNKENOUTER  = 0x0002;
    const long BDR_RAISEDINNER  = 0x0004;
    const long BDR_SUNKENINNER  = 0x0008;

    const long BDR_OUTER        = 0x0003;
    const long BDR_INNER        = 0x000c;
    const long BDR_RAISED       = 0x0005;
    const long BDR_SUNKEN       = 0x000a;
	
    const long EDGE_RAISED      = 0x0005; // BDR_RAISEDOUTER | BDR_RAISEDINNER
    const long EDGE_SUNKEN      = 0x000a; // BDR_SUNKENOUTER | BDR_SUNKENINNER
    const long EDGE_ETCHED      = 0x0006; // BDR_SUNKENOUTER | BDR_RAISEDINNER
    const long EDGE_BUMP        = 0x0009; // BDR_RAISEDOUTER | BDR_SUNKENINNER

    // Border flags
    const long BF_LEFT          = 0x0001;
    const long BF_TOP           = 0x0002;
    const long BF_RIGHT         = 0x0004;
    const long BF_BOTTOM        = 0x0008;

    const long BF_TOPLEFT       = 0x0003; // BF_TOP | BF_LEFT
    const long BF_TOPRIGHT      = 0x0006; // BF_TOP | BF_RIGHT
    const long BF_BOTTOMLEFT    = 0x0009; // BF_BOTTOM | BF_LEFT
    const long BF_BOTTOMRIGHT   = 0x000c; // BF_BOTTOM | BF_RIGHT
    const long BF_RECT          = 0x000f; // BF_LEFT | BF_TOP | BF_RIGHT | BF_BOTTOM

    // Flags for DrawFrameControl
    const long DFC_CAPTION          = 1;
    const long DFC_MENU             = 2;
    const long DFC_SCROLL           = 3;
    const long DFC_BUTTON           = 4;

    const long DFCS_CAPTIONCLOSE    = 0x0000;
    const long DFCS_CAPTIONMIN      = 0x0001;
    const long DFCS_CAPTIONMAX      = 0x0002;
    const long DFCS_CAPTIONRESTORE  = 0x0003;
    const long DFCS_CAPTIONHELP     = 0x0004;

    const long DFCS_MENUARROW       = 0x0000;
    const long DFCS_MENUCHECK       = 0x0001;
    const long DFCS_MENUBULLET      = 0x0002;
    const long DFCS_MENUARROWRIGHT  = 0x0004;
	
    const long DFCS_SCROLLUP        = 0x0000;
    const long DFCS_SCROLLDOWN      = 0x0001;
    const long DFCS_SCROLLLEFT      = 0x0002;
    const long DFCS_SCROLLRIGHT     = 0x0003;
    const long DFCS_SCROLLCOMBOBOX  = 0x0005;
    const long DFCS_SCROLLSIZEGRIP  = 0x0008;
    const long DFCS_SCROLLSIZEGRIPRIGHT = 0x0010;

    const long DFCS_BUTTONCHECK     = 0x0000;
    const long DFCS_BUTTONRADIOIMAGE =  0x0001;
    const long DFCS_BUTTONRADIOMASK = 0x0002;
    const long DFCS_BUTTONRADIO     = 0x0004;
    const long DFCS_BUTTON3STATE    = 0x0008;
    const long DFCS_BUTTONPUSH      = 0x0010;

    const long DFCS_INACTIVE        = 0x0100;
    const long DFCS_PUSHED          = 0x0200;
    const long DFCS_CHECKED         = 0x0400;
    const long DFCS_ADJUSTRECT      = 0x2000;
    const long DFCS_FLAT            = 0x4000;
    const long DFCS_MONO            = 0x8000;

    // Flags for DrawCaption
    const long DC_ACTIVE        = 0x0001;
    const long DC_SMALLCAP      = 0x0002;
    const long DC_ICON          = 0x0004;
    const long DC_TEXT          = 0x0008;
    const long DC_INBUTTON      = 0x0010;
    const long DC_CAPTION       = 0x001C;   // (DC_ICON | DC_TEXT | DC_BUTTONS)
    const long DC_NC            = 0x001C;   // (DC_CAPTION | DC_FRAME)

	// Draw Animated Rects flasgs
    const long IDANI_OPEN       = 1;
    const long IDANI_CLOSE      = 2;
    const long IDANI_CAPTION    = 3;

	// Predefined Clipboard Formats
    const long CF_TEXT          = 1;
    const long CF_BITMAP        = 2;
    const long CF_METAFILEPICT  = 3;
    const long CF_SYLK          = 4;
    const long CF_DIF           = 5;
    const long CF_TIFF          = 6;
    const long CF_OEMTEXT       = 7;
    const long CF_DIB           = 8;
    const long CF_PALETTE       = 9;
    const long CF_PENDATA       = 10;
    const long CF_RIFF          = 11;
    const long CF_WAVE          = 12;
	const long CF_UNICODETEXT      = 13;
	const long CF_ENHMETAFILE      = 14;
	const long CF_HDROP            = 15;
	const long CF_LOCALE           = 16;
	const long CF_MAX              = 17;
	
    const long CF_OWNERDISPLAY  = 0x0080;
    const long CF_DSPTEXT      = 0x0081;
    const long CF_DSPBITMAP        = 0x0082;
    const long CF_DSPMETAFILEPICT  = 0x0083;

    // Private formats don't get GlobalFree()'d
    const long CF_PRIVATEFIRST  = 0x0200;
    const long CF_PRIVATELAST      = 0x02FF;

    // GDIOBJ formats do get DeleteObject()'d
    const long CF_GDIOBJFIRST      = 0x0300;
    const long CF_GDIOBJLAST       = 0x03FF;


	// SendMessageTimeout values
	const long SMTO_NORMAL         = 0x0000;
	const long SMTO_BLOCK          = 0x0001;
	const long SMTO_ABORTIFHUNG    = 0x0002;

	// Broadcast Special Message Recipient list
	const long BSM_ALLCOMPONENTS       = 0x00000000;
	const long BSM_VXDS                = 0x00000001;
	const long BSM_NETDRIVER           = 0x00000002;
	const long BSM_INSTALLABLEDRIVERS  = 0x00000004;
	const long BSM_APPLICATIONS        = 0x00000008;
	const long BSM_ALLDESKTOPS         = 0x00000010;

	// Broadcast Special Message Flags
	const long BSF_QUERY               = 0x00000001;
	const long BSF_IGNORECURRENTTASK   = 0x00000002;
	const long BSF_FLUSHDISK           = 0x00000004;
	const long BSF_NOHANG              = 0x00000008;
	const long BSF_POSTMESSAGE         = 0x00000010;
	const long BSF_FORCEIFHUNG         = 0x00000020;
	const long BSF_NOTIMEOUTIFNOTHUNG  = 0x00000040;

	const long BROADCAST_QUERY_DENY    = 0x424D5144;  // Return this value to deny a query.

	/*
	* PeekMessage() Options
	*/
	const long PM_NOREMOVE         = 0x0000;
	const long PM_REMOVE           = 0x0001;
	const long PM_NOYIELD          = 0x0002;

	const long MOD_ALT         = 0x0001;
	const long MOD_CONTROL     = 0x0002;
	const long MOD_SHIFT       = 0x0004;
	const long MOD_WIN         = 0x0008;

	const long IDHOT_SNAPWINDOW        = -1;    // SHIFT-PRINTSCRN  
	const long IDHOT_SNAPDESKTOP       = -2;    // PRINTSCRN

    const long EW_RESTARTWINDOWS = 0x42;
    const long EW_REBOOTSYSTEM   = 0x43;
    const long EW_EXITANDEXECAPP = 0x0044;

    const long EWX_LOGOFF   = 0x0;
    const long EWX_SHUTDOWN = 0x1;
    const long EWX_REBOOT   = 0x2;
    const long EWX_FORCE    = 0x4;
    const long EWX_POWEROFF = 0x8;

    // ******** Window Styles *********

    const long WS_OVERLAPPED    = 0x00000000;
    const long WS_POPUP         = 0x80000000;
    const long WS_CHILD         = 0x40000000;
    const long WS_MINIMIZE      = 0x20000000;
    const long WS_VISIBLE       = 0x10000000;
    const long WS_DISABLED      = 0x08000000;
    const long WS_CLIPSIBLINGS  = 0x04000000;
    const long WS_CLIPCHILDREN  = 0x02000000;
    const long WS_MAXIMIZE      = 0x01000000;
    const long WS_CAPTION       = 0x00C00000;   // WS_BORDER | WS_DLGFRAME
    const long WS_BORDER        = 0x00800000;
    const long WS_DLGFRAME      = 0x00400000;
    const long WS_VSCROLL       = 0x00200000;
    const long WS_HSCROLL       = 0x00100000;
    const long WS_SYSMENU       = 0x00080000;
    const long WS_THICKFRAME    = 0x00040000;
    const long WS_GROUP         = 0x00020000;
    const long WS_TABSTOP       = 0x00010000;

    const long WS_MINIMIZEBOX   = 0x00020000;
    const long WS_MAXIMIZEBOX   = 0x00010000;


    // Obsolete style names
    const long WS_TILED         = 0x00000000;
    const long WS_ICONIC        = 0x20000000;
    const long WS_SIZEBOX       = 0x00040000;
    const long WS_TILEDWINDOW   = 0x00CF0000;
    
    // Common Window Styles
    const long WS_OVERLAPPEDWINDOW = 0x00CF0000;
    const long WS_POPUPWINDOW   = 0x80880000;
    const long WS_CHILDWINDOW   = 0x40000000;

    // Extended Window Styles
    const long WS_EX_DLGMODALFRAME  = 0x00000001;
    const long WS_EX_NOPARENTNOTIFY = 0x00000004;
    const long WS_EX_TOPMOST        = 0x00000008;
    const long WS_EX_ACCEPTFILES    = 0x00000010;
    const long WS_EX_TRANSPARENT    = 0x00000020;

    const long WS_EX_MDICHILD       = 0x00000040;
    const long WS_EX_TOOLWINDOW     = 0x00000080;
    const long WS_EX_WINDOWEDGE     = 0x00000100;
    const long WS_EX_CLIENTEDGE     = 0x00000200;
    const long WS_EX_CONTEXTHELP    = 0x00000400;

    const long WS_EX_RIGHT          = 0x00001000;
    const long WS_EX_LEFT           = 0x00000000;
    const long WS_EX_RTLREADING     = 0x00002000;
    const long WS_EX_LTRREADING     = 0x00000000;
    const long WS_EX_LEFTSCROLLBAR  = 0x00004000;
    const long WS_EX_RIGHTSCROLLBAR = 0x00000000;

    const long WS_EX_CONTROLPARENT  = 0x00010000;
    const long WS_EX_STATICEDGE     = 0x00020000;
    const long WS_EX_APPWINDOW      = 0x00040000;


    const long WS_EX_OVERLAPPEDWINDOW = 0x00000300; // (WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE)
    const long WS_EX_PALETTEWINDOW  = 0x00000188; // (WS_EX_WINDOWEDGE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST)


    // Class styles
    const long CS_VREDRAW      = 0x0001;
    const long CS_HREDRAW      = 0x0002;
    const long CS_KEYCVTWINDOW    = 0x0004;
    const long CS_DBLCLKS      = 0x0008;
    const long CS_OWNDC        = 0x0020;
    const long CS_CLASSDC      = 0x0040;
    const long CS_PARENTDC     = 0x0080;
    const long CS_NOKEYCVT      = 0x0100;
    const long CS_NOCLOSE      = 0x0200;
    const long CS_SAVEBITS     = 0x0800;
    const long CS_BYTEALIGNCLIENT  = 0x1000;
    const long CS_BYTEALIGNWINDOW  = 0x2000;
    const long CS_GLOBALCLASS      = 0x4000;
	const long CS_PUBLICCLASS	   = 0x4000;
	const long CS_IME              = 0x00010000;
	
    // Window field offsets for GetWindowLong() and GetWindowWord()
    const long GWL_WNDPROC       = -4;
    const long GWL_HINSTANCE     = -6;
    const long GWL_HWNDPARENT    = -8;
    const long GWL_ID            = -12;
    const long GWL_STYLE         = -16;
    const long GWL_EXSTYLE       = -20;
    const long GWL_USERDATA      = -21;
    const long GWD_HINSTANCE     = -6;
    const long GWD_HWNDPARENT    = -8;
    const long GWD_ID            = -12;

    // Class field offsets for GetClassLong() and GetClassWord()
    const long GCL_MENUNAME      = -8;
    const long GCL_HBRBACKGROUND = -10;
    const long GCL_HCURSOR       = -12;
    const long GCL_HICON         = -14;
    const long GCL_HMODULE       = -16;
    const long GCL_CBWNDEXTRA    = -18;
    const long GCL_CBCLSEXTRA    = -20;
    const long GCL_WNDPROC       = -24;
    const long GCL_STYLE         = -26;
    const long GCW_ATOM          = -32;
    const long GCD_HBRBACKGROUND = -10;
    const long GCD_HCURSOR       = -12;
    const long GCD_HICON         = -14;
    const long GCD_HMODULE       = -16;
    const long GCD_CBWNDEXTRA    = -18;
    const long GCD_CBCLSEXTRA    = -20;
    const long GCD_STYLE         = -26;


	// *********** Show Windows Consts *********

    const long SW_HIDE               = 0;
    const long SW_SHOWNORMAL         = 1;
    const long SW_NORMAL             = 1;
    const long SW_SHOWMINIMIZED      = 2;
    const long SW_SHOWMAXIMIZED      = 3;
    const long SW_MAXIMIZE           = 3;
    const long SW_SHOWNOACTIVATE     = 4;
    const long SW_SHOW               = 5;
    const long SW_MINIMIZE           = 6;
    const long SW_SHOWMINNOACTIVE    = 7;
    const long SW_SHOWNA             = 8;
    const long SW_RESTORE            = 9;
    const long SW_SHOWDEFAULT        = 10;
    const long SW_MAX                = 10;

    const long FLASHW_STOP        = 0;
    const long FLASHW_CAPTION     = 0x00000001;
    const long FLASHW_TRAY        = 0x00000002;
    const long FLASHW_ALL         = 0x00000003; //FLASHW_CAPTION | FLASHW_TRAY
    const long FLASHW_TIMER       = 0x00000004;
    const long FLASHW_TIMERNOFG   = 0x0000000C;

    // ********** SetWindowPos() and WINDOWPOS flags **********
    const long SWP_NOSIZE      = 0x0001;
    const long SWP_NOMOVE      = 0x0002;
    const long SWP_NOZORDER        = 0x0004;
    const long SWP_NOREDRAW        = 0x0008;
    const long SWP_NOACTIVATE      = 0x0010;
    const long SWP_FRAMECHANGED    = 0x0020; // The frame changed: send WM_NCCALCSIZE
    const long SWP_SHOWWINDOW      = 0x0040;
    const long SWP_HIDEWINDOW      = 0x0080;
    const long SWP_NOCOPYBITS      = 0x0100;
    const long SWP_NOOWNERZORDER   = 0x0200; // Don't do owner Z ordering
    const long SWP_NOSENDCHANGING  = 0x0400;

    const long SWP_DRAWFRAME       = 0x0020;
    const long SWP_NOREPOSITION    = 0x0200;

    const long SWP_DEFERERASE      = 0x2000;
    const long SWP_ASYNCWINDOWPOS  = 0x4000;

    // SetWindowPos() hwndInsertAfter field values
    const long HWND_TOP            = 0;
    const long HWND_BOTTOM         = 1;
    const long HWND_TOPMOST        = -1;
    const long HWND_NOTOPMOST      = -2;

	const long WPF_SETMINPOSITION      = 0x0001;
    const long WPF_RESTORETOMAXIMIZED  = 0x0002;

	const long EMPTY_BYTE = 0x45;
	const long KEYEVENTF_EXTENDEDKEY = 0x0001;
	const long KEYEVENTF_KEYUP       = 0x0002;

	const long MOUSEEVENTF_MOVE        = 0x0001; // mouse move 
	const long MOUSEEVENTF_LEFTDOWN    = 0x0002; // left button down 
	const long MOUSEEVENTF_LEFTUP      = 0x0004; // left button up 
	const long MOUSEEVENTF_RIGHTDOWN   = 0x0008; // right button down 
	const long MOUSEEVENTF_RIGHTUP     = 0x0010; // right button up 
	const long MOUSEEVENTF_MIDDLEDOWN  = 0x0020; // middle button down 
	const long MOUSEEVENTF_MIDDLEUP    = 0x0040; // middle button up 
	const long MOUSEEVENTF_WHEEL       = 0x0800; // wheel button rolled 
	const long MOUSEEVENTF_ABSOLUTE    = 0x8000; // absolute move 

    // GetSystemMetrics() codes
    const long SM_CXSCREEN     = 0;
    const long SM_CYSCREEN     = 1;
    const long SM_CXVSCROLL        = 2;
    const long SM_CYHSCROLL        = 3;
    const long SM_CYCAPTION        = 4;
    const long SM_CXBORDER     = 5;
    const long SM_CYBORDER     = 6;
    const long SM_CXDLGFRAME   = 7;
    const long SM_CYDLGFRAME   = 8;
    const long SM_CYVTHUMB     = 9;
    const long SM_CXHTHUMB     = 10;
    const long SM_CXICON       = 11;
    const long SM_CYICON       = 12;
    const long SM_CXCURSOR     = 13;
    const long SM_CYCURSOR     = 14;
    const long SM_CYMENU           = 15;
    const long SM_CXFULLSCREEN = 16;
    const long SM_CYFULLSCREEN = 17;
    const long SM_CYKANJIWINDOW    = 18;
    const long SM_MOUSEPRESENT = 19;
    const long SM_CYVSCROLL        = 20;
    const long SM_CXHSCROLL        = 21;
    const long SM_DEBUG            = 22;
    const long SM_SWAPBUTTON   = 23;
    const long SM_RESERVED1        = 24;
    const long SM_RESERVED2        = 25;
    const long SM_RESERVED3        = 26;
    const long SM_RESERVED4        = 27;
    const long SM_CXMIN            = 28;
    const long SM_CYMIN            = 29;
    const long SM_CXSIZE           = 30;
    const long SM_CYSIZE           = 31;
    const long SM_CXFRAME      = 32;
    const long SM_CYFRAME      = 33;
    const long SM_CXMINTRACK   = 34;
    const long SM_CYMINTRACK   = 35;

    const long SM_CXDOUBLECLK  = 36;
    const long SM_CYDOUBLECLK  = 37;
    const long SM_CXICONSPACING    = 38;
    const long SM_CYICONSPACING    = 39;
    const long SM_MENUDROPALIGNMENT = 40;
    const long SM_PENWINDOWS       = 41;
    const long SM_DBCSENABLED  = 42;

    const long SM_CMOUSEBUTTONS = 43;

	// win40 name changes 
    const long SM_CXFIXEDFRAME   = 7; //same as SM_CXFRAME
    const long SM_CYFIXEDFRAME   = 8; //Same as SM_CYFRAME
    const long SM_CXSIZEFRAME      = 32;
    const long SM_CYSIZEFRAME      = 33;

	const long SM_SECURE               = 44;
	const long SM_CXEDGE               = 45;
	const long SM_CYEDGE               = 46;
	const long SM_CXMINSPACING         = 47;
	const long SM_CYMINSPACING         = 48;
	const long SM_CXSMICON             = 49;
	const long SM_CYSMICON             = 50;
	const long SM_CYSMCAPTION          = 51;
	const long SM_CXSMSIZE             = 52;
	const long SM_CYSMSIZE             = 53;
	const long SM_CXMENUSIZE           = 54;
	const long SM_CYMENUSIZE           = 55;
	const long SM_ARRANGE              = 56;
	const long SM_CXMINIMIZED          = 57;
	const long SM_CYMINIMIZED          = 58;
	const long SM_CXMAXTRACK           = 59;
	const long SM_CYMAXTRACK           = 60;
	const long SM_CXMAXIMIZED          = 61;
	const long SM_CYMAXIMIZED          = 62;
	const long SM_NETWORK              = 63;
	const long SM_CLEANBOOT            = 67;
	const long SM_CXDRAG               = 68;
	const long SM_CYDRAG               = 69;
	const long SM_SHOWSOUNDS           = 70;
	const long SM_CXMENUCHECK          = 71;   // Use instead of GetMenuCheckMarkDimensions()! 
	const long SM_CYMENUCHECK          = 72;
	const long SM_SLOWMACHINE          = 73;
	const long SM_MIDEASTENABLED       = 74;
	const long SM_MOUSEWHEELPRESENT    = 75;
	const long SM_CMETRICS             = 76;

    // System Menu Command Values
    const long SC_SIZE      = 0xF000;
    const long SC_MOVE      = 0xF010;
    const long SC_MINIMIZE  = 0xF020;
    const long SC_MAXIMIZE  = 0xF030;
    const long SC_NEXTWINDOW = 0xF040;
    const long SC_PREVWINDOW = 0xF050;
    const long SC_CLOSE     = 0xF060;
    const long SC_VSCROLL   = 0xF070;
    const long SC_HSCROLL   = 0xF080;
    const long SC_MOUSEMENU = 0xF090;
    const long SC_KEYMENU   = 0xF100;
    const long SC_ARRANGE   = 0xF110;
    const long SC_RESTORE   = 0xF120;
    const long SC_TASKLIST  = 0xF130;
    const long SC_SCREENSAVE = 0xF140;
    const long SC_HOTKEY    = 0xF150;
    const long SC_DEFAULT      = 0xF160;
    const long SC_MONITORPOWER = 0xF170;
    const long SC_CONTEXTHELP  = 0xF180;
    const long SC_SEPARATOR    = 0xF00F;

	// Obsolete names
 	//define SC_ICON         SC_MINIMIZE
	//define SC_ZOOM         SC_MAXIMIZE

/* ;win40  -- A lot of MF_* flags have been renamed as MFT_* and MFS_* flags */
/*
 * Menu flags for Add/Check/EnableMenuItem()
 */
	const long MF_INSERT           = 0x00000000;
	const long MF_CHANGE           = 0x00000080;
	const long MF_APPEND           = 0x00000100;
	const long MF_DELETE           = 0x00000200;
	const long MF_REMOVE           = 0x00001000;

	const long MF_BYCOMMAND        = 0x00000000;
	const long MF_BYPOSITION       = 0x00000400;

	const long MF_SEPARATOR        = 0x00000800;

	const long MF_ENABLED          = 0x00000000;
	const long MF_GRAYED           = 0x00000001;
	const long MF_DISABLED         = 0x00000002;

	const long MF_UNCHECKED        = 0x00000000;
	const long MF_CHECKED          = 0x00000008;
	const long MF_USECHECKBITMAPS  = 0x00000200;

	const long MF_STRING           = 0x00000000;
	const long MF_BITMAP           = 0x00000004;
	const long MF_OWNERDRAW        = 0x00000100;

	const long MF_POPUP            = 0x00000010;
	const long MF_MENUBARBREAK     = 0x00000020;
	const long MF_MENUBREAK        = 0x00000040;

	const long MF_UNHILITE         = 0x00000000;
	const long MF_HILITE           = 0x00000080;

	const long MF_DEFAULT          = 0x00001000;
	const long MF_SYSMENU          = 0x00002000;
	const long MF_HELP             = 0x00004000;
	const long MF_RIGHTJUSTIFY     = 0x00004000;

	const long MF_MOUSESELECT      = 0x00008000;
	const long MF_END              = 0x00000080;  /* Obsolete -- only used by old RES files */


	const long MFT_STRING          = 0x00000000;
	const long MFT_BITMAP          = 0x00000004;
	const long MFT_MENUBARBREAK    = 0x00000020;
	const long MFT_MENUBREAK       = 0x00000040;
	const long MFT_OWNERDRAW       = 0x00000100;
	const long MFT_RADIOCHECK      = 0x00000200;
	const long MFT_SEPARATOR       = 0x00000800;
	const long MFT_RIGHTORDER      = 0x00002000;
	const long MFT_RIGHTJUSTIFY    = 0x00004000;

/* Menu flags for Add/Check/EnableMenuItem() */
	const long MFS_GRAYED          = 0x00000003;
	const long MFS_DISABLED        = 0x00000003;
	const long MFS_CHECKED         = 0x00000008;
	const long MFS_HILITE          = 0x00000080;
	const long MFS_ENABLED         = 0x00000000;
	const long MFS_UNCHECKED       = 0x00000000;
	const long MFS_UNHILITE        = 0x00000000;
	const long MFS_DEFAULT         = 0x00001000;
	//If WINVER >= 0x0500
	const long MFS_MASK            = 0x0000108B;
	const long MFS_HOTTRACKDRAWN   = 0x10000000;
	const long MFS_CACHEDBMP       = 0x20000000;
	const long MFS_BOTTOMGAPDROP   = 0x40000000;
	const long MFS_TOPGAPDROP      = 0x80000000;
	const long MFS_GAPDROP         = 0xC0000000;
	//End if WINVER >= 0x0500

	const long MNC_IGNORE  = 0;
	const long MNC_CLOSE   = 1;
	const long MNC_EXECUTE = 2;
	const long MNC_SELECT  = 3;

	const long MIIM_STATE       = 0x00000001;
	const long MIIM_ID          = 0x00000002;
	const long MIIM_SUBMENU     = 0x00000004;
	const long MIIM_CHECKMARKS  = 0x00000008;
	const long MIIM_TYPE        = 0x00000010;
	const long MIIM_DATA        = 0x00000020;
	const long MIIM_STRING      = 0x00000040;
	const long MIIM_BITMAP      = 0x00000080;
	const long MIIM_FTYPE       = 0x00000100;

	const long MNS_NOCHECK        = 0x80000000;
	const long MNS_MODELESS       = 0x40000000;
	const long MNS_DRAGDROP       = 0x20000000;
	const long MNS_AUTODISMISS    = 0x10000000;
	const long MNS_NOTIFYBYPOS    = 0x08000000;
	const long MNS_CHECKORBMP     = 0x04000000;

	const long MIM_MAXHEIGHT              = 0x00000001;
	const long MIM_BACKGROUND             = 0x00000002;
	const long MIM_HELPID                 = 0x00000004;
	const long MIM_MENUDATA               = 0x00000008;
	const long MIM_STYLE                  = 0x00000010;
	const long MIM_APPLYTOSUBMENUS        = 0x80000000;

    // Flags for TrackPopupMenu
    const long TPM_LEFTBUTTON   = 0x0000;
    const long TPM_RIGHTBUTTON  = 0x0002;
    const long TPM_LEFTALIGN    = 0x0000;
    const long TPM_CENTERALIGN  = 0x0004;
    const long TPM_RIGHTALIGN   = 0x0008;

    const long TPM_TOPALIGN     = 0x0000;
    const long TPM_VCENTERALIGN = 0x0010;
    const long TPM_BOTTOMALIGN  = 0x0020;

    const long TPM_HORIZONTAL   = 0x0000;     // Horz alignment matters more
    const long TPM_VERTICAL     = 0x0040;     // Vert alignment matters more
    const long TPM_NONOTIFY     = 0x0080;     // Don't send any notification msgs
    const long TPM_RETURNCMD    = 0x0100;


	// ********** Drag-and-drop support *********

	const long DOF_EXECUTABLE      = 0x8001;
	const long DOF_DOCUMENT        = 0x8002;
	const long DOF_DIRECTORY       = 0x8003;
	const long DOF_MULTIPLE        = 0x8004;
	const long DOF_PROGMAN         = 0x0001;
	const long DOF_SHELLDATA       = 0x0002;

	const long DO_DROPFILE         = 0x454C4946;
	const long DO_PRINTFILE        = 0x544E5250;

    // DrawText() Format Flags
    const long DT_TOP				= 0x0000;
    const long DT_LEFT				= 0x0000;
    const long DT_CENTER			= 0x0001;
    const long DT_RIGHT				= 0x0002;
    const long DT_VCENTER			= 0x0004;
    const long DT_BOTTOM			= 0x0008;
    const long DT_WORDBREAK			= 0x0010;
    const long DT_SINGLELINE		= 0x0020;
    const long DT_EXPANDTABS		= 0x0040;
    const long DT_TABSTOP			= 0x0080;
    const long DT_NOCLIP			= 0x0100;
    const long DT_EXTERNALLEADING	= 0x0200;
    const long DT_CALCRECT			= 0x0400;
    const long DT_NOPREFIX			= 0x0800;
    const long DT_INTERNAL			= 0x1000;
	
	const long DT_EDITCONTROL      = 0x00002000;
	const long DT_PATH_ELLIPSIS    = 0x00004000;
	const long DT_END_ELLIPSIS     = 0x00008000;
	const long DT_MODIFYSTRING     = 0x00010000;
	const long DT_RTLREADING       = 0x00020000;
	const long DT_WORD_ELLIPSIS    = 0x00040000;

	// Image type 
	const long DST_COMPLEX     = 0x0000;
	const long DST_TEXT        = 0x0001;
	const long DST_PREFIXTEXT  = 0x0002;
	const long DST_ICON        = 0x0003;
	const long DST_BITMAP      = 0x0004;

	// State type 
	const long DSS_NORMAL      = 0x0000;
	const long DSS_UNION       = 0x0010;  // Gray string appearance 
	const long DSS_DISABLED    = 0x0020;
	const long DSS_MONO        = 0x0080;
	const long DSS_RIGHT       = 0x8000;

	[helpstring("ScrollWindowEx: Scroll children within *lprcScroll.")]
    const long SW_SCROLLCHILDREN   = 0x0001; 
	[helpstring("ScrollWindowEx: Invalidate after scrolling")]
    const long SW_INVALIDATE       = 0x0002; 
	[helpstring("ScrollWindowEx: If SW_INVALIDATE, don't send WM_ERASEBACKGROUND")]
    const long SW_ERASE            = 0x0004; 
	[helpstring("ScrollWindowEx: Use smooth scrolling")]
	const long SW_SMOOTHSCROLL	   = 0x0010; 


    // ****** Scroll bar support ***

    // Scroll bar selection constants
    const long SB_HORZ         = 0;
    const long SB_VERT         = 1;
    const long SB_CTL          = 2;
    const long SB_BOTH         = 3;

    // WM_H/VSCROLL commands
    const long SB_LINEUP        = 0;
    const long SB_LINELEFT      = 0;
    const long SB_LINEDOWN      = 1;
    const long SB_LINERIGHT     = 1;
    const long SB_PAGEUP        = 2;
    const long SB_PAGELEFT      = 2;
    const long SB_PAGEDOWN      = 3;
    const long SB_PAGERIGHT     = 3;
    const long SB_THUMBPOSITION = 4;
    const long SB_THUMBTRACK    = 5;
    const long SB_TOP           = 6;
    const long SB_LEFT          = 6;
    const long SB_BOTTOM        = 7;
    const long SB_RIGHT         = 7;
    const long SB_ENDSCROLL     = 8;

	//Scroll bar messages
    const long SBM_SETPOS                 = 0x00E0; /*not in win3.1 */
    const long SBM_GETPOS                 = 0x00E1; /*not in win3.1 */
    const long SBM_SETRANGE               = 0x00E2; /*not in win3.1 */
    const long SBM_SETRANGEREDRAW         = 0x00E6; /*not in win3.1 */
    const long SBM_GETRANGE               = 0x00E3; /*not in win3.1 */
    const long SBM_ENABLE_ARROWS          = 0x00E4; /*not in win3.1 */
    const long SBM_SETSCROLLINFO          = 0x00E9;
    const long SBM_GETSCROLLINFO          = 0x00EA;

    const long SIF_RANGE           = 0x0001;
    const long SIF_PAGE            = 0x0002;
    const long SIF_POS             = 0x0004;
    const long SIF_DISABLENOSCROLL = 0x0008;
    const long SIF_TRACKPOS        = 0x0010;
//#define SIF_ALL             (SIF_RANGE | SIF_PAGE | SIF_POS | SIF_TRACKPOS)

	// MessageBox() Flags
    const long MB_OK                 = 0x0000;
    const long MB_OKCANCEL           = 0x0001;
    const long MB_ABORTRETRYIGNORE   = 0x0002;
    const long MB_YESNOCANCEL        = 0x0003;
    const long MB_YESNO              = 0x0004;
    const long MB_RETRYCANCEL        = 0x0005;

    const long MB_ICONHAND           = 0x0010;
    const long MB_ICONQUESTION       = 0x0020;
    const long MB_ICONEXCLAMATION    = 0x0030;
    const long MB_ICONASTERISK       = 0x0040;

	const long MB_USERICON           = 0x00000080;
	const long MB_ICONWARNING        = 0x0030; // MB_ICONEXCLAMATION
	const long MB_ICONERROR          = 0x0010; // MB_ICONHAND

    const long MB_ICONINFORMATION    = 0x0040;
    const long MB_ICONSTOP           = 0x0010;

    const long MB_DEFBUTTON1         = 0x0000;
    const long MB_DEFBUTTON2         = 0x0100;
    const long MB_DEFBUTTON3         = 0x0200;
    const long MB_DEFBUTTON4         = 0x0300;

    const long MB_APPLMODAL          = 0x0000;
    const long MB_SYSTEMMODAL        = 0x1000;
    const long MB_TASKMODAL          = 0x2000;
    const long MB_HELP               = 0x4000;	// Help Button

    const long MB_NOFOCUS            = 0x8000;
	const long MB_SETFOREGROUND            = 0x00010000;
	const long MB_DEFAULT_DESKTOP_ONLY     = 0x00020000;

	const long MB_TOPMOST            = 0x00040000;
	const long MB_RIGHT              = 0x00080000;
	const long MB_RTLREADING         = 0x00100000;

	const long MB_SERVICE_NOTIFICATION     = 0x00200000;

    const long MB_TYPEMASK           = 0x000F;
    const long MB_ICONMASK           = 0x00F0;
    const long MB_DEFMASK            = 0x0F00;
	const long MB_MODEMASK           = 0x00003000;
	const long MB_MISCMASK           = 0x0000C000;
	[helpstring("ChildWindowFromPointEx: flag")]
	const long CWP_ALL             = 0x0000;
	[helpstring("ChildWindowFromPointEx: flag")]
	const long CWP_SKIPINVISIBLE   = 0x0001;
	[helpstring("ChildWindowFromPointEx: flag")]
	const long CWP_SKIPDISABLED    = 0x0002;
	[helpstring("ChildWindowFromPointEx: flag")]
	const long CWP_SKIPTRANSPARENT = 0x0004;

    // Color types 
    const long CTLCOLOR_MSGBOX			= 0;
    const long CTLCOLOR_EDIT			= 1;
    const long CTLCOLOR_LISTBOX			= 2;
    const long CTLCOLOR_BTN				= 3;
    const long CTLCOLOR_DLG				= 4;
    const long CTLCOLOR_SCROLLBAR		= 5;
    const long CTLCOLOR_STATIC			= 6;
    const long CTLCOLOR_MAX				= 7;

    // GetSysColor/SetSysColors
    const long COLOR_SCROLLBAR             = 0;
    const long COLOR_BACKGROUND            = 1;
    const long COLOR_ACTIVECAPTION         = 2;
    const long COLOR_INACTIVECAPTION       = 3;
    const long COLOR_MENU                  = 4;
    const long COLOR_WINDOW                = 5;
    const long COLOR_WINDOWFRAME           = 6;
    const long COLOR_MENUTEXT              = 7;
    const long COLOR_WINDOWTEXT            = 8;
    const long COLOR_CAPTIONTEXT           = 9;
    const long COLOR_ACTIVEBORDER          = 10;
    const long COLOR_INACTIVEBORDER        = 11;
    const long COLOR_APPWORKSPACE          = 12;
    const long COLOR_HIGHLIGHT             = 13;
    const long COLOR_HIGHLIGHTTEXT         = 14;
    const long COLOR_BTNFACE               = 15;
    const long COLOR_BTNSHADOW             = 16;
    const long COLOR_GRAYTEXT              = 17;
    const long COLOR_BTNTEXT               = 18;
    const long COLOR_INACTIVECAPTIONTEXT   = 19;
    const long COLOR_BTNHIGHLIGHT          = 20;

    const long COLOR_3DDKSHADOW            = 21;
    const long COLOR_3DLIGHT               = 22;
    const long COLOR_INFOTEXT              = 23;
    const long COLOR_INFOBK                = 24;

    const long COLOR_DESKTOP               = 1;  // COLOR_BACKGROUND
    const long COLOR_3DFACE                = 15; // COLOR_BTNFACE
    const long COLOR_3DSHADOW              = 16; // COLOR_BTNSHADOW
    const long COLOR_3DHIGHLIGHT           = 20; // COLOR_BTNHIGHLIGHT
    const long COLOR_3DHILIGHT             = 20; // COLOR_BTNHIGHLIGHT
    const long COLOR_BTNHILIGHT            = 20; // COLOR_BTNHIGHLIGHT

	//The "real" ancestor window
	const long     GA_MIC         = 1;
	const long     GA_PARENT      = 1;
	const long     GA_ROOT        = 2;
	const long     GA_ROOTOWNER   = 3;
	const long     GA_MAC         = 4;

    // GetWindow() constants
    const long GW_HWNDFIRST  = 0;
    const long GW_HWNDLAST   = 1;
    const long GW_HWNDNEXT   = 2;
    const long GW_HWNDPREV   = 3;
    const long GW_OWNER      = 4;
    const long GW_CHILD      = 5;
    const long GW_MAX		= 5;

    // WH_MSGFILTER Filter Proc Codes
    const long MSGF_DIALOGBOX    = 0;
	const long MSGF_MESSAGEBOX   = 1;
    const long MSGF_MENU         = 2;
    const long MSGF_MOVE         = 3;
    const long MSGF_SIZE         = 4;
    const long MSGF_SCROLLBAR    = 5;
    const long MSGF_NEXTWINDOW   = 6;
    const long MSGF_MAINLOOP     = 8;
    const long MSGF_MAX			 = 8;
    const long MSGF_USER         = 4096;

	// Shell support
    const long HSHELL_WINDOWCREATED			= 1;
    const long HSHELL_WINDOWDESTROYED		= 2;
    const long HSHELL_ACTIVATESHELLWINDOW	= 3;
	const long HSHELL_WINDOWACTIVATED		= 4;
	const long HSHELL_GETMINRECT			= 5;
	const long HSHELL_REDRAW				= 6;
	const long HSHELL_TASKMAN				= 7;
	const long HSHELL_LANGUAGE				= 8;

    // Standard cursor resource IDs
    const long IDC_ARROW          = 32512;
    const long IDC_IBEAM          = 32513;
    const long IDC_WAIT           = 32514;
    const long IDC_CROSS          = 32515;
    const long IDC_UPARROW        = 32516;
    const long IDC_SIZE           = 32640;
    const long IDC_ICON           = 32641;
    const long IDC_SIZENWSE       = 32642;
    const long IDC_SIZENESW       = 32643;
    const long IDC_SIZEWE         = 32644;
    const long IDC_SIZENS         = 32645;
    const long IDC_SIZEALL		  = 32646;
    const long IDC_NO		  = 32648;
    const long IDC_APPSTARTING		  = 32650;
    const long IDC_HELP		  = 32651;

    const long IMAGE_BITMAP         = 0;
    const long IMAGE_ICON           = 1;
    const long IMAGE_CURSOR         = 2;
    const long IMAGE_ENHMETAFILE    = 3;

    const long LR_DEFAULTCOLOR      = 0x0000;
    const long LR_MONOCHROME        = 0x0001;
    const long LR_COLOR             = 0x0002;
    const long LR_COPYRETURNORG     = 0x0004;
    const long LR_COPYDELETEORG     = 0x0008;
    const long LR_LOADFROMFILE      = 0x0010;
    const long LR_LOADTRANSPARENT   = 0x0020;
    const long LR_DEFAULTSIZE       = 0x0040;
    const long LR_VGACOLOR          = 0x0080;
    const long LR_LOADMAP3DCOLORS   = 0x1000;
    const long LR_CREATEDIBSECTION  = 0x2000;
    const long LR_COPYFROMRESOURCE  = 0x4000;
    const long LR_SHARED            = 0x8000;

    const long DI_MASK              = 0x0001;
    const long DI_IMAGE             = 0x0002;
    const long DI_NORMAL            = 0x0003;
    const long DI_COMPAT            = 0x0004;
    const long DI_DEFAULTSIZE       = 0x0008;

	const long RES_ICON    = 1;
	const long RES_CURSOR  = 2;

    // OEM Resource Ordinal Numbers
    const long OBM_CLOSE           = 32754;
    const long OBM_UPARROW         = 32753;
    const long OBM_DNARROW         = 32752;
    const long OBM_RGARROW         = 32751;
    const long OBM_LFARROW         = 32750;
    const long OBM_REDUCE          = 32749;
    const long OBM_ZOOM            = 32748;
    const long OBM_RESTORE         = 32747;
    const long OBM_REDUCED         = 32746;
    const long OBM_ZOOMD           = 32745;
    const long OBM_RESTORED        = 32744;
    const long OBM_UPARROWD        = 32743;
    const long OBM_DNARROWD        = 32742;
    const long OBM_RGARROWD        = 32741;
    const long OBM_LFARROWD        = 32740;
    const long OBM_MNARROW         = 32739;
    const long OBM_COMBO           = 32738;
    const long OBM_UPARROWI        = 32737;
    const long OBM_DNARROWI        = 32736;
    const long OBM_RGARROWI        = 32735;
    const long OBM_LFARROWI        = 32734;

    const long OBM_OLD_CLOSE       = 32767;
    const long OBM_SIZE            = 32766;
    const long OBM_OLD_UPARROW     = 32765;
    const long OBM_OLD_DNARROW     = 32764;
    const long OBM_OLD_RGARROW     = 32763;
    const long OBM_OLD_LFARROW     = 32762;
    const long OBM_BTSIZE          = 32761;
    const long OBM_CHECK           = 32760;
    const long OBM_CHECKBOXES      = 32759;
    const long OBM_BTNCORNERS      = 32758;
    const long OBM_OLD_REDUCE      = 32757;
    const long OBM_OLD_ZOOM        = 32756;
    const long OBM_OLD_RESTORE     = 32755;

    const long OCR_NORMAL          = 32512;
    const long OCR_IBEAM           = 32513;
    const long OCR_WAIT            = 32514;
    const long OCR_CROSS           = 32515;
    const long OCR_UP              = 32516;
    const long OCR_SIZE            = 32640;
    const long OCR_ICON            = 32641;
    const long OCR_SIZENWSE        = 32642;
    const long OCR_SIZENESW        = 32643;
    const long OCR_SIZEWE          = 32644;
    const long OCR_SIZENS          = 32645;
    const long OCR_SIZEALL         = 32646;
    const long OCR_ICOCUR          = 32647;
	const long OCR_NO              = 32648;
    const long OCR_APPSTARTING	   = 32650;

    const long OIC_SAMPLE          = 32512;
    const long OIC_HAND            = 32513;
    const long OIC_QUES            = 32514;
    const long OIC_BANG            = 32515;
    const long OIC_NOTE            = 32516;

    const long OIC_WINLOGO         = 32517;
    const long OIC_WARNING         = 32515; // OIC_BANG
    const long OIC_ERROR           = 32513; // OIC_HAND
    const long OIC_INFORMATION     = 32516; // OIC_NOTE


    // Standard icon resource IDs
    const long IDI_APPLICATION  = 32512;
    const long IDI_HAND         = 32513;
    const long IDI_QUESTION     = 32514;
    const long IDI_EXCLAMATION  = 32515;
    const long IDI_ASTERISK     = 32516;
    const long IDI_WINLOGO      =  32517;


    const long IDI_WARNING      = 32515;
    const long IDI_ERROR        = 32513;
    const long IDI_INFORMATION  = 32516;

    // Standard dialog button IDs
    const long IDOK         = 1;
    const long IDCANCEL     = 2;
    const long IDABORT		= 3;
    const long IDRETRY		= 4;
    const long IDIGNORE     = 5;
    const long IDYES        = 6;
    const long IDNO         = 7;

    const long IDCLOSE      = 8;
    const long IDHELP       = 9;

    // Edit control styles
    const long ES_LEFT        = 0x00000000;
    const long ES_CENTER      = 0x00000001;
    const long ES_RIGHT       = 0x00000002;
    const long ES_MULTILINE   = 0x00000004;
    const long ES_UPPERCASE   = 0x00000008;
    const long ES_LOWERCASE   = 0x00000010;
    const long ES_PASSWORD    = 0x00000020;
    const long ES_AUTOVSCROLL = 0x00000040;
    const long ES_AUTOHSCROLL = 0x00000080;
    const long ES_NOHIDESEL   = 0x00000100;
    const long ES_OEMCONVERT  = 0x00000400;
    const long ES_READONLY    = 0x00000800;
    const long ES_WANTRETURN  = 0x00001000;
    const long ES_NUMBER      = 0x00002000;

    // Edit control notification codes
    const long EN_SETFOCUS      = 0x0100;
    const long EN_KILLFOCUS     = 0x0200;
    const long EN_CHANGE        = 0x0300;
    const long EN_UPDATE        = 0x0400;
    const long EN_ERRSPACE      = 0x0500;
    const long EN_MAXTEXT       = 0x0501;
    const long EN_HSCROLL       = 0x0601;
    const long EN_VSCROLL       = 0x0602;

    // Edit control EM_SETMARGIN parameters
    const long EC_LEFTMARGIN       = 0x0001;
    const long EC_RIGHTMARGIN      = 0x0002;
    const long EC_USEFONTINFO      = 0xffff;

    // Edit control messages
    const long EM_GETSEL              = 0x00B0;
    const long EM_SETSEL              = 0x00B1;
    const long EM_GETRECT             = 0x00B2;
    const long EM_SETRECT             = 0x00B3;
    const long EM_SETRECTNP           = 0x00B4;
    const long EM_SCROLL              = 0x00B5;
    const long EM_LINESCROLL          = 0x00B6;
    const long EM_GETMODIFY           = 0x00B8;
    const long EM_SETMODIFY           = 0x00B9;
    const long EM_GETLINECOUNT        = 0x00BA;
    const long EM_LINEINDEX           = 0x00BB;
    const long EM_SETHANDLE           = 0x00BC;
    const long EM_GETHANDLE           = 0x00BD;
    const long EM_GETTHUMB            = 0x00BE;
    const long EM_LINELENGTH          = 0x00C1;
    const long EM_REPLACESEL          = 0x00C2;
    const long EM_GETLINE             = 0x00C4;
    const long EM_LIMITTEXT           = 0x00C5;
    const long EM_CANUNDO             = 0x00C6;
    const long EM_UNDO                = 0x00C7;
    const long EM_FMTLINES            = 0x00C8;
    const long EM_LINEFROMCHAR        = 0x00C9;
    const long EM_SETTABSTOPS         = 0x00CB;
    const long EM_SETPASSWORDCHAR     = 0x00CC;
    const long EM_EMPTYUNDOBUFFER     = 0x00CD;
    const long EM_GETFIRSTVISIBLELINE = 0x00CE;
    const long EM_SETREADONLY         = 0x00CF;
    const long EM_SETWORDBREAKPROC    = 0x00D0;
    const long EM_GETWORDBREAKPROC    = 0x00D1;
    const long EM_GETPASSWORDCHAR     = 0x00D2;
    const long EM_SETMARGINS          = 0x00D3;
    const long EM_GETMARGINS          = 0x00D4;
    const long EM_SETLIMITTEXT        = 0x00C5; // EM_LIMITTEXT
    const long EM_GETLIMITTEXT        = 0x00D5;
    const long EM_POSFROMCHAR         = 0x00D6;
    const long EM_CHARFROMPOS         = 0x00D7;


    // EDITWORDBREAKPROC code values
    const long WB_LEFT         = 0;
    const long WB_RIGHT        = 1;
    const long WB_ISDELIMITER  = 2;

    // Button Control Styles
    const long BS_PUSHBUTTON      = 0x00000000;
    const long BS_DEFPUSHBUTTON    = 0x00000001;
    const long BS_CHECKBOX        = 0x00000002;
    const long BS_AUTOCHECKBOX     = 0x00000003;
    const long BS_RADIOBUTTON     = 0x00000004;
    const long BS_3STATE      = 0x00000005;
    const long BS_AUTO3STATE      = 0x00000006;
    const long BS_GROUPBOX        = 0x00000007;
    const long BS_USERBUTTON      = 0x00000008;
    const long BS_AUTORADIOBUTTON  = 0x00000009;
    const long BS_OWNERDRAW       = 0x0000000B;
    const long BS_LEFTTEXT        = 0x00000020;
	const long BS_TEXT             = 0x00000000;
	const long BS_ICON             = 0x00000040;
	const long BS_BITMAP           = 0x00000080;
	const long BS_LEFT             = 0x00000100;
	const long BS_RIGHT            = 0x00000200;
	const long BS_CENTER           = 0x00000300;
	const long BS_TOP              = 0x00000400;
	const long BS_BOTTOM           = 0x00000800;
	const long BS_VCENTER          = 0x00000C00;
	const long BS_PUSHLIKE         = 0x00001000;
	const long BS_MULTILINE        = 0x00002000;
	const long BS_NOTIFY           = 0x00004000;
	const long BS_FLAT             = 0x00008000;
	const long BS_RIGHTBUTTON      = 0x00000020;

    // User Button Notification Codes
    const long BN_CLICKED       = 0;
    const long BN_PAINT         = 1;
    const long BN_HILITE        = 2;
    const long BN_UNHILITE      = 3;
    const long BN_DISABLE       = 4;
    const long BN_DOUBLECLICKED = 5;

    const long BN_PUSHED        = 2; // BN_HILITE
    const long BN_UNPUSHED      = 3; // BN_UNHILITE
    const long BN_DBLCLK		   = 5;
    const long BN_SETFOCUS      = 6;
    const long BN_KILLFOCUS     = 7;

    // Button Control Messages
	const long BM_GETCHECK        = 0x00F0;
	const long BM_SETCHECK        = 0x00F1;
	const long BM_GETSTATE        = 0x00F2;
	const long BM_SETSTATE        = 0x00F3;
	const long BM_SETSTYLE        = 0x00F4;
	const long BM_CLICK           = 0x00F5;
	const long BM_GETIMAGE        = 0x00F6;
	const long BM_SETIMAGE        = 0x00F7;

	const long BST_UNCHECKED      = 0x0000;
	const long BST_CHECKED        = 0x0001;
	const long BST_INDETERMINATE  = 0x0002;
	const long BST_PUSHED         = 0x0004;
	const long BST_FOCUS          = 0x0008;


    // ****** Static control *******
 
    // Static Control Styles
    const long SS_LEFT			   = 0x00000000;
    const long SS_CENTER           = 0x00000001;
    const long SS_RIGHT            = 0x00000002;
    const long SS_ICON             = 0x00000003;
    const long SS_BLACKRECT        = 0x00000004;
    const long SS_GRAYRECT         = 0x00000005;
    const long SS_WHITERECT        = 0x00000006;
    const long SS_BLACKFRAME       = 0x00000007;
    const long SS_GRAYFRAME        = 0x00000008;
    const long SS_WHITEFRAME       = 0x00000009;
    const long SS_SIMPLE		   = 0x0000000B;
    const long SS_LEFTNOWORDWRAP   = 0x0000000C;
	const long SS_OWNERDRAW        = 0x0000000D;
	const long SS_BITMAP           = 0x0000000E;
	const long SS_ENHMETAFILE      = 0x0000000F;
	const long SS_ETCHEDHORZ       = 0x00000010;
	const long SS_ETCHEDVERT       = 0x00000011;
	const long SS_ETCHEDFRAME      = 0x00000012;
	const long SS_TYPEMASK         = 0x0000001F;
    const long SS_NOPREFIX         = 0x00000080; // Don't do "&" character translation
	const long SS_NOTIFY           = 0x00000100;
	const long SS_CENTERIMAGE      = 0x00000200;
	const long SS_RIGHTJUST        = 0x00000400;
	const long SS_REALSIZEIMAGE    = 0x00000800;
	const long SS_SUNKEN           = 0x00001000;
	const long SS_ENDELLIPSIS      = 0x00004000;
	const long SS_PATHELLIPSIS     = 0x00008000;
	const long SS_WORDELLIPSIS     = 0x0000C000;
	const long SS_ELLIPSISMASK     = 0x0000C000;

    // Static Control Mesages
    const long STM_SETICON	       = 0x0170;
    const long STM_GETICON		   = 0x0171;
	const long STM_SETIMAGE        = 0x0172;
	const long STM_GETIMAGE        = 0x0173;
	const long STN_CLICKED         = 0;
	const long STN_DBLCLK          = 1;
	const long STN_ENABLE          = 2;
	const long STN_DISABLE         = 3;
	const long STM_MSGMAX          = 0x0174;


    const long WC_DIALOG  = 0x8002;

    // Get/SetWindowWord/Long offsets for use with WC_DIALOG windows
    const long DWL_MSGRESULT	= 0;
    const long DWL_DLGPROC		= 4;
    const long DWL_USER			= 8;

    // DlgDirList, DlgDirListComboBox flags values
    const long DDL_READWRITE          = 0x0000;
    const long DDL_READONLY           = 0x0001;
    const long DDL_HIDDEN             = 0x0002;
    const long DDL_SYSTEM             = 0x0004;
    const long DDL_DIRECTORY          = 0x0010;
    const long DDL_ARCHIVE            = 0x0020;

    const long DDL_POSTMSGS           = 0x2000;
    const long DDL_DRIVES             = 0x4000;
    const long DDL_EXCLUSIVE          = 0x8000;

    // Dialog styles
    const long DS_ABSALIGN        = 0x01;
    const long DS_SYSMODAL        = 0x02;
    const long DS_LOCALEDIT        = 0x20;
    const long DS_SETFONT          = 0x40;
    const long DS_MODALFRAME       = 0x80;
    const long DS_NOIDLEMSG        = 0x100;
	const long DS_3DLOOK           = 0x0004;
	const long DS_FIXEDSYS         = 0x0008;
	const long DS_NOFAILCREATE     = 0x0010;
	const long DS_CONTROL          = 0x0400;
	const long DS_CENTER           = 0x0800;
	const long DS_CENTERMOUSE      = 0x1000;
	const long DS_CONTEXTHELP      = 0x2000;

    // Dialog messages
    const long DM_GETDEFID      = 1024; //WM_USER+0;
    const long DM_SETDEFID      = 1025; //WM_USER+1;
	const long DM_REPOSITION    = 1026; //WM_USER+2;

	const long PSM_PAGEINFO        = 1124; //(WM_USER+100);
	const long PSM_SHEETINFO       = 1125; //(WM_USER+101);

	const long PSI_SETACTIVE       = 0x0001;
	const long PSI_KILLACTIVE      = 0x0002;
	const long PSI_APPLY           = 0x0003;
	const long PSI_RESET           = 0x0004;
	const long PSI_HASHELP         = 0x0005;
	const long PSI_HELP            = 0x0006;

	const long PSI_CHANGED         = 0x0001;
	const long PSI_GUISTART        = 0x0002;
	const long PSI_REBOOT          = 0x0003;
	const long PSI_GETSIBLINGS     = 0x0004;
	// Returned in HIWORD() of DM_GETDEFID result if msg is supported
    const long DC_HASDEFID      = 0x534B;

    // Dialog codes
    const long DLGC_WANTARROWS       = 0x0001;
    const long DLGC_WANTTAB          = 0x0002;
    const long DLGC_WANTALLKEYS      = 0x0004;
    const long DLGC_WANTMESSAGE      = 0x0004;
    const long DLGC_HASSETSEL        = 0x0008;
    const long DLGC_DEFPUSHBUTTON    = 0x0010;
    const long DLGC_UNDEFPUSHBUTTON  = 0x0020;
    const long DLGC_RADIOBUTTON      = 0x0040;
    const long DLGC_WANTCHARS        = 0x0080;
    const long DLGC_STATIC           = 0x0100;
    const long DLGC_BUTTON           = 0x2000;

    const long LB_CTLCODE      = 0;

    // Listbox message return values
    const long LB_OKAY          = 0;
    const long LB_ERR           = -1;
    const long LB_ERRSPACE      = -2;

    // Listbox notification codes
    const long LBN_ERRSPACE     = -2;
    const long LBN_SELCHANGE    = 1;
    const long LBN_DBLCLK       = 2;
    const long LBN_SELCANCEL    = 3;
    const long LBN_SETFOCUS     = 4;
    const long LBN_KILLFOCUS    = 5;

    // Listbox messages
    const long LB_ADDSTRING      = 0x180;
    const long LB_INSERTSTRING   = 0x181;
    const long LB_DELETESTRING   = 0x182;
    const long LB_SELITEMRANGEEX = 0x183;
    const long LB_RESETCONTENT   = 0x184;
    const long LB_SETSEL         = 0x185;
    const long LB_SETCURSEL      = 0x186;
    const long LB_GETSEL         = 0x187;
    const long LB_GETCURSEL      = 0x188;
    const long LB_GETTEXT        = 0x189;
    const long LB_GETTEXTLEN     = 0x18A;
    const long LB_GETCOUNT       = 0x18B;
    const long LB_SELECTSTRING   = 0x18C;
    const long LB_DIR            = 0x18D;
    const long LB_GETTOPINDEX    = 0x18E;
    const long LB_FINDSTRING     = 0x18F;
    const long LB_GETSELCOUNT    = 0x190;
    const long LB_GETSELITEMS    = 0x191;
    const long LB_SETTABSTOPS    = 0x192;
    const long LB_GETHORIZONTALEXTENT = 0x193;
    const long LB_SETHORIZONTALEXTENT = 0x194;
    const long LB_SETCOLUMNWIDTH = 0x195;
    const long LB_ADDFILE        = 0x196;
    const long LB_SETTOPINDEX    = 0x197;
    const long LB_GETITEMRECT    = 0x198;
    const long LB_GETITEMDATA    = 0x199;
    const long LB_SETITEMDATA    = 0x19A;
    const long LB_SELITEMRANGE   = 0x19B;
    const long LB_SETANCHORINDEX = 0x19C;
    const long LB_GETANCHORINDEX = 0x19D;
    const long LB_SETCARETINDEX  = 0x19E;
    const long LB_GETCARETINDEX  = 0x19F;
    const long LB_SETITEMHEIGHT  = 0x1A0;
    const long LB_GETITEMHEIGHT  = 0x1A1;
    const long LB_FINDSTRINGEXACT = 0x1A2;
    const long LB_SETLOCALE      = 0x1A5;
    const long LB_GETLOCALE      = 0x1A6;
    const long LB_SETCOUNT       = 0x1A7;
    const long LB_INITSTORAGE    = 0x1A8;
    const long LB_ITEMFROMPOINT  = 0x1A9;

    // Listbox styles
    const long LBS_NOTIFY            = 0x00000001;
    const long LBS_SORT              = 0x00000002;
    const long LBS_NOREDRAW          = 0x00000004;
    const long LBS_MULTIPLESEL       = 0x00000008;
    const long LBS_OWNERDRAWFIXED    = 0x00000010;
    const long LBS_OWNERDRAWVARIABLE = 0x00000020;
    const long LBS_HASSTRINGS        = 0x00000040;
    const long LBS_USETABSTOPS       = 0x00000080;
    const long LBS_NOINTEGRALHEIGHT  = 0x00000100;
    const long LBS_MULTICOLUMN       = 0x00000200;
    const long LBS_WANTKEYBOARDINPUT = 0x00000400;
    const long LBS_EXTENDEDSEL       = 0x00000800;
    const long LBS_DISABLENOSCROLL   = 0x00001000;
    const long LBS_NODATA            = 0x00002000;
    const long LBS_NOSEL             = 0x00004000;
    const long LBS_STANDARD          = 0x00A00003;  // (LBS_NOTIFY | LBS_SORT | WS_VSCROLL | WS_BORDER)

    // Combo box message return values
    const long CB_OKAY          = 0;
    const long CB_ERR           = -1;
    const long CB_ERRSPACE      = -2;

    // Combo box notification codes
    const long CBN_ERRSPACE     = -1;
    const long CBN_SELCHANGE    = 1;
    const long CBN_DBLCLK       = 2;
    const long CBN_SETFOCUS     = 3;
    const long CBN_KILLFOCUS    = 4;
    const long CBN_EDITCHANGE   = 5;
    const long CBN_EDITUPDATE   = 6;
    const long CBN_DROPDOWN     = 7;
    const long CBN_CLOSEUP      = 8;
    const long CBN_SELENDOK     = 9;
    const long CBN_SELENDCANCEL = 10;

    // Combo box styles
    const long CBS_SIMPLE			 = 0x0001;
    const long CBS_DROPDOWN          = 0x0002;
    const long CBS_DROPDOWNLIST      = 0x0003;
    const long CBS_OWNERDRAWFIXED    = 0x0010;
    const long CBS_OWNERDRAWVARIABLE = 0x0020;
    const long CBS_AUTOHSCROLL       = 0x0040;
    const long CBS_OEMCONVERT        = 0x0080;
    const long CBS_SORT              = 0x0100;
    const long CBS_HASSTRINGS        = 0x0200;
    const long CBS_NOINTEGRALHEIGHT  = 0x0400;
    const long CBS_DISABLENOSCROLL   = 0x0800;
    const long CBS_UPPERCASE		 = 0x2000;
    const long CBS_LOWERCASE		 = 0x4000;

    // Combo box messages
    const long CB_GETEDITSEL	= 0x140;
    const long CB_LIMITTEXT     = 0x141;
    const long CB_SETEDITSEL    = 0x142;
    const long CB_ADDSTRING     = 0x143;
    const long CB_DELETESTRING  = 0x144;
    const long CB_DIR           = 0x145;
    const long CB_GETCOUNT      = 0x146;
    const long CB_GETCURSEL     = 0x147;
    const long CB_GETLBTEXT     = 0x148;
    const long CB_GETLBTEXTLEN  = 0x149;
    const long CB_INSERTSTRING  = 0x14A;
    const long CB_RESETCONTENT  = 0x14B;
    const long CB_FINDSTRING    = 0x14C;
    const long CB_SELECTSTRING  = 0x14D;
    const long CB_SETCURSEL     = 0x14E;
    const long CB_SHOWDROPDOWN  = 0x14F;
    const long CB_GETITEMDATA   = 0x150;
    const long CB_SETITEMDATA   = 0x151;
    const long CB_GETDROPPEDCONTROLRECT	= 0x152;
    const long CB_SETITEMHEIGHT = 0x153;
    const long CB_GETITEMHEIGHT = 0x154;
    const long CB_SETEXTENDEDUI = 0x155;
    const long CB_GETEXTENDEDUI = 0x156;
    const long CB_GETDROPPEDSTATE = 0x157;
    const long CB_FINDSTRINGEXACT = 0x158;

    const long CB_SETLOCALE				= 0x159;
    const long CB_GETLOCALE				= 0x15A;
    const long CB_GETTOPINDEX			= 0x15B;
    const long CB_SETTOPINDEX			= 0x15C;
    const long CB_GETHORIZONTALEXTENT   = 0x15D;
    const long CB_SETHORIZONTALEXTENT   = 0x15E;
    const long CB_GETDROPPEDWIDTH       = 0x15F;
    const long CB_SETDROPPEDWIDTH       = 0x160;
    const long CB_INITSTORAGE           = 0x161;

    // Scroll bar styles
    const long SBS_HORZ						= 0x0000;
    const long SBS_VERT						= 0x0001;
    const long SBS_TOPALIGN					= 0x0002;
    const long SBS_LEFTALIGN				= 0x0002;
    const long SBS_BOTTOMALIGN				= 0x0004;
    const long SBS_RIGHTALIGN				= 0x0004;
    const long SBS_SIZEBOXTOPLEFTALIGN      = 0x0002;
    const long SBS_SIZEBOXBOTTOMRIGHTALIGN	= 0x0004;
    const long SBS_SIZEBOX					= 0x0008;
    const long SBS_SIZEGRIP					= 0x0010;

	// MDI client style bits
    const long MDIS_ALLCHILDSTYLES   = 0x0001;

    // wParam values for WM_MDITILE and WM_MDICASCADE messages.
    const long MDITILE_VERTICAL      = 0x0000;
    const long MDITILE_HORIZONTAL    = 0x0001;
    const long MDITILE_SKIPDISABLED  = 0x0002;

    // WinHelp() commands
    const long HELP_CONTEXT       = 0x0001;
    const long HELP_QUIT          = 0x0002;
    const long HELP_INDEX		  = 0x0003;
    const long HELP_CONTENTS      = 0x0003;
    const long HELP_HELPONHELP    = 0x0004;
    const long HELP_SETINDEX      = 0x0005;
    const long HELP_SETCONTENTS   = 0x0005;
    const long HELP_CONTEXTPOPUP  = 0x0008;
    const long HELP_FORCEFILE     = 0x0009;
    const long HELP_KEY           = 0x0101;
    const long HELP_COMMAND       = 0x0102;
    const long HELP_PARTIALKEY    = 0x0105;
    const long HELP_MULTIKEY      = 0x0201;
    const long HELP_SETWINPOS     = 0x0203;

	const long HELP_CONTEXTMENU  = 0x000a;
	const long HELP_FINDER       = 0x000b;
	const long HELP_WM_HELP      = 0x000c;
	const long HELP_SETPOPUP_POS = 0x000d;

	const long HELP_TCARD              = 0x8000;
	const long HELP_TCARD_DATA         = 0x0010;
	const long HELP_TCARD_OTHER_CALLER = 0x0011;

	const long IDH_NO_HELP                     = 28440;
	const long IDH_MISSING_CONTEXT             = 28441; // Control doesn't have matching help context
	const long IDH_GENERIC_HELP_BUTTON         = 28442; // Property sheet help button
	const long IDH_OK                          = 28443;
	const long IDH_CANCEL                      = 28444;
	const long IDH_HELP                        = 28445;

    const long SPI_GETBEEP                 = 1;
    const long SPI_SETBEEP                 = 2;
    const long SPI_GETMOUSE                = 3;
    const long SPI_SETMOUSE                = 4;
    const long SPI_GETBORDER               = 5;
    const long SPI_SETBORDER               = 6;
    const long SPI_GETKEYBOARDSPEED        = 10;
    const long SPI_SETKEYBOARDSPEED        = 11;
    const long SPI_LANGDRIVER              = 12;
    const long SPI_ICONHORIZONTALSPACING   = 13;
    const long SPI_GETSCREENSAVETIMEOUT    = 14;
    const long SPI_SETSCREENSAVETIMEOUT    = 15;
    const long SPI_GETSCREENSAVEACTIVE     = 16;
    const long SPI_SETSCREENSAVEACTIVE     = 17;
    const long SPI_GETGRIDGRANULARITY      = 18;
    const long SPI_SETGRIDGRANULARITY      = 19;
    const long SPI_SETDESKWALLPAPER        = 20;
    const long SPI_SETDESKPATTERN          = 21;
    const long SPI_GETKEYBOARDDELAY        = 22;
    const long SPI_SETKEYBOARDDELAY        = 23;
    const long SPI_ICONVERTICALSPACING     = 24;
    const long SPI_GETICONTITLEWRAP        = 25;
    const long SPI_SETICONTITLEWRAP        = 26;
    const long SPI_GETMENUDROPALIGNMENT    = 27;
    const long SPI_SETMENUDROPALIGNMENT    = 28;
    const long SPI_SETDOUBLECLKWIDTH       = 29;
    const long SPI_SETDOUBLECLKHEIGHT      = 30;
    const long SPI_GETICONTITLELOGFONT     = 31;
    const long SPI_SETDOUBLECLICKTIME      = 32;
    const long SPI_SETMOUSEBUTTONSWAP      = 33;
    const long SPI_SETICONTITLELOGFONT     = 34;
    const long SPI_GETFASTTASKSWITCH       = 35;
    const long SPI_SETFASTTASKSWITCH       = 36;
	const long SPI_SETDRAGFULLWINDOWS     = 37;

	const long SPI_GETDRAGFULLWINDOWS     = 38;
	const long SPI_GETNONCLIENTMETRICS    = 41;
	const long SPI_SETNONCLIENTMETRICS    = 42;
	const long SPI_GETMINIMIZEDMETRICS    = 43;
	const long SPI_SETMINIMIZEDMETRICS    = 44;
	const long SPI_GETICONMETRICS         = 45;
	const long SPI_SETICONMETRICS         = 46;
	const long SPI_SETWORKAREA            = 47;
	const long SPI_GETWORKAREA            = 48;
	const long SPI_SETPENWINDOWS          = 49;

	const long SPI_GETHIGHCONTRAST        = 66;
	const long SPI_SETHIGHCONTRAST        = 67;
	const long SPI_GETKEYBOARDPREF        = 68;
	const long SPI_SETKEYBOARDPREF        = 69;
	const long SPI_GETSCREENREADER        = 70;
	const long SPI_SETSCREENREADER        = 71;
	const long SPI_GETANIMATION           = 72;
	const long SPI_SETANIMATION           = 73;
	const long SPI_GETFONTSMOOTHING       = 74;
	const long SPI_SETFONTSMOOTHING       = 75;
	const long SPI_SETDRAGWIDTH           = 76;
	const long SPI_SETDRAGHEIGHT          = 77;
	const long SPI_SETHANDHELD            = 78;
	const long SPI_GETLOWPOWERTIMEOUT     = 79;
	const long SPI_GETPOWEROFFTIMEOUT     = 80;
	const long SPI_SETLOWPOWERTIMEOUT     = 81;
	const long SPI_SETPOWEROFFTIMEOUT     = 82;
	const long SPI_GETLOWPOWERACTIVE      = 83;
	const long SPI_GETPOWEROFFACTIVE      = 84;
	const long SPI_SETLOWPOWERACTIVE      = 85;
	const long SPI_SETPOWEROFFACTIVE      = 86;
	const long SPI_SETCURSORS             = 87;
	const long SPI_SETICONS               = 88;
	const long SPI_GETDEFAULTINPUTLANG    = 89;
	const long SPI_SETDEFAULTINPUTLANG    = 90;
	const long SPI_SETLANGTOGGLE          = 91;
	const long SPI_GETWINDOWSEXTENSION    = 92;
	const long SPI_SETMOUSETRAILS         = 93;
	const long SPI_GETMOUSETRAILS         = 94;
	const long SPI_SCREENSAVERRUNNING     = 97;
	const long SPI_GETFILTERKEYS          = 50;
	const long SPI_SETFILTERKEYS          = 51;
	const long SPI_GETTOGGLEKEYS          = 52;
	const long SPI_SETTOGGLEKEYS          = 53;
	const long SPI_GETMOUSEKEYS           = 54;
	const long SPI_SETMOUSEKEYS           = 55;
	const long SPI_GETSHOWSOUNDS          = 56;
	const long SPI_SETSHOWSOUNDS          = 57;
	const long SPI_GETSTICKYKEYS          = 58;
	const long SPI_SETSTICKYKEYS          = 59;
	const long SPI_GETACCESSTIMEOUT       = 60;
	const long SPI_SETACCESSTIMEOUT       = 61;
	const long SPI_GETSERIALKEYS          = 62;
	const long SPI_SETSERIALKEYS          = 63;
	const long SPI_GETSOUNDSENTRY         = 64;
	const long SPI_SETSOUNDSENTRY         = 65;
	const long SPI_GETSNAPTODEFBUTTON     = 95;
	const long SPI_SETSNAPTODEFBUTTON     = 96;
	const long SPI_GETMOUSEHOVERWIDTH     = 98;
	const long SPI_SETMOUSEHOVERWIDTH     = 99;
	const long SPI_GETMOUSEHOVERHEIGHT   = 100;
	const long SPI_SETMOUSEHOVERHEIGHT   = 101;
	const long SPI_GETMOUSEHOVERTIME     = 102;
	const long SPI_SETMOUSEHOVERTIME     = 103;
	const long SPI_GETWHEELSCROLLLINES   = 104;
	const long SPI_SETWHEELSCROLLLINES   = 105;

    // SystemParametersInfo flags
    const long SPIF_UPDATEINIFILE       = 0x0001;
    const long SPIF_SENDWININICHANGE    = 0x0002;
    const long SPIF_SENDCHANGE          = 0x0002; // SPIF_SENDWININICHANGE    

	const long METRICS_USEDEFAULT = -1;


	const long ARW_BOTTOMLEFT              = 0x0000;
	const long ARW_BOTTOMRIGHT             = 0x0001;
	const long ARW_TOPLEFT                 = 0x0002;
	const long ARW_TOPRIGHT                = 0x0003;
	const long ARW_STARTMASK               = 0x0003;
	const long ARW_STARTRIGHT              = 0x0001;
	const long ARW_STARTTOP                = 0x0002;

	const long ARW_LEFT                    = 0x0000;
	const long ARW_RIGHT                   = 0x0000;
	const long ARW_UP                      = 0x0004;
	const long ARW_DOWN                    = 0x0004;
	const long ARW_HIDE                    = 0x0008;
	const long ARW_VALID                   = 0x000F;

	const long SERKF_SERIALKEYSON  = 0x00000001;
	const long SERKF_AVAILABLE     = 0x00000002;
	const long SERKF_INDICATOR     = 0x00000004;

	// flags for HIGHCONTRAST dwFlags field 
	const long HCF_HIGHCONTRASTON  = 0x00000001;
	const long HCF_AVAILABLE       = 0x00000002;
	const long HCF_HOTKEYACTIVE    = 0x00000004;
	const long HCF_CONFIRMHOTKEY   = 0x00000008;
	const long HCF_HOTKEYSOUND     = 0x00000010;
	const long HCF_INDICATOR       = 0x00000020;
	const long HCF_HOTKEYAVAILABLE = 0x00000040;

	// Flags for ChangeDisplaySettings 
	const long CDS_UPDATEREGISTRY  = 0x00000001;
	const long CDS_TEST            = 0x00000002;
	const long CDS_FULLSCREEN      = 0x00000004;
	const long CDS_GLOBAL          = 0x00000008;
	const long CDS_SET_PRIMARY     = 0x00000010;
	const long CDS_RESET           = 0x40000000;
	const long CDS_SETRECT         = 0x20000000;
	const long CDS_NORESET         = 0x10000000;

	// Return values for ChangeDisplaySettings 
	const long DISP_CHANGE_SUCCESSFUL       = 0;
	const long DISP_CHANGE_RESTART          = 1;
	const long DISP_CHANGE_FAILED          = -1;
	const long DISP_CHANGE_BADMODE         = -2;
	const long DISP_CHANGE_NOTUPDATED      = -3;
	const long DISP_CHANGE_BADFLAGS        = -4;
	const long DISP_CHANGE_BADPARAM        = -5;

//Form wingdi.h
	const long CCHDEVICENAME = 32;
	const long CCHFORMNAME = 32;

	const long ENUM_CURRENT_SETTINGS       = -1;
	const long ENUM_REGISTRY_SETTINGS      = -2;

	// FILTERKEYS dwFlags field
	const long FKF_FILTERKEYSON    = 0x00000001;
	const long FKF_AVAILABLE       = 0x00000002;
	const long FKF_HOTKEYACTIVE    = 0x00000004;
	const long FKF_CONFIRMHOTKEY   = 0x00000008;
	const long FKF_HOTKEYSOUND     = 0x00000010;
	const long FKF_INDICATOR       = 0x00000020;
	const long FKF_CLICKON         = 0x00000040;

	// STICKYKEYS dwFlags field
	const long SKF_STICKYKEYSON    = 0x00000001;
	const long SKF_AVAILABLE       = 0x00000002;
	const long SKF_HOTKEYACTIVE    = 0x00000004;
	const long SKF_CONFIRMHOTKEY   = 0x00000008;
	const long SKF_HOTKEYSOUND     = 0x00000010;
	const long SKF_INDICATOR       = 0x00000020;
	const long SKF_AUDIBLEFEEDBACK = 0x00000040;
	const long SKF_TRISTATE        = 0x00000080;
	const long SKF_TWOKEYSOFF      = 0x00000100;

	// MOUSEKEYS dwFlags field
	const long MKF_MOUSEKEYSON     = 0x00000001;
	const long MKF_AVAILABLE       = 0x00000002;
	const long MKF_HOTKEYACTIVE    = 0x00000004;
	const long MKF_CONFIRMHOTKEY   = 0x00000008;
	const long MKF_HOTKEYSOUND     = 0x00000010;
	const long MKF_INDICATOR       = 0x00000020;
	const long MKF_MODIFIERS       = 0x00000040;
	const long MKF_REPLACENUMBERS  = 0x00000080;

	// ACCESSTIMEOUT dwFlags field
	const long ATF_TIMEOUTON       = 0x00000001;
	const long ATF_ONOFFFEEDBACK   = 0x00000002;

	// values for SOUNDSENTRY iFSGrafEffect field 
	const long SSGF_NONE       = 0;
	const long SSGF_DISPLAY    = 3;

	// values for SOUNDSENTRY iFSTextEffect field 
	const long SSTF_NONE       = 0;
	const long SSTF_CHARS      = 1;
	const long SSTF_BORDER     = 2;
	const long SSTF_DISPLAY    = 3;

	// values for SOUNDSENTRY iWindowsEffect field 
	const long SSWF_NONE     = 0;
	const long SSWF_TITLE    = 1;
	const long SSWF_WINDOW   = 2;
	const long SSWF_DISPLAY  = 3;
	const long SSWF_CUSTOM   = 4;

	// SOUNDSENTRY dwFlags field
	const long SSF_SOUNDSENTRYON   = 0x00000001;
	const long SSF_AVAILABLE       = 0x00000002;
	const long SSF_INDICATOR       = 0x00000004;

	// TOGGLEKEYS dwFlags field
	const long TKF_TOGGLEKEYSON    = 0x00000001;
	const long TKF_AVAILABLE       = 0x00000002;
	const long TKF_HOTKEYACTIVE    = 0x00000004;
	const long TKF_CONFIRMHOTKEY   = 0x00000008;
	const long TKF_HOTKEYSOUND     = 0x00000010;
	const long TKF_INDICATOR       = 0x00000020;


	// SetLastErrorEx() types.
	const long SLE_ERROR       = 0x00000001;
	const long SLE_MINORERROR  = 0x00000002;
	const long SLE_WARNING     = 0x00000003;

	// Special HWND value for use with PostMessage() and SendMessage()
    [helpstring("Special HWND value for use with PostMessage() and SendMessage()")]
	const long HWND_BROADCAST  = 0xffff;

/*
 * Low level hook flags
 */
	const long LLKHF_EXTENDED      = 0x00000001;
	const long LLKHF_INJECTED      = 0x00000010;
	const long LLKHF_ALTDOWN       = 0x00000020;
	const long LLKHF_UP            = 0x00000080;

	const long LLMHF_INJECTED      = 0x00000001;

/*
 * SetWindowsHook() codes
 */

const long WH_MIN              = (-1);
const long WH_MSGFILTER        = (-1);
const long WH_JOURNALRECORD    = 0;
const long WH_JOURNALPLAYBACK  = 1;
const long WH_KEYBOARD         = 2;
const long WH_GETMESSAGE       = 3;
const long WH_CALLWNDPROC      = 4;
const long WH_CBT              = 5;
const long WH_SYSMSGFILTER     = 6;
const long WH_MOUSE            = 7;
const long WH_HARDWARE         = 8;
const long WH_DEBUG            = 9;
const long WH_SHELL           = 10;
const long WH_FOREGROUNDIDLE  = 11;
const long WH_CALLWNDPROCRET  = 12;
const long WH_KEYBOARD_LL     = 13;
const long WH_MOUSE_LL        = 14;

const long WH_MAX             = 11;

/*
 * Hook Codes
 */
const long HC_ACTION           = 0;
const long HC_GETNEXT          = 1;
const long HC_SKIP             = 2;
const long HC_NOREMOVE         = 3;
const long HC_NOREM            = 3;
const long HC_SYSMODALON       = 4;
const long HC_SYSMODALOFF      = 5;

/*
 * CBT Hook Codes
 */
const long HCBT_MOVESIZE       = 0;
const long HCBT_MINMAX         = 1;
const long HCBT_QS             = 2;
const long HCBT_CREATEWND      = 3;
const long HCBT_DESTROYWND     = 4;
const long HCBT_ACTIVATE       = 5;
const long HCBT_CLICKSKIPPED   = 6;
const long HCBT_KEYSKIPPED     = 7;
const long HCBT_SYSCOMMAND     = 8;
const long HCBT_SETFOCUS       = 9;


/*
 * Owner draw control types
 */
const long  ODT_MENU        = 1;
const long  ODT_LISTBOX     = 2;
const long  ODT_COMBOBOX    = 3;
const long  ODT_BUTTON      = 4;
const long  ODT_STATIC      = 5;
const long  ODT_TAB			= 101;

/*
 * Owner draw actions
 */
const long  ODA_DRAWENTIRE = 0x0001;
const long  ODA_SELECT     = 0x0002;
const long  ODA_FOCUS      = 0x0004;

/*
 * Owner draw state
 */
const long  ODS_SELECTED   = 0x0001;
const long  ODS_GRAYED     = 0x0002;
const long  ODS_DISABLED   = 0x0004;
const long  ODS_CHECKED    = 0x0008;
const long  ODS_FOCUS      = 0x0010;
const long  ODS_DEFAULT        = 0x0020;
const long  ODS_COMBOBOXEDIT   = 0x1000;
const long  ODS_HOTLIGHT       = 0x0040;
const long  ODS_INACTIVE       = 0x0080;


	// RedrawWindow() flags 
    const DWORD RDW_INVALIDATE      = 0x0001;
    const DWORD RDW_INTERNALPAINT   = 0x0002;
    const DWORD RDW_ERASE           = 0x0004;

    const DWORD RDW_VALIDATE        = 0x0008;
    const DWORD RDW_NOINTERNALPAINT = 0x0010;
    const DWORD RDW_NOERASE         = 0x0020;

    const DWORD RDW_NOCHILDREN      = 0x0040;
    const DWORD RDW_ALLCHILDREN     = 0x0080;

    const DWORD RDW_UPDATENOW       = 0x0100;
    const DWORD RDW_ERASENOW        = 0x0200;

    const DWORD RDW_FRAME           = 0x0400;
    const DWORD RDW_NOFRAME         = 0x0800;

	const long CW_USEDEFAULT       = 0x80000000;
	
	const long INVALID_SET_FILE_POINTER = -1;

};