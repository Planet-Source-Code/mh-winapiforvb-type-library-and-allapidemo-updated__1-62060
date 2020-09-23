[
	helpstring("Error Constants"),
	dllname("NOSUCHG.DLL")
]
module GDIConstant {

    // Binary raster ops 
    const long R2_BLACK            = 1;   //  0
    const long R2_NOTMERGEPEN      = 2;   // DPon
    const long R2_MASKNOTPEN       = 3;   // DPna
    const long R2_NOTCOPYPEN       = 4;   // PN
    const long R2_MASKPENNOT       = 5;   // PDna
    const long R2_NOT              = 6;   // Dn
    const long R2_XORPEN           = 7;   // DPx
    const long R2_NOTMASKPEN       = 8;   // DPan
    const long R2_MASKPEN          = 9;   // DPa
    const long R2_NOTXORPEN        = 10;  // DPxn
    const long R2_NOP              = 11;  // D
    const long R2_MERGENOTPEN      = 12;  // DPno
    const long R2_COPYPEN          = 13;  // P
    const long R2_MERGEPENNOT      = 14;  // PDno
    const long R2_MERGEPEN         = 15;  // DPo
    const long R2_WHITE            = 16;  //  1
    const long R2_LAST             = 16;

    // Ternary raster operations
    const long SRCCOPY     = 0x00CC0020; // dest = source
    const long SRCPAINT    = 0x00EE0086; // dest = source OR dest
    const long SRCAND      = 0x008800C6; // dest = source AND dest
    const long SRCINVERT   = 0x00660046; // dest = source XOR dest
    const long SRCERASE    = 0x00440328; // dest = source AND (NOT dest )
    const long NOTSRCCOPY  = 0x00330008; // dest = (NOT source)
    const long NOTSRCERASE = 0x001100A6; // dest = (NOT src) AND (NOT dest)
    const long MERGECOPY   = 0x00C000CA; // dest = (source AND pattern)
    const long MERGEPAINT  = 0x00BB0226; // dest = (NOT source) OR dest
    const long PATCOPY     = 0x00F00021; // dest = pattern
    const long PATPAINT    = 0x00FB0A09; // dest = DPSnoo
    const long PATINVERT   = 0x005A0049; // dest = pattern XOR dest
    const long DSTINVERT   = 0x00550009; // dest = (NOT dest)
    const long BLACKNESS   = 0x00000042; // dest = BLACK
    const long WHITENESS   = 0x00FF0062; // dest = WHITE


    const long GDI_ERROR             = 0xFFFFFFFF;
    const HANDLE HGDI_ERROR         = 0xFFFFFFFF;

    // Region Flags 
    //const long ERROR               = 0;	// Conflict with VB's Error function
    const long NULLREGION          = 1;
    const long SIMPLEREGION        = 2;
    const long COMPLEXREGION       = 3;
    const long RGN_ERROR           = 0; // ERROR

    // CombineRgn() Styles 
    const long RGN_AND             = 1;
    const long RGN_OR              = 2;
    const long RGN_XOR             = 3;
    const long RGN_DIFF            = 4;
    const long RGN_COPY            = 5;
    const long RGN_MIN             = 1; // RGN_AND
    const long RGN_MAX             = 5; // RGN_COPY

    // StretchBlt() Modes 
    const long BLACKONWHITE                 = 1;
    const long WHITEONBLACK                 = 2;
    const long COLORONCOLOR                 = 3;
    const long HALFTONE                     = 4;
    const long MAXSTRETCHBLTMODE            = 4;

    // new StretchBlt() Modes (simpler names)
    const long STRETCH_ANDSCANS   = 1;
    const long STRETCH_ORSCANS        = 2;
    const long STRETCH_DELETESCANS    = 3;

    // PolyFill() Modes 
    const long ALTERNATE                    = 1;
    const long WINDING                      = 2;
    const long POLYFILL_LAST                = 2;

    // Text Alignment Options 
    const long TA_NOUPDATECP                = 0;
    const long TA_UPDATECP                  = 1;

    const long TA_LEFT                      = 0;
    const long TA_RIGHT                     = 2;
    const long TA_CENTER                    = 6;

    const long TA_TOP                       = 0;
    const long TA_BOTTOM                    = 8;
    const long TA_BASELINE                  = 24;
    const long TA_RTLREADING                = 256;
    const long TA_MASK       = 24 + 6 + 1 + 256; // (TA_BASELINE+TA_CENTER+TA_UPDATECP+TA_RTLREADING);

    const long VTA_BASELINE = 24; // TA_BASELINE;
    const long VTA_LEFT     = 8;  // TA_BOTTOM;
    const long VTA_RIGHT    = 0;  // TA_TOP;
    const long VTA_CENTER   = 6;  // TA_CENTER;
    const long VTA_BOTTOM   = 8;  // TA_RIGHT;
    const long VTA_TOP      = 0;  // TA_LEFT;

    const UINT ETO_OPAQUE                   = 0x0002;
    const UINT ETO_CLIPPED                  = 0x0004;
    const UINT ETO_GLYPH_INDEX              = 0x0010;
    const UINT ETO_RTLREADING               = 0x0080;

    const UINT ASPECT_FILTERING             = 0x0001;

    // Bounds Accumulation APIs 

    const UINT DCB_RESET       = 0x0001;
    const UINT DCB_ACCUMULATE  = 0x0002;
    const UINT DCB_DIRTY       = 0x0002; // DCB_ACCUMULATE
    const UINT DCB_SET         = 0x0003; // DCB_RESET | DCB_ACCUMULATE)
    const UINT DCB_ENABLE      = 0x0004;
    const UINT DCB_DISABLE     = 0x0008;

    // Metafile Functions 
    const UINT META_SETBKCOLOR              = 0x0201;
    const UINT META_SETBKMODE               = 0x0102;
    const UINT META_SETMAPMODE              = 0x0103;
    const UINT META_SETROP2                 = 0x0104;
    const UINT META_SETRELABS               = 0x0105;
    const UINT META_SETPOLYFILLMODE         = 0x0106;
    const UINT META_SETSTRETCHBLTMODE       = 0x0107;
    const UINT META_SETTEXTCHAREXTRA        = 0x0108;
    const UINT META_SETTEXTCOLOR            = 0x0209;
    const UINT META_SETTEXTJUSTIFICATION    = 0x020A;
    const UINT META_SETWINDOWORG            = 0x020B;
    const UINT META_SETWINDOWEXT            = 0x020C;
    const UINT META_SETVIEWPORTORG          = 0x020D;
    const UINT META_SETVIEWPORTEXT          = 0x020E;
    const UINT META_OFFSETWINDOWORG         = 0x020F;
    const UINT META_SCALEWINDOWEXT          = 0x0410;
    const UINT META_OFFSETVIEWPORTORG       = 0x0211;
    const UINT META_SCALEVIEWPORTEXT        = 0x0412;
    const UINT META_LINETO                  = 0x0213;
    const UINT META_MOVETO                  = 0x0214;
    const UINT META_EXCLUDECLIPRECT         = 0x0415;
    const UINT META_INTERSECTCLIPRECT       = 0x0416;
    const UINT META_ARC                     = 0x0817;
    const UINT META_ELLIPSE                 = 0x0418;
    const UINT META_FLOODFILL               = 0x0419;
    const UINT META_PIE                     = 0x081A;
    const UINT META_RECTANGLE               = 0x041B;
    const UINT META_ROUNDRECT               = 0x061C;
    const UINT META_PATBLT                  = 0x061D;
    const UINT META_SAVEDC                  = 0x001E;
    const UINT META_SETPIXEL                = 0x041F;
    const UINT META_OFFSETCLIPRGN           = 0x0220;
    const UINT META_TEXTOUT                 = 0x0521;
    const UINT META_BITBLT                  = 0x0922;
    const UINT META_STRETCHBLT              = 0x0B23;
    const UINT META_POLYGON                 = 0x0324;
    const UINT META_POLYLINE                = 0x0325;
    const UINT META_ESCAPE                  = 0x0626;
    const UINT META_RESTOREDC               = 0x0127;
    const UINT META_FILLREGION              = 0x0228;
    const UINT META_FRAMEREGION             = 0x0429;
    const UINT META_INVERTREGION            = 0x012A;
    const UINT META_PAINTREGION             = 0x012B;
    const UINT META_SELECTCLIPREGION        = 0x012C;
    const UINT META_SELECTOBJECT            = 0x012D;
    const UINT META_SETTEXTALIGN            = 0x012E;
    const UINT META_CHORD                   = 0x0830;
    const UINT META_SETMAPPERFLAGS          = 0x0231;
    const UINT META_EXTTEXTOUT              = 0x0a32;
    const UINT META_SETDIBTODEV             = 0x0d33;
    const UINT META_SELECTPALETTE           = 0x0234;
    const UINT META_REALIZEPALETTE          = 0x0035;
    const UINT META_ANIMATEPALETTE          = 0x0436;
    const UINT META_SETPALENTRIES           = 0x0037;
    const UINT META_POLYPOLYGON             = 0x0538;
    const UINT META_RESIZEPALETTE           = 0x0139;
    const UINT META_DIBBITBLT               = 0x0940;
    const UINT META_DIBSTRETCHBLT           = 0x0b41;
    const UINT META_DIBCREATEPATTERNBRUSH   = 0x0142;
    const UINT META_STRETCHDIB              = 0x0f43;
    const UINT META_EXTFLOODFILL            = 0x0548;
    const UINT META_DELETEOBJECT            = 0x01f0;
    const UINT META_CREATEPALETTE           = 0x00f7;
    const UINT META_CREATEPATTERNBRUSH      = 0x01F9;
    const UINT META_CREATEPENINDIRECT       = 0x02FA;
    const UINT META_CREATEFONTINDIRECT      = 0x02FB;
    const UINT META_CREATEBRUSHINDIRECT     = 0x02FC;
    const UINT META_CREATEREGION            = 0x06FF;

    // GDI Escapes 
    const long ESC_NEWFRAME                  = 1;
    const long ESC_ABORTDOC                  = 2;
    const long ESC_NEXTBAND                  = 3;
    const long ESC_SETCOLORTABLE             = 4;
    const long ESC_GETCOLORTABLE             = 5;
    const long ESC_FLUSHOUTPUT               = 6;
    const long ESC_DRAFTMODE                 = 7;
    const long ESC_QUERYESCSUPPORT           = 8;
    const long ESC_SETABORTPROC              = 9;
    const long ESC_STARTDOC                  = 10;
    const long ESC_ENDDOC                    = 11;
    const long ESC_GETPHYSPAGESIZE           = 12;
    const long ESC_GETPRINTINGOFFSET         = 13;
    const long ESC_GETSCALINGFACTOR          = 14;
    const long ESC_MFCOMMENT                 = 15;
    const long ESC_GETPENWIDTH               = 16;
    const long ESC_SETCOPYCOUNT              = 17;
    const long ESC_SELECTPAPERSOURCE         = 18;
    const long ESC_DEVICEDATA                = 19;
    const long ESC_PASSTHROUGH               = 19;
    const long ESC_GETTECHNOLGY              = 20;
    const long ESC_GETTECHNOLOGY             = 20;
    const long ESC_SETLINECAP                = 21;
    const long ESC_SETLINEJOIN               = 22;
    const long ESC_SETMITERLIMIT             = 23;
    const long ESC_BANDINFO                  = 24;
    const long ESC_DRAWPATTERNRECT           = 25;
    const long ESC_GETVECTORPENSIZE          = 26;
    const long ESC_GETVECTORBRUSHSIZE        = 27;
    const long ESC_ENABLEDUPLEX              = 28;
    const long ESC_GETSETPAPERBINS           = 29;
    const long ESC_GETSETPRINTORIENT         = 30;
    const long ESC_ENUMPAPERBINS             = 31;
    const long ESC_SETDIBSCALING             = 32;
    const long ESC_EPSPRINTING               = 33;
    const long ESC_ENUMPAPERMETRICS          = 34;
    const long ESC_GETSETPAPERMETRICS        = 35;
    const long ESC_POSTSCRIPT_DATA           = 37;
    const long ESC_POSTSCRIPT_IGNORE         = 38;
    const long ESC_MOUSETRAILS               = 39;
    const long ESC_GETDEVICEUNITS            = 42;

    const long ESC_GETEXTENDEDTEXTMETRICS    = 256;
    const long ESC_GETEXTENTTABLE            = 257;
    const long ESC_GETPAIRKERNTABLE          = 258;
    const long ESC_GETTRACKKERNTABLE         = 259;
    const long ESC_EXTTEXTOUT                = 512;
    const long ESC_GETFACENAME               = 513;
    const long ESC_DOWNLOADFACE              = 514;
    const long ESC_ENABLERELATIVEWIDTHS      = 768;
    const long ESC_ENABLEPAIRKERNING         = 769;
    const long ESC_SETKERNTRACK              = 770;
    const long ESC_SETALLJUSTVALUES          = 771;
    const long ESC_SETCHARSET                = 772;

    const long ESC_STRETCHBLT                = 2048;
    const long ESC_GETSETSCREENPARAMS        = 3072;
    const long ESC_QUERYDIBSUPPORT           = 3073;
    const long ESC_BEGIN_PATH                = 4096;
    const long ESC_CLIP_TO_PATH              = 4097;
    const long ESC_END_PATH                  = 4098;
    const long ESC_EXT_DEVICE_CAPS           = 4099;
    const long ESC_RESTORE_CTM               = 4100;
    const long ESC_SAVE_CTM                  = 4101;
    const long ESC_SET_ARC_DIRECTION         = 4102;
    const long ESC_SET_BACKGROUND_COLOR      = 4103;
    const long ESC_SET_POLY_MODE             = 4104;
    const long ESC_SET_SCREEN_ANGLE          = 4105;
    const long ESC_SET_SPREAD                = 4106;
    const long ESC_TRANSFORM_CTM             = 4107;
    const long ESC_SET_CLIP_BOX              = 4108;
    const long ESC_SET_BOUNDS                = 4109;
    const long ESC_SET_MIRROR_MODE           = 4110;
    const long ESC_OPENCHANNEL               = 4110;
    const long ESC_DOWNLOADHEADER            = 4111;
    const long ESC_CLOSECHANNEL              = 4112;
    const long ESC_POSTSCRIPT_PASSTHROUGH    = 4115;
    const long ESC_ENCAPSULATED_POSTSCRIPT   = 4116;
    // Flag returned from QUERYDIBSUPPORT
    const long QDI_SETDIBITS                = 1;
    const long QDI_GETDIBITS                = 2;
    const long QDI_DIBTOSCREEN              = 4;
    const long QDI_STRETCHDIB               = 8;

    // Spooler Error Codes
    const long SP_NOTREPORTED               = 0x4000;
    const long SP_ERROR                     = -1;
    const long SP_APPABORT                  = -2;
    const long SP_USERABORT                 = -3;
    const long SP_OUTOFDISK                 = -4;
    const long SP_OUTOFMEMORY               = -5;

    const UINT PR_JOBSTATUS                = 0x0000;

    // Object Definitions for EnumObjects() 
    const long OBJ_PEN             = 1;
    const long OBJ_BRUSH           = 2;
    const long OBJ_DC              = 3;
    const long OBJ_METADC          = 4;
    const long OBJ_PAL             = 5;
    const long OBJ_FONT            = 6;
    const long OBJ_BITMAP          = 7;
    const long OBJ_REGION          = 8;
    const long OBJ_METAFILE        = 9;
    const long OBJ_MEMDC           = 10;
    const long OBJ_EXTPEN          = 11;
    const long OBJ_ENHMETADC       = 12;
    const long OBJ_ENHMETAFILE     = 13;

    // xform stuff 
    const long MWT_IDENTITY        = 1;
    const long MWT_LEFTMULTIPLY    = 2;
    const long MWT_RIGHTMULTIPLY   = 3;

    const long MWT_MIN             = 1; // MWT_IDENTITY
    const long MWT_MAX             = 3; // MWT_RIGHTMULTIPLY

    // Image Color Matching color definitions 

    const ULONG LCS_CALIBRATED_RGB        = 0x00000000;
    const ULONG LCS_DEVICE_RGB            = 0x00000001;
    const ULONG LCS_DEVICE_CMYK           = 0x00000002;

    const ULONG LCS_GM_BUSINESS           = 0x00000001;
    const ULONG LCS_GM_GRAPHICS           = 0x00000002;
    const ULONG LCS_GM_IMAGES             = 0x00000004;

    // ICM Defines for results from CheckColorInGamut() 
    const long CM_OUT_OF_GAMUT       = 255;
    const long CM_IN_GAMUT           = 0;

    // constants for the biCompression field 
    const long BI_RGB        = 0;
    const long BI_RLE8       = 1;
    const long BI_RLE4       = 2;
    const long BI_BITFIELDS  = 3;

    const long TCI_SRCCHARSET    = 1;
    const long TCI_SRCCODEPAGE   = 2;
    const long TCI_SRCFONTSIG    = 3;

    // tmPitchAndFamily flags 
    const long TMPF_FIXED_PITCH    = 0x01;
    const long TMPF_VECTOR         = 0x02;
    const long TMPF_DEVICE         = 0x08;
    const long TMPF_TRUETYPE       = 0x04;

    // ntmFlags field flags 
    const long NTM_REGULAR     = 0x00000040;
    const long NTM_BOLD        = 0x00000020;
    const long NTM_ITALIC      = 0x00000001;

    // Logical Font 
    const long LF_FACESIZE         = 32;

    const long LF_FULLFACESIZE     = 64;

    const long OUT_DEFAULT_PRECIS      = 0;
    const long OUT_STRING_PRECIS       = 1;
    const long OUT_CHARACTER_PRECIS    = 2;
    const long OUT_STROKE_PRECIS       = 3;
    const long OUT_TT_PRECIS           = 4;
    const long OUT_DEVICE_PRECIS       = 5;
    const long OUT_RASTER_PRECIS       = 6;
    const long OUT_TT_ONLY_PRECIS      = 7;
    const long OUT_OUTLINE_PRECIS      = 8;
    const UINT CLIP_DEFAULT_PRECIS     = 0x00;
    const UINT CLIP_CHARACTER_PRECIS   = 0x01;
    const UINT CLIP_STROKE_PRECIS      = 0x02;
    const UINT CLIP_MASK               = 0x0f;
    const UINT CLIP_LH_ANGLES          = 0x04; // (1<<4)
    const UINT CLIP_TT_ALWAYS          = 0x40; // (2<<4)
    const UINT CLIP_EMBEDDED           = 0x80; // (8<<4)

    const long DEFAULT_QUALITY         = 0;
    const long DRAFT_QUALITY           = 1;
    const long PROOF_QUALITY           = 2;
    const long NONANTIALIASED_QUALITY  = 3;
    const long ANTIALIASED_QUALITY     = 4;

    const long DEFAULT_PITCH           = 0;
    const long FIXED_PITCH             = 1;
    const long VARIABLE_PITCH          = 2;
    const long MONO_FONT               = 8;

    const long ANSI_CHARSET            = 0;
    const long DEFAULT_CHARSET         = 1;
    const long SYMBOL_CHARSET          = 2;
    const long SHIFTJIS_CHARSET        = 128;
    const long HANGEUL_CHARSET         = 129;
    const long GB2312_CHARSET          = 134;
    const long CHINESEBIG5_CHARSET     = 136;
    const long OEM_CHARSET             = 255;
    const long JOHAB_CHARSET           = 130;
    const long HEBREW_CHARSET          = 177;
    const long ARABIC_CHARSET          = 178;
    const long GREEK_CHARSET           = 161;
    const long TURKISH_CHARSET         = 162;
    const long THAI_CHARSET            = 222;
    const long EASTEUROPE_CHARSET      = 238;
    const long RUSSIAN_CHARSET         = 204;

    const long MAC_CHARSET             = 77;
    const long BALTIC_CHARSET          = 186;

    const UINT FS_LATIN1         = 0x00000001;
    const UINT FS_LATIN2         = 0x00000002;
    const UINT FS_CYRILLIC       = 0x00000004;
    const UINT FS_GREEK          = 0x00000008;
    const UINT FS_TURKISH        = 0x00000010;
    const UINT FS_HEBREW         = 0x00000020;
    const UINT FS_ARABIC         = 0x00000040;
    const UINT FS_BALTIC         = 0x00000080;
    const UINT FS_THAI           = 0x00010000;
    const UINT FS_JISJAPAN       = 0x00020000;
    const UINT FS_CHINESESIMP    = 0x00040000;
    const UINT FS_WANSUNG        = 0x00080000;
    const UINT FS_CHINESETRAD    = 0x00100000;
    const UINT FS_JOHAB          = 0x00200000;
    const UINT FS_SYMBOL         = 0x80000000;

    // Font Families 
    const UINT FF_DONTCARE       = 0x00; // (0<<4)  Don't care or don't know.
    const UINT FF_ROMAN          = 0x14; // (1<<4)  Variable stroke width, serifed.
                                         //         Times Roman, Century Schoolbook, etc.
    const UINT FF_SWISS          = 0x20; // (2<<4)  Variable stroke width, sans-serifed.
                                         //         Helvetica, Swiss, etc.
    const UINT FF_MODERN         = 0x30; // (3<<4)  Constant stroke width, serifed or sans-serifed.
                                         //         Pica, Elite, Courier, etc.
    const UINT FF_SCRIPT         = 0x40; // (4<<4)  Cursive, etc.
    const UINT FF_DECORATIVE     = 0x50; // (5<<4)  Old English, etc.

    // Font Weights 
    const long FW_DONTCARE         = 0;
    const long FW_THIN             = 100;
    const long FW_EXTRALIGHT       = 200;
    const long FW_LIGHT            = 300;
    const long FW_NORMAL           = 400;
    const long FW_MEDIUM           = 500;
    const long FW_SEMIBOLD         = 600;
    const long FW_BOLD             = 700;
    const long FW_EXTRABOLD        = 800;
    const long FW_HEAVY            = 900;

    const long FW_ULTRALIGHT       = 200; // FW_EXTRALIGHT
    const long FW_REGULAR          = 400; // FW_NORMAL
    const long FW_DEMIBOLD         = 600; // FW_SEMIBOLD
    const long FW_ULTRABOLD        = 800; // FW_EXTRABOLD
    const long FW_BLACK            = 900; // FW_HEAVY

    const long PANOSE_COUNT                = 10;
    const long PAN_FAMILYTYPE_INDEX        = 0;
    const long PAN_SERIFSTYLE_INDEX        = 1;
    const long PAN_WEIGHT_INDEX            = 2;
    const long PAN_PROPORTION_INDEX        = 3;
    const long PAN_CONTRAST_INDEX          = 4;
    const long PAN_STROKEVARIATION_INDEX   = 5;
    const long PAN_ARMSTYLE_INDEX          = 6;
    const long PAN_LETTERFORM_INDEX        = 7;
    const long PAN_MIDLINE_INDEX           = 8;
    const long PAN_XHEIGHT_INDEX           = 9;

    const long PAN_CULTURE_LATIN           = 0;

    const long PAN_ANY                         = 0;  // Any
    const long PAN_NO_FIT                      = 1;  // No Fit

    const long PAN_FAMILY_TEXT_DISPLAY         = 2;  // Text and Display
    const long PAN_FAMILY_SCRIPT               = 3;  // Script
    const long PAN_FAMILY_DECORATIVE           = 4;  // Decorative
    const long PAN_FAMILY_PICTORIAL            = 5;  // Pictorial

    const long PAN_SERIF_COVE                  = 2;  // Cove
    const long PAN_SERIF_OBTUSE_COVE           = 3;  // Obtuse Cove
    const long PAN_SERIF_SQUARE_COVE           = 4;  // Square Cove
    const long PAN_SERIF_OBTUSE_SQUARE_COVE    = 5;  // Obtuse Square Cove
    const long PAN_SERIF_SQUARE                = 6;  // Square
    const long PAN_SERIF_THIN                  = 7;  // Thin
    const long PAN_SERIF_BONE                  = 8;  // Bone
    const long PAN_SERIF_EXAGGERATED           = 9;  // Exaggerated
    const long PAN_SERIF_TRIANGLE              = 10; // Triangle
    const long PAN_SERIF_NORMAL_SANS           = 11; // Normal Sans
    const long PAN_SERIF_OBTUSE_SANS           = 12; // Obtuse Sans
    const long PAN_SERIF_PERP_SANS             = 13; // Prep Sans
    const long PAN_SERIF_FLARED                = 14; // Flared
    const long PAN_SERIF_ROUNDED               = 15; // Rounded

    const long PAN_WEIGHT_VERY_LIGHT           = 2;  // Very Light
    const long PAN_WEIGHT_LIGHT                = 3;  // Light
    const long PAN_WEIGHT_THIN                 = 4;  // Thin
    const long PAN_WEIGHT_BOOK                 = 5;  // Book
    const long PAN_WEIGHT_MEDIUM               = 6;  // Medium
    const long PAN_WEIGHT_DEMI                 = 7;  // Demi
    const long PAN_WEIGHT_BOLD                 = 8;  // Bold
    const long PAN_WEIGHT_HEAVY                = 9;  // Heavy
    const long PAN_WEIGHT_NORD                 = 11; // Nord

    const long PAN_PROP_OLD_STYLE              = 2;  // Old Style
    const long PAN_PROP_MODERN                 = 3;  // Modern
    const long PAN_PROP_EVEN_WIDTH             = 4;  // Even Width
    const long PAN_PROP_EXPANDED               = 5;  // Expanded
    const long PAN_PROP_CONDENSED              = 6;  // Condensed
    const long PAN_PROP_VERY_EXPANDED          = 7;  // Very Expanded
    const long PAN_PROP_VERY_CONDENSED         = 8;  // Very Condensed
    const long PAN_PROP_MONOSPACED             = 9;  // Monospaced

    const long PAN_CONTRAST_NONE               = 2;  // None
    const long PAN_CONTRAST_VERY_LOW           = 3;  // Very Low
    const long PAN_CONTRAST_LOW                = 4;  // Low
    const long PAN_CONTRAST_MEDIUM_LOW         = 5;  // Medium Low
    const long PAN_CONTRAST_MEDIUM             = 6;  // Medium
    const long PAN_CONTRAST_MEDIUM_HIGH        = 7;  // Mediim High
    const long PAN_CONTRAST_HIGH               = 8;  // High
    const long PAN_CONTRAST_VERY_HIGH          = 9;  // Very High

    const long PAN_STROKE_GRADUAL_DIAG         = 2;  // Gradual/Diagonal
    const long PAN_STROKE_GRADUAL_TRAN         = 3;  // Gradual/Transitional
    const long PAN_STROKE_GRADUAL_VERT         = 4;  // Gradual/Vertical
    const long PAN_STROKE_GRADUAL_HORZ         = 5;  // Gradual/Horizontal
    const long PAN_STROKE_RAPID_VERT           = 6;  // Rapid/Vertical
    const long PAN_STROKE_RAPID_HORZ           = 7;  // Rapid/Horizontal
    const long PAN_STROKE_INSTANT_VERT         = 8;  // Instant/Vertical

    const long PAN_STRAIGHT_ARMS_HORZ          = 2;  // Straight Arms/Horizontal
    const long PAN_STRAIGHT_ARMS_WEDGE         = 3;  // Straight Arms/Wedge
    const long PAN_STRAIGHT_ARMS_VERT          = 4;  // Straight Arms/Vertical
    const long PAN_STRAIGHT_ARMS_SINGLE_SERIF  = 5;  // Straight Arms/Single-Serif
    const long PAN_STRAIGHT_ARMS_DOUBLE_SERIF  = 6;  // Straight Arms/Double-Serif
    const long PAN_BENT_ARMS_HORZ              = 7;  // Non-Straight Arms/Horizontal
    const long PAN_BENT_ARMS_WEDGE             = 8;  // Non-Straight Arms/Wedge
    const long PAN_BENT_ARMS_VERT              = 9;  // Non-Straight Arms/Vertical
    const long PAN_BENT_ARMS_SINGLE_SERIF      = 10; // Non-Straight Arms/Single-Serif
    const long PAN_BENT_ARMS_DOUBLE_SERIF      = 11; // Non-Straight Arms/Double-Serif

    const long PAN_LETT_NORMAL_CONTACT         = 2;  // Normal/Contact
    const long PAN_LETT_NORMAL_WEIGHTED        = 3;  // Normal/Weighted
    const long PAN_LETT_NORMAL_BOXED           = 4;  // Normal/Boxed
    const long PAN_LETT_NORMAL_FLATTENED       = 5;  // Normal/Flattened
    const long PAN_LETT_NORMAL_ROUNDED         = 6;  // Normal/Rounded
    const long PAN_LETT_NORMAL_OFF_CENTER      = 7;  // Normal/Off Center
    const long PAN_LETT_NORMAL_SQUARE          = 8;  // Normal/Square
    const long PAN_LETT_OBLIQUE_CONTACT        = 9;  // Oblique/Contact
    const long PAN_LETT_OBLIQUE_FLATTENED      = 12; // Oblique/Flattened
    const long PAN_LETT_OBLIQUE_ROUNDED        = 13; // Oblique/Rounded
    const long PAN_LETT_OBLIQUE_OFF_CENTER     = 14; // Oblique/Off Center
    const long PAN_LETT_OBLIQUE_SQUARE         = 15; // Oblique/Square

    const long PAN_MIDLINE_STANDARD_TRIMMED    = 2;  // Standard/Trimmed
    const long PAN_MIDLINE_STANDARD_POINTED    = 3;  // Standard/Pointed
    const long PAN_MIDLINE_STANDARD_SERIFED    = 4;  // Standard/Serifed
    const long PAN_MIDLINE_HIGH_TRIMMED        = 5;  // High/Trimmed
    const long PAN_MIDLINE_HIGH_POINTED        = 6;  // High/Pointed
    const long PAN_MIDLINE_HIGH_SERIFED        = 7;  // High/Serifed
    const long PAN_MIDLINE_CONSTANT_TRIMMED    = 8;  // Constant/Trimmed
    const long PAN_MIDLINE_CONSTANT_POINTED    = 9;  // Constant/Pointed
    const long PAN_MIDLINE_CONSTANT_SERIFED    = 10; // Constant/Serifed
    const long PAN_MIDLINE_LOW_TRIMMED         = 11; // Low/Trimmed
    const long PAN_MIDLINE_LOW_POINTED         = 12; // Low/Pointed
    const long PAN_MIDLINE_LOW_SERIFED         = 13; // Low/Serifed

    const long PAN_XHEIGHT_CONSTANT_SMALL      = 2;  // Constant/Small
    const long PAN_XHEIGHT_CONSTANT_STD        = 3;  // Constant/Standard
    const long PAN_XHEIGHT_CONSTANT_LARGE      = 4;  // Constant/Large
    const long PAN_XHEIGHT_DUCKING_SMALL       = 5;  // Ducking/Small
    const long PAN_XHEIGHT_DUCKING_STD         = 6;  // Ducking/Standard
    const long PAN_XHEIGHT_DUCKING_LARGE       = 7;  // Ducking/Large


    const long ELF_VENDOR_SIZE     = 4;
                                  
    const long ELF_VERSION         = 0;
    const long ELF_CULTURE_LATIN   = 0;

    // EnumFonts Masks 
    const UINT RASTER_FONTTYPE     = 0x0001;
    const UINT DEVICE_FONTTYPE     = 0x0002;
    const UINT TRUETYPE_FONTTYPE   = 0x0004;

    // palette entry flags 

    const UINT PC_RESERVED     = 0x01;    // palette index used for animation
    const UINT PC_EXPLICIT     = 0x02;    // palette index is explicit to device
    const UINT PC_NOCOLLAPSE   = 0x04;    // do not match color to system palette

    // Background Modes
    const long TRANSPARENT    = 1;
    const long OPAQUE         = 2;
    const long NEWTRANSPARENT = 3;

    // Graphics Modes 

    const long GM_COMPATIBLE       = 1;
    const long GM_ADVANCED         = 2;
    const long GM_LAST             = 2;

    // PolyDraw and GetPath point types 
    const UINT PT_CLOSEFIGURE      = 0x01;
    const UINT PT_LINETO           = 0x02;
    const UINT PT_BEZIERTO         = 0x04;
    const UINT PT_MOVETO           = 0x06;

    // Mapping Modes 
    const long MM_TEXT             = 1;
    const long MM_LOMETRIC         = 2;
    const long MM_HIMETRIC         = 3;
    const long MM_LOENGLISH        = 4;
    const long MM_HIENGLISH        = 5;
    const long MM_TWIPS            = 6;
    const long MM_ISOTROPIC        = 7;
    const long MM_ANISOTROPIC      = 8;

    // Min and Max Mapping Mode values 
    const long MM_MIN              = 1; // MM_TEXT
    const long MM_MAX              = 8; // MM_ANISOTROPIC
    const long MM_MAX_FIXEDSCALE   = 6; // MM_TWIPS

    // Coordinate Modes 
    const long ABSOLUTE            = 1;
    const long RELATIVE            = 2;

    // Stock Logical Objects 
    const long WHITE_BRUSH         = 0;
    const long LTGRAY_BRUSH        = 1;
    const long GRAY_BRUSH          = 2;
    const long DKGRAY_BRUSH        = 3;
    const long BLACK_BRUSH         = 4;
    const long NULL_BRUSH          = 5;
    const long HOLLOW_BRUSH        = 5; // NULL_BRUSH
    const long WHITE_PEN           = 6;
    const long BLACK_PEN           = 7;
    const long NULL_PEN            = 8;
    const long OEM_FIXED_FONT      = 10;
    const long ANSI_FIXED_FONT     = 11;
    const long ANSI_VAR_FONT       = 12;
    const long SYSTEM_FONT         = 13;
    const long DEVICE_DEFAULT_FONT = 14;
    const long DEFAULT_PALETTE     = 15;
    const long SYSTEM_FIXED_FONT   = 16;
    const long DEFAULT_GUI_FONT    = 17;

    const UINT CLR_INVALID     = 0xFFFFFFFF;

    // Brush Styles 
    const long BS_SOLID            = 0;
    const long BS_NULL             = 1;
    const long BS_HOLLOW           = 1; // BS_NULL
    const long BS_HATCHED          = 2;
    const long BS_PATTERN          = 3;
    const long BS_INDEXED          = 4;
    const long BS_DIBPATTERN       = 5;
    const long BS_DIBPATTERNPT     = 6;
    const long BS_PATTERN8X8       = 7;
    const long BS_DIBPATTERN8X8    = 8;

    // Hatch Styles 
    const long HS_HORIZONTAL       = 0;      // -----
    const long HS_VERTICAL         = 1;      // |||||
    const long HS_FDIAGONAL        = 2;      //
    const long HS_BDIAGONAL        = 3;      //
    const long HS_CROSS            = 4;      // +++++
    const long HS_DIAGCROSS        = 5;      // xxxxx

    // Pen Styles 
    const long PS_SOLID            = 0;
    const long PS_DASH             = 1;      // -------
    const long PS_DOT              = 2;      // .......
    const long PS_DASHDOT          = 3;      // _._._._
    const long PS_DASHDOTDOT       = 4;      // _.._.._
    const long PS_NULL             = 5;
    const long PS_INSIDEFRAME      = 6;
    const long PS_USERSTYLE        = 7;
    const long PS_ALTERNATE        = 8;
    const long PS_STYLE_MASK       = 0x0000000F;

    const long PS_ENDCAP_ROUND     = 0x00000000;
    const long PS_ENDCAP_SQUARE    = 0x00000100;
    const long PS_ENDCAP_FLAT      = 0x00000200;
    const long PS_ENDCAP_MASK      = 0x00000F00;

    const long PS_JOIN_ROUND       = 0x00000000;
    const long PS_JOIN_BEVEL       = 0x00001000;
    const long PS_JOIN_MITER       = 0x00002000;
    const long PS_JOIN_MASK        = 0x0000F000;

    const long PS_COSMETIC         = 0x00000000;
    const long PS_GEOMETRIC        = 0x00010000;
    const long PS_TYPE_MASK        = 0x000F0000;

    const UINT AD_COUNTERCLOCKWISE = 1;
    const UINT AD_CLOCKWISE        = 2;

    // Device Parameters for GetDeviceCaps() 
    const long DRIVERVERSION = 0;
    const long TECHNOLOGY    = 2;
    const long HORZSIZE      = 4;
    const long VERTSIZE      = 6;
    const long HORZRES       = 8;
    const long VERTRES       = 10;
    const long BITSPIXEL     = 12;
    const long PLANES        = 14;
    const long NUMBRUSHES    = 16;
    const long NUMPENS       = 18;
    const long NUMMARKERS    = 20;
    const long NUMFONTS      = 22;
    const long NUMCOLORS     = 24;
    const long PDEVICESIZE   = 26;
    const long CURVECAPS     = 28;
    const long LINECAPS      = 30;
    const long POLYGONALCAPS = 32;
    const long TEXTCAPS    = 34;
    const long CLIPCAPS      = 36;
    const long RASTERCAPS    = 38;
    const long ASPECTX       = 40;
    const long ASPECTY       = 42;
    const long ASPECTXY      = 44;
    const long LOGPIXELSX    = 88;
    const long LOGPIXELSY    = 90;
    const long SIZEPALETTE   = 104;
    const long NUMRESERVED   = 106;
    const long COLORRES      = 108;

    // GetDeviceCaps() return value masks

    // Printing related DeviceCaps. These replace the appropriate Escapes

    const long PHYSICALWIDTH   = 110; // Physical Width in device units
    const long PHYSICALHEIGHT  = 111; // Physical Height in device units
    const long PHYSICALOFFSETX = 112; // Physical Printable Area x margin
    const long PHYSICALOFFSETY = 113; // Physical Printable Area y margin
    const long SCALINGFACTORX  = 114; // Scaling factor x
    const long SCALINGFACTORY  = 115; // Scaling factor y

    // Display driver specific

    const long VREFRESH        = 116; // Current vertical refresh rate of the
                                     // display device (for displays only) in Hz
    const long DESKTOPVERTRES  = 117; // Horizontal width of entire desktop in
                                     // pixels
    const long DESKTOPHORZRES  = 118; // Vertical height of entire desktop in
                                     // pixels
    const long BLTALIGNMENT    = 119; // Preferred blt alignment

    // TECHNOLOGY
    const long DT_PLOTTER        = 0;
    const long DT_RASDISPLAY     = 1;
    const long DT_RASPRINTER     = 2;
    const long DT_RASCAMERA      = 3;
    const long DT_CHARSTREAM     = 4;
    const long DT_METAFILE       = 5;
    const long DT_DISPFILE       = 6;

    // Curve Capabilities 
    const UINT CC_NONE          = 0x0000;
    const UINT CC_CIRCLES       = 0x0001;
    const UINT CC_PIE           = 0x0002;
    const UINT CC_CHORD         = 0x0004;
    const UINT CC_ELLIPSES      = 0x0008;
    const UINT CC_WIDE          = 0x0010;
    const UINT CC_STYLED        = 0x0020;
    const UINT CC_WIDESTYLED    = 0x0040;
    const UINT CC_INTERIORS     = 0x0080;
    const UINT CC_ROUNDRECT     = 0x0100;

    // Line Capabilities 
    const UINT LC_NONE          = 0x0000;
    const UINT LC_POLYLINE      = 0x0002;
    const UINT LC_MARKER            = 0x0004;
    const UINT LC_POLYMARKER        = 0x0008;
    const UINT LC_WIDE          = 0x0010;
    const UINT LC_STYLED            = 0x0020;
    const UINT LC_WIDESTYLED        = 0x0040;
    const UINT LC_INTERIORS     = 0x0080;

    // Polygonal Capabilities 
    const UINT PC_NONE          = 0x0000;
    const UINT PC_POLYGON       = 0x0001;
    const UINT PC_RECTANGLE     = 0x0002;
    const UINT PC_WINDPOLYGON   = 0x0004;
    const UINT PC_SCANLINE      = 0x0008;
    const UINT PC_WIDE          = 0x0010;
    const UINT PC_STYLED        = 0x0020;
    const UINT PC_WIDESTYLED    = 0x0040;
    const UINT PC_INTERIORS     = 0x0080;
    const long PC_POLYPOLYGON    = 256; // Can do polypolygons
    const long PC_PATHS          = 512; // Can do paths

    // CLIPCAPS
    const UINT CP_NONE          = 0x0000;
    const UINT CP_RECTANGLE     = 0x0001;
    const UINT CP_REGION        = 0x0002;

    // TEXTCAPS
    const UINT TC_OP_CHARACTER  = 0x0001;
    const UINT TC_OP_STROKE     = 0x0002;
    const UINT TC_CP_STROKE     = 0x0004;
    const UINT TC_CR_90         = 0x0008;
    const UINT TC_CR_ANY        = 0x0010;
    const UINT TC_SF_X_YINDEP   = 0x0020;
    const UINT TC_SA_DOUBLE     = 0x0040;
    const UINT TC_SA_INTEGER    = 0x0080;
    const UINT TC_SA_CONTIN     = 0x0100;
    const UINT TC_EA_DOUBLE     = 0x0200;
    const UINT TC_IA_ABLE       = 0x0400;
    const UINT TC_UA_ABLE       = 0x0800;
    const UINT TC_SO_ABLE       = 0x1000;
    const UINT TC_RA_ABLE       = 0x2000;
    const UINT TC_VA_ABLE       = 0x4000;
    const UINT TC_RESERVED      = 0x00008000;
    const UINT TC_SCROLLBLT     = 0x00010000;  // Don't do text scroll with blt

    // RASTERCAPS
    const UINT RC_BITBLT        = 0x0001;
    const UINT RC_BANDING       = 0x0002;
    const UINT RC_SCALING       = 0x0004;
    const UINT RC_BITMAP64      = 0x0008;
    const UINT RC_GDI20_OUTPUT  = 0x0010;
    const UINT RC_GDI20_STATE   = 0x0020;
    const UINT RC_SAVEBITMAP    = 0x0040;
    const UINT RC_DI_BITMAP     = 0x0080;
    const UINT RC_PALETTE       = 0x0100;
    const UINT RC_DIBTODEV      = 0x0200;
    const UINT RC_BIGFONT       = 0x0400;
    const UINT RC_STRETCHBLT    = 0x0800;
    const UINT RC_FLOODFILL     = 0x1000;
    const UINT RC_STRETCHDIB    = 0x2000;
    const UINT RC_OP_DX_OUTPUT  = 0x4000;
    const UINT RC_DEVBITS		= 0x8000;

    const long CAPS1         = 94;
    const UINT C1_TRANSPARENT   = 0x0001;   // new raster cap


    // DIB color table identifiers 

    const long DIB_RGB_COLORS    = 0; // color table in RGBs
    const long DIB_PAL_COLORS    = 1; // color table in palette indices

    // constants for Get/SetSystemPaletteUse() 

    const long SYSPAL_ERROR      = 0;
    const long SYSPAL_STATIC     = 1;
    const long SYSPAL_NOSTATIC   = 2;

    // constants for CreateDIBitmap 
    const UINT CBM_INIT        = 0x04;   // initialize bitmap

    // ExtFloodFill style flags
    const long FLOODFILLBORDER   = 0;
    const long FLOODFILLSURFACE  = 1;

    // size of a device name string 
    const long CCHDEVICENAME = 32;

    // size of a form name string 
    const long CCHFORMNAME = 32;

    // current version of specification 
    const UINT DM_SPECVERSION = 0x0400;

    // field selection bits 
    const UINT DM_ORIENTATION      = 0x00000001;
    const UINT DM_PAPERSIZE        = 0x00000002;
    const UINT DM_PAPERLENGTH      = 0x00000004;
    const UINT DM_PAPERWIDTH       = 0x00000008;
    const UINT DM_SCALE            = 0x00000010;
    const UINT DM_COPIES           = 0x00000100;
    const UINT DM_DEFAULTSOURCE    = 0x00000200;
    const UINT DM_PRINTQUALITY     = 0x00000400;
    const UINT DM_COLOR            = 0x00000800;
    const UINT DM_DUPLEX           = 0x00001000;
    const UINT DM_YRESOLUTION      = 0x00002000;
    const UINT DM_TTOPTION         = 0x00004000;
    const UINT DM_COLLATE          = 0x00008000;
    const UINT DM_FORMNAME         = 0x00010000;
    const UINT DM_LOGPIXELS        = 0x00020000;
    const UINT DM_BITSPERPEL       = 0x00040000;
    const UINT DM_PELSWIDTH        = 0x00080000;
    const UINT DM_PELSHEIGHT       = 0x00100000;
    const UINT DM_DISPLAYFLAGS     = 0x00200000;
    const UINT DM_DISPLAYFREQUENCY = 0x00400000;
    const UINT DM_ICMMETHOD        = 0x00800000;
    const UINT DM_ICMINTENT        = 0x01000000;
    const UINT DM_MEDIATYPE        = 0x02000000;
    const UINT DM_DITHERTYPE       = 0x04000000;

    // orientation selections 
    const long DMORIENT_PORTRAIT     = 1;
    const long DMORIENT_LANDSCAPE    = 2;

    // paper selections 
    const long DMPAPER_FIRST                = 1;  // DMPAPER_LETTER
    const long DMPAPER_LETTER               = 1;  // Letter 8 1/2 x 11 in
    const long DMPAPER_LETTERSMALL          = 2;  // Letter Small 8 1/2 x 11 in
    const long DMPAPER_TABLOID              = 3;  // Tabloid 11 x 17 in
    const long DMPAPER_LEDGER               = 4;  // Ledger 17 x 11 in
    const long DMPAPER_LEGAL                = 5;  // Legal 8 1/2 x 14 in
    const long DMPAPER_STATEMENT            = 6;  // Statement 5 1/2 x 8 1/2 in
    const long DMPAPER_EXECUTIVE            = 7;  // Executive 7 1/4 x 10 1/2 in
    const long DMPAPER_A3                   = 8;  // A3 297 x 420 mm
    const long DMPAPER_A4                   = 9;  // A4 210 x 297 mm
    const long DMPAPER_A4SMALL             = 10;  // A4 Small 210 x 297 mm
    const long DMPAPER_A5                  = 11;  // A5 148 x 210 mm
    const long DMPAPER_B4                  = 12;  // B4 (JIS) 250 x 354
    const long DMPAPER_B5                  = 13;  // B5 (JIS) 182 x 257 mm
    const long DMPAPER_FOLIO               = 14;  // Folio 8 1/2 x 13 in
    const long DMPAPER_QUARTO              = 15;  // Quarto 215 x 275 mm
    const long DMPAPER_10X14               = 16;  // 10x14 in
    const long DMPAPER_11X17               = 17;  // 11x17 in
    const long DMPAPER_NOTE                = 18;  // Note 8 1/2 x 11 in
    const long DMPAPER_ENV_9               = 19;  // Envelope #9 3 7/8 x 8 7/8
    const long DMPAPER_ENV_10              = 20;  // Envelope #10 4 1/8 x 9 1/2
    const long DMPAPER_ENV_11              = 21;  // Envelope #11 4 1/2 x 10 3/8
    const long DMPAPER_ENV_12              = 22;  // Envelope #12 4 \276 x 11
    const long DMPAPER_ENV_14              = 23;  // Envelope #14 5 x 11 1/2
    const long DMPAPER_CSHEET              = 24;  // C size sheet
    const long DMPAPER_DSHEET              = 25;  // D size sheet
    const long DMPAPER_ESHEET              = 26;  // E size sheet
    const long DMPAPER_ENV_DL              = 27;  // Envelope DL 110 x 220mm
    const long DMPAPER_ENV_C5              = 28;  // Envelope C5 162 x 229 mm
    const long DMPAPER_ENV_C3              = 29;  // Envelope C3  324 x 458 mm
    const long DMPAPER_ENV_C4              = 30;  // Envelope C4  229 x 324 mm
    const long DMPAPER_ENV_C6              = 31;  // Envelope C6  114 x 162 mm
    const long DMPAPER_ENV_C65             = 32;  // Envelope C65 114 x 229 mm
    const long DMPAPER_ENV_B4              = 33;  // Envelope B4  250 x 353 mm
    const long DMPAPER_ENV_B5              = 34;  // Envelope B5  176 x 250 mm
    const long DMPAPER_ENV_B6              = 35;  // Envelope B6  176 x 125 mm
    const long DMPAPER_ENV_ITALY           = 36;  // Envelope 110 x 230 mm
    const long DMPAPER_ENV_MONARCH         = 37;  // Envelope Monarch 3.875 x 7.5 in
    const long DMPAPER_ENV_PERSONAL        = 38;  // 6 3/4 Envelope 3 5/8 x 6 1/2 in
    const long DMPAPER_FANFOLD_US          = 39;  // US Std Fanfold 14 7/8 x 11 in
    const long DMPAPER_FANFOLD_STD_GERMAN  = 40;  // German Std Fanfold 8 1/2 x 12 in
    const long DMPAPER_FANFOLD_LGL_GERMAN  = 41;  // German Legal Fanfold 8 1/2 x 13 in
    const long DMPAPER_ISO_B4              = 42;  // B4 (ISO) 250 x 353 mm
    const long DMPAPER_JAPANESE_POSTCARD   = 43;  // Japanese Postcard 100 x 148 mm
    const long DMPAPER_9X11                = 44;  // 9 x 11 in
    const long DMPAPER_10X11               = 45;  // 10 x 11 in
    const long DMPAPER_15X11               = 46;  // 15 x 11 in
    const long DMPAPER_ENV_INVITE          = 47;  // Envelope Invite 220 x 220 mm
    const long DMPAPER_RESERVED_48         = 48;  // RESERVED--DO NOT USE
    const long DMPAPER_RESERVED_49         = 49;  // RESERVED--DO NOT USE
    const long DMPAPER_LETTER_EXTRA        = 50;  // Letter Extra 9 \275 x 12 in
    const long DMPAPER_LEGAL_EXTRA         = 51;  // Legal Extra 9 \275 x 15 in
    const long DMPAPER_TABLOID_EXTRA       = 52;  // Tabloid Extra 11.69 x 18 in
    const long DMPAPER_A4_EXTRA            = 53;  // A4 Extra 9.27 x 12.69 in
    const long DMPAPER_LETTER_TRANSVERSE   = 54;  // Letter Transverse 8 \275 x 11 in
    const long DMPAPER_A4_TRANSVERSE       = 55;  // A4 Transverse 210 x 297 mm
    const long DMPAPER_LETTER_EXTRA_TRANSVERSE = 56; // Letter Extra Transverse 9\275 x 12 in
    const long DMPAPER_A_PLUS              = 57;  // SuperA/SuperA/A4 227 x 356 mm
    const long DMPAPER_B_PLUS              = 58;  // SuperB/SuperB/A3 305 x 487 mm
    const long DMPAPER_LETTER_PLUS         = 59;  // Letter Plus 8.5 x 12.69 in
    const long DMPAPER_A4_PLUS             = 60;  // A4 Plus 210 x 330 mm
    const long DMPAPER_A5_TRANSVERSE       = 61;  // A5 Transverse 148 x 210 mm
    const long DMPAPER_B5_TRANSVERSE       = 62;  // B5 (JIS) Transverse 182 x 257 mm
    const long DMPAPER_A3_EXTRA            = 63;  // A3 Extra 322 x 445 mm
    const long DMPAPER_A5_EXTRA            = 64;  // A5 Extra 174 x 235 mm
    const long DMPAPER_B5_EXTRA            = 65;  // B5 (ISO) Extra 201 x 276 mm
    const long DMPAPER_A2                  = 66;  // A2 420 x 594 mm
    const long DMPAPER_A3_TRANSVERSE       = 67;  // A3 Transverse 297 x 420 mm
    const long DMPAPER_A3_EXTRA_TRANSVERSE = 68;  // A3 Extra Transverse 322 x 445 mm

    const long DMPAPER_USER                = 256;

    // bin selections 
    const long DMBIN_FIRST         = 1; // DMBIN_UPPER
    const long DMBIN_UPPER         = 1;
    const long DMBIN_ONLYONE       = 1;
    const long DMBIN_LOWER         = 2;
    const long DMBIN_MIDDLE        = 3;
    const long DMBIN_MANUAL        = 4;
    const long DMBIN_ENVELOPE      = 5;
    const long DMBIN_ENVMANUAL     = 6;
    const long DMBIN_AUTO          = 7;
    const long DMBIN_TRACTOR       = 8;
    const long DMBIN_SMALLFMT      = 9;
    const long DMBIN_LARGEFMT      = 10;
    const long DMBIN_LARGECAPACITY = 11;
    const long DMBIN_CASSETTE      = 14;
    const long DMBIN_FORMSOURCE    = 15;
    const long DMBIN_LAST          = 15; // DMBIN_FORMSOURCE

    const long DMBIN_USER          = 256;   // device specific bins start here

    // prlong qualities
    const long DMRES_DRAFT         = -1;
    const long DMRES_LOW           = -2;
    const long DMRES_MEDIUM        = -3;
    const long DMRES_HIGH          = -4;

    // color enable/disable for color printers 
    const long DMCOLOR_MONOCHROME  = 1;
    const long DMCOLOR_COLOR       = 2;

    // duplex enable
    const long DMDUP_SIMPLEX    = 1;
    const long DMDUP_VERTICAL   = 2;
    const long DMDUP_HORIZONTAL = 3;

    // TrueType options 
    const long DMTT_BITMAP     = 1;       // prlong TT fonts as graphics
    const long DMTT_DOWNLOAD   = 2;       // download TT fonts as soft fonts
    const long DMTT_SUBDEV     = 3;       // substitute device fonts for TT fonts
    const long DMTT_DOWNLOAD_OUTLINE = 4; // download TT fonts as outline soft fonts

    // Collation selections 
    const long DMCOLLATE_FALSE  = 0;
    const long DMCOLLATE_TRUE   = 1;

    // DEVMODE dmDisplayFlags flags 

    const UINT DM_GRAYSCALE  = 0x00000001;
    const UINT DM_INTERLACED = 0x00000002;

    // ICM methods 
    const long DMICMMETHOD_NONE    = 1;   // ICM disabled
    const long DMICMMETHOD_SYSTEM  = 2;   // ICM handled by system
    const long DMICMMETHOD_DRIVER  = 3;   // ICM handled by driver
    const long DMICMMETHOD_DEVICE  = 4;   // ICM handled by device

    const long DMICMMETHOD_USER  = 256;   // Device-specific methods start here

    // ICM Intents 
    const long DMICM_SATURATE      = 1;   // Maximize color saturation
    const long DMICM_CONTRAST      = 2;   // Maximize color contrast
    const long DMICM_COLORMETRIC   = 3;   // Use specific color metric

    const long DMICM_USER        = 256;   // Device-specific intents start here

    // Media types 

    const long DMMEDIA_STANDARD      = 1;   // Standard paper
    const long DMMEDIA_TRANSPARENCY  = 2;   // Transparency
    const long DMMEDIA_GLOSSY        = 3;   // Glossy paper

    const long DMMEDIA_USER        = 256;   // Device-specific media start here

    // Dither types 
    const long DMDITHER_NONE         = 1;   // No dithering
    const long DMDITHER_COARSE       = 2;   // Dither with a coarse brush
    const long DMDITHER_FINE         = 3;   // Dither with a fine brush
    const long DMDITHER_LINEART      = 4;   // LineArt dithering

    const long DMDITHER_ERRORDIFFUSION = 5; // LineArt dithering
    const long DMDITHER_RESERVED6    = 6;   // LineArt dithering
    const long DMDITHER_RESERVED7    = 7;   // LineArt dithering
    const long DMDITHER_RESERVED8    = 8;   // LineArt dithering
    const long DMDITHER_RESERVED9    = 9;   // LineArt dithering

    const long DMDITHER_GRAYSCALE    = 10;   // Device does grayscaling

    const long DMDITHER_USER         = 256;   // Device-specific dithers start here

    // GetRegionData/ExtCreateRegion 

    const long RDH_RECTANGLES        = 1;

    //  GetGlyphOutline constants

    const long GGO_METRICS           = 0;
    const long GGO_BITMAP            = 1;
    const long GGO_NATIVE            = 2;

    const long  GGO_GRAY2_BITMAP     = 4;
    const long  GGO_GRAY4_BITMAP     = 5;
    const long  GGO_GRAY8_BITMAP     = 6;
    const long  GGO_GLYPH_INDEX      = 0x0080;

    const long TT_POLYGON_TYPE       = 24;

    const long TT_PRIM_LINE          = 1;
    const long TT_PRIM_QSPLINE       = 2;

    const UINT GCP_DBCS            = 0x0001;
    const UINT GCP_REORDER         = 0x0002;
    const UINT GCP_USEKERNING      = 0x0008;
    const UINT GCP_GLYPHSHAPE      = 0x0010;
    const UINT GCP_LIGATE          = 0x0020;
    //const UINT GCP_GLYPHINDEXING = 0x0080;
    const UINT GCP_DIACRITIC       = 0x0100;
    const UINT GCP_KASHIDA         = 0x0400;
    const UINT GCP_ERROR           = 0x8000;
    const UINT FLI_MASK            = 0x103B;

    const UINT GCP_JUSTIFY         = 0x00010000;
    //const UINT GCP_NODIACRITICS  = 0x00020000;
    const UINT FLI_GLYPHS          = 0x00040000;
    const UINT GCP_CLASSIN         = 0x00080000;
    const UINT GCP_MAXEXTENT       = 0x00100000;
    const UINT GCP_JUSTIFYIN       = 0x00200000;
    const UINT GCP_DISPLAYZWG      = 0x00400000;
    const UINT GCP_SYMSWAPOFF      = 0x00800000;
    const UINT GCP_NUMERICOVERRIDE = 0x01000000;
    const UINT GCP_NEUTRALOVERRIDE = 0x02000000;
    const UINT GCP_NUMERICSLATIN   = 0x04000000;
    const UINT GCP_NUMERICSLOCAL   = 0x08000000;

    const UINT GCPCLASS_LATIN                  = 1;
    const UINT GCPCLASS_HEBREW                 = 2;
    const UINT GCPCLASS_ARABIC                 = 2;
    const UINT GCPCLASS_NEUTRAL                = 3;
    const UINT GCPCLASS_LOCALNUMBER            = 4;
    const UINT GCPCLASS_LATINNUMBER            = 5;
    const UINT GCPCLASS_LATINNUMERICTERMINATOR = 6;
    const UINT GCPCLASS_LATINNUMERICSEPARATOR  = 7;
    const UINT GCPCLASS_NUMERICSEPARATOR       = 8;
    const UINT GCPCLASS_PREBOUNDLTR         = 0x80;
    const UINT GCPCLASS_PREBOUNDRTL         = 0x40;
    const UINT GCPCLASS_POSTBOUNDLTR        = 0x20;
    const UINT GCPCLASS_POSTBOUNDRTL        = 0x10;

    const UINT GCPGLYPH_LINKBEFORE          = 0x8000;
    const UINT GCPGLYPH_LINKAFTER           = 0x4000;

    // bits defined in wFlags of RASTERIZER_STATUS 
    const long TT_AVAILABLE    = 0x0001;
    const long TT_ENABLED      = 0x0002;

    // pixel types 
    const UINT PFD_TYPE_RGBA        = 0;
    const UINT PFD_TYPE_COLORINDEX  = 1;

    // layer types 
    const UINT PFD_MAIN_PLANE       = 0;
    const UINT PFD_OVERLAY_PLANE    = 1;
    const UINT PFD_UNDERLAY_PLANE   = -1;

    // PIXELFORMATDESCRIPTOR flags 
    const UINT PFD_DOUBLEBUFFER            = 0x00000001;
    const UINT PFD_STEREO                  = 0x00000002;
    const UINT PFD_DRAW_TO_WINDOW          = 0x00000004;
    const UINT PFD_DRAW_TO_BITMAP          = 0x00000008;
    const UINT PFD_SUPPORT_GDI             = 0x00000010;
    const UINT PFD_SUPPORT_OPENGL          = 0x00000020;
    const UINT PFD_GENERIC_FORMAT          = 0x00000040;
    const UINT PFD_NEED_PALETTE            = 0x00000080;
    const UINT PFD_NEED_SYSTEM_PALETTE     = 0x00000100;
    const UINT PFD_SWAP_EXCHANGE           = 0x00000200;
    const UINT PFD_SWAP_COPY               = 0x00000400;

    // PIXELFORMATDESCRIPTOR flags for use in ChoosePixelFormat only 
    const UINT PFD_DOUBLEBUFFER_DONTCARE   = 0x40000000;
    const UINT PFD_STEREO_DONTCARE         = 0x80000000;

    // Mode selections for the device mode function 
    const long DM_UPDATE           = 1;
    const long DM_COPY             = 2;
    const long DM_PROMPT           = 4;
    const long DM_MODIFY           = 8;

    const long DM_IN_BUFFER        = 8; // DM_MODIFY
    const long DM_IN_PROMPT        = 4; // DM_PROMPT
    const long DM_OUT_BUFFER       = 2; // DM_COPY
    const long DM_OUT_DEFAULT      = 1; // DM_UPDATE

    // device capabilities indices 
    const long DC_FIELDS           = 1;
    const long DC_PAPERS           = 2;
    const long DC_PAPERSIZE        = 3;
    const long DC_MINEXTENT        = 4;
    const long DC_MAXEXTENT        = 5;
    const long DC_BINS             = 6;
    const long DC_DUPLEX           = 7;
    const long DC_SIZE             = 8;
    const long DC_EXTRA            = 9;
    const long DC_VERSION          = 10;
    const long DC_DRIVER           = 11;
    const long DC_BINNAMES         = 12;
    const long DC_ENUMRESOLUTIONS  = 13;
    const long DC_FILEDEPENDENCIES = 14;
    const long DC_TRUETYPE         = 15;
    const long DC_PAPERNAMES       = 16;
    const long DC_ORIENTATION      = 17;
    const long DC_COPIES           = 18;
    const long DC_BINADJUST            = 19;
    const long DC_EMF_COMPLIANT        = 20;
    const long DC_DATATYPE_PRODUCED    = 21;
    const long DC_MANUFACTURER     = 23;
    const long DC_MODEL            = 24;

    // bit fields of the return value (DWORD) for DC_TRUETYPE 
    const UINT DCTT_BITMAP             = 0x0000001;
    const UINT DCTT_DOWNLOAD           = 0x0000002;
    const UINT DCTT_SUBDEV             = 0x0000004;
    const UINT DCTT_DOWNLOAD_OUTLINE   = 0x0000008;

    // return values for DC_BINADJUST 
    const UINT DCBA_FACEUPNONE         = 0x0000;
    const UINT DCBA_FACEUPCENTER       = 0x0001;
    const UINT DCBA_FACEUPLEFT         = 0x0002;
    const UINT DCBA_FACEUPRIGHT        = 0x0003;
    const UINT DCBA_FACEDOWNNONE       = 0x0100;
    const UINT DCBA_FACEDOWNCENTER     = 0x0101;
    const UINT DCBA_FACEDOWNLEFT       = 0x0102;
    const UINT DCBA_FACEDOWNRIGHT      = 0x0103;

    const long WGL_FONT_LINES      = 0;
    const long WGL_FONT_POLYGONS   = 1;

    const UINT GDICOMMENT_IDENTIFIER           = 0x43494447;
    const UINT GDICOMMENT_WINDOWS_METAFILE     = 0x80000001;
    const UINT GDICOMMENT_BEGINGROUP           = 0x00000002;
    const UINT GDICOMMENT_ENDGROUP             = 0x00000003;
    const UINT GDICOMMENT_MULTIFORMATS         = 0x40000004;
    const UINT EPS_SIGNATURE                   = 0x46535045;

    // Enhanced metafile constants.

    const UINT ENHMETA_SIGNATURE       = 0x464D4520;

    // Stock object flag used in the object handle index in the enhanced
    // metafile records.
    // E.g. The object handle index (META_STOCK_OBJECT | BLACK_BRUSH)
    // represents the stock object BLACK_BRUSH.

    const UINT ENHMETA_STOCK_OBJECT    = 0x80000000;

    // Enhanced metafile record types.

    const long EMR_HEADER                      = 1;
    const long EMR_POLYBEZIER                  = 2;
    const long EMR_POLYGON                     = 3;
    const long EMR_POLYLINE                    = 4;
    const long EMR_POLYBEZIERTO                = 5;
    const long EMR_POLYLINETO                  = 6;
    const long EMR_POLYPOLYLINE                = 7;
    const long EMR_POLYPOLYGON                 = 8;
    const long EMR_SETWINDOWEXTEX              = 9;
    const long EMR_SETWINDOWORGEX              = 10;
    const long EMR_SETVIEWPORTEXTEX            = 11;
    const long EMR_SETVIEWPORTORGEX            = 12;
    const long EMR_SETBRUSHORGEX               = 13;
    const long EMR_EOF                         = 14;
    const long EMR_SETPIXELV                   = 15;
    const long EMR_SETMAPPERFLAGS              = 16;
    const long EMR_SETMAPMODE                  = 17;
    const long EMR_SETBKMODE                   = 18;
    const long EMR_SETPOLYFILLMODE             = 19;
    const long EMR_SETROP2                     = 20;
    const long EMR_SETSTRETCHBLTMODE           = 21;
    const long EMR_SETTEXTALIGN                = 22;
    const long EMR_SETCOLORADJUSTMENT          = 23;
    const long EMR_SETTEXTCOLOR                = 24;
    const long EMR_SETBKCOLOR                  = 25;
    const long EMR_OFFSETCLIPRGN               = 26;
    const long EMR_MOVETOEX                    = 27;
    const long EMR_SETMETARGN                  = 28;
    const long EMR_EXCLUDECLIPRECT             = 29;
    const long EMR_INTERSECTCLIPRECT           = 30;
    const long EMR_SCALEVIEWPORTEXTEX          = 31;
    const long EMR_SCALEWINDOWEXTEX            = 32;
    const long EMR_SAVEDC                      = 33;
    const long EMR_RESTOREDC                   = 34;
    const long EMR_SETWORLDTRANSFORM           = 35;
    const long EMR_MODIFYWORLDTRANSFORM        = 36;
    const long EMR_SELECTOBJECT                = 37;
    const long EMR_CREATEPEN                   = 38;
    const long EMR_CREATEBRUSHINDIRECT         = 39;
    const long EMR_DELETEOBJECT                = 40;
    const long EMR_ANGLEARC                    = 41;
    const long EMR_ELLIPSE                     = 42;
    const long EMR_RECTANGLE                   = 43;
    const long EMR_ROUNDRECT                   = 44;
    const long EMR_ARC                         = 45;
    const long EMR_CHORD                       = 46;
    const long EMR_PIE                         = 47;
    const long EMR_SELECTPALETTE               = 48;
    const long EMR_CREATEPALETTE               = 49;
    const long EMR_SETPALETTEENTRIES           = 50;
    const long EMR_RESIZEPALETTE               = 51;
    const long EMR_REALIZEPALETTE              = 52;
    const long EMR_EXTFLOODFILL                = 53;
    const long EMR_LINETO                      = 54;
    const long EMR_ARCTO                       = 55;
    const long EMR_POLYDRAW                    = 56;
    const long EMR_SETARCDIRECTION             = 57;
    const long EMR_SETMITERLIMIT               = 58;
    const long EMR_BEGINPATH                   = 59;
    const long EMR_ENDPATH                     = 60;
    const long EMR_CLOSEFIGURE                 = 61;
    const long EMR_FILLPATH                    = 62;
    const long EMR_STROKEANDFILLPATH           = 63;
    const long EMR_STROKEPATH                  = 64;
    const long EMR_FLATTENPATH                 = 65;
    const long EMR_WIDENPATH                   = 66;
    const long EMR_SELECTCLIPPATH              = 67;
    const long EMR_ABORTPATH                   = 68;

    const long EMR_GDICOMMENT                  = 70;
    const long EMR_FILLRGN                     = 71;
    const long EMR_FRAMERGN                    = 72;
    const long EMR_INVERTRGN                   = 73;
    const long EMR_PAINTRGN                    = 74;
    const long EMR_EXTSELECTCLIPRGN            = 75;
    const long EMR_BITBLT                      = 76;
    const long EMR_STRETCHBLT                  = 77;
    const long EMR_MASKBLT                     = 78;
    const long EMR_PLGBLT                      = 79;
    const long EMR_SETDIBITSTODEVICE           = 80;
    const long EMR_STRETCHDIBITS               = 81;
    const long EMR_EXTCREATEFONTINDIRECTW      = 82;
    const long EMR_EXTTEXTOUTA                 = 83;
    const long EMR_EXTTEXTOUTW                 = 84;
    const long EMR_POLYBEZIER16                = 85;
    const long EMR_POLYGON16                   = 86;
    const long EMR_POLYLINE16                  = 87;
    const long EMR_POLYBEZIERTO16              = 88;
    const long EMR_POLYLINETO16                = 89;
    const long EMR_POLYPOLYLINE16              = 90;
    const long EMR_POLYPOLYGON16               = 91;
    const long EMR_POLYDRAW16                  = 92;
    const long EMR_CREATEMONOBRUSH             = 93;
    const long EMR_CREATEDIBPATTERNBRUSHPT     = 94;
    const long EMR_EXTCREATEPEN                = 95;
    const long EMR_POLYTEXTOUTA                = 96;
    const long EMR_POLYTEXTOUTW                = 97;
    const long EMR_SETICMMODE                  = 98;
    const long EMR_CREATECOLORSPACE            = 99;
    const long EMR_SETCOLORSPACE               = 100;
    const long EMR_DELETECOLORSPACE            = 101;
    const long EMR_MIN                         = 1;
    //#if (WINVER >= 0x0400)
    const long EMR_MAX_XP                      = 101;
    const long EMR_MAX                         = 97;

    const long ICM_OFF   = 1;
    const long ICM_ON    = 2;
    const long ICM_QUERY = 3;
	
	const long FONTMAPPER_MAX = 10;

	const long DI_APPBANDING   = 0x0001;

    // Flags value for COLORADJUSTMENT 
    const UINT CA_NEGATIVE                 = 0x0001;
    const UINT CA_LOG_FILTER               = 0x0002;

    // IlluminantIndex values 
    const long ILLUMINANT_DEVICE_DEFAULT   = 0;
    const long ILLUMINANT_A                = 1;
    const long ILLUMINANT_B                = 2;
    const long ILLUMINANT_C                = 3;
    const long ILLUMINANT_D50              = 4;
    const long ILLUMINANT_D55              = 5;
    const long ILLUMINANT_D65              = 6;
    const long ILLUMINANT_D75              = 7;
    const long ILLUMINANT_F2               = 8;
    const long ILLUMINANT_MAX_INDEX        = 8; // ILLUMINANT_F2

    const long ILLUMINANT_TUNGSTEN         = 1; // ILLUMINANT_A
    const long ILLUMINANT_DAYLIGHT         = 3; // ILLUMINANT_C
    const long ILLUMINANT_FLUORESCENT      = 8; // ILLUMINANT_F2
    const long ILLUMINANT_NTSC             = 3; // ILLUMINANT_C

    // Min and max for RedGamma, GreenGamma, BlueGamma 
    const UINT RGB_GAMMA_MIN               = 02500;
    const UINT RGB_GAMMA_MAX               = 65000;

    // Min and max for ReferenceBlack and ReferenceWhite 
    const UINT REFERENCE_WHITE_MIN         = 6000;
    const UINT REFERENCE_WHITE_MAX         = 10000;
    const UINT REFERENCE_BLACK_MIN         = 0;
    const UINT REFERENCE_BLACK_MAX         = 4000;

    // Min and max for Contrast, Brightness, Colorfulness, RedGreenTint 
    const short COLOR_ADJ_MIN               = -100;
    const short COLOR_ADJ_MAX               = 100;

};