//Structures that are using Padding mechanism to cope with the way VB
//tends to naturally align variables
//NMPGSCROLLVB
//A_CHARFORMAT
//W_CHARFORMAT
//A_CHARFORMAT2
//W_CHARFORMAT2

typedef struct DOCINFO { 
  long	cbSize; 
  long	lpszDocName;  //LPSTR LPWSTR
  long	lpszOutput;   //LPSTR LPWSTR
  long	lpszDatatype; //LPSTR LPWSTR
  DWORD	fwType; 
} DOCINFO; 

typedef struct POINTAPI { 
  long x; 
  long y; 
} POINTAPI;

//Used by GetMessagePos
typedef struct POINTS { 
  SHORT x; 
  SHORT y; 
} POINTS;

typedef struct RECT {
    long Left;
    long Top;
    long Right;
    long Bottom;
} RECT;

typedef struct SIZE {
    long cx;
    long cy;
} SIZE;

typedef struct SIZEL {
    LONG cx;
    LONG cy;
} SIZEL;

typedef struct COPY_DATA {
    LONG dwData;
    LONG cbData;
    LONG lpData;
} COPY_DATA;

// ************

typedef struct _GENERIC_MAPPING {
    DWORD GenericRead;
    DWORD GenericWrite;
    DWORD GenericExecute;
    DWORD GenericAll;
} GENERIC_MAPPING;

typedef struct LUID {
    DWORD LowPart;
    LONG HighPart;
} LUID;

typedef struct LUID_AND_ATTRIBUTES {
    LUID Luid;
    DWORD Attributes;
} LUID_AND_ATTRIBUTES;

typedef struct TOKEN_PRIVILEGES {
    DWORD PrivilegeCount;
    LUID_AND_ATTRIBUTES Privileges[1];
} TOKEN_PRIVILEGES;

typedef struct PRIVILEGE_SET {
    DWORD PrivilegeCount;
    DWORD Control;
    LUID_AND_ATTRIBUTES Privilege[1];
} PRIVILEGE_SET;

typedef struct ACL {
    BYTE  AclRevision;
    BYTE  Sbz1;
    WORD   AclSize;
    WORD   AceCount;
    WORD   Sbz2;
} ACL;

typedef struct SECURITY_DESCRIPTOR_RELATIVE {
    BYTE  Revision;
    BYTE  Sbz1;
    long Control;
    DWORD Owner;
    DWORD Group;
    DWORD Sacl;
    DWORD Dacl;
    } SECURITY_DESCRIPTOR_RELATIVE;

typedef struct SECURITY_DESCRIPTOR {
   BYTE  Revision;
   BYTE  Sbz1;
   long Control;
   long Owner;
   long Group;
   ACL Sacl;
   ACL Dacl;

   } SECURITY_DESCRIPTOR;

// ***********
//contains information about a low-level keyboard input event. 
typedef struct KBDLLHOOKSTRUCT {
    DWORD vkCode;
    DWORD scanCode;
    DWORD flags;
    DWORD time;
    long  dwExtraInfo;
} KBDLLHOOKSTRUCT;

typedef struct {
    POINTAPI	pt;
    HWND		hwnd;
    UINT		wHitTestCode;
    long		dwExtraInfo;
} MOUSEHOOKSTRUCT;

typedef struct MOUSEHOOKSTRUCTEX {
    MOUSEHOOKSTRUCT MOUSEHOOKSTRUCT;
    DWORD mouseData;
} MOUSEHOOKSTRUCTEX;

typedef struct {
    POINTAPI	pt;
    DWORD		mouseData;
    DWORD		flags;
    DWORD		time;
    long		dwExtraInfo;
} MSLLHOOKSTRUCT;

// ***********

typedef struct LARGE_INTEGER {
	DWORD	LowPart;		//Low-order + High-order 32 bits. Use MAKEDWORD
	DWORD	HighPart;		//Low-order + High-order 32 bits
    //double	QuadPart;	//Signed 64-bit integer 8 bytes. LONGLONG
} LARGE_INTEGER;

typedef struct ULARGE_INTEGER {
	DWORD	LowPart;		//Low-order + High-order 32 bits
	DWORD	HighPart;		//Low-order + High-order 32 bits 
    //double	QuadPart;	// ULONGLONG
} ULARGE_INTEGER;

typedef struct LARGE_INTEGER_D {
    LONGLONG QuadPart;
} LARGE_INTEGER_D;

typedef struct ULARGE_INTEGER_D {
    ULONGLONG QuadPart;
} ULARGE_INTEGER_D;

//typedef struct VB_LARGE_INTEGER {
//	DWORD  LowPart;         
//	DWORD  HighPart; 
//} VB_LARGE_INTEGER;

typedef struct SHQUERYRBINFO {
  Long				cbSize;
  LARGE_INTEGER	i64Size;
  LARGE_INTEGER	i64NumItems;
} SHQUERYRBINFO;


//GDI related

typedef struct KERNINGPAIR { 
  WORD  wFirst; 
  WORD  wSecond; 
  long  iKernAmount; 
} KERNINGPAIR; 

typedef struct RGNDATAHEADER { 
    DWORD dwSize; 
    DWORD iType; 
    DWORD nCount; 
    DWORD nRgnSize; 
    RECT  rcBound; 
} RGNDATAHEADER; 

typedef struct _RGNDATA { 
    RGNDATAHEADER rdh; 
    BYTE          Buffer[1]; 
} RGNDATA; 

typedef struct PALETTEENTRY {
    BYTE peRed;
    BYTE peGreen;
    BYTE peBlue;
    BYTE peFlags;
} PALETTEENTRY;

typedef struct LOGPALETTE {
    short palVersion;
    short palNumEntries;
    PALETTEENTRY palPalEntry;
} LOGPALETTE;

typedef struct LOGPALETTE256 {
    WORD			palVersion;
    WORD			palNumEntries;
    PALETTEENTRY	palPalEntry[256];
} LOGPALETTE256;

typedef struct XFORM {
    FLOAT   eM11;
    FLOAT   eM12;
    FLOAT   eM21;
    FLOAT   eM22;
    FLOAT   eDx;
    FLOAT   eDy;
} XFORM;

// Bitmap Header Definition 
typedef struct BITMAP {
    LONG        bmType;
    LONG        bmWidth;
    LONG        bmHeight;
    LONG        bmWidthBytes;
    WORD        bmPlanes;
    WORD        bmBitsPixel;
    long 		bmBits; // LPVOID 
} BITMAP;

typedef struct RGBTRIPLE {
        BYTE    rgbtBlue;
        BYTE    rgbtGreen;
        BYTE    rgbtRed;
} RGBTRIPLE;

typedef struct RGBQUAD {
        BYTE    rgbBlue;
        BYTE    rgbGreen;
        BYTE    rgbRed;
        BYTE    rgbReserved;
} RGBQUAD;

// ICM Color Definitions 
// The following two structures are used for defining RGB's in terms of
// CIEXYZ. The values are fixed point 16.16.

typedef struct CIEXYZ {
	long ciexyzX;
	long ciexyzY;
	long ciexyzZ;
} CIEXYZ;

typedef struct CIEXYZTRIPLE {
	CIEXYZ  ciexyzRed;
	CIEXYZ  ciexyzGreen;
	CIEXYZ  ciexyzBlue;
} CIEXYZTRIPLE;

// The next structures the logical color space. Unlike pens and brushes,
// but like palettes, there is only one way to create a LogColorSpace.
// A pointer to it must be passed, its elements can't be pushed as
// arguments.

//LOGCOLORSPACEA
typedef struct A_LOGCOLORSPACE {
    DWORD lcsSignature;
    DWORD lcsVersion;
    DWORD lcsSize;
    long	lcsCSType;
    long	lcsIntent;
    CIEXYZTRIPLE 	lcsEndpoints;
    DWORD lcsGammaRed;
    DWORD lcsGammaGreen;
    DWORD lcsGammaBlue;
    BYTE  lcsFilename[260]; // MAX_PATH
} A_LOGCOLORSPACEA;

//LOGCOLORSPACEW
typedef struct W_LOGCOLORSPACE {
    DWORD lcsSignature;
    DWORD lcsVersion;
    DWORD lcsSize;
    long	lcsCSType;
    long	lcsIntent;
    CIEXYZTRIPLE 	lcsEndpoints;
    DWORD lcsGammaRed;
    DWORD lcsGammaGreen;
    DWORD lcsGammaBlue;
    BYTE  lcsFilename[260]; // MAX_PATH
} W_LOGCOLORSPACE;

// structures for defining DIBs 
typedef struct BITMAPCOREHEADER {
        DWORD   bcSize;                 // used to get to color table 
        WORD    bcWidth;
        WORD    bcHeight;
        WORD    bcPlanes;
        WORD    bcBitCount;
} BITMAPCOREHEADER;


typedef struct BITMAPINFOHEADER {
        DWORD      biSize;
        LONG       biWidth;
        LONG       biHeight;
        WORD       biPlanes;
        WORD       biBitCount;
        DWORD      biCompression;
        DWORD      biSizeImage;
        LONG       biXPelsPerMeter;
        LONG       biYPelsPerMeter;
        DWORD      biClrUsed;
        DWORD      biClrImportant;
} BITMAPINFOHEADER;


typedef struct BITMAPV4HEADER {
        DWORD        bV4Size;
        LONG         bV4Width;
        LONG         bV4Height;
        WORD         bV4Planes;
        WORD         bV4BitCount;
        DWORD        bV4V4Compression;
        DWORD        bV4SizeImage;
        LONG         bV4XPelsPerMeter;
        LONG         bV4YPelsPerMeter;
        DWORD        bV4ClrUsed;
        DWORD        bV4ClrImportant;
        DWORD        bV4RedMask;
        DWORD        bV4GreenMask;
        DWORD        bV4BlueMask;
        DWORD        bV4AlphaMask;
        DWORD        bV4CSType;
        long		 bV4Endpoints;
        DWORD        bV4GammaRed;
        DWORD        bV4GammaGreen;
        DWORD        bV4GammaBlue;
} BITMAPV4HEADER;

typedef struct BITMAPV5HEADER { 
  DWORD        bV5Size; 
  LONG         bV5Width; 
  LONG         bV5Height; 
  WORD         bV5Planes; 
  WORD         bV5BitCount; 
  DWORD        bV5Compression; 
  DWORD        bV5SizeImage; 
  LONG         bV5XPelsPerMeter; 
  LONG         bV5YPelsPerMeter; 
  DWORD        bV5ClrUsed; 
  DWORD        bV5ClrImportant; 
  DWORD        bV5RedMask; 
  DWORD        bV5GreenMask; 
  DWORD        bV5BlueMask; 
  DWORD        bV5AlphaMask; 
  DWORD        bV5CSType; 
  CIEXYZTRIPLE bV5Endpoints; 
  DWORD        bV5GammaRed; 
  DWORD        bV5GammaGreen; 
  DWORD        bV5GammaBlue; 
  DWORD        bV5Intent; 
  DWORD        bV5ProfileData; 
  DWORD        bV5ProfileSize; 
  DWORD        bV5Reserved; 
} BITMAPV5HEADER; 

typedef struct BITMAPINFO {
    BITMAPINFOHEADER    bmiHeader;
    RGBQUAD             bmiColors[1];
} BITMAPINFO;

typedef struct BITMAPINFO256 {
    BITMAPINFOHEADER    bmiHeader;
    RGBQUAD             bmiColors[256];
} BITMAPINFO256;

typedef struct BITMAPCOREINFO {
    BITMAPCOREHEADER    bmciHeader;
    RGBTRIPLE           bmciColors[1];
} BITMAPCOREINFO;

typedef struct BITMAPCOREINFO256 {
    BITMAPCOREHEADER    bmciHeader;
    RGBTRIPLE           bmciColors[256];
} BITMAPCOREINFO256;

typedef struct BITMAPFILEHEADER {
        WORD    bfType;
        DWORD   bfSize;
        WORD    bfReserved1;
        WORD    bfReserved2;
        DWORD   bfOffBits;
} BITMAPFILEHEADER;

typedef struct FONTSIGNATURE {
    DWORD fsUsb[4];
    DWORD fsCsb[2];
} FONTSIGNATURE;

typedef struct CHARSETINFO {
    UINT ciCharset;
    UINT ciACP;
    FONTSIGNATURE fs;
} CHARSETINFO;

typedef struct LOCALESIGNATURE {
    DWORD lsUsb[4];
    DWORD lsCsbDefault[2];
    DWORD lsCsbSupported[2];
} LOCALESIGNATURE;

// Clipboard Metafile Picture Structure 
typedef struct HANDLETABLE {
    long objectHandle[1];
} HANDLETABLE;

typedef struct METARECORD {
    DWORD       rdSize;
    WORD        rdFunction;
    WORD        rdParm[1];
} METARECORD;

typedef struct METAFILEPICT {
    LONG        mm;
    LONG        xExt;
    LONG        yExt;
    long		hMF;
} METAFILEPICT;

typedef struct METAHEADER {
    WORD        mtType;
    WORD        mtHeaderSize;
    WORD        mtVersion;
    DWORD       mtSize;
    WORD        mtNoObjects;
    DWORD       mtMaxRecord;
    WORD        mtNoParameters;
} METAHEADER;

// Enhanced Metafile structures 
typedef struct ENHMETARECORD {
    DWORD   iType;              // Record type EMR_XXX
    DWORD   nSize;              // Record size in bytes
    DWORD   dParm[1];           // Parameters
} ENHMETARECORD;

typedef struct ENHMETAHEADER {
    DWORD   iType;              // Record type EMR_HEADER
    DWORD   nSize;              // Record size in bytes.  This may be greater
                                // than the sizeof(ENHMETAHEADER).
    RECT    rclBounds;          // Inclusive-inclusive bounds in device units
    RECT    rclFrame;           // Inclusive-inclusive Picture Frame of metafile in .01 mm units
    DWORD   dSignature;         // Signature.  Must be ENHMETA_SIGNATURE.
    DWORD   nVersion;           // Version number
    DWORD   nBytes;             // Size of the metafile in bytes
    DWORD   nRecords;           // Number of records in the metafile
    WORD    nHandles;           // Number of handles in the handle table
                                // Handle index zero is reserved.
    WORD    sReserved;          // Reserved.  Must be zero.
    DWORD   nDescription;       // Number of chars in the unicode description string
                                // This is 0 if there is no description string
    DWORD   offDescription;     // Offset to the metafile description record.
                                // This is 0 if there is no description string
    DWORD   nPalEntries;        // Number of entries in the metafile palette.
    SIZEL   szlDevice;          // Size of the reference device in pels
    SIZEL   szlMillimeters;     // Size of the reference device in millimeters
} ENHMETAHEADER;


typedef struct A_TEXTMETRIC {
    LONG        tmHeight;
    LONG        tmAscent;
    LONG        tmDescent;
    LONG        tmInternalLeading;
    LONG        tmExternalLeading;
    LONG        tmAveCharWidth;
    LONG        tmMaxCharWidth;
    LONG        tmWeight;
    LONG        tmOverhang;
    LONG        tmDigitizedAspectX;
    LONG        tmDigitizedAspectY;
    BYTE       tmFirstChar;
    BYTE       tmLastChar;
    BYTE       tmDefaultChar;
    BYTE       tmBreakChar;
    BYTE        tmItalic;
    BYTE        tmUnderlined;
    BYTE        tmStruckOut;
    BYTE        tmPitchAndFamily;
    BYTE        tmCharSet;
} A_TEXTMETRIC;

typedef struct W_TEXTMETRIC {
    LONG        tmHeight;
    LONG        tmAscent;
    LONG        tmDescent;
    LONG        tmInternalLeading;
    LONG        tmExternalLeading;
    LONG        tmAveCharWidth;
    LONG        tmMaxCharWidth;
    LONG        tmWeight;
    LONG        tmOverhang;
    LONG        tmDigitizedAspectX;
    LONG        tmDigitizedAspectY;
    short       tmFirstChar;
    short       tmLastChar;
    short       tmDefaultChar;
    short       tmBreakChar;
    BYTE        tmItalic;
    BYTE        tmUnderlined;
    BYTE        tmStruckOut;
    BYTE        tmPitchAndFamily;
    BYTE        tmCharSet;
} W_TEXTMETRIC;

typedef struct A_NEWTEXTMETRIC {
    LONG        tmHeight;
    LONG        tmAscent;
    LONG        tmDescent;
    LONG        tmInternalLeading;
    LONG        tmExternalLeading;
    LONG        tmAveCharWidth;
    LONG        tmMaxCharWidth;
    LONG        tmWeight;
    LONG        tmOverhang;
    LONG        tmDigitizedAspectX;
    LONG        tmDigitizedAspectY;
    BYTE       tmFirstChar;
    BYTE       tmLastChar;
    BYTE       tmDefaultChar;
    BYTE       tmBreakChar;
    BYTE        tmItalic;
    BYTE        tmUnderlined;
    BYTE        tmStruckOut;
    BYTE        tmPitchAndFamily;
    BYTE        tmCharSet;
    DWORD   ntmFlags;
    UINT    ntmSizeEM;
    UINT    ntmCellHeight;
    UINT    ntmAvgWidth;
} A_NEWTEXTMETRIC;

typedef struct W_NEWTEXTMETRIC {
    LONG        tmHeight;
    LONG        tmAscent;
    LONG        tmDescent;
    LONG        tmInternalLeading;
    LONG        tmExternalLeading;
    LONG        tmAveCharWidth;
    LONG        tmMaxCharWidth;
    LONG        tmWeight;
    LONG        tmOverhang;
    LONG        tmDigitizedAspectX;
    LONG        tmDigitizedAspectY;
    short       tmFirstChar;
    short       tmLastChar;
    short       tmDefaultChar;
    short       tmBreakChar;
    BYTE        tmItalic;
    BYTE        tmUnderlined;
    BYTE        tmStruckOut;
    BYTE        tmPitchAndFamily;
    BYTE        tmCharSet;
    DWORD   ntmFlags;
    UINT    ntmSizeEM;
    UINT    ntmCellHeight;
    UINT    ntmAvgWidth;
} W_NEWTEXTMETRIC;

typedef struct A_NEWTEXTMETRICEX {
    A_NEWTEXTMETRIC   ntmTm;
    FONTSIGNATURE   ntmFontSig;
} A_NEWTEXTMETRICEX;

typedef struct W_NEWTEXTMETRICEX {
    W_NEWTEXTMETRIC   ntmTm;
    FONTSIGNATURE   ntmFontSig;
} W_NEWTEXTMETRICEX;

// GDI Logical Objects:

// Pel Array 
typedef struct PELARRAY {
    LONG        paXCount;
    LONG        paYCount;
    LONG        paXExt;
    LONG        paYExt;
    BYTE        paRGBs;
  } PELARRAY;

// Logical Brush (or Pattern) 
typedef struct LOGBRUSH {
    UINT        lbStyle;
    COLORREF    lbColor;
    LONG        lbHatch;
} LOGBRUSH;

// Logical Pen 
typedef struct LOGPEN {
    UINT        lopnStyle;
    POINTAPI    lopnWidth;
    COLORREF    lopnColor;
} LOGPEN;

typedef struct EXTLOGPEN {
    DWORD       elpPenStyle;
    DWORD       elpWidth;
    UINT        elpBrushStyle;
    COLORREF    elpColor;
    LONG        elpHatch;
    DWORD       elpNumEntries;
    DWORD       elpStyleEntry[1];
} EXTLOGPEN;

typedef struct A_LOGFONT {
    LONG      lfHeight;
    LONG      lfWidth;
    LONG      lfEscapement;
    LONG      lfOrientation;
    LONG      lfWeight;
    BYTE      lfItalic;
    BYTE      lfUnderline;
    BYTE      lfStrikeOut;
    BYTE      lfCharSet;
    BYTE      lfOutPrecision;
    BYTE      lfClipPrecision;
    BYTE      lfQuality;
    BYTE      lfPitchAndFamily;
    BYTE	  lfFaceName[32];		// LF_FACESIZE
} A_LOGFONT;

typedef struct W_LOGFONT {
    LONG      lfHeight;
    LONG      lfWidth;
    LONG      lfEscapement;
    LONG      lfOrientation;
    LONG      lfWeight;
    BYTE      lfItalic;
    BYTE      lfUnderline;
    BYTE      lfStrikeOut;
    BYTE      lfCharSet;
    BYTE      lfOutPrecision;
    BYTE      lfClipPrecision;
    BYTE      lfQuality;
    BYTE      lfPitchAndFamily;
    BYTE	  lfFaceName[64];	// LF_FACESIZE
} W_LOGFONT;

// Structure passed to FONTENUMPROC 
typedef struct A_ENUMLOGFONT {
    A_LOGFONT elfLogFont;
    BYTE     elfFullName[64];		// LF_FULLFACESIZE
    BYTE	 elfStyle[32];			// LF_FACESIZE
} A_ENUMLOGFONT;

typedef struct W_ENUMLOGFONT {
    W_LOGFONT elfLogFont;
    BYTE     elfFullName[128];	// LF_FULLFACESIZE
    BYTE	 elfStyle[64];		// LF_FACESIZE
} W_ENUMLOGFONT;

typedef struct A_ENUMLOGFONTEX {
    A_LOGFONT    elfLogFont;
    BYTE        elfFullName[64]; // LF_FULLFACESIZE
    BYTE        elfStyle[32]; // LF_FACESIZE
    BYTE        elfScript[32]; // LF_FACESIZE
} A_ENUMLOGFONTEX;

typedef struct W_ENUMLOGFONTEX {
    W_LOGFONT    elfLogFont;
    BYTE        elfFullName[64]; // LF_FULLFACESIZE
    BYTE        elfStyle[32]; // LF_FACESIZE
    BYTE        elfScript[32]; // LF_FACESIZE
} W_ENUMLOGFONTEX;

typedef struct PANOSE {
    BYTE    bFamilyType;
    BYTE    bSerifStyle;
    BYTE    bWeight;
    BYTE    bProportion;
    BYTE    bContrast;
    BYTE    bStrokeVariation;
    BYTE    bArmStyle;
    BYTE    bLetterform;
    BYTE    bMidline;
    BYTE    bXHeight;
} PANOSE;


// The extended logical font A
// An extension of the ENUMLOGFONT 

typedef struct A_EXTLOGFONT {
    A_LOGFONT     elfLogFont;
    BYTE        elfFullName[64];		// FULLFACESIZE
    BYTE        elfStyle[64];			// FACESIZE
    DWORD       elfVersion;     // 0 for the first release of NT 
    DWORD       elfStyleSize;
    DWORD       elfMatch;
    DWORD       elfReserved;
    BYTE        elfVendorId[4]; // ELF_VENDOR_SIZE     
    DWORD       elfCulture;     // 0 for Latin                   
    PANOSE      elfPanose;
} A_EXTLOGFONT;

// The extended logical font W
// An extension of the ENUMLOGFONT 

typedef struct W_EXTLOGFONT {
    W_LOGFONT     elfLogFont;
    BYTE        elfFullName[128];	// FULLFACESIZE
    BYTE        elfStyle[128];		// FACESIZE
    DWORD       elfVersion;     // 0 for the first release of NT 
    DWORD       elfStyleSize;
    DWORD       elfMatch;
    DWORD       elfReserved;
    BYTE        elfVendorId[4]; // ELF_VENDOR_SIZE     
    DWORD       elfCulture;     // 0 for Latin                   
    PANOSE      elfPanose;
} W_EXTLOGFONT;

typedef struct A_DEVMODE{
    BYTE dmDeviceName[32];
    short dmSpecVersion;
    short dmDriverVersion;
    short dmSize;
    short dmDriverExtra;
    long  dmFields;
    short dmOrientation;
    short dmPaperSize;
    short dmPaperLength;
    short dmPaperWidth;
    short dmScale;
    short dmCopies;
    short dmDefaultSource;
    short dmPrintQuality;
    short dmColor;
    short dmDuplex;
    short dmYResolution;
    short dmTTOption;
    short dmCollate;
    BYTE  dmFormName[32];
    short dmLogPixels;
    long  dmBitsPerPel;
    long  dmPelsWidth;
    long  dmPelsHeight;
    long  dmDisplayFlags;
    long  dmDisplayFrequency;
    long  dmICMMethod;
    long  dmICMIntent;
    long  dmMediaType;
    long  dmDitherType;
    long  dmReserved1;
    long  dmReserved2;
} A_DEVMODE;

typedef struct W_DEVMODE{
    BYTE dmDeviceName[64];
    short dmSpecVersion;
    short dmDriverVersion;
    short dmSize;
    short dmDriverExtra;
    long  dmFields;
    short dmOrientation;
    short dmPaperSize;
    short dmPaperLength;
    short dmPaperWidth;
    short dmScale;
    short dmCopies;
    short dmDefaultSource;
    short dmPrintQuality;
    short dmColor;
    short dmDuplex;
    short dmYResolution;
    short dmTTOption;
    short dmCollate;
    BYTE  dmFormName[64];
    short dmLogPixels;
    long  dmBitsPerPel;
    long  dmPelsWidth;
    long  dmPelsHeight;
    long  dmDisplayFlags;
    long  dmDisplayFrequency;
    long  dmICMMethod;
    long  dmICMIntent;
    long  dmMediaType;
    long  dmDitherType;
    long  dmReserved1;
    long  dmReserved2;
} W_DEVMODE;


typedef struct ANEW_DEVMODE{
    BYTE dmDeviceName[32];
    short dmSpecVersion;
    short dmDriverVersion;
    short dmSize;
    short dmDriverExtra;
    long  dmFields;
    short dmOrientation;
    short dmPaperSize;
    short dmPaperLength;
    short dmPaperWidth;
    short dmScale;
    short dmCopies;
    short dmDefaultSource;
    short dmPrintQuality;
    short dmColor;
    short dmDuplex;
    short dmYResolution;
    short dmTTOption;
    short dmCollate;
    BYTE  dmFormName[32];
    short dmLogPixels;
    long  dmBitsPerPel;
    long  dmPelsWidth;
    long  dmPelsHeight;
    long  dmDisplayFlags;
    long  dmDisplayFrequency;
    long  dmICMMethod;
    long  dmICMIntent;
    long  dmMediaType;
    long  dmDitherType;
    long  dmReserved1;
    long  dmReserved2;
//if (WINVER >= 0x0500) || (_WIN32_WINNT >= 0x0400)
    long  dmPanningWidth;
    long  dmPanningHeight;
} ANEW_DEVMODE;


typedef struct WNEW_DEVMODE{
    BYTE dmDeviceName[64];
    short dmSpecVersion;
    short dmDriverVersion;
    short dmSize;
    short dmDriverExtra;
    long  dmFields;
    short dmOrientation;
    short dmPaperSize;
    short dmPaperLength;
    short dmPaperWidth;
    short dmScale;
    short dmCopies;
    short dmDefaultSource;
    short dmPrintQuality;
    short dmColor;
    short dmDuplex;
    short dmYResolution;
    short dmTTOption;
    short dmCollate;
    BYTE  dmFormName[64];
    short dmLogPixels;
    long  dmBitsPerPel;
    long  dmPelsWidth;
    long  dmPelsHeight;
    long  dmDisplayFlags;
    long  dmDisplayFrequency;
    long  dmICMMethod;
    long  dmICMIntent;
    long  dmMediaType;
    long  dmDitherType;
    long  dmReserved1;
    long  dmReserved2;
//if (WINVER >= 0x0500) || (_WIN32_WINNT >= 0x0400)
    long  dmPanningWidth;
    long  dmPanningHeight;
} WNEW_DEVMODE;

typedef struct A_OUTLINETEXTMETRIC {
    UINT    otmSize;
    A_TEXTMETRIC otmTextMetrics;
    BYTE    otmFiller;
    PANOSE  otmPanoseNumber;
    UINT    otmfsSelection;
    UINT    otmfsType;
    long    otmsCharSlopeRise;
    long    otmsCharSlopeRun;
    long    otmItalicAngle;
    UINT    otmEMSquare;
    long    otmAscent;
    long    otmDescent;
    UINT    otmLineGap;
    UINT    otmsCapEmHeight;
    UINT    otmsXHeight;
    RECT    otmrcFontBox;
    long    otmMacAscent;
    long    otmMacDescent;
    UINT    otmMacLineGap;
    UINT    otmusMinimumPPEM;
    POINTAPI  otmptSubscriptSize;
    POINTAPI  otmptSubscriptOffset;
    POINTAPI  otmptSuperscriptSize;
    POINTAPI  otmptSuperscriptOffset;
    UINT    otmsStrikeoutSize;
    long    otmsStrikeoutPosition;
    long    otmsUnderscoreSize;
    long    otmsUnderscorePosition;
    long   otmpFamilyName; //LPSTR
    long   otmpFaceName; //LPSTR
    long   otmpStyleName; //LPSTR
    long   otmpFullName; //LPSTR
} A_OUTLINETEXTMETRIC;

typedef struct W_OUTLINETEXTMETRIC {
    UINT    otmSize;
    W_TEXTMETRIC otmTextMetrics;
    BYTE    otmFiller;
    PANOSE  otmPanoseNumber;
    UINT    otmfsSelection;
    UINT    otmfsType;
    long    otmsCharSlopeRise;
    long    otmsCharSlopeRun;
    long    otmItalicAngle;
    UINT    otmEMSquare;
    long    otmAscent;
    long    otmDescent;
    UINT    otmLineGap;
    UINT    otmsCapEmHeight;
    UINT    otmsXHeight;
    RECT    otmrcFontBox;
    long    otmMacAscent;
    long    otmMacDescent;
    UINT    otmMacLineGap;
    UINT    otmusMinimumPPEM;
    POINTAPI  otmptSubscriptSize;
    POINTAPI  otmptSubscriptOffset;
    POINTAPI  otmptSuperscriptSize;
    POINTAPI  otmptSuperscriptOffset;
    UINT    otmsStrikeoutSize;
    long    otmsStrikeoutPosition;
    long    otmsUnderscoreSize;
    long    otmsUnderscorePosition;
    long   otmpFamilyName;
    long   otmpFaceName;
    long   otmpStyleName;
    long   otmpFullName;
} W_OUTLINETEXTMETRIC;

typedef struct POLYTEXT {
    long    x;
    long    y;
    UINT	n;
    long	lpstr; //LPSTR  LPWSTR
    UINT    uiFlags;
    RECT    rcl;
    long    pdx;
} POLYTEXT;

typedef struct FIXED {
    WORD    fract;
    short   value;
} FIXED;

typedef struct MAT2 {
     FIXED  eM11;
     FIXED  eM12;
     FIXED  eM21;
     FIXED  eM22;
} MAT2;

typedef struct GLYPHMETRICS {
    UINT		gmBlackBoxX;
    UINT		gmBlackBoxY;
    POINTAPI	gmptGlyphOrigin;
    short		gmCellIncX;
    short		gmCellIncY;
} GLYPHMETRICS;

typedef struct POINTFX {
    FIXED x;
    FIXED y;
} POINTFX;

typedef struct TTPOLYCURVE {
    WORD    wType;
    WORD    cpfx;
    POINTFX apfx[1];
} TTPOLYCURVE;

typedef struct TTPOLYGONHEADER {
    DWORD   cb;
    DWORD   dwType;
    POINTFX pfxStart;
} TTPOLYGONHEADER;

typedef struct GCP_RESULTS {
    DWORD   lStructSize;
    long    lpOutString;	//LPSTR LPWSTR
    UINT	lpOrder;
    long	lpDx;
    long	lpCaretPos;
    long    lpClass;
    long	lpGlyphs;	//LPSTR LPWSTR
    UINT    nGlyphs;
    long     nMaxFit;
} GCP_RESULTS;

typedef struct RASTERIZER_STATUS {
    short   nSize;
    short   wFlags;
    short   nLanguageID;
} RASTERIZER_STATUS;

// Pixel format descriptor 
typedef struct PIXELFORMATDESCRIPTOR {
    WORD  nSize;
    WORD  nVersion;
    DWORD dwFlags;
    BYTE  iPixelType;
    BYTE  cColorBits;
    BYTE  cRedBits;
    BYTE  cRedShift;
    BYTE  cGreenBits;
    BYTE  cGreenShift;
    BYTE  cBlueBits;
    BYTE  cBlueShift;
    BYTE  cAlphaBits;
    BYTE  cAlphaShift;
    BYTE  cAccumBits;
    BYTE  cAccumRedBits;
    BYTE  cAccumGreenBits;
    BYTE  cAccumBlueBits;
    BYTE  cAccumAlphaBits;
    BYTE  cDepthBits;
    BYTE  cStencilBits;
    BYTE  cAuxBuffers;
    BYTE  iLayerType;
    BYTE  bReserved;
    DWORD dwLayerMask;
    DWORD dwVisibleMask;
    DWORD dwDamageMask;
} PIXELFORMATDESCRIPTOR;

typedef struct DIBSECTION {
    BITMAP              dsBm;
    BITMAPINFOHEADER    dsBmih;
    DWORD               dsBitfields[3];
    HANDLE              dshSection;
    DWORD               dsOffset;
} DIBSECTION;

typedef struct  COLORADJUSTMENT {
    WORD   caSize;
    WORD   caFlags;
    WORD   caIlluminantIndex;
    WORD   caRedGamma;
    WORD   caGreenGamma;
    WORD   caBlueGamma;
    WORD   caReferenceBlack;
    WORD   caReferenceWhite;
    SHORT  caContrast;
    SHORT  caBrightness;
    SHORT  caColorfulness;
    SHORT  caRedGreenTint;
} COLORADJUSTMENT;

// Base record type for the enhanced metafile.

typedef struct EMR {
    DWORD   iType;              // Enhanced metafile record type
    DWORD   nSize;              // Length of the record in bytes.
                                // This must be a multiple of 4.
} EMR;

// Base text record type for the enhanced metafile.

typedef struct EMRTEXT {
    POINTAPI   ptlReference;
    DWORD   nChars;
    DWORD   offString;          // Offset to the string
    DWORD   fOptions;
    RECT	rcl;
    DWORD   offDx;              // Offset to the inter-character spacing array.
                                // This is always given.
} EMRTEXT;

typedef struct ABORTPATH {
    EMR     emr;
} ABORTPATH;

typedef struct EMRSELECTCLIPPATH {
    EMR     emr;
    DWORD   iMode;
} EMRSELECTCLIPPATH;

typedef struct EMRSETMITERLIMIT {
    EMR     emr;
    FLOAT   eMiterLimit;
} EMRSETMITERLIMIT;

typedef struct EMRRESTOREDC {
    EMR     emr;
    LONG    iRelative;          // Specifies a relative instance
} EMRRESTOREDC;

typedef struct EMRSETARCDIRECTION {
    EMR     emr;
    DWORD   iArcDirection;      // Specifies the arc direction in the
                                // advanced graphics mode.
} EMRSETARCDIRECTION;

typedef struct EMRSETMAPPERFLAGS {
    EMR     emr;
    DWORD   dwFlags;
} EMRSETMAPPERFLAGS;

typedef struct EMRSETTEXTCOLOR {
    EMR     emr;
    COLORREF crColor;
} EMRSETBKCOLOR;

typedef struct EMRSELECTOBJECT {
    EMR     emr;
    DWORD   ihObject;           // Object handle index
} EMRSELECTOBJECT;

typedef struct EMRSELECTCOLORSPACE {
    EMR     emr;
    DWORD   ihCS;               // ColorSpace handle index
} EMRSELECTCOLORSPACE;

typedef struct EMRSELECTPALETTE {
    EMR     emr;
    DWORD   ihPal;              // Palette handle index, background mode only
} EMRSELECTPALETTE;

typedef struct EMRRESIZEPALETTE {
    EMR     emr;
    DWORD   ihPal;              // Palette handle index
    DWORD   cEntries;
} EMRRESIZEPALETTE;

typedef struct EMRSETPALETTEENTRIES {
    EMR     emr;
    DWORD   ihPal;              // Palette handle index
    DWORD   iStart;
    DWORD   cEntries;
    PALETTEENTRY aPalEntries[1];// The peFlags fields do not contain any flags
} EMRSETPALETTEENTRIES;

typedef struct EMRSETCOLORADJUSTMENT {
    EMR				emr;
    COLORADJUSTMENT ColorAdjustment;
} EMRSETCOLORADJUSTMENT;

typedef struct EMRGDICOMMENT {
    EMR     emr;
    DWORD   cbData;             // Size of data in bytes
    BYTE    Data[1];
} EMRGDICOMMENT;

typedef struct EMREOF {
    EMR     emr;
    DWORD   nPalEntries;        // Number of palette entries
    DWORD   offPalEntries;      // Offset to the palette entries
    DWORD   nSizeLast;          // Same as nSize and must be the last DWORD
                                // of the record.  The palette entries,
                                // if exist, precede this field.
} EMREOF;

typedef struct EMRLINETO {
    EMR     emr;
    POINTAPI   ptl;
} EMRLINETO;

typedef struct EMROFFSETCLIPRGN {
    EMR			emr;
    POINTAPI	ptlOffset;
} EMROFFSETCLIPRGN;

typedef struct EMRFILLPATH {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
} EMRFILLPATH;

typedef struct EMREXCLUDECLIPRECT {
    EMR     emr;
    RECT    rclClip;
} EMREXCLUDECLIPRECT;

typedef struct EMRSETVIEWPORTORGEX {
    EMR     emr;
    POINTAPI   ptlOrigin;
} EMRSETVIEWPORTORGEX;

typedef struct EMRSETVIEWPORTEXTEX {
    EMR     emr;
    SIZEL    szlExtent;
} EMRSETVIEWPORTEXTEX;

typedef struct EMRSCALEVIEWPORTEXTEX {
    EMR     emr;
    LONG    xNum;
    LONG    xDenom;
    LONG    yNum;
    LONG    yDenom;
} EMRSCALEVIEWPORTEXTEX;

typedef struct EMRSETWORLDTRANSFORM {
    EMR     emr;
    XFORM   xform;
} EMRSETWORLDTRANSFORM;

typedef struct EMRMODIFYWORLDTRANSFORM {
    EMR     emr;
    XFORM   xform;
    DWORD   iMode;
} EMRMODIFYWORLDTRANSFORM;

typedef struct EMRSETPIXELV {
    EMR     emr;
    POINTAPI  ptlPixel;
    COLORREF crColor;
} EMRSETPIXELV;

typedef struct EMREXTFLOODFILL {
    EMR     emr;
    POINTAPI  ptlStart;
    COLORREF crColor;
    DWORD   iMode;
} EMREXTFLOODFILL;

typedef struct EMRELLIPSE {
    EMR     emr;
    RECT   rclBox;             // Inclusive-inclusive bounding rectangle
} EMRELLIPSE;

typedef struct EMRROUNDRECT {
    EMR     emr;
    RECT   rclBox;             // Inclusive-inclusive bounding rectangle
    SIZEL   szlCorner;
} EMRROUNDRECT;

typedef struct EMRARC {
    EMR     emr;
    RECT   rclBox;             // Inclusive-inclusive bounding rectangle
    POINTAPI  ptlStart;
    POINTAPI  ptlEnd;
} EMRARC;

typedef struct EMRANGLEARC {
    EMR     emr;
    POINTAPI  ptlCenter;
    DWORD   nRadius;
    FLOAT   eStartAngle;
    FLOAT   eSweepAngle;
} EMRANGLEARC;

typedef struct EMRPOLYLINE {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   cptl;
    POINTAPI  aptl[1];
} EMRPOLYLINE;

typedef struct EMRPOLYLINE16 {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   cpts;
    POINTS  apts[1];
} EMRPOLYLINE16;

typedef struct EMRPOLYDRAW {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   cptl;               // Number of points
    POINTAPI  aptl[1];            // Array of points
    BYTE    abTypes[1];         // Array of point types
} EMRPOLYDRAW;

typedef struct EMRPOLYDRAW16 {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   cpts;               // Number of points
    POINTS  apts[1];            // Array of points
    BYTE    abTypes[1];         // Array of point types
} EMRPOLYDRAW16;

typedef struct EMRPOLYPOLYLINE {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   nPolys;             // Number of polys
    DWORD   cptl;               // Total number of points in all polys
    DWORD   aPolyCounts[1];     // Array of point counts for each poly
    POINTAPI  aptl[1];            // Array of points
} EMRPOLYPOLYLINE;

typedef struct EMRPOLYPOLYLINE16 {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   nPolys;             // Number of polys
    DWORD   cpts;               // Total number of points in all polys
    DWORD   aPolyCounts[1];     // Array of point counts for each poly
    POINTS  apts[1];            // Array of points
} EMRPOLYPOLYLINE16;

typedef struct EMRINVERTRGN {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   cbRgnData;          // Size of region data in bytes
    BYTE    RgnData[1];
} EMRINVERTRGN;

typedef struct EMRFILLRGN {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   cbRgnData;          // Size of region data in bytes
    DWORD   ihBrush;            // Brush handle index
    BYTE    RgnData[1];
} EMRFILLRGN;

typedef struct EMRFRAMERGN {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   cbRgnData;          // Size of region data in bytes
    DWORD   ihBrush;            // Brush handle index
    SIZEL   szlStroke;
    BYTE    RgnData[1];
} EMRFRAMERGN;

typedef struct EMREXTSELECTCLIPRGN {
    EMR     emr;
    DWORD   cbRgnData;          // Size of region data in bytes
    DWORD   iMode;
    BYTE    RgnData[1];
} EMREXTSELECTCLIPRGN;

typedef struct A_EMREXTTEXTOUT {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   iGraphicsMode;      // Current graphics mode
    FLOAT   exScale;            // X and Y scales from Page units to .01mm units
    FLOAT   eyScale;            //   if graphics mode is GM_COMPATIBLE.
    EMRTEXT emrtext;            // This is followed by the string and spacing
                                // array
} A_EMREXTTEXTOUT;

typedef struct W_EMRPOLYTEXTOUT {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    DWORD   iGraphicsMode;      // Current graphics mode
    FLOAT   exScale;            // X and Y scales from Page units to .01mm units
    FLOAT   eyScale;            //   if graphics mode is GM_COMPATIBLE.
    LONG    cStrings;
    EMRTEXT aemrtext[1];        // Array of EMRTEXT structures.  This is
                                // followed by the strings and spacing arrays.
} W_EMRPOLYTEXTOUT;

typedef struct EMRBITBLT {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    LONG    xDest;
    LONG    yDest;
    LONG    cxDest;
    LONG    cyDest;
    DWORD   dwRop;
    LONG    xSrc;
    LONG    ySrc;
    XFORM   xformSrc;           // Source DC transform
    COLORREF crBkColorSrc;      // Source DC BkColor in RGB
    DWORD   iUsageSrc;          // Source bitmap info color table usage
                                // (DIB_RGB_COLORS)
    DWORD   offBmiSrc;          // Offset to the source BITMAPINFO structure
    DWORD   cbBmiSrc;           // Size of the source BITMAPINFO structure
    DWORD   offBitsSrc;         // Offset to the source bitmap bits
    DWORD   cbBitsSrc;          // Size of the source bitmap bits
} EMRBITBLT;

typedef struct EMRSTRETCHBLT {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    LONG    xDest;
    LONG    yDest;
    LONG    cxDest;
    LONG    cyDest;
    DWORD   dwRop;
    LONG    xSrc;
    LONG    ySrc;
    XFORM   xformSrc;           // Source DC transform
    COLORREF crBkColorSrc;      // Source DC BkColor in RGB
    DWORD   iUsageSrc;          // Source bitmap info color table usage
                                // (DIB_RGB_COLORS)
    DWORD   offBmiSrc;          // Offset to the source BITMAPINFO structure
    DWORD   cbBmiSrc;           // Size of the source BITMAPINFO structure
    DWORD   offBitsSrc;         // Offset to the source bitmap bits
    DWORD   cbBitsSrc;          // Size of the source bitmap bits
    LONG    cxSrc;
    LONG    cySrc;
} EMRSTRETCHBLT;

typedef struct EMRMASKBLT {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    LONG    xDest;
    LONG    yDest;
    LONG    cxDest;
    LONG    cyDest;
    DWORD   dwRop;
    LONG    xSrc;
    LONG    ySrc;
    XFORM   xformSrc;           // Source DC transform
    COLORREF crBkColorSrc;      // Source DC BkColor in RGB
    DWORD   iUsageSrc;          // Source bitmap info color table usage
                                // (DIB_RGB_COLORS)
    DWORD   offBmiSrc;          // Offset to the source BITMAPINFO structure
    DWORD   cbBmiSrc;           // Size of the source BITMAPINFO structure
    DWORD   offBitsSrc;         // Offset to the source bitmap bits
    DWORD   cbBitsSrc;          // Size of the source bitmap bits
    LONG    xMask;
    LONG    yMask;
    DWORD   iUsageMask;         // Mask bitmap info color table usage
    DWORD   offBmiMask;         // Offset to the mask BITMAPINFO structure if any
    DWORD   cbBmiMask;          // Size of the mask BITMAPINFO structure if any
    DWORD   offBitsMask;        // Offset to the mask bitmap bits if any
    DWORD   cbBitsMask;         // Size of the mask bitmap bits if any
} EMRMASKBLT;

typedef struct EMRPLGBLT {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    POINTAPI  aptlDest[3];
    LONG    xSrc;
    LONG    ySrc;
    LONG    cxSrc;
    LONG    cySrc;
    XFORM   xformSrc;           // Source DC transform
    COLORREF crBkColorSrc;      // Source DC BkColor in RGB
    DWORD   iUsageSrc;          // Source bitmap info color table usage
                                // (DIB_RGB_COLORS)
    DWORD   offBmiSrc;          // Offset to the source BITMAPINFO structure
    DWORD   cbBmiSrc;           // Size of the source BITMAPINFO structure
    DWORD   offBitsSrc;         // Offset to the source bitmap bits
    DWORD   cbBitsSrc;          // Size of the source bitmap bits
    LONG    xMask;
    LONG    yMask;
    DWORD   iUsageMask;         // Mask bitmap info color table usage
    DWORD   offBmiMask;         // Offset to the mask BITMAPINFO structure if any
    DWORD   cbBmiMask;          // Size of the mask BITMAPINFO structure if any
    DWORD   offBitsMask;        // Offset to the mask bitmap bits if any
    DWORD   cbBitsMask;         // Size of the mask bitmap bits if any
} EMRPLGBLT;

typedef struct EMRSETDIBITSTODEVICE {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    LONG    xDest;
    LONG    yDest;
    LONG    xSrc;
    LONG    ySrc;
    LONG    cxSrc;
    LONG    cySrc;
    DWORD   offBmiSrc;          // Offset to the source BITMAPINFO structure
    DWORD   cbBmiSrc;           // Size of the source BITMAPINFO structure
    DWORD   offBitsSrc;         // Offset to the source bitmap bits
    DWORD   cbBitsSrc;          // Size of the source bitmap bits
    DWORD   iUsageSrc;          // Source bitmap info color table usage
    DWORD   iStartScan;
    DWORD   cScans;
} EMRSETDIBITSTODEVICE;

typedef struct EMRSTRETCHDIBITS {
    EMR     emr;
    RECT   rclBounds;          // Inclusive-inclusive bounds in device units
    LONG    xDest;
    LONG    yDest;
    LONG    xSrc;
    LONG    ySrc;
    LONG    cxSrc;
    LONG    cySrc;
    DWORD   offBmiSrc;          // Offset to the source BITMAPINFO structure
    DWORD   cbBmiSrc;           // Size of the source BITMAPINFO structure
    DWORD   offBitsSrc;         // Offset to the source bitmap bits
    DWORD   cbBitsSrc;          // Size of the source bitmap bits
    DWORD   iUsageSrc;          // Source bitmap info color table usage
    DWORD   dwRop;
    LONG    cxDest;
    LONG    cyDest;
} EMRSTRETCHDIBITS;

typedef struct A_EMREXTCREATEFONTINDIRECT {
    EMR				emr;
    DWORD			ihFont;             // Font handle index
    A_EXTLOGFONT	elfw;
} A_EMREXTCREATEFONTINDIRECT;

typedef struct W_EMREXTCREATEFONTINDIRECT {
    EMR				emr;
    DWORD			ihFont;             // Font handle index
    W_EXTLOGFONT	elfw;
} W_EMREXTCREATEFONTINDIRECT;

typedef struct EMRCREATEPALETTE {
    EMR     emr;
    DWORD   ihPal;              // Palette handle index
    LOGPALETTE lgpl;            // The peFlags fields in the palette entries
                                // do not contain any flags
} EMRCREATEPALETTE;

typedef struct EMRCREATEPALETTE256 {
    EMR     emr;
    DWORD   ihPal;              // Palette handle index
    LOGPALETTE256 lgpl;            // The peFlags fields in the palette entries
                                // do not contain any flags
} EMRCREATEPALETTE256;

typedef struct A_EMRCREATECOLORSPACE {
    EMR         emr;
    DWORD       ihCS;   // ColorSpace handle index
    A_LOGCOLORSPACEA  lcs;
} A_EMRCREATECOLORSPACE;

typedef struct W_EMRCREATECOLORSPACE {
    EMR         emr;
    DWORD       ihCS;   // ColorSpace handle index
    W_LOGCOLORSPACE  lcs;
} W_EMRCREATECOLORSPACE;

typedef struct EMRCREATEPEN {
    EMR     emr;
    DWORD   ihPen;              // Pen handle index
    LOGPEN  lopn;
} EMRCREATEPEN;

typedef struct EMREXTCREATEPEN {
    EMR     emr;
    DWORD   ihPen;              // Pen handle index
    DWORD   offBmi;             // Offset to the BITMAPINFO structure if any
    DWORD   cbBmi;              // Size of the BITMAPINFO structure if any
                                // The bitmap info is followed by the bitmap
                                // bits to form a packed DIB.
    DWORD   offBits;            // Offset to the brush bitmap bits if any
    DWORD   cbBits;             // Size of the brush bitmap bits if any
    EXTLOGPEN elp;              // The extended pen with the style array.
} EMREXTCREATEPEN;

typedef struct EMRCREATEBRUSHINDIRECT {
    EMR     emr;
    DWORD   ihBrush;            // Brush handle index
    LOGBRUSH lb;                // The style must be BS_SOLID, BS_HOLLOW,
                                // BS_NULL or BS_HATCHED.
} EMRCREATEBRUSHINDIRECT;

typedef struct EMRCREATEMONOBRUSH {
    EMR     emr;
    DWORD   ihBrush;            // Brush handle index
    DWORD   iUsage;             // Bitmap info color table usage
    DWORD   offBmi;             // Offset to the BITMAPINFO structure
    DWORD   cbBmi;              // Size of the BITMAPINFO structure
    DWORD   offBits;            // Offset to the bitmap bits
    DWORD   cbBits;             // Size of the bitmap bits
} EMRCREATEMONOBRUSH;

typedef struct EMRCREATEDIBPATTERNBRUSHPT {
    EMR     emr;
    DWORD   ihBrush;            // Brush handle index
    DWORD   iUsage;             // Bitmap info color table usage
    DWORD   offBmi;             // Offset to the BITMAPINFO structure
    DWORD   cbBmi;              // Size of the BITMAPINFO structure
                                // The bitmap info is followed by the bitmap
                                // bits to form a packed DIB.
    DWORD   offBits;            // Offset to the bitmap bits
    DWORD   cbBits;             // Size of the bitmap bits
} EMRCREATEDIBPATTERNBRUSHPT;

typedef struct EMRFORMAT {
    DWORD   dSignature;         // Format signature, e.g. ENHMETA_SIGNATURE.
    DWORD   nVersion;           // Format version number.
    DWORD   cbData;             // Size of data in bytes.
    DWORD   offData;            // Offset to data from GDICOMMENT_IDENTIFIER.
                                // It must begin at a DWORD offset.
} EMRFORMAT;

typedef struct POINTFLOAT {
    long   x;
    long   y;
} POINTFLOAT;

typedef struct GLYPHMETRICSFLOAT {
    FLOAT       gmfBlackBoxX;
    FLOAT       gmfBlackBoxY;
    POINTFLOAT  gmfptGlyphOrigin;
    FLOAT       gmfCellIncX;
    FLOAT       gmfCellIncY;
} GLYPHMETRICSFLOAT;

//END GDI

// Array related

    typedef struct TYPEDESC {
        LONG pTypeDesc;
        short vt;
    } TYPEDESC;

    typedef struct SAFEARRAYBOUND {
        LONG cElements;
        LONG lLbound;
    } SAFEARRAYBOUND;

    typedef struct ARRAYDESC {
        TYPEDESC tdescElem;
        short cDims;
    } ARRAYDESC;

// User + Kernel

// Queued message structure
typedef struct MSG {
    long		hwnd;
    long		message;
    long		wParam;
    long		lParam; // LPARAM  
    long		time;
    POINTAPI	pt;
} MSG;

typedef struct WNDCLASSEX {
    long        cbSize;
    // Win 3.x /
    long         style;
    long         lpfnWndProc;
    long         cbClsExtra;
    long         cbWndExtra;
    long		 hInstance;
    long	     hIcon;
    long		 hCursor;
    long		 hbrBackground;
    long	     lpszMenuName;
    long	     lpszClassName;
    // Win 4.0 
    long	     hIconSm;
} WNDCLASSEX;

typedef struct WNDCLASS {
    long	     style;
    long	     lpfnWndProc;
    long         cbClsExtra;
    long         cbWndExtra;
    long		 hInstance;
    long		 hIcon;
    long		 hCursor;
    long		 hbrBackground;
    long		 lpszMenuName;
    long		 lpszClassName;
} WNDCLASS;

typedef struct MINMAXINFO {
    POINTAPI  ptReserved;
    POINTAPI  ptMaxSize;
    POINTAPI  ptMaxPosition;
    POINTAPI  ptMinTrackSize;
    POINTAPI  ptMaxTrackSize;
} MINMAXINFO;

typedef struct COPYDATASTRUCT {
    long dwData;
    long cbData;
    long lpData; // PVOID
} COPYDATASTRUCT;

typedef struct ICONINFO {
    long    fIcon;
    long	xHotspot;
    long	yHotspot;
    long	hbmMask;
    long	hbmColor;
} ICONINFO;


// Icon/Cursor header
typedef struct CURSORSHAPE
{
    long			xHotSpot;
    long			yHotSpot;
    long			cx;
    long			cy;
    long			cbWidth;
    BYTE		    Planes;
    BYTE		    BitsPixel;
} CURSORSHAPE;

typedef struct MDINEXTMENU {
    long   hmenuIn;
    long   hmenuNext;
    long   hwndNext;
} MDINEXTMENU;

//Used with SendInput
typedef struct KEYBDINPUT {
    long wVk;
    long wScan;
    long dwFlags;
    long time;
    long dwExtraInfo;
} KEYBDINPUT;

//Used with SendInput
typedef struct MOUSEINPUT {
    LONG dx;
    LONG dy;
    long mouseData;
    long dwFlags;
    long time;
    long dwExtraInfo;
} MOUSEINPUT;

//Used with SendInput
typedef struct HARDWAREINPUT {
    long uMsg;
    long wParamL;
    long wParamH;
} HARDWAREINPUT;

//Used with SendInput
typedef struct INPUT { 
  long			type; 
  MOUSEINPUT	mi; 
  KEYBDINPUT	ki;
  HARDWAREINPUT hi;
}INPUT;

// WM_WINDOWPOSCHANGING/CHANGED struct pointed to by lParam
typedef struct WINDOWPOS {
    long     hwnd;
    long     hwndInsertAfter;
    long     x;
    long     y;
    long     cx;
    long     cy;
    long     flags;
} WINDOWPOS;

// WM_NCCALCSIZE parameter structure
typedef struct NCCALCSIZE_PARAMS {
    RECT       rgrc[3];
    long	   lppos;
} NCCALCSIZE_PARAMS;

typedef struct CREATESTRUCT {
    long	lpCreateParams; // void *
    long	hInstance;
    long    hMenu;
    long    hwndParent;
    long    cy;
    long    cx;
    long    y;
    long    x;
    long    style;
    long	lpszName; //LPCSTR
    long	lpszClass; //LPCSTR
    long    dwExStyle;
} CREATESTRUCT;

typedef struct WINDOWPLACEMENT {
    long  Length;
    long  Flags;
    long  ShowCmd;
    POINTAPI  ptMinPosition;
    POINTAPI  ptMaxPosition;
    RECT  rcNormalPosition;
} WINDOWPLACEMENT;

typedef struct NMHDR {
	long	hwndFrom;
	long	idFrom;
	long	code;
} NMHDR;

// MEASUREITEMSTRUCT for ownerdraw
typedef struct MEASUREITEMSTRUCT {
    long       CtlType;
    long       CtlID;
    long       itemID;
    long       itemWidth;
    long       itemHeight;
    long       itemData;
} MEASUREITEMSTRUCT;

// DRAWITEMSTRUCT for ownerdraw
typedef struct DRAWITEMSTRUCT {
    long        CtlType;
    long        CtlID;
    long        itemID;
    long        itemAction;
    long        itemState;
    long        hwndItem;
    long        hDC;
    RECT        rcItem;
    long       itemData;
} DRAWITEMSTRUCT;

// DELETEITEMSTRUCT for ownerdraw
typedef struct DELETEITEMSTRUCT {
    long       CtlType;
    long       CtlID;
    long       itemID;
    long       hwndItem;
    long       itemData;
} DELETEITEMSTRUCT;

// COMPAREITEMSTUCT for ownerdraw sorting
typedef struct COMPAREITEMSTRUCT {
    long        CtlType;
    long        CtlID;
    long        hwndItem;
    long        itemID1;
    long        itemData1;
    long        itemID2;
    long        itemData2;
    long        dwLocaleId;
} COMPAREITEMSTRUCT;

typedef struct BROADCASTSYSMSG {
    long  uiMessage;
    long  wParam;
    long  lParam;
}BROADCASTSYSMSG;

typedef struct TPMPARAMS {
    long    cbSize;     // Size of structure 
    RECT    rcExclude;  // Screen coordinates of rectangle to exclude when positioning 
}TPMPARAMS;

typedef struct MENUITEMINFO {
    long    cbSize;
    long    fMask;
    long    fType;          // used if MIIM_TYPE (4.0) or MIIM_STRING (>4.0)
    long    fState;         // used if MIIM_STATE
    long    wID;            // used if MIIM_ID
    long	hSubMenu;       // used if MIIM_SUBMENU
    long	hbmpChecked;    // used if MIIM_CHECKMARKS
    long	hbmpUnchecked;  // used if MIIM_CHECKMARKS
    long	dwItemData;     // used if MIIM_DATA 
    long	dwTypeData;     // LPTSTR/LPTWSTR used if MIIM_TYPE (4.0) or MIIM_STRING (>4.0)
    long    cch;            // used if MIIM_TYPE (4.0) or MIIM_STRING (>4.0)
//#if (_WIN32_WINNT >= 0x0500)
//    HBITMAP hbmpItem;       // used if MIIM_BITMAP
}   MENUITEMINFO;

typedef struct MENUINFO {
  DWORD   cbSize;
  DWORD   fMask;
  DWORD   dwStyle;
  UINT    cyMax;
  long    hbrBack;
  DWORD   dwContextHelpID;
  long    dwMenuData;
} MENUINFO;


typedef struct DROPSTRUCT {
    long		hwndSource;
    long		hwndSink;
    long		wFmt;
    long		dwData;
    POINTAPI	ptDrop;
    long		dwControlData;
} DROPSTRUCT;

typedef struct DRAWTEXTPARAMS {
    long     cbSize;
    long     iTabLength;
    long     iLeftMargin;
    long     iRightMargin;
    long     uiLengthDrawn;
} DRAWTEXTPARAMS;
//typedef DRAWTEXTPARAMS * LPDRAWTEXTPARAMS;

typedef struct MSGBOXPARAMS {
    long        cbSize;
    long        hwndOwner;
    long        hInstance;
    long		lpszText; /* LPCTSTR/LPCWSTR */
    long		lpszCaption; /* LPCTSTR/LPCWSTR */
    long		dwStyle;
    long		lpszIcon; /* LPCTSTR/LPCWSTR */
    long		dwContextHelpId;
    long		lpfnMsgBoxCallback;
    long		dwLanguageId;
} MSGBOXPARAMS;

//Registry related

typedef struct SECURITY_ATTRIBUTES {  
	long nLength;
	long lpSecurityDescriptor;
	BOOL bInheritHandle;
} SECURITY_ATTRIBUTES;

// ******* Kernel related

typedef struct FILETIME {
    long dwLowDateTime;
    long dwHighDateTime;
} FILETIME;

typedef struct SYSTEMTIME {
    short wYear;
    short wMonth;
    short wDayOfWeek;
    short wDay;
    short wHour;
    short wMinute;
    short wSecond;
    short wMilliseconds;
} SYSTEMTIME;

typedef struct A_WIN32_FIND_DATA {
    long dwFileAttributes;
    CURRENCY ftCreationTime;
    CURRENCY ftLastAccessTime;
    CURRENCY ftLastWriteTime;
    long nFileSizeHigh;
    long nFileSizeLow;
    long dwReserved0;
    long dwReserved1;
    BYTE cFileName[260];
    BYTE cAlternateFileName[14];
} A_WIN32_FIND_DATA;

typedef struct W_WIN32_FIND_DATA {
    long dwFileAttributes;
    CURRENCY ftCreationTime;
    CURRENCY ftLastAccessTime;
    CURRENCY ftLastWriteTime;
    long nFileSizeHigh;
    long nFileSizeLow;
    long dwReserved0;
    long dwReserved1;
    BYTE cFileName[520];
    BYTE cAlternateFileName[28];
} W_WIN32_FIND_DATA;

typedef struct AFT_WIN32_FIND_DATA {
    long dwFileAttributes;
    FILETIME ftCreationTime;
    FILETIME ftLastAccessTime;
    FILETIME ftLastWriteTime;
    long nFileSizeHigh;
    long nFileSizeLow;
    long dwReserved0;
    long dwReserved1;
    BYTE cFileName[260];
    BYTE cAlternateFileName[14];
} AFT_WIN32_FIND_DATA;

typedef struct WFT_WIN32_FIND_DATA {
    long dwFileAttributes;
    FILETIME ftCreationTime;
    FILETIME ftLastAccessTime;
    FILETIME ftLastWriteTime;
    long nFileSizeHigh;
    long nFileSizeLow;
    long dwReserved0;
    long dwReserved1;
    BYTE cFileName[520];
    BYTE cAlternateFileName[28];
} WFT_WIN32_FIND_DATAOW;

typedef struct A_OSVERSIONINFO {
    long dwOSVersionInfoSize;
    long dwMajorVersion;
    long dwMinorVersion;
    long dwBuildNumber;
    long dwPlatformId;
    BYTE szCSDVersion[128];     // Maintenance string for PSS usage
} A_OSVERSIONINFO;

typedef struct W_OSVERSIONINFO {
    long dwOSVersionInfoSize;
    long dwMajorVersion;
    long dwMinorVersion;
    long dwBuildNumber;
    long dwPlatformId;
    BYTE szCSDVersion[256];	// Maintenance string for PSS usage
} W_OSVERSIONINFO;

typedef struct STARTUPINFO {
    long   cb;
    long  lpzReserved; //LPSTR
    long  lpzDesktop; //LPSTR
    long  lpzTitle; //LPSTR
    long   dwX;
    long   dwY;
    long   dwXSize;
    long   dwYSize;
    long   dwXCountChars;
    long   dwYCountChars;
    long   dwFillAttribute;
    long   dwFlags;
    short  wShowWindow;
	short  cbReserved2;
    BYTE   lpReserved2;
    long   hStdInput;
    long   hStdOutput;
    long   hStdError;
} STARTUPINFO;


typedef struct BY_HANDLE_FILE_INFORMATION {
    long dwFileAttributes;
    CURRENCY ftCreationTime;
    CURRENCY ftLastAccessTime;
    CURRENCY ftLastWriteTime;
    long dwVolumeSerialNumber;
    long nFileSizeHigh;
    long nFileSizeLow;
    long nNumberOfLinks;
    long nFileIndexHigh;
    long nFileIndexLow;
} BY_HANDLE_FILE_INFORMATION;

typedef struct BY_HANDLE_FILE_INFORMATIONO {
    long dwFileAttributes;
    FILETIME ftCreationTime;
    FILETIME ftLastAccessTime;
    FILETIME ftLastWriteTime;
    long dwVolumeSerialNumber;
    long nFileSizeHigh;
    long nFileSizeLow;
    long nNumberOfLinks;
    long nFileIndexHigh;
    long nFileIndexLow;
} BY_HANDLE_FILE_INFORMATIONO;

typedef struct A_TIME_ZONE_INFORMATION {
    long Bias;
    BYTE StandardName[32];
    SYSTEMTIME StandardDate;
    long StandardBias;
    BYTE DaylightName[32];
    SYSTEMTIME DaylightDate;
    long DaylightBias;
} A_TIME_ZONE_INFORMATION;

typedef struct W_TIME_ZONE_INFORMATION {
    long Bias;
    BYTE StandardName[64];
    SYSTEMTIME StandardDate;
    long StandardBias;
    BYTE DaylightName[64];
    SYSTEMTIME DaylightDate;
    LONG DaylightBias;
} W_TIME_ZONE_INFORMATION;

typedef struct A_OFSTRUCT {
    BYTE cBytes;
    BYTE fFixedDisk;
    short nErrCode;
    short Reserved1;
    short Reserved2;
    BYTE szPathName[128]; // OFS_MAXPATHNAME
} A_OFSTRUCT;

typedef struct W_OFSTRUCT {
    BYTE cBytes;
    BYTE fFixedDisk;
    short nErrCode;
    short Reserved1;
    short Reserved2;
    BYTE szPathName[256]; // OFS_MAXPATHNAME
} W_OFSTRUCT;


typedef struct SYSTEM_INFO {
    // DWORD dwOemId;    // Obsolete, replaced by following:
    short	wProcessorArchitecture;
    short	wReserved;
    long	dwPageSize;
    long    lpMinimumApplicationAddress; // LPVOID 
    long    lpMaximumApplicationAddress; // LPVOID 
    long	dwActiveProcessorMask;
    long	dwNumberOfProcessors;
    long	dwProcessorType;
    long	dwAllocationGranularity;
    short	wProcessorLevel;
    short	wProcessorRevision;
} SYSTEM_INFO;

typedef struct MEMORYSTATUS {
    long dwLength;
    long dwMemoryLoad;
    long dwTotalPhys;
    long dwAvailPhys;
    long dwTotalPageFile;
    long dwAvailPageFile;
    long dwTotalVirtual;
    long dwAvailVirtual;
} MEMORYSTATUS;

typedef struct MEMORYSTATUSEXL {
    long dwLength;
    long dwMemoryLoad;
    LARGE_INTEGER ullTotalPhys;
    LARGE_INTEGER ullAvailPhys;
    LARGE_INTEGER ullTotalPageFile;
    LARGE_INTEGER ullAvailPageFile;
    LARGE_INTEGER ullTotalVirtual;
    LARGE_INTEGER ullAvailVirtual;
    LARGE_INTEGER ullAvailExtendedVirtual;
} MEMORYSTATUSEXL;

typedef struct MEMORYSTATUSEX {
    long dwLength;
    long dwMemoryLoad;
    CURRENCY ullTotalPhys;
    CURRENCY ullAvailPhys;
    CURRENCY ullTotalPageFile;
    CURRENCY ullAvailPageFile;
    CURRENCY ullTotalVirtual;
    CURRENCY ullAvailVirtual;
    CURRENCY ullAvailExtendedVirtual;
} MEMORYSTATUSEX;

typedef struct OVERLAPPED {
    long   Internal;
    long   InternalHigh;
    long   Offset;
    long   OffsetHigh;
    long  hEvent;
} OVERLAPPED;

typedef struct PROCESS_INFORMATION {
	long hProcess;
	long hThread;
	long dwProcessId;
	long dwThreadId;
} PROCESS_INFORMATION;

typedef struct EXCEPTION_POINTERS {
     long pExceptionRecord;       //pointer to an EXCEPTION_RECORD struct
     long pContextRecord;         //pointer to a CONTEXT struct
} EXCEPTION_POINTERS;

typedef struct EXCEPTION_RECORD {
     long ExceptionCode;
     long ExceptionFlags;
     long pExceptionRecord;    //Pointer to an EXCEPTION_RECORD structure
     long ExceptionAddress;
     long NumberParameters;
     long ExceptionInformation[15];
} EXCEPTION_RECORD;


typedef struct FLASHWINFO { 
	UINT cbSize;
	HWND hwnd;
	DWORD dwFlags;
	UINT uCount;
	DWORD dwTimeout;
} FLASHWINFO; 

//
//  Number format.
//

typedef struct NUMBERFMT {
    UINT    NumDigits;                 // number of decimal digits
    UINT    LeadingZero;               // if leading zero in decimal fields
    UINT    Grouping;                  // group size left of decimal
    long   lpzDecimalSep;              // ptr to decimal separator string
    long   lpzThousandSep;             // ptr to thousand separator string
    UINT    NegativeOrder;             // negative number ordering
} NUMBERFMT;

//
//  Currency format.
//

typedef struct CURRENCYFMT {
    UINT    NumDigits;                 // number of decimal digits
    UINT    LeadingZero;               // if leading zero in decimal fields
    UINT    Grouping;                  // group size left of decimal
    long   lpzDecimalSep;              // ptr to decimal separator string
    long   lpzThousandSep;             // ptr to thousand separator string
    UINT    NegativeOrder;             // negative currency ordering
    UINT    PositiveOrder;             // positive currency ordering
    long   lpzCurrencySymbol;          // ptr to currency symbol string
} CURRENCYFMT;

// ********* Ole object related enums and structs
typedef struct UUID { //16 bytes
    LONG Data1;
    SHORT Data2;
    SHORT Data3;
    BYTE Data4[8];
} UUID;

typedef struct STRRET {
    ESTRRET uType;
    BYTE cStr[260];
} STRRET;

typedef struct STATSTG {
    LONG pwcsName;
    STGTY type;
    CURRENCY cbSize;
    CURRENCY mtime;
    CURRENCY ctime;
    CURRENCY atime;
    STGM grfMode;
    LOCKTYPE grfLocksSupported;
    UUID clsid;
    LONG grfStateBits;
    LONG reserved;
} STATSTG;

    typedef struct BIND_OPTS {
        LONG cbStruct;       //  sizeof(BIND_OPTS)
        BIND_FLAGS grfFlags;
        STGM grfMode;
        LONG dwTickCountDeadline;
    } BIND_OPTS;

    typedef struct BIND_OPTS2 {
        LONG cbStruct;       //  sizeof(BIND_OPTS2)
        BIND_FLAGS grfFlags;
        STGM grfMode;
        LONG dwTickCountDeadline;
        LONG dwTrackFlags;
        LONG dwClassContext;
        LONG locale;
        LONG pServerInfo;
    } BIND_OPTS2;

    typedef struct DISPPARAMS {
        LONG rgPointerToVariantArray;
        LONG rgPointerToLONGNamedArgs;
        LONG cArgs;
        LONG cNamedArgs;
    } DISPPARAMS;


    typedef [helpstring("DISP_E_EXCEPTION Error info")] struct EXCEPINFO {
        short wCode;
        short wReserved;
        BSTR Source;
        BSTR Description;
        BSTR HelpFile;
        LONG dwHelpContext;
        LONG pvReserved;
        LONG pfnDeferredFillIn;
        LONG scode;
    } EXCEPINFO;

typedef struct DVASPECTINFO {
    long cb;
    DVASPECTINFOFLAG dwFlags;
} DVASPECTINFO;

typedef struct FORMATETC {
    short cfFormat;
    long pDVTARGETDEVICE;
    DVASPECT dwAspect;
    long lindex;
    TYMED TYMED;
} FORMATETC;

typedef struct STGMEDIUM {
    TYMED TYMED;
    long data;
    long pUnkForRelease;
} STGMEDIUM;

typedef struct STATDATA {
    FORMATETC FORMATETC;
    ADVF grfAdvf;
    IAdviseSink *pIAdviseSink;
    long dwConnection;
} STATDATA;

typedef struct OLEVERB {
    long lVerb;
    long lpszVerbName;
    long fuFlags;
    OLEVERBATTRIB grfAttribs;
} OLEVERB;

typedef struct OLEMENUGROUPWIDTHS {
    LONG width[6];
} OLEMENUGROUPWIDTHS;

typedef struct OLEINPLACEFRAMEINFO {
    long cb;
    long fMDIApp;
    long hwndFrame;
    long haccel;
    long cAccelEntries;
} OLEINPLACEFRAMEINFO;


typedef struct CUSTDATA {
     [helpstring("number of custom data items in rgCustData")]
    LONG cCustData;
    LONG prgCustData;
} CUSTDATA;

    typedef struct IDLDESC {
        LONG dwReserved;
        short wIDLFlags;
    } IDLDESC;

    typedef struct PARAMDESCEX {
        LONG cBytes;
    LONG ignore;
        VARIANT varDefaultValue;
    } PARAMDESCEX;

    typedef struct PARAMDESC {
        [helpstring("valid if PARAMFLAG_FHASDEFAULT bit is set")]
        LONG pPARAMDESCEX;
        short wParamFlags;
    } PARAMDESC;

    typedef struct ELEMDESC {
        TYPEDESC tdesc;
        PARAMDESC paramdesc;
    } ELEMDESC;

typedef struct VARDESC {
    LONG memid;
    LONG lpstrSchema;
    LONG oInst_varValue;
    ELEMDESC elemdescVar;
    short wVarFlags;
    VARKIND varkind;
} VARDESC;

typedef struct FUNCDESC {
    LONG memid;
    LONG lprgscode;
     [helpstring("array of param types")]
    LONG lprgELEMDESCParam;
    FUNCKIND funckind;
    INVOKEKIND invkind;
    CALLCONV callconv;
    short cParams;
    short cParamsOpt;
    short oVft;
    short cScodes;
    ELEMDESC elemdescFunc;
    short wFuncFlags;
} FUNCDESC;

// ******** SHLWAPI related

typedef struct DLLVERSIONINFO {
    LONG cbSize;
    LONG dwMajorVersion; // Major version
    LONG dwMinorVersion; // Minor version
    LONG dwBuildNumber; // Build number
    LONG dwPlatformID; // DLLVER_PLATFORM_*
} DLLVERSIONINFO;

typedef struct {
    DLLVERSIONINFO info1;
    LONG dwFlags; // No flags currently defined
    CURRENCY ullVersion; // Encoded as:
    // Major 0xFFFF 0000 0000 0000
    // Minor 0x0000 FFFF 0000 0000
    // Build 0x0000 0000 FFFF 0000
    // QFE 0x0000 0000 0000 FFFF
} DLLVERSIONINFO2;

// ********** SHELL32 related

//
// SHITEMID -- Item ID
//
typedef struct SHITEMID        // mkid
{
    long	    cb;             // Size of the ID (including cb itself)
    BYTE        abID[1];        // The item ID (variable length)
} SHITEMID;

//
// ITEMIDLIST -- List if item IDs (combined with 0-terminator)
//
typedef struct ITEMIDLIST      // idl
{
    SHITEMID    mkid;
} ITEMIDLIST;

typedef struct A_SHFILEINFO {
        long        hIcon;                      // out: icon
        long        iIcon;                      // out: icon index
        DWORD       dwAttributes;               // out: SFGAO_ flags
        BYTE        szDisplayName[260];         // out: display name (or path)
        BYTE        szTypeName[80];             // out: type name
} A_SHFILEINFO;

typedef struct W_SHFILEINFO {
        long        hIcon;                      // out: icon
        long        iIcon;                      // out: icon index
        DWORD       dwAttributes;               // out: SFGAO_ flags
        BYTE        szDisplayName[520];     // out: display name (or path)
        BYTE        szTypeName[160];         // out: type name
} W_SHFILEINFO;

typedef struct SHFILEOPSTRUCT {
        HWND            hwnd;
        UINT            wFunc;
        LONG            lpszFrom;
        LONG            lpszTo;
        WORD		    fFlags;
        BOOL            fAnyOperationsAborted;
        long 			hNameMappings;	// LPVOID          
        long			lpszProgressTitle; // only used if FOF_SIMPLEPROGRESS
} SHFILEOPSTRUCT;

typedef struct A_NOTIFYICONDATA {
        DWORD cbSize;
        HWND  hWnd;
        UINT  uID;
        UINT  uFlags;
        UINT  uCallbackMessage;
        long hIcon;
        BYTE  szTip[64];
} A_NOTIFYICONDATA;

typedef struct W_NOTIFYICONDATA {
        DWORD cbSize;
        HWND  hWnd;
        UINT  uID;
        UINT  uFlags;
        UINT  uCallbackMessage;
        long hIcon;
        BYTE  szTip[128];
} W_NOTIFYICONDATA;


typedef struct BROWSEINFO {
    long        hwndOwner;
    long		pidlRoot;
    long		lpszDisplayName;// Return display name of item selected.
    long		lpszTitle; // text to go in the banner over the tree.
    UINT        ulFlags;   // Flags that control the return stuff
    LONG		lpfn;
    LONG        lParam;    // extra info that's passed back in callbacks
    long        iImage;    // output var: where to return the Image index.
} BROWSEINFO;

    typedef struct SHELLEXECUTEINFO {
            long	cbSize;
            long	fMask;
            long	hwnd;
            long   lpzVerb;
            long   lpzFile;
            long   lpzParameters;
            long   lpzDirectory;
            long    nShow;
            long	hInstApp;
            long    lpIDList;
            long   lpzClass;
            long	hkeyClass;
            long	dwHotKey;
            long	hIcon;
            long	hProcess;
    } SHELLEXECUTEINFO;

//
// *************** Network Related
//

typedef struct  NETRESOURCE {
    DWORD   dwScope;
    DWORD   dwType;
    DWORD   dwDisplayType;
    DWORD   dwUsage;
    long	lpLocalName; // LPSTR  LPWSTR
    long	lpRemoteName; // LPSTR  LPWSTR
    long	lpComment ; // LPSTR  LPWSTR
    long	lpProvider; // LPSTR  LPWSTR
} NETRESOURCE;

typedef struct CONNECTDLGSTRUCT {
    DWORD cbStructure;       // size of this structure in bytes
    HWND hwndOwner;          // owner window for the dialog
    // DWORD NETRESOURCE lpConnRes; // Requested Resource info
    DWORD dwFlags;           // flags (see below)
    DWORD dwDevNum;          // number of devices connected to
} CONNECTDLGSTRUCT;

typedef struct UNIVERSAL_NAME_INFO {
    long lpUniversalName; // LPSTR  LPWSTR
} UNIVERSAL_NAME_INFO;

typedef struct REMOTE_NAME_INFO {
    long lpUniversalName; // LPSTR  LPWSTR
    long lpConnectionName; // LPSTR  LPWSTR
    long lpRemainingPath; // LPSTR  LPWSTR
} REMOTE_NAME_INFO;

typedef struct NETINFOSTRUCT {
    DWORD cbStructure;
    DWORD dwProviderVersion;
    DWORD dwStatus;
    DWORD dwCharacteristics;
    DWORD dwHandle;
    WORD  wNetType;
    DWORD dwPrinters;
    DWORD dwDrives;
} NETINFOSTRUCT;

typedef struct DISCDLGSTRUCT {
    DWORD cbStructure;      // size of this structure in bytes
    HWND  hwndOwner;        // owner window for the dialog
    long  lpLocalName;      // LPSTR  LPWSTR local device name
    long  lpRemoteName;     // LPSTR  LPWSTR network resource name
    DWORD dwFlags;          // flags
} DISCDLGSTRUCT;

typedef struct NETCONNECTINFOSTRUCT {
    DWORD cbStructure;
    DWORD dwFlags;
    DWORD dwSpeed;
    DWORD dwDelay;
    DWORD dwOptDataSize;
} NETCONNECTINFOSTRUCT;


//
// INTERNET_SCHEME - enumerated URL scheme type
//

typedef struct {
    DWORD dwResult;
    DWORD dwError;
} INTERNET_ASYNC_RESULT;

typedef struct {
    DWORD dwAccessType;
    long lpszProxy;			//LPCTSTR
    long lpszProxyBypass;	//LPCTSTR
} INTERNET_PROXY_INFO;

typedef struct {
    DWORD dwMajorVersion;
    DWORD dwMinorVersion;
} INTERNET_VERSION_INFO;

typedef struct {
    DWORD dwMajorVersion;
    DWORD dwMinorVersion;
} HTTP_VERSION_INFO;

typedef struct {
    DWORD dwConnectedState;
    DWORD dwFlags;
} INTERNET_CONNECTED_INFO;

typedef struct {
    DWORD   dwStructSize;       // size of this structure. Used in version check
    long    lpszScheme;         //LPSTR LPWSTR pointer to scheme name
    DWORD   dwSchemeLength;     // length of scheme name
    long    nScheme;            //INTERNET_SCHEME enumerated scheme type (if known)
    long    lpszHostName;       //LPSTR LPWSTR pointer to host name
    DWORD   dwHostNameLength;   // length of host name
    INTERNET_PORT    nPort;              // converted port number
    long    lpszUserName;       //LPSTR LPWSTR pointer to user name
    DWORD   dwUserNameLength;   // length of user name
    long    lpszPassword;       //LPSTR LPWSTR pointer to password
    DWORD   dwPasswordLength;   // length of password
    long    lpszUrlPath;        //LPSTR LPWSTR pointer to URL-path
    DWORD   dwUrlPathLength;    // length of URL-path
    long   lpszExtraInfo;       //LPSTR LPWSTR pointer to extra information (e.g. ?foo or #foo)
    DWORD   dwExtraInfoLength;  // length of extra information
} URL_COMPONENTS;

typedef struct {
    FILETIME ftExpiry;
    FILETIME ftStart;
    long lpszSubjectInfo;	//LPTSTR LPWTSTR
    long lpszIssuerInfo;	//LPTSTR LPWTSTR
    long lpszProtocolName;	//LPTSTR LPWTSTR
    long lpszSignatureAlgName;	//LPTSTR LPWTSTR
    long lpszEncryptionAlgName;	//LPTSTR LPWTSTR
    DWORD dwKeySize;
} INTERNET_CERTIFICATE_INFO;


//
// INTERNET_BUFFERS - combines headers and data. May be chained for e.g. file
// upload or scatter/gather operations. For chunked read/write, lpcszHeader
// contains the chunked-ext
//

typedef struct INTERNET_BUFFERS {
    DWORD	dwStructSize;               // used for API versioning. Set to sizeof(INTERNET_BUFFERS)
    long *  Next;						// chain of buffers, struct _INTERNET_BUFFERSA
    long *  lpcszHeader;				//LPTSTR LPWTSTR pointer to headers (may be NULL)
    DWORD	dwHeadersLength;            // length of headers if not NULL
    DWORD	dwHeadersTotal;             // size of headers if not enough buffer
    LPVOID	lpvBuffer;                  // pointer to data buffer (may be NULL)
    DWORD	dwBufferLength;             // length of data buffer if not NULL
    DWORD	dwBufferTotal;              // total size of chunk, or content-length if not chunked
    DWORD	dwOffsetLow;                // used for read-ranges (only used in HttpSendRequest2)
    DWORD	dwOffsetHigh;
} INTERNET_BUFFERS;

// If SERIALIZE_DIALOGS flag set, client should implement thread-safe non-blocking callback...

//DWORD InternetAuthNotifyCallback
//(
//    DWORD           dwContext,    // as passed to InternetErrorDlg
//    DWORD           dwReturn,     // error code: success, resend, or cancel
//    LPVOID          lpReserved    // reserved: will be set to null
//);
//typedef DWORD (CALLBACK * PFN_AUTH_NOTIFY) (DWORD, DWORD, LPVOID);

// ... and last parameter of InternetErrorDlg should point to...

typedef struct INTERNET_AUTH_NOTIFY_DATA {
    DWORD            cbStruct;    // size of this structure
    DWORD            dwOptions;   // reserved: must set to 0
    long			 pfnNotify;   // notification callback to retry InternetErrorDlg
    DWORD            dwContext;   // context to pass to to notification function
} INTERNET_AUTH_NOTIFY_DATA;

//
// INTERNET_CACHE_ENTRY_INFO -
//


typedef struct INTERNET_CACHE_ENTRY_INFO {
    DWORD dwStructSize;         // version of cache system.
    long  lpszSourceUrlName;    // embedded pointer to the URL name string.
    long  lpszLocalFileName;    // embedded pointer to the local file name.
    DWORD CacheEntryType;       // cache type bit mask.
    DWORD dwUseCount;           // current users count of the cache entry.
    DWORD dwHitRate;            // num of times the cache entry was retrieved.
    DWORD dwSizeLow;            // low DWORD of the file size.
    DWORD dwSizeHigh;           // high DWORD of the file size.
    FILETIME LastModifiedTime;  // last modified time of the file in GMT format.
    FILETIME ExpireTime;        // expire time of the file in GMT format
    FILETIME LastAccessTime;    // last accessed time in GMT format
    FILETIME LastSyncTime;      // last time the URL was synchronized
                                // with the source
    long  lpHeaderInfo;         // embedded pointer to the header info.
    DWORD dwHeaderInfoSize;     // size of the above header.
    long  lpszFileExtension;    // File extension used to retrive the urldata as a file.
    DWORD dwReserved;
    DWORD dwExemptDelta;
} INTERNET_CACHE_ENTRY_INFO;

typedef struct INITCOMMONCONTROLSEX {
    DWORD dwSize;             // size of this structure
    DWORD dwICC;              // flags indicating which classes to be initialized
} INITCOMMONCONTROLSEX;

/*
Private Const CCM_FIRST = 8192 &H2000
Private Const CCM_LAST = (8192 + 512)
Private Const CCM_SETVERSION = (8192 + 7)
Private Const CCM_GETCOLORSCHEME = (8192 + 3)
Private Const CCM_GETDROPTARGET = (8192 + 4)
Private Const CCM_GETUNICODEFORMAT = (8192 + 6)
Private Const CCM_GETVERSION = (8192 + 8)
Private Const CCM_SETBKCOLOR = (8192 + 1)
Private Const CCM_SETCOLORSCHEME = (8192 + 2)
Private Const CCM_SETNOTIFYWINDOW = (8192 + 9)
Private Const CCM_SETUNICODEFORMAT = (8192 + 5)
Private Const CCM_SETWINDOWTHEME = (8192 + 11)

Private Const CCS_ADJUSTABLE = 32
Private Const CCS_BOTTOM = 3 &H3&
Private Const CCS_LEFT = 129 (CCS_VERT Or CCS_TOP)
Private Const CCS_NODIVIDER = 64 &H40&
Private Const CCS_NOMOVEX = 130 (CCS_VERT Or CCS_NOMOVEY)
Private Const CCS_NOMOVEY = 2 &H2&
Private Const CCS_NOPARENTALIGN = 8 &H8&
Private Const CCS_NORESIZE = 4 &H4&
Private Const CCS_RIGHT = 131 (CCS_VERT Or CCS_BOTTOM)
Private Const CCS_TOP = 1 &H1&
Private Const CCS_VERT = 128 &H80&
*/

    typedef struct COLORSCHEME {
        LONG dwSize;
        LONG clrBtnHighlight; // highlight color
        LONG clrBtnShadow; // shadow color
    } COLORSCHEME;

    typedef struct NMTOOLTIPSCREATED {
        NMHDR hdr;
        LONG hwndToolTips;
    } NMTOOLTIPSCREATED;

    typedef struct NMMOUSE {
        NMHDR   hdr;
        LONG dwItemSpec;
        LONG dwItemData;
        POINTAPI pt;
        LONG dwHitInfo;
    } NMMOUSE;

	typedef NMMOUSE NMCLICK;

    typedef struct NMOBJECTNOTIFY {
        NMHDR   hdr;
        LONG iItem;
        LONG piid;
        LONG pObject;
        LONG hResult;
        LONG dwFlags;
    } NMOBJECTNOTIFY;

    typedef struct NMKEY {
        NMHDR hdr;
        short wVKey;
        LONG flags;
    } NMKEY;

    typedef struct NMCHAR {
        NMHDR hdr;
        LONG ch;
        LONG dwItemPrev; // Item previously selected
        LONG dwItemNext; // Item to be selected
    } NMCHAR;

    typedef struct NMCUSTOMDRAW {
        NMHDR hdr;
        LONG dwDrawStage;
        LONG hdc;
        RECT rc;
        LONG dwItemSpec;
        ItemStateFlags uItemState;
        LONG lParam;
    } NMCUSTOMDRAW;

    typedef struct NMTTCUSTOMDRAW {
        NMCUSTOMDRAW nmcd;
        LONG uDrawFlags;
    } NMTTCUSTOMDRAW;

 // IMAGELIST

    typedef struct IMAGELISTDRAWPARAMSv5 {
        LONG cbSize;
        LONG himl;
        LONG i;
        LONG hdcDst;
        LONG x;
        LONG y;
        LONG cx;
        LONG cy;
        LONG xBitmap;
        LONG yBitmap;
        LONG rgbBk;
        LONG rgbFg;
        LONG fStyle;
        LONG dwRop;
    } IMAGELISTDRAWPARAMSv5;


    typedef struct IMAGELISTDRAWPARAMSv6 {
        DWORD cbSize;
        HIMAGELIST himl;
        LONG i;
        HDC hdcDst;
        LONG x;
        LONG y;
        LONG cx;
        LONG cy;
        LONG xBitmap;
        LONG yBitmap;
        COLORREF rgbBk;
        COLORREF rgbFg;
        UINT fStyle;
        DWORD dwRop;
        DWORD fState;
        DWORD Frame;
        DWORD crEffect;
    } IMAGELISTDRAWPARAMSv6;

    typedef struct IMAGEINFO {
        LONG hbmImage;
        LONG hbmMask;
        LONG Unused1;
        LONG Unused2;
        RECT rcImage;
    } IMAGEINFO;

    typedef struct HD_TEXTFILTER {
        LONG pszText; // [in] pointer to the buffer containing the filter (ANSI)
        LONG cchTextMax; // [in] max size of buffer/edit control buffer
    } HD_TEXTFILTER;

    typedef struct HDITEM {
        HeaderItemMask    mask;
        LONG cxy;
        LONG pszText;
        LONG hbm;
        LONG cchTextMax;
        LONG fmt;
        LONG lParam;
        LONG iImage; // index of bitmap in ImageList
        LONG iOrder; // where to draw this item
    } HDITEM;

    typedef struct HDITEM_V5 {
        HeaderItemMask    mask;
        LONG cxy;
        LONG pszText;
        LONG hbm;
        LONG cchTextMax;
        LONG fmt;
        LONG lParam;
        LONG iImage; // index of bitmap in ImageList
        LONG iOrder; // where to draw this item
        LONG type;
        LONG pvFilter;
    } HDITEM_V5;

    typedef struct HDLAYOUT {
        LONG pRECT;
        LONG pWINDOWPOS;
    } HDLAYOUT;

    typedef struct HDHITTESTINFO {
        POINTAPI pt;
        LONG flags;
        LONG iItem;
    } HDHITTESTINFO;

    typedef struct NMHEADER {
        NMHDR   hdr;
        LONG iItem;
        LONG iButton;
        LONG pHDITEM;
    }  NMHEADER;

    typedef struct NMHDDISPINFO {
        NMHDR hdr;
        LONG iItem;
        LONG mask;
        LONG pszText;
        LONG cchTextMax;
        LONG iImage;
        LONG lParam;
    } NMHDDISPINFO;

    typedef struct NMHDFILTERBTNCLICK {
        NMHDR hdr;
        LONG iItem;
        RECT rc;
    } NMHDFILTERBTNCLICK;

 // = = ====TOOLBAR CONTROL = = ====================================================

    typedef struct TBBUTTON {
        LONG iBitmap;
        LONG idCommand;
        BYTE fsState;
        BYTE fsStyle;
        BYTE bReserved[2];
        LONG dwData;
        LONG iString;
    } TBBUTTON;

    typedef struct COLORMAP {
        LONG from;
        LONG to;
    } COLORMAP;

    typedef struct NMTBCUSTOMDRAW {
        NMCUSTOMDRAW nmcd;
        LONG hbrMonoDither;
        LONG hbrLines;
        LONG hpenLines;
        LONG clrText; // Color of text
        LONG clrMark; // Color of text bk when marked. (only if TBSTATE_MARKED)
        LONG clrTextHighlight; // Color of text when highlighted
        LONG clrBtnFace; // Background of the button
        LONG clrBtnHighlight; // 3D highlight
        LONG clrHighlightHotTrack; // In conjunction with fHighlightHotTrack
 // will cause button to highlight like a menu
        RECT rcText; // Rect for text
        LONG nStringBkMode;
        LONG nHLStringBkMode;
    } NMTBCUSTOMDRAW;

    typedef struct TBADDBITMAP {
        LONG hInst;
        LONG nID;
    } TBADDBITMAP;

    typedef struct TBSAVEPARAMS {
        LONG hkr;
        LONG pszSubKey;
        LONG pszValueName;
    } TBSAVEPARAMS;

    typedef struct TBINSERTMARK {
        LONG iButton;
        LONG dwFlags;
    } TBINSERTMARK;

    typedef struct TBREPLACEBITMAP {
        LONG hInstOld;
        LONG nIDOld;
        LONG hInstNew;
        LONG nIDNew;
        LONG nButtons;
    } TBREPLACEBITMAP;

    typedef struct TBBUTTONINFO {
        LONG cbSize;
        LONG dwMask;
        LONG idCommand;
        LONG iImage;
        BYTE fsState;
        BYTE fsStyle;
        short cx;
        LONG lParam;
        LONG pszText;
        LONG cchText;
    } TBBUTTONINFO;

 // Hot item change flags

    typedef struct NMTBHOTITEM {
        NMHDR   hdr;
        LONG idOld;
        LONG idNew;
        ToolbarHotItemFlags dwFlags;
    } NMTBHOTITEM;

    typedef struct NMTBDISPINFO {
        NMHDR  hdr;
        LONG dwMask; // [in] Specifies the values requested .[out] Client ask the data to be set for future use
        LONG idCommand; // [in] id of button we're requesting info for
        LONG lParam; // [in] lParam of button
        LONG iImage; // [out] image index
        LONG pszText; // [out] new text for item
        LONG cchText; // [in] size of buffer pointed to by pszText
    } NMTBDISPINFO;

    typedef struct NMTBSAVE {
        NMHDR hdr;
        LONG pData;
        LONG pCurrent;
        LONG cbData;
        LONG iItem;
        LONG cButtons;
        TBBUTTON tbButton;
    } NMTBSAVE;

    typedef struct NMTBRESTORE {
        NMHDR hdr;
        LONG pData;
        LONG pCurrent;
        LONG cbData;
        LONG iItem;
        LONG cButtons;
        LONG cbBytesPerRecord;
        TBBUTTON tbButton;
    } NMTBRESTORE;

    typedef struct NMTBGETINFOTIP {
        NMHDR hdr;
        LONG pszText;
        LONG cchTextMax;
        LONG iItem;
        LONG lParam;
    } NMTBGETINFOTIP;

    typedef struct NMTOOLBAR {
        NMHDR hdr;
        LONG iItem;
        TBBUTTON tbButton;
        LONG cchText;
        LONG pszText;
    } NMTOOLBAR;

// Rebar control

    typedef struct REBARINFO {
        LONG cbSize;
        LONG fMask;
        LONG himl;
    } REBARINFO;

    typedef struct REBARBANDINFO {
        LONG cbSize;
        ReBarBandInfoMask    fMask;
        LONG fStyle;
        LONG clrFore;
        LONG clrBack;
        LONG lpText;
        LONG cch;
        LONG iImage;
        LONG hwndChild;
        LONG cxMinChild;
        LONG cyMinChild;
        LONG cx;
        LONG hbmBack;
        LONG wID;
        LONG cyChild;
        LONG cyMaxChild;
        LONG cyIntegral;
        LONG cxIdeal;
        LONG lParam;
        LONG cxHeader;
    } REBARBANDINFO;

    typedef struct NMREBARCHILDSIZE {
        NMHDR hdr;
        LONG uBand;
        LONG wID;
        RECT rcChild;
        RECT rcBand;
    } NMREBARCHILDSIZE;

    typedef struct NMREBAR {
        NMHDR   hdr;
        NMREBARMasks dwMask; // RBNM_*
        LONG uBand;
        LONG fStyle;
        LONG wID;
        LONG lParam;
    } NMREBAR;

    typedef struct NMRBAUTOSIZE {
        NMHDR hdr;
        BOOL fChanged;
        RECT rcTarget;
        RECT rcActual;
    } NMRBAUTOSIZE;

    typedef struct NMREBARCHEVRON {
        NMHDR hdr;
        LONG uBand;
        LONG wID;
        LONG lParam;
        RECT rc;
        LONG lParamNM;
    } NMREBARCHEVRON;

    typedef struct RBHITTESTINFO {
        POINTAPI pt;
        ReBarHitTestFlags flags;
        LONG iBand;
    } RBHITTESTINFO;

// ********* Tooltip Control


    typedef struct TTTOOLINFO {
        LONG cbSize;
        LONG uFlags;
        LONG hwnd;
        LONG uId;
        RECT rect;
        LONG hinst;
        LONG lpszText;
        LONG lParam;
    } TTTOOLINFO;

    typedef struct TTHITTESTINFO {
        LONG hwnd;
        POINTAPI pt;
        TTTOOLINFO ti;
    } TTHITTESTINFO;

    typedef enum Tooltip_NotifycationCodes {
        TTN_GETDISPINFO = (-520 - 0),
        TTN_GETDISPINFOW = (-520 - 10),
        TTN_SHOW = (-520 - 1),
        TTN_POP = (-520 - 2),
        TTN_NEEDTEXT = (-520 - 0),
        TTN_NEEDTEXTW = (-520 - 10),
    } Tooltip_NotifycationCodes;

    typedef struct NMTTDISPIFNO {
        NMHDR hdr;
        LONG lpszText;
		BYTE szText[80];
        LONG hinst;
        LONG uFlags;
        LONG lParam;
    } NMTTDISPINFO;

    typedef struct NMTTDISPINFOW {
        NMHDR hdr;
        LONG lpszText;
        BYTE szText[160];
        LONG hinst;
        LONG uFlags;
        LONG lParam;
    } NMTTDISPINFOW;

// ******* DRAG LIST CONTROL


    typedef struct DRAGLISTINFO {
        LONG uNotification;
        LONG hWnd;
        POINTAPI ptCursor;
    } DRAGLISTINFO;

// *********** UPDOWN Control

    typedef struct UDACCEL {
        LONG nSec;
        LONG nInc;
    } UDACCEL;

    typedef struct NMUPDOWN {
        NMHDR hdr;
        LONG iPos;
        LONG iDelta;
    } NMUPDOWN;

    typedef enum UpDownNotifications {
        UDN_DELTAPOS = (-721 - 1)
    } UpDownNotifications;


// ********* PROGRESS CONTROL

    typedef struct PBRANGE {
        LONG iLow;
        LONG iHigh;
    } PBRANGE;

// ***************** LISTVIEW CONTROL

typedef struct LVITEM {
    LONG mask;
    LONG iItem;
    LONG iSubItem;
    LONG state;
    LONG stateMask;
    LONG pszText;
    LONG cchTextMax;
    LONG iImage;
    LONG lParam;
    LONG iIndent;
} LVITEM;

typedef struct LVITEMv6 {
    LONG mask;
    LONG iItem;
    LONG iSubItem;
    LONG state;
    LONG stateMask;
    LONG pszText;
    LONG cchTextMax;
    LONG iImage;
    LONG lParam;
    LONG iIndent;
    LONG iGroupId;
    UINT cColumns; // tile view columns
    UINT puColumns;
} LVITEMv6;

typedef struct LVGROUP {
    UINT cbSize;
    UINT mask;
    long pszHeader; //LPWSTR
    long cchHeader;
    long pszFooter; //LPWSTR
    long cchFooter;
    long iGroupId;
    UINT stateMask;
    UINT state;
    UINT uAlign;
} LVGROUP;

typedef struct LVGROUPMETRICS {
    UINT cbSize;
    UINT mask;
    UINT Left;
    UINT Top;
    UINT Right;
    UINT Bottom;
    COLORREF crLeft;
    COLORREF crTop;
    COLORREF crRight;
    COLORREF crBottom;
    COLORREF crHeader;
    COLORREF crFooter;
} LVGROUPMETRICS;

typedef struct LVFINDINFO
{
    ListViewFindIndoFlags flags;
    LONG psz;
    LONG lParam;
    POINTAPI pt;
    LONG vkDirection;
} LVFINDINFO;

typedef struct LVHITTESTINFO
{
    POINTAPI pt;
    LONG flags;
    LONG iItem;
    LONG iSubItem; // valid only for LVM_SUBITEMHITTEST
} LVHITTESTINFO;

typedef struct LVCOLUMN
{
    ListViewColumnMasks mask;
    ListViewColumnFormats fmt;
    LONG cx;
    LONG pszText;
    LONG cchTextMax;
    LONG iSubItem;
    LONG iImage;
    LONG iOrder;
} LVCOLUMN;

typedef struct LVBKIMAGE {
    ListViewBkImageFlags ulFlags; // LVBKIF_*
    LONG hbm;
    LONG pszImage;
    LONG cchImageMax;
    LONG xOffsetPercent;
    LONG yOffsetPercent;
} LVBKIMAGE;


typedef struct NMLISTVIEW {
    NMHDR   hdr;
    LONG iItem;
    LONG iSubItem;
    LONG uNewState;
    LONG uOldState;
    LONG uChanged;
    POINTAPI ptAction;
    LONG lParam;
} NMLISTVIEW;

typedef struct NMITEMACTIVATE {
    NMHDR hdr;
    LONG iItem;
    LONG iSubItem;
    LONG uNewState;
    LONG uOldState;
    LONG uChanged;
    POINTAPI ptAction;
    LONG lParam;
    ListViewKeyFlags uKeyFlags;
} NMITEMACTIVATE;

typedef struct NMLVCUSTOMDRAW {
    NMCUSTOMDRAW nmcd;
    LONG clrText;
    LONG clrTextBk;
    LONG iSubItem;
} NMLVCUSTOMDRAW;

typedef struct NMLVCACHEHINT {
    NMHDR   hdr;
    LONG iFrom;
    LONG iTo;
} NMLVCACHEHINT;

typedef struct NMLVFINDITEM {
    NMHDR   hdr;
    LONG iStart;
    LVFINDINFO lvfi;
} NMLVFINDITEM;

typedef struct NMLVODSTATECHANGE {
    NMHDR hdr;
    LONG iFrom;
    LONG iTo;
    LONG uNewState;
    LONG uOldState;
} NMLVODSTATECHANGE;

typedef struct NMLVDISPINFO {
    NMHDR hdr;
    LVITEM item;
} NMLVDISPINFO;

typedef struct NMLVDISPINFOv6 {
    NMHDR hdr;
    LVITEMv6 item;
} NMLVDISPINFOv6;

typedef struct NMLVKEYDOWN {
    NMHDR hdr;
    short wVKey;
    LONG flags;
} NMLVKEYDOWN;

typedef struct NMLVGETINFOTIP {
    NMHDR hdr;
    ListViewGetInfoTipFlags dwFlags;
    LONG pszText;
    LONG cchTextMax;
    LONG iItem;
    LONG iSubItem;
    LONG lParam;
} NMLVGETINFOTIP;

//********************* TREEVIEW CONTROL

typedef struct TVITEM {
    LONG mask;
    LONG hItem;
    LONG state;
    LONG stateMask;
    LONG pszText;
    LONG cchTextMax;
    LONG iImage;
    LONG iSelectedImage;
    LONG cChildren;
    LONG lParam;
} TVITEM;

typedef struct TVITEMEX {
    LONG mask;
    LONG hItem;
    LONG state;
    LONG stateMask;
    LONG pszText;
    LONG cchTextMax;
    LONG iImage;
    LONG iSelectedImage;
    LONG cChildren;
    LONG lParam;
    LONG iIntegral;
} TVITEMEX;

typedef struct TVINSERTSTRUCTEX {
    LONG hParent;
    LONG hInsertAfter;
    TVITEMEX itemex;
} TVINSERTSTRUCTEX;

typedef struct TVINSERTSTRUCT {
    LONG hParent;
    LONG hInsertAfter;
    TVITEM item;
} TVINSERTSTRUCT;

typedef struct TVHITTESTINFO {
	POINTAPI pt;
    LONG flags;
    LONG hItem;
} TVHITTESTINFO;

typedef struct TVSORTCB {
    LONG hParent;
    LONG lpfnCompare;
    LONG lParam;
} TVSORTCB;

typedef struct NM_TREEVIEW {
    NMHDR    hdr;
    LONG action;
    TVITEM itemOld;
    TVITEM itemNew;
    POINTAPI ptDrag;
} NMTREEVIEW;

typedef struct NMTVDISPINFO {
    NMHDR hdr;
    TVITEM item;
} NMTVDISPINFO;

typedef struct TV_KEYDOWN {
    NMHDR hdr;
    short wVKey;
    LONG flags;
} TVKEYDOWN;

typedef struct NMTVCUSTOMDRAW {
    NMCUSTOMDRAW nmcd;
    LONG clrText;
    LONG clrTextBk;
    LONG iLevel;
} NMTVCUSTOMDRAW;

typedef struct TV_GETINFOTIP {
    NMHDR hdr;
    LONG pszText;
    LONG cchTextMax;
    LONG hItem;
    LONG lParam;
} TVGETINFOTIP;

// ******************* COMBOBOXEX CONTROL

/*
 * Combobox information
 */
typedef struct tagCOMBOBOXINFO
{
    DWORD cbSize;
    RECT  rcItem;
    RECT  rcButton;
    DWORD stateButton;
    HWND  hwndCombo;
    HWND  hwndItem;
    HWND  hwndList;
} COMBOBOXINFO;

    typedef struct COMBOBOXEXITEM {
        ComboBoxExItemFlags mask;
        LONG iItem;
        LONG pszText;
        LONG cchTextMax;
        LONG iImage;
        LONG iSelectedImage;
        LONG iOverlay;
        LONG iIndent;
        LONG lParam;
    } COMBOBOXEXITEM;

    typedef struct NMCOMBOBOXEX {
        NMHDR hdr;
        COMBOBOXEXITEM ceItem;
    } NMCOMBOBOXEX;

 // CBEN_DRAGBEGIN sends this information ...

    typedef struct NMCBEDRAGBEGINW {
        NMHDR hdr;
        LONG iItemid;
        BYTE szText[520];
    } NMCBEDRAGBEGINW;

    typedef struct NMCBEDRAGBEGIN {
        NMHDR hdr;
        LONG iItemid;
        BYTE szText[260];
    } NMCBEDRAGBEGIN;

    typedef struct NMCBEENDEDITW {
        NMHDR hdr;
        BOOL fChanged;
        LONG iNewSelection;
        BYTE szText[520];
        LONG iWhy;
    } NMCBEENDEDITW;

    typedef struct NMCBEENDEDIT {
        NMHDR hdr;
        BOOL fChanged;
        LONG iNewSelection;
        BYTE szText[260];
        LONG iWhy;
    } NMCBEENDEDIT;

    typedef struct TCITEMHEADER {
        TabItemFlags mask;
        LONG lpReserved1;
        LONG lpReserved2;
        LONG pszText;
        LONG cchTextMax;
        LONG iImage;
    } TCITEMHEADER;

    typedef struct TCITEM {
        TabItemFlags mask;
        TabItemState dwState;
        TabItemState dwStateMask;
        LONG pszText;
        LONG cchTextMax;
        LONG iImage;
        LONG lParam;
    } TCITEM;

    typedef enum TabHitTestFlags {
        TCHT_NOWHERE = 0x0001,
        TCHT_ONITEMICON = 0x0002,
        TCHT_ONITEMLABEL = 0x0004,
        TCHT_ONITEM = 0x0006,
    } TabHitTestFlags;

    typedef struct TCHITTESTINFO {
        POINTAPI pt;
        TabHitTestFlags flags;
    } TCHITTESTINFO;

    typedef struct NMTCKEYDOWN {
        NMHDR hdr;
        short wVKey;
        LONG flags;
    } NMTCKEYDOWN;

    typedef struct MCHITTESTINFO {
        LONG cbSize;
        POINTAPI pt;
        LONG uHit; // out param
        SYSTEMTIME st;
    } MCHITTESTINFO;

    typedef struct NMSELCHANGE {
        NMHDR       nmhdr;
        SYSTEMTIME      stSelStart;
        SYSTEMTIME      stSelEnd;
    } NMSELCHANGE;

    typedef struct NMDAYSTATE {
        NMHDR      nmhdr;
        SYSTEMTIME stStart;
        LONG cDayState;
        LONG prgDayState; // points to cDayState MONTHDAYSTATEs
    } NMDAYSTATE;

    typedef struct NMDATETIMECHANGE {
        NMHDR       nmhdr;
        DateTimePickDateTimeChangeFlags dwFlags; // GDT_VALID or GDT_NONE
        SYSTEMTIME  st; // valid iff dwFlags = = GDT_VALID
    } NMDATETIMECHANGE;

    typedef struct NMDATETIMESTRING {
        NMHDR      nmhdr;
        LONG pszUserString; // string user entered
        SYSTEMTIME st; // app fills this in
        DateTimePickDateTimeChangeFlags dwFlags; // GDT_VALID or GDT_NONE
    } NMDATETIMESTRING;

    typedef struct NMDATETIMEWMKEYDOWN {
        NMHDR      nmhdr;
        LONG nVirtKey; // virtual key code of WM_KEYDOWN which MODIFIES an X field
        LONG pszFormat; // format substring
        SYSTEMTIME st; // current systemtime, app should modify based on key
    } NMDATETIMEWMKEYDOWN;

    typedef struct NMDATETIMEFORMAT {
        NMHDR      nmhdr;
        LONG pszFormat; // format substring
        SYSTEMTIME st; // current systemtime
        LONG pszDisplay; // string to display
        BYTE       szDisplay[64]; // buffer pszDisplay originally points at
    } NMDATETIMEFORMAT;

    typedef struct NMDATETIMEFORMATW {
        NMHDR nmhdr;
        LONG pszFormat; // format substring
        SYSTEMTIME st; // current systemtime
        LONG pszDisplay; // string to display
        BYTE szDisplay[128]; // buffer pszDisplay originally points at
    } NMDATETIMEFORMATW;

    typedef struct NMDATETIMEFORMATQUERY {
        NMHDR nmhdr;
        LONG pszFormat; // format substring
        SIZE szMax; // max bounding rectangle app will use for this format string
    } NMDATETIMEFORMATQUERY;

    typedef struct NMIPADDRESS {
        NMHDR hdr;
        LONG iField;
        LONG iValue;
    } NMIPADDRESS;
//---------------------------------------------------------------------------------------
//Pager Control Notification Messages
//---------------------------------------------------------------------------------------

	// This structure is sent along with PGN_SCROLL notifications
    typedef struct NMPGSCROLL {
        NMHDR hdr;
		WORD fwKeys;
        RECT	rcParent;	// Contains Parent Window Rect
		LONG	iDir;		// Scrolling Direction
        LONG	iXpos;		// Horizontal scroll position
        LONG	iYpos;		// Vertical scroll position
        LONG	iScroll;	// [in/out] Amount to scroll
    } NMPGSCROLL;

    typedef struct NMPGSCROLLVB {
        NMHDR hdr;
		BYTE  TheRest[34];	//Bytes are manullay copied and then accessed using copymemory
    } NMPGSCROLLVB;


    typedef struct NMPGCALCSIZE {
        NMHDR   hdr;
        LONG dwFlag;
        LONG iWidth;
        LONG iHeight;
    } NMPGCALCSIZE;

// ====== TrackMouseEvent =======

typedef struct TRACKMOUSEEVENT {
	long cbSize;
	long dwFlags;
	long hwndTrack;
	long dwHoverTime;
} TRACKMOUSEEVENT;

// FLAT ScrollBars

    typedef struct SCROLLINFO {
        LONG cbSize;
        LONG fMask;
        LONG nMin;
        LONG nMax;
        LONG nPage;
        LONG nPos;
        LONG nTrackPos;
    } SCROLLINFO;

//************ RICH EDIT

// enum for use with undocumented EM_GET/SETTEXTMODE which don't work
typedef enum RE_TEXTMODE {
	TM_PLAINTEXT			= 1,
	TM_RICHTEXT				= 2,	// default behavior 
	TM_SINGLELEVELUNDO		= 4,
	TM_MULTILEVELUNDO		= 8,	// default behavior 
	TM_SINGLECODEPAGE		= 16,
	TM_MULTICODEPAGE		= 32	// default behavior 
} RE_TEXTMODE;

// all character format measurements are in twips
typedef struct A_CHARFORMAT {
	UINT		cbSize;
	DWORD		dwMask;
	DWORD		dwEffects;
	LONG		yHeight;
	LONG		yOffset;
	COLORREF	crTextColor;
	BYTE		bCharSet;
	BYTE		bPitchAndFamily;
    BYTE        szFaceName[32];			// LF_FACESIZE
	short		wPad; //This is to align the var boundries so the len(CHARFORMAT) = 60, not 58
} A_CHARFORMAT;

typedef struct W_CHARFORMAT {
	UINT		cbSize;
	DWORD		dwMask;
	DWORD		dwEffects;
	LONG		yHeight;
	LONG		yOffset;
	COLORREF	crTextColor;
	BYTE		bCharSet;
	BYTE		bPitchAndFamily;
    BYTE        szFaceName[64];		// LF_FACESIZE
	short		wPad;
} W_CHARFORMAT;

typedef struct CHARRANGE {
	LONG	cpMin;
	LONG	cpMax;
} CHARRANGE;

typedef struct TEXTRANGE {
	CHARRANGE	chrg;
    long		lpstrText;    // allocated by caller, zero terminated by RichEdit
} TEXTRANGE;

typedef struct EDITSTREAM {
    DWORD dwCookie;     // user value passed to callback as first parameter
    DWORD dwError;      // last error
	long  pfnCallback;
} EDITSTREAM;

typedef struct FINDTEXT {
	CHARRANGE	chrg;
    long		lpstrText;
} FINDTEXT;

typedef struct FINDTEXTEX {
	CHARRANGE chrg;
    long lpstrText;
	CHARRANGE chrgText;
} FINDTEXTEX;

typedef struct FORMATRANGE {
	HDC hdc;
	HDC hdcTarget;
	RECT rc;
	RECT rcPage;
	CHARRANGE chrg;
} FORMATRANGE;

typedef struct PARAFORMAT {
	UINT	cbSize;
	DWORD	dwMask;
	WORD	wNumbering;
    WORD    wEffects;
	LONG	dxStartIndent;
	LONG	dxRightIndent;
	LONG	dxOffset;
	WORD	wAlignment;
	SHORT	cTabCount;
	LONG	rgxTabs[32]; // MAX_TAB_STOPS
} PARAFORMAT;

// CHARFORMAT2 and PARAFORMAT2 structures

typedef struct A_CHARFORMAT2 {
	UINT		cbSize;
	DWORD		dwMask;
	DWORD		dwEffects;
	LONG		yHeight;
    LONG        yOffset;            // > 0 for superscript, < 0 for subscript
	COLORREF	crTextColor;
	BYTE		bCharSet;
	BYTE		bPitchAndFamily;
    BYTE        szFaceName[32]; // LF_FACESIZE
	short		wPad;
    WORD        wWeight;            // Font weight (LOGFONT value)
    SHORT       sSpacing;           // Amount to space between letters
    COLORREF    crBackColor;        // Background color
    long        lcid;               // Locale ID
    DWORD       dwReserved;         // Reserved. Must be 0
    SHORT       sStyle;             // Style handle
    WORD        wKerning;           // Twip size above which to kern char pair
    BYTE        bUnderlineType;     // Underline type
    BYTE        bAnimation;         // Animated text like marching ants
    BYTE        bRevAuthor;         // Revision author index
	BYTE		bReserved1;
} A_CHARFORMAT2;

typedef struct W_CHARFORMAT2 {
	UINT		cbSize;
	DWORD		dwMask;
	DWORD		dwEffects;
	LONG		yHeight;
    LONG        yOffset;            // > 0 for superscript, < 0 for subscript
	COLORREF	crTextColor;
	BYTE		bCharSet;
	BYTE		bPitchAndFamily;
    BYTE        szFaceName[64]; // LF_FACESIZE
	short		wPad;
    WORD        wWeight;            // Font weight (LOGFONT value)
    SHORT       sSpacing;           // Amount to space between letters
    COLORREF    crBackColor;        // Background color
    long        lcid;               // Locale ID
    DWORD       dwReserved;         // Reserved. Must be 0
    SHORT       sStyle;             // Style handle
    WORD        wKerning;           // Twip size above which to kern char pair
    BYTE        bUnderlineType;     // Underline type
    BYTE        bAnimation;         // Animated text like marching ants
    BYTE        bRevAuthor;         // Revision author index
	BYTE		bReserved1;
} W_CHARFORMAT2;


typedef struct PARAFORMAT2 {
	UINT	cbSize;
	DWORD	dwMask;
	WORD	wNumbering;
	WORD	wReserved;
	LONG	dxStartIndent;
	LONG	dxRightIndent;
	LONG	dxOffset;
	WORD	wAlignment;
	SHORT	cTabCount;
	LONG	rgxTabs[32]; // MAX_TAB_STOPS
    LONG    dySpaceBefore;          // Vertical spacing before para
    LONG    dySpaceAfter;           // Vertical spacing after para
    LONG    dyLineSpacing;          // Line spacing depending on Rule
    SHORT   sStyle;                 // Style handle
    BYTE    bLineSpacingRule;       // Rule for line spacing (see tom.doc)
    BYTE    bCRC;                   // Reserved for CRC for rapid searching
    WORD    wShadingWeight;         // Shading in hundredths of a per cent
    WORD    wShadingStyle;          // Nibble 0: style, 1: cfpat, 2: cbpat
    WORD    wNumberingStart;        // Starting value for numbering
    WORD    wNumberingStyle;        // Alignment, roman/arabic, (), ), ., etc.
    WORD    wNumberingTab;          // Space bet 1st indent and 1st-line text
    WORD    wBorderSpace;           // Space between border and text (twips)
    WORD    wBorderWidth;           // Border pen width (twips)
    WORD    wBorders;               // Byte 0: bits specify which borders
                                    // Nibble 2: border style, 3: color index
} PARAFORMAT2;

typedef struct REQRESIZE {
	NMHDR nmhdr;
	RECT rc;
} REQRESIZE;

typedef struct SELCHANGE {
	NMHDR nmhdr;
	CHARRANGE chrg;
	WORD seltyp;
} SELCHANGE;

typedef struct ENDROPFILES {
	NMHDR nmhdr;
	HANDLE hDrop;
	LONG cp;
	BOOL fProtected;
} ENDROPFILES;

typedef struct ENPROTECTED {
	NMHDR   nmhdr;
	UINT    msg;
	long    wParam;
    LONG	lParam; // LPARAM  
	CHARRANGE chrg;
} ENPROTECTED;

typedef struct ENSAVECLIPBOARD {
	NMHDR nmhdr;
	LONG cObjectCount;
    LONG cch;
} ENSAVECLIPBOARD;

typedef struct ENOLEOPFAILED {
	NMHDR	nmhdr;
	LONG	iob;
	LONG	lOper;
	long	hr;
} ENOLEOPFAILED;

typedef struct OBJECTPOSITIONS {
    NMHDR nmhdr;
    LONG cObjectCount;
    LONG pcpPositions;
} OBJECTPOSITIONS;

typedef struct ENLINK {
    NMHDR nmhdr;
    UINT msg;
    long wParam;
    LONG	lParam; // LPARAM  
    CHARRANGE chrg;
} ENLINK;

// PenWin specific
typedef struct ENCORRECTTEXT {
	NMHDR nmhdr;
	CHARRANGE chrg;
	WORD seltyp;
} ENCORRECTTEXT;


// Far East specific
typedef struct PUNCTUATION {
	UINT	iSize;
	long	szPunctuation;
} PUNCTUATION;

// Far East specific
typedef struct COMPCOLOR {
	COLORREF crText;
	COLORREF crBackground;
	DWORD dwEffects;
}COMPCOLOR;


// Paste Special
typedef struct REPASTESPECIAL {
	DWORD	dwAspect;
	DWORD	dwParam;
} REPASTESPECIAL;

//  UndoName info
typedef enum UNDONAMEID {
    UID_UNKNOWN     = 0,
	UID_TYPING		= 1,
	UID_DELETE 		= 2,
	UID_DRAGDROP	= 3,
	UID_CUT			= 4,
	UID_PASTE		= 5
} UNDONAMEID;

// EM_GETTEXTEX info; this struct is passed in the wparam of the message 
typedef struct GETTEXTEX {
	DWORD	cb;				// count of bytes in the string				
	DWORD	flags;			// flags (see the GT_XXX defines			
	UINT	codepage;		// code page for translation (CP_ACP for default,
						    //    1200 for Unicode							
	long    lpDefaultChar;	// replacement for unmappable chars //LPSTR			
	BOOL    lpUsedDefChar;	// pointer to flag set when def char used	
} GETTEXTEX;

// EM_GETTEXTLENGTHEX info; this struct is passed in the wparam of the msg 
typedef struct GETTEXTLENGTHEX {
	DWORD	flags;			// flags (see GTL_XXX defines)				
	UINT	codepage;		// code page for translation (CP_ACP for default,
							//   1200 for Unicode							
} GETTEXTLENGTHEX;

typedef struct SETTEXTEX {
    DWORD flags;
    UINT codepage;
} SETTEXTEX;

typedef struct MSGFILTER {
    NMHDR nmhdr;
    UINT msg;
    long wParam;
    long lParam;
} MSGFILTER;

typedef struct PAINTSTRUCT { 
  HDC  hdc; 
  BOOL fErase; 
  RECT rcPaint; 
  BOOL fRestore; 
  BOOL fIncUpdate; 
  BYTE rgbReserved[32]; 
} PAINTSTRUCT;

//URL History
typedef struct STATURL {
    [helpstring("Struct size")]
    long cbSize;
    [helpstring("URL")]
    long pwcsUrl;
    [helpstring("Page title")]
    long pwcsTitle;
    [helpstring("Last visited date (UTC)")]
    FILETIME ftLastVisited;
    [helpstring("Last updated date (UTC)")]
    FILETIME ftLastUpdated;
    [helpstring("Expiration date (UTC)")]
    FILETIME ftExpires;
    [helpstring("Flags")]
    STATURLFLAGS dwFlags;
} STATURL;