
[
	helpstring("RichEdit consts"),
	dllname("NOSUCHR.DLL")
]
module RichEdit {

	const long cchTextLimitDefault = 32767;

	// RichEdit messages 

    const long EM_SCROLLCARET          = 1024 + 49;
    const long EM_CANPASTE             = 1024 + 50;
    const long EM_DISPLAYBAND          = 1024 + 51;
    const long EM_EXGETSEL             = 1024 + 52;
    const long EM_EXLIMITTEXT          = 1024 + 53;
    const long EM_EXLINEFROMCHAR       = 1024 + 54;
    const long EM_EXSETSEL             = 1024 + 55;
    const long EM_FINDTEXT             = 1024 + 56;
    const long EM_FORMATRANGE          = 1024 + 57;
    const long EM_GETCHARFORMAT        = 1024 + 58;
    const long EM_GETEVENTMASK         = 1024 + 59;
    const long EM_GETOLEINTERFACE      = 1024 + 60;
    const long EM_GETPARAFORMAT        = 1024 + 61;
    const long EM_GETSELTEXT           = 1024 + 62;
    const long EM_HIDESELECTION        = 1024 + 63;
    const long EM_PASTESPECIAL         = 1024 + 64;
    const long EM_REQUESTRESIZE        = 1024 + 65;
    const long EM_SELECTIONTYPE        = 1024 + 66;
    const long EM_SETBKGNDCOLOR        = 1024 + 67;
    const long EM_SETCHARFORMAT        = 1024 + 68;
    const long EM_SETEVENTMASK         = 1024 + 69;
    const long EM_SETOLECALLBACK       = 1024 + 70;
    const long EM_SETPARAFORMAT        = 1024 + 71;
    const long EM_SETTARGETDEVICE      = 1024 + 72;
    const long EM_STREAMIN             = 1024 + 73;
    const long EM_STREAMOUT            = 1024 + 74;
    const long EM_GETTEXTRANGE         = 1024 + 75;
    const long EM_FINDWORDBREAK        = 1024 + 76;
    const long EM_SETOPTIONS           = 1024 + 77;
    const long EM_GETOPTIONS           = 1024 + 78;
    const long EM_FINDTEXTEX           = 1024 + 79;
    const long EM_GETWORDBREAKPROCEX   = 1024 + 80;
    const long EM_SETWORDBREAKPROCEX   = 1024 + 81;

    // Richedit v2.0 messages
    const long EM_SETUNDOLIMIT         = 1024 + 82;
    const long EM_REDO                 = 1024 + 84;
    const long EM_CANREDO              = 1024 + 85;
    const long EM_GETUNDONAME          = 1024 + 86;
    const long EM_GETREDONAME          = 1024 + 87;
    const long EM_STOPGROUPTYPING      = 1024 + 88;

    const long EM_SETTEXTMODE         = 1024 + 89;
    const long EM_GETTEXTMODE         = 1024 + 90;
	const long EM_AUTOURLDETECT		  = 1024 + 91;
	const long EM_GETAUTOURLDETECT	  = 1024 + 92;
	const long EM_SETPALETTE		  = 1024 + 93;
	const long EM_GETTEXTEX			  = 1024 + 94;
	const long EM_GETTEXTLENGTHEX	  = 1024 + 95;
	const long EM_SETTEXTEX			  = 1024 + 97;


    // Far East specific messages
    const long EM_SETPUNCTUATION       = 1024 + 100;
    const long EM_GETPUNCTUATION       = 1024 + 101;
    const long EM_SETWORDWRAPMODE      = 1024 + 102;
    const long EM_GETWORDWRAPMODE      = 1024 + 103;
    const long EM_SETIMECOLOR          = 1024 + 104;
    const long EM_GETIMECOLOR          = 1024 + 105;
    const long EM_SETIMEOPTIONS        = 1024 + 106;
    const long EM_GETIMEOPTIONS        = 1024 + 107;
    const long EM_CONVPOSITION         = 1024 + 108;

    const long EM_SETLANGOPTIONS       = 1024 + 120;
    const long EM_GETLANGOPTIONS       = 1024 + 121;
    const long EM_GETIMECOMPMODE       = 1024 + 122;

    // Options for EM_SETLANGOPTIONS and EM_GETLANGOPTIONS
    const long IMF_AUTOKEYBOARD        = 0x0001;
    const long IMF_AUTOFONT            = 0x0002;
	const long IMF_IMECANCELCOMPLETE   = 0x0004;	// high completes the comp string when aborting, low cancels.
	const long IMF_IMEALWAYSSENDNOTIFY = 0x0008;

    // Values for EM_GETIMECOMPMODE
    const long ICM_NOTOPEN             = 0x0000;
    const long ICM_LEVEL3              = 0x0001;
    const long ICM_LEVEL2              = 0x0002;
    const long ICM_LEVEL2_5            = 0x0003;
    const long ICM_LEVEL2_SUI          = 0x0004;

	// New notifications 

    const long EN_MSGFILTER            = 0x0700;
    const long EN_REQUESTRESIZE        = 0x0701;
    const long EN_SELCHANGE            = 0x0702;
    const long EN_DROPFILES            = 0x0703;
    const long EN_PROTECTED            = 0x0704;
    const long EN_CORRECTTEXT          = 0x0705;          // PenWin specific
    const long EN_STOPNOUNDO           = 0x0706;
    const long EN_IMECHANGE            = 0x0707;          // Far East specific
    const long EN_SAVECLIPBOARD        = 0x0708;
    const long EN_OLEOPFAILED          = 0x0709;
    const long EN_OBJECTPOSITIONS      = 0x070a;
    const long EN_LINK                 = 0x070b;
	const long EN_DRAGDROPDONE		   = 0x070c;

    // Event notification masks

    const long ENM_NONE                = 0x00000000;
    const long ENM_CHANGE              = 0x00000001;
    const long ENM_UPDATE              = 0x00000002;
    const long ENM_SCROLL              = 0x00000004;
    const long ENM_KEYEVENTS           = 0x00010000;
    const long ENM_MOUSEEVENTS         = 0x00020000;
    const long ENM_REQUESTRESIZE       = 0x00040000;
    const long ENM_SELCHANGE           = 0x00080000;
    const long ENM_DROPFILES           = 0x00100000;
    const long ENM_PROTECTED           = 0x00200000;
    const long ENM_CORRECTTEXT         = 0x00400000;      // PenWin specific

    // Far East specific notification mask
    const long ENM_IMECHANGE           = 0x00800000;
    const long ENM_LANGCHANGE          = 0x01000000;
    const long ENM_OBJECTPOSITIONS     = 0x02000000;
    const long ENM_LINK                = 0x04000000;

    // New edit control styles
    const long ES_SAVESEL              = 0x00008000;
    const long ES_SUNKEN               = 0x00004000;
    const long ES_DISABLENOSCROLL      = 0x00002000;
    // same as WS_MAXIMIZE, but that doesn't make sense so we re-use the value
    const long ES_SELECTIONBAR         = 0x01000000;
	// same as ES_UPPERCASE, but re-used to completely disable OLE drag'n'drop 
	const long ES_NOOLEDRAGDROP		= 0x00000008;

    // New edit control extended style
    const long ES_EX_NOCALLOLEINIT     = 0x01000000;

    // These flags are used in FE Windows
    const long ES_VERTICAL             = 0x00400000;
    const long ES_NOIME                = 0x00080000;
    const long ES_SELFIME              = 0x00040000;

    // Edit control options
    const long ECO_AUTOWORDSELECTION   = 0x00000001;
    const long ECO_AUTOVSCROLL         = 0x00000040;
    const long ECO_AUTOHSCROLL         = 0x00000080;
    const long ECO_NOHIDESEL           = 0x00000100;
    const long ECO_READONLY            = 0x00000800;
    const long ECO_WANTRETURN          = 0x00001000;
    const long ECO_SAVESEL             = 0x00008000;
    const long ECO_SELECTIONBAR        = 0x01000000;
    const long ECO_VERTICAL            = 0x00400000;      // FE specific


    // ECO operations
    const long ECOOP_SET               = 0x0001;
    const long ECOOP_OR                = 0x0002;
    const long ECOOP_AND               = 0x0003;
    const long ECOOP_XOR               = 0x0004;

    // new word break function actions
    const long WB_CLASSIFY         = 3;
    const long WB_MOVEWORDLEFT     = 4;
    const long WB_MOVEWORDRIGHT    = 5;
    const long WB_LEFTBREAK        = 6;
    const long WB_RIGHTBREAK       = 7;

    // Far East specific flags
    const long WB_MOVEWORDPREV     = 4;
    const long WB_MOVEWORDNEXT     = 5;
    const long WB_PREVBREAK        = 6;
    const long WB_NEXTBREAK        = 7;

    const long PC_FOLLOWING        = 1;
    const long PC_LEADING          = 2;
    const long PC_OVERFLOW         = 3;
    const long PC_DELIMITER        = 4;
    const long WBF_WORDWRAP        = 0x010;
    const long WBF_WORDBREAK       = 0x020;
    const long WBF_OVERFLOW        = 0x040;
    const long WBF_LEVEL1          = 0x080;
    const long WBF_LEVEL2          = 0x100;
    const long WBF_CUSTOM          = 0x200;

    // Far East specific flags
    const long IMF_FORCENONE           = 0x0001;
    const long IMF_FORCEENABLE         = 0x0002;
    const long IMF_FORCEDISABLE        = 0x0004;
    const long IMF_CLOSESTATUSWINDOW   = 0x0008;
    const long IMF_VERTICAL            = 0x0020;
    const long IMF_FORCEACTIVE         = 0x0040;
    const long IMF_FORCEINACTIVE       = 0x0080;
    const long IMF_FORCEREMEMBER       = 0x0100;
    const long IMF_MULTIPLEEDIT        = 0x0400;

    // Word break flags (used with WB_CLASSIFY)
    const long WBF_CLASS           = 0x0F;
    const long WBF_ISWHITE         = 0x10;
    const long WBF_BREAKLINE       = 0x20;
    const long WBF_BREAKAFTER      = 0x40;

    // CHARFORMAT masks
    const long CFM_BOLD        = 0x00000001;
    const long CFM_ITALIC      = 0x00000002;
    const long CFM_UNDERLINE   = 0x00000004;
    const long CFM_STRIKEOUT   = 0x00000008;
    const long CFM_PROTECTED   = 0x00000010;
    const long CFM_LINK        = 0x00000020;      // Exchange hyperlink extension
    const long CFM_SIZE        = 0x80000000;
    const long CFM_COLOR       = 0x40000000;
    const long CFM_FACE        = 0x20000000;
    const long CFM_OFFSET      = 0x10000000;
    const long CFM_CHARSET     = 0x08000000;

    // CHARFORMAT effects
    const long CFE_BOLD        = 0x0001;
    const long CFE_ITALIC      = 0x0002;
    const long CFE_UNDERLINE   = 0x0004;
    const long CFE_STRIKEOUT   = 0x0008;
    const long CFE_PROTECTED   = 0x0010;
    const long CFE_LINK        = 0x0020;
    const long CFE_AUTOCOLOR   = 0x40000000;      // NOTE: this corresponds to
                                               // CFM_COLOR, which controls it
    const long yHeightCharPtsMost = 1638;

    // EM_SETCHARFORMAT wParam masks
    const long SCF_SELECTION   = 0x0001;
    const long SCF_WORD        = 0x0002;
    const long SCF_DEFAULT     = 0x0000;      // set the default charformat or paraformat
    const long SCF_ALL         = 0x0004;      // not valid with SCF_SELECTION or SCF_WORD
	const long SCF_USEUIRULES  = 0x0008;		// modifier for SCF_SELECTION; says that
											// the format came from a toolbar, etc. and
											// therefore UI formatting rules should be
											// used instead of strictly formatting the
											// selection.


    // stream formats

    const long SF_TEXT         = 0x0001;
    const long SF_RTF          = 0x0002;
    const long SF_RTFNOOBJS    = 0x0003;      // outbound only
    const long SF_TEXTIZED     = 0x0004;      // outbound only
    const long SF_UNICODE      = 0x0010;      // Unicode file of some kind

    // Flag telling stream operations to operate on the selection only
    // EM_STREAMIN will replace the current selection
    // EM_STREAMOUT will stream out the current selection
    const long SFF_SELECTION   = 0x8000;

    // Flag telling stream operations to operate on the common RTF keyword only
    // EM_STREAMIN will accept the only common RTF keyword
    // EM_STREAMOUT will stream out the only common RTF keyword
    const long SFF_PLAINRTF    = 0x4000;

    // all paragraph measurements are in twips

    const long MAX_TAB_STOPS = 32;
    const long lDefaultTab = 720;

    // PARAFORMAT mask values
    const long PFM_STARTINDENT         = 0x00000001;
    const long PFM_RIGHTINDENT         = 0x00000002;
    const long PFM_OFFSET              = 0x00000004;
    const long PFM_ALIGNMENT           = 0x00000008;
    const long PFM_TABSTOPS            = 0x00000010;
    const long PFM_NUMBERING           = 0x00000020;
    const long PFM_OFFSETINDENT        = 0x80000000;

    // PARAFORMAT numbering options
    const long PFN_BULLET      = 0x0001;

    // PARAFORMAT alignment options
    const long PFA_LEFT    = 0x0001;
    const long PFA_RIGHT   = 0x0002;
    const long PFA_CENTER  = 0x0003;

    // New masks and effects -- a parenthesized asterisk indicates that
    // the data is stored by RichEdit2.0, but not displayed

    const long CFM_SMALLCAPS       = 0x0040;
    const long CFM_ALLCAPS         = 0x0080;
    const long CFM_HIDDEN          = 0x0100;
    const long CFM_OUTLINE         = 0x0200;
    const long CFM_SHADOW          = 0x0400;
    const long CFM_EMBOSS          = 0x0800;
    const long CFM_IMPRlong        = 0x1000;
    const long CFM_DISABLED        = 0x2000;
    const long CFM_REVISED         = 0x4000;

    const long CFM_BACKCOLOR       = 0x04000000;
    const long CFM_LCID            = 0x02000000;
    const long CFM_UNDERLINETYPE   = 0x00800000;
    const long CFM_WEIGHT          = 0x00400000;
    const long CFM_SPACING         = 0x00200000;
    const long CFM_KERNING         = 0x00100000;
    const long CFM_STYLE           = 0x00080000;
    const long CFM_ANIMATION       = 0x00040000;
    const long CFM_REVAUTHOR       = 0x00008000;

    const long CFE_SUBSCRIPT       = 0x00010000;      // Superscript and subscript are
    const long CFE_SUPERSCRIPT     = 0x00020000;      //  mutually exclusive

    const long CFM_SUBSCRIPT       = 0x00030000; // CFE_SUBSCRIPT | CFE_SUPERSCRIPT
    const long CFM_SUPERSCRIPT     = 0x00030000; // CFM_SUBSCRIPT

    // Underline types
    const long CFU_CF1UNDERLINE    = 0xFF;   // map charformat's bit underline to CF2.
    const long CFU_INVERT          = 0xFE;   // For IME composition fake a selection.
    const long CFU_UNDERLINEDOTTED = 0x3;    // (*) displayed as ordinary underline
    const long CFU_UNDERLINEDOUBLE = 0x2;    // (*) displayed as ordinary underline
    const long CFU_UNDERLINEWORD   = 0x1;    // (*) displayed as ordinary underline
    const long CFU_UNDERLINE       = 0x0;


	// this is a hack to make PARAFORMAT code more readable.
	// it applies to the wReserved field in PARAFORMAT, which
	// in PARAFORMAT2 is now used. 

	// const long wEffects			= wReserved;

    // PARAFORMAT 2.0 masks and effects

    const long PFM_SPACEBEFORE         = 0x00000040;
    const long PFM_SPACEAFTER          = 0x00000080;
    const long PFM_LINESPACING         = 0x00000100;
    const long PFM_STYLE               = 0x00000400;
    const long PFM_BORDER              = 0x00000800;
    const long PFM_SHADING             = 0x00001000;
    const long PFM_NUMBERINGSTYLE      = 0x00002000;
    const long PFM_NUMBERINGTAB        = 0x00004000;
    const long PFM_NUMBERINGSTART      = 0x00008000;

    const long PFM_RTLPARA             = 0x00010000;
    const long PFM_KEEP                = 0x00020000;
    const long PFM_KEEPNEXT            = 0x00040000;
    const long PFM_PAGEBREAKBEFORE     = 0x00080000;
    const long PFM_NOLINENUMBER        = 0x00100000;
    const long PFM_NOWIDOWCONTROL      = 0x00200000;
    const long PFM_DONOTHYPHEN         = 0x00400000;
    const long PFM_SIDEBYSIDE          = 0x00800000;

    const long PFM_TABLE               = 0xc0000000;

    const long PFE_TABLEROW            = 0xc000;      // These 3 options are mutually
    const long PFE_TABLECELLEND        = 0x8000;      //  exclusive and each imply
    const long PFE_TABLECELL           = 0x4000;      //  that para is part of a table


    /*
     *  PARAFORMAT numbering options (values for wNumbering):
     *
     *      Numbering Type      Value   Meaning
     *      tomNoNumbering        0     Turn off paragraph numbering
     *      tomNumberAsLCLetter   1     a, b, c, ...
     *      tomNumberAsUCLetter   2     A, B, C, ...
     *      tomNumberAsLCRoman    3     i, ii, iii, ...
     *      tomNumberAsUCRoman    4     I, II, III, ...
     *      tomNumberAsSymbols    5     default is bullet
     *      tomNumberAsNumber     6     0, 1, 2, ...
     *      tomNumberAsSequence   7     tomNumberingStart is first Unicode to use
     *
     *  Other valid Unicode chars are Unicodes for bullets.
     */


    const long PFA_JUSTIFY         = 4;   // New paragraph-alignment option 2.0 (*)


    // notification structures


    // UNICODE embedding character
    const long WCH_EMBEDDING   = 0xFFFC;

    const long SEL_EMPTY       = 0x0000;
    const long SEL_TEXT        = 0x0001;
    const long SEL_OBJECT      = 0x0002;
    const long SEL_MULTICHAR   = 0x0004;
    const long SEL_MULTIOBJECT = 0x0008;


	// used with IRichEditOleCallback::GetContextMenu, this flag will be
	// passed as a "selection type".  It indicates that a context menu for
	// a right-mouse drag drop should be generated.  The IOleObject parameter
	// will really be the IDataObject for the drop
 
	const long GCM_RIGHTMOUSEDROP  = 0x8000;

    const long OLEOP_DOVERB    = 1;

	// flags for the GETEXTEX data structure 
	const long GT_DEFAULT		= 0;
	const long GT_USECRLF		= 1;

	// flags for the GETTEXTLENGTHEX data structure							
	const long GTL_DEFAULT		= 0;	// do the default (return # of chars)		
	const long GTL_USECRLF		= 1;	// compute answer using CRLFs for paragraphs
	const long GTL_PRECISE		= 2;	// compute a precise answer					
	const long GTL_CLOSE			= 4;	// fast computation of a "close" answer		
	const long GTL_NUMCHARS		= 8;	// return the number of characters		
	const long GTL_NUMBYTES		= 16;	// return the number of _bytes_	


};