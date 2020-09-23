// ******* GDI functions
[
	//helpstring("Windows GDI Functions")
	dllname("GDI32.DLL")
]
module GDI {

    [
    entry("AnimatePalette")
    ]
    BOOL AnimatePalette([in] HPALETTE hpal, 
                               [in] UINT iStartIndex, 
                               [in] UINT cEntries, 
                               [in] LPVOID ppe);

    [entry("Arc")]
    BOOL Arc([in] HDC hDC,
                    [in] long nLeftRect, [in] long nTopRect,
                    [in] long nRightRect, [in] long nBottomRect,
                    [in] long nXRadial1, [in] long nYRadial1,
                    [in] long nXRadial2, [in] long nYRadial2);
    
    [entry("BitBlt")]
    BOOL BitBlt([in] HDC hdcDest,
                    [in] long nXDest, [in] long nYDest,
                    [in] long nWidth, [in] long nHeight,
                    [in] HDC hdcSrc,
                    [in] long nXSrc, [in] long nYSrc,
                    [in] DWORD dwRop);

    [entry("CancelDC")]
    BOOL CancelDC(
			[in] HDC hdc);

    [entry("Chord")]
    BOOL Chord([in] HDC hDC,
                      [in] long nLeftRect, [in] long nTopRect,
                      [in] long nRightRect, [in] long nBottomRect,
                      [in] long nXRadial1, [in] long nYRadial1,
                      [in] long nXRadial2, [in] long nYRadial2);

    [entry("CloseMetaFile")]
    HMETAFILE CloseMetaFile([in] HDC hdc);

    [entry("CombineRgn")]
    long CombineRgn([in] HRGN hrgnDst, [in] HRGN hrgnSrc1, 
                          [in] HRGN hrgnSrc2,[in] long fnCombineMode);

    [entry("CopyMetaFileA")]
    HMETAFILE A_CopyMetaFile(
				[in] HMETAFILE hmfSrc,
                [in] LPSTR lpszFile);

    [entry("CopyMetaFileW")]
    HMETAFILE W_CopyMetaFile(
				[in] HMETAFILE hmfSrc,
                [in] LPWSTR lpszFile);

    [entry("CreateBitmap")]
    HBITMAP CreateBitmap([in] long nWidth, [in] long nHeight,
                                [in] UINT cbPlanes,[in] UINT cbBits,
                                [in] void *lpvBits);

    [entry("CreateBitmapIndirect")]
    HBITMAP CreateBitmapIndirect([in] LPVOID lpbm);

    [entry("CreateBrushIndirect")]
    HBRUSH  CreateBrushIndirect([in] LPVOID lplb);

    [entry("CreateCompatibleBitmap")]
    HBITMAP CreateCompatibleBitmap([in] HDC hdc,
                                          [in] long nWidth,
                                          [in] long nHeight);

    [entry("CreateDiscardableBitmap")    ]
    HBITMAP CreateDiscardableBitmap([in] HDC hdc,[in] long nWidth,[in] long nHeight);

    [entry("CreateCompatibleDC")    ]
    HDC CreateCompatibleDC([in] HDC hdc);

    [entry("CreateDCA")    ]
    HDC A_CreateDC([in] LPSTR lpszDriver,[in] LPSTR lpszDevice, 
                        [in] LPSTR lpszOutput,
						[in] void * lpInitData);

    [entry("CreateDCA")    ]
    HDC W_CreateDC([in] LPWSTR lpszDriver,[in] LPWSTR lpszDevice, 
                        [in] LPWSTR lpszOutput,
						[in] void * lpInitData);

    [entry("CreateDIBitmap")    ]
    HBITMAP CreateDIBitmap([in] HDC hdc, 
                                  [in] LPVOID lpbmih,
                                  [in] DWORD fdwinit, 
                                  [in] void *lpbInit, 
                                  [in] LPVOID lpbmi,
                                  [in] UINT fuUsage);
    
    [entry("CreateDIBPatternBrush")    ]
    HBRUSH  CreateDIBPatternBrush([in] HGLOBAL hglbDIBPacked, 
                                         [in] UINT fuColorSpec);
    [entry("CreateDIBPatternBrushPt")    ]
    HBRUSH  CreateDIBPatternBrushPt([in] void *lpPackedDIB, 
                                           [in] UINT iUsage);
    [entry("CreateEllipticRgn")    ]
    HRGN CreateEllipticRgn([in] long nLeftRect, 
                                     [in] long nTopRect, 
                                     [in] long nRightRect, 
                                     [in] long nBottomRect);
    [entry("CreateEllipticRgnIndirect")    ]
    HRGN CreateEllipticRgnIndirect([in] RECT *lprc);
    
    [entry("CreateFontIndirectA")    ]
    HFONT A_CreateFontIndirect([in] LPVOID lplf);

    [entry("CreateFontIndirectW")    ]
    HFONT W_CreateFontIndirect([in] LPVOID lplf);

    [    entry("CreateFontA")    ]
    HFONT A_CreateFont([in] long nHeight, 
                              [in] long nWidth, 
                              [in] long nEscapement, 
                              [in] long nOrientation, 
                              [in] long fnWeight, 
                              [in] DWORD fdwItalid,
                              [in] DWORD fdwUnderline, 
                              [in] DWORD fdwStrikeOut, 
                              [in] DWORD fdwCharSet, 
                              [in] DWORD fdwOutputPrecision, 
                              [in] DWORD fdwClipPrecision, 
                              [in] DWORD fdwQuality, 
                              [in] DWORD fdwPitchAndFamily, 
                              [in] LPSTR lpszFace);

    [    entry("CreateFontW")    ]
    HFONT W_CreateFont([in] long nHeight, 
                              [in] long nWidth, 
                              [in] long nEscapement, 
                              [in] long nOrientation, 
                              [in] long fnWeight, 
                              [in] DWORD fdwItalid,
                              [in] DWORD fdwUnderline, 
                              [in] DWORD fdwStrikeOut, 
                              [in] DWORD fdwCharSet, 
                              [in] DWORD fdwOutputPrecision, 
                              [in] DWORD fdwClipPrecision, 
                              [in] DWORD fdwQuality, 
                              [in] DWORD fdwPitchAndFamily, 
                              [in] LPWSTR lpszFace);

    [    entry("CreateHatchBrush")    ]
    HBRUSH CreateHatchBrush([in] long fnStyle,[in] COLORREF clrref);
    
    [    entry("CreateICA")    ]
    HDC A_CreateIC([in] LPSTR lpszDriver,[in] LPSTR lpszDevice, 
                        [in] LPSTR lpszOutput, [in] void * lpdvminit);

    [    entry("CreateICW")    ]
    HDC W_CreateIC([in] LPWSTR lpszDriver,[in] LPWSTR lpszDevice, 
                        [in] LPWSTR lpszOutput, [in] void * lpdvminit);

    [    entry("CreateMetaFileA")    ]
    HDC A_CreateMetaFile([in] LPSTR lpszFile);

    [    entry("CreateMetaFileA")    ]
    HDC W_CreateMetaFile([in] LPWSTR lpszFile);

    [    entry("CreatePalette")    ]
    HPALETTE CreatePalette([in] LPVOID lplgpl);

    [    entry("CreatePen")    ]
    HPEN CreatePen([in] long fnPenStyle, [in] long nWidth,[in] COLORREF crColor);

    [    entry("CreatePenIndirect")    ]
    HPEN CreatePenIndirect([in] LPVOID lplgpn);

    [    entry("ExtCreatePen")    ]
	HPEN ExtCreatePen(
			[in] DWORD dwPenStyle,      // pen style
			[in] DWORD dwWidth,         // pen width
			[in] LPVOID lplb,  // brush attributes
			[in] DWORD dwStyleCount,    // length of custom style array
			[in] DWORD *lpStyle);   // custom style array

    [    entry("CreatePolyPolygonRgn")    ]
    HRGN CreatePolyPolygonRgn([in] POINTAPI *lppt,
                                        [in] long *lpPolyCounts, 
                                        [in] long nCount, 
                                        [in] long fnPolyFillMode);

    [    entry("CreatePatternBrush")    ]
    HBRUSH CreatePatternBrush([in] HBITMAP hbmp);

    [    entry("CreateRectRgn")    ]
    HRGN CreateRectRgn([in] long nLeftRect, [in] long nTopRect,
                                 [in] long nRightRect, [in] long nBottomRect);

    [    entry("CreateRectRgnIndirect")    ]
    HRGN CreateRectRgnIndirect([in] RECT * lprc);

    [    entry("CreateRoundRectRgn")    ]
    HRGN CreateRoundRectRgn([in] long nLeftRect, [in] long nTopRect,
                                      [in] long nRightRect, [in] long nBottomRect,
                                      [in] long nWidthEllipse, 
                                      [in] long nHeightEllipse);

    [    entry("CreateScalableFontResourceA")    ]
    BOOL A_CreateScalableFontResource([in] DWORD fdwHidden, 
                                              [in] LPSTR lpszFontRes, 
                                              [in] LPSTR lpszFontFile, 
                                              [in] LPSTR lpszCurrentPath);

    [    entry("CreateScalableFontResourceW")    ]
    BOOL W_CreateScalableFontResource([in] DWORD fdwHidden, 
                                              [in] LPWSTR lpszFontRes, 
                                              [in] LPWSTR lpszFontFile, 
                                              [in] LPWSTR lpszCurrentPath);
    [    entry("CreateSolidBrush")    ]
    HBRUSH CreateSolidBrush([in] COLORREF crColor);

    [    entry("DeleteDC")    ]
    BOOL DeleteDC([in] HDC hdc);

    [    entry("DeleteMetaFile")    ]
    BOOL DeleteMetaFile([in] HMETAFILE hmf);

    [    entry("DeleteObject")    ]
    BOOL DeleteObject([in] HGDIOBJ hgdiobj);

    [    entry("DeviceCapabilitiesA")    ]
    long A_DeviceCapabilities([in] LPSTR pDevice, 
                                   [in] LPSTR pPort, 
                                   [in] WORD fwCapability, 
                                   [in] void * pOutput, 
                                   [in] void * pDevMode);
    [entry("DeviceCapabilitiesW")    ]
    long W_DeviceCapabilities([in] LPWSTR pDevice, 
                                   [in] LPWSTR pPort, 
                                   [in] WORD fwCapability, 
                                   [in] void * pOutput, 
                                   [in] void * pDevMode);

    [
    entry("DrawEscape")
    ]
    long DrawEscape([in] HDC hdc, [in] long nEscape, 
                           [in] long cbInput,[in] LPSTR lpszInData);

    [
    entry("Ellipse")
    ]
    BOOL Ellipse([in] HDC hdc,
                 [in] long nLeft, [in] long nTop,
                 [in] long nRight, [in] long nBottom);


    [

    entry("EnumFontFamiliesExA")
    ]
    long A_EnumFontFamiliesEx([in] HDC hdc, 
                                   [in] LPVOID lpLogFont,
                                   [in] long lpEnumFontFamExProc, 
                                   [in] long lParam,
                                   [in] DWORD dwFlags);
    [

    entry("EnumFontFamiliesExW")
    ]
    long W_EnumFontFamiliesEx([in] HDC hdc, 
                                   [in] LPVOID lpLogFont,
                                   [in] long lpEnumFontFamExProc, 
                                   [in] long lParam,
                                   [in] DWORD dwFlags);

    [

    entry("EnumFontFamiliesA")
    ]
    long A_EnumFontFamilies([in] HDC hdc, 
                                 [in] LPSTR lpszFamily, 
                                 [in] long lpEnumFontFamProc, 
                                 [in] long lParam);

    [

    entry("EnumFontFamiliesW")
    ]
    long W_EnumFontFamilies([in] HDC hdc, 
                                 [in] LPWSTR lpszFamily, 
                                 [in] long lpEnumFontFamProc, 
                                 [in] long lParam);

    [

    entry("EnumFontsA")
    ]
    long A_EnumFonts([in] HDC hdc, 
                          [in] LPSTR lpFaceName,  
                          [in] long lpFontFunc, 
                          [in] long lParam);
    [

    entry("EnumFontsW")
    ]
    long W_EnumFonts([in] HDC hdc, 
                          [in] LPWSTR lpFaceName,  
                          [in] long lpFontFunc, 
                          [in] long lParam);

    [

    entry("EnumObjects")
    ]
    long EnumObjects([in] HDC hdc, 
							[in] long nObjectType, 
							[in] long lpObjectFunc,	
							[in] long lParam);

    [

    entry("EqualRgn")
    ]
    BOOL EqualRgn([in] HRGN hSrcRgn1,[in] HRGN hSrcRgn2);


    [

    entry("ExcludeClipRect")
    ]
    long ExcludeClipRect([in] HDC hdc, 
								[in] long nLeftRect, [in] long nTopRect,
								[in] long nRightRect, [in] long nBottomRect);

    [

    entry("ExtCreateRegion")
    ]
    HRGN ExtCreateRegion([in] LPVOID lpXform,
								[in] DWORD nCount,
								[in] LPVOID lpRgnData);

    [

    entry("ExtFloodFill")
    ]
    BOOL ExtFloodFill([in] HDC hdc,
                                [in] long nXStart, [in] long nYStart,
                                [in] COLORREF clrref,
                                [in] UINT fuFillType);
	
    [

    entry("FillRgn")
    ]
    BOOL FillRgn([in] HDC hdc,[in] HRGN hrgn,[in] HBRUSH hbr);

    [

    entry("FloodFill")
    ]
    BOOL FloodFill([in] HDC hdc,
                             [in] long nXStart, [in] long nYStart,
                             [in] COLORREF clrref);

    [

    entry("FrameRgn")
    ]
    BOOL FrameRgn([in] HDC hdc, 
						 [in] HRGN hrgn, 
						 [in] HBRUSH hbr, 
						 [in] long nWidth, 
						 [in] long nHeight);

    [

    entry("GetROP2")
    ]
    long GetROP2([in] HDC hdc);

    [

    entry("GetAspectRatioFilterEx")
    ]
    BOOL GetAspectRatioFilterEx([in] HDC hdc, 
									    [out] SIZE * lpAspectRatio);

    [

    entry("GetBkColor")
    ]
    COLORREF GetBkColor([in] HDC hdc);

    [

    entry("GetBkMode")
    ]
    long GetBkMode([in] HDC hdc);

    [

    entry("GetBitmapBits")
    ]
    LONG GetBitmapBits([in] HBITMAP hbm, 
								 [in] LONG cbBuffer, 
								 [out] BYTE * lpvBits);

    [

    entry("GetBitmapDimensionEx")
    ]
    BOOL GetBitmapDimensionEx([in] HBITMAP hbm, 
									  [out] SIZE * lpDimension);

    [

    entry("GetBoundsRect")
    ]
    UINT GetBoundsRect([in] HDC hdc, 
							   [out] RECT * lprc, 
							   [in] UINT flags);

    [

    entry("GetBrushOrgEx")
    ]
    BOOL GetBrushOrgEx([in] HDC hdc,[out] POINTAPI * lpPoint);

    [

    entry("GetCharWidthA")
    ]
    BOOL A_GetCharWidth([in] HDC hdc, 
							  [in] UINT iFirstChar, 
							  [in] UINT iLastChar, 
							  [out] long * lpBuffer);
    [

    entry("GetCharWidthW")
    ]
    BOOL W_GetCharWidth([in] HDC hdc, 
							  [in] UINT iFirstChar, 
							  [in] UINT iLastChar, 
							  [out] long * lpBuffer);

    [

    entry("GetCharWidth32A")
    ]
    BOOL A_GetCharWidth32([in] HDC hdc,
	 						    [in] UINT iFirstChar, 
							    [in] UINT iLastChar, 
								[out] long * lpBuffer);

    [

    entry("GetCharWidth32W")
    ]
    BOOL W_GetCharWidth32([in] HDC hdc,
	 						    [in] UINT iFirstChar, 
							    [in] UINT iLastChar, 
								[out] long * lpBuffer);

    [

    entry("GetClipBox")
    ]
    long GetClipBox([in] HDC hdc,[out] RECT * lprc);

    [

    entry("GetClipRgn")
    ]
    long GetClipRgn([in] HDC hdc,[in] HRGN hrgn);

    [

    entry("GetMetaRgn")
    ]
    long GetMetaRgn([in] HDC hdc, [in] HRGN hrgn);

    [

    entry("GetCurrentObject")
    ]
    HGDIOBJ GetCurrentObject([in] HDC hdc, [in] UINT uObjectType);


    [

    entry("GetCurrentPositionEx")
    ]
    BOOL GetCurrentPositionEx([in] HDC hdc,[out] POINTAPI * lpPoint);

    [

    entry("GetDeviceCaps")
    ]
    long GetDeviceCaps([in] HDC hdc,[in] long iCapability);

    [

    entry("GetDIBits")
    ]
    long GetDIBits([in] HDC hdc, 
						   [in] HBITMAP hbm, 
						   [in] UINT uStartScan, 
						   [in] UINT cScanLines, 
                           [out] LPVOID lpvBits, //LPBITMAPINFO
                           [in,out] LPVOID lpbi, //LPBITMAPINFO
						   [in] UINT uUsage);

    [

    entry("GetFontData")
    ]
    DWORD GetFontData([in] HDC hdc, 
							 [in] DWORD dwTable, 
							 [in] DWORD dwOffset, 
							 [out] LPVOID lpvBuffer, 
							 [in] DWORD cbData);

    [

    entry("GetGlyphOutlineA")
    ]
    DWORD A_GetGlyphOutline([in] HDC hdc, 
								 [in] UINT uChar, 
								 [in] UINT uFormat, 
								 [out] LPVOID  lpgm, 
								 [in] DWORD cbBuffer, 
								 [out] LPVOID lpvBuffer, 
								 [in] LPVOID  lpmat2);
    [

    entry("GetGlyphOutlineW")
    ]
    DWORD W_GetGlyphOutline([in] HDC hdc, 
								 [in] UINT uChar, 
								 [in] UINT uFormat, 
								 [out] LPVOID  lpgm, 
								 [in] DWORD cbBuffer, 
								 [out] LPVOID lpvBuffer, 
								 [in] LPVOID  lpmat2);

    [

    entry("GetGraphicsMode")
    ]
    long GetGraphicsMode([in] HDC hdc);

    [

    entry("GetMapMode")
    ]
    long GetMapMode([in] HDC hdc);

    [

    entry("GetMetaFileBitsEx")
    ]
    UINT GetMetaFileBitsEx([in] HMETAFILE hmf, [in] UINT nSize,
                                   [out] LPVOID lpvData);

    [

    entry("GetNearestColor")
    ]
    COLORREF GetNearestColor([in] HDC hdc, [in] COLORREF clrref);

    [

    entry("GetNearestPaletteIndex")
    ]
    UINT  GetNearestPaletteIndex([in] HPALETTE hpal, [in] COLORREF clrref);

    [

    entry("GetObjectType")
    ]
    DWORD GetObjectType([in] HGDIOBJ hgdiobj);

    [

    entry("GetOutlineTextMetricsA")
    ]
    UINT A_GetOutlineTextMetrics([in] HDC hdc,[in] UINT iFlags,[out] LPVOID lpOTM);

    [

    entry("GetOutlineTextMetricsW")
    ]
    UINT W_GetOutlineTextMetrics([in] HDC hdc,[in] UINT iFlags,[out] LPVOID lpOTM);

    [

    entry("GetPaletteEntries")
    ]
    UINT  GetPaletteEntries([in] HPALETTE hpal, 
								   [in] UINT iStartIndex, [in] UINT nEntries, 
								   [out] LPVOID lppe);

    [

    entry("GetPixel")
    ]
    COLORREF GetPixel([in] HDC hdc,
                             [in] long nXPos, [in] long nYPos);

    [

    entry("GetPixelFormat")
    ]
    long  GetPixelFormat([in] HDC hdc);
	
    [

    entry("DescribePixelFormat")
    ]
	long DescribePixelFormat([in] HDC hdc,
					[in] long iPixelFormat,
					[in] UINT nBytes, 
					[out] LPVOID ppfd);

    [

    entry("GetPolyFillMode")
    ]
    long GetPolyFillMode([in] HDC hdc);

    [

    entry("GetRasterizerCaps")
    ]
    BOOL  GetRasterizerCaps([out] LPVOID lprs,[in] UINT cp);

    [

    entry("GetRegionData")
    ]
    DWORD GetRegionData([in] HRGN hrgn,[in] DWORD dwCount,[out] LPVOID lpRgnData);

    [

    entry("GetRgnBox")
    ]
    long  GetRgnBox([in] HRGN hrgn,[out] RECT * lprc);

    [

    entry("GetStockObject")
    ]
    HGDIOBJ GetStockObject([in] long fnObject);

    [

    entry("GetStretchBltMode")
    ]
    long  GetStretchBltMode([in] HDC hdc);

    [

    entry("GetSystemPaletteEntries")
    ]
    UINT  GetSystemPaletteEntries([in] HDC hdc, 
										 [in] UINT iStartIndex, 
										 [in] UINT nEntries,	
										 [out] LPVOID lppe);

    [

    entry("GetSystemPaletteUse")
	]
    UINT  GetSystemPaletteUse([in] HDC hdc);

    [

    entry("GetTextCharacterExtra")
    ]
    long  GetTextCharacterExtra([in] HDC hdc);

    [

    entry("GetTextAlign")
    ]
    UINT  GetTextAlign([in] HDC hdc);

    [

    entry("GetTextColor")
    ]
    COLORREF GetTextColor([in] HDC hdc);

    [

    entry("GetTextExtentPointA")
    ]
    BOOL  A_GetTextExtentPoint([in] HDC hDC, [in] LPSTR lpString, 
									[in] long cbString,[out] SIZEL *lpSize);

    [

    entry("GetTextExtentPointW")
    ]
    BOOL  W_GetTextExtentPoint([in] HDC hDC, [in] LPWSTR lpString, 
									[in] long cbString,[out] SIZEL *lpSize);

    [

    entry("GetTextExtentPoint32A")
    ]
    BOOL  A_GetTextExtentPoint32([in] HDC hdc, [in] LPSTR lpString, 
									    [in] long cbString, [out] SIZEL *lpSize);

    [

    entry("GetTextExtentPoint32W")
    ]
    BOOL  W_GetTextExtentPoint32([in] HDC hdc, [in] LPWSTR lpString, 
									    [in] long cbString, [out] SIZEL *lpSize);

    [

    entry("GetTextExtentExPointA")
    ]
    BOOL  A_GetTextExtentExPoint([in] HDC hdc, [in] LPSTR lpszStr,	
									  [in] long cchString, [in] long nMaxExtent,	
									  [out] long * lpnFit,	[out] long * alpDx,	
									  [out] SIZEL * lpSize);

    [

    entry("GetTextExtentExPointW")
    ]
    BOOL  W_GetTextExtentExPoint([in] HDC hdc, [in] LPWSTR lpszStr,	
									  [in] long cchString, [in] long nMaxExtent,	
									  [out] long * lpnFit,	[out] long * alpDx,	
									  [out] SIZEL * lpSize);

    [

    entry("GetWindowExtEx")
    ]
    BOOL  GetWindowExtEx([in] HDC hdc,[out] SIZE * lpSize);

    [

    entry("IntersectClipRect")
    ]
    long  IntersectClipRect([in] HDC hdc, 
				[in] long nLeftRect, 
				[in] long nTopRect, 
				[in] long nRightRect, 
				[in] long nBottomRect);

    [

    entry("InvertRgn")
    ]
    BOOL InvertRgn([in] HDC hdc, [in] HRGN hrgn);

    [

    entry("LineTo")
    ]
    BOOL LineTo([in] HDC hdc, [in] long X, [in] long Y);

    [

    entry("MaskBlt")
    ]
    BOOL  MaskBlt([in] HDC hdcDest,
                           [in] long nXDest, [in] long nYDest,
                           [in] long nWidth, [in] long nHeight,
                           [in] HDC hdcSrc,
                           [in] long nXSrc, [in] long nYSrc,
                           [in] HBITMAP hbmMask,
                           [in] long xMask, [in] long yMask,
                           [in] DWORD dwRop);

    [

    entry("PlgBlt")
    ]
    BOOL  PlgBlt([in] HDC hdcDest,
                          [in] POINTAPI * lpPoint,
                          [in] HDC hdcSrc,
                          [in] long nXSrc, [in] long nYSrc,
                          [in] long nWidth, [in] long nHeight,
                          [in] HBITMAP hbmMask,
                          [in] long sMask, [in] long yMask);

    [

    entry("OffsetClipRgn")
    ]
    long  OffsetClipRgn([in] HDC hdc,
				[in] long nXOffset, 
				[in] long nYOffset);

    [

    entry("OffsetRgn")
    ]
    long  OffsetRgn([in] HRGN hrgn,
				[in] long nXOffset,
				[in] long nYOffset);

    [

    entry("PatBlt")
    ]
    BOOL PatBlt([in] HDC hdc,
                          [in] long nLeftRect, [in] long nTopRect,
                          [in] long nWidth, [in] long nHeight,
                          [in] DWORD fdwRop);

    [

    entry("Pie")
    ]
    BOOL Pie([in] HDC hDC,
                    [in] long nLeftRect, [in] long nTopRect,
                    [in] long nRightRect, [in] long nBottomRect,
                    [in] long nXRadial1, [in] long nYRadial1,
                    [in] long nXRadial2, [in] long nYRadial2);

    [

    entry("PlayMetaFile")
    ]
    BOOL  PlayMetaFile([in] HDC hdc, [in] HMETAFILE hmf);

    [

    entry("PaintRgn")
    ]
    BOOL PaintRgn([in] HDC hdc, [in] HRGN hrgn);

    [

    entry("PolyPolygon")
    ]
    BOOL PolyPolygon([in] HDC hdc,[in] POINTAPI * lpPoints,
                            [in] long * lpPolyCounts, [in] long cPoints);

    [

    entry("PtInRegion")
    ]
    BOOL PtInRegion([in] HRGN hrgn, [in] long X, [in] long Y);

    [

    entry("PtVisible")
    ]
    BOOL PtVisible([in] HDC hdc, [in] long X, [in] long Y);

    [

    entry("RectInRegion")
    ]
    BOOL RectInRegion([in] HRGN hrgn, [in] RECT * lprc);

    [

    entry("RectVisible")
    ]
    BOOL RectVisible([in] HDC hdc, [in] RECT * lprc);

    [

    entry("Rectangle")
    ]
    BOOL Rectangle([in] HDC hdc,
                   [in] long nLeft, [in] long nTop,
                   [in] long nRight, [in] long nBottom);

    [

    entry("RestoreDC")
    ]
    BOOL RestoreDC([in] HDC hdc, [in] long nSavedDC);

    [

    entry("ResetDCA")
    ]
    HDC A_ResetDC([in] HDC hdc,[in] void *lpInitData); //DEVMODE

    [

    entry("RealizePalette")
    ]
    UINT RealizePalette([in] HDC hdc);

    [

    entry("RemoveFontResourceA")
    ]
    BOOL A_RemoveFontResource([in] LPSTR lpszFontname);

    [

    entry("RemoveFontResourceW")
    ]
    BOOL W_RemoveFontResource([in] LPWSTR lpszFontname);

    [

    entry("AddFontResourceExA")
    ]
	long A_AddFontResourceEx(
				[in] LPSTR lpszFilename, // font file name
				[in] DWORD fl,             // font characteristics
				[in] long pdv); //Reserved
    [

    entry("AddFontResourceExW")
    ]
	long W_AddFontResourceEx(
				[in] LPWSTR lpszFilename, // font file name
				[in] DWORD fl,             // font characteristics
				[in] long pdv); //Reserved

    [

    entry("RemoveFontResourceExA")
    ]
	BOOL A_RemoveFontResourceEx(
				[in] LPSTR lpFileName,  // name of font file
				[in] DWORD fl,            // font characteristics
				[in] long pdv);            // Reserved.
    [

    entry("RemoveFontResourceExW")    ]
	BOOL W_RemoveFontResourceEx(
				[in] LPSTR lpFileName,  // name of font file
				[in] DWORD fl,            // font characteristics
				[in] long pdv);            // Reserved.

    [

    entry("RoundRect")
    ]
    BOOL  RoundRect([in] HDC hdc,
                             [in] long nLeft, [in] long nTop,
                             [in] long nRight, [in] long nBottom,
                             [in] long nWidth, [in] long nHeight);

    [

    entry("ResizePalette")
    ]
    BOOL ResizePalette([in] HPALETTE hpal, [in] UINT nEntries);

    [

    entry("SaveDC")
    ]
    long  SaveDC([in] HDC hdc);

    [

    entry("SelectClipRgn")
    ]
    long  SelectClipRgn([in] HDC hdc, [in] HRGN hrgn);

    [

    entry("ExtSelectClipRgn")
    ]
    long  ExtSelectClipRgn([in] HDC hdc, [in] HRGN hrgn, [in] long fnMode);

    [

    entry("SetMetaRgn")
    ]
    long  SetMetaRgn([in] HDC hdc);

    [

    entry("SelectObject")
    ]
    HGDIOBJ SelectObject([in] HDC hdc, [in] HGDIOBJ hgdiobj);

    [

    entry("SelectPalette")
    ]
    HPALETTE SelectPalette([in] HDC hdc, [in] HPALETTE hpal, 
								  [in] BOOL bForceBackground);

    [

    entry("SetBkColor")
    ]
    COLORREF SetBkColor([in] HDC hdc, [in] COLORREF clrref);

    [

    entry("SetBkMode")
    ]
    long  SetBkMode([in] HDC hdc, [in] long fnBkMode);

    [

    entry("SetBitmapBits")
    ]
    LONG  SetBitmapBits([in] HBITMAP hbm, [in] LONG cbBuffer, [in] BYTE * lpvBits);

    [

    entry("SetBoundsRect")
    ]
    UINT  SetBoundsRect([in] HDC hdc, [in] RECT * lprc, [in] UINT flags);

    [

    entry("SetDIBits")
    ]
    long  SetDIBits([in] HDC hdc, 
						   [in] HBITMAP hbmp,	
						   [in] UINT uStartScan,
						   [in] UINT cScanLines,
                           [in] LPVOID lpvBits,
                           [in] LPVOID lpbmi,
						   [in] UINT fuColorUse); 	

    [

    entry("SetDIBitsToDevice")
    ]
    long  SetDIBitsToDevice([in] HDC hdc, 
									[in] long XDest,	
									[in] long YDest,	
									[in] DWORD dwWidth,	
									[in] DWORD dwHeight,	
									[in] long XSrc,	
									[in] long YSrc,	
									[in] UINT uStartScan,	
									[in] UINT cScanLines,	
                                    [in] LPVOID lpvBits,
                                    [in] LPVOID lpbmi,
									[in] UINT fuColorUse);

    [

    entry("SetMapperFlags")
    ]
    DWORD SetMapperFlags([in] HDC hdc, [in] DWORD dwFlags);

    [

    entry("SetGraphicsMode")
    ]
    long  SetGraphicsMode([in] HDC hdc, [in] long iMode);

    [

    entry("SetMapMode")
    ]
    long   SetMapMode([in] HDC hdc, [in] long fnMapMode);

    [

    entry("SetMetaFileBitsEx")
    ]
    HMETAFILE  SetMetaFileBitsEx([in] UINT nSize, [in] BYTE * lpData);

    [

    entry("SetPaletteEntries")
    ]
    UINT  SetPaletteEntries([in] HPALETTE hpal, 
								   [in] UINT iStart, [in] UINT cEntries, 
								   [in] LPVOID lppe);

    [

    entry("SetPixel")
    ]
    COLORREF SetPixel([in] HDC hdc,
                             [in] long nXPos, [in] long nYPos,
                             [in] COLORREF clrref);

    [

    entry("SetPixelV")
    ]
    BOOL   SetPixelV([in] HDC hdc, [in] long X, [in] long Y, [in] COLORREF crColor);

    [

    entry("SetPixelFormat")
    ]
    BOOL  SetPixelFormat([in] HDC hdc, [in] long iPixelFormat,[in]  LPVOID ppfd);

    [

    entry("SetPolyFillMode")
    ]
    long  SetPolyFillMode([in] HDC hdc, [in] long iPolyFillMode);

    [

    entry("StretchBlt")
    ]
    BOOL StretchBlt([in] HDC hdcDest,
                           [in] long nXOriginDest, [in] long nYOriginDest,
                           [in] long bWidthDest, [in] long nHeightDest,
                           [in] HDC hdcSrc,
                           [in] long nXOriginSrc, [in] long nYOriginSrc,
                           [in] long nWidthSrc, [in] long nHeightSrc,
                           [in] DWORD dwRop);

    [

    entry("SetRectRgn")
    ]
    BOOL   SetRectRgn([in] HRGN hrgn, 
							 [in] long nLeftRect,	
							 [in] long nTopRect,	
							 [in] long nRightRect,	
							 [in] long nBottomRect);
  
    [

    entry("StretchDIBits")
    ]
    long   StretchDIBits([in] HDC hdc, 
							   [in] long XDest,	
							   [in] long YDest,	
							   [in] long nDestWidth,	
							   [in] long nDestHeight,	
							   [in] long XSrc,	
							   [in] long YSrc,	
							   [in] long nSrcWidth,	
							   [in] long nSrcHeight,	
                               [in] LPVOID lpBits,
                               [in] LPVOID lpBitsInfo,
							   [in] UINT iUsage,	
							   [in] DWORD dwRop);

    [

    entry("SetROP2")
    ]
    long   SetROP2([in] HDC hdc, [in] long fnDrawMode);

    [

    entry("SetStretchBltMode")
    ]
    long   SetStretchBltMode([in] HDC hdc, [in] long fnStretchMode);

    [

    entry("SetSystemPaletteUse")
    ]
    UINT  SetSystemPaletteUse([in] HDC hdc, [in] UINT uUsage);

    [

    entry("SetTextCharacterExtra")
    ]
    long   SetTextCharacterExtra([in] HDC hdc, [in] long nCharExtra);

    [

    entry("SetTextColor")
    ]
    COLORREF SetTextColor([in] HDC hdc, [in] COLORREF clrref);

    [

    entry("SetTextAlign")
    ]
    UINT  SetTextAlign([in] HDC hdc, [in] UINT fMode);

    [

    entry("SetTextJustification")
    ]
    BOOL  SetTextJustification([in] HDC hdc, [in] long nBreakExtra, [in] long nBreakCount);

    [

    entry("UpdateColors")
    ]
    BOOL  UpdateColors([in] HDC hdc);

    [

    entry("EnumMetaFile")
    ]
    BOOL  EnumMetaFile([in] HDC hdc, 
							  [in] HMETAFILE hmf, 
							  [in] long lpMetaFunc, 
							  [in] long lParam);

    // Enhanced Metafile Function Declarations
    [

    entry("CloseEnhMetaFile")
    ]
    HENHMETAFILE CloseEnhMetaFile([in] HDC hdc);
	
    [

    entry("CopyEnhMetaFileA")
    ]
    HENHMETAFILE A_CopyEnhMetaFile([in] HENHMETAFILE hemf,
                                        [in] LPSTR lpszFile);
    [

    entry("CopyEnhMetaFileW")
    ]
    HENHMETAFILE W_CopyEnhMetaFile([in] HENHMETAFILE hemf,
                                        [in] LPWSTR lpszFile);

    [

    entry("CreateEnhMetaFileA")
    ]
    HDC A_CreateEnhMetaFile([in] HDC hdc, [in] LPSTR lpFileName,
                                   [in] long * lpRect,
                                   [in] LPSTR lpDescription);
    [

    entry("CreateEnhMetaFileW")
    ]
    HDC W_CreateEnhMetaFile([in] HDC hdc, [in] LPWSTR lpFileName,
                                   [in] long * lpRect,
                                   [in] LPWSTR lpDescription);

    [

    entry("DeleteEnhMetaFile")
    ]
    BOOL  DeleteEnhMetaFile([in] HENHMETAFILE hemf);
	
    [

    entry("EnumEnhMetaFile")
    ]
    BOOL  EnumEnhMetaFile([in] HDC hdc, 
								 [in] HENHMETAFILE hemf, 
								 [in] long lpEnhMetaFunc, 
								 [in] LPVOID lpData, 
								 [in] RECT * lpRect);
	
    [

    entry("GetEnhMetaFileA")
    ]
    HENHMETAFILE A_GetEnhMetaFile([in] LPSTR lpszMetaFile);

    [

    entry("GetEnhMetaFileW")
    ]
    HENHMETAFILE W_GetEnhMetaFile([in] LPWSTR lpszMetaFile);

    [

    entry("GetEnhMetaFileBits")
    ]
    UINT  GetEnhMetaFileBits([in] HENHMETAFILE hemf,
                                    [in] UINT cbBuffer,
                                    [out] BYTE * lpbBuffer);

    [

    entry("GetEnhMetaFileDescriptionA")
    ]
    UINT  A_GetEnhMetaFileDescription([in] HENHMETAFILE hemf,
                                           [in] UINT cchBuffer,
                                           [out] LPSTR lpszDescription);
    [

    entry("GetEnhMetaFileDescriptionW")
    ]
    UINT  W_GetEnhMetaFileDescription([in] HENHMETAFILE hemf,
                                           [in] UINT cchBuffer,
                                           [out] LPWSTR lpszDescription);

    [

    entry("GetEnhMetaFileHeader")
    ]
    UINT  GetEnhMetaFileHeader([in] HENHMETAFILE hemf,
                                      [in] UINT cbBuffer,
                                      [out] LPVOID lpemnh);

    [

    entry("GetEnhMetaFilePaletteEntries")
    ]
    UINT  GetEnhMetaFilePaletteEntries([in] HENHMETAFILE hemf,
											  [in] UINT cEntries, 
											  [out] LPVOID lppe);

    [

    entry("GetWinMetaFileBits")
    ]
    UINT  GetWinMetaFileBits([in] HENHMETAFILE hemf,
                                    [in] UINT cbBuffer,
                                    [out] BYTE * lpbBuffer,
                                    [in] long fnMapMode,
                                    [in] HDC hdc);

    [

    entry("PlayEnhMetaFile")
    ]
    BOOL  PlayEnhMetaFile([in] HDC hdc, [in] HENHMETAFILE hemf,
                                 [in] RECT * lpRect);

    [

    entry("PlayEnhMetaFileRecord")
    ]
    BOOL  PlayEnhMetaFileRecord([in] HDC hdc, 
					[in] LPVOID lpHandletable,
					[in] LPVOID lpEnhMetaRecord,
					[in] UINT nHandles);

    [

    entry("SetEnhMetaFileBits")
    ]
    HENHMETAFILE SetEnhMetaFileBits([in] UINT cbBuffer,
                                           [in] BYTE * lpData);

    [

    entry("SetWinMetaFileBits")
    ]
    HENHMETAFILE  SetWinMetaFileBits([in] UINT cbBuffer,
                                            [in] BYTE * lpbBuffer,
                                            [in] HDC hdc,
                                            [in] LPVOID lpmfp);

    [

    entry("GdiComment")
    ]
    BOOL  GdiComment([in] HDC hdc, [in] UINT cbSize,
                            [in] BYTE * lpData);

    [entry("GetTextMetricsA")]
    BOOL A_GetTextMetrics([in] HDC hdc, [out] void *lptm); //A_TEXOMETRIC

    [entry("GetTextMetricsW")]
    BOOL W_GetTextMetrics([in] HDC hdc, [out] void *lptm); //W_TEXOMETRIC

    // new GDI 

    [

    entry("AngleArc")
    ]
    BOOL AngleArc([in] HDC hdc, 
				[in] long X, 
				[in] long Y, 
				[in] DWORD dwRadius, 
				[in] FLOAT eStartAngle, 
				[in] FLOAT eSweepAngle);

    [

    entry("PolyPolyline")
    ]
    BOOL PolyPolyline([in] HDC hdc, [in] POINTAPI * lppt, [in] UINT *lpdwPolyPoints, [in] DWORD cCount);

    [

    entry("GetWorldTransform")
    ]
    BOOL GetWorldTransform([in] HDC hdc, [out] LPVOID lpXform);

    [

    entry("SetWorldTransform")
    ]
    BOOL SetWorldTransform([in] HDC hdc, [in] LPVOID lpXform);

    [

    entry("ModifyWorldTransform")
    ]
    BOOL ModifyWorldTransform([in] HDC hdc, [in] LPVOID lpXform, [in] DWORD iMode);

    [

    entry("CombineTransform")
    ]
    BOOL CombineTransform([out] LPVOID lpxformResult , [in] LPVOID lpXform1, [in] LPVOID lpXform2);

    [

    entry("CreateDIBSection")
    ]
    HBITMAP CreateDIBSection([in] HDC hdc, 
                                    [in] LPVOID pbmi,
									[in] UINT iUsage,	
									[out] void * ppvBits,	
									[in] HANDLE hSection,
									[in] DWORD dwOffset);

    [

    entry("GetDIBColorTable")
    ]
    UINT GetDIBColorTable([in] HDC hdc, 
								 [in] UINT uStartIndex,	
								 [in] UINT cEntries,	
								 [out] LPVOID pColors);	

    [

    entry("SetDIBColorTable")
    ]
    UINT SetDIBColorTable([in] HDC hdc, 
								 [in] UINT uStartIndex,	
								 [in] UINT cEntries,	
								 [in] LPVOID pColors);

    [

    entry("SetColorAdjustment")
    ]
    BOOL SetColorAdjustment([in] HDC hdc, [in] LPVOID lpca);

    [

    entry("GetColorAdjustment")
    ]
    BOOL GetColorAdjustment([in] HDC hdc, [out] LPVOID lpca);

    [

    entry("CreateHalftonePalette")
    ]
    HPALETTE CreateHalftonePalette([in] HDC hdc);

    [

    entry("StartDocA")
    ]
    long A_StartDoc([in] HDC hdc, [in] LPVOID lpdi);

    [

    entry("StartDocW")
    ]
    long W_StartDoc([in] HDC hdc, [in] LPVOID lpdi);

    [

    entry("EndDoc")
    ]
    long EndDoc([in] HDC hdc);

    [

    entry("StartPage")
    ]
    long StartPage([in] HDC hdc);

    [

    entry("EndPage")
    ]
    long EndPage([in] HDC hdc);

    [

    entry("AbortDoc")
    ]
    long AbortDoc([in] HDC hdc);

    [

    entry("SetAbortProc")
    ]
    long SetAbortProc([in] HDC hdc, [in] long *lpAbortProc);

    [

    entry("AbortPath")
    ]
    BOOL AbortPath([in] HDC hDC);

    [

    entry("ArcTo")
    //helpstring("Draws an elliptical arc (NT only)")
    ]
    BOOL ArcTo([in] HDC hDC,
                      [in] long nLeftRect, [in] long nTopRect,
                      [in] long nRightRect, [in] long nBottomRect,
                      [in] long nXRadial1, [in] long nYRadial1,
                      [in] long nXRadial2, [in] long nYRadial2);

    [

    entry("BeginPath")
    ]
    BOOL BeginPath([in] HDC hDC);

    [

    entry("CloseFigure")
    ]
    BOOL CloseFigure([in] HDC hDC);

    [

    entry("EndPath")
    ]
    BOOL EndPath([in] HDC hDC);

    [

    entry("FillPath")
    ]
    BOOL FillPath([in] HDC hDC);

    [

    entry("FlattenPath")
    ]
    BOOL FlattenPath([in] HDC hDC);

    [

    entry("GetPath")
    ]
    long GetPath([in] HDC hDC, [out] POINTAPI * lpPoints,
                       [out] BYTE * lpTypes, [in] long nSize);

    [

    entry("PathToRegion")
    ]
    HRGN PathToRegion([in] HDC hDC);

    [

    entry("PolyDraw")
    ]
    BOOL PolyDraw([in] HDC hDC, [in] POINTAPI * lppt,
                         [in] BYTE * lpbTypes, [in] long cCount);

    [

    entry("SelectClipPath")
    ]
    BOOL SelectClipPath([in] HDC hDC, [in] long iMode);

    [

    entry("SetArcDirection")
    ]
    long SetArcDirection([in] HDC hDC, [in] long ArcDirection);

    [

    entry("GetMiterLimit")
    ]
    BOOL GetMiterLimit([in] HDC hDC, [out] FLOAT * peOldLimit);

    [

    entry("SetMiterLimit")
    ]
    BOOL SetMiterLimit([in] HDC hDC, [in] FLOAT eNewLimit,
                              [out] FLOAT * peOldLimit);

    [

    entry("StrokeAndFillPath")
    ]
    BOOL StrokeAndFillPath([in] HDC hDC);

    [

    entry("StrokePath")
    ]
    BOOL StrokePath([in] HDC hDC);

    [

    entry("WidenPath")
    ]
    BOOL WidenPath([in] HDC hDC);

    // HPEN ExtCreatePen([in] DWORD, [in] DWORD, [in] LOGBRUSH *, [in] DWORD, [in] UINT *);

    [

    entry("GetArcDirection")
    ]
    long GetArcDirection([in] HDC hDC);

    [

    entry("GetObjectA")
	]
    long A_GetObject([in] HGDIOBJ hgdiobj, [in] long cbBuffer, [out] void * lpvObject);

    [

    entry("GetObjectW")
	]
    long W_GetObject([in] HGDIOBJ hgdiobj, [in] long cbBuffer, [out] void * lpvObject);

    [

    entry("MoveToEx")
    ]
    BOOL    MoveToEx([in] HDC hdc, [in] long X, [in] long Y, [out] POINTAPI * lpPoint);

    [

    entry("MoveToEx")
    ]
    BOOL    MoveToExNoRet([in] HDC hdc, [in] long X, [in] long Y, [out] LPVOID lpNull);

    [

    entry("TextOutA")
    ]
    BOOL  A_TextOut([in] HDC hdc, [in] long nXStart, [in] long nYStart,	
						 [in] LPSTR lpString, [in] long cbString);

    [

    entry("TextOutW")
    ]
    BOOL  W_TextOut([in] HDC hdc, [in] long nXStart, [in] long nYStart,	
						 [in] LPWSTR lpString, [in] long cbString);
	  
    [

    entry("ExtTextOutA")
    ]
    BOOL  A_ExtTextOut([in] HDC hdc, [in] long X,	[in] long Y,	
							[in] UINT fuOptions, [in] RECT *lprc,	
							[in] LPSTR lpString, [in] UINT cbCount,	
							[in] long * lpDx);
    [
	usesgetlasterror,
    entry("ExtTextOutW")
    ]
    BOOL  W_ExtTextOut([in] HDC hdc, [in] long X,	[in] long Y,	
							[in] UINT fuOptions, [in] RECT *lprc,	
							[in] LPWSTR lpString, [in] UINT cbCount,	
							[in] long * lpDx);
    [

    entry("PolyTextOutA")
    ]
    BOOL  A_PolyTextOut([in] HDC hdc, [in] LPVOID pptxt, 
							 [in] long cStrings);

    [

    entry("PolyTextOutW")
    ]
    BOOL  W_PolyTextOut([in] HDC hdc,
							[in] LPVOID pptxt, 
							[in] long cStrings);

    [

    entry("CreatePolygonRgn")
    ]
    HRGN  CreatePolygonRgn([in] POINTAPI * lppt, [in] long cPoints,	
								  [in] long fnPolyFillMode);

    [

    entry("DPtoLP")
    ]
    BOOL  DPtoLP([in] HDC hdc, [in,out] POINTAPI * lpPoint, [in] long nCount);

    [

    entry("LPtoDP")
    ]
    BOOL  LPtoDP([in] HDC hdc, [in,out] POINTAPI * lpPoint, [in] long nCount);
 
    [

    entry("Polygon")
    ]
    BOOL Polygon([in] HDC hdc, [in] POINTAPI * lpPoints,
                        [in] long cPoints);

    [

    entry("Polyline")
    ]
    BOOL  Polyline([in] HDC hdc, [in] POINTAPI * lppt, [in] long cPoints);

    [

    entry("PolyBezier")
    ]
    BOOL    PolyBezier([in] HDC hdc, [in] POINTAPI * lppt,
                              [in] DWORD cPoints);

    [

    entry("PolyBezierTo")
    ]
    BOOL    PolyBezierTo([in] HDC hdc, [in] POINTAPI * lppt,
                                [in] DWORD cPoints);

    [

    entry("PolylineTo")
    ]
    BOOL  PolylineTo([in] HDC hdc, [in] POINTAPI * lppt, [in] DWORD cCount);

    [

    entry("SetViewportExtEx")
    ]
    BOOL  SetViewportExtEx([in] HDC hdc, [in] long nXExtent, [in] long nYExtent, [in] LPVOID lpSize);

    [

    entry("SetViewportOrgEx")
    ]
    BOOL  SetViewportOrgEx([in] HDC hdc, [in] long X, [in] long Y, [out] LPVOID lpPoint);

    [

    entry("SetWindowExtEx")
    ]
    BOOL  SetWindowExtEx([in] HDC hdc, [in] long nXExtent, [in] long nYExtent, [out] LPVOID lpSize);

    [

    entry("SetWindowOrgEx")
    ]
    BOOL  SetWindowOrgEx([in] HDC hdc, [in] long X, [in] long Y, [out] LPVOID lpPoint);

    [

    entry("OffsetViewportOrgEx")
    ]
    BOOL  OffsetViewportOrgEx([in] HDC hdc, [in] long nXOffset, [in] long nYOffset, [out] LPVOID lpPoint);

    [

    entry("OffsetWindowOrgEx")
    ]
    BOOL  OffsetWindowOrgEx([in] HDC hdc, [in] long nXOffset, [in] long nYOffset, [in] LPVOID lpPoint);

    [

    entry("ScaleViewportExtEx")
    ]
    BOOL  ScaleViewportExtEx([in] HDC hdc, 
				[in] long Xnum, 
				[in] long Xdenom , 
				[in] long Ynum, 
				[in] long Ydenom, 
				[out] LPVOID lpSize);

    [

    entry("ScaleWindowExtEx")
    ]
    BOOL  ScaleWindowExtEx([in] HDC hdc, 
				[in] long Xnum, 
				[in] long Xdenom, 
				[in] long Ynum, 
				[in] long Ydenom, 
				[in] LPVOID lpSize);

    [

    entry("SetBitmapDimensionEx")
    ]
    BOOL  SetBitmapDimensionEx([in] HBITMAP hbm, [in] long nWidth, [in] long nHeight, [in] SIZE *lpSize);

    [

    entry("SetBrushOrgEx")
    ]
    BOOL  SetBrushOrgEx([in] HDC hdc, [in] long nXOrg, [in] long nYOrg, [in] POINTAPI * lpPoint);

    [

    entry("GetTextFaceA")
    ]
    long   A_GetTextFace([in] HDC hdc, [in] long nCount, [in] LPSTR lpFaceName);

    [

    entry("GetTextFaceW")
    ]
    long   W_GetTextFace([in] HDC hdc, [in] long nCount, [in] LPWSTR lpFaceName);

    [

    entry("GetKerningPairsA")
    ]
    DWORD A_GetKerningPairs([in] HDC hdc, [in] DWORD nNumpairs, [in] LPVOID lpkrnpair);

    [

    entry("GetKerningPairsW")
    ]
    DWORD W_GetKerningPairs([in] HDC hdc, [in] DWORD nNumpairs, [in] LPVOID lpkrnpair);

    [

    entry("GetDCOrgEx")
    ]
    BOOL  GetDCOrgEx([in] HDC hdc,[out] POINTAPI * lpPoint);

    [

    entry("UnrealizeObject")
	]
    BOOL UnrealizeObject([in] HGDIOBJ hgdiobj);

    [

    entry("GdiFlush")
	]
    BOOL  GdiFlush(void);

    [

    entry("GdiSetBatchLimit")
	]
    DWORD GdiSetBatchLimit([in] DWORD dwLimit);

    [

    entry("GdiGetBatchLimit")
	]
    DWORD GdiGetBatchLimit(void);

    [

    entry("SetICMMode")
    ]
    long SetICMMode([in] HDC hdc, [in] long iEnableICM);

    [

    entry("GetICMProfileA")
    ]
    BOOL A_GetICMProfile([in] HDC hdc,[in] DWORD *lpcbName ,[in] LPSTR lpszFilename);

    [

    entry("GetICMProfileW")
    ]
    BOOL W_GetICMProfile([in] HDC hdc,[in] DWORD *lpcbName ,[in] LPWSTR lpszFilename);

    [

    entry("SetICMProfileA")
    ]
    BOOL A_SetICMProfile([in] HDC hdc,[in] LPSTR lpFileName);

    [

    entry("SetICMProfileW")
    ]
    BOOL W_SetICMProfile([in] HDC hdc,[in] LPWSTR lpFileName);

    [

    entry("EnumICMProfilesA")
    ]
    long A_EnumICMProfiles([in] HDC hdc,
							   [in] long lpICMProfilesFunc,
							   [in] long lParam);
    [

    entry("EnumICMProfilesW")
    ]
    long W_EnumICMProfiles([in] HDC hdc,
							   [in] long lpICMProfilesFunc,
							   [in] long lParam);

    [

    entry("Escape")
    ]
    long Escape([in] HDC hdc, 
					   [in] long nEscape, 
					   [in] long cbInput, 
					   [in] LPSTR lpvInData, 
					   [out] LPVOID lpvOutData);

}; //END GDI module
