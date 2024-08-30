Attribute VB_Name = "API_GDI32"

#If VBA7 And Win64 Then

    Declare PtrSafe Function AbortDoc Lib "gdi32" Alias "AbortDoc" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function AbortPath Lib "gdi32" Alias "AbortPath" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
    Declare PtrSafe Function AngleArc Lib "gdi32" Alias "AngleArc" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Double, ByVal eSweepAngle As Double) As Long
    Declare PtrSafe Function AnimatePalette Lib "gdi32" Alias "AnimatePaletteA" (ByVal hPalette As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteColors As PALETTEENTRY) As Long
    Declare PtrSafe Function Arc Lib "gdi32" Alias "Arc" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
    Declare PtrSafe Function ArcTo Lib "gdi32" Alias "ArcTo" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
    Declare PtrSafe Function BeginPath Lib "gdi32" Alias "BeginPath" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function BitBlt Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Declare PtrSafe Function CancelDC Lib "gdi32" Alias "CancelDC" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function CheckColorsInGamut Lib "gdi32" Alias "CheckColorsInGamut" (ByVal hdc As LongPtr, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long
    Declare PtrSafe Function ChoosePixelFormat Lib "gdi32" Alias "ChoosePixelFormat" (ByVal hDC As LongPtr, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
    Declare PtrSafe Function Chord Lib "gdi32" Alias "Chord" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
    Declare PtrSafe Function CloseEnhMetaFile Lib "gdi32" Alias "CloseEnhMetaFile" (ByVal hdc As LongPtr) As LongPtr
    Declare PtrSafe Function CloseFigure Lib "gdi32" Alias "CloseFigure" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function CloseMetaFile Lib "gdi32" Alias "CloseMetaFile" (ByVal hMF As LongPtr) As LongPtr
    Declare PtrSafe Function ColorMatchToTarget Lib "gdi32" Alias "ColorMatchToTarget" (ByVal hdc As LongPtr, ByVal hdc2 As LongPtr, ByVal dw As Long) As Long
    Declare PtrSafe Function CombineRgn Lib "gdi32" Alias "CombineRgn" (ByVal hDestRgn As LongPtr, ByVal hSrcRgn1 As LongPtr, ByVal hSrcRgn2 As LongPtr, ByVal nCombineMode As Long) As Long
    Declare PtrSafe Function CombineTransform Lib "gdi32" Alias "CombineTransform" (lpxformResult As xform, lpxform1 As xform, lpxform2 As xform) As Long
    Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
    Declare PtrSafe Function CopyMetaFile Lib "gdi32" Alias "CopyMetaFileA" (ByVal hMF As LongPtr, ByVal lpFileName As String) As LongPtr
    Declare PtrSafe Function CreateBitmap Lib "gdi32" Alias "CreateBitmap" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As LongPtr
    Declare PtrSafe Function CreateBitmapIndirect Lib "gdi32" Alias "CreateBitmapIndirect" (lpBitmap As BITMAP) As LongPtr
    Declare PtrSafe Function CreateBrushIndirect Lib "gdi32" Alias "CreateBrushIndirect" (lpLogBrush As LOGBRUSH) As LongPtr
    Declare PtrSafe Function CreateColorSpace Lib "gdi32" Alias "CreateColorSpaceA" (lplogcolorspace As LOGCOLORSPACE) As LongPtr
    Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" Alias "CreateCompatibleBitmap" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" Alias "CreateCompatibleDC" (ByVal hdc As LongPtr) As LongPtr
    Declare PtrSafe Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As LongPtr
    Declare PtrSafe Function CreateDIBitmap Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As LongPtr, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As LongPtr
    Declare PtrSafe Function CreateDIBPatternBrush Lib "gdi32" Alias "CreateDIBPatternBrush" (ByVal hPackedDIB As LongPtr, ByVal wUsage As Long) As LongPtr
    Declare PtrSafe Function CreateDIBPatternBrushPt Lib "gdi32" Alias "CreateDIBPatternBrushPt" (lpPackedDIB As Any, ByVal iUsage As Long) As LongPtr
    Declare PtrSafe Function CreateDIBSection Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As LongPtr, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As LongPtr, ByVal handle As LongPtr, ByVal dw As Long) As LongPtr
    Declare PtrSafe Function CreateDiscardableBitmap Lib "gdi32" Alias "CreateDiscardableBitmap" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Declare PtrSafe Function CreateEllipticRgn Lib "gdi32" Alias "CreateEllipticRgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr
    Declare PtrSafe Function CreateEllipticRgnIndirect Lib "gdi32" Alias "CreateEllipticRgnIndirect" (lpRect As RECT) As LongPtr
    Declare PtrSafe Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As LongPtr, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As LongPtr
    Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As LongPtr
    Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As LongPtr
    Declare PtrSafe Function CreateHalftonePalette Lib "gdi32" Alias "CreateHalftonePalette" (ByVal hdc As LongPtr) As LongPtr
    Declare PtrSafe Function CreateHatchBrush Lib "gdi32" Alias "CreateHatchBrush" (ByVal nIndex As Long, ByVal crColor As Long) As LongPtr
    Declare PtrSafe Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As LongPtr
    Declare PtrSafe Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As LongPtr
    Declare PtrSafe Function CreatePalette Lib "gdi32" Alias "CreatePalette" (lpLogPalette As LOGPALETTE) As LongPtr
    Declare PtrSafe Function CreatePatternBrush Lib "gdi32" Alias "CreatePatternBrush" (ByVal hBitmap As LongPtr) As LongPtr
    Declare PtrSafe Function CreatePen Lib "gdi32" Alias "CreatePen" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
    Declare PtrSafe Function CreatePenIndirect Lib "gdi32" Alias "CreatePenIndirect" (lpLogPen As LOGPEN) As LongPtr
    Declare PtrSafe Function CreatePolygonRgn Lib "gdi32" Alias "CreatePolygonRgn" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
    Declare PtrSafe Function CreatePolyPolygonRgn Lib "gdi32" Alias "CreatePolyPolygonRgn" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
    Declare PtrSafe Function CreateRectRgn Lib "gdi32" Alias "CreateRectRgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr
    Declare PtrSafe Function CreateRectRgnIndirect Lib "gdi32" Alias "CreateRectRgnIndirect" (lpRect As RECT) As LongPtr
    Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" Alias "CreateRoundRectRgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As LongPtr
    Declare PtrSafe Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long
    Declare PtrSafe Function CreateSolidBrush Lib "gdi32" Alias "CreateSolidBrush" (ByVal crColor As Long) As LongPtr
    Declare PtrSafe Function DeleteColorSpace Lib "gdi32" Alias "DeleteColorSpace" (ByVal hcolorspace As LongPtr) As Long
    Declare PtrSafe Function DeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function DeleteEnhMetaFile Lib "gdi32" Alias "DeleteEnhMetaFile" (ByVal hemf As LongPtr) As Long
    Declare PtrSafe Function DeleteMetaFile Lib "gdi32" Alias "DeleteMetaFile" (ByVal hMF As LongPtr) As Long
    Declare PtrSafe Function DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As LongPtr) As Long
    Declare PtrSafe Function DescribePixelFormat Lib "gdi32" Alias "DescribePixelFormat" (ByVal hDC As LongPtr, ByVal n As Long, ByVal un As Long, lpPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
    Declare PtrSafe Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long
    Declare PtrSafe Function DPtoLP Lib "gdi32" Alias "DPtoLP" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Declare PtrSafe Function DrawEscape Lib "gdi32" Alias "DrawEscape" (ByVal hdc As LongPtr, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String) As Long
    Declare PtrSafe Function Ellipse Lib "gdi32" Alias "Ellipse" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare PtrSafe Function EndDoc Lib "gdi32" Alias "EndDoc" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function EndPage Lib "gdi32" Alias "EndPage" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function EndPath Lib "gdi32" Alias "EndPath" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function EnumEnhMetaFile Lib "gdi32" Alias "EnumEnhMetaFile" (ByVal hdc As LongPtr, ByVal hemf As LongPtr, ByVal lpEnhMetaFunc As LongPtr, lpData As Any, lpRect As RECT) As Long
    Declare PtrSafe Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As LongPtr, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare PtrSafe Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As LongPtr, lpLogFont As LOGFONT, ByVal lpEnumFontProc As LongPtr, ByVal lParam As LongPtr, ByVal dw As Long) As Long
    Declare PtrSafe Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hDC As LongPtr, ByVal lpsz As String, ByVal lpFontEnumProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare PtrSafe Function EnumICMProfiles Lib "gdi32" Alias "EnumICMProfilesA" (ByVal hdc As LongPtr, ByVal icmEnumProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare PtrSafe Function EnumMetaFile Lib "gdi32" Alias "EnumMetaFile" (ByVal hDC As LongPtr, ByVal hMetafile As LongPtr, ByVal lpMFEnumProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare PtrSafe Function EnumObjects Lib "gdi32" Alias "EnumObjects" (ByVal hDC As LongPtr, ByVal n As Long, ByVal lpGOBJEnumProc As LongPtr, lpVoid As Any) As Long
    Declare PtrSafe Function EqualRgn Lib "gdi32" Alias "EqualRgn" (ByVal hSrcRgn1 As LongPtr, ByVal hSrcRgn2 As LongPtr) As Long
    Declare PtrSafe Function Escape Lib "gdi32" Alias "Escape" (ByVal hdc As LongPtr, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
    Declare PtrSafe Function ExcludeClipRect Lib "gdi32" Alias "ExcludeClipRect" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare PtrSafe Function ExtCreatePen Lib "gdi32" Alias "ExtCreatePen" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As LongPtr
    Declare PtrSafe Function ExtCreateRegion Lib "gdi32" Alias "ExtCreateRegion" (lpXform As XFORM, ByVal nCount As Long, lpRgnData As RGNDATA) As LongPtr
    Declare PtrSafe Function ExtEscape Lib "gdi32" Alias "ExtEscape" (ByVal hdc As LongPtr, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String, ByVal cbOutput As Long, ByVal lpszOutData As String) As Long
    Declare PtrSafe Function ExtFloodFill Lib "gdi32" Alias "ExtFloodFill" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
    Declare PtrSafe Function ExtSelectClipRgn Lib "gdi32" Alias "ExtSelectClipRgn" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal fnMode As Long) As Long
    Declare PtrSafe Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
    Declare PtrSafe Function FillPath Lib "gdi32" Alias "FillPath" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function FillRgn Lib "gdi32" Alias "FillRgn" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr) As Long
    Declare PtrSafe Function FixBrushOrgEx Lib "gdi32" Alias "FixBrushOrgEx" (ByVal hDC As LongPtr, ByVal n1 As Long, ByVal n2 As Long, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function FlattenPath Lib "gdi32" Alias "FlattenPath" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function FloodFill Lib "gdi32" Alias "FloodFill" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
    Declare PtrSafe Function FrameRgn Lib "gdi32" Alias "FrameRgn" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Declare PtrSafe Function GdiComment Lib "gdi32" Alias "GdiComment" (ByVal hdc As LongPtr, ByVal cbSize As Long, lpData As Byte) As Long
    Declare PtrSafe Function GdiFlush Lib "gdi32" Alias "GdiFlush" () As Long
    Declare PtrSafe Function GdiGetBatchLimit Lib "gdi32" Alias "GdiGetBatchLimit" () As Long
    Declare PtrSafe Function GdiSetBatchLimit Lib "gdi32" Alias "GdiSetBatchLimit" (ByVal dwLimit As Long) As Long
    Declare PtrSafe Function GetArcDirection Lib "gdi32" Alias "GetArcDirection" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetAspectRatioFilterEx Lib "gdi32" Alias "GetAspectRatioFilterEx" (ByVal hdc As LongPtr, lpAspectRatio As SIZE) As Long
    Declare PtrSafe Function GetBitmapBits Lib "gdi32" Alias "GetBitmapBits" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long
    Declare PtrSafe Function GetBitmapDimensionEx Lib "gdi32" Alias "GetBitmapDimensionEx" (ByVal hBitmap As LongPtr, lpDimension As SIZE) As Long
    Declare PtrSafe Function GetBkColor Lib "gdi32" Alias "GetBkColor" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetBkMode Lib "gdi32" Alias "GetBkMode" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetBoundsRect Lib "gdi32" Alias "GetBoundsRect" (ByVal hdc As LongPtr, lprcBounds As RECT, ByVal flags As Long) As Long
    Declare PtrSafe Function GetBrushOrgEx Lib "gdi32" Alias "GetBrushOrgEx" (ByVal hDC As LongPtr, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function GetCharABCWidths Lib "gdi32" Alias "GetCharABCWidthsA" (ByVal hdc As LongPtr, ByVal uFirstChar As Long, ByVal uLastChar As Long, lpabc As ABC) As Long
    Declare PtrSafe Function GetCharABCWidthsFloat Lib "gdi32" Alias "GetCharABCWidthsFloatA" (ByVal hdc As LongPtr, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpABCF As ABCFLOAT) As Long
    Declare PtrSafe Function GetCharacterPlacement Lib "gdi32" Alias " GetCharacterPlacementA" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal n1 As Long, ByVal n2 As Long, lpGcpResults As GCP_RESULTS, ByVal dw As Long) As Long
    Declare PtrSafe Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As LongPtr, ByVal wFirstChar As Long, ByVal wLastChar As Long, lpBuffer As Long) As Long
    Declare PtrSafe Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As LongPtr, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long
    Declare PtrSafe Function GetCharWidthFloat Lib "gdi32" Alias "GetCharWidthFloatA" (ByVal hdc As LongPtr, ByVal iFirstChar As Long, ByVal iLastChar As Long, pxBuffer As Double) As Long
    Declare PtrSafe Function GetClipBox Lib "gdi32" Alias "GetClipBox" (ByVal hdc As LongPtr, lpRect As RECT) As Long
    Declare PtrSafe Function GetClipRgn Lib "gdi32" Alias "GetClipRgn" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare PtrSafe Function GetColorAdjustment Lib "gdi32" Alias "GetColorAdjustment" (ByVal hdc As LongPtr, lpca As COLORADJUSTMENT) As Long
    Declare PtrSafe Function GetColorSpace Lib "gdi32" Alias "GetColorSpace" (ByVal hdc As LongPtr) As LongPtr
    Declare PtrSafe Function GetCurrentObject Lib "gdi32" Alias "GetCurrentObject" (ByVal hdc As LongPtr, ByVal uObjectType As Long) As LongPtr
    Declare PtrSafe Function GetCurrentPositionEx Lib "gdi32" Alias "GetCurrentPositionEx" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function GetDCOrgEx Lib "gdi32" Alias "GetDCOrgEx" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function GetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Declare PtrSafe Function GetDeviceGammaRamp Lib "gdi32" Alias "GetDeviceGammaRamp" (ByVal hdc As LongPtr, lpv As Any) As Long
    Declare PtrSafe Function GetDIBColorTable Lib "gdi32" Alias "GetDIBColorTable" (ByVal hDC As LongPtr, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
    Declare PtrSafe Function GetDIBits Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Declare PtrSafe Function GetEnhMetaFile Lib "gdi32" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As LongPtr
    Declare PtrSafe Function GetEnhMetaFileBits Lib "gdi32" Alias "GetEnhMetaFileBits" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpbBuffer As Byte) As Long
    Declare PtrSafe Function GetEnhMetaFileDescription Lib "gdi32" Alias "GetEnhMetaFileDescriptionA" (ByVal hemf As LongPtr, ByVal cchBuffer As Long, ByVal lpszDescription As String) As Long
    Declare PtrSafe Function GetEnhMetaFileHeader Lib "gdi32" Alias "GetEnhMetaFileHeader" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
    Declare PtrSafe Function GetEnhMetaFilePaletteEntries Lib "gdi32" Alias "GetEnhMetaFilePaletteEntries" (ByVal hemf As LongPtr, ByVal cEntries As Long, lppe As PALETTEENTRY) As Long
    Declare PtrSafe Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As LongPtr, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
    Declare PtrSafe Function GetFontLanguageInfo Lib "gdi32" Alias "GetFontLanguageInfo" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hdc As LongPtr, ByVal uChar As Long, ByVal fuFormat As Long, lpgm As GLYPHMETRICS, ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) As Long
    Declare PtrSafe Function GetGraphicsMode Lib "gdi32" Alias "GetGraphicsMode" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetICMProfile Lib "gdi32" Alias "GetICMProfileA" (ByVal hdc As LongPtr, ByVal dw As LongPtr, ByVal lpStr As String) As Long
    Declare PtrSafe Function GetKerningPairs Lib "gdi32" Alias "GetKerningPairsA" (ByVal hdc As LongPtr, ByVal cPairs As Long, lpkrnpair As KERNINGPAIR) As Long
    Declare PtrSafe Function GetLogColorSpace Lib "gdi32" Alias "GetLogColorSpaceA" (ByVal hcolorspace As LongPtr, lplogcolorspace As LOGCOLORSPACE, ByVal dw As Long) As Long
    Declare PtrSafe Function GetMapMode Lib "gdi32" Alias "GetMapMode" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As LongPtr
    Declare PtrSafe Function GetMetaFileBitsEx Lib "gdi32" Alias "GetMetaFileBitsEx" (ByVal hMF As LongPtr, ByVal nSize As Long, lpvData As Any) As Long
    Declare PtrSafe Function GetMetaRgn Lib "gdi32" Alias "GetMetaRgn" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare PtrSafe Function GetMiterLimit Lib "gdi32" Alias "GetMiterLimit" (ByVal hdc As LongPtr, peLimit As Double) As Long
    Declare PtrSafe Function GetNearestColor Lib "gdi32" Alias "GetNearestColor" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
    Declare PtrSafe Function GetNearestPaletteIndex Lib "gdi32" Alias "GetNearestPaletteIndex" (ByVal hPalette As LongPtr, ByVal crColor As Long) As Long
    Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
    Declare PtrSafe Function GetObjectType Lib "gdi32" Alias "GetObjectType" (ByVal hgdiobj As LongPtr) As Long
    Declare PtrSafe Function GetOutlineTextMetrics Lib "gdi32" Alias "GetOutlineTextMetricsA" (ByVal hdc As LongPtr, ByVal cbData As Long, lpotm As OUTLINETEXTMETRIC) As Long
    Declare PtrSafe Function GetPaletteEntries Lib "gdi32" Alias "GetPaletteEntries" (ByVal hPalette As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
    Declare PtrSafe Function GetPath Lib "gdi32" Alias "GetPath" (ByVal hdc As LongPtr, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
    Declare PtrSafe Function GetPixel Lib "gdi32" Alias "GetPixel" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare PtrSafe Function GetPixelFormat Lib "gdi32" Alias "GetPixelFormat" (ByVal hDC As LongPtr) As Long
    Declare PtrSafe Function GetPolyFillMode Lib "gdi32" Alias "GetPolyFillMode" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetProcAddress Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Declare PtrSafe Function GetRasterizerCaps Lib "gdi32" Alias "GetRasterizerCaps" (lpraststat As RASTERIZER_STATUS, ByVal cb As Long) As Long
    Declare PtrSafe Function GetRegionData Lib "gdi32" Alias "GetRegionDataA" (ByVal hRgn As LongPtr, ByVal dwCount As Long, lpRgnData As RgnData) As Long
    Declare PtrSafe Function GetRgnBox Lib "gdi32" Alias "GetRgnBox" (ByVal hRgn As LongPtr, lpRect As RECT) As Long
    Declare PtrSafe Function GetROP2 Lib "gdi32" Alias "GetROP2" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetStockObject Lib "gdi32" Alias "GetStockObject" (ByVal nIndex As Long) As LongPtr
    Declare PtrSafe Function GetStretchBltMode Lib "gdi32" Alias "GetStretchBltMode" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetSystemPaletteEntries Lib "gdi32" Alias "GetSystemPaletteEntries" (ByVal hdc As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
    Declare PtrSafe Function GetSystemPaletteUse Lib "gdi32" Alias "GetSystemPaletteUse" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetTextAlign Lib "gdi32" Alias "GetTextAlign" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetTextCharacterExtra Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetTextCharset Lib "gdi32" Alias "GetTextCharset" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetTextCharsetInfo Lib "gdi32" Alias "GetTextCharsetInfo" (ByVal hdc As LongPtr, lpSig As FONTSIGNATURE, ByVal dwFlags As Long) As Long
    Declare PtrSafe Function GetTextColor Lib "gdi32" Alias "GetTextColor" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function GetTextExtentExPoint Lib "gdi32" Alias "GetTextExtentExPointA" (ByVal hdc As LongPtr, ByVal lpszStr As String, ByVal cchString As Long, ByVal nMaxExtent As Long, lpnFit As Long, alpDx As Long, lpSize As SIZE) As Long
    Declare PtrSafe Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As LongPtr, ByVal lpszString As String, ByVal cbString As Long, lpSize As SIZE) As Long
    Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
    Declare PtrSafe Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As LongPtr, ByVal nCount As Long, ByVal lpFacename As String) As Long
    Declare PtrSafe Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As LongPtr, lpMetrics As TEXTMETRIC) As Long
    Declare PtrSafe Function GetViewportExtEx Lib "gdi32" Alias "GetViewportExtEx" (ByVal hdc As LongPtr, lpSize As SIZE) As Long
    Declare PtrSafe Function GetViewportOrgEx Lib "gdi32" Alias "GetViewportOrgEx" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function GetWindowExtEx Lib "gdi32" Alias "GetWindowExtEx" (ByVal hdc As LongPtr, lpSize As SIZE) As Long
    Declare PtrSafe Function GetWindowOrgEx Lib "gdi32" Alias "GetWindowOrgEx" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function GetWinMetaFileBits Lib "gdi32" Alias "GetWinMetaFileBits" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal fnMapMode As Long, ByVal hdcRef As LongPtr) As Long
    Declare PtrSafe Function GetWorldTransform Lib "gdi32" Alias "GetWorldTransform" (ByVal hdc As LongPtr, lpXform As xform) As Long
    Declare PtrSafe Function IntersectClipRect Lib "gdi32" Alias "IntersectClipRect" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare PtrSafe Function InvertRgn Lib "gdi32" Alias "InvertRgn" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare PtrSafe Function LineDDA Lib "gdi32" Alias "LineDDA" (ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal lpLineDDAProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare PtrSafe Function LineTo Lib "gdi32" Alias "LineTo" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare PtrSafe Function LPtoDP Lib "gdi32" Alias "LPtoDP" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Declare PtrSafe Function MaskBlt Lib "gdi32" Alias "MaskBlt" (ByVal hdcDest As LongPtr, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As LongPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As LongPtr, ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As Long) As Long
    Declare PtrSafe Function ModifyWorldTransform Lib "gdi32" Alias "ModifyWorldTransform" (ByVal hdc As LongPtr, lpXform As xform, ByVal iMode As Long) As Long
    Declare PtrSafe Function MoveToEx Lib "gdi32" Alias "MoveToEx" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function OffsetClipRgn Lib "gdi32" Alias "OffsetClipRgn" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare PtrSafe Function OffsetRgn Lib "gdi32" Alias "OffsetRgn" (ByVal hRgn As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare PtrSafe Function OffsetViewportOrgEx Lib "gdi32" Alias "OffsetViewportOrgEx" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function OffsetWindowOrgEx Lib "gdi32" Alias "OffsetWindowOrgEx" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function PaintRgn Lib "gdi32" Alias "PaintRgn" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare PtrSafe Function PatBlt Lib "gdi32" Alias "PatBlt" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
    Declare PtrSafe Function PathToRegion Lib "gdi32" Alias "PathToRegion" (ByVal hdc As LongPtr) As LongPtr
    Declare PtrSafe Function Pie Lib "gdi32" Alias "Pie" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
    Declare PtrSafe Function PlayEnhMetaFile Lib "gdi32" Alias "PlayEnhMetaFile" (ByVal hdc As LongPtr, ByVal hemf As LongPtr, lpRect As RECT) As Long
    Declare PtrSafe Function PlayEnhMetaFileRecord Lib "gdi32" Alias "PlayEnhMetaFileRecord" (ByVal hdc As LongPtr, lpHandletable As HANDLETABLE, lpEnhMetaRecord As ENHMETARECORD, ByVal nHandles As Long) As Long
    Declare PtrSafe Function PlayMetaFile Lib "gdi32" Alias "PlayMetaFile" (ByVal hdc As LongPtr, ByVal hMF As LongPtr) As Long
    Declare PtrSafe Function PlayMetaFileRecord Lib "gdi32" Alias "PlayMetaFileRecord" (ByVal hdc As LongPtr, lpHandletable As HANDLETABLE, lpMetaRecord As METARECORD, ByVal nHandles As Long) As Long
    Declare PtrSafe Function PlgBlt Lib "gdi32" Alias "PlgBlt" (ByVal hdcDest As LongPtr, lpPoint As POINTAPI, ByVal hdcSrc As LongPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As LongPtr, ByVal xMask As Long, ByVal yMask As Long) As Long
    Declare PtrSafe Function PolyBezier Lib "gdi32" Alias "PolyBezier" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cPoints As Long) As Long
    Declare PtrSafe Function PolyBezierTo Lib "gdi32" Alias "PolyBezierTo" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cCount As Long) As Long
    Declare PtrSafe Function PolyDraw Lib "gdi32" Alias "PolyDraw" (ByVal hdc As LongPtr, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
    Declare PtrSafe Function Polygon Lib "gdi32" Alias "Polygon" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Declare PtrSafe Function Polyline Lib "gdi32" Alias "Polyline" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Declare PtrSafe Function PolylineTo Lib "gdi32" Alias "PolylineTo" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cCount As Long) As Long
    Declare PtrSafe Function PolyPolygon Lib "gdi32" Alias "PolyPolygon" (ByVal hdc As LongPtr, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
    Declare PtrSafe Function PolyPolyline Lib "gdi32" Alias "PolyPolyline" (ByVal hdc As LongPtr, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
    Declare PtrSafe Function PolyTextOut Lib "gdi32" Alias "PolyTextOutA" (ByVal hdc As LongPtr, pptxt As POLYTEXT, ByVal cStrings As Long) As Long
    Declare PtrSafe Function PtInRegion Lib "gdi32" Alias "PtInRegion" (ByVal hRgn As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare PtrSafe Function PtVisible Lib "gdi32" Alias "PtVisible" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare PtrSafe Function RealizePalette Lib "gdi32" Alias "RealizePalette" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function Rectangle Lib "gdi32" Alias "Rectangle" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare PtrSafe Function RectInRegion Lib "gdi32" Alias "RectInRegion" (ByVal hRgn As LongPtr, lpRect As RECT) As Long
    Declare PtrSafe Function RectVisible Lib "gdi32" Alias "RectVisible" (ByVal hdc As LongPtr, lpRect As RECT) As Long
    Declare PtrSafe Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
    Declare PtrSafe Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As LongPtr, lpInitData As DEVMODE) As LongPtr
    Declare PtrSafe Function ResizePalette Lib "gdi32" Alias "ResizePalette" (ByVal hPalette As LongPtr, ByVal nNumEntries As Long) As Long
    Declare PtrSafe Function RestoreDC Lib "gdi32" Alias "RestoreDC" (ByVal hdc As LongPtr, ByVal nSavedDC As Long) As Long
    Declare PtrSafe Function RoundRect Lib "gdi32" Alias "RoundRect" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    Declare PtrSafe Function SaveDC Lib "gdi32" Alias "SaveDC" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function ScaleViewportExtEx Lib "gdi32" Alias "ScaleViewportExtEx" (ByVal hdc As LongPtr, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As SIZE) As Long
    Declare PtrSafe Function ScaleWindowExtEx Lib "gdi32" Alias "ScaleWindowExtEx" (ByVal hdc As LongPtr, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As SIZE) As Long
    Declare PtrSafe Function SelectClipPath Lib "gdi32" Alias "SelectClipPath" (ByVal hdc As LongPtr, ByVal iMode As Long) As Long
    Declare PtrSafe Function SelectClipRgn Lib "gdi32" Alias "SelectClipRgn" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare PtrSafe Function SelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Declare PtrSafe Function SelectPalette Lib "gdi32" Alias "SelectPalette" (ByVal hdc As LongPtr, ByVal hPalette As LongPtr, ByVal bForceBackground As Long) As LongPtr
    Declare PtrSafe Function SetAbortProc Lib "gdi32" Alias "SetAbortProc" (ByVal hDC As LongPtr, ByVal lpAbortProc As LongPtr) As Long
    Declare PtrSafe Function SetArcDirection Lib "gdi32" Alias "SetArcDirection" (ByVal hdc As LongPtr, ByVal ArcDirection As Long) As Long
    Declare PtrSafe Function SetBitmapBits Lib "gdi32" Alias "SetBitmapBits" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long
    Declare PtrSafe Function SetBitmapDimensionEx Lib "gdi32" Alias "SetBitmapDimensionEx" (ByVal hbm As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
    Declare PtrSafe Function SetBkColor Lib "gdi32" Alias "SetBkColor" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
    Declare PtrSafe Function SetBkMode Lib "gdi32" Alias "SetBkMode" (ByVal hdc As LongPtr, ByVal nBkMode As Long) As Long
    Declare PtrSafe Function SetBoundsRect Lib "gdi32" Alias "SetBoundsRect" (ByVal hdc As LongPtr, lprcBounds As RECT, ByVal flags As Long) As Long
    Declare PtrSafe Function SetBrushOrgEx Lib "gdi32" Alias "SetBrushOrgEx" (ByVal hdc As LongPtr, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
    Declare PtrSafe Function SetColorAdjustment Lib "gdi32" Alias "SetColorAdjustment" (ByVal hdc As LongPtr, lpca As COLORADJUSTMENT) As Long
    Declare PtrSafe Function SetColorSpace Lib "gdi32" Alias "SetColorSpace" (ByVal hdc As LongPtr, ByVal hcolorspace As LongPtr) As LongPtr
    Declare PtrSafe Function SetDeviceGammaRamp Lib "gdi32" Alias "SetDeviceGammaRamp" (ByVal hdc As LongPtr, lpv As Any) As Long
    Declare PtrSafe Function SetDIBColorTable Lib "gdi32" Alias "SetDIBColorTable" (ByVal hDC As LongPtr, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
    Declare PtrSafe Function SetDIBits Lib "gdi32" Alias "SetDIBits" (ByVal hdc As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Declare PtrSafe Function SetDIBitsToDevice Lib "gdi32" Alias "SetDIBitsToDevice" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
    Declare PtrSafe Function SetEnhMetaFileBits Lib "gdi32" Alias "SetEnhMetaFileBits" (ByVal cbBuffer As Long, lpData As Byte) As LongPtr
    Declare PtrSafe Function SetGraphicsMode Lib "gdi32" Alias "SetGraphicsMode" (ByVal hdc As LongPtr, ByVal iMode As Long) As Long
    Declare PtrSafe Function SetICMMode Lib "gdi32" Alias "SetICMMode" (ByVal hdc As LongPtr, ByVal n As Long) As Long
    Declare PtrSafe Function SetICMProfile Lib "gdi32" Alias "SetICMProfileA" (ByVal hdc As LongPtr, ByVal lpStr As String) As Long
    Declare PtrSafe Function SetMapMode Lib "gdi32" Alias "SetMapMode" (ByVal hdc As LongPtr, ByVal nMapMode As Long) As Long
    Declare PtrSafe Function SetMapperFlags Lib "gdi32" Alias "SetMapperFlags" (ByVal hdc As LongPtr, ByVal dwFlag As Long) As Long
    Declare PtrSafe Function SetMetaFileBitsEx Lib "gdi32" Alias "SetMetaFileBitsEx" (ByVal nSize As Long, lpData As Byte) As LongPtr
    Declare PtrSafe Function SetMetaRgn Lib "gdi32" Alias "SetMetaRgn" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function SetMiterLimit Lib "gdi32" Alias "SetMiterLimit" (ByVal hdc As LongPtr, ByVal eNewLimit As Double, peOldLimit As Double) As Long
    Declare PtrSafe Function SetPaletteEntries Lib "gdi32" Alias "SetPaletteEntries" (ByVal hPalette As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
    Declare PtrSafe Function SetPixel Lib "gdi32" Alias "SetPixel" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
    Declare PtrSafe Function SetPixelFormat Lib "gdi32" Alias "SetPixelFormat" (ByVal hDC As LongPtr, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
    Declare PtrSafe Function SetPixelV Lib "gdi32" Alias "SetPixelV" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
    Declare PtrSafe Function SetPolyFillMode Lib "gdi32" Alias "SetPolyFillMode" (ByVal hdc As LongPtr, ByVal nPolyFillMode As Long) As Long
    Declare PtrSafe Function SetRectRgn Lib "gdi32" Alias "SetRectRgn" (ByVal hRgn As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare PtrSafe Function SetROP2 Lib "gdi32" Alias "SetROP2" (ByVal hdc As LongPtr, ByVal nDrawMode As Long) As Long
    Declare PtrSafe Function SetStretchBltMode Lib "gdi32" Alias "SetStretchBltMode" (ByVal hdc As LongPtr, ByVal nStretchMode As Long) As Long
    Declare PtrSafe Function SetSystemPaletteUse Lib "gdi32" Alias "SetSystemPaletteUse" (ByVal hdc As LongPtr, ByVal wUsage As Long) As Long
    Declare PtrSafe Function SetTextAlign Lib "gdi32" Alias "SetTextAlign" (ByVal hdc As LongPtr, ByVal wFlags As Long) As Long
    Declare PtrSafe Function SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtraA" (ByVal hdc As LongPtr, ByVal nCharExtra As Long) As Long
    Declare PtrSafe Function SetTextColor Lib "gdi32" Alias "SetTextColor" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
    Declare PtrSafe Function SetTextJustification Lib "gdi32" Alias "SetTextJustification" (ByVal hdc As LongPtr, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
    Declare PtrSafe Function SetViewportExtEx Lib "gdi32" Alias "SetViewportExtEx" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
    Declare PtrSafe Function SetViewportOrgEx Lib "gdi32" Alias "SetViewportOrgEx" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function SetWindowExtEx Lib "gdi32" Alias "SetWindowExtEx" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
    Declare PtrSafe Function SetWindowOrgEx Lib "gdi32" Alias "SetWindowOrgEx" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
    Declare PtrSafe Function SetWinMetaFileBits Lib "gdi32" Alias "SetWinMetaFileBits" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As LongPtr, lpmfp As METAFILEPICT) As LongPtr
    Declare PtrSafe Function SetWorldTransform Lib "gdi32" Alias "SetWorldTransform" (ByVal hdc As LongPtr, lpXform As xform) As Long
    Declare PtrSafe Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As LongPtr, lpdi As DOCINFO) As Long
    Declare PtrSafe Function StartPage Lib "gdi32" Alias "StartPage" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function StretchBlt Lib "gdi32" Alias "StretchBlt" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Declare PtrSafe Function StretchDIBits Lib "gdi32" Alias "StretchDIBits" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
    Declare PtrSafe Function StrokeAndFillPath Lib "gdi32" Alias "StrokeAndFillPath" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function StrokePath Lib "gdi32" Alias "StrokePath" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function SwapBuffers Lib "gdi32" Alias "SwapBuffers" (ByVal hDC As LongPtr) As Long
    Declare PtrSafe Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
    Declare PtrSafe Function TranslateCharsetInfo Lib "gdi32" Alias "TranslateCharsetInfo" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long
    Declare PtrSafe Function UnrealizeObject Lib "gdi32" Alias "UnrealizeObject" (ByVal hObject As LongPtr) As Long
    Declare PtrSafe Function UpdateColors Lib "gdi32" Alias "UpdateColors" (ByVal hdc As LongPtr) As Long
    Declare PtrSafe Function WidenPath Lib "gdi32" Alias "WidenPath" (ByVal hdc As LongPtr) As Long

#Else

    Declare Function AbortDoc Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function AbortPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
    Declare Function AngleArc Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Double, ByVal eSweepAngle As Double) As Long
    Declare Function AnimatePalette Lib "gdi32" Alias "AnimatePaletteA" (ByVal hPalette As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteColors As PALETTEENTRY) As Long
    Declare Function Arc Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
    Declare Function ArcTo Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
    Declare Function BeginPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Declare Function CancelDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function CheckColorsInGamut Lib "gdi32" (ByVal hdc As LongPtr, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long
    Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As LongPtr, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
    Declare Function Chord Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
    Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
    Declare Function CloseFigure Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As LongPtr) As LongPtr
    Declare Function ColorMatchToTarget Lib "gdi32" (ByVal hdc As LongPtr, ByVal hdc2 As LongPtr, ByVal dw As Long) As Long
    Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As LongPtr, ByVal hSrcRgn1 As LongPtr, ByVal hSrcRgn2 As LongPtr, ByVal nCombineMode As Long) As Long
    Declare Function CombineTransform Lib "gdi32" (lpxformResult As xform, lpxform1 As xform, lpxform2 As xform) As Long
    Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
    Declare Function CopyMetaFile Lib "gdi32" Alias "CopyMetaFileA" (ByVal hMF As LongPtr, ByVal lpFileName As String) As LongPtr
    Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As LongPtr
    Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As LongPtr
    Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As LongPtr
    Declare Function CreateColorSpace Lib "gdi32" Alias "CreateColorSpaceA" (lplogcolorspace As LOGCOLORSPACE) As LongPtr
    Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
    Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As LongPtr
    Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As LongPtr, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As LongPtr
    Declare Function CreateDIBPatternBrush Lib "gdi32" (ByVal hPackedDIB As LongPtr, ByVal wUsage As Long) As LongPtr
    Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As LongPtr
    Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As LongPtr, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As LongPtr, ByVal handle As LongPtr, ByVal dw As Long) As LongPtr
    Declare Function CreateDiscardableBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr
    Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As LongPtr
    Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As LongPtr, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As LongPtr
    Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As LongPtr
    Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As LongPtr
    Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
    Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As LongPtr
    Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As LongPtr
    Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As LongPtr
    Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As LongPtr
    Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As LongPtr) As LongPtr
    Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
    Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As LongPtr
    Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
    Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
    Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr
    Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As LongPtr
    Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As LongPtr
    Declare Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long
    Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
    Declare Function DeleteColorSpace Lib "gdi32" (ByVal hcolorspace As LongPtr) As Long
    Declare Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As LongPtr) As Long
    Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As LongPtr) As Long
    Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Declare Function DescribePixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByVal n As Long, ByVal un As Long, lpPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
    Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long
    Declare Function DPtoLP Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Declare Function DrawEscape Lib "gdi32" (ByVal hdc As LongPtr, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String) As Long
    Declare Function Ellipse Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function EndDoc Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function EndPage Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function EndPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function EnumEnhMetaFile Lib "gdi32" (ByVal hdc As LongPtr, ByVal hemf As LongPtr, ByVal lpEnhMetaFunc As LongPtr, lpData As Any, lpRect As RECT) As Long
    Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As LongPtr, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As LongPtr, lpLogFont As LOGFONT, ByVal lpEnumFontProc As LongPtr, ByVal lParam As LongPtr, ByVal dw As Long) As Long
    Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal lpFontEnumProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare Function EnumICMProfiles Lib "gdi32" Alias "EnumICMProfilesA" (ByVal hdc As LongPtr, ByVal icmEnumProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare Function EnumMetaFile Lib "gdi32" (ByVal hdc As LongPtr, ByVal hMetafile As LongPtr, ByVal lpMFEnumProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare Function EnumObjects Lib "gdi32" (ByVal hdc As LongPtr, ByVal n As Long, ByVal lpGOBJEnumProc As LongPtr, lpVoid As Any) As Long
    Declare Function EqualRgn Lib "gdi32" (ByVal hSrcRgn1 As LongPtr, ByVal hSrcRgn2 As LongPtr) As Long
    Declare Function Escape Lib "gdi32" (ByVal hdc As LongPtr, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
    Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As LongPtr
    Declare Function ExtCreateRegion Lib "gdi32" (lpXform As xform, ByVal nCount As Long, lpRgnData As RgnData) As LongPtr
    Declare Function ExtEscape Lib "gdi32" (ByVal hdc As LongPtr, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String, ByVal cbOutput As Long, ByVal lpszOutData As String) As Long
    Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
    Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal fnMode As Long) As Long
    Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
    Declare Function FillPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function FillRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr) As Long
    Declare Function FixBrushOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal n1 As Long, ByVal n2 As Long, lpPoint As POINTAPI) As Long
    Declare Function FlattenPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function FloodFill Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
    Declare Function FrameRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Declare Function GdiComment Lib "gdi32" (ByVal hdc As LongPtr, ByVal cbSize As Long, lpData As Byte) As Long
    Declare Function GdiFlush Lib "gdi32" () As Long
    Declare Function GdiGetBatchLimit Lib "gdi32" () As Long
    Declare Function GdiSetBatchLimit Lib "gdi32" (ByVal dwLimit As Long) As Long
    Declare Function GetArcDirection Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetAspectRatioFilterEx Lib "gdi32" (ByVal hdc As LongPtr, lpAspectRatio As size) As Long
    Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long
    Declare Function GetBitmapDimensionEx Lib "gdi32" (ByVal hBitmap As LongPtr, lpDimension As size) As Long
    Declare Function GetBkColor Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetBkMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetBoundsRect Lib "gdi32" (ByVal hdc As LongPtr, lprcBounds As RECT, ByVal Flags As Long) As Long
    Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
    Declare Function GetCharABCWidths Lib "gdi32" Alias "GetCharABCWidthsA" (ByVal hdc As LongPtr, ByVal uFirstChar As Long, ByVal uLastChar As Long, lpabc As ABC) As Long
    Declare Function GetCharABCWidthsFloat Lib "gdi32" Alias "GetCharABCWidthsFloatA" (ByVal hdc As LongPtr, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpABCF As ABCFLOAT) As Long
    Declare Function GetCharacterPlacement Lib "gdi32" Alias " GetCharacterPlacementA" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal n1 As Long, ByVal n2 As Long, lpGcpResults As GCP_RESULTS, ByVal dw As Long) As Long
    Declare Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As LongPtr, ByVal wFirstChar As Long, ByVal wLastChar As Long, lpBuffer As Long) As Long
    Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As LongPtr, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long
    Declare Function GetCharWidthFloat Lib "gdi32" Alias "GetCharWidthFloatA" (ByVal hdc As LongPtr, ByVal iFirstChar As Long, ByVal iLastChar As Long, pxBuffer As Double) As Long
    Declare Function GetClipBox Lib "gdi32" (ByVal hdc As LongPtr, lpRect As RECT) As Long
    Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare Function GetColorAdjustment Lib "gdi32" (ByVal hdc As LongPtr, lpca As ColorAdjustment) As Long
    Declare Function GetColorSpace Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
    Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal uObjectType As Long) As LongPtr
    Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
    Declare Function GetDCOrgEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
    Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Declare Function GetDeviceGammaRamp Lib "gdi32" (ByVal hdc As LongPtr, lpv As Any) As Long
    Declare Function GetDIBColorTable Lib "gdi32" (ByVal hdc As LongPtr, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
    Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Declare Function GetEnhMetaFile Lib "gdi32" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As LongPtr
    Declare Function GetEnhMetaFileBits Lib "gdi32" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpbBuffer As Byte) As Long
    Declare Function GetEnhMetaFileDescription Lib "gdi32" Alias "GetEnhMetaFileDescriptionA" (ByVal hemf As LongPtr, ByVal cchBuffer As Long, ByVal lpszDescription As String) As Long
    Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
    Declare Function GetEnhMetaFilePaletteEntries Lib "gdi32" (ByVal hemf As LongPtr, ByVal cEntries As Long, lppe As PALETTEENTRY) As Long
    Declare Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As LongPtr, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
    Declare Function GetFontLanguageInfo Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hdc As LongPtr, ByVal uChar As Long, ByVal fuFormat As Long, lpgm As GLYPHMETRICS, ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) As Long
    Declare Function GetGraphicsMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetICMProfile Lib "gdi32" Alias "GetICMProfileA" (ByVal hdc As LongPtr, ByVal dw As LongPtr, ByVal lpStr As String) As Long
    Declare Function GetKerningPairs Lib "gdi32" Alias "GetKerningPairsA" (ByVal hdc As LongPtr, ByVal cPairs As Long, lpkrnpair As KERNINGPAIR) As Long
    Declare Function GetLogColorSpace Lib "gdi32" Alias "GetLogColorSpaceA" (ByVal hcolorspace As LongPtr, lplogcolorspace As LOGCOLORSPACE, ByVal dw As Long) As Long
    Declare Function GetMapMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As LongPtr
    Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As LongPtr, ByVal nSize As Long, lpvData As Any) As Long
    Declare Function GetMetaRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare Function GetMiterLimit Lib "gdi32" (ByVal hdc As LongPtr, peLimit As Double) As Long
    Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
    Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As LongPtr, ByVal crColor As Long) As Long
    Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
    Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As LongPtr) As Long
    Declare Function GetOutlineTextMetrics Lib "gdi32" Alias "GetOutlineTextMetricsA" (ByVal hdc As LongPtr, ByVal cbData As Long, lpotm As OUTLINETEXTMETRIC) As Long
    Declare Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
    Declare Function GetPath Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
    Declare Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetPolyFillMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Declare Function GetRasterizerCaps Lib "gdi32" (lpraststat As RASTERIZER_STATUS, ByVal cb As Long) As Long
    Declare Function GetRegionData Lib "gdi32" Alias "GetRegionDataA" (ByVal hRgn As LongPtr, ByVal dwCount As Long, lpRgnData As RgnData) As Long
    Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As LongPtr, lpRect As RECT) As Long
    Declare Function GetROP2 Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As LongPtr
    Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
    Declare Function GetSystemPaletteUse Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetTextAlign Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetTextCharacterExtra Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetTextCharset Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetTextCharsetInfo Lib "gdi32" (ByVal hdc As LongPtr, lpSig As FONTSIGNATURE, ByVal dwFlags As Long) As Long
    Declare Function GetTextColor Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function GetTextExtentExPoint Lib "gdi32" Alias "GetTextExtentExPointA" (ByVal hdc As LongPtr, ByVal lpszStr As String, ByVal cchString As Long, ByVal nMaxExtent As Long, lpnFit As Long, alpDx As Long, lpSize As size) As Long
    Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As LongPtr, ByVal lpszString As String, ByVal cbString As Long, lpSize As size) As Long
    Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long
    Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As LongPtr, ByVal nCount As Long, ByVal lpFacename As String) As Long
    Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As LongPtr, lpMetrics As TEXTMETRIC) As Long
    Declare Function GetViewportExtEx Lib "gdi32" (ByVal hdc As LongPtr, lpSize As size) As Long
    Declare Function GetViewportOrgEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
    Declare Function GetWindowExtEx Lib "gdi32" (ByVal hdc As LongPtr, lpSize As size) As Long
    Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
    Declare Function GetWinMetaFileBits Lib "gdi32" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal fnMapMode As Long, ByVal hdcRef As LongPtr) As Long
    Declare Function GetWorldTransform Lib "gdi32" (ByVal hdc As LongPtr, lpXform As xform) As Long
    Declare Function IntersectClipRect Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function InvertRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare Function LineDDA Lib "gdi32" (ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal lpLineDDAProc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare Function LineTo Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare Function LPtoDP Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Declare Function MaskBlt Lib "gdi32" (ByVal hDCDest As LongPtr, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As LongPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As LongPtr, ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As Long) As Long
    Declare Function ModifyWorldTransform Lib "gdi32" (ByVal hdc As LongPtr, lpXform As xform, ByVal iMode As Long) As Long
    Declare Function MoveToEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
    Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare Function OffsetViewportOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
    Declare Function OffsetWindowOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
    Declare Function PaintRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare Function PatBlt Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
    Declare Function PathToRegion Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
    Declare Function Pie Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
    Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As LongPtr, ByVal hemf As LongPtr, lpRect As RECT) As Long
    Declare Function PlayEnhMetaFileRecord Lib "gdi32" (ByVal hdc As LongPtr, lpHandletable As HANDLETABLE, lpEnhMetaRecord As ENHMETARECORD, ByVal nHandles As Long) As Long
    Declare Function PlayMetaFile Lib "gdi32" (ByVal hdc As LongPtr, ByVal hMF As LongPtr) As Long
    Declare Function PlayMetaFileRecord Lib "gdi32" (ByVal hdc As LongPtr, lpHandletable As HANDLETABLE, lpMetaRecord As METARECORD, ByVal nHandles As Long) As Long
    Declare Function PlgBlt Lib "gdi32" (ByVal hDCDest As LongPtr, lpPoint As POINTAPI, ByVal hdcSrc As LongPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As LongPtr, ByVal xMask As Long, ByVal yMask As Long) As Long
    Declare Function PolyBezier Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cPoints As Long) As Long
    Declare Function PolyBezierTo Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cCount As Long) As Long
    Declare Function PolyDraw Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
    Declare Function Polygon Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Declare Function Polyline Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Declare Function PolylineTo Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cCount As Long) As Long
    Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
    Declare Function PolyPolyline Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
    Declare Function PolyTextOut Lib "gdi32" Alias "PolyTextOutA" (ByVal hdc As LongPtr, pptxt As POLYTEXT, ByVal cStrings As Long) As Long
    Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare Function PtVisible Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
    Declare Function RealizePalette Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function Rectangle Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As LongPtr, lpRect As RECT) As Long
    Declare Function RectVisible Lib "gdi32" (ByVal hdc As LongPtr, lpRect As RECT) As Long
    Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
    Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As LongPtr, lpInitData As DEVMODE) As LongPtr
    Declare Function ResizePalette Lib "gdi32" (ByVal hPalette As LongPtr, ByVal nNumEntries As Long) As Long
    Declare Function RestoreDC Lib "gdi32" (ByVal hdc As LongPtr, ByVal nSavedDC As Long) As Long
    Declare Function RoundRect Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    Declare Function SaveDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function ScaleViewportExtEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As size) As Long
    Declare Function ScaleWindowExtEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As size) As Long
    Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As LongPtr, ByVal iMode As Long) As Long
    Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
    Declare Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Declare Function SelectPalette Lib "gdi32" (ByVal hdc As LongPtr, ByVal hPalette As LongPtr, ByVal bForceBackground As Long) As LongPtr
    Declare Function SetAbortProc Lib "gdi32" (ByVal hdc As LongPtr, ByVal lpAbortProc As LongPtr) As Long
    Declare Function SetArcDirection Lib "gdi32" (ByVal hdc As LongPtr, ByVal ArcDirection As Long) As Long
    Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long
    Declare Function SetBitmapDimensionEx Lib "gdi32" (ByVal hbm As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As size) As Long
    Declare Function SetBkColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
    Declare Function SetBkMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nBkMode As Long) As Long
    Declare Function SetBoundsRect Lib "gdi32" (ByVal hdc As LongPtr, lprcBounds As RECT, ByVal Flags As Long) As Long
    Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
    Declare Function SetColorAdjustment Lib "gdi32" (ByVal hdc As LongPtr, lpca As ColorAdjustment) As Long
    Declare Function SetColorSpace Lib "gdi32" (ByVal hdc As LongPtr, ByVal hcolorspace As LongPtr) As LongPtr
    Declare Function SetDeviceGammaRamp Lib "gdi32" (ByVal hdc As LongPtr, lpv As Any) As Long
    Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As LongPtr, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
    Declare Function SetDIBits Lib "gdi32" (ByVal hdc As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
    Declare Function SetEnhMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpData As Byte) As LongPtr
    Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal iMode As Long) As Long
    Declare Function SetICMMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal n As Long) As Long
    Declare Function SetICMProfile Lib "gdi32" Alias "SetICMProfileA" (ByVal hdc As LongPtr, ByVal lpStr As String) As Long
    Declare Function SetMapMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nMapMode As Long) As Long
    Declare Function SetMapperFlags Lib "gdi32" (ByVal hdc As LongPtr, ByVal dwFlag As Long) As Long
    Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As LongPtr
    Declare Function SetMetaRgn Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function SetMiterLimit Lib "gdi32" (ByVal hdc As LongPtr, ByVal eNewLimit As Double, peOldLimit As Double) As Long
    Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
    Declare Function SetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
    Declare Function SetPixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
    Declare Function SetPixelV Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
    Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nPolyFillMode As Long) As Long
    Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function SetROP2 Lib "gdi32" (ByVal hdc As LongPtr, ByVal nDrawMode As Long) As Long
    Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nStretchMode As Long) As Long
    Declare Function SetSystemPaletteUse Lib "gdi32" (ByVal hdc As LongPtr, ByVal wUsage As Long) As Long
    Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As LongPtr, ByVal wFlags As Long) As Long
    Declare Function SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtraA" (ByVal hdc As LongPtr, ByVal nCharExtra As Long) As Long
    Declare Function SetTextColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
    Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As LongPtr, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
    Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As size) As Long
    Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
    Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As size) As Long
    Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
    Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As LongPtr, lpmfp As METAFILEPICT) As LongPtr
    Declare Function SetWorldTransform Lib "gdi32" (ByVal hdc As LongPtr, lpXform As xform) As Long
    Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As LongPtr, lpdi As DOCINFO) As Long
    Declare Function StartPage Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function StretchBlt Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
    Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function StrokePath Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function SwapBuffers Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
    Declare Function TranslateCharsetInfo Lib "gdi32" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long
    Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Declare Function UpdateColors Lib "gdi32" (ByVal hdc As LongPtr) As Long
    Declare Function WidenPath Lib "gdi32" (ByVal hdc As LongPtr) As Long

#End If


