
                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||        IMAGES - EXIF PROPERTIES       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
                                                                                                                                                                                            ' _
    AUTHOR:   Daniel Pineault   (64-bit compatibility added by K. Willock)                                                                                                                   ' _
    PURPOSE:  GDI+ routines to read and write EXIF properties to JPG image files.                                                                                                           ' _
    LICENSE:  Attribution-ShareAlike 4.0 International                                                                                                                                      ' _
              (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/                                                                                                              ' _
    VERSION:  1.0        22/03/2022                                                                                                                                                         ' _
                                                                                                                                                                                            ' _
    NOTES:    https://www.devhut.net/getting-all-of-an-images-properties-using-the-gdi-api/                                                                                                 ' _
                                                                                                                                                                                            ' _
    USAGE:                                                                                                                                                                                  ' _
              - SetImageProperty("C:\Temp\IMG01 - Copy.jpg", EquipModel , "Huawei")                                                                                                         ' _
              - SetImageProperty("C:\Temp\IMG01 - Copy.jpg", ExifISOSpeed , 250)                                                                                                            ' _
              - SetImageProperty("C:\Temp\IMG01 - Copy.jpg", ImageDescription , "View from the house")                                                                                      ' _
                                                                                                                                                                                            ' _
              - ImageProps = GetImageProperties("C:\Temp\IMG_20210508_170154.jpg")
    
    Option Explicit
     
    'API Declarations, ENUMS, TYPES, Global Variables, ...
    '-------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------
    Private Const GdiplusVersion  As Long = 1
     
    Private Const ImageCodecBMP = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
    Private Const ImageCodecGIF = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
    Private Const ImageCodecJPG = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
    Private Const ImageCodecPNG = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
    Private Const ImageCodecTIF = "{557CF405-1A04-11D3-9A73-0000F81EF32E}"
    Private Const EncoderQuality = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
    Private Const EncoderCompression = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
    Private Const TiffCompressionNone = 6
    Private Const EncoderParameterValueTypeLong = 4
     
    Private Type GUID
       Data1 As Long
       Data2 As Integer
       Data3 As Integer
       Data4(0 To 7) As Byte
    End Type
     
    Private Type EncoderParameter
        GUID            As GUID
        NumberOfValues  As Long
        Type            As Long
        value           As Long
    End Type
     
    Private Type EncoderParameters
        Count           As Long
        Parameter(15)   As EncoderParameter
    End Type
     
    Private Type PropertyItem
        id                        As Long    'PropertyTagId
        Length                    As Long
        Type                      As Integer 'PropertyTagType
        value                     As LongPtr
    End Type
     
    'GDI+ Status Constants
    Private Enum Status
        'https://docs.microsoft.com/en-us/windows/win32/api/gdiplustypes/ne-gdiplustypes-status
        OK = 0
        GenericError = 1
        InvalidParameter = 2
        OutOfMemory = 3
        ObjectBusy = 4
        InsufficientBuffer = 5
        NotImplemented = 6
        Win32Error = 7
        WrongState = 8
        Aborted = 9
        FileNotFound = 10
        ValueOverflow = 11
        AccessDenied = 12
        UnknownImageFormat = 13
        FontFamilyNotFound = 14
        FontStyleNotFound = 15
        NotTrueTypeFont = 16
        UnsupportedGdiplusVersion = 17
        GdiplusNotInitialized = 18
        PropertyNotFound = 19
        PropertyNotSupported = 20
        ProfileNotFound = 21
    End Enum
     
    'Image Property Tag Constants
    Public Enum PropertyTagId
        'https://docs.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-property-tags-in-alphabetical-order
        'https://docs.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-property-tags-in-numerical-order
        '   0x0... => &H...
        '   https://docs.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-property-item-descriptions
        GpsVer = 0    '&H0&
        GpsLatitudeRef = 1    '&H1&
        GpsLatitude = 2    '&H2&
        GpsLongitudeRef = 3    '&H3&
        GpsLongitude = 4    '&H4&
        GpsAltitudeRef = 5    '&H5&
        GpsAltitude = 6    '&H6&
        GpsGpsTime = 7    '&H7&
        GpsGpsSatellites = 8    '&H8&
        GpsGpsStatus = 9    '&H9&
        GpsGpsMeasureMode = 10    '&HA&
        GpsGpsDop = 11    '&HB&
        GpsSpeedRef = 12    '&HC&
        GpsSpeed = 13    '&HD&
        GpsTrackRef = 14    '&HE&
        GpsTrack = 15    '&HF&
        GpsImgDirRef = 16    '&H10&
        GpsImgDir = 17    '&H11&
        GpsMapDatum = 18    '&H12&
        GpsDestLatRef = 19    '&H13&
        GpsDestLat = 20    '&H14&
        GpsDestLongRef = 21    '&H15&
        GpsDestLong = 22    '&H16
        GpsDestBearRef = 23    '&H17&
        GpsDestBear = 24    '&H18&
        GpsDestDistRef = 25    '&H19&
        GpsDestDist = 26    '&H1A&
        NewSubfileType = 254    '&HFE&
        SubfileType = 255    '&HFF&
        ImageWidth = 256    '&H100&
        ImageHeight = 257    '&H101&
        BitsPerSample = 258    '&H102&
        Compression = 259    '&H103&
        PhotometricInterp = 262    '&H106&
        ThreshHolding = 263    '&H107&
        CellWidth = 264    '&H108&
        CellHeight = 265    '&H109&
        FillOrder = 266    '&H10A&
        DocumentName = 269    '&H10D&
        ImageDescription = 270    '&H10E&
        EquipMake = 271    '&H10F&
        EquipModel = 272    '&H110&
        StripOffsets = 273    '&H111&
        Orientation = 274    '&H112&
        SamplesPerPixel = 277    '&H115&
        RowsPerStrip = 278    '&H116&
        StripBytesCount = 279    '&H117&
        MinSampleValue = 280    '&H118&
        MaxSampleValue = 281    '&H119&
        XResolution = 282    '&H11A&
        YResolution = 283    '&H11B&
        PlanarConfig = 284    '&H11C&
        PageName = 285    '&H11D&
        XPosition = 286    '&H11E&
        YPosition = 287    '&H11F&
        FreeOffset = 288    '&H120&
        FreeByteCounts = 289    '&H121&
        GrayResponseUnit = 290    '&H122&
        GrayResponseCurve = 291    '&H123&
        T4Option = 292    '&H124&
        T6Option = 293    '&H125&
        ResolutionUnit = 296    '&H128&
        PageNumber = 297    '&H129&
        TransferFunction = 301    '&H12D&
        SoftwareUsed = 305    '&H131&
        DateTime = 306    '&H132&
        Artist = 315    '&H13B&
        HostComputer = 316    '&H13C&
        Predictor = 317    '&H13D&
        WhitePoint = 318    '&H13E&
        PrimaryChromaticities = 319    '&H13F&
        ColorMap = 320    '&H140&
        HalftoneHints = 321    '&H141&
        TileWidth = 322    '&H142&
        TileLength = 323    '&H143&
        TileOffset = 324    '&H144&
        TileByteCounts = 325    '&H145&
        InkSet = 332    '&H14C&
        InkNames = 333    '&H14D&
        NumberOfInks = 334    '&H14E&
        DotRange = 336    '&H150&
        TargetPrinter = 337    '&H151&
        ExtraSamples = 338    '&H152&
        SampleFormat = 339    '&H153&
        TransferRange = 342    '&H156&
        JPEGProc = 512    '&H200&
        JPEGInterFormat = 513    '&H201&
        JPEGInterLength = 514    '&H202&
        JPEGRestartInterval = 515    '&H203&
        JPEGLosslessPredictors = 517    '&H205&
        JPEGPointTransforms = 518    '&H206&
        JPEGQTables = 519    '&H207&
        JPEGDCTables = 520    '&H208&
        JPEGACTables = 521    '&H209&
        YCbCrCoefficients = 529    '&H211&
        YCbCrSubsampling = 530    '&H212&
        YCbCrPositioning = 531    '&H213&
        REFBlackWhite = 532    '&H214&
        Gamma = 769    '&H301&
        ICCProfileDescriptor = 770    '&H302&
        SRGBRenderingIntent = 771    '&H303&
        ImageTitle = 800    '&H320&
        ResolutionXUnit = 20481    '&H5001&
        ResolutionYUnit = 20482    '&H5002&
        ResolutionXLengthUnit = 20483    '&H5003&
        ResolutionYLengthUnit = 20484    '&H5004&
        PrintFlags = 20485    '&H5005&
        PrintFlagsVersion = 20486    '&H5006&
        PrintFlagsCrop = 20487    '&H5007&
        PrintFlagsBleedWidth = 20488    '&H5008&
        PrintFlagsBleedWidthScale = 20489    '&H5009&
        HalftoneLPI = 20490    '&H500A&
        HalftoneLPIUnit = 20491    '&H500B&
        HalftoneDegree = 20492    '&H500C&
        HalftoneShape = 20493    '&H500D&
        HalftoneMisc = 20494    '&H500E&
        HalftoneScreen = 20495    '&H500F&
        JPEGQuality = 20496    '&H5010&
        GridSize = 20497    '&H5011&
        ThumbnailFormat = 20498    '&H5012&
        ThumbnailWidth = 20499    '&H5013&
        ThumbnailHeight = 20500    '&H5014&
        ThumbnailColorDepth = 20501    '&H5015&
        ThumbnailPlanes = 20502    '&H5016&
        ThumbnailRawBytes = 20503    '&H5017&
        ThumbnailSize = 20504    '&H5018&
        ThumbnailCompressedSize = 20505    '&H5019&
        ColorTransferFunction = 20506    '&H501A&
        ThumbnailData = 20507    '&H501B&
        ThumbnailImageWidth = 20512    '&H5020&
        ThumbnailImageHeight = 20513    '&H5021&
        ThumbnailBitsPerSample = 20514    '&H5022&
        ThumbnailCompression = 20515    '&H5023&
        ThumbnailPhotometricInterp = 20516    '&H5024&
        ThumbnailImageDescription = 20517    '&H5025&
        ThumbnailEquipMake = 20518    '&H5026&
        ThumbnailEquipModel = 20519    '&H5027&
        ThumbnailStripOffsets = 20520    '&H5028&
        ThumbnailOrientation = 20521    '&H5029&
        ThumbnailSamplesPerPixel = 20522    '&H502A&
        ThumbnailRowsPerStrip = 20523    '&H502B&
        ThumbnailStripBytesCount = 20524    '&H502C&
        ThumbnailResolutionX = 20525    '&H502D&
        ThumbnailResolutionY = 20526    '&H502E&
        ThumbnailPlanarConfig = 20527    '&H502F&
        ThumbnailResolutionUnit = 20528    '&H5030&
        ThumbnailTransferFunction = 20529    '&H5031&
        ThumbnailSoftwareUsed = 20530    '&H5032&
        ThumbnailDateTime = 20531    '&H5033&
        ThumbnailArtist = 20532    '&H5034&
        ThumbnailWhitePoint = 20533    '&H5035&
        ThumbnailPrimaryChromaticities = 20534    '&H5036&
        ThumbnailYCbCrCoefficients = 20535    '&H5037&
        ThumbnailYCbCrSubsampling = 20536    '&H5038&
        ThumbnailYCbCrPositioning = 20537    '&H5039&
        ThumbnailRefBlackWhite = 20538    '&H503A&
        ThumbnailCopyRight = 20539    '&H503B&
        LuminanceTable = 20624    '&H5090&
        ChrominanceTable = 20625    '&H5091&
        FrameDelay = 20736    '&H5100&
        LoopCount = 20737    '&H5101&
        GlobalPalette = 20738    '&H5102&
        IndexBackground = 20739    '&H5103&
        IndexTransparent = 20740    '&H5104&
        PixelUnit = 20752    '&H5110&
        PixelPerUnitX = 20753    '&H5111&
        PixelPerUnitY = 20754    '&H5112&
        PaletteHistogram = 20755    '&H5113&
        Copyright = 33432    '&H8298&
        ExifExposureTime = 33434    '&H829A&
        ExifFNumber = 33437    '&H829D&
        ExifIFD = 34665    '&H8769&
        ICCProfile = 34675    '&H8773&
        ExifExposureProg = 34850    '&H8822&
        ExifSpectralSense = 34852    '&H8824&
        GpsIFD = 34853    '&H8825&
        ExifISOSpeed = 34855    '&H8827&
        ExifOECF = 34856    '&H8828&
        ExifVer = 36864    '&H9000&
        ExifDTOrig = 36867    '&H9003&
        ExifDTDigitized = 36868    '&H9004&
        ExifCompConfig = 37121    '&H9101&
        ExifCompBPP = 37122    '&H9102&
        ExifShutterSpeed = 37377    '&H9201&
        ExifAperture = 37378    '&H9202&
        ExifBrightness = 37379    '&H9203&
        ExifExposureBias = 37380    '&H9204&
        ExifMaxAperture = 37381    '&H9205&
        ExifSubjectDist = 37382    '&H9206&
        ExifMeteringMode = 37383    '&H9207&
        ExifLightSource = 37384    '&H9208&
        ExifFlash = 37385    '&H9209&
        ExifFocalLength = 37386    '&H920A&
        ExifMakerNote = 37500    '&H927C&
        ExifUserComment = 37510    '&H9286&
        ExifDTSubsec = 37520    '&H9290&
        ExifDTOrigSS = 37521    '&H9291&
        ExifDTDigSS = 37522    '&H9292&
        ExifFPXVer = 40960    '&HA000&
        ExifColorSpace = 40961    '&HA001&
        ExifPixXDim = 40962    '&HA002&
        ExifPixYDim = 40963    '&HA003&
        ExifRelatedWav = 40964    '&HA004&
        ExifInterop = 40965    '&HA005&
        ExifFlashEnergy = 41483    '&HA20B&
        ExifSpatialFR = 41484    '&HA20C&
        ExifFocalXRes = 41486    '&HA20E&
        ExifFocalYRes = 41487    '&HA20F&
        ExifFocalResUnit = 41488    '&HA210&
        ExifSubjectLoc = 41492    '&HA214&
        ExifExposureIndex = 41493    '&HA215&
        ExifSensingMethod = 41495    '&HA217&
        ExifFileSource = 41728    '&HA300&
        ExifSceneType = 41729    '&HA301&
        ExifCfaPattern = 41730    '&HA302&
    End Enum
     
    Public Enum PropertyTagType
        'https://docs.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-image-property-tag-type-constants
        TypeByte = 1
        TypeASCII = 2
        TypeShort = 3
        TypeLong = 4
        TypeRational = 5
        TypeUndefined = 7
        TypeSLong = 9
        TypeSRational = 10
    End Enum
    
    #If Win64 Then
        Private Type GdiplusStartupInput
            GdiplusVersion                      As Long
            DebugEventCallback                  As LongPtr
            SuppressBackgroundThread            As LongPtr
            SuppressExternalCodecs              As LongPtr
        End Type
    
        ' GDI - General
        Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, LInput As GdiplusStartupInput, Optional ByVal lOutPut As LongPtr = 0) As Long
        Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
        Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As LongPtr, ByRef bitmap As LongPtr) As Long
        Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
        ' GDI - Image / Properties
        Private Declare PtrSafe Function GdipGetAllPropertyItems Lib "gdiplus" (ByVal Image As LongPtr, ByVal totalBufferSize As Long, ByVal numProperties As Long, allItems As Any) As Long
        Private Declare PtrSafe Function GdipGetPropertySize Lib "gdiplus" (ByVal Image As LongPtr, totalBufferSize As Long, numProperties As Long) As Long
        Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal FileName As LongPtr, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
        Private Declare PtrSafe Function GdipSetPropertyItem Lib "gdiplus" (ByVal nImage As LongPtr, Item As PropertyItem) As Long
        
        ' Helper API Declarations
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
        Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal str As LongPtr, id As GUID) As Long
        Private Declare PtrSafe Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As LongPtr) As Long
        Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr
                
        Dim lGDIpToken                          As LongPtr
        Dim lBitmap                             As LongPtr
    #Else
        Private Type GdiplusStartupInput
            GdiplusVersion            As Long
            DebugEventCallback        As Long
            SuppressBackgroundThread  As Long
            SuppressExternalCodecs    As Long
        End Type
            
       'G DI - General
        Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Status
        Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Status
        Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef Bitmap As Long) As Status
        Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Status
        ' GDI - Image / Properties
        Private Declare Function GdipGetAllPropertyItems Lib "gdiplus" (ByVal Image As Long, ByVal totalBufferSize As Long, ByVal numProperties As Long, ByRef allItems As PropertyItem) As Status
        Private Declare Function GdipGetPropertySize Lib "gdiplus" (ByVal Image As Long, ByRef totalBufferSize As Long, ByRef numProperties As Long) As Status
        Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Status
        Private Declare Function GdipSetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByRef item As Long) As Status
         
        ' Helper API Declarations
        Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
        Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
        Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
        Dim lGDIpToken                 As Long
        Dim lBitmap                    As Long
    #End If
    
    Dim bGDIpInitialized           As Boolean
     
    '---------------------------------------------------------------------------------------
    ' Procedure : SetImageProperty
    ' Author    : Daniel Pineault, CARDA Consultants Inc.
    ' Website   : http://www.cardaconsultants.com
    ' Purpose   : Set the image property with the specified value
    ' Copyright : The following is release as Attribution-ShareAlike 4.0 International
    '             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
    '
    ' Input Variables:
    ' ~~~~~~~~~~~~~~~~
    ' sFile             : Fully qualified path and filename of the image file to get info about
    ' lPropertyTagId    : Property to the set the value of
    ' vPropertyValue    : Value to set for the property.  In the case of Rational / Nominator value
    ' vPropertyValue2   : Rational only / Denominator value
    '
    ' Usage:
    ' ~~~~~~
    ' - SetImageProperty("C:\Temp\IMG01 - Copy.jpg", EquipModel , "Huawei")
    '   Returns -> True/False
    ' - SetImageProperty("C:\Temp\IMG01 - Copy.jpg", ImageWidth , 179306496)
    '   Returns -> True/False
    ' - SetImageProperty("C:\Temp\IMG01 - Copy.jpg", ExifISOSpeed , 250)
    '   Returns -> True/False
    ' - SetImageProperty("C:\Temp\IMG01 - Copy.jpg", ImageDescription , "View from the house")
    '   Returns -> True/False
    ' - SetImageProperty("C:\Temp\IMG01 - Copy.jpg",ExifFNumber , 180, 100)
    '   Returns -> True/False
    '
    ' Revision History:
    ' Rev       Date(yyyy-mm-dd)        Description
    ' **************************************************************************************
    ' 1         2022-01-09              Initial Blog Release
    '---------------------------------------------------------------------------------------
    Public Function SetImageProperty(ByVal sFile As String, ByVal lPropertyTagId As PropertyTagId, ByVal vPropertyValue As Variant, Optional ByVal vPropertyValue2 As Variant) As Boolean
        On Error GoTo Error_Handler
        Dim GDIpStartupInput      As GdiplusStartupInput
        Dim GDIStatus             As Status
        Dim PI                    As PropertyItem
        Dim byteValue             As Byte
        Dim sValue                As String
        Dim lValue                As Long
        Dim iValue                As Integer
        Dim lPropertyType         As PropertyTagType
        Dim tEncoder              As GUID
        Dim tParams               As EncoderParameters
        Dim sTmp                  As String
        Dim sExt                  As String
     
        'Start GDI
        '-------------------------------------------------------------------------------------
        If bGDIpInitialized = False Then
            GDIpStartupInput.GdiplusVersion = 1
            GDIStatus = GdiplusStartup(lGDIpToken, GDIpStartupInput, ByVal 0)
            If GDIStatus <> Status.OK Then
                MsgBox "Unable to start the GDI+ API" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                GoTo Error_Handler_Exit
            Else
                bGDIpInitialized = True
            End If
        End If
     
        'Load our Image to work with
        '-------------------------------------------------------------------------------------
        'In case we already have something in memory let's dispose of it properly
        If lBitmap <> 0 Then
            GDIStatus = GdipDisposeImage(lBitmap)
            If GDIStatus <> Status.OK Then
                MsgBox "Unable to dispose of the current image in memory" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                GoTo Error_Handler_Exit
            End If
        End If
        'Now let's proceed with loading the actual image we want to work with
        GDIStatus = GdipCreateBitmapFromFile(StrPtr(sFile), lBitmap)
        If GDIStatus <> Status.OK Then
            MsgBox "Unable to load the specified image" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
            GoTo Error_Handler_Exit
        End If
     
        'Work with the image
        '-------------------------------------------------------------------------------------
        'Get Property Type
        lPropertyType = GetPropertyType(lPropertyTagId)
     
        PI.id = lPropertyTagId
        PI.Type = lPropertyType
     
        Select Case lPropertyType
            Case PropertyTagType.TypeByte    '1
                byteValue = vPropertyValue
     
                PI.Length = LenB(byteValue)
                PI.value = VarPtr(byteValue)
            Case PropertyTagType.TypeASCII    '2
                sValue = vPropertyValue & vbNullChar
     
                PI.Length = Len(sValue)
                PI.value = StrPtr(StrConv(sValue, vbFromUnicode))
            Case PropertyTagType.TypeShort    '3
                iValue = vPropertyValue
     
                PI.Length = LenB(iValue)
                PI.value = VarPtr(iValue)
            Case PropertyTagType.TypeLong, PropertyTagType.TypeSLong    '4, 9
                Select Case lPropertyType
                    Case PropertyTagType.TypeLong
                        lValue = Abs(vPropertyValue)
                    Case PropertyTagType.TypeSLong
                        lValue = vPropertyValue
                End Select
     
                PI.Length = LenB(lValue)
                PI.value = VarPtr(lValue)
            Case PropertyTagType.TypeRational, PropertyTagType.TypeSRational    '5, 10
                Dim dValue        As Double
                Dim bytProperty() As Byte
                Dim DataSize      As Long
                Dim lNumerator    As Long
                Dim lDenominator  As Long
     
                Select Case lPropertyType
                    Case PropertyTagType.TypeRational
                        lNumerator = Abs(vPropertyValue)
                        lDenominator = Abs(vPropertyValue2)
                    Case PropertyTagType.TypeSRational
                        lNumerator = vPropertyValue
                        lDenominator = vPropertyValue2
                End Select
     
                DataSize = LenB(dValue)
                ReDim bytProperty(DataSize - 1)
                Call CopyMemory(bytProperty(0), lNumerator, DataSize / 2)
                Call CopyMemory(bytProperty(0 + (DataSize / 2)), lDenominator, DataSize / 2)
                Call CopyMemory(dValue, bytProperty(0), DataSize)
                Erase bytProperty
     
                PI.Length = DataSize    '8
                PI.value = VarPtr(dValue)
            Case Else    'PropertyTagType.TypeUndefined    '7
                Exit Function
        End Select
     
        GDIStatus = GdipSetPropertyItem(lBitmap, PI)
        If GDIStatus <> Status.OK Then
            MsgBox "Unable to set the specified property" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
            GoTo Error_Handler_Exit
        Else
            'Save the changes
            '   Cannot save over the file itself, so create a tmp file, release the current one then copy the image over
            sExt = Mid(sFile, InStrRev(sFile, ".") + 1)
            Select Case sExt
                Case "bmp", "dib"
                    CLSIDFromString StrPtr(ImageCodecBMP), tEncoder
                Case "gif"
                    CLSIDFromString StrPtr(ImageCodecGIF), tEncoder
                Case "jpg", "jpeg", "jpe", "jfif"
                    CLSIDFromString StrPtr(ImageCodecJPG), tEncoder
     
                    With tParams
                        .Count = 1
                        .Parameter(0).NumberOfValues = 1
                        .Parameter(0).Type = EncoderParameterValueTypeLong
                        .Parameter(0).value = 100    '100% Quality
                        CLSIDFromString StrPtr(EncoderQuality), .Parameter(0).GUID
                    End With
                Case "png"
                    CLSIDFromString StrPtr(ImageCodecPNG), tEncoder
                Case "tif", "tiff"
                    CLSIDFromString StrPtr(ImageCodecTIF), tEncoder
     
                    With tParams
                        .Count = 1
                        .Parameter(0).NumberOfValues = 1
                        .Parameter(0).Type = EncoderParameterValueTypeLong
                        .Parameter(0).value = TiffCompressionNone
                        CLSIDFromString StrPtr(EncoderCompression), .Parameter(0).GUID
                    End With
                Case Else
                    Exit Function
            End Select
     
            sTmp = Environ("Temp") & "\TempSave." & sExt
            GDIStatus = GdipSaveImageToFile(lBitmap, StrPtr(sTmp), tEncoder, ByVal tParams)
            If GDIStatus <> Status.OK Then
                MsgBox "Unable to save the image" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                Exit Function
            End If
        End If
     
Error_Handler_Exit:
        On Error Resume Next
        'Shutdown GDI
        '-------------------------------------------------------------------------------------
        If bGDIpInitialized = True Then
            If lBitmap <> 0 Then
                GDIStatus = GdipDisposeImage(lBitmap)
                If GDIStatus <> Status.OK Then
                    MsgBox "Unable to dispose of the processed image" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                    Exit Function
                Else
                    lBitmap = 0
                    'Overwrite the file with our Temp File and Cleanup
                    If sTmp <> "" Then
                        Call FileCopy(sTmp, sFile)
                        If Len(Dir(sFile)) > 0 Then
                            SetImageProperty = True
                            Kill sTmp
                        End If
                    End If
                End If
            End If
            GDIStatus = GdiplusShutdown(lGDIpToken)
            If GDIStatus <> Status.OK Then
                MsgBox "Unable to shutdown the GDI+ API" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                Exit Function
            Else
                bGDIpInitialized = False
            End If
        End If
        Exit Function
     
Error_Handler:
        MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Source: GetImageProperty" & vbCrLf & _
               "Error Description: " & Err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occured!"
        Resume Error_Handler_Exit
    End Function
    
    Private Function GetPropertyValue(ByVal PropertyItemLength As Long, ByVal PropertyItemValue As LongPtr, ByVal lPropertyType As PropertyTagType) As String
        Dim byteRetValue          As Byte
        Dim lRetValue             As Long
        Dim iRetValue             As Integer
        Dim bytProperty()         As Byte
        Dim sProperty             As String
        Dim DataLength            As Long
        Dim Numerator             As Long
        Dim Denominator           As Long
        Dim sTemp                 As String
        Dim i                     As Long
     
        If PropertyItemLength = 0 Then
            GetPropertyValue = ""
            Exit Function
        End If
     
        Select Case lPropertyType
            Case PropertyTagType.TypeByte
                DataLength = 1
                ReDim bytProperty(PropertyItemLength - 1)
                Call CopyMemory(bytProperty(0), ByVal PropertyItemValue, PropertyItemLength)
                Call CopyMemory(byteRetValue, bytProperty(0), DataLength)
                Erase bytProperty
                GetPropertyValue = byteRetValue
            Case PropertyTagType.TypeASCII
                sProperty = Space$(lstrlen(ByVal PropertyItemValue))
                Call lstrcpy(ByVal StrPtr(sProperty), ByVal PropertyItemValue)
                GetPropertyValue = Trim$(Left$(StrConv(sProperty, vbUnicode), PropertyItemLength - 1))
            Case PropertyTagType.TypeShort
                DataLength = 2
                ReDim bytProperty(PropertyItemLength - 1)
                Call CopyMemory(bytProperty(0), ByVal PropertyItemValue, PropertyItemLength)
                Call CopyMemory(iRetValue, bytProperty(0), DataLength)
                Erase bytProperty
                GetPropertyValue = iRetValue
            Case PropertyTagType.TypeLong, PropertyTagType.TypeSLong
                DataLength = 4
                ReDim bytProperty(PropertyItemLength - 1)
                Call CopyMemory(bytProperty(0), ByVal PropertyItemValue, PropertyItemLength)
                Call CopyMemory(lRetValue, bytProperty(0), DataLength)
                Erase bytProperty
                GetPropertyValue = lRetValue
            Case PropertyTagType.TypeRational, PropertyTagType.TypeSRational
                DataLength = 8
                ReDim bytProperty(PropertyItemLength - 1)
                Call CopyMemory(bytProperty(0), ByVal PropertyItemValue, PropertyItemLength)
                Call CopyMemory(Numerator, bytProperty(0), DataLength / 2)
                Call CopyMemory(Denominator, bytProperty(0 + (DataLength / 2)), DataLength / 2)
                Erase bytProperty
                GetPropertyValue = CStr(Numerator) & "/" & CStr(Denominator)
                GetPropertyValue = Application.Evaluate(GetPropertyValue)
            Case PropertyTagType.TypeUndefined
                ReDim bytProperty(PropertyItemLength - 1)
                Call CopyMemory(bytProperty(0), ByVal PropertyItemValue, PropertyItemLength)
                For i = 1 To PropertyItemLength - 1
                    If i > 1 Then sProperty = sProperty & Chr$(32)
                    sTemp = Hex$(bytProperty(i - 1))
                    If Len(sTemp) = 1 Then sTemp = "0" & sTemp
                    sProperty = sProperty & sTemp
                Next i
                Erase bytProperty
                GetPropertyValue = sProperty
        End Select
    End Function
    
    '---------------------------------------------------------------------------------------
    ' Procedure : GetImageProperties
    ' Author    : Daniel Pineault, CARDA Consultants Inc.
    ' Website   : http://www.cardaconsultants.com
    ' Purpose   : Returns all the properties of a given image file
    ' Copyright : The following is release as Attribution-ShareAlike 4.0 International
    '             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
    '
    ' Input Variables:
    ' ~~~~~~~~~~~~~~~~
    ' sFile     : Fully qualified path and filename of the image file to get info about
    '
    ' Usage:
    ' ~~~~~~
    ' sImgProperties = GetImageProperties("C:\Temp\IMG_20210508_170154.jpg")
    '
    ' Revision History:
    ' Rev       Date(yyyy-mm-dd)        Description
    ' **************************************************************************************
    ' 1         2022-01-09              Initial Blog Release
    '---------------------------------------------------------------------------------------
    Public Function GetImageProperties(sFile As String) As String
        On Error GoTo Error_Handler
        Dim GDIpStartupInput      As GdiplusStartupInput
        Dim GDIStatus             As Status
        Dim lAllPropertiesSize    As Long
        Dim lAllPropertiesCount   As Long
        Dim PI()                  As PropertyItem
        Dim sOutput               As String
        Dim propItem              As Long
     
        'Start GDI
        '-------------------------------------------------------------------------------------
        If bGDIpInitialized = False Then
            GDIpStartupInput.GdiplusVersion = 1
            GDIStatus = GdiplusStartup(lGDIpToken, GDIpStartupInput, ByVal 0)
            If GDIStatus <> Status.OK Then
                MsgBox "Unable to start the GDI+ API" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                GoTo Error_Handler_Exit
            Else
                bGDIpInitialized = True
            End If
        End If
     
        'Load our Image to work with
        '-------------------------------------------------------------------------------------
        'In case we already have something in memory let's dispose of it properly
        If lBitmap <> 0 Then
            GDIStatus = GdipDisposeImage(lBitmap)
            If GDIStatus <> Status.OK Then
                MsgBox "Unable to dispose of the current image in memory" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                GoTo Error_Handler_Exit
            End If
        End If
        'Now let's proceed with loading the actual image we want to work with
        GDIStatus = GdipCreateBitmapFromFile(StrPtr(sFile), lBitmap)
        If GDIStatus <> Status.OK Then
            MsgBox "Unable to load the specified image" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
            GoTo Error_Handler_Exit
        End If
     
        'Work with the image
        '-------------------------------------------------------------------------------------
        If lBitmap Then
            GDIStatus = GdipGetPropertySize(lBitmap, lAllPropertiesSize, lAllPropertiesCount)
            If GDIStatus <> Status.OK Then
                MsgBox "Unable to determine the size of all the properties" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                GoTo Error_Handler_Exit
            End If
            If (lAllPropertiesCount > 0) Then
                ReDim PI(0 To lAllPropertiesSize \ Len(PI(0)) - 1)
                GDIStatus = GdipGetAllPropertyItems(lBitmap, lAllPropertiesSize, lAllPropertiesCount, PI(0))
                If GDIStatus <> Status.OK Then
                    MsgBox "Unable to get all the image properties" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                    GoTo Error_Handler_Exit
                End If
                For propItem = 0 To lAllPropertiesCount - 1
                    If propItem <> 0 Then sOutput = sOutput & vbNewLine
                    sOutput = sOutput & PropertyTagIdToString(PI(propItem).id) & ": " & GetPropertyValue(PI(propItem).Length, PI(propItem).value, PI(propItem).Type)
                Next propItem
                GetImageProperties = sOutput
     
                Erase PI
            End If
        End If
     
Error_Handler_Exit:
        On Error Resume Next
        'Shutdown GDI
        '-------------------------------------------------------------------------------------
        If bGDIpInitialized = True Then
            If lBitmap <> 0 Then
                GDIStatus = GdipDisposeImage(lBitmap)
                If GDIStatus <> Status.OK Then
                    MsgBox "Unable to dispose of the processed image" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                    Exit Function
                Else
                    lBitmap = 0
                End If
            End If
            GDIStatus = GdiplusShutdown(lGDIpToken)
            If GDIStatus <> Status.OK Then
                MsgBox "Unable to shutdown the GDI+ API" & vbCrLf & vbCrLf & GDIErrorToString(GDIStatus), vbCritical Or vbOKOnly, "Operation Aborted"
                Exit Function
            Else
                bGDIpInitialized = False
            End If
        End If
        Exit Function
     
Error_Handler:
        MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Source: GetImageProperty" & vbCrLf & _
               "Error Description: " & Err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occured!"
        Resume Error_Handler_Exit
    End Function
    
    Private Function PropertyTagIdToString(ByVal lPropertyTagId As PropertyTagId) As String
        Select Case lPropertyTagId
            Case &H0&
                PropertyTagIdToString = "GPS Ver"
            Case &H1&
                PropertyTagIdToString = "GPS Latitude Ref"
            Case &H2&
                PropertyTagIdToString = "GPS Latitude"
            Case &H3&
                PropertyTagIdToString = "GPS Longitude Ref"
            Case &H4&
                PropertyTagIdToString = "GPS Longitude"
            Case &H5&
                PropertyTagIdToString = "GPS Altitude Ref"
            Case &H6&
                PropertyTagIdToString = "GPS Altitude"
            Case &H7&
                PropertyTagIdToString = "GPS Time"
            Case &H8&
                PropertyTagIdToString = "GPS Satellites"
            Case &H9&
                PropertyTagIdToString = "GPS Status"
            Case &HA&
                PropertyTagIdToString = "GPS Measure Mode"
            Case &HB&
                PropertyTagIdToString = "GPS Dop"
            Case &HC&
                PropertyTagIdToString = "GPS Speed Ref"
            Case &HD&
                PropertyTagIdToString = "GPS Speed"
            Case &HE&
                PropertyTagIdToString = "GPS Track Ref"
            Case &HF&
                PropertyTagIdToString = "GPS Track"
            Case &H10&
                PropertyTagIdToString = "GPS Img Dir Ref"
            Case &H11&
                PropertyTagIdToString = "GPS Img Dir"
            Case &H12&
                PropertyTagIdToString = "GPS Map Datum"
            Case &H13&
                PropertyTagIdToString = "GPS Dest Lat Ref"
            Case &H14&
                PropertyTagIdToString = "GPS Dest Lat"
            Case &H15&
                PropertyTagIdToString = "GPS Dest Long Ref"
            Case &H16&
                PropertyTagIdToString = "GPS Dest Long"
            Case &H17&
                PropertyTagIdToString = "GPS Dest Bear Ref"
            Case &H18&
                PropertyTagIdToString = "GPS Dest Bear"
            Case &H19&
                PropertyTagIdToString = "GPS Dest Dist Ref"
            Case &H1A&
                PropertyTagIdToString = "GPS Des tDist"
            Case &HFE&
                PropertyTagIdToString = "New Subfile Type"
            Case &HFF&
                PropertyTagIdToString = "Sub file Type"
            Case &H100&
                PropertyTagIdToString = "Image Width"
            Case &H101&
                PropertyTagIdToString = "Image Height"
            Case &H102&
                PropertyTagIdToString = "Bits Per Sample"
            Case &H103&
                PropertyTagIdToString = "Compression"
            Case &H106&
                PropertyTagIdToString = "Photometric Interp"
            Case &H107&
                PropertyTagIdToString = "Thresh Holding"
            Case &H108&
                PropertyTagIdToString = "Cell Width"
            Case &H109&
                PropertyTagIdToString = "Cell Height"
            Case &H10A&
                PropertyTagIdToString = "Fill Order"
            Case &H10D&
                PropertyTagIdToString = "Document Name"
            Case &H10E&
                PropertyTagIdToString = "Image Description"
            Case &H10F&
                PropertyTagIdToString = "Equip Make"
            Case &H110&
                PropertyTagIdToString = "Equip Model"
            Case &H111&
                PropertyTagIdToString = "Strip Offsets"
            Case &H112&
                PropertyTagIdToString = "Orientation"
            Case &H115&
                PropertyTagIdToString = "Samples Per Pixel"
            Case &H116&
                PropertyTagIdToString = "Rows Per Strip"
            Case &H117&
                PropertyTagIdToString = "Strip Bytes Count"
            Case &H118&
                PropertyTagIdToString = "Min Sample Value"
            Case &H119&
                PropertyTagIdToString = "Max Sample Value"
            Case &H11A&
                PropertyTagIdToString = "X Resolution"
            Case &H11B&
                PropertyTagIdToString = "Y Resolution"
            Case &H11C&
                PropertyTagIdToString = "Planar Config"
            Case &H11D&
                PropertyTagIdToString = "Page Name"
            Case &H11E&
                PropertyTagIdToString = "X Position"
            Case &H11F&
                PropertyTagIdToString = "Y Position"
            Case &H120&
                PropertyTagIdToString = "Free Offset"
            Case &H121&
                PropertyTagIdToString = "Free Byte Counts"
            Case &H122&
                PropertyTagIdToString = "Gray Response Unit"
            Case &H123&
                PropertyTagIdToString = "Gray Response Curve"
            Case &H124&
                PropertyTagIdToString = "T4 Option"
            Case &H125&
                PropertyTagIdToString = "T6 Option"
            Case &H128&
                PropertyTagIdToString = "Resolution Unit"
            Case &H129&
                PropertyTagIdToString = "Page Number"
            Case &H12D&
                PropertyTagIdToString = "Transfer Function"
            Case &H131&
                PropertyTagIdToString = "Software Used"
            Case &H132&
                PropertyTagIdToString = "Date/Time"
            Case &H13B&
                PropertyTagIdToString = "Artist"
            Case &H13C&
                PropertyTagIdToString = "Host Computer"
            Case &H13D&
                PropertyTagIdToString = "Predictor"
            Case &H13E&
                PropertyTagIdToString = "White Point"
            Case &H13F&
                PropertyTagIdToString = "Primary Chromaticities"
            Case &H140&
                PropertyTagIdToString = "Color Map"
            Case &H141&
                PropertyTagIdToString = "Halftone Hints"
            Case &H142&
                PropertyTagIdToString = "Tile Width"
            Case &H143&
                PropertyTagIdToString = "Tile Length"
            Case &H144&
                PropertyTagIdToString = "Tile Offset"
            Case &H145&
                PropertyTagIdToString = "Tile Byte Counts"
            Case &H14C&
                PropertyTagIdToString = "Ink Set"
            Case &H14D&
                PropertyTagIdToString = "Ink Names"
            Case &H14E&
                PropertyTagIdToString = "Number Of Inks"
            Case &H150&
                PropertyTagIdToString = "Dot Range"
            Case &H151&
                PropertyTagIdToString = "Target Printer"
            Case &H152&
                PropertyTagIdToString = "ExtraS amples"
            Case &H153&
                PropertyTagIdToString = "Sample Format"
            Case &H156&
                PropertyTagIdToString = "Transfer Range"
            Case &H200&
                PropertyTagIdToString = "JPEG Proc"
            Case &H201&
                PropertyTagIdToString = "JPEG Inter Format"
            Case &H202&
                PropertyTagIdToString = "JPEG Inter Length"
            Case &H203&
                PropertyTagIdToString = "JPEG Restart Interval"
            Case &H205&
                PropertyTagIdToString = "JPEG Lossless Predictors"
            Case &H206&
                PropertyTagIdToString = "JPEG Point Transforms"
            Case &H207&
                PropertyTagIdToString = "JPEG Q Tables"
            Case &H208&
                PropertyTagIdToString = "JPEG DC Tables"
            Case &H209&
                PropertyTagIdToString = "JPEG AC Tables"
            Case &H211&
                PropertyTagIdToString = "YCb Cr Coefficients"
            Case &H212&
                PropertyTagIdToString = "YCb Cr Subsampling"
            Case &H213&
                PropertyTagIdToString = "YCb Cr Positioning"
            Case &H214&
                PropertyTagIdToString = "REF Black White"
            Case &H301&
                PropertyTagIdToString = "Gamma"
            Case &H302&
                PropertyTagIdToString = "ICC Profile Descriptor"
            Case &H303&
                PropertyTagIdToString = "SRGB Rendering Intent"
            Case &H320&
                PropertyTagIdToString = "Image Title"
            Case &H5001&
                PropertyTagIdToString = "Resolution X Unit"
            Case &H5002&
                PropertyTagIdToString = "Resolution Y Unit"
            Case &H5003&
                PropertyTagIdToString = "Resolution X Length Unit"
            Case &H5004&
                PropertyTagIdToString = "Resolution Y Length Unit"
            Case &H5005&
                PropertyTagIdToString = "Print Flags"
            Case &H5006&
                PropertyTagIdToString = "Print Flags Version"
            Case &H5007&
                PropertyTagIdToString = "Print Flags Crop"
            Case &H5008&
                PropertyTagIdToString = "Print Flags Bleed Width"
            Case &H5009&
                PropertyTagIdToString = "Print Flags Bleed Width Scale"
            Case &H500A&
                PropertyTagIdToString = "Halftone LPI"
            Case &H500B&
                PropertyTagIdToString = "Halftone LPI Unit"
            Case &H500C&
                PropertyTagIdToString = "Halftone Degree"
            Case &H500D&
                PropertyTagIdToString = "Halftone Shape"
            Case &H500E&
                PropertyTagIdToString = "Halftone Misc"
            Case &H500F&
                PropertyTagIdToString = "Halftone Screen"
            Case &H5010&
                PropertyTagIdToString = "JPEG Quality"
            Case &H5011&
                PropertyTagIdToString = "Grid Size"
            Case &H5012&
                PropertyTagIdToString = "Thumbnail Format"
            Case &H5013&
                PropertyTagIdToString = "Thumbnail Width"
            Case &H5014&
                PropertyTagIdToString = "Thumbnail Height"
            Case &H5015&
                PropertyTagIdToString = "Thumbnail Color Depth"
            Case &H5016&
                PropertyTagIdToString = "Thumbnail Planes"
            Case &H5017&
                PropertyTagIdToString = "Thumbnail Raw Bytes"
            Case &H5018&
                PropertyTagIdToString = "Thumbnail Size"
            Case &H5019&
                PropertyTagIdToString = "Thumbnail Compressed Size"
            Case &H501A&
                PropertyTagIdToString = "Color Transfer Function"
            Case &H501B&
                PropertyTagIdToString = "Thumbnail Data"
            Case &H5020&
                PropertyTagIdToString = "Thumbnail Image Width"
            Case &H5021&
                PropertyTagIdToString = "Thumbnail Image Height"
            Case &H5022&
                PropertyTagIdToString = "Thumbnail Bits Per Sample"
            Case &H5023&
                PropertyTagIdToString = "Thumbnail Compression"
            Case &H5024&
                PropertyTagIdToString = "Thumbnail Photometric Interp"
            Case &H5025&
                PropertyTagIdToString = "Thumbnail Image Description"
            Case &H5026&
                PropertyTagIdToString = "Thumbnail EquipMake"
            Case &H5027&
                PropertyTagIdToString = "Thumbnail EquipModel"
            Case &H5028&
                PropertyTagIdToString = "Thumbnail Strip Offsets"
            Case &H5029&
                PropertyTagIdToString = "Thumbnail Orientation"
            Case &H502A&
                PropertyTagIdToString = "Thumbnail Samples Per Pixel"
            Case &H502B&
                PropertyTagIdToString = "Thumbnail Rows Per Strip"
            Case &H502C&
                PropertyTagIdToString = "Thumbnail Strip Bytes Count"
            Case &H502D&
                PropertyTagIdToString = "Thumbnail Resolution X"
            Case &H502E&
                PropertyTagIdToString = "Thumbnail Resolution Y"
            Case &H502F&
                PropertyTagIdToString = "Thumbnail Planar Config"
            Case &H5030&
                PropertyTagIdToString = "Thumbnail Resolution Unit"
            Case &H5031&
                PropertyTagIdToString = "Thumbnail Transfer Function"
            Case &H5032&
                PropertyTagIdToString = "Thumbnail Software Used"
            Case &H5033&
                PropertyTagIdToString = "Thumbnail Date/Time"
            Case &H5034&
                PropertyTagIdToString = "Thumbnail Artist"
            Case &H5035&
                PropertyTagIdToString = "Thumbnail White Point"
            Case &H5036&
                PropertyTagIdToString = "Thumbnail Primary Chromaticities"
            Case &H5037&
                PropertyTagIdToString = "Thumbnail YCb Cr Coefficients"
            Case &H5038&
                PropertyTagIdToString = "Thumbnail YCb Cr Subsampling"
            Case &H5039&
                PropertyTagIdToString = "Thumbnail YCb Cr Positioning"
            Case &H503A&
                PropertyTagIdToString = "Thumbnail Ref Black White"
            Case &H503B&
                PropertyTagIdToString = "Thumbnail Copyright"
            Case &H5090&
                PropertyTagIdToString = "Luminance Table"
            Case &H5091&
                PropertyTagIdToString = "Chrominance Table"
            Case &H5100&
                PropertyTagIdToString = "Frame Delay"
            Case &H5101&
                PropertyTagIdToString = "Loop Count"
            Case &H5102&
                PropertyTagIdToString = "Global Palette"
            Case &H5103&
                PropertyTagIdToString = "Index Background"
            Case &H5104&
                PropertyTagIdToString = "Index Transparent"
            Case &H5110&
                PropertyTagIdToString = "Pixe lUnit"
            Case &H5111&
                PropertyTagIdToString = "Pixel Per Unit X"
            Case &H5112&
                PropertyTagIdToString = "Pixel Per Unit Y"
            Case &H5113&
                PropertyTagIdToString = "Palette Histogram"
            Case &H8298&
                PropertyTagIdToString = "Copyright"
            Case &H829A&
                PropertyTagIdToString = "Exif Exposure Time"
            Case &H829D&
                PropertyTagIdToString = "Exif F Number"
            Case &H8769&
                PropertyTagIdToString = "Exif IFD"
            Case &H8773&
                PropertyTagIdToString = "ICC Profile"
            Case &H8822&
                PropertyTagIdToString = "Exif Exposure Prog"
            Case &H8824&
                PropertyTagIdToString = "Exif Spectral Sense"
            Case &H8825&
                PropertyTagIdToString = "GPS IFD"
            Case &H8827&
                PropertyTagIdToString = "Exif ISOS peed"
            Case &H8828&
                PropertyTagIdToString = "Exif OECF"
            Case &H9000&
                PropertyTagIdToString = "Exif Ver"
            Case &H9003&
                PropertyTagIdToString = "Exif DT Orig"
            Case &H9004&
                PropertyTagIdToString = "Exif DT Digitized"
            Case &H9101&
                PropertyTagIdToString = "Exif Comp Config"
            Case &H9102&
                PropertyTagIdToString = "Exif Comp BPP"
            Case &H9201&
                PropertyTagIdToString = "Exif Shutter Speed"
            Case &H9202&
                PropertyTagIdToString = "Exif Aperture"
            Case &H9203&
                PropertyTagIdToString = "Exif Brightness"
            Case &H9204&
                PropertyTagIdToString = "Exif Exposure Bias"
            Case &H9205&
                PropertyTagIdToString = "Exif Max Aperture"
            Case &H9206&
                PropertyTagIdToString = "Exif Subject Dist"
            Case &H9207&
                PropertyTagIdToString = "Exif Metering Mode"
            Case &H9208&
                PropertyTagIdToString = "Exif Light Source"
            Case &H9209&
                PropertyTagIdToString = "Exif Flash"
            Case &H920A&
                PropertyTagIdToString = "Exif Focal Length"
            Case &H927C&
                PropertyTagIdToString = "Exif Maker Note"
            Case &H9286&
                PropertyTagIdToString = "Exif User Comment"
            Case &H9290&
                PropertyTagIdToString = "Exif DT Subsec"
            Case &H9291&
                PropertyTagIdToString = "Exif DT Orig SS"
            Case &H9292&
                PropertyTagIdToString = "Exif DT Dig SS"
            Case &HA000&
                PropertyTagIdToString = "Exif FPX Ver"
            Case &HA001&
                PropertyTagIdToString = "Exif Color Space"
            Case &HA002&
                PropertyTagIdToString = "Exif Pix X Dim"
            Case &HA003&
                PropertyTagIdToString = "Exif Pix Y Dim"
            Case &HA004&
                PropertyTagIdToString = "Exif Related Wav"
            Case &HA005&
                PropertyTagIdToString = "Exif Interop"
            Case &HA20B&
                PropertyTagIdToString = "Exif Flash Energy"
            Case &HA20C&
                PropertyTagIdToString = "Exif Spatial FR"
            Case &HA20E&
                PropertyTagIdToString = "Exif Focal X Res"
            Case &HA20F&
                PropertyTagIdToString = "Exif Focal Y Res"
            Case &HA210&
                PropertyTagIdToString = "Exif Focal Res Unit"
            Case &HA214&
                PropertyTagIdToString = "Exif Subject Loc"
            Case &HA215&
                PropertyTagIdToString = "Exif Exposure Index"
            Case &HA217&
                PropertyTagIdToString = "Exif Sensing Method"
            Case &HA300&
                PropertyTagIdToString = "Exif File Source"
            Case &HA301&
                PropertyTagIdToString = "Exif Scene Type"
            Case &HA302&
                PropertyTagIdToString = "Exif Cfa Pattern"
            Case Else
                PropertyTagIdToString = "Unknown Property"
        End Select
    End Function
     
    Private Function GDIErrorToString(ByVal lGDIError As Status) As String
        Select Case lGDIError
            Case GenericError
                GDIErrorToString = "Generic Error."
            Case InvalidParameter
                GDIErrorToString = "Invalid Parameter."
            Case OutOfMemory
                GDIErrorToString = "Out Of Memory."
            Case ObjectBusy
                GDIErrorToString = "Object Busy."
            Case InsufficientBuffer
                GDIErrorToString = "Insufficient Buffer."
            Case NotImplemented
                GDIErrorToString = "Not Implemented."
            Case Win32Error
                GDIErrorToString = "Win32 Error."
            Case WrongState
                GDIErrorToString = "Wrong State."
            Case Aborted
                GDIErrorToString = "Aborted."
            Case FileNotFound
                GDIErrorToString = "File Not Found."
            Case ValueOverflow
                GDIErrorToString = "Value Overflow."
            Case AccessDenied
                GDIErrorToString = "Access Denied."
            Case UnknownImageFormat
                GDIErrorToString = "Unknown Image Format."
            Case FontFamilyNotFound
                GDIErrorToString = "FontFamily Not Found."
            Case FontStyleNotFound
                GDIErrorToString = "FontStyle Not Found."
            Case NotTrueTypeFont
                GDIErrorToString = "Not TrueType Font."
            Case UnsupportedGdiplusVersion
                GDIErrorToString = "Unsupported Gdiplus Version."
            Case GdiplusNotInitialized
                GDIErrorToString = "Gdiplus Not Initialized."
            Case PropertyNotFound
                GDIErrorToString = "Property Not Found."
            Case PropertyNotSupported
                GDIErrorToString = "Property Not Supported."
            Case Else
                GDIErrorToString = "Unknown Error."
        End Select
    End Function
     
    Private Function GetPropertyType(lPropertyTagId As PropertyTagId) As PropertyTagType
    'https://docs.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-property-item-descriptions
        Select Case lPropertyTagId
            Case PropertyTagId.GpsVer: GetPropertyType = TypeByte
            Case PropertyTagId.GpsLatitudeRef: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsLatitude: GetPropertyType = TypeRational
            Case PropertyTagId.GpsLongitudeRef: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsLongitude: GetPropertyType = TypeRational
            Case PropertyTagId.GpsAltitudeRef: GetPropertyType = TypeByte
            Case PropertyTagId.GpsAltitude: GetPropertyType = TypeRational
            Case PropertyTagId.GpsGpsTime: GetPropertyType = TypeRational
            Case PropertyTagId.GpsGpsSatellites: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsGpsStatus: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsGpsMeasureMode: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsGpsDop: GetPropertyType = TypeRational
            Case PropertyTagId.GpsSpeedRef: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsSpeed: GetPropertyType = TypeRational
            Case PropertyTagId.GpsTrackRef: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsTrack: GetPropertyType = TypeRational
            Case PropertyTagId.GpsImgDirRef: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsImgDir: GetPropertyType = TypeRational
            Case PropertyTagId.GpsMapDatum: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsDestLatRef: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsDestLat: GetPropertyType = TypeRational
            Case PropertyTagId.GpsDestLongRef: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsDestLong: GetPropertyType = TypeRational
            Case PropertyTagId.GpsDestBearRef: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsDestBear: GetPropertyType = TypeRational
            Case PropertyTagId.GpsDestDistRef: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsDestDist: GetPropertyType = TypeRational
            Case PropertyTagId.NewSubfileType: GetPropertyType = TypeLong
            Case PropertyTagId.SubfileType: GetPropertyType = TypeShort
            Case PropertyTagId.ImageWidth: GetPropertyType = TypeLong 'Short or Long?
            Case PropertyTagId.ImageHeight: GetPropertyType = TypeLong
            Case PropertyTagId.BitsPerSample: GetPropertyType = TypeShort
            Case PropertyTagId.Compression: GetPropertyType = TypeShort
            Case PropertyTagId.PhotometricInterp: GetPropertyType = TypeShort
            Case PropertyTagId.ThreshHolding: GetPropertyType = TypeShort
            Case PropertyTagId.CellWidth: GetPropertyType = TypeShort
            Case PropertyTagId.CellHeight: GetPropertyType = TypeShort
            Case PropertyTagId.FillOrder: GetPropertyType = TypeShort
            Case PropertyTagId.DocumentName: GetPropertyType = TypeASCII
            Case PropertyTagId.ImageDescription: GetPropertyType = TypeASCII
            Case PropertyTagId.EquipMake: GetPropertyType = TypeASCII
            Case PropertyTagId.EquipModel: GetPropertyType = TypeASCII
            Case PropertyTagId.StripOffsets: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.Orientation: GetPropertyType = TypeShort
            Case PropertyTagId.SamplesPerPixel: GetPropertyType = TypeShort
            Case PropertyTagId.RowsPerStrip: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.StripBytesCount: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.MinSampleValue: GetPropertyType = TypeShort
            Case PropertyTagId.MaxSampleValue: GetPropertyType = TypeShort
            Case PropertyTagId.XResolution: GetPropertyType = TypeRational
            Case PropertyTagId.YResolution: GetPropertyType = TypeRational
            Case PropertyTagId.PlanarConfig: GetPropertyType = TypeShort
            Case PropertyTagId.PageName: GetPropertyType = TypeASCII
            Case PropertyTagId.XPosition: GetPropertyType = TypeRational
            Case PropertyTagId.YPosition: GetPropertyType = TypeRational
            Case PropertyTagId.FreeOffset: GetPropertyType = TypeLong
            Case PropertyTagId.FreeByteCounts: GetPropertyType = TypeLong
            Case PropertyTagId.GrayResponseUnit: GetPropertyType = TypeShort
            Case PropertyTagId.GrayResponseCurve: GetPropertyType = TypeShort
            Case PropertyTagId.T4Option: GetPropertyType = TypeLong
            Case PropertyTagId.T6Option: GetPropertyType = TypeLong
            Case PropertyTagId.ResolutionUnit: GetPropertyType = TypeShort
            Case PropertyTagId.PageNumber: GetPropertyType = TypeShort
            Case PropertyTagId.TransferFunction: GetPropertyType = TypeShort
            Case PropertyTagId.SoftwareUsed: GetPropertyType = TypeASCII
            Case PropertyTagId.DateTime: GetPropertyType = TypeASCII
            Case PropertyTagId.Artist: GetPropertyType = TypeASCII
            Case PropertyTagId.HostComputer: GetPropertyType = TypeASCII
            Case PropertyTagId.Predictor: GetPropertyType = TypeShort
            Case PropertyTagId.WhitePoint: GetPropertyType = TypeRational
            Case PropertyTagId.PrimaryChromaticities: GetPropertyType = TypeRational
            Case PropertyTagId.ColorMap: GetPropertyType = TypeShort
            Case PropertyTagId.HalftoneHints: GetPropertyType = TypeShort
            Case PropertyTagId.TileWidth: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.TileLength: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.TileOffset: GetPropertyType = TypeLong
            Case PropertyTagId.TileByteCounts: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.InkSet: GetPropertyType = TypeShort
            Case PropertyTagId.InkNames: GetPropertyType = TypeASCII
            Case PropertyTagId.NumberOfInks: GetPropertyType = TypeShort
            Case PropertyTagId.DotRange: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.TargetPrinter: GetPropertyType = TypeASCII
            Case PropertyTagId.ExtraSamples: GetPropertyType = TypeShort
            Case PropertyTagId.SampleFormat: GetPropertyType = TypeShort
            Case PropertyTagId.TransferRange: GetPropertyType = TypeShort
            Case PropertyTagId.JPEGProc: GetPropertyType = TypeShort
            Case PropertyTagId.JPEGInterFormat: GetPropertyType = TypeLong
            Case PropertyTagId.JPEGInterLength: GetPropertyType = TypeLong
            Case PropertyTagId.JPEGRestartInterval: GetPropertyType = TypeShort
            Case PropertyTagId.JPEGLosslessPredictors: GetPropertyType = TypeShort
            Case PropertyTagId.JPEGPointTransforms: GetPropertyType = TypeShort
            Case PropertyTagId.JPEGQTables: GetPropertyType = TypeLong
            Case PropertyTagId.JPEGDCTables: GetPropertyType = TypeLong
            Case PropertyTagId.JPEGACTables: GetPropertyType = TypeLong
            Case PropertyTagId.YCbCrCoefficients: GetPropertyType = TypeRational
            Case PropertyTagId.YCbCrSubsampling: GetPropertyType = TypeShort
            Case PropertyTagId.YCbCrPositioning: GetPropertyType = TypeShort
            Case PropertyTagId.REFBlackWhite: GetPropertyType = TypeRational
            Case PropertyTagId.Gamma: GetPropertyType = TypeRational
            Case PropertyTagId.ICCProfileDescriptor: GetPropertyType = TypeASCII
            Case PropertyTagId.SRGBRenderingIntent: GetPropertyType = TypeByte
            Case PropertyTagId.ImageTitle: GetPropertyType = TypeASCII
            Case PropertyTagId.ResolutionXUnit: GetPropertyType = TypeShort
            Case PropertyTagId.ResolutionYUnit: GetPropertyType = TypeShort
            Case PropertyTagId.ResolutionXLengthUnit: GetPropertyType = TypeShort
            Case PropertyTagId.ResolutionYLengthUnit: GetPropertyType = TypeShort
            Case PropertyTagId.PrintFlags: GetPropertyType = TypeASCII
            Case PropertyTagId.PrintFlagsVersion: GetPropertyType = TypeShort
            Case PropertyTagId.PrintFlagsCrop: GetPropertyType = TypeByte
            Case PropertyTagId.PrintFlagsBleedWidth: GetPropertyType = TypeLong
            Case PropertyTagId.PrintFlagsBleedWidthScale: GetPropertyType = TypeShort
            Case PropertyTagId.HalftoneLPI: GetPropertyType = TypeRational
            Case PropertyTagId.HalftoneLPIUnit: GetPropertyType = TypeShort
            Case PropertyTagId.HalftoneDegree: GetPropertyType = TypeRational
            Case PropertyTagId.HalftoneShape: GetPropertyType = TypeShort
            Case PropertyTagId.HalftoneMisc: GetPropertyType = TypeLong
            Case PropertyTagId.HalftoneScreen: GetPropertyType = TypeByte
            Case PropertyTagId.JPEGQuality: GetPropertyType = TypeShort
            Case PropertyTagId.GridSize: GetPropertyType = TypeUndefined
            Case PropertyTagId.ThumbnailFormat: GetPropertyType = TypeLong
            Case PropertyTagId.ThumbnailWidth: GetPropertyType = TypeLong
            Case PropertyTagId.ThumbnailHeight: GetPropertyType = TypeLong
            Case PropertyTagId.ThumbnailColorDepth: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailPlanes: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailRawBytes: GetPropertyType = TypeLong
            Case PropertyTagId.ThumbnailSize: GetPropertyType = TypeLong
            Case PropertyTagId.ThumbnailCompressedSize: GetPropertyType = TypeLong
            Case PropertyTagId.ColorTransferFunction: GetPropertyType = TypeUndefined
            Case PropertyTagId.ThumbnailData: GetPropertyType = TypeByte
            Case PropertyTagId.ThumbnailImageWidth: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.ThumbnailImageHeight: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.ThumbnailBitsPerSample: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailCompression: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailPhotometricInterp: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailImageDescription: GetPropertyType = TypeASCII
            Case PropertyTagId.ThumbnailEquipMake: GetPropertyType = TypeASCII
            Case PropertyTagId.ThumbnailEquipModel: GetPropertyType = TypeASCII
            Case PropertyTagId.ThumbnailStripOffsets: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.ThumbnailOrientation: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailSamplesPerPixel: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailRowsPerStrip: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.ThumbnailStripBytesCount: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.ThumbnailResolutionX: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailResolutionY: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailPlanarConfig: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailResolutionUnit: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailTransferFunction: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailSoftwareUsed: GetPropertyType = TypeASCII
            Case PropertyTagId.ThumbnailDateTime: GetPropertyType = TypeASCII
            Case PropertyTagId.ThumbnailArtist: GetPropertyType = TypeASCII
            Case PropertyTagId.ThumbnailWhitePoint: GetPropertyType = TypeRational
            Case PropertyTagId.ThumbnailPrimaryChromaticities: GetPropertyType = TypeRational
            Case PropertyTagId.ThumbnailYCbCrCoefficients: GetPropertyType = TypeRational
            Case PropertyTagId.ThumbnailYCbCrSubsampling: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailYCbCrPositioning: GetPropertyType = TypeShort
            Case PropertyTagId.ThumbnailRefBlackWhite: GetPropertyType = TypeRational
            Case PropertyTagId.ThumbnailCopyRight: GetPropertyType = TypeASCII
            Case PropertyTagId.LuminanceTable: GetPropertyType = TypeShort
            Case PropertyTagId.ChrominanceTable: GetPropertyType = TypeShort
            Case PropertyTagId.FrameDelay: GetPropertyType = TypeLong
            Case PropertyTagId.LoopCount: GetPropertyType = TypeShort
            Case PropertyTagId.GlobalPalette: GetPropertyType = TypeByte
            Case PropertyTagId.IndexBackground: GetPropertyType = TypeByte
            Case PropertyTagId.IndexTransparent: GetPropertyType = TypeByte
            Case PropertyTagId.PixelUnit: GetPropertyType = TypeByte
            Case PropertyTagId.PixelPerUnitX: GetPropertyType = TypeLong
            Case PropertyTagId.PixelPerUnitY: GetPropertyType = TypeLong
            Case PropertyTagId.PaletteHistogram: GetPropertyType = TypeByte
            Case PropertyTagId.Copyright: GetPropertyType = TypeASCII
            Case PropertyTagId.ExifExposureTime: GetPropertyType = TypeRational
            Case PropertyTagId.ExifFNumber: GetPropertyType = TypeRational
            Case PropertyTagId.ExifIFD: GetPropertyType = TypeLong
            Case PropertyTagId.ICCProfile: GetPropertyType = TypeByte
            Case PropertyTagId.ExifExposureProg: GetPropertyType = TypeShort
            Case PropertyTagId.ExifSpectralSense: GetPropertyType = TypeASCII
            Case PropertyTagId.GpsIFD: GetPropertyType = TypeLong
            Case PropertyTagId.ExifISOSpeed: GetPropertyType = TypeShort
            Case PropertyTagId.ExifOECF: GetPropertyType = TypeUndefined
            Case PropertyTagId.ExifVer: GetPropertyType = TypeUndefined
            Case PropertyTagId.ExifDTOrig: GetPropertyType = TypeASCII
            Case PropertyTagId.ExifDTDigitized: GetPropertyType = TypeASCII
            Case PropertyTagId.ExifCompConfig: GetPropertyType = TypeUndefined
            Case PropertyTagId.ExifCompBPP: GetPropertyType = TypeRational
            Case PropertyTagId.ExifShutterSpeed: GetPropertyType = TypeSRational
            Case PropertyTagId.ExifAperture: GetPropertyType = TypeRational
            Case PropertyTagId.ExifBrightness: GetPropertyType = TypeSRational
            Case PropertyTagId.ExifExposureBias: GetPropertyType = TypeSRational
            Case PropertyTagId.ExifMaxAperture: GetPropertyType = GetPropertyType
            Case PropertyTagId.ExifSubjectDist: GetPropertyType = TypeRational
            Case PropertyTagId.ExifMeteringMode: GetPropertyType = TypeShort
            Case PropertyTagId.ExifLightSource: GetPropertyType = TypeShort
            Case PropertyTagId.ExifFlash: GetPropertyType = TypeShort
            Case PropertyTagId.ExifFocalLength: GetPropertyType = TypeRational
            Case PropertyTagId.ExifMakerNote: GetPropertyType = TypeUndefined
            Case PropertyTagId.ExifUserComment: GetPropertyType = TypeUndefined
            Case PropertyTagId.ExifDTSubsec: GetPropertyType = TypeASCII
            Case PropertyTagId.ExifDTOrigSS: GetPropertyType = TypeASCII
            Case PropertyTagId.ExifDTDigSS: GetPropertyType = TypeASCII
            Case PropertyTagId.ExifFPXVer: GetPropertyType = TypeUndefined
            Case PropertyTagId.ExifColorSpace: GetPropertyType = TypeShort
            Case PropertyTagId.ExifPixXDim: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.ExifPixYDim: GetPropertyType = TypeLong 'Short or Long
            Case PropertyTagId.ExifRelatedWav: GetPropertyType = TypeASCII
            Case PropertyTagId.ExifInterop: GetPropertyType = TypeLong
            Case PropertyTagId.ExifFlashEnergy: GetPropertyType = TypeRational
            Case PropertyTagId.ExifSpatialFR: GetPropertyType = TypeUndefined
            Case PropertyTagId.ExifFocalXRes: GetPropertyType = TypeRational
            Case PropertyTagId.ExifFocalYRes: GetPropertyType = TypeRational
            Case PropertyTagId.ExifFocalResUnit: GetPropertyType = TypeShort
            Case PropertyTagId.ExifSubjectLoc: GetPropertyType = TypeShort
            Case PropertyTagId.ExifExposureIndex: GetPropertyType = TypeRational
            Case PropertyTagId.ExifSensingMethod: GetPropertyType = TypeShort
            Case PropertyTagId.ExifFileSource: GetPropertyType = TypeUndefined
            Case PropertyTagId.ExifSceneType: GetPropertyType = TypeUndefined
            Case PropertyTagId.ExifCfaPattern: GetPropertyType = TypeUndefined
            Case Else: GetPropertyType = TypeUndefined
        End Select
    End Function
    
