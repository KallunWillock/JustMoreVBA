Attribute VB_Name = "modImages_PastePicture"
'@Lang VBA
    '***************************************************************************
    '* AUTHOR:          STEPHEN BULLEN, Business Modelling Solutions Ltd.   (Stephen@BMSLtd.co.uk | https://web.archive.org/web/*/https://www.bmsltd.co.uk/)
    '                   15 Nov 1998 - Created this module comprising
    '                   KALLUN WILLOCK  (https://github.com/kallunwillock/)
    '                   12 Sep 2022 - Updated API declaration and related data types for 64bit compatibility
    '                   23 Jun 2023 - Updated code - streamlined per xlBangAnywhere
    '* URL:             https://github.dev/KallunWillock/JustMoreVBA/blob/main/Images/modImages_PastePicture.bas
    '* DESCRIPTION:     Creates a standard Picture object from the contents of the clipboard. This stdole.stdPicture can then be assigned to (for example) 
    '                   an Image control on a userform. The code in this module has been derived from a number of sources discovered on MSDN.
    '* NOTES:           23 June 2023 - A number of necessary updates / stylistic adjustments to the Stephen Bullen's code. 
    '                   (1) Replaced xlBitmap and xlPicture constants - these are meaningless in the rest of the VBA ecosystem.
    '                   (2) API Declarations - various corrections
    '                   (3) Shortened code / removed and rewrote comments
    '* USAGE:           Set Image1.Picture = PastePicture(vbaBitmap)
    '* PROCEDURES:
    '    PastePicture   Public function used to assign image on clipboard to stdOle.stdPicture object
    '    CreatePicture  Private function to convert a bitmap or metafile handle to an OLE reference
    '    fnOLEError     Private function that returns the error text for an OLE error code
    '***************************************************************************
    
    Option Explicit
    Option Compare Text
   
    Public Enum vbaPictureFormat
        vbaBitmap           = 2                 ' Mirrors xlBitmap - a bitmap (.bmp, .jpg, .gif).
        vbaPicture	        = -4147	            ' Mirrors xlPicture - a drawn picture (.png, .wmf, .mix).
    End Enum
        
    Private Type GUID                           ' Declare a UDT to store a GUID for the IPicture OLE Interface
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
    End Type
    
    'Declare a UDT to store the bitmap information
    Private Type uPicDesc
        Size As Long
        Type As Long
        hPic As LongPtr
        hPal As LongPtr
    End Type

    #If VBA7 Then                    
        Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long         ' Does the clipboard contain a bitmap/metafile?
        Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long                      ' Open the clipboard to read
        Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr                ' Get a pointer to the bitmap/metafile
        Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long                                          ' Close the clipboard
        ' Convert the handle into an OLE IPicture interface.
        Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleAut32.dll" (ByRef lpPictDesc As uPicDesc, ByRef riid As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
        ' Create copy of the metafile/bitmap
        Private Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
        Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
    #Else
        Private Enum LongPtr
            [_]
        End Enum
        Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
        Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
        Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
        Private Declare Function CloseClipboard Lib "user32" () As Long
        Private Declare Function OleCreatePictureIndirect Lib "oleAut32.dll" (ByRef lpPictDesc As uPicDesc, ByRef riid As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
        Private Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
        Private Declare Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
    #End If

    Private Const CF_BITMAP             As Long = 2
    Private Const CF_PALETTE            As Long = 9
    Private Const CF_ENHMETAFILE        As Long = 14
    Private Const IMAGE_BITMAP          As Long = 0
    Private Const LR_COPYRETURNORG      As Long = &H4
    
    Public Function PastePicture(Optional Byval SelectedPictureType As vbaPictureFormat = vbaBitmap) As IPicture
        ''' Purpose:    Get a Picture object showing whatever's on the clipboard.
        ''' Arguments:  SelectedPictureType - either vbaBitmap (Default) or vbaPicture
        ''' --------------------------------------------------------------------------
        ''' 30 Oct 98   Stephen Bullen      Created
        ''' 15 Nov 98   Stephen Bullen      Updated to create our own copies of the clipboard images
        ''' 12 Sep 22   Kallun Willock      Updated for 32/64bit compatibility
        ''' 23 Jun 23   Kallun Willock      Updated per xlBangAnywhere - Replaced Excel constants
        Dim hPicAvail As LongPtr, hPtr As LongPtr, hPal As LongPtr, hCopy As LongPtr
        Dim Result As Long, lPicType As Long
        lPicType = IIf(SelectedPictureType = vbaBitmap, CF_BITMAP, CF_ENHMETAFILE)
        hPicAvail = IsClipboardFormatAvailable(lPicType)                                ' Check if the clipboard contains the required format
        If hPicAvail <> 0 Then
            Result = OpenClipboard(0&)                                                  ' Get access to the clipboard
            If Result <> 0 Then
                hPtr = GetClipboardData(lPicType)                                       ' Get a handle to the image data
                If lPicType = CF_BITMAP Then                                            ' Create our own copy of the image on the clipboard, in the appropriate format.
                    hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
                Else
                    hCopy = CopyEnhMetaFile(hPtr, vbNullString)
                End If
                Result = CloseClipboard                                                 ' Release the clipboard to other programs
                If hPtr <> 0 Then Set PastePicture = CreatePicture(hCopy, lPicType)     ' If function returns a pointer, convert it into an IPicture object and return it
            End If
        End If
    End Function
    
    Private Function CreatePicture(ByVal hPic As LongPtr, Optional ByVal lPicType) As IPicture
        ''' Purpose:    Converts an image handle into a Picture object. The IPicture Interface 
        '''             requires a reference to the "OLE Automation" type library
        ''' --------------------------------------------------------------------------
        ''' 30 Oct 98  Stephen Bullen      Created
        ''' 12 Sep 22  Kallun Willock      Updated for both 32/64bit compatibility        
        Dim r As Long, uPicinfo As uPicDesc, IID_IDispatch As GUID, IPic As IPicture
        
        Const PICTYPE_BITMAP = 1                ' OLE Picture types
        Const PICTYPE_ENHMETAFILE = 4
    
        With IID_IDispatch                      ' Create the IPicture Interface GUID
            .Data1 = &H7BF80980
            .Data2 = &HBF32
            .Data3 = &H101A
            .Data4(0) = &H8B
            .Data4(1) = &HBB
            .Data4(2) = &H0
            .Data4(3) = &HAA
            .Data4(4) = &H0
            .Data4(5) = &H30
            .Data4(6) = &HC
            .Data4(7) = &HAB
        End With
        
        With uPicinfo
            .Size = LenB(uPicinfo)                                                  ' Size of struct
            .Type = IIf(lPicType = CF_BITMAP, PICTYPE_BITMAP, PICTYPE_ENHMETAFILE)  ' Type of picture
            .hPic = hPic                                                            ' Handle to image
            .hPal = 0
        End With
        
        r = OleCreatePictureIndirect(uPicinfo, IID_IDispatch, True, IPic)           ' Create and return the Picture object.
        If r <> 0 Then Debug.Print "CreatePicture: " & fnOLEError(r)
        Set CreatePicture = IPic                                                    
    End Function
        
    Private Function fnOLEError(ByVal lErrNum As Long) As String
        ''' Purpose:    Gets the message text for standard OLE errors
        ''' Arguments:  OLECreatePictureIndirect return value
        ''' --------------------------------------------------------------------------
        ''' 30 Oct 98   Stephen Bullen      Created
        ''' 23 Jun 23   Kallun Willock      Removed constants
        Select Case lErrNum
            Case &H0: fnOLEError = "Success"                                ' S_OK
            Case &H80004001: fnOLEError = "Not Implemented"                 ' E_NOTIMPL
            case &H80004002: fnOLEError = "No Interface"                    ' E_NOINTERFACE
            Case &H80004003: fnOLEError = "Invalid Pointer"                 ' E_POINTER
            Case &H80004004: fnOLEError = "Aborted"                         ' E_ABORT
            Case &H80004005: fnOLEError = "General Failure"                 ' E_FAIL
            Case &H80070005: fnOLEError = "Access Denied"                   ' E_ACCESSDENIED
            Case &H80070006: fnOLEError = "Bad/Missing Handle"              ' E_HANDLE
            Case &H8007000E: fnOLEError = "Out of Memory"                   ' E_OUTOFMEMORY
            Case &H80070057: fnOLEError = "Invalid Argument"                ' E_INVALIDARG
            Case &H8000FFFF: fnOLEError = "Unknown Error"                   ' E_UNEXPECTED
        End Select
    End Function
