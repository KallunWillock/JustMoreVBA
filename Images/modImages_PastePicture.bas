    '***************************************************************************
    '*
    '* MODULE NAME:     Paste Picture
    '* AUTHOR & DATE:   STEPHEN BULLEN, Business Modelling Solutions Ltd.
    '*                  15 November 1998
    '*
    '*                  12 September 2022 - Updated API declaration and related data types for 64bit compatibility
    '*
    '* CONTACT:         Stephen@BMSLtd.co.uk
    '* WEB SITE:        http://www.BMSLtd.co.uk
    '*
    '* DESCRIPTION:     Creates a standard Picture object from whatever is on the clipboard.
    '*                  This object can then be assigned to (for example) and Image control
    '*                  on a userform.  The PastePicture function takes an optional argument of
    '*                  the picture type - xlBitmap or xlPicture.
    '*
    '*                  The code requires a reference to the "OLE Automation" type library
    '*
    '*                  The code in this module has been derived from a number of sources
    '*                  discovered on MSDN.
    '*
    '*                  To use it, just copy this module into your project, then you can use:
    ''
    '*                      Set Image1.Picture = PastePicture(xlPicture)
    '*
    '*                  to paste a picture of whatever is on the clipboard into a standard image control.
    '*
    '* PROCEDURES:
    '*   PastePicture   The entry point for the routine
    '*   CreatePicture  Private function to convert a bitmap or metafile handle to an OLE reference
    '*   fnOLEError     Get the error text for an OLE error code
    '***************************************************************************
    
    Option Explicit
    Option Compare Text
    
    '''Windows API Function Declarations

    #If VBA7 Then
    
        'Declare a UDT to store the bitmap information
        Private Type uPicDesc
            Size As Long
            Type As Long
            hPic As LongPtr
            hPal As LongPtr
        End Type
        
        'Does the clipboard contain a bitmap/metafile?
        Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
        
        'Open the clipboard to read
        Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
        
        'Get a pointer to the bitmap/metafile
        Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
        
        'Close the clipboard
        Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
        
        'Convert the handle into an OLE IPicture interface.
        Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleAut32.dll" (ByRef lpPictDesc As uPicDesc, ByRef riid As GUID, ByVal fPictureOwnsHandle As LongPtr, ByRef IPic As IPicture) As Long
        
        'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
        Private Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
        
        'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
        Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
        
        'Some pointers
        Dim H As LongPtr, hPicAvail As LongPtr, hPtr As LongPtr, hPal As LongPtr, lPicType As Long, hCopy As LongPtr
        
    #Else
        'Declare a UDT to store the bitmap information
        Private Type uPicDesc
            Size As Long
            Type As Long
            hPic As Long
            hPal As Long
        End Type
        
        'Does the clipboard contain a bitmap/metafile?
        Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
        
        'Open the clipboard to read
        Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
        
        'Get a pointer to the bitmap/metafile
        Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
        
        'Close the clipboard
        Private Declare Function CloseClipboard Lib "user32" () As Long
        
        'Convert the handle into an OLE IPicture interface.
        Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
        'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
        Private Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
        
        'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
        Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
        
        'Some pointers
        Dim H As Long, hPicAvail As Long, hPtr As Long, hPal As Long, lPicType As Long, hCopy As Long
    #End If
    
    'Declare a UDT to store a GUID for the IPicture OLE Interface
    Private Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
    End Type
    
    'The API format types we're interested in
    Private Const CF_BITMAP             As Long = 2
    Private Const CF_PALETTE            As Long = 9
    Private Const CF_ENHMETAFILE        As Long = 14
    Private Const IMAGE_BITMAP          As Long = 0
    Private Const LR_COPYRETURNORG      As Long = &H4
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Subroutine: PastePicture
    '''
    ''' Purpose:    Get a Picture object showing whatever's on the clipboard.
    '''
    ''' Arguments:  lXlPicType - The type of picture to create.  Can be one of:
    '''                          xlPicture to create a metafile (default)
    '''                          xlBitmap to create a bitmap
    '''
    ''' Date        Developer           Action
    ''' --------------------------------------------------------------------------
    ''' 30 Oct 98   Stephen Bullen      Created
    ''' 15 Nov 98   Stephen Bullen      Updated to create our own copies of the clipboard images
    ''' 12 Sep 22   Kallun Willock      Updated for both 32/64bit compatibility
    '''
    
    Function PastePicture(Optional lXlPicType As XlCopyPictureFormat = xlPicture) As IPicture
        
        'Convert the type of picture requested from the xl constant to the API constant
        lPicType = IIf(lXlPicType = xlBitmap, CF_BITMAP, CF_ENHMETAFILE)
        
        'Check if the clipboard contains the required format
        hPicAvail = IsClipboardFormatAvailable(lPicType)
        
        If hPicAvail <> 0 Then
            
            'Get access to the clipboard
            H = OpenClipboard(0&)
        
            If H > 0 Then
                'Get a handle to the image data
                hPtr = GetClipboardData(lPicType)
        
                'Create our own copy of the image on the clipboard, in the appropriate format.
                If lPicType = CF_BITMAP Then
                    hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
                Else
                    hCopy = CopyEnhMetaFile(hPtr, vbNullString)
                End If
        
                'Release the clipboard to other programs
                H = CloseClipboard
        
                'If we got a handle to the image, convert it into a Picture object and return it
                If hPtr <> 0 Then Set PastePicture = CreatePicture(hCopy, 0, lPicType)
            End If
        End If
    
    End Function
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Subroutine: CreatePicture
    '''
    ''' Purpose:    Converts a image (and palette) handle into a Picture object.
    '''
    '''             Requires a reference to the "OLE Automation" type library
    '''
    ''' Arguments:  None
    '''
    ''' Date        Developer           Action
    ''' --------------------------------------------------------------------------
    ''' 30 Oct 98  Stephen Bullen      Created
    ''' 12 Sep 22  Kallun Willock      Updated for both 32/64bit compatibility
    '''
    #If VBA7 Then
        Private Function CreatePicture(ByVal hPic As LongPtr, ByVal hPal As LongPtr, ByVal lPicType) As IPicture
    #Else
        Private Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, ByVal lPicType) As IPicture
    #End If
        
        ' IPicture requires a reference to "OLE Automation"
        Dim r As Long, uPicinfo As uPicDesc, IID_IDispatch As GUID, IPic As IPicture
        
        'OLE Picture types
        Const PICTYPE_BITMAP = 1
        Const PICTYPE_ENHMETAFILE = 4
        
        ' Create the Interface GUID (for the IPicture interface)
        With IID_IDispatch
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
        
        ' Fill uPicInfo with necessary parts.
        With uPicinfo
            .Size = LenB(uPicinfo)                                                  ' Length of structure.
            .Type = IIf(lPicType = CF_BITMAP, PICTYPE_BITMAP, PICTYPE_ENHMETAFILE)  ' Type of Picture
            .hPic = hPic                                                            ' Handle to image.
            .hPal = IIf(lPicType = CF_BITMAP, hPal, 0)                              ' Handle to palette (if bitmap).
        End With
        
        ' Create the Picture object.
        r = OleCreatePictureIndirect(uPicinfo, IID_IDispatch, True, IPic)
        
        ' If an error occured, show the description
        If r <> 0 Then Debug.Print "Create Picture: " & fnOLEError(r)
        
        ' Return the new Picture object.
        Set CreatePicture = IPic
        
    End Function
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Subroutine: fnOLEError
    '''
    ''' Purpose:    Gets the message text for standard OLE errors
    '''
    ''' Arguments:  None
    '''
    ''' Date        Developer           Action
    ''' --------------------------------------------------------------------------
    ''' 30 Oct 98   Stephen Bullen      Created
    '''
        
    Private Function fnOLEError(lErrNum As Long) As String
        
        'OLECreatePictureIndirect return values
        Const E_ABORT                   As Long = &H80004004
        Const E_ACCESSDENIED            As Long = &H80070005
        Const E_FAIL                    As Long = &H80004005
        Const E_HANDLE                  As Long = &H80070006
        Const E_INVALIDARG              As Long = &H80070057
        Const E_NOINTERFACE             As Long = &H80004002
        Const E_NOTIMPL                 As Long = &H80004001
        Const E_OUTOFMEMORY             As Long = &H8007000E
        Const E_POINTER                 As Long = &H80004003
        Const E_UNEXPECTED              As Long = &H8000FFFF
        Const S_OK                      As Long = &H0
        
        Select Case lErrNum
            Case E_ABORT:               fnOLEError = " Aborted"
            Case E_ACCESSDENIED:        fnOLEError = " Access Denied"
            Case E_FAIL:                fnOLEError = " General Failure"
            Case E_HANDLE:              fnOLEError = " Bad/Missing Handle"
            Case E_INVALIDARG:          fnOLEError = " Invalid Argument"
            Case E_NOINTERFACE:         fnOLEError = " No Interface"
            Case E_NOTIMPL:             fnOLEError = " Not Implemented"
            Case E_OUTOFMEMORY:         fnOLEError = " Out of Memory"
            Case E_POINTER:             fnOLEError = " Invalid Pointer"
            Case E_UNEXPECTED:          fnOLEError = " Unknown Error"
            Case S_OK:                  fnOLEError = " Success!"
        End Select
        
    End Function
