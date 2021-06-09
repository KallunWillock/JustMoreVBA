Attribute VB_Name = "modImages_WIA_Convert"
    Enum wiaFormat
        BMP = 0
        GIF = 1
        JPEG = 2
        PNG = 3
        TIFF = 4
    End Enum
     
    '---------------------------------------------------------------------------------------
    ' Procedure : WIA_ConvertImage
    ' Author    : Daniel Pineault, CARDA Consultants Inc.
    ' Website   : http://www.cardaconsultants.com
    ' Purpose   : Convert an image's format using WIA
    ' Copyright : The following is release as Attribution-ShareAlike 4.0 International
    '             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
    ' Req'd Refs: Uses Late Binding, so none required
    '
    ' Windows Image Acquisition (WIA)
    '             https://msdn.microsoft.com/en-us/library/windows/desktop/ms630368(v=vs.85).aspx
    '
    ' Input Variables:
    ' ~~~~~~~~~~~~~~~~
    ' sInitialImage : Fully qualified path and filename of the original image to resize
    ' sOutputImage  : Fully qualified path and filename of where to save the new image
    ' lFormat       : Format to convert the image into
    ' lQuality      : Quality level to be used for the conversion process (1-100)
    '
    ' Usage:
    ' ~~~~~~
    ' Call WIA_ConvertImage("C:\Users\Public\Pictures\Sample Pictures\Chrysanthemum.jpg", _
    '                       "C:\Users\MyUser\Desktop\Chrysanthemum_2.jpg", _
    '                       JPEG)
    '
    ' Revision History:
    ' Rev       Date(yyyy/mm/dd)        Description
    ' **************************************************************************************
    ' 1         2017-01-18              Initial Release
    ' 2         2018-09-20              Updated Copyright
    '---------------------------------------------------------------------------------------
    Function WIA_ConvertImage(sInitialImage As String, _
                                     sOutputImage As String, _
                                     lFormat As wiaFormat, _
                                     Optional lQuality As Long = 85) As Boolean
        On Error GoTo Error_Handler
        Dim oWIA                  As Object    'WIA.ImageFile
        Dim oIP                   As Object    'ImageProcess
        Dim sFormatID             As String
        Dim sExt                  As String
     
        'Convert our Enum over to the proper value used by WIA
        Select Case lFormat
            Case 0
                sFormatID = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
                sExt = "BMP"
            Case 1
                sFormatID = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
                sExt = "GIF"
            Case 2
                sFormatID = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
                sExt = "JPEG"
            Case 3
                sFormatID = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
                sExt = "PNG"
            Case 4
                sFormatID = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
                sExt = "TIFF"
        End Select
     
        If lQuality > 100 Then lQuality = 100
     
        'Should check if the output file already exists and if so,
        'prompt the user to overwrite it or not
     
        Set oWIA = CreateObject("WIA.ImageFile")
        Set oIP = CreateObject("WIA.ImageProcess")
     
        oIP.Filters.ADD oIP.FilterInfos("Convert").FilterID
        oIP.Filters(1).Properties("FormatID") = sFormatID
        oIP.Filters(1).Properties("Quality") = lQuality
     
        oWIA.LoadFile sInitialImage
        Set oWIA = oIP.Apply(oWIA)
        
        'Overide the specified ext with the appropriate one for the choosen format
        oWIA.SaveFile Left(sOutputImage, InStrRev(sOutputImage, ".")) & LCase(sExt)
        WIA_ConvertImage = True
     
Error_Handler_Exit:
        On Error Resume Next
        If Not oIP Is Nothing Then Set oIP = Nothing
        If Not oWIA Is Nothing Then Set oWIA = Nothing
        Exit Function
     
Error_Handler:
        MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
               "Error Number: " & err.Number & vbCrLf & _
               "Error Source: WIA_ConvertImage" & vbCrLf & _
               "Error Description: " & err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occured!"
        Resume Error_Handler_Exit
    End Function
    
