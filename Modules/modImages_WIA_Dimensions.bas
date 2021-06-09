Attribute VB_Name = "modImages_WIA_Dimensions"
    Option Explicit
    
    Private Type ImgageInfo
       Height                    As Long
       width                     As Long
       fileExtension             As String
       HorizontalResolution      As Double
       VerticalResolution        As Double
       PixelDepth As Long
    End Type
    
    Dim img                    As ImgageInfo
    
    '---------------------------------------------------------------------------------------
    ' Procedure : WIA_GetImgDimensions
    ' Author    : Daniel Pineault, CARDA Consultants Inc.
    ' Website   : http://www.cardaconsultants.com
    ' Purpose   : Retrieve various properties (dimensions, extension, resolution) of an image
    ' Copyright : The following is release as Attribution-ShareAlike 4.0 International
    '             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
    ' Req'd Refs: Uses Late Binding, so none required
    '
    ' Input Variables:
    ' ~~~~~~~~~~~~~~~~
    ' sFile     : Fully qualified path, filename and extension of the image file to check
    '
    ' Usage:
    ' ~~~~~~
    ' Call WIA_GetImgDimensions(C:\Tmp\database.png )
    ' Debug.Print sFile, "Width: " & Img.Width, "Height: " & Img.Height, "FileExtension: " & _
    '                  Img.FileExtension, "HorizontalResolution: " & Img.HorizontalResolution, _
    '                  "VerticalResolution: " & Img.VerticalResolution, _
    '                  "PixelDepth: " & Img.PixelDepth
    '
    ' Revision History:
    ' Rev       Date(yyyy/mm/dd)        Description
    ' **************************************************************************************
    ' 1         2018-10-23              Initial Release
    '---------------------------------------------------------------------------------------
    Function WIA_GetImgDimensions(ByVal sFile As String) As Boolean
       ' For a complete listing of available WIA ImageFile properties
       ' Ref: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-imagefile
       On Error GoTo Error_Handler
       Dim oWIA                  As Object
    
       Set oWIA = CreateObject("WIA.ImageFile")
       oWIA.LoadFile sFile
       img.width = oWIA.width
       img.Height = oWIA.Height
       img.fileExtension = oWIA.fileExtension
       img.HorizontalResolution = oWIA.HorizontalResolution
       img.VerticalResolution = oWIA.VerticalResolution
       img.PixelDepth = oWIA.PixelDepth
    
Error_Handler_Exit:
       On Error Resume Next
       If Not oWIA Is Nothing Then Set oWIA = Nothing
       Exit Function
    
Error_Handler:
       MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
              "Error Number: " & err.Number & vbCrLf & _
              "Error Source: WIA_GetImgDimensions" & vbCrLf & _
              "Error Description: " & err.Descappliription & _
              Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
              , vbOKOnly + vbCritical, "An Error has Occured!"
       Resume Error_Handler_Exit
    End Function
    
    Private Sub TestDimensionsProcedure()
       Dim sFiles                 As Variant
       Dim sFile                  As Variant
    
       sFiles = Array(Filename1, Filename2, Filename3)
                     
       For Each sFile In sFiles
           
           WIA_GetImgDimensions sFile
           Debug.Print sFile
           Debug.Print "Width: " & img.width & vbTab & "Height: " & img.Height
           Debug.Print "FileExtension: " & img.fileExtension
           Debug.Print "HorizontalResolution: " & img.HorizontalResolution & vbTab & "VerticalResolution: " & img.VerticalResolution
           Debug.Print "PixelDepth: " & img.PixelDepth & vbNewLine & vbNewLine
    
       Next
    
    End Sub
