Attribute VB_Name = "modForum_GoogleSheets_Image"
                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||   FORUMS - GOOGLE SHEETS - IMAGE      ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
                                                                                                                                                                                            ' _
    AUTHOR:   Dan_W / Kallun Willock                                                                                                                                                        ' _
    PURPOSE:  Google Sheets function - IMAGE - rewritten in VBA for use in Excel.                                                                                                           ' _
    LICENSE:  MIT                                                                                                                                                                           ' _
    VERSION:  1.0        22/03/2022                                                                                                                                                         ' _
                                                                                                                                                                                            ' _
    USAGE:    To be used on the worksheet:                                                                                                                                                  ' _
              =IMAGE("https://www.mysite.com/myimage.jpg")                 ' This will insert the image at the current cell                                                                 ' _
              =IMAGE("https://www.mysite.com/myimage.jpg", 2)              ' This will fill the current cell with the image                                                                 ' _
              =IMAGE("https://www.mysite.com/myimage.jpg", 4, 200, 100)    ' This will insert the image at the current cell, setting custom height/width properties                         ' _

    Option Explicit
    
    #If Win64 Then
        Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    #Else
        Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    #End If
    
    Private Const ERROR_SUCCESS As Long = 0
    Private Const BINDF_GETNEWESTVERSION As Long = &H10
    Private Const INTERNET_FLAG_RELOAD As Long = &H80000000
    
    Enum ImageModeEnum
        FitCell_MaintainAspect = 1              ' Resizes the image to fit inside the cell, maintaining aspect ratio.
        FitCell_IgnoreAspect = 2                ' Stretches or compresses the image to fit inside the cell, ignoring aspect ratio.
        OriginalSize = 3                        ' Leaves the image at original size, which may cause cropping.
        CustomSize = 4                          ' Allows the specification of a custom size.
    End Enum
    
    Function IMAGE(TargetURL As String, Optional ImageMode As ImageModeEnum = 3, Optional CustomHeight As Long = -1, Optional CustomWidth As Long = -1)
        
        Application.Volatile
        
        Dim TargetCell As Range, Img As Object
        Dim CallerCell As Variant, TargetFileName As String
        
        Const BASEPATH = "D:\TEMP\"
        
        If Len(Dir(BASEPATH, vbDirectory)) = 0 Then MkDir BASEPATH
            
        CallerCell = Application.Caller.Address
        
        If VarType(CallerCell) = vbString Then
            Set TargetCell = Range(CallerCell)
            
            TargetFileName = BASEPATH & GetFilenameFromURL(TargetURL)
            If Len(Dir(TargetFileName)) <> 0 Then Kill TargetFileName
            DownloadFile TargetURL, TargetFileName
            TargetCell.Parent.Shapes.AddPicture FileName:=TargetFileName, linktofile:=msoFalse, savewithdocument:=msoTrue, Top:=TargetCell.Top, left:=TargetCell.left, Width:=CustomWidth, Height:=CustomHeight
            Set Img = TargetCell.Parent.Shapes(TargetCell.Parent.Shapes.count)
            With Img
                .LockAspectRatio = IIf(ImageMode = 2, msoFalse, msoTrue)
                .Placement = xlMoveAndSize
                Select Case ImageMode
                    Case FitCell_MaintainAspect
                        If .Width > .Height Then
                            .Width = TargetCell.Width
                        Else
                            .Height = TargetCell.Height
                        End If
                    Case FitCell_IgnoreAspect
                        .Width = TargetCell.Width
                        .Height = TargetCell.Height
                    Case CustomSize
                        .ShapeRange.LockAspectRatio = msoFalse
                        If CustomHeight >= 0 And CustomWidth >= 0 Then
                            .Width = CustomWidth
                            .Height = CustomHeight
                        ElseIf CustomHeight >= 0 Then
                            .Height = CustomHeight
                        Else
                            .Width = CustomWidth
                        End If
                End Select
            End With
            
        End If
        
        IMAGE = ""
    End Function
    
    Private Function DownloadFile(ByVal SourceURL As String, ByVal LocalFile As String) As Boolean

        DownloadFile = URLDownloadToFile(0&, SourceURL, LocalFile, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS
    
    End Function
    
    Private Function GetFilenameFromURL(ByVal FilePath As String) As String
        
        If right$(FilePath, 1) <> "/" And Len(FilePath) > 0 Then
            If InStr(FilePath, "?") > 0 Then FilePath = Split(FilePath, "?")(0)
            GetFilenameFromURL = GetFilenameFromURL(left$(FilePath, Len(FilePath) - 1)) + right$(FilePath, 1)
        End If
    
    End Function
    


