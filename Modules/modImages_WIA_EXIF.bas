Attribute VB_Name = "modImages_WIA_EXIF"

                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||          WIA - EXIF PROPERTIES        ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
                                                                                                                                                                                            ' _
    AUTHOR:   Dan_W and Kallun Willock                                                                                                                                                      ' _
    PURPOSE:  Routines to read and write EXIF properties to JPG image files using                                                                                                           ' _
              the Windows Image Acquisition (WIA) COM Object.                                                                                                                               ' _
                                                                                                                                                                                            ' _
    LICENSE:  MIT                                                                                                                                                                       ' _
                                                                                                                                                                                            ' _
    VERSION:  1.0        25/03/2022                                                                                                                                                         ' _
                                                                                                                                                                                            ' _
    USAGE:                                                                                                                                                                                  ' _
              WriteEXIFData Filename, PropertyName, PropertyValue, (Opt) WriteOverOriginal = True, (Opt) CreateBackup                                                                       ' _
              - WriteEXIFData "C:\Temp\IMG01 - Copy.jpg", ImageTitle, "New Image Title"                                                                                                     ' _
              - WriteEXIFData "C:\Temp\IMG01 - Copy.jpg", ImageTitle, "Live Life On The Edge", True                                                                                         ' _
                                                                                                                                                                                            ' _
              - Comments = GetEXIFData("C:\Temp\IMG_20210508_170154.jpg", ImageComments)
    
    Option Explicit

    Public Enum PropertyNameEnum
        ImageDateTimeOriginal = 36867
        ImageTitle = 40091
        ImageComments = 40092
        ImageAuthor = 40093
        ImageKeywords = 40094
        ImageSubject = 40095
    End Enum
    
    Private Enum WIAImagePropertyType
        UndefinedImagePropertyType = 1000
        ByteImagePropertyType = 1001
        StringImagePropertyType = 1002
        UnsignedIntegerImagePropertyType = 1003
        LongImagePropertyType = 1004
        UnsignedLongImagePropertyType = 1005
        RationalImagePropertyType = 1006
        UnsignedRationalImagePropertyType = 1007
        VectorOfUndefinedImagePropertyType = 1100
        VectorOfBytesImagePropertyType = 1101
        VectorOfUnsignedIntegersImagePropertyType = 1102
        VectorOfLongsImagePropertyType = 1103
        VectorOfUnsignedLongsImagePropertyType = 1104
        VectorOfRationalsImagePropertyType = 1105
        VectorOfUnsignedRationalsImagePropertyType = 1106
    End Enum
    
    ' The sample photo is available at https://drive.google.com/file/d/1ZwU5L2HUi6pbehoYZ3-S2d5YIin8Lp8r/view?usp=sharing
    Const TargetFileName = "C:\PATHTOFILE\pexels-jill-evans-11567527.jpg"
    
    Sub Test_WriteProperties()
    
        Dim NewFileName As String
        NewFileName = WriteEXIFData(TargetFileName, ImageTitle, "White Concrete Building Under White Sky", True, True)
        WriteEXIFData NewFileName, ImageAuthor, "Jill Evans"
        WriteEXIFData NewFileName, ImageSubject, "Photo by Jill Evans from Pexels"
        WriteEXIFData NewFileName, ImageComments, "Source: https://www.pexels.com/photo/white-concrete-building-under-white-sky-11567527/"
        
        Debug.Print NewFileName
        
    End Sub
    
    Sub Test_ReadProperties()
        
        Dim Title As String
        Dim Subject As String
        Dim Comments As String
        Dim Author As String
        
        Title = GetEXIFData(TargetFileName, ImageTitle)
        Subject = GetEXIFData(TargetFileName, ImageSubject)
        Comments = GetEXIFData(TargetFileName, ImageComments)
        Author = GetEXIFData(TargetFileName, ImageAuthor)
        
        MsgBox "Title: " & Title & vbNewLine & _
               "Author: " & Author & vbNewLine & _
               "Subject: " & Subject & vbNewLine & _
               "Comments: " & Comments
    End Sub
    
    Public Function GetEXIFData(Filename As String, PropertyName As PropertyNameEnum) As String
    
        Dim Image               As Object
        Dim ImageProperty       As Object
        Dim Result              As String
        
        Set Image = CreateObject("WIA.ImageFile")
        Image.LoadFile Filename
    
        For Each ImageProperty In Image.Properties
            If ImageProperty.PropertyID = PropertyName Then
                If TypeName(ImageProperty.Value) = "String" Then
                    Result = ImageProperty.Value
                Else
                    Result = Replace(StrConv(ImageProperty.Value.BinaryData, vbUnicode), Chr(0), "")
                End If
                Exit For
            End If
        Next
        
        GetEXIFData = Result
    
    End Function
    
    Public Function WriteEXIFData(ByVal Filename As String, ByVal PropertyName As PropertyNameEnum, ByVal PropertyValue As Variant, Optional ByVal OverWriteOriginal As Boolean = True, Optional ByVal CreateBackup As Boolean)
    
        Dim Image               As Object
        Dim ImageProcess        As Object
        Dim ImageVector         As Object
        Dim NewFileName         As String
        
        If CreateBackup = True Then
            Dim BackUpFilename  As String
            BackUpFilename = Replace(Filename, ".jpg", "_BACKUP(" & Format(Now, "ddmmyyyy-hhnn") & ").jpg")
            FileCopy Filename, BackUpFilename
        End If
        
        Set Image = CreateObject("WIA.ImageFile")
        Set ImageProcess = CreateObject("WIA.ImageProcess")
        Set ImageVector = CreateObject("WIA.Vector")
        
        Image.LoadFile Filename
        
        ImageProcess.Filters.Add ImageProcess.FilterInfos("Exif").FilterID
        ImageProcess.Filters(1).Properties("ID") = PropertyName
        
        Select Case PropertyName
            
            Case PropertyNameEnum.ImageDateTimeOriginal
                Dim StringValue As String
                StringValue = Format(PropertyValue, "YYYY:MM:DD HH:MM:SS")
                ImageProcess.Filters(1).Properties("Type") = StringImagePropertyType
                ImageProcess.Filters(1).Properties("Value") = StringValue
            
            Case Else
                ImageProcess.Filters(1).Properties("Type") = VectorOfBytesImagePropertyType
                ImageVector.SetFromString PropertyValue
                ImageProcess.Filters(1).Properties("Value") = ImageVector
        
        End Select
        
        Set Image = ImageProcess.Apply(Image)
        
        If OverWriteOriginal = True Then
            NewFileName = Filename
            Kill Filename
        Else
            NewFileName = Replace(Filename, ".jpg", "_metadata.jpg")
            If Len(Dir(NewFileName)) > 0 Then Kill NewFileName
        End If
        
        Image.SaveFile NewFileName
    
        WriteEXIFData = NewFileName
    
    End Function
