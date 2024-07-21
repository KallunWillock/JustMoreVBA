Attribute VB_Name = "modImages_ExportPNG"
'@Lang VBA

                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||         PNG EXPORTER (v1.1)           ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                            ' _
    AUTHOR:   Kallun Willock                                                                                                                                                                ' _
                                                                                                                                                                                            ' _
    PURPOSE:  The PublishObjects.add method converts a designated worksheet/workbook into web-ready files                                                                                   ' _
              for publicatoin online. During this process, Excel will convert all images/shapes in the                                                                                      ' _
              designated area into 32bit high quality PNG files.                                                                                                                            ' _
                                                                                                                                                                                            ' _
    LICENSE:  MIT                                                                                                                                                                           ' _
                                                                                                                                                                                            ' _
    VERSION:  1.0        12/12/2019     Finalised v1                                                                                                                                        ' _
              1.1        14/09/2022     Rewrote library for Mac VBA compatability
                                                                                                                                                                                            ' _
    NOTES:    The main function, ExportPNG, accepts two optional arguments: PNG Filename/Path, and the target                                                                               ' _
              object for export. If no arguments are passed to the function, it reverts to defaults.                                                                                        ' _
                                                                                                                                                                                            ' _
              Error-handling is currently minimal to non-existant. Given that there is a lot clipboard/image/                                                                               ' _
              file IO activity, errors owing to issues with speed are likely to occur, so the error routines are                                                                            ' _
              primarily designed to force VBA to take a timeout for a few seconds, before attempting to proceeed.
                                                                                                                                                                                            ' _
    TODO:     [ ] Add user transparency colour settings                                                                                                                                     ' _
              [ ] Add webpublishing options/enumerations                                                                                                                                    ' _
              [ ] Add rooutine to get image dimensions
    
    Option Compare Text
    
    Sub TestExportPNG()
    
        Sheets("Demo Sheet").Activate
        ExportPNG "D:\REDDIT_TEST.PNG"
        Range("A1").Select
    
    End Sub


    Function ExportPNG(Optional ByVal PNGFilename As String, Optional ByVal Target As Variant) As Boolean
        
        Application.ScreenUpdating = False
        
        Dim ThePublisher        As PublishObject
        Dim TempWS              As Worksheet
        Dim HTMLOutput          As String
        Dim ExportPath          As String
        Dim Counter             As Long
        
        On Error GoTo ErrHandler
        
        ' If a PNG filename/path is not provided, the routine will default to this formulation.
        
        If Len(PNGFilename) = 0 Then PNGFilename = "D:\NewPNG_(" & Format(Now, "yymmddhhnnss") & ").png"
        
        ' If a target is not provided, the routine will use the currently available selection.
        
        If IsMissing(Target) Then Set Target = Application.Selection
        
        If TypeName(Target) = "Range" Then
            If Target.Cells.Count = 1 And IsEmpty(Target) Then
                Target.Parent.Shapes.SelectAll
                Set Target = Selection.Duplicate
                Target.Cut
            Else
                Target.CopyPicture xlScreen, XLPicture
            End If
        Else
            Target.Copy
        End If
        
        ' It appears that the generated PNG files, etc are saved in auto-generated subfolder, the name of which
        ' is derived from the basename of the HTML output file.
        
        HTMLOutput = Environ("TEMP") & "\HTMLOutput.htm"
        ExportPath = Environ("TEMP") & "\HTMLOutput_files\"
        
        ' Create a temporary worksheet to separate the item(s) for export and any
        ' other items on the worksheet. This TempWS will be deleted at the end of the routine.
        Set TempWS = Application.ActiveWorkbook.Sheets.Add
        TempWS.Paste
        
        ' Configure the PublishObject settings
        
        Set ThePublisher = Application.ActiveWorkbook.PublishObjects.Add(SourceType:=xlSourceSheet, _
            Filename:=HTMLOutput, sheet:=TempWS.Name, HtmlType:=xlHtmlStatic, Title:="PNG")
        
        ' VBA will throw errors if it tries to move too quickly through this process. Here, VBA will
        ' pause for approx 1.5 seconds.
        
        Pause 1.5
        
        ' Generate the necessary files for publication on the internet.
        
        ThePublisher.Publish True
        
        ' Get a list of the PNG files that have been generated and that located in the temporary export path
        
        FileList = GetFileList(ExportPath)
        
        ' Loop through each of the PNG files, assign an index number, and move to the destination path
        For Counter = LBound(FileList) To UBound(FileList)
            TmpPNGName = FileList(Counter)
            Name TmpPNGName As Replace(PNGFilename, ".png", "_" & Counter + 1 & ".png")
        Next
        
        ' Clean-up - the following code deletes the generate files, folders, and the temporary worksheet.
        
        On Error Resume Next
        Kill HTMLOutput                '    Environ("TEMP") & "\HTMLOutput.htm"
        Kill ExportPath & "*.*"        '    Environ("TEMP") & "\HTMLOutput_files\*.*"
        RmDir ExportPath               '    Environ("TEMP") & "\HTMLOutput_files\"
        On Error GoTo 0
        
        Application.DisplayAlerts = False
        TempWS.Delete
        Application.DisplayAlerts = True
        
        Set TempWS = Nothing
        Set Target = Nothing
        ExportPNG = CBool((UBound(FileList) >= 0))
        Application.ScreenUpdating = True
        Exit Function
    
ErrHandler:
        Pause 0.5
        Debug.Print "Error " & Err.Number & " - " & Err.Description
        Resume Next
    
    End Function

    Private Sub Pause(Optional ByVal Period As Single = 1)
        
        Dim TimeOut             As Single
        TimeOut = Timer + Period
        Do
            DoEvents
        Loop Until TimeOut < Timer
        
    End Sub

    Private Function GetFileList(ByVal BasePath As String) As Variant
        
        ' In addition to PNG image files of the target object(s), the PublishObjects.Add method
        ' generates a FileList.XML file containing details about each of the PNG files. This routine
        ' extracts those PNG files names and returns them as an array
        
        Dim FileListXML         As String
        Dim FileList            As Variant
        Dim XMLCode             As Variant
        Dim Counter             As Long
        
        FileListXML = BasePath & "FileList.xml"
        If Len(Dir(FileListXML)) = 0 Then Exit Function
        XMLCode = GetText(FileListXML, True)
        
        FileList = VBA.Filter(XMLCode, ".png", True, vbTextCompare)
        If UBound(FileList) >= -1 Then
            For Counter = LBound(FileList) To UBound(FileList)
                FileList(Counter) = BasePath & URLDecode(CStr(Split(Split(FileList(Counter), "HRef=" & Chr(34))(1), Chr(34))(0)))
            Next
            GetFileList = FileList
        End If
    
    End Function

    Private Function URLDecode(ByVal StringToDecode As String) As String
    
        ' The files generated through PublishObjects.Add method will have their filenames
        ' URL encoded automatically. This routine reverses that process. This is necessary, otherwise the
        ' generated files according to the XML file will not match the names of the files located on the
        ' filesystem.
    
        Dim TempResult          As String
        Dim CurrentCharacter    As Long
        
        CurrentCharacter = 1
        
        Do Until CurrentCharacter - 1 = Len(StringToDecode)
            Select Case Mid(StringToDecode, CurrentCharacter, 1)
                Case "+"
                    TempResult = TempResult & " "
                Case "%"
                    TempResult = TempResult & Chr(val("&h" & _
                    Mid(StringToDecode, CurrentCharacter + 1, 2)))
                    CurrentCharacter = CurrentCharacter + 2
                Case Else
                    TempResult = TempResult & Mid(StringToDecode, CurrentCharacter, 1)
            End Select
            CurrentCharacter = CurrentCharacter + 1
        Loop
        URLDecode = TempResult
    
    End Function
       
    Private Function GetText(ByVal TargetFilename As String, Optional ReturnArray As Boolean = False) As Variant
    
        Dim FileNumber          As Long
        Dim FileLength          As Long
        Dim FileContents        As String
        
        On Error GoTo ErrHandler
        
        FileNumber = FreeFile
        Open TargetFilename For Binary As #FileNumber
        FileLength = LOF(FileNumber)
        FileContents = Space$(FileLength)
        Get #FileNumber, , FileContents
        
        If ReturnArray Then
            GetText = Split(FileContents, vbCr)
        Else
            GetText = FileContents
        End If
    
ErrHandler:
        Close #FileNumber
    
    End Function
        


