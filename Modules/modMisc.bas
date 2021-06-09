                                                                                                                                          ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||             MISCELLANEOUS             ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
                                                                                                                                          ' _
    AUTHOR:   Kallun Willock                                                                                                              ' _
    PURPOSE:  Collection of misc. procedures                                                                            ' _
                                                                                                                                          ' _
    VERSION:  1.0        25/05/2021                                                                                                       ' _
                                                                                                                                          ' _
    NOTES:    [•]                                  ' _
															
              -  [•]:	   [•]                                                      							   ' _
                                                                                                                                          ' _
    TODO:     [•]                                                                                       				  ' _
              																  ' _
    ...................................................................................................                                   ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Sub INCREMENT(ByRef ItemValue As Long, Optional Step As Long = 1)
    ItemValue = ItemValue + Step
End Sub
Sub DECREMENT(ByRef ItemValue As Long, Optional Step As Long = 1)
    ItemValue = ItemValue - Step
End Sub
Function ISITODD(Target) As Boolean
    ISITODD = WorksheetFunction.IsOdd(Target)
End Function
Function ISITEVEN(Target) As Boolean
    ISITEVEN = WorksheetFunction.IsEven(Target)
End Function

Function HOWMANY(Source As String, TargetText As String) As Long
    If InStr(Source, TargetText) = 0 Then HOWMANY = 0: Exit Function
    whenremoved = Len(Source) - Len(Replace(Source, TargetText, ""))
    HOWMANY = whenremoved / Len(TargetText)
End Function

'  Procedures:   Information re: files and folders

Function GETFILENAME(Filename As String) As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GETFILENAME = FSO.GETFILENAME(Filename)
    Set FSO = Nothing
End Function
Function GETEXTENSION(Filename As String) As String
    On Error Resume Next
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GETEXTENSION = FSO.getextensionname(Filename)
    Set FSO = Nothing
End Function
Function GETPATH(Filename As String) As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GETPATH = Replace(FSO.GetAbsolutePathName(Filename), FSO.GETFILENAME(Filename), "", , , vbDatabaseCompare)
    Set FSO = Nothing
End Function

Function DOWNLOAD(URL As String, FILENAME As String) As String
    WGH = WorksheetFunction.WebService(strurl)
    CreateObject("Scripting.FileSystemObject").CreateTextFile(FILENAME, True).WriteLine ("Code")
End Function