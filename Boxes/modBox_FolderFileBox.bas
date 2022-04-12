Attribute VB_Name = "modBox_FolderFileBox"

                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||             FOLDERFILEBOX             ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                            ' _
    AUTHOR:   Kallun Willock                                                                                                                                                                ' _
    PURPOSE:  The FolderFileBox creates a standard Open File/Save As File/Folder Browser dialog box, using                                                                                  ' _
              either API calls or the Shell Application COM object. The need for this came arises from the                                                                                  ' _
              the fact that not all MS Office products have access to these dialog boxes (for reasons I cannot begin to                                                                     ' _
              understand.
                                                                                                                                                                                            ' _
    LICENSE:  MIT                                                                                                                                                                           ' _
                                                                                                                                                                                            ' _
    VERSION:  1.0        12/04/2022
                                                                                                                                                                                            ' _
    NOTES:    Note that the filter currently doesn't work.                                                                                                                                  ' _
              With SelectFolder, if you designate a default path, you will not be able to navigate 'up' to the parent folder                                                                ' _
                                                                                                                                                                                            ' _
              To use GetOpenFilename:           Filename = GetFileName                                                                                                                      ' _
              To use GetSaveFilename:           Filename = GetFileName(False)                                                                                                               ' _
                                                                                                                                                                                            ' _
    TODO:     [ ] Fix the filter code                                                                                                                                                       ' _
              [ ] Add other system dialog boxes? API version of Select Folder                                                                                                               ' _


    Option Explicit
    
    #If VBA7 Then
        Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
        Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
        
        Private Type OPENFILENAME
            lStructSize As Long
            hwndOwner As LongPtr
            hInstance As LongPtr
            lpstrFilter As String
            lpstrCustomFilter As String
            nMaxCustFilter As Long
            nFilterIndex As Long
            lpstrFile As String
            nMaxFile As Long
            lpstrFileTitle As String
            nMaxFileTitle As Long
            lpstrInitialDir As String
            lpTitle As String
            flags As Long
            nFileOffset As Integer
            nFileExtension As Integer
            lpstrDefExt As String
            lCustData As LongPtr
            lpfnHook As LongPtr
            lpTemplateName As String
        End Type
      
    #Else
    
        Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
        Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
        
        Private Type OPENFILENAME
            lStructSize As Long
            hwndOwner As Long
            hInstance As Long
            lpstrFilter As String
            lpstrCustomFilter As String
            nMaxCustFilter As Long
            nFilterIndex As Long
            lpstrFile As String
            nMaxFile As Long
            lpstrFileTitle As String
            nMaxFileTitle As Long
            lpstrInitialDir As String
            lpTitle As String
            flags As Long
            nFileOffset As Integer
            nFileExtension As Integer
            lpstrDefExt As String
            lCustData As Long
            lpfnHook As Long
            lpTemplateName As String
        End Type
    #End If
    
     Function GetFileName(Optional OpenDLG As Boolean = True, Optional Title As String, Optional DefaultPath As String) As String
    
        Dim OpenFile    As OPENFILENAME
        Dim lReturn     As Long
      
        OpenFile.lpstrFilter = "All files (*.*) | *.*" & Chr$(0) & Chr$(0)
        OpenFile.nFilterIndex = 1
        OpenFile.hwndOwner = 0
        OpenFile.lpstrFile = String(257, 0)
        #If VBA7 Then
            OpenFile.nMaxFile = LenB(OpenFile.lpstrFile) - 1
            OpenFile.lStructSize = LenB(OpenFile)
        #Else
            OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
            OpenFile.lStructSize = Len(OpenFile)
        #End If
        OpenFile.lpstrFileTitle = OpenFile.lpstrFile
        OpenFile.nMaxFileTitle = OpenFile.nMaxFile
        OpenFile.lpstrInitialDir = DefaultPath & Chr$(0) & Chr$(0)
        OpenFile.lpTitle = Title
        OpenFile.flags = 0
        If OpenDLG Then
            lReturn = GetOpenFileName(OpenFile)
        Else
            lReturn = GetSaveFileName(OpenFile)
        End If
        
        If lReturn = 0 Then
            GetFileName = ""
        Else
            GetFileName = Trim(Left(OpenFile.lpstrFile, InStr(1, OpenFile.lpstrFile, vbNullChar) - 1))
        End If
      
    End Function
    
    Function SelectFolder(Optional DefaultPath As Variant) As String
        On Error Resume Next
        SelectFolder = CreateObject("Shell.Application").BrowseForFolder(0, "Select the folder you want to export to", 0, DefaultPath).self.path
    End Function
