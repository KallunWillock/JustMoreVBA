Attribute VB_Name = "modBox_BrowseForFolder"
  
                                                                                                                                                                                              ' _
      |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
      ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
      ||||||||||||||||||||||||||         BROWSEFORFOLDER (v1.1)        ||||||||||||||||||||||||||||||||||                                                                                     ' _
      ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
      |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                              ' _
      AUTHOR:   Kallun Willock                                                                                                                                                                ' _
      URL:      https://github.com/KallunWillock/JustMoreVBA/blob/main/Boxes/modBox_BrowseForFolder.bas                                                                                       ' _
      PURPOSE:  Calls Win32 APIs to create the Browse For Folder dialog box.                                                                                                                  ' _
      LICENSE:  MIT                                                                                                                                                                           ' _
      VERSION:  1.1        13/10/2024         Updated code to be Unicode enabled                                                                                                              ' _
                1.0        22/09/2024         Uploaded module demonstrating how to call the BrowseForFolder Dialog Box in 64bit Office.                                                       ' _
                                                                                                                                                                                              ' _
      NOTES:    The BrowseForFolder Win32 API requires a callback, otherwise it will crash the housing application.                                                                           ' _
                As such, the routines should be stored in a standard module.                                                                                                                  ' _
                Also, I would advise reading the Remarks section of the MSDN entry for SHBowserForFolderW:
                                                                                                                                                                                              ' _
                https://learn.microsoft.com/en-us/windows/win32/api/shlobj_core/nf-shlobj_core-shbrowseforfolderw
  
  Option Explicit
  
  #If VBA7 Then
    Private Declare PtrSafe Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As LongPtr, ByVal nFolder As Long, pidl As ITEMIDLIST) As LongPtr
    Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As LongPtr, ByVal pszPath As String) As Boolean
    Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" (lpBrowseInfo As BrowseInfo) As LongPtr
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal pidl As LongPtr, ByVal pszPath As String) As Long
    Private Declare PtrSafe Function SHGetPathFromIDListW Lib "shell32" (ByVal pidList As LongPtr, ByVal lpBuffer As LongPtr) As Long
    Private Declare PtrSafe Function SHBrowseForFolderA Lib "shell32.dll" (lpBrowseInfo As BrowseInfo) As LongPtr
    Private Declare PtrSafe Function SHBrowseForFolderW Lib "shell32" (lpbi As BrowseInfoW) As LongPtr
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As Any) As Long
    Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As Any) As Long
  #Else
    Private Enum LongPtr
    [_]
    End Enum
    Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As LongPtr, ByVal nFolder As Long, pidl As ITEMIDLIST) As LongPtr
    Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As LongPtr, ByVal pszPath As String) As Boolean
    Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpBrowseInfo As BrowseInfo) As LongPtr
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal pidl As LongPtr, ByVal pszPath As String) As Long
    Private Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidList As LongPtr, ByVal lpBuffer As LongPtr) As Long
    Private Declare Function SHBrowseForFolderA Lib "shell32.dll" (lpBrowseInfo As BrowseInfo) As LongPtr
    Private Declare Function SHBrowseForFolderW Lib "shell32" (lpbi As BrowseInfow) As LongPtr
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As Any) As Long
    Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As Any) As Long
  #End If
  
  Private Type BrowseInfo
    Owner           As LongPtr
    RootIdl         As LongPtr
    DisplayName     As String
    Title           As String
    Flags           As Long
    CallbackAddress As LongPtr
    CallbackParam   As LongPtr
    Image           As Long
  End Type
  
  Private Type BrowseInfoW
    Owner           As LongPtr
    RootIdl         As LongPtr
    DisplayName     As LongPtr
    Title           As LongPtr
    Flags           As Long
    CallbackAddress As LongPtr
    CallbackParam   As LongPtr
    Image           As Long
  End Type
  
  Private Type SHITEMID
    cb              As Long
    abID            As Byte
  End Type
  
  Private Type ITEMIDLIST
    mkid            As SHITEMID
  End Type
  
  'Private Const MAX_PATH = 260
  Private Const MAX_PATH_UNICODE As Long = 519 ' = 260 * 2 - 1
  Private Const MAX_PATH  As Long = MAX_PATH_UNICODE
  
  Private Const WM_USER = &H400
  Private Const BFFM_INITIALIZED As Long = &H1
  Private Const BFFM_SELCHANGED As Long = &H2
  Private Const BFFM_VALIDATEFAILEDA As Long = &H3
  Private Const BFFM_VALIDATEFAILED = BFFM_VALIDATEFAILEDA
  Private Const BFFM_VALIDATEFAILEDW As Long = &H4
  Private Const BFFM_IUNKNOWN = &H5
  Private Const BFFM_SETSTATUSTEXTA As Long = WM_USER + &H64      '+ 100
  Private Const BFFM_ENABLEOK As Long = WM_USER + &H65            '+ 101
  Private Const BFFM_SETSELECTIONA = WM_USER + &H66               '+ 102
  Private Const BFFM_SETSELECTION = BFFM_SETSELECTIONA
  Private Const BFFM_SETSELECTIONW As Long = WM_USER + &H67       '+ 103
  Private Const BFFM_SETSTATUSTEXTW As Long = WM_USER + &H68      '+ 104
  Private Const BFFM_SETOKTEXT As Long = WM_USER + &H69           '+ 105      Unicode only
  Private Const BFFM_SETEXPANDED As Long = WM_USER + &H6A         '+ 106      Unicode only
  
  ' From MSDN @ https://learn.microsoft.com/en-us/windows/win32/api/shlobj_core/nf-shlobj_core-shbrowseforfolderw
  
  ' There are two styles of dialog box available. The older style is displayed by default and is not resizable.
  ' The newer style provides a number of additional features, including drag-and-drop capability within the dialog box,
  ' reordering, deletion, shortcut menus, the ability to create new folders, and other shortcut menu commands.
  ' Initially, it is larger than the older dialog box, but the user can resize it. To specify a dialog box using the
  ' newer style, set the BIF_USENEWUI flag in the ulFlags member of the BROWSEINFO structure.
  
  Public Enum BIF_OPTIONS
    BIF_RETURNONLYFSDIRS = &H1&           ' Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
    BIF_DONTGOBELOWDOMAIN = &H2&          ' Do not include network folders below the domain level in the dialog box's tree view control.
    BIF_STATUSTEXT = &H4&                 ' include status area for callback
    BIF_RETURNFSANCESTORS = &H8&
    BIF_EDITBOX = &H10&                   ' Include an edit control in the browse dialog box that allows the user to type the name of an item.
    BIF_VALIDATE = &H20&
    BIF_NEWDIALOGSTYLE = &H40&
    BIF_USENEWUI = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
    BIF_BROWSEINCLUDEURLS = &H80&
    BIF_UAHINT = &H100&
    BIF_NONEWFOLDERBUTTON = &H200&
    BIF_NOTRANSLATETARGETS = &H400&
    BIF_BROWSEFORCOMPUTER = &H1000&
    BIF_BROWSEFORPRINTER = &H2000&
    BIF_BROWSEINCLUDEFILES = &H4000&
    BIF_SHAREABLE = &H8000&
    BIF_BROWSEFILEJUNCTIONS = &H100000
  End Enum
  
  Public Function BrowseForFolder(Optional ByVal Flags As BIF_OPTIONS, Optional ByVal Title As String, Optional ByVal StartFolder As String = "C:\") As String
      
    Dim BI          As BrowseInfoW
    Dim FolderPath  As String
    Dim Result      As Long
    Dim IDL         As ITEMIDLIST
    Dim Handle      As LongPtr
    
    FolderPath = Space(MAX_PATH)
    With BI
      .Owner = 0^
      .RootIdl = 0^
      .DisplayName = StrPtr(Space(MAX_PATH))
      .Title = StrPtr(Title & vbNullChar)
      .Flags = Flags
      If StartFolder <> "" Then
        .CallbackParam = StrPtr(StartFolder & vbNullChar)
        .CallbackAddress = GetAddressofFunction(AddressOf BrowseCallbackProc)
      End If
    End With
    
    Handle = SHBrowseForFolderW(BI)
    If (Handle <> 0) Then
      FolderPath = Space(MAX_PATH)
      If (CBool(SHGetPathFromIDListW(Handle, StrPtr(FolderPath)))) Then
        BrowseForFolder = TrimAtNull(FolderPath)
      Else
        BrowseForFolder = TrimAtNull(FolderPath = BI.Title)
      End If
    End If
    'Call GlobalFree(Handle)
    CoTaskMemFree Handle
       
  End Function
  
  Private Function TrimAtNull(ByVal SourceString As Variant) As String
    If SourceString = vbNullString Then Exit Function
    TrimAtNull = Split(SourceString, vbNullChar)(0)
  End Function
  
  Private Function BrowseCallbackProc(ByVal hwnd As LongPtr, ByVal Msg As LongPtr, ByVal Pointer As LongPtr, ByVal Data As LongPtr) As LongPtr
    On Error Resume Next
  
    Dim Result      As Long
    Dim Buffer      As String
    
    Select Case Msg
      Case BFFM_INITIALIZED
        Call SendMessageW(hwnd, BFFM_SETSELECTION, 1&, Data)
      Case BFFM_SELCHANGED
        Buffer = Space(MAX_PATH)
        Result = SHGetPathFromIDListW(Pointer, StrPtr(Buffer))
        If Result = 1 Then
          Call SendMessageW(hwnd, BFFM_SETSTATUSTEXTA, 0, StrPtr(Buffer))
        End If
    End Select
    BrowseCallbackProc = 0
  End Function
    
  Private Function GetAddressofFunction(PtrAddress As LongPtr) As LongPtr
    GetAddressofFunction = PtrAddress
  End Function
  
  Public Function BrowseFolders(Optional StartFolder As String = "C:\") As String
    BrowseFolders = BrowseForFolder(BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE, "Browse for Folder", StartFolder)
  End Function
  
  ' **********************************************************
  
  Sub TestBrowseForFolder()
    Dim Result          As String
    Dim StartFolder     As String
    
    ' The BIF_BROWSEINCLUDEFILES Flag extends the functionality of the Browse For Folder dialog box to allow
    ' the user to select a file.
    
    Result = BrowseForFolder(BIF_USENEWUI Or BIF_BROWSEINCLUDEFILES, "Select a file")
    Debug.Print Result
    
    StartFolder = Environ("USERPROFILE")
    Result = BrowseForFolder(BIF_RETURNONLYFSDIRS Or BIF_USENEWUI, "Browse for Folder", StartFolder)
    Debug.Print Result
  End Sub
