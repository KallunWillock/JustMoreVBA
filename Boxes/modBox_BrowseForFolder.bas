Attribute VB_Name = "modBox_BrowseForFolder"

                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||         BROWSEFORFOLDER (v1.0)        ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                            
Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As LongPtr, ByVal nFolder As Long, pidl As ITEMIDLIST) As LongPtr
  Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As LongPtr, ByVal pszPath As String) As Boolean
  Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" (lpBrowseInfo As BrowseInfo) As LongPtr
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal Msg As Long, wParam As Any, lParam As Any) As LongPtr
#Else
  Private Enum LongPtr
  [_]
  End Enum
  Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As LongPtr, ByVal nFolder As Long, pidl As ITEMIDLIST) As LongPtr
  Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As LongPtr, ByVal pszPath As String) As Boolean
  Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpBrowseInfo As BrowseInfo) As LongPtr
  Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal Msg As Long, wParam As Any, lParam As Any) As LongPtr
#End If

Private Type BrowseInfo
  Owner As LongPtr
  RootIdl As LongPtr
  DisplayName As String
  Title As String
  Flags As Long
  CallbackAddress As LongPtr
  CallbackParam As LongPtr
  Image As Long
End Type

Private Type SHITEMID
  cb As Long
  abID As Byte
End Type

Private Type ITEMIDLIST
  mkid As SHITEMID
End Type

Private Const MAX_PATH = 260

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILEDA As Long = 3
Private Const BFFM_VALIDATEFAILED = BFFM_VALIDATEFAILEDA
Private Const BFFM_VALIDATEFAILEDW As Long = 4
Private Const BFFM_SETSTATUSTEXTA As Long = WM_USER + 100
Private Const BFFM_ENABLEOK As Long = WM_USER + 101
Private Const BFFM_SETSELECTIONA = WM_USER + 102
Private Const BFFM_SETSELECTION = BFFM_SETSELECTIONA
Private Const BFFM_SETSELECTIONW As Long = WM_USER + 103
Private Const BFFM_SETSTATUSTEXTW As Long = WM_USER + 104
Private Const BFFM_SETOKTEXT As Long = WM_USER + 105
Private Const BFFM_SETEXPANDED As Long = WM_USER + 106

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

Private mStartFolder As String

Public Function BrowseForFolder(Optional ByVal Flags As BIF_OPTIONS, Optional ByVal Title As String, Optional ByVal StartFolder As String = "C:\") As String
    
  Dim BI          As BrowseInfo
  Dim FolderPath  As String
  Dim Result      As Long
  Dim IDL         As ITEMIDLIST
  Dim Handle      As LongPtr
  
  FolderPath = Space(MAX_PATH)
  With BI
    .Owner = 0
    .RootIdl = 0
    .DisplayName = Space(MAX_PATH)
    .Title = Title
    .Flags = Flags
  End With
  
  If StartFolder <> "" Then
    mStartFolder = StartFolder & vbNullChar
    BI.CallbackAddress = GetAddressofFunction(AddressOf BrowseCallbackProc)
  End If
   
  Handle = SHBrowseForFolder(BI)
  If (Handle <> 0) Then
    FolderPath = Space(MAX_PATH)
    If (CBool(SHGetPathFromIDList(Handle, FolderPath))) Then
      BrowseForFolder = TrimAtNull(FolderPath)
    Else
      BrowseForFolder = TrimAtNull(FolderPath = BI.Title)
    End If
  End If
  Call GlobalFree(Handle)
     
End Function

Private Function TrimAtNull(ByVal SourceString As String) As String
  If SourceString = vbNullString Then Exit Function
  TrimAtNull = Split(SourceString, vbNullChar)(0)
End Function

Private Function BrowseCallbackProc(ByVal hWnd As LongPtr, ByVal Msg As LongPtr, ByVal Pointer As LongPtr, ByVal Data As String) As LongPtr
  On Error Resume Next

  Dim Result          As Long
  Dim Buffer          As String
  
  Select Case Msg
    Case BFFM_VALIDATEFAILED
      '
    Case BFFM_INITIALIZED
      Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal mStartFolder)
    Case BFFM_SELCHANGED
      Buffer = Space(MAX_PATH)
      Result = SHGetPathFromIDList(Pointer, Buffer)
      If Result = 1 Then
        Call SendMessage(hWnd, BFFM_SETSTATUSTEXTA, 0, Buffer)
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

