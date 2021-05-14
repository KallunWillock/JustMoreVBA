Attribute VB_Name = "modRegistry_UITheme"

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                                                               ' _
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                                                              ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                ' _
||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                               ' _
||||||||||||||||||||||||||              UI - THEMES              ||||||||||||||||||||||||||||||||||                                                                                ' _
||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                               ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                            ' _

'  AUTHOR:   Kallun Willock
'  PURPOSE:  
'  VERSION:  1.0 	20/05/2021

#If Win64 Then
    Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
#Else
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
#End If
                                                                               ' _
Private Enum UITHEME
    DARKGREY = 3
    BLACK = 4
    WHITE = 5
End Enum

...................................................................................................                                                                               ' _
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: 

Private Const REG_UITHEME As String = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\UI Theme"

' Procedures:   IsDarkTheme / IsWhiteTheme / IsTheme
' Purpose:      Checks the registry for th eOffice Application UI theme settings of the current user 
' Returns:      Boolean / Boolean / UITHEME enum    

Function IsDarkTheme() As Boolean
    IsDarkTheme = ReadRegistry(REG_UITHEME) = UITHEME.BLACK
End Function
Function IsWhiteTheme() As Boolean
    IsWhiteTheme = ReadRegistry(REG_UITHEME) = UITHEME.WHITE
End Function
Function IsTheme(Theme as UITHEME) as Boolean
    IsTheme = ReadRegistry(REG_UITHEME) = Theme
End Function

' Procedure:    ReadRegistry
' Purpose:      Generic function to check registry settings - late-binding
' Returns:      Long

Private Function ReadRegistry(Path As String) As Long
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    ReadRegistry = WshShell.RegRead(Path)
End Function

' Procedures:   PrintColourList
' Purpose:      Output the current VBALong values of the spectrum of system colours. Useful to get default values.

Sub PrintColourList()
    Dim i As Long
    For i = 0 To 20
        Debug.Print i, GetSysColor(i)
    Next
End Sub

Sub TestSetSysColours()
   Const CHANGE_INDEX = 1
   SetSysColors CHANGE_INDEX, COLOR_HIGHLIGHT, vbYellow
   SetSysColors CHANGE_INDEX, COLOR_HIGHLIGHTTEXT, vbBlack
End Sub