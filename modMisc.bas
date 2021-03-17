Attribute VB_Name = "modRegistry_UITheme"
Const REG_UITHEME As String = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\UI Theme"

Enum UITHEME
    DARKGREY = 3
    BLACK = 4
    WHITE = 5
End Enum
    
Function IsDarkTheme() As Boolean
    IsDarkTheme = ReadRegistry(REG_UITHEME) = UITHEME.BLACK
End Function
Function IsWhiteTheme() As Boolean
    IsWhiteTheme = ReadRegistry(REG_UITHEME) = UITHEME.WHITE
End Function
Private Function ReadRegistry(Path As String) As Long
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    ReadRegistry = WshShell.RegRead(Path)
End Function

