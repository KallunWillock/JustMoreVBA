Attribute VB_Name = "modUI_Themes"
                                                                                                                                          ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||              UI - THEMES              ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
                                                                                                                                          ' _
    AUTHOR:   Kallun Willock                                                                                                              ' _
    PURPOSE:  API procedure to generate balloon tooltips / toast ui element                                                               ' _
                                                                                                                                          ' _
    VERSION:  1.1        09/06/2021                                                                                                       ' _
                                                                                                                                          ' _
    NOTES:    Using the Toast function is relatively straight-forward. Through trial-and-error,                                           ' _
              I have worked out that:                                                                                                     ' _
                                                                                                                                          ' _
              -  TITLE:   The TITLE argument will accept a maximum of 63 characters.                                                         ' _
              -  CONTENT: The CONTENT argument will accept a maximum of 154 characters.                                                      ' _
              -  TYPE:    The first four types are accompanied with a system beep, whereas the latter four are silent.                       ' _
                                                                                                                                          ' _
    TODO:     Icons - should be able to add icons                                                                                         ' _
              Timeout - NB: uTimeout - utility?                                                                                           ' _

    Option Explicit
    
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
    
    Private Const REG_UITHEME As String = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\UI Theme"
                                                                                                                                            ' _
        ...................................................................................................                                 ' _
        :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    
    
    ' Procedures:   IsDarkTheme / IsWhiteTheme / IsTheme
    ' Purpose:      Checks the registry for th eOffice Application UI theme settings of the current user
    ' Returns:      Boolean / Boolean / UITHEME enum
    
    Function IsDarkTheme() As Boolean
        IsDarkTheme = ReadRegistry(REG_UITHEME) = UITHEME.BLACK
    End Function
    
    Function IsWhiteTheme() As Boolean
        IsWhiteTheme = ReadRegistry(REG_UITHEME) = UITHEME.WHITE
    End Function
    
    Function IsTheme(Theme As UITHEME) As Boolean
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
