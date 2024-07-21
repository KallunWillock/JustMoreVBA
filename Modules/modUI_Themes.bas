Attribute VB_Name = "modUI_Themes"
'@Lang VBA

                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||              UI  - THEME              ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
                                                                                                                                                                                            ' _
    AUTHOR:     Kallun Willock                                                                                                                                                              ' _
    PURPOSE:    Code to:    (1) engage the GetSysColor / SetSysColor APIs.                                                                                                                  ' _
                            (2) check the registry to ascertain which UI mode the user has set Office products: dark grey, black, light                                                     ' _
    LICENSE:    MIT
                                                                                                                                                                                            ' _
    VERSION:    1.4         12/04/2022          Added further details for SysColour API calls; added dark theme toggle routine                                                              ' _
                1.3         04/03/2022          Added code to change theme settings in the Registry for the VBIDE                                                                           ' _
                1.2         10/06/2021          Corrected these module details                                                                                                              ' _
                                                                                                                                                                                            ' _
    NOTES:      The SetSysColors function sends a WM_SYSCOLORCHANGE message to all windows to inform them                                                                                   ' _
                of the change in color. It also directs the system to repaint the affected portions of                                                                                      ' _
                all currently visible windows. This function affects only the current session. The new colors                                                                               ' _
                are not saved when the system terminates.

    Option Explicit
    
    #If Win64 Then
        Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
        Private Declare PtrSafe Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
    #Else
        Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
        Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
    #End If
                                                                                
    Enum UITHEME
        DARKGREY = 3
        BLACK = 4
        WHITE = 5
    End Enum
    
    Enum SystemColoursList
        COLOR_3DDKSHADOW = 21
        COLOR_3DFACE = 15                       ' COLOR_BTNFACE
        COLOR_3DHIGHLIGHT = 20                  ' COLOR_BTNHIGHLIGHT
        COLOR_3DHILIGHT = 20                    ' COLOR_BTNHIGHLIGHT
        COLOR_3DLIGHT = 22
        COLOR_3DSHADOW = 16                     ' COLOR_BTNSHADOW
        COLOR_ACTIVEBORDER = 10
        COLOR_ACTIVECAPTION = 2
        COLOR_APPWORKSPACE = 12
        COLOR_BACKGROUND = 1
        COLOR_BTNFACE = 15
        COLOR_BTNHIGHLIGHT = 20
        COLOR_BTNHILIGHT = 20                   ' COLOR_BTNHIGHLIGHT
        COLOR_BTNSHADOW = 16
        COLOR_BTNTEXT = 18
        COLOR_CAPTIONTEXT = 9
        COLOR_DESKTOP = 1                       ' COLOR_BACKGROUND
        COLOR_GRADIENTACTIVECAPTION = 27
        COLOR_GRADIENTINACTIVECAPTION = 28
        COLOR_GRAYTEXT = 17
        COLOR_HIGHLIGHT = 13
        COLOR_HIGHLIGHTTEXT = 14
        COLOR_HOTLIGHT = 26
        COLOR_INACTIVEBORDER = 11
        COLOR_INACTIVECAPTION = 3
        COLOR_INACTIVECAPTIONTEXT = 19
        COLOR_INFOBK = 24
        COLOR_INFOTEXT = 23
        COLOR_MENU = 4
        COLOR_MENUTEXT = 7
        COLOR_SCROLLBAR = 0
        COLOR_WINDOW = 5
        COLOR_WINDOWFRAME = 6
        COLOR_WINDOWTEXT = 8
    End Enum

    Private Const REG_BACKCOLOUR                As String = "HKEY_CURRENT_USER\Software\Microsoft\VBA\7.1\Common\CodeBackColors"
    Private Const REG_FORECOLOUR                As String = "HKEY_CURRENT_USER\Software\Microsoft\VBA\7.1\Common\CodeForeColors"
    Private Const BACKCOLOUR_DARK_THEME         As String = "4 0 4 7 6 4 4 4 11 4 0 0 0 0 0 0"
    Private Const FORECOLOUR_DARK_THEME         As String = "1 0 5 14 1 9 11 15 4 1 0 0 0 0 0 0"
    Private Const BACKCOLOUR_WHITE_THEME        As String = "0 0 0 7 6 0 0 0 0 0 0 0 0 0 0 0"
    Private Const FORECOLOUR_WHITE_THEME        As String = "0 0 5 0 1 10 14 0 0 0 0 0 0 0 0 0"
    
    Private Const REG_UITHEME                   As String = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\%VERSION%\Common\UI Theme"
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
    
    Function IsTheme(ByVal Theme As UITHEME) As Boolean
    
        IsTheme = ReadRegistry(REG_UITHEME) = Theme
    
    End Function

    ' Procedures:   ToggleDarkMode
    ' Purpose:      Available in Word only - toggles the application of the dark theme to the Active Document
            
    Sub ToggleDarkMode()

        Application.CommandBars.ExecuteMso "DarkModeOn"

    End Sub
    
    ' Procedures:   ToggleVBIDEThemeSettings
    ' Purpose:      Changes registry settings used to set the colour theme used in the VBIDE
    
    Sub ToggleVBIDEThemeSettings()
        
        If ReadRegistry(REG_BACKCOLOUR) = BACKCOLOUR_WHITE_THEME Then
            WriteRegistry REG_BACKCOLOUR, BACKCOLOUR_DARK_THEME, "REG_SZ"
            WriteRegistry REG_FORECOLOUR, FORECOLOUR_DARK_THEME, "REG_SZ"
        Else
            WriteRegistry REG_BACKCOLOUR, BACKCOLOUR_WHITE_THEME, "REG_SZ"
            WriteRegistry REG_FORECOLOUR, FORECOLOUR_WHITE_THEME, "REG_SZ"
        End If
    
    End Sub

    ' Procedure:    ReadRegistry / WriteRegistry
    ' Purpose:      Generic function to read/write registry settings - late-binding
    ' Returns:      Variant / NA
    
    Function ReadRegistry(ByVal RegPath As Variant) As Variant
        
        Dim WshShell    As Object
            
        If InStr(RegPath, "%VERSION%") Then
            RegPath = Replace(RegPath, "%VERSION%", Application.Version)
        End If
        Set WshShell = CreateObject("WScript.Shell")
        ReadRegistry = WshShell.RegRead(RegPath)
        
        Set WshShell = Nothing
    
    End Function
    
    Sub WriteRegistry(ByVal RegPath As Variant, ByVal RegValue As Variant, ByVal RegType As Variant)
        
        Dim WshShell    As Object
        
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.RegWrite RegPath, RegValue, RegType
        
        Set WshShell = Nothing
    
    End Sub
    
    ' Procedures:   GenerateSystemColourList
    ' Purpose:      Output the current VBALong values of the spectrum of system colours. Useful to get default values.
    
    Sub GenerateSystemColourList()

        Dim i           As Long
        For i = 0 To 20
            Debug.Print i, GetSysColor(i)
        Next

    End Sub
    
    Sub TestSetSysColours()
        Const CHANGE_INDEX = 1
        
        SetSysColors CHANGE_INDEX, COLOR_HIGHLIGHT, vbYellow
        SetSysColors CHANGE_INDEX, COLOR_HIGHLIGHTTEXT, vbBlack
        
    End Sub



