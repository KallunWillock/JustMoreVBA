
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
    LICENSE:    MIT                                                                                                                                                                                         ' _
    VERSION:    1.3         04/03/2022          Added code to change theme settings in the Registry for the VBIDE                                                                           ' _
                1.2         10/06/2021          Corrected these module details                                                                                                              ' _
                                                                                                                                                                                            ' _

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

    Sub ToggleThemeSettings()
        
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

