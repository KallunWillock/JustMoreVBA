
                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||             ICONMSGBOX (v1.3)         ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                            ' _
    AUTHOR:   Kallun Willock                                                                                                                                                                ' _
    PURPOSE:  The IconMsgBox is a Unicode-enabled MessageBox that can display a custom icon (ICO file).                                                                                     ' _
              It also features a timeout feature resulting in it closing down after a designated period of time.                                                                            ' _
    LICENSE:  MIT                                                                                                                                                                           ' _
    VERSION:  1.3        07/04/2022         Corrections; Fixed sound notifications; implemented right-justified text;                                                                       ' _
                                            Improved existing demo and added two more samples to demonstrate                                                                                ' _
                                            functionality: including Japanese and simplified Chinese script, custom icon file, etc.                                                         ' _
              1.2        22/03/2022         Improved timeout functiontionality. Fixed 32-bit compatibility.                                                                                 ' _
              1.1        16/03/2022         Added timeout functiontionality. Improved comments and corrected errors.                                                                        ' _
              1.0        18/02/2022         Version 1 uploaded to Github. Compatible with 32-bit and 64-bit Office
                                                                                                                                                                                            ' _
    NOTES:    A return value of 32000 indicates that the user did not press a button.                                                                                                       ' _
              The timeout period is measured in milliseconds, but where a whole number under 60 has been passed                                                                             ' _
              to IconMsgBox for the timeout parameter, that will be interpreted as seconds.
                                                                                                                                                                                            ' _
    TODO:     [X] Add access to system DLL icons                                                                                                                                            ' _
              [X] Allow use of custom ICO files                                                                                                                                             ' _
              [X] Unicode compatibility                                                                                                                                                     ' _
              [X] Add timeout feature                                                                                                                                                       ' _
              [X] Sound notification                                                                                                                                                        ' _
              [X] Left / Right Justification of content                                                                                                                                     ' _
              [X] Add unicode conversion compatibility                                                                                                                                      ' _
              [ ] RTL Support                                                                                                                                                               ' _
              [ ] Custom button labels?

    Option Explicit

    Public Enum ImageDLL
        icn_shell32                 '        C:\Windows\System32\shell32.dll                - 329   icons
        icn_imageres                '        C:\Windows\System32\imageres.dll               - 334   icons
        icn_pifmgr                  '        C:\Windows\System32\pifmgr.dll                 - 38    icons
        icn_accessibilitycpl        '        C:\Windows\System32\accessibilitycpl.dll       - 24    icons
        icn_ddores                  '        C:\Windows\System32\ddores.dll                 - 149   icons
        icn_moricons                '        C:\Windows\System32\moricons.dll               - 113   icons
        icn_explorer                '        C:\Windows\explorer.exe                        - 28    icons
        icn_mmcndmgr                '        C:\Windows\System32\mmcndmgr.dll               - 129   icons
        icn_mmres                   '        C:\Windows\System32\mmres.dll                  - 18    icons
        icn_netcenter               '        C:\Windows\System32\netcenter.dll              - 14    icons
        icn_netshell                '        C:\Windows\System32\netshell.dll               - 165   icons
        icn_networkexplorer         '        C:\Windows\System32\networkexplorer.dll        - 20    icons
        icn_pnidui                  '        C:\Windows\System32\pnidui.dll                 - 43    icons
        icn_sensorscpl              '        C:\Windows\System32\sensorscpl.dll             - 22    icons
        icn_mshtml                  '        C:\Windows\System32\mshtml.dll                 - 27    icons
        icn_diagcpl                 '        C:\Windows\System32\diagcpl.dll                - 9     icons
    End Enum
    
    #If VBA7 And Win64 Then

        Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
        Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
        Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInstance As LongPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As LongPtr
        Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
        Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
        Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
        Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
        Private Declare PtrSafe Function MessageBoxTimeoutW Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long, ByVal wLanguageID As Long, ByVal lngMilliseconds As Long) As Long
        Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
        Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
        Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long

        Private pHook                               As LongPtr
        Private hIcon                               As LongPtr
        Private hIconWnd                            As LongPtr
    #Else

        Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
        Private Declare Function ExtractIcon Lib "SHELL32.DLL" Alias "ExtractIconA" (ByVal hInst As LongPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As LongPtr
        Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
        Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
        Private Declare Function GetCurrentThreadId Lib "KERNEL32" () As Long
        Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
        Private Declare Function MessageBoxTimeoutW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long, ByVal wLanguageID As Long, ByVal lngMilliseconds As Long) As Long
        Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
        Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadID As Long) As Long
        Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

        Private pHook                               As Long
        Private hIcon                               As Long
        Private hIconWnd                            As Long

    #End If
    
    Private Const WH_CBT                            As Long = &H5
    Private Const HCBT_ACTIVATE                     As Long = &H5
    Private Const STM_SETICON                       As Long = &H170
    Private Const MSGBOX_CLASSNAME                  As String = "#32770"
    Private Const MB_SETFOREGROUND                  As Long = &H10000
    Private Const MB_TOPMOST                        As Long = &H40000
    Private Const MB_RIGHT                          As Long = &H80000
    Private Const MB_RTLREADING                     As Long = &H100000
    Private Const ICNMB_ADDBEEP                     As Long = &H10
    Private Const ICNMB_ADDICON                     As Long = &H20
    
    Sub IconMsgBox_Demo1()
    ' IconMsgBox with a 15 second timeout - Unicode compatible (displays Hello World in Japanese script - katakana) - uses icon#77 of the ImageRes.dll
        
        Dim Content                                 As String
        Dim Title                                   As String
        Dim IconFilePath                            As String
        Dim HelloWorld_JP                           As Variant
        Dim TimeOutPeriod                           As Long
        
        HelloWorld_JP = Array(12495, 12525, 12540, 12539, 12527, 12540, 12523, 12489)
        TimeOutPeriod = 15000
        
        Title = "Title - IconMsgBox_Demo1 - " & GetUnicodeMessage(HelloWorld_JP)
        Content = "IconMsgBox (v1.3) allows for:" & vbNewLine & vbNewLine & "1. Custom icons" & vbNewLine & "2. Unicode text"
        Content = Content & vbNewLine & "3. Timeout feature" & vbNewLine & "4. Sound notification" & vbNewLine & "5. Right-justified text support"
        Content = Content & vbNewLine & vbNewLine & "This routine uses Icon#77 of the ImageRes.dll library, displays Japanese text in the title, and with timeout set for " & TimeOutPeriod / 1000 & " seconds."
        Content = Content & vbNewLine & vbNewLine & "Note that there is no sound notification and the text is left-justified."
        
        Debug.Print IconMsgBox(Content, vbYesNo, Title, IconFilePath, icn_imageres, 77, TimeOutPeriod)
        
    End Sub
    
    Sub IconMsgBox_Demo2()
    ' IconMsgBox with no timeout - Unicode compatible (displays Hello World in simplified Chinese script) - uses icon#12 of the ddores.dll - right-justified content - beep notification
        
        Dim Content                                 As String
        Dim Title                                   As String
        Dim BellIcon                                As String
        Dim HelloWorld_ZHCN                         As Variant
        
        HelloWorld_ZHCN = Array(20320, 22909, -244, 19990, 30028)
        BellIcon = GetUnicodeMessage(128276)
        
        Title = "Title - IconMsgBox_Demo2 - " & GetUnicodeMessage(HelloWorld_ZHCN) & "  " & BellIcon
        Content = "IconMsgBox (v1.3) allows for:" & vbNewLine & vbNewLine & "1. Custom icons" & vbNewLine & "2. Unicode text"
        Content = Content & vbNewLine & "3. Timeout feature" & vbNewLine & "4. Sound notification" & vbNewLine & "5. Right-justified text support"
        Content = Content & vbNewLine & vbNewLine & "This routine uses Icon#12 of the ddores.dll library, displays simplified Chinese script in the title, and with no timeout set."
        Content = Content & vbNewLine & vbNewLine & "Note that it includes a sound notification " & BellIcon & " and the text is right-justified."
        
        Debug.Print IconMsgBox(Content, vbYesNo, Title, , icn_ddores, 12, , , True, True)

    End Sub

    Sub IconMsgBox_Demo3()
    ' IconMsgBox with 5 second timeout - Unicode compatible (displays unicode tick marks) - uses custom icon file, github.ico - left-justified content - no beep
        
        Dim Content                                 As String
        Dim Title                                   As String
        Dim IconFilePath                            As String
        Dim TimeOutPeriod                           As Long
        Dim TickMark                                As String
        
        TimeOutPeriod = 5
        
        TickMark = ChrW(10004)
        
        Title = "Title - IconMsgBox_Demo3"
        Content = "IconMsgBox (v1.3) allows for:" & vbNewLine & vbNewLine & TickMark & " Custom icons" & vbNewLine & TickMark & " Unicode text"
        Content = Content & vbNewLine & TickMark & " Timeout feature" & vbNewLine & TickMark & " Sound notification" & vbNewLine & TickMark & " Right-justified text support"
        Content = Content & vbNewLine & vbNewLine & "This routine uses a custom icon file (github.ico), displays unicode characters (tick mark), and has a timeout set for " & TimeOutPeriod & " seconds."
        Content = Content & vbNewLine & vbNewLine & "Note that there is no sound notification and the text is left-justified."
        IconFilePath = ThisWorkbook.Path & "\github.ico"
        
        Debug.Print IconMsgBox(Content, vbYesNo, Title, IconFilePath, , , TimeOutPeriod)

    End Sub
    
    Sub IconMsgBox_Demo4()
    ' IconMsgBox with no timeout - Unicode compatible (displays unicode tick marks) - custom width - atttempts to use non-existant custom icon file, nogithub.ico - left-justified content - no beep
    
        Dim Content                                 As String
        Dim Title                                   As String * 100
        Dim IconFilePath                            As String
        Dim TickMark                                As String
        Dim BellIcon                                As String
        
        TickMark = ChrW(10003)
        BellIcon = GetUnicodeMessage(128276)
        
        Title = "Title - IconMsgBox_Demo4 " & BellIcon
        Content = "IconMsgBox (v1.3) allows for:" & vbNewLine & vbNewLine & TickMark & " Custom icons" & vbNewLine & TickMark & " Unicode text"
        Content = Content & vbNewLine & TickMark & " Timeout feature" & vbNewLine & TickMark & " Sound notification" & vbNewLine & TickMark & " Right-justified text support"
        Content = Content & vbNewLine & vbNewLine & "This routine demonstrates that " & vbNewLine & "there is some (limited) flexibility " & vbNewLine & "in setting the width of the IconMsgBox."
        Content = Content & vbNewLine & vbNewLine & "It also demonstrates " & vbNewLine & "the default icon if no valid icon is " & vbNewLine & "found at the designated filepath."
        Content = Content & vbNewLine & vbNewLine & "Note that it includes a sound" & vbNewLine & "notification " & BellIcon & " and the text is left-justified."
        
        IconFilePath = "C:\PATHTOFILE\nogithub.ico"
        
        Debug.Print IconMsgBox(Content, vbOKOnly, Title, IconFilePath, , , , , True)

    End Sub

    Private Function GetUnicodeMessage(ByVal UnicodeCharacters As Variant) As String
    
        Dim Counter                                 As Long
        Dim TempMessage                             As String
        
        If IsArray(UnicodeCharacters) = False Then UnicodeCharacters = Array(UnicodeCharacters)
        
        For Counter = LBound(UnicodeCharacters) To UBound(UnicodeCharacters)
            TempMessage = TempMessage & UnicodeConverter(UnicodeCharacters(Counter))
        Next
        
        GetUnicodeMessage = TempMessage
        
    End Function
    
    Private Function UnicodeConverter(ByVal Code As Variant) As String
        
        If IsNumeric(Code) = False Then
            If Left(Code, 2) = "U+" Then
                Code = CLng(Replace(Code, "U+", "&H"))
            ElseIf Left(Code, 2) = "0x" Then
                Code = CLng(Replace(Code, "0x", "&H"))
            Else
                Code = CLng("&H" & Code)
            End If
        End If
        
        ' Conversion algorithm below partially based on code by GSerg at
        ' https://stackoverflow.com/questions/57158679/alternative-of-chrw-function
        ' Revised to allow for negative values (see Demo2, comma) | Sourced: 07/04/2022
        If (Code >= &H8000 And Code <= &HD7FF&) Or (Code >= &HE000& And Code <= &HFFFF&) Then
            UnicodeConverter = ChrW(Code)
        Else
            Code = Code - &H10000
            UnicodeConverter = ChrW(&HD800 Or (Code \ &H400&)) & ChrW(&HDC00 Or (Code And &H3FF&))
        End If
        
    End Function

    Public Function IconMsgBox(ByVal Content As String, _
                      Optional ByVal Style As VbMsgBoxStyle = vbOKOnly, _
                      Optional ByVal Title As String = "", _
                      Optional ByVal IconFilePath As String, _
                      Optional ByVal IconLibrary As ImageDLL, _
                      Optional ByVal IconNumber As Long = 0, _
                      Optional ByVal TimeOut As Long = -1, _
                      Optional ByVal RightToLeft As Boolean = False, _
                      Optional ByVal BeepNotification As Boolean = False, _
                      Optional ByVal RightJustified As Boolean = False) As VbMsgBoxResult

        Dim IconPath                                As String
        Dim TargetThreadID                          As Long
        
        If Len(Dir(IconFilePath)) = 0 Then IconFilePath = vbNullString
        
        If IconFilePath = vbNullString Then
            Dim ImageLibraryPaths                   As Variant
            ImageLibraryPaths = Array("System32\shell32.dll", "system32\imageres.dll", "system32\pifmgr.dll", "system32\accessibilitycpl.dll", "system32\ddores.dll", "system32\moricons.dll", _
                                      "explorer.exe", "system32\mmcndmgr.dll", "system32\mmres.dll", "system32\netcenter.dll", "system32\netshell.dll", "system32\networkexplorer.dll", _
                                      "system32\pnidui.dll", "system32\sensorscpl.dll", "System32\mshtml.dll", "System32\diagcpl.dll")

            IconPath = Environ("SystemRoot") & Application.PathSeparator & ImageLibraryPaths(IconLibrary)
        Else
            IconPath = IconFilePath
            IconNumber = 0
        End If

        hIcon = ExtractIcon(0&, IconPath, IconNumber)
        
        If hIcon <> 0 Then Style = Style Or ICNMB_ADDICON
        Style = IIf(BeepNotification, Style Or ICNMB_ADDBEEP, Style)
        Style = IIf(RightJustified, Style Or MB_RIGHT, Style)
        Style = Style Or MB_TOPMOST
        
                        
        If TimeOut < 60 Then TimeOut = TimeOut * 1000
        
        TargetThreadID = GetCurrentThreadId()
        pHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHookProc, 0&, TargetThreadID)
        
        If TimeOut > -1 Then
            IconMsgBox = MessageBoxTimeoutW(0, StrPtr(Content), StrPtr(Title), Style, 0, TimeOut)
        Else
            IconMsgBox = MessageBoxW(0, StrPtr(Content), StrPtr(Title), Style)
        End If
        
        DestroyIcon hIcon
        
    End Function

    Private Function MsgBoxHookProc(ByVal CodeNo As Long, ByVal wParam As LongPtr, ByVal lParam As Long) As LongPtr

        Dim ClassNameSize                           As Long
        Dim CurrWindowClassName                     As String
        
      ' Hook the process
        MsgBoxHookProc = CallNextHookEx(pHook, CodeNo, wParam, lParam)

        If CodeNo = HCBT_ACTIVATE Then
            CurrWindowClassName = Space(32)
           
          ' This function call will populate both the CurrWindowClassName and ClassNameSize variables:- 6 and #32770 respectively
            ClassNameSize = GetClassName(wParam, CurrWindowClassName, 32)
           
            If Left(CurrWindowClassName, ClassNameSize) <> MSGBOX_CLASSNAME Then Exit Function
           
          ' If hIcon has been assigned a pointer then get the handle for the STATIC control (which houses the icon),
          ' and then assign that icon to the msgbox with SendMessage - STM_SETICON
          
            If hIcon <> 0 Then
                hIconWnd = FindWindowEx(wParam, 0&, "Static", vbNullString)
                SendMessage hIconWnd, STM_SETICON, hIcon, ByVal 0&
            End If
            
          ' Unhook the process
            UnhookWindowsHookEx pHook
        
        End If

    End Function