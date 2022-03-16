Attribute VB_Name = "modBox_IconMsgBox"
                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||             ICONMSGBOX (v1.1)         ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
                                                                                                                                                                                            ' _
    AUTHOR:   Kallun Willock                                                                                                                                                                ' _
    PURPOSE:  The IconMsgBox is a Unicode-enabled MessageBox that can display a custom icon (ICO file).                                                                                     ' _
              It also features a timeout feature resulting in it closing down after a designated period of time.                                                                            ' _
    LICENSE:  MIT                                                                                                                                                                           ' _
    VERSION:  1.1        16/03/2022         Added timeout functiontionality. Improved comments and corrected errors.                                                                        ' _
              1.0        18/02/2022         Version 1 uploaded to Github. Compatible with 32-bit and 64-bit Office                                                                          ' _
                                                                                                                                                                                            ' _
    NOTES:    A return value of 32000 indicates that the user did not press a button.                                                                                                       ' _
                                                                                                                                                                                            ' _
    TODO:     [X] Add access to system DLL icons                                                                                                                                            ' _
              [x] Allow use of custom ICO files                                                                                                                                             ' _
              [x] Unicode compatibility                                                                                                                                                     ' _
              [x] Add timeout feature                                                                                                                                                       ' _
              [ ] Custom buttons labels?

            Option Explicit

            Public Enum ImageDLL
                icn_shell32
                icn_imageres
                icn_accessibilitycpl
                icn_ddores
                icn_moricons
                icn_pifmgr
                icn_explorer
                icn_mmcndmgr
                icn_mmres
                icn_netcenter
                icn_netshell
                icn_networkexplorer
                icn_pnidui
                icn_sensorscpl
            End Enum
        
        '        C:\Windows\System32\accessibilitycpl.dll       - 24    icons
        '        C:\Windows\System32\ActionCenterCPL.dll        - 7     icons
        '        C:\Windows\System32\mshtml.dll                 - 27    icons
        '        C:\Windows\System32\taskbarcpl.dll             - 15    icons
        '        C:\Windows\System32\powercpl.dll               - 6     icons
        '        C:\Windows\System32\Diagcpl.dll                - 9     icons
        '        C:\Windows\System32\Usercpl.dll                - 1     icon
        '        C:\Windows\System32\Themecpl.dll               - 2     icons
        '        C:\Windows\System32\iscsicpl.dll               - 2     icons
        '        C:\Windows\System32\sdcpl.dll                  - 6     icons
        '        C:\Windows\System32\hgcpl.dll                  - 4     icons
        '        C:\Windows\System32\fhcpl.dll                  - 1     icon
        '        C:\Windows\System32\fvecpl.dll                 - 1     icon
        '        C:\Windows\System32\werconcpl.dll              - 5     icons
        '        C:\Windows\System32\user32.dll                 - 7     icons
        '        C:\Windows\System32\shell32.dll                - 329   icons
        
            #If VBA7 Then

                Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
                Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
                Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInstance As LongPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As LongPtr
                Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
                Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
                Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
                Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
                Private Declare PtrSafe Function MessageBoxTimeoutW Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long, ByVal wLanguageID As Long, ByVal lngMilliseconds As Long) As Long
                Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
                Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
                Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long

                Private pHook As LongPtr
                Private hIcon As LongPtr

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

                Private pHook As Long
                Private hIcon As Long

            #End If
            
            Private Const WH_CBT                            As Long = &H5
            Private Const HCBT_ACTIVATE                     As Long = &H5
            Private Const STM_SETICON                       As Long = &H170
            Private Const MSGBOX_CLASSNAME                  As String = "#32770"

            Sub IconMsgBox_Demo1()

                Dim Content                                 As String
                Dim Title                                   As String
                Dim IconFilePath                            As String

                Title = "Title - IconMsgBox_Demo1"
                Content = "Is this a useless example question?" & vbNewLine & vbNewLine & "Anything is better than the previous example, OK? This will timeout in 5 seconds."
                IconFilePath = "D:\discord1.ico"

                Debug.Print IconMsgBox(Content, vbQuestion + vbYesNo + vbDefaultButton1, Title, IconFilePath, icn_imageres, 25, 10000)

            End Sub

            Public Function IconMsgBox(ByVal Content As String, Optional ByVal Style As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal IconFilePath As String, Optional ByVal IconLibrary As ImageDLL, Optional ByVal IconNumber As Long = 0, Optional TimeOut As Long = -1, Optional BeepNotification As Boolean = False) As VbMsgBoxResult

                Dim IconPath                                As String
                Dim TargetThreadID                          As Long
                
                If Len(Dir(IconFilePath)) = 0 Then IconFilePath = vbNullString
                
                If IconFilePath = vbNullString Then
                    Dim ImageLibraryPaths                   As Variant
                    ImageLibraryPaths = Array("C:\WINDOWS\System32\shell32.dll", "C:\Windows\system32\imageres.dll", _
                                              "C:\Windows\system32\pifmgr.dll", "C:\Windows\system32\accessibilitycpl.dll", _
                                              "C:\Windows\system32\ddores.dll", "C:\Windows\system32\moricons.dll", _
                                              "C:\windows\explorer.exe", "C:\windows\system32\mmcndmgr.dll", _
                                              "C:\windows\system32\mmres.dll", "C:\windows\system32\netcenter.dll", _
                                              "C:\windows\system32\netshell.dll", "C:\windows\system32\networkexplorer.dll", _
                                              "C:\windows\system32\pnidui.dll", "C:\windows\system32\sensorscpl.dll")

                    IconPath = ImageLibraryPaths(IconLibrary)
                Else
                    IconPath = IconFilePath
                    IconNumber = 0
                End If

                hIcon = ExtractIcon(0&, IconPath, IconNumber)
                
                If BeepNotification = True Then Style = Style - vbCritical
                
                If TimeOut < 60 Then TimeOut = TimeOut * 1000
                
                TargetThreadID = GetCurrentThreadId()
                pHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHookProc, 0^, TargetThreadID)
                
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
                Dim hIconWnd                                As LongPtr

              ' Hook the process
                MsgBoxHookProc = CallNextHookEx(pHook, CodeNo, wParam, lParam)

                If CodeNo = HCBT_ACTIVATE Then
                    CurrWindowClassName = Space(32)
                   
                  ' This function call will populate both the CurrWindowClassName and ClassNameSize variables:- 6 and #32770 respectively
                    ClassNameSize = GetClassName(wParam, CurrWindowClassName, 32)
                   
                    If Left(CurrWindowClassName, ClassNameSize) <> MSGBOX_CLASSNAME Then Exit Function
                   
                  ' If phIcon has been assigned a pointer then get the handle for the STATIC control (which houses the icon), and then
                  ' assign that icon to the msgbox with SendMessage - STM_SETICON
                  
                    If hIcon <> 0 Then
                        hIconWnd = FindWindowEx(wParam, 0^, "Static", vbNullString)
                        SendMessage hIconWnd, STM_SETICON, hIcon, ByVal 0^
                    End If
                    
                  ' Unhook the process
                    UnhookWindowsHookEx pHook
                
                End If

            End Function
