Attribute VB_Name = "modUI_Notifications"
'@Lang VBA
    ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||           UI - NOTIFICATIONS          ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
                                                                                                                                        ' _
    AUTHOR:   Kallun Willock                                                                                                              ' _
    PURPOSE:  API procedure to generate balloon tooltips / toast UI element                                                               ' _
                                                                                                                                        ' _
    VERSION:  1.2        12/06/2022         32-bit compatibility.                                                                         ' _
              1.1        10/08/2021         Repurpoed this module to focus on notifications.                                              ' _
              1.0        21/05/2021
                                                                                                                                        ' _
    NOTES:    Using the Toast function is relatively straight-forward. Through trial-and-error,                                           ' _
              I have worked out that:                                                                                                     ' _
                                                                                                                                        ' _
              -  TITLE:   The TITLE argument will accept a maximum of 63 characters.                                                      ' _
              -  CONTENT: The CONTENT argument will accept a maximum of 154 characters.                                                   ' _
              -  TYPE:    The first four types are accompanied with a system beep, whereas the latter four are silent.                    ' _
                                                                                                                                        ' _
    TODO:     Icons - should be able to add icons                                                                                         ' _
              Timeout - NB: uTimeout - utility?                                                                                           ' _

    Option Explicit

    #If Win64 Then
        Private Type NOTIFYICONDATA
            cbSize As Long
            hwnd As LongPtr
            uID As Long
            uFlags As Long
            uCallbackMessage As Long
            hIcon As LongPtr
            szTip As String * 128
            dwState As Long
            dwStateMask As Long
            szInfo As String * 256
            uTimeout As Long
            szInfoTitle As String * 64
            dwInfoFlags As Long
        End Type

        Private Declare PtrSafe Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, ByRef nfIconData As NOTIFYICONDATA) As LongPtr
    #Else
        Private Type NOTIFYICONDATA
            cbSize As Long
            hwnd As Long
            uID As Long
            uFlags As Long
            uCallbackMessage As Long
            hIcon As Long
            szTip As String * 128
            dwState As Long
            dwStateMask As Long
            szInfo As String * 256
            uTimeout As Long
            szInfoTitle As String * 64
            dwInfoFlags As Long
        End Type

        Private Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, ByRef nfIconData As NOTIFYICONDATA) As Long

    #End If

    Enum ToastType
        NoIcon = 0
        Information = 1
        Exclamation = 2
        Critical = 3
        Silent_NoIcon = 16
        Silent_Information = 17
        Silent_Excamation = 18
        Silent_Critical = 19
    End Enum

    Private nfIconData As NOTIFYICONDATA
                                                                                                                                ' _
    ...................................................................................................                              ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    Sub TestToast()

        Dim Title               As String
        Dim Content             As String
        Dim TestType            As ToastType

        Title = "TestToast!"
        Content = "This is some sample content."
        TestType = Information

        TOAST Title, Content, TestType

    End Sub

    Sub TOAST(Optional ByVal Title As String, Optional ByVal Content As String, Optional ByVal Flag As ToastType)

        With nfIconData
            .dwInfoFlags = Flag
            .uFlags = &H10
            .szInfoTitle = Title
            .szInfo = Content
            .cbSize = &H1F8
        End With

        Shell_NotifyIconA &H0, nfIconData
        Shell_NotifyIconA &H1, nfIconData

    End Sub
                                                                                                                                    ' _
    ...................................................................................................                                 ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    '  Procedures:   ClearStatusBar; ResetStatusBar
    '  Purpose:      Method of displaying a final message to the user before resetting the StatusBar.

    Sub ClearStatusBar(Optional Message As String, Optional DelaySeconds As Long = 0)
        Application.StatusBar = Message
        If DelaySeconds > 59 Then DelaySeconds = 59
        Application.OnTime Now + TimeValue("00:00:" & Format(DelaySeconds, "00")), "ResetStatusBar"
    End Sub

    Sub ResetStatusBar()
        Application.DisplayStatusBar = True
        Application.StatusBar = ""
    End Sub

        
