Attribute VB_Name = "modUI_Toast"

                                                                                                                                          ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||              UI - TOAST               ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
                                                                                                                                          ' _
    AUTHOR:   Kallun Willock                                                                                                              ' _
    PURPOSE:  API procedure to generate balloon tooltips / toast UI element                                                               ' _
                                                                                                                                          ' _
    VERSION:  1.0        21/05/2021                                                                                                       ' _
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
    
    Private Declare PtrSafe Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, ByRef nfIconData As NOTIFYICONDATA) As LongPtr
    
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
