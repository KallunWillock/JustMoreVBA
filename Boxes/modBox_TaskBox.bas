Attribute VB_Name = "modBox_TaskBox"
'@Lang VBA

                                                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                                                     ' _
    ||||||||||||||||||||||||||              TASKBOX (v1.3)           ||||||||||||||||||||||||||||||||||                                                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                                                     ' _
                                                                                                                                                                                                                            ' _
    AUTHOR:   Kallun Willock                                                                                                                                                                                                ' _
    PURPOSE:  A basic implementation of the TaskDialog, here referred to as the TaskBox.                                                                                                                                    ' _
    REFERENCE:http://vbnet.mvps.org/index.html?code/comdlg/taskdialog.htm                                                                                                                                                   ' _
    LICENSE:  MIT                                                                                                                                                                                                           ' _
    VERSION:  1.3        19/09/2023         Various corrections to the code                                                                                                                                                 ' _
              1.2        15/06/2022         Various amendments and addition of further example                                                                                                                              ' _
              1.1        07/04/2022         Improved support for Unicode text.                                                                                                                                              ' _
              1.0        18/02/2022         Version 1 uploaded to Github. Compatible with 32-bit and 64-bit Office                                                                                                          ' _
                                                                                                                                                                                                                            ' _
    NOTES:    See following for VB6 (32-bit) implementation of TaskDialogIndirect (which provides broader scope for customisation):                                                                                         ' _
              https://www.vbforums.com/showthread.php?777021-VB6-TaskDialogIndirect-Complete-class-implementation-of-Vista-Task-Dialogs                                                                                     ' _

    '   The TaskDialog allows for the standard system icons: Information, Question, Warning, Error
                                                                                                                                
    Option Explicit
    
    Public Enum TDICONS
        TD_NO_ICON = 0                              '  No icon - MainInstruction and Contents against a white background
        IDI_APPLICATION = 32512                     '  Generic icon of an application - imageres.dll - index 11
        
        TD_WARNING_ICON = -1                        '  vbExclamation equivalent
        TD_ERROR_ICON = -2                          '  vbCritical equivalent
        TD_INFORMATION_ICON = -3                    '  vbInformation equivalent
        IDI_QUESTION = 32512                        '  vbQuestion equivalent

        TD_SHIELD_ICON = -4                         '  Icon of a security shield
        TD_SHIELD_GRADIENT_ICON = -5                '  Icon of a security shield against a gradient blue/teal colour bar - default setting
        TD_SHIELD_WARNING_ICON = -6                 '  Exclamation point in shield icon against a gradient orange/yellow colour bar
        TD_SHIELD_ERROR_ICON = -7                   '  X in shield icon against a gradient red colour bar
        TD_SHIELD_OK_ICON = -8                      '  Tick in shield icon against a gradient green colour bar
        TD_SHIELD_GRAY_ICON = -9                    '  Icon of a security shield against a grey colour bar
    End Enum
    
    '   The Task Dialog allows for any combination from the common button set: OK, Yes, No, Cancel, Retry, Close
    
    Public Enum TDBUTTONS
        TDCBF_OK_BUTTON = &H1&                      '  Return: 1 (IDOK)
        TDCBF_YES_BUTTON = &H2&                     '  Return: 6 (IDYES)
        TDCBF_NO_BUTTON = &H4&                      '  Return: 7 (IDNO)
        TDCBF_CANCEL_BUTTON = &H8&                  '  Return: 2 (IDCANCEL)
        TDCBF_RETRY_BUTTON = &H10&                  '  Return: 4 (IDRETRY)
        TDCBF_CLOSE_BUTTON = &H20&                  '  Return: 8 (IDCLOSE)
    End Enum
    
    Public Enum TDBUTTONS_RETURN_CODES
        IDOK = 1
        IDCANCEL = 2
        IDRETRY = 4
        IDYES = 6
        IDNO = 7
        IDCLOSE = 8
    End Enum
    
    '  HRESULT TaskDialog(
    '    HWND                           hwndOwner,
    '    HINSTANCE                      hInstance,
    '    PCWSTR                         pszWindowTitle,
    '    PCWSTR                         pszMainInstruction,
    '    PCWSTR                         pszContent,
    '    TASKDIALOG_COMMON_BUTTON_FLAGS dwCommonButtons,
    '    PCWSTR                         pszIcon,
    '    int                            *pnButton
    '  );
    
    #If VBA7 Then
        Private Declare PtrSafe Function TaskDialog Lib "comctl32.dll" (ByVal hWndParent As LongPtr, ByVal hInstance As LongPtr, ByVal pszWindowTitle As LongPtr, ByVal pszMainInstruction As LongPtr, ByVal pszContent As LongPtr, ByVal dwCommonButtons As Long, ByVal pszIcon As LongPtr, pnButton As Long) As Long
    #Else
        Private Enum LongPtr
            [_]
        End Enum
        Private Declare Function TaskDialog Lib "comctl32.dll" (ByVal hwndParent As LongPtr, ByVal hInstance As LongPtr, ByVal pszWindowTitle As LongPtr, ByVal pszMainInstruction As LongPtr, ByVal pszContent As LongPtr, ByVal dwCommonButtons As Long, ByVal pszIcon As LongPtr, pnButton As Long) As Long
    #End If

'      :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'                                           DEMO ROUTINES

'      ...................................................................................................

    Sub TaskBox_Demo1()
        
        Dim Title               As String
        Dim MainInstruction     As String
        Dim Content             As String
        Dim Result              As TDBUTTONS_RETURN_CODES
            
        Title = "Title - TaskBox_Demo1"
        MainInstruction = "MainInstruction"
        Content = "Content" & vbNewLine & vbNewLine & "This TaskBox uses one of the five available TaskDialog colour bars:- GRADIENT"
        
        Result = TaskBox(MainInstruction, Content, Title, TDBUTTONS.TDCBF_OK_BUTTON, TDICONS.TD_SHIELD_GRADIENT_ICON)
        
        Debug.Print Result
        
    End Sub
    
    Sub TaskBox_Demo2()
        
        Dim Title               As String
        Dim MainInstruction     As String
        Dim Content             As String
        Dim Result              As TDBUTTONS_RETURN_CODES
            
        Title = "Title - TaskBox_Demo2"
        MainInstruction = "MainInstruction"
        Content = "Content" & vbNewLine & vbNewLine & "This TaskBox does not use any of the available TaskDialog colour bars, " & _
                  "but still displays the MainInstruction header together with the Information icon." & vbNewLine & vbNewLine & _
                  "All of the system icons used with the MsgBox (equivalents of each vbInformation, vbExclamation, vbCritical, vbQuestion) are available in the TaskDialog."
        
        Result = TaskBox(MainInstruction, Content, Title, TDBUTTONS.TDCBF_OK_BUTTON Or TDBUTTONS.TDCBF_CLOSE_BUTTON, TDICONS.TD_INFORMATION_ICON)
        
        Debug.Print Result
        
    End Sub
    
    Sub TaskBox_Demo3()
        
        Dim Title               As String
        Dim MainInstruction     As String
        Dim Content             As String
        Dim Result              As TDBUTTONS_RETURN_CODES
        Dim SecondResult        As TDBUTTONS_RETURN_CODES
            
        Title = "Application Process ABCD - TaskBox_Demo3"
        MainInstruction = "Do you want to proceed?"
        Content = "Please confirm whether or not you would like to proceed to the next stage of the application process."
        
        Result = TaskBox(MainInstruction, _
                         Content, Title, TDBUTTONS.TDCBF_YES_BUTTON Or TDBUTTONS.TDCBF_NO_BUTTON, TDICONS.TD_SHIELD_GRAY_ICON)
        
        Select Case Result
            Case TDBUTTONS_RETURN_CODES.IDYES
                
                SecondResult = TaskBox("Proceed to next stage...", _
                                       "You confirmed that you would like to proceed to the next stage of the application process.", _
                                       Title, TDBUTTONS.TDCBF_OK_BUTTON, TDICONS.TD_SHIELD_OK_ICON)
            
            Case TDBUTTONS_RETURN_CODES.IDNO
                
                ' Note the order of the buttons in the code and as presented on the screen. Here, the code provides for the
                ' Close button to be first, but on the screen, the first button is retry and the second is close. The buttons will
                ' be always be presented in a certain order (see the order set out above in the enumerations).
                
                SecondResult = TaskBox("Withdraw application", _
                                       "You responded that you do not wish to continue with the application process." & _
                                       vbNewLine & vbNewLine & "Please note that unless you elect to retry making an" & _
                                       " application by selecting 'Retry' below, your application will be closed.", _
                                       Title, _
                                       TDBUTTONS.TDCBF_CLOSE_BUTTON Or TDBUTTONS.TDCBF_RETRY_BUTTON, _
                                       TDICONS.TD_SHIELD_WARNING_ICON)
                
                If SecondResult = TDBUTTONS_RETURN_CODES.IDRETRY Then
                    
                    ' Retry code here
                
                ElseIf SecondResult = TDBUTTONS_RETURN_CODES.IDCLOSE Then
                
                    ' Close code here
                
                End If
                
        End Select
        
    End Sub

    Sub TaskBox_Demo4()
        
        Dim Title               As String
        Dim MainInstruction     As String
        Dim Result              As TDBUTTONS_RETURN_CODES
            
        Title = "Title - TaskBox_Demo4"
        MainInstruction = "As an alternative, the entirety of the content can be displayed as the Main Instruction."
        MainInstruction = MainInstruction & vbNewLine & vbNewLine & "This would give the appearance of a larger font messagebox."
        
        Result = TaskBox(MainInstruction, , Title, TDBUTTONS.TDCBF_OK_BUTTON, TD_NO_ICON)
        
        Debug.Print Result
        
    End Sub
'      :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'                                    MAIN TASKBOX ROUTINE

'      ...................................................................................................

    Public Function TaskBox(ByVal TaskBoxMainInstruction As String, Optional TaskBoxContent As String = "", _
                            Optional ByVal TaskBoxTitle As String = " ", _
                            Optional ByVal dwButtons As TDBUTTONS = TDCBF_OK_BUTTON, _
                            Optional ByVal lIcon As TDICONS = TD_SHIELD_GRADIENT_ICON, Optional ByVal hWndParent As LongPtr = -1) As TDBUTTONS

        Dim dwIcon              As LongPtr
        Dim pnButton            As Long
        Dim Result              As TDBUTTONS_RETURN_CODES
        
        Const IDPROMPT          As Long = &HFFFF&
        
        '  Make the IntResource
        dwIcon = IDPROMPT And lIcon
        
        '  From MSDN: "If you create a task dialog while a dialog box is present, use a handle to the dialog box as the hWndParent parameter.
        '              The hWndParent parameter should not identify a child window, such as a control in a dialog box."
        If hWndParent = -1 Then hWndParent = Application.hWnd
        Result = TaskDialog(hWndParent, 0&, StrPtr(TaskBoxTitle), StrPtr(TaskBoxMainInstruction), StrPtr(TaskBoxContent), dwButtons, dwIcon, pnButton)
    
        '   From VBNET MVPS: "The value of the button the user pressed is not returned as a result of the function call but rather as a parameter passed
        '   ByRef to the function. The return value of the call now represents success (0) or OUTOFMEMORY, INVALIDARG, or simply "FAIL"."
        TaskBox = pnButton
    
    End Function

'      :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'                                    HELPER FUNCTIONS

'      ...................................................................................................
    
    Private Function GetUnicodeMessage(ByVal UnicodeCharacters As Variant) As String
    
        Dim Counter             As Long
        Dim TempMessage         As String
        
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



