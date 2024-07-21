Attribute VB_Name = "modUserForm_KeyRoutines"
'@Lang VBA
                                                                                                                                            ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                     ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                     ' _
    ||||||||||||||||||||||||||        USERFORM - KEY ROUTINES        ||||||||||||||||||||||||||||||||||                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                     ' _
                                                                                                                                            ' _
    AUTHOR:   Kallun Willock                                                                                                                ' _
    PURPOSE:  Collection of standard Userform-related routines                                                                              ' _
                                                                                                                                            ' _
    VERSION:    1.5         11/08/2021          Added offset details, relocated shape-related routines                                      ' _
                                                in separate file, added freezeform procedure                                                ' _
                1.4         02/08/2021          Added AlwaysOnTop and Polygon-shaped userform routines                                      ' _
                1.3         22/06/2021          Fixes to 'FormTransparent' subroutine, reordering routines,                                 ' _
                                                adding userform shape-related functions                                                     ' _
                1.2         20/06/2021          Added 'SetFocusToMainApp' subroutine and further edits                                      ' _
                1.1         09/06/2021                                                                                                      ' _
                1.0         21/05/2021                                                                                                      ' _
                                                                                                                                            ' _
    
    Option Explicit
    
    Private Type RECT
        Left                As Long
        Top                 As Long
        Right               As Long
        Bottom              As Long
    End Type

    Private Type POINTAPI
        X                   As Long
        Y                   As Long
    End Type

    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
    
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
    Private Declare PtrSafe Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr
    Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hwnd As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Long) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
                                                                                                                                            
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

    Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hwndLock As LongPtr) As Long

    Private Const WM_NCLBUTTONDOWN = &HA1                           '  CONSTANTS FOR MOVING THE USERFORM
    Private Const HTCAPTION = 2
    
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOSIZE = &H1
    Private Const HWND_TOPMOST = -1

                                                                    '  CONSTANTS FOR HIDETITLEBORDER
    Private Const GWL_STYLE As Long = (-16)                         '  Window style offset
    Private Const GWL_EXSTYLE As Long = (-20)                       '  Window extended style offset
    Private Const WS_CAPTION As Long = &HC00000                     '  Titlebar
    Private Const WS_EX_DLGMODALFRAME As Long = &H1                 '  Controls if the window has an icon or not
    
    Private Const WS_SYSMENU As Long = &H80000                      '  Controls Close button
    
    Private Const WS_EX_LAYERED = &H80000                           '  CONSTANTS FOR TRANSLUCENT
    Private Const LWA_COLORKEY = &H1
    Private Const LWA_ALPHA = &H2                                                                                                         ' _
                                                                                                                                        ' _
    ...................................................................................................                                   ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    
    '  Procedure:    HideTitleBorder
    '  Purpose:      Removes the border and titlebar from the standard windows userform.
    '  Note:         Removing the title bar and the border will necessarily move all the controls on the userform.
    '                The resulting offset is Left = 6 pixels and Top = 15 pixels

    Sub HideTitleBorder(ByVal UserformCaption As String)
    
        Dim lngWindow       As LongPtr
        Dim hWndForm        As LongPtr
    
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
    
        lngWindow = GetWindowLong(hWndForm, GWL_STYLE)   '  Removes the title bar
        lngWindow = lngWindow And (Not WS_CAPTION)
        SetWindowLong hWndForm, GWL_STYLE, lngWindow
    
        lngWindow = GetWindowLong(hWndForm, GWL_EXSTYLE) '  Removes the border
        lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
        SetWindowLong hWndForm, GWL_EXSTYLE, lngWindow
    
        DrawMenuBar hWndForm
    
    End Sub
    
    Function GetHWND(UserformCaption As String) As LongPtr
    
        GetHWND = FindWindow("ThunderDFrame", UserformCaption)
        
    End Function
    
    '  Procedure:    FormOpacity
    '  Purpose:      Method of adjusting the userform opacity that changes - from fully
    '                transparent/translucent to fully opaque.
    
    Sub FormOpacity(ByVal UserformCaption As String, ByVal Opacity As Long)
    
        Dim Index           As LongPtr
        Dim hWndForm        As LongPtr
        Dim Percentage      As Double
    
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        Percentage = (255 * Opacity) / 100
        Index = GetWindowLong(hWndForm, GWL_EXSTYLE)
    
        SetWindowLong hWndForm, GWL_EXSTYLE, Index Or WS_EX_LAYERED
        SetLayeredWindowAttributes hWndForm, 0, Percentage, LWA_ALPHA
    
    End Sub
    
    '  Procedure:    FormTransparent
    '  Purpose:      Method of making a certain given colour on the userform transparent.
    
    Sub FormTransparent(ByVal UserformCaption As String, ByVal Color As Variant)
        
        Dim Index           As LongPtr
        Dim hWndForm        As LongPtr
        Dim bytOpacity      As Byte
        
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        bytOpacity = 100
        Index = GetWindowLong(hWndForm, GWL_EXSTYLE)
        
        SetWindowLong hWndForm, GWL_EXSTYLE, Index Or WS_EX_LAYERED
        SetLayeredWindowAttributes hWndForm, Color, bytOpacity, LWA_COLORKEY
    
    End Sub
    
    '  Procedure:    RemoveCloseButton
    '  Purpose:      Removes the close button, but also removes the title bar.
    '                It leaves a white bar where the title bar would have been.
    
    Sub RemoveCloseButton(ByVal UserformCaption As String)
    
        Dim hWndForm        As LongPtr
        Dim lStyle          As LongPtr

        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
    
        lStyle = GetWindowLong(hWndForm, GWL_STYE)
        SetWindowLong hWndForm, GWL_STYLE, (lStyle And Not WS_SYSMENU)
        
    End Sub

    '  Procedures:   SetFocusToMainApp
    '  Purpose:      Set focus back to the main Excel Application
    '  Note:         Remember to set the Userform's ShowModal property to False.

    Sub SetFocusToMainApp()

        Dim hWndForm        As LongPtr
        
        hWndForm = FindWindow("XLMAIN", Application.Caption)
        SetForegroundWindow hWndForm

    End Sub

    '  Procedures:   Frozen (experimental?)
    '  Purpose:      Freezes redrawing of the screen to reduce/prevent flicker.
    '  Note:         Originally in a class module, so not sure if it's wise to remove it. Important
    '                to unfreeze at the destruction of the userform.

    Sub Frozen(ByVal UserformCaption As String, ByVal State As Boolean)

        Dim hWndForm        As LongPtr

        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        LockWindowUpdate IIf(State, hWndForm, 0&)

    End Sub

                                                                                                                                          ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _
                                      USERFORM - POSITION                                                                                 ' _
    ...................................................................................................                                   ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    '  Procedure:    MoveUserForm
    '  Purpose:      Method of moving userform when there is no titlebar.
    '  Notes:        This is usually put in the userform in the MouseMove event.
    '                If it is going to sit outside the userform, the subroutine
    '                should be called conditionally - usually, by checking that
    '                the Button is being pressed - that needs to be checked at the form level.
    
    Sub MoveUserForm(ByVal UserformCaption As String)
    
        Dim Res             As LongPtr
        Dim hWndForm        As LongPtr
    
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
    
        ReleaseCapture
        Res = SendMessage(hWndForm, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
        
    End Sub
    
    '  Procedure:    AlwaysOnTop
    '  Purpose:      Method of positioning the userform above all others.
    
    Sub AlwaysOnTop(ByVal UserformCaption As String)
    
        Dim Res             As LongPtr
        Dim hWndForm        As LongPtr
    
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        Res = SetWindowPos(hWndForm, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        
    End Sub

