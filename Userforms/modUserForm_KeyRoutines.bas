Attribute VB_Name = "modUserForm_KeyRoutines"
                                                                                                                                         ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||        USERFORM - KEY ROUTINES        ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
                                                                                                                                          ' _
    AUTHOR:   Kallun Willock                                                                                                              ' _
    PURPOSE:  Collection of standard Userform-related routines                                                                            ' _
                                                                                                                                          ' _
    VERSION:  1.1        09/06/2021                                                                                                       ' _
              1.0        21/05/2021                                                                                                       ' _
                                                                                                                                          ' _
    NOTES:    See procedure notes.                                                                                                        ' _

    Private Type RECT
        Left                                As Long
        Top                                 As Long
        Right                               As Long
        Bottom                              As Long
    End Type
    
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
    
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    
    Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hwnd As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Long) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
                                                                                                                                            
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
                                                                                                                                            
    Private Const WM_NCLBUTTONDOWN = &HA1                           '  CONSTANTS FOR MOVING THE USERFORM
    Private Const HTCAPTION = 2
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
    
    Sub HideTitleBorder(UserformCaption As String)
    
        Dim lngWindow       As LongPtr
        Dim hWndForm        As LongPtr
    
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
    
        lngWindow = GetWindow(hWndForm, GWL_STYLE)       '  Removes the title bar
        lngWindow = lngWindow And (Not WS_CAPTION)
        SetWindow hWndForm, GWL_STYLE, lngWindow
    
        lngWindow = GetWindow(hWndForm, GWL_EXSTYLE)     '  Removes the border
        lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
        SetWindow hWndForm, GWL_EXSTYLE, lngWindow
    
        DrawMenuBar hWndForm
    
    End Sub
    
    Function GetHWND(UserformCaption As String) As LongPtr
    
        GetHWND = FindWindow("ThunderDFrame", UserformCaption)
        
    End Function
    
    '  Procedure:    MoveUserForm
    '  Purpose:      Method of moving userform when there is no titlebar.
    '  Notes:        This is usually put in the userform in the MouseMove event.
    '                If it is going to sit outside the userform, the subroutine
    '                should be called conditionally - usually, by checking that
    '                the Button is being pressed - that needs to be checked at the form level.
    
    Sub MoveUserForm(UserformCaption As String)
    
        Dim Res             As LongPtr
        Dim hWndForm        As LongPtr
    
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
    
        ReleaseCapture
        Res = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
        
    End Sub
    
    '  Procedure:    FormOpacity
    '  Purpose:      Method of adjusting the userform opacity that changes - from fully
    '                transparent/translucent to fully opaque.
    
    Sub FormOpacity(UserformCaption As String, Opacity As Long)
    
        Dim Index           As LongPtr
        Dim hWndForm        As LongPtr
        Dim Percentage      As Double
    
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        Percentage = (255 * Opacity) / 100
        Index = GetWindow(hWndForm, GWL_EXSTYLE)
    
        SetWindowLong hWndForm, GWL_EXSTYLE, Index Or WS_EX_LAYERED
        SetLayeredWindowAttributes hWndForm, 0, Percentage, LWA_ALPHA
    
    End Sub
    
    '  Procedure:    FormTransparent
    '  Purpose:      Method of making a certain given colour on the userform transparent.
    
    Sub FormTransparent(UserformCaption As String, Color As Variant)
        
        Dim hWndForm        As LongPtr
        Dim bytOpacity      As Byte
        
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        bytOpacity = 100
        
        SetWindowLong hWndForm, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes hWndForm, Color, bytOpacity, LWA_COLORKEY
    
    End Sub
    
    '  Procedure:    RemoveCloseButton
    '  Purpose:      Removes the close button, but also removes the title bar.
    '                It leaves a white bar where the title bar would have been.
    
    Sub RemoveCloseButton(UserformCaption As String)
    
        Dim hWndForm            As LongPtr
        Dim lStyle              As LongPtr
       
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
    
        lStyle = GetWindow(hWndForm, GWL_STYE)
        SetWindow hWndForm, GWL_STYLE, (lStyle And Not WS_SYSMENU)
        
    End Sub
    
    '  Procedure:    GetWindowDimensions
    '  Purpose:      Gets the dimensions of a given window. Need to pass the RECT struct.
    
    Sub GetWindowDimensions(UserformCaption As String, ByRef TargetRect As RECT)
         
        ' Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
        ' Private Type RECT
        '     Left                                As Long
        '     Top                                 As Long
        '     Right                               As Long
        '     Bottom                              As Long
        ' End Type
        
        Dim hWndForm            As LongPtr
        
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        GetWindowRect hWndForm, TargetRect
        
    End Sub
    
    '  Procedure:    RoundedCorners
    '  Purpose:      Replace the corners of a given window with 'rounded corners' - set by parameters X3 and Y3.

    Sub RoundedCorners(UserformCaption As String, X3 As Long, Y3 As Long)
        
        '  CreateRoundRectRgn, SetWindowRgn, DeleteObject, CreatePolygonRgn, CreateEllipticRgn, SendMessage, ReleaseCapture
        
        Dim hWndForm            As LongPtr
        Dim DefinedRegion       As LongPtr
        Dim hWndRect            As RECT
        
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        GetWindowDimensions UserformCaption, hWndRect
        DefinedRegion = CreateRoundRectRgn(hWndRect.Left, hWndRect.Top, hWndRect.Right, hWndRect.Bottom, X3, Y3)
        SetWindowRgn hWndForm, DefinedRegion, True
        DeleteObject DefinedRegion
        
    End Sub
