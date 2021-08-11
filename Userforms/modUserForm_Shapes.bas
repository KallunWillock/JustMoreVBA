                                                                                                                                            ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                     ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                     ' _
    ||||||||||||||||||||||||||           USERFORM - SHAPES           ||||||||||||||||||||||||||||||||||                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                     ' _
                                                                                                                                            ' _
    AUTHOR:   Kallun Willock                                                                                                                ' _
    PURPOSE:  Collection of standard Userform shape-related routines                                                                        ' _
                                                                                                                                            ' _
    VERSION:  1.0         10/08/2021          Created new file related to userform shapes.                                                  ' _
                                                                                                                                            ' _
    NOTES:    To get the best result from adjusting the userform shape, remove the titlebar and border first - HideTitleBorder

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
                                                                                                                                            
    Private Const WM_NCLBUTTONDOWN = &HA1                           '  CONSTANTS FOR MOVING THE USERFORM
    Private Const HTCAPTION = 2
    
                                                                    '  CONSTANTS FOR HIDETITLEBORDER
    Private Const GWL_STYLE As Long = (-16)                         '  Window style offset
    Private Const GWL_EXSTYLE As Long = (-20)                       '  Window extended style offset
    Private Const WS_CAPTION As Long = &HC00000                     '  Titlebar
    Private Const WS_EX_DLGMODALFRAME As Long = &H1                 '  Controls if the window has an icon or not
    
    Private Const WS_SYSMENU As Long = &H80000                      '  Controls Close button
    
                                                                                                                                      ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _

    '  Procedure:    GetWindowDimensions
    '  Purpose:      Gets the dimensions of a given window. Need to pass the RECT struct.
    
    Sub GetWindowDimensions(UserformCaption As String, ByRef TargetRect As RECT)

        ' Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
        ' Private Type RECT
        '     Left          As Long
        '     Top           As Long
        '     Right         As Long
        '     Bottom        As Long
        ' End Type
        
        Dim hWndForm            As LongPtr
        
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        GetWindowRect hWndForm, TargetRect
        
    End Sub
    
    '  Procedure:    RoundedCorners
    '  Purpose:      Replace the corners of a given window with 'rounded corners' - set by parameters X3 and Y3.

    Sub RoundedCorners(UserformCaption As String, X3 As Long, Y3 As Long)
             
        Dim hWndForm        As LongPtr
        Dim DefinedRegion   As LongPtr
        Dim hWndRect        As RECT
        
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        GetWindowDimensions UserformCaption, hWndRect
        DefinedRegion = CreateRoundRectRgn(hWndRect.Left, hWndRect.Top, hWndRect.Right, hWndRect.Bottom, X3, Y3)
        SetWindowRgn hWndForm, DefinedRegion, True
        DeleteObject DefinedRegion
        
    End Sub

    '  Procedure:    EllipticalWindow
    '  Purpose:      Converts an existing window into an elliptical shape - using the upper-left and lower-right
    '                coordinates of the window as the bounding box coordinates.

    Sub EllipticalWindow(UserformCaption As String)
        
        Dim hWndForm        As LongPtr
        Dim DefinedRegion   As LongPtr
        Dim hWndRect        As RECT
        
        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        GetWindowDimensions UserformCaption, hWndRect
        DefinedRegion = CreateEllipticRgn(hWndRect.Left, hWndRect.Top, hWndRect.Right, hWndRect.Bottom)
        SetWindowRgn hWndForm, DefinedRegion, True
        DeleteObject DefinedRegion
        
    End Sub

    '  Procedure:    PolygonalWindow
    '  Purpose:      Converts an existing window into a polygon defined by a set of vertices.

    Sub PolygonalWindow(UserformCaption As String)
        
        Dim hWndForm        As LongPtr
        Dim DefinedRegion   As LongPtr
        Dim PolyRegn(10)    As POINTAPI

        hWndForm = FindWindow("ThunderDFrame", UserformCaption)
        '   Load the Polygon shape data
        
        ThunderAnd 3

        DefinedRegion = CreatePolygonRgn(PolyShape(0), 1 + UBound(PolyShape), 1)
        SetWindowRgn hWndForm, DefinedRegion, True
        DeleteObject DefinedRegion

    End Sub

    Sub ThunderAnd(Optional SizeScale As Single = 2)
    
        PolyShape(0).X = 95: PolyShape(0).Y = 105
        PolyShape(1).X = 202: PolyShape(1).Y = 48
        PolyShape(2).X = 263: PolyShape(2).Y = 137
        PolyShape(3).X = 240: PolyShape(3).Y = 144
        PolyShape(4).X = 307: PolyShape(4).Y = 209
        PolyShape(5).X = 307: PolyShape(5).Y = 209
        PolyShape(6).X = 283.5: PolyShape(6).Y = 219
        PolyShape(7).X = 367.5: PolyShape(7).Y = 328.5
        PolyShape(8).X = 213.5: PolyShape(8).Y = 242
        PolyShape(9).X = 240.5: PolyShape(9).Y = 229.5
        PolyShape(10).X = 151.5: PolyShape(10).Y = 174
        PolyShape(11).X = 183: PolyShape(11).Y = 159
        PolyShape(12).X = 99.5: PolyShape(12).Y = 112
        For i = 0 To 12
            PolyShape(i).X = PolyShape(i).X * SizeScale
            PolyShape(i).Y = PolyShape(i).Y * SizeScale
        Next

    End Sub


