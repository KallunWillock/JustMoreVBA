Attribute VB_Name = "modMISC"
                                                                                                                                          ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||             MISCELLANEOUS             ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
                                                                                                                                          ' _
    AUTHOR:     Kallun Willock                                                                                                            ' _
    PURPOSE:    Collection of misc. procedures                                                                                            ' _
    LICENSE:    MIT                                                                                                                                       ' _
    VERSION:    1.0         11/08/2021                                                                                                    ' _
                                                                                                                                          ' _
                                                                                                                                          ' _
    ...................................................................................................                                   ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    '  Procedures:   SpeedUp, SlowDown, and Pause
    '  Purpose:      Basic procedures at the start, the end and throughout procedures to aid performance.
    
    Sub SpeedUp()
        With Application
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .DisplayStatusBar = False
            .EnableEvents = False
        End With
    End Sub

    Sub SlowDown()
        With Application
            .Calculation = xlCalculationAutomatic
            .ScreenUpdating = True
            .DisplayStatusBar = True
            .EnableEvents = True
            .ActiveWindow.DisplayWorkbookTabs = True
        End With
    End Sub

    Sub Pause(Period As Single)
        Dim WaitTime As Single
        WaitTime = Timer + Period 
        Do
            DoEvents
        Loop While WaitTime > Timer
    End Sub
    
    '  Procedures:   CalledFromWhere
    '  Purpose:      A function to demonstarate the use of Application.Caller.
    
    Function CalledFromWhere() As String
        
        Dim rng As Range, Result As String
        On Error Resume Next
        
        Select Case TypeName(Application.Caller)
            Case "String"
                ' Assigning a macro to a shape will return the name of the shape
                Result = Application.Caller
            
            Case "Range":
                ' Calling from the Worksheet will return a range from which you can
                ' identify the address of the calling function
                Set rng = Application.Caller
                Result = "WORKSHEET @ " & Application.Caller.address(False, False)
            
            Case Else
                ' Calling direct from another procedure will return an error type
                err.Clear
                Result = "VBA"
        End Select
        
        CalledFromWhere = Result

    End Function

    Function CalledFromVBA() As Boolean
        Dim rng As Range
        On Error Resume Next
        Set rng = Application.Caller
        Err.Clear
        If rng Is Nothing Then CalledFromVBA = True Else CalledFromVBA = False
    End Function

    Function CalledFromWorksheet() As Boolean
        Dim rng As Range
        On Error Resume Next
        Set rng = Application.Caller
        Err.Clear
        If Not rng Is Nothing Then CalledFromWorksheet = True Else CalledFromWorksheet = False
    End Function



