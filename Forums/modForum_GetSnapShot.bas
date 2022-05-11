Attribute VB_Name = "modForums_GetSnapShot"
                                                                                                                                                                                                ' _
        |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
        ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
        ||||||||||||||||||||||||||    GETSNAPSHOT - PNG/PDF (v1.1)       ||||||||||||||||||||||||||||||||||                                                                                     ' _
        ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
        |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
                                                                                                                                                                                                ' _
        AUTHOR:     Dan_W / Kallun Willock                                                                                                                                                      ' _
        PURPOSE:    Automates the generation of a PNG / PDF capture of a given URL using the headless                                                                                           ' _
                    browser feature found in chromium-based browsers. Google set out how to go about it                                                                                         ' _
                    from the command line: https://developers.google.com/web/updates/2017/04/headless-chrome#cli                                                                                ' _
        URL:        https://www.mrexcel.com/board/excel-articles/website-snapshots.55/                                                                                                          ' _
        LICENSE:    MIT                                                                                                                                                                         ' _
        VERSION:    1.0        22/03/2022       Published on Mr Excel site                                                                                                                      ' _
                    1.1        01/05/2022       Improved scope of options / flexibility
                                                                                                                                                                                                ' _
        USAGE:      Structured as a function that returns the output filename.                                                                                                                  ' _
                                                                                                                                                                                                ' _
                    PNGSnapshot = GetSnapShot(URL, OutputPath, ScreenShotPNG, Chrome)                                                                                                           ' _
                    PDFCapture = GetSnapShot(URL, OutputPath, FullPagePDF)                                                                                                                      ' _
                                                                                                                                                                                                ' _
        NOTE:       It may take a few moments for a browser to generate the file.

    Option Explicit
    
    Enum SnapShotType
        ScreenShotPNG
        FullPagePDF
    End Enum
    
    Enum PreferredBrowser
        MSEdge
        Chrome
        Brave
    End Enum

    Const ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    
    #If VBA7 Then
        Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal DirPath As String) As Long
    #Else
        Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal DirPath As String) As Long
    #End If

    Sub TestSnapshot()

        Dim URL As String, OutputPath As String

        URL = "https://news.microsoft.com/2001/04/11/farewell-clippy-whats-happening-to-the-infamous-office-assistant-in-office-xp/"
        OutputPath = "D:\TEMP\"

        Debug.Print GetSnapShot(URL, OutputPath, "ClippyHistory", ScreenShotPNG, True, Brave, 1280, 768)

    End Sub

    Function GetSnapShot(ByVal TargetURL As String, _
                         ByVal OutputPath As String, _
                         Optional ByVal ProjectName As String = "Project", _
                         Optional ByVal Snap As SnapShotType = ScreenShotPNG, _
                         Optional ByVal DisplayFile As Boolean = True, _
                         Optional ByVal BrowserName As PreferredBrowser = Chrome, _
                         Optional ByVal WindowWidth As Long = 1280, Optional ByVal WindowHeight As Long = 768, _
                         Optional ByVal JavascriptEnabled As Boolean = True)

        Const ARGUMENT = " --headless --disable-gpu --blink-settings=scriptEnabled=%JSENABLED% --window-size=%WIDTH%,%HEIGHT% "
        
        Dim BrowserPath As String, Browser As String, Filename As String, Extension As String, Ret As Long, PID As Long
        Dim CommandLine As String, CLArguments As String, AdditionalParameter As String, TempFileName

        Browser = Switch(BrowserName = MSEdge, "msedge.exe", BrowserName = Chrome, "chrome.exe", BrowserName = Brave, "brave.exe")
        BrowserPath = DoubleQuote(GetProgramLocation(Browser))
        
        If Len(Trim(BrowserPath)) <= 2 Then
            MsgBox "Unable to locate the designated browser." & vbNewLine & "Exiting the procedure.", _
            vbCritical Or vbOKOnly, "Cannot locate browser."
            Exit Function
        End If
        
        ' The MakeSureDirectoryPathExists API will check whether a given path exists, and if not, will create the necessary directories.
        Ret = MakeSureDirectoryPathExists(OutputPath)
        Extension = IIf(Snap = FullPagePDF, ".pdf", ".png")
        AdditionalParameter = IIf(Snap = FullPagePDF, "--print-to-pdf=", " --screenshot=")
        If Len(ProjectName) > 0 Then ProjectName = ProjectName & "_"
        Filename = DoubleQuote(OutputPath & ProjectName & "SnapShot_" & Format(Now, "yyyymmdd-hhmmss") & Extension, True)
        TempFileName = VBA.Mid(Trim(Filename), 2, Len(Trim(Filename)) - 2)
        CLArguments = Replace(Replace(Replace(ARGUMENT, "%JSENABLED%", JavascriptEnabled), "%WIDTH%", WindowWidth), "%HEIGHT%", WindowHeight)
        CLArguments = CLArguments & AdditionalParameter & Filename & TargetURL
        CommandLine = BrowserPath & CLArguments
        
        ' Execute the instructions in a minimised window - other options include: vbHide, vbNormalFocus
        Shell CommandLine, vbMinimizedNoFocus
        
        GetSnapShot = TempFileName
        
        If DisplayFile Then
            Application.OnTime Now + TimeSerial(0, 0, 5), "'DisplayExportedFile " & DoubleQuote(TempFileName) & "'"
        End If
        
    End Function

    Function DoubleQuote(Optional ByVal SourceText As String, Optional ByVal TrailingSpace As Boolean = False)

        DoubleQuote = Chr(34) & SourceText & Chr(34) & IIf(TrailingSpace, Chr(32), vbNullString)

    End Function

    Function GetProgramLocation(ByVal ExeFilename As String)
        
        ' This function will check the registry for any registered applications with the given filename;
        ' it will return the full path if successful.
        Const REGISTRYADDRESS = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\"
        
        On Error Resume Next
        GetProgramLocation = CreateObject("WScript.Shell").RegRead(REGISTRYADDRESS & ExeFilename & "\")

    End Function

    Sub DisplayExportedFile(ByVal Filename As String)
    
        If Len(Dir(Filename)) Then
            Shell "Explorer.exe " & Filename, vbNormalFocus
        End If
    
    End Sub

