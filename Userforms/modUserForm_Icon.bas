
Attribute VB_Name = "modUserform_Icon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Lang VBA
                                                                                                                                          ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||       USERFORM - USERFORM ICON        ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
                                                                                                                                          ' _
    AUTHOR:   Kallun Willock                                                                                                              ' _
    PURPOSE:  Procedures to convert and embed icon files in userforms, and then load icon on userform initialisation.                     ' _
                                                                                                                                          ' _
    VERSION:  1.0        09/06/2021                                                                                                       ' _
                                                                                                                                          ' _
    NOTES:    This userform is an implementation of the procedures set out in a corresponding module. I                                   ' _
              wrote the code in response to a request on a forum - the OP needed to be able to share a                                    ' _
              workbook containing the userform, which itself loaded an embedded icon. I had originally                                    ' _
              responeded that it would be possible, suggesing that the icon file could be converted into                                  ' _
              Base64, stored in a variable or in a cell on a hidden sheet in the workbook, and then reconsituted                          ' _
              into an temporary icon file for loading into the userform on initialisation.                                                                                    ' _
                                                                                                                                          ' _
    TODO:     Implement Base64 conversion functionality - probably requires less storage                                                  ' _
              Rewrite hex conversion code to remove delimiter                                                                             ' _
              Fix Hex> > ICO conversion to include alpha transparency

    #If Win64 Then
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
        Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As LongPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As LongPtr
        Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As LongPtr) As LongPtr
    #Else
        Private Declare Function FindWindow Lib "user32" Alias "FindWindowA"(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        Private Declare Function ExtractIcon  Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
        Private Declare Function SendMessage" Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    #End If
    
    Private Const WM_SETICON = &H80
                                                                                                                                          ' _
    ...................................................................................................                                   ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Private Sub UserForm_Initialize()
    
    Dim IconPath As String
    Dim lngIcon As LongPtr
    Dim lnghWnd As LongPtr
    
    ' Change to the path and filename of an icon file
    ' Reconstitute the hexcode into an icon file
    IconPath = HexToIconFile(GetIconCode3)
    
    ' Get the icon from the source
    lngIcon = ExtractIcon(0, IconPath, 0)
    
    ' Get the window handle of the userform
    lnghWnd = FindWindow("ThunderDFrame", Me.Caption)
    
    'Set the big (32x32) and small (16x16) icons
    SendMessage lnghWnd, WM_SETICON, True, lngIcon
    SendMessage lnghWnd, WM_SETICON, False, lngIcon

End Sub

Private Sub UserForm_Activate()

    Me.tbFilename.ShowDropButtonWhen = 2    '   fmShowDropButtonWhenAlways
    Me.tbFilename.DropButtonStyle = 2

End Sub

Sub TestConversionToHex()
    
    Dim Filename    As String
    Dim Cll         As Range
    
    ' Select the cell with the full path to the Icon file.
    If Selection.Cells.Count = 1 Then
        Filename = Selection.Value
        If Filename = vbNullString Then Filename = Application.GetOpenFilename
        ConvertToHex Filename, Selection.Offset(0, 1)
    Else
        For Each Cll In Selection
            ConvertToHex Cll.Value, Cll.Offset(0, 1)
        Next
    End If
End Sub

Private Sub tbFilename_Change()

    ProcessFileSize Me.tbFilename.Text
    If IsValidFile(tbFilename.Text) Then
        If FileLen(tbFilename.Text) < 10000 Then
            Me.tbCode.Text = CodeToVariable(IcoToHex(tbFilename.Text))
        End If
    End If

End Sub

Private Sub tbFilename_DropButtonClick()
    
    Dim Filename        As String

    Filename = Application.GetOpenFilename
    Me.tbFilename.Text = Filename

    ProcessFileSize Filename
    
End Sub

Sub ConvertToHex(ByVal Filename As String, ByVal Target As Range)
    
    Dim FileNum As Long
    Dim IconHexCode As String
    Dim IconBytes() As Byte
    Dim IconByte As Variant
    Dim Result As VbMsgBoxResult
    
    If FileLen(Filename) > 32000 Then
        Result = MsgBox("This file is likely too large or of the incorrect format. Do you want to continue?", vbCritical + vbYesNo, "Icon file")
        If Result = vbNo Then Exit Sub
    End If
    
    FileNum = FreeFile
    
    Open Filename For Binary Access Read As #FileNum
    ReDim IconBytes(LOF(FileNum) - 1)
    Get FileNum, , IconBytes
    Close FileNum
    ' Each byte is converted to hexidecimal and separated by a '|'
    For Each IconByte In IconBytes
        IconHexCode = IconHexCode & Hex(IconByte) & "|"
    Next
    ' Remove the final '|' character from teh string
    IconHexCode = Left(IconHexCode, Len(IconHexCode) - 1)
    
    Target.Value = IconHexCode

End Sub

Function HexToIconFile(ByVal Target As Variant) As String
    
    Dim IconHexCode As String
    
    If TypeName(Target) = "Range" Then
        IconHexCode = Target.Value
    ElseIf TypeName(Target) = "String" Then
        IconHexCode = Target
    Else: Debug.Print "Error: input needs to be either a string or a range": Exit Function
    End If
    
    If Len(IconHexCode) = 32767 Then Debug.Print "Hex code is likely incomplete.": Exit Function
    If Right(IconHexCode, 1) = "|" Then IconHexCode = Left(IconHexCode, Len(IconHexCode) - 1)
    
    Dim IconBytes() As String
    IconBytes = Split(IconHexCode, "|")

    Dim FileNum As Long, Filename As String
    FileNum = FreeFile
    Filename = Environ("Temp") & "\TempUFrmIcon.ICO"

    ' Note that the code will delete any file with the same name in the Temp folder.
    If Len(Dir(Filename)) > 0 Then Kill Filename
        
    Open Filename For Binary As #FileNum
    
    Dim i As Long
    
    For i = LBound(IconBytes) To UBound(IconBytes)
        Put #FileNum, , CByte("&H" & IconBytes(i))
    Next i
    
    Close #FileNum
    
    HexToIconFile = Filename

End Function

Function IcoToHex(ByVal Filename As String) As String

    Dim FileNum As Long
    Dim IconHexCode As String
    Dim IconBytes() As Byte
    Dim IconByte As Variant
    Dim Result As VbMsgBoxResult
    
    FileNum = FreeFile
    
    Open Filename For Binary Access Read As #FileNum
    ReDim IconBytes(LOF(FileNum) - 1)
    Get FileNum, , IconBytes
    Close FileNum
    
    ' Each byte is converted to hexidecimal and separated by a '|'
    For Each IconByte In IconBytes
        IconHexCode = IconHexCode & Hex(IconByte) & "|"
    Next
    
    ' Remove the final '|' character from teh string
    IcoToHex = Left(IconHexCode, Len(IconHexCode) - 1)
    
End Function

Function IsValidFile(ByVal Filename As String) As Boolean
    
    If InStr(Filename, ":\") And Right(Filename, 4) = ".ico" Then
         IsValidFile = Len(Dir(Filename)) > 0
    End If

End Function

Sub ProcessFileSize(ByVal Filename As String)
    
    Dim Filesize        As Long
    
    If IsValidFile(Filename) Then
        Filesize = FileLen(Filename)
        Me.lbFileSize.Caption = "File size: " & vbTab & Filesize
        If Filesize > 20000 Then Me.lbFileSize.ForeColor = vbRed Else Me.lbFileSize.ForeColor = vbBlack
    End If

End Sub

Function CodeToVariable(ByVal HexCode As String) As String
    
    Dim CodeLength      As Long
    Dim MaxLength       As Long
    Dim Counter         As Long
    Dim Code            As String
    
    CodeLength = Len(HexCode)
    MaxLength = 950
    
    Code = Code & "Function GetIconCode() as string"
    Code = Code & vbNewLine & "   Dim HexCode as String" & vbNewLine
    
    For Counter = 1 To CodeLength Step MaxLength
        Code = Code & vbNewLine & "   HexCode = HexCode & " & Chr(34) & Mid(HexCode, Counter, MaxLength) & Chr(34)
    Next
    
    Code = Code & vbNewLine & "   GetIconCode = HexCode"
    Code = Code & vbNewLine & "End Function"
    
    CodeToVariable = Code

End Function

' SAMPLE ICONS

Function GetIconCode() As String
   ' USES THE GITHUB FAVICON
   Dim HexCode As String

   HexCode = HexCode & "0|0|1|0|2|0|10|10|0|0|1|0|20|0|28|5|0|0|26|0|0|0|20|20|0|0|1|0|20|0|28|14|0|0|4E|5|0|0|28|0|0|0|10|0|0|0|20|0|0|0|1|0|20|0|0|0|0|0|0|5|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|11|11|13|76|13|13|13|C5|E|E|E|12|0|0|0|0|0|0|0|0|F|F|F|11|11|11|14|B1|13|13|13|69|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|14|14|14|96|13|13|14|FC|13|13|14|ED|0|0|0|19|0|0|0|0|0|0|0|0|0|0|0|18|15|15|17|FF|15|15|17|FF|11|11|13|85|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|11|11|12|C1|13|13|14|EE|11|11|11|1E|10|10|10|10|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|D|13|13|14|F5|15|15|17|FF|15|15|17|FF|11|11|14|AF|0|0|0|0|0|0|0|0|0|0|0|0|14|14|14|99|15|15|17|FF|6|6|11|2C|E|E|E|5C|F|F|F|C1|F|F|F|22|0|0|0|0|0|0|0|0|F|F|F|34|10|10|10|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|14|14|14|8F|0|0|0|0|10|10|10|30|F|D|F|FF|0|0|0|F9|1|1|1|ED|2|2|2|FF|2|2|2|F6|E|E|E|38|0|0|0|0|0|0|0|0|8|8|8|40|2|2|2|EB|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|11|11|11|"
   HexCode = HexCode & "2D|14|14|15|9C|14|14|15|FF|1|1|1|FC|F|F|11|FB|D|D|11|3B|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|D|D|12|3A|13|13|14|E7|15|15|17|FF|15|15|17|FF|12|12|12|9A|13|13|13|D9|15|15|17|FF|15|15|17|FF|13|13|13|4F|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|11|11|11|4C|15|15|17|FF|15|15|17|FF|13|13|13|DA|13|13|14|F6|15|15|17|FF|14|14|14|F0|0|0|0|2|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|2|13|13|14|F1|15|15|17|FF|13|13|14|F6|13|13|14|F7|15|15|17|FF|14|14|14|E1|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|14|14|14|E1|15|15|17|FF|13|13|14|F7|14|14|14|DE|15|15|17|FF|13|13|14|F9|F|F|F|21|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|10|10|10|1F|13|13|14|F8|15|15|17|FF|14|14|14|DE|11|11|14|A2|15|15|17|FF|15|15|17|FF|F|F|F|34|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|10|10|10|40|15|15|17|FF|15|15|17|FF|11|11|14|A2|E|E|E|38|1"
   HexCode = HexCode & "5|15|17|FF|15|15|17|FF|12|12|12|98|0|0|0|0|0|0|0|C|0|0|0|A|0|0|0|0|0|0|0|0|0|0|0|C|0|0|0|1|0|0|0|0|12|12|12|98|15|15|17|FF|15|15|17|FF|E|E|E|38|0|0|0|0|11|11|14|A4|15|15|17|FF|11|11|12|C1|E|E|E|36|0|0|0|81|D|D|D|DC|12|12|14|D8|12|12|14|D8|13|13|14|F7|0|0|0|74|5|5|5|37|11|11|12|C1|15|15|17|FF|11|11|14|A4|0|0|0|0|0|0|0|0|0|0|0|3|13|13|13|C6|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|13|13|13|C6|0|0|0|3|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|3|11|11|14|A2|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|11|11|14|A2|0|0|0|3|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|10|10|10|3E|13|13|13|97|13|13|13|D9|12|12|14|F2|12|12|14|F2|13|13|13|D9|13|13|13|97|10|10|10|3E|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
   HexCode = HexCode & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|28|0|0|0|20|0|0|0|40|0|0|0|1|0|20|0|0|0|0|0|0|14|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|2B|C|1E|1E|1E|11|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|1B|1B|1B|1C|24|24|24|E|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|16|16|1D|23|17|17|18|92|15|15|17|F1|16|16|17|F3|40|40|40|4|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|16|16|18|ED|16|16|17|F3|16|16|18|95|1"
   HexCode = HexCode & "C|1C|1C|25|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|24|24|24|7|16|16|18|80|16|16|18|F8|15|15|17|FF|15|15|17|FF|15|15|17|FF|20|20|20|8|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|17|FE|15|15|17|FF|15|15|17|FF|16|16|18|F9|16|16|18|82|20|20|20|8|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|1B|1B|1B|1C|16|16|17|D0|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|2B|2B|2B|6|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|17|FD|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|17|D2|1A|1A|1A|1E|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|16|16|1B|2F|15|15|17|E6|15|15|17|FF|15|15|17|FC|16|16|18|B8|16|16|18|74|16|16|19|67|16|16|18|7E|55|55|55|3|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|17|FC|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|"
   HexCode = HexCode & "15|15|17|FF|15|15|17|E6|16|16|1B|2F|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|1A|1A|1A|1D|15|15|17|E6|15|15|17|FF|15|15|17|FC|18|18|18|49|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|17|FB|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|E6|1A|1A|1A|1D|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|24|24|24|7|16|16|17|D1|15|15|17|FF|15|15|17|FF|15|15|18|9D|0|0|0|0|15|15|20|18|16|16|18|73|15|15|17|90|17|17|19|66|24|24|24|7|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|1C|1C|1C|12|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|17|D1|24|24|24|7|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|16|16|18|81|15|15|17|FF|15|15|17|FF|15|15|17|F1|1B|1B|1B|1C|1C|1C|1C|25|16|16|18|EB|15|15|17|FF|15|15|17|FF|15|15|17|FF|17|17|1A|4E|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|18|18|18|40|15|15|17|FF|15|15|17|FF|1"
   HexCode = HexCode & "5|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|80|0|0|0|0|0|0|0|0|0|0|0|0|15|15|1C|24|16|16|18|F9|15|15|17|FF|15|15|18|EE|16|16|1A|45|15|15|2B|C|16|16|17|CF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|17|C4|80|80|80|2|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|18|BF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|F8|16|16|1D|23|0|0|0|0|0|0|0|0|16|16|18|94|15|15|17|FF|15|15|17|FF|16|16|17|8E|17|17|1A|5A|16|16|17|D1|15|15|17|FF|15|15|17|FF|15|15|18|E2|16|16|18|80|16|16|1A|45|1C|1C|1C|12|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|22|22|22|F|17|17|17|42|17|17|19|7B|16|16|17|DB|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|17|17|18|93|0|0|0|0|27|27|27|D|15|15|17|F2|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FE|16|16|18|82|33|33|33|5|0|0|0|0|0|0|0|0|"
   HexCode = HexCode & "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|80|80|80|2|16|16|18|74|15|15|17|FC|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|F2|15|15|2B|C|16|16|19|52|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|74|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|18|60|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|19|52|15|15|19|91|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|18|CA|FF|FF|FF|1|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|16|16|18|B7|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|19|91|16|16|18|C9|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|19|5C|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
   HexCode = HexCode & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|16|16|19|47|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|C8|16|16|18|E1|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|17|17|17|16|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|20|20|20|8|16|16|18|F8|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|E0|16|16|18|F5|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|F2|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|16|16|18|DE|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|F5|16|16|17|F3|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|DE|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|"
   HexCode = HexCode & "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|18|CA|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|17|F3|15|15|18|D9|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|F4|FF|FF|FF|1|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|16|16|18|E1|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|18|D9|15|15|18|BF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|1C|1C|1C|25|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|20|20|20|10|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|18|BF|16|16|18|95|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|76|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
   HexCode = HexCode & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|18|61|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|95|16|16|19|47|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|F4|19|19|19|1F|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|1B|1B|1B|13|16|16|18|EB|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|19|47|2B|2B|2B|6|15|15|17|F1|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|19|5D|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|18|18|18|49|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|F1|2B|2B|2B|6|0|0|0|0|16|16|18|97|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|19|19|19|33|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
   HexCode = HexCode & "|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|1A|1A|1A|1E|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|97|0|0|0|0|0|0|0|0|15|15|20|18|16|16|18|F4|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|18|18|18|35|0|0|0|0|0|0|0|0|0|0|0|0|15|15|2B|C|18|18|18|2A|80|80|80|2|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|FF|FF|FF|1|1B|1B|1B|26|1E|1E|1E|11|0|0|0|0|0|0|0|0|0|0|0|0|17|17|17|21|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|F4|15|15|20|18|0|0|0|0|0|0|0|0|0|0|0|0|16|16|18|82|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|17|17|19|66|0|0|0|0|40|40|40|4|17|17|17|62|16|16|17|E7|15|15|17|FF|16|16|17|F3|16|16|17|D2|15|15|18|C1|15|15|18|C0|16|16|17|D1|15|15|17|F0|15|15|17|FF|16|16|18|ED|15|15|18|6C|2B|2B|2B|6|0|0|0|0|16|16|19|52|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|82|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|24|24|24|7|16|16|18|C8|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|18|D6|15|15|18|A8|16|16|18|EC|15|"
   HexCode = HexCode & "15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|EF|15|15|18|AA|15|15|18|CD|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|C8|24|24|24|7|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|20|18|15|15|18|E3|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|18|E3|15|15|20|18|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|16|16|1C|2E|15|15|18|E3|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|18|E3|16|16|1C|2E|0|0|0|0|0|0|0|0|0|"
   HexCode = HexCode & "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|20|18|16|16|18|C8|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|C8|15|15|20|18|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|24|24|24|7|16|16|18|82|16|16|18|F4|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|16|16|18|F4|16|16|18|82|24|24|24|7|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|15|15|20|18|16|16|18|97|15|15|17|F1|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|FF|15|15|17|F1|16|16|18|97|15|15|20|18|0|"
   HexCode = HexCode & "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|2B|2B|2B|6|16|16|19|47|16|16|18|95|15|15|18|BF|15|15|18|D9|16|16|17|F3|16|16|17|F3|15|15|18|D9|15|15|18|BF|16|16|18|95|16|16|19|47|2B|2B|2B|6|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|"
   HexCode = HexCode & "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|"
   HexCode = HexCode & "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"

   GetIconCode = HexCode
End Function

Function GetIconCode2() As String
   
   ' COMMODORE 64 ICON
   Dim HexCode As String
   HexCode = "0|0|1|0|1|0|10|10|10|0|1|0|4|0|28|1|0|0|16|0|0|0|28|0|0|0|10|0|0|0|20|0|0|0|1|0|4|0|0|0|0|0|C0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|80|0|0|80|0|0|0|80|80|0|80|0|0|0|80|0|80|0|80|80|0|0|C0|C0|C0|0|80|80|80|0|0|0|FF|0|0|FF|0|0|0|FF|FF|0|FF|0|0|0|FF|0|FF|0|FF|FF|0|0|FF|FF|FF|0|0|0|0|0|0|0|0|0|0|0|C|CC|CC|C0|0|0|0|C|CC|CC|CC|C0|0|0|0|CC|CC|CC|CC|C0|0|0|0|CC|CC|0|0|0|0|0|C|CC|C0|0|0|0|0|0|C|CC|0|0|0|0|99|0|C|CC|0|0|0|0|0|0|C|CC|0|0|0|0|CC|0|C|CC|0|0|0|0|0|0|C|CC|C0|0|0|0|0|0|0|CC|CC|0|0|0|0|0|0|CC|CC|CC|CC|C0|0|0|0|C|CC|CC|CC|C0|0|0|0|0|C|CC|CC|C0|0|0|0|0|0|0|0|0|0|0|F8|1F|0|0|E0|F|0|0|C0|F|0|0|80|F|0|0|80|F|0|0|3|C0|0|0|7|E1|0|0|7|E3|0|0|7|E1|0|0|7|E0|0|0|3|CF|0|0|80|F|0|0|80|F|0|0|C0|F|0|0|E0|F|0|0|F8|1F|0|0"
   GetIconCode2 = HexCode

End Function

Function GetIconCode3() As String
   
   Dim HexCode As String
   HexCode = HexCode & "0|0|1|0|1|0|10|10|0|0|1|0|18|0|56|3|0|0|16|0|0|0|28|0|0|0|10|0|0|0|20|0|0|0|1|0|18|0|0|0|0|0|0|3|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FB|FB|F5|F8|F9|EF|F8|F9|EF|F8|F9|EF|F8|F9|EF|F8|F9|EF|FB|FC|F6|FF|FF|FF|FF|FF|FF|F8|F9|F0|DB|E0|B1|D8|DD|AB|D8|DD|AB|D8|DD|AB|D8|DD|AB|DB|DF|B0|D9|DD|AB|CF|D5|98|CF|D4|97|CF|D4|97|CF|D5|98|D8|DD|AB|E1|E5|BE|F8|F9|EF|FF|FF|FF|C9|D0|8A|8B|99|6|89|97|0|89|97|0|89|97|0|89|97|0|97|A3|1E|D3|D8|9F|A6|B0|3E|9C|A8|2A|9C|A8|2A|98|A4|21|AD|B7|4F|E1|E4|BD|F7|F8|ED|FF|FF|FF|B0|BA|55|91|9E|13|91|9E|13|89|97|0|89|97|0|89|97|0|97|A3|1E|DC|E0|B3|C9|CF|8B|BE|C6|73|C7|CD|86|95|A1|1B|C3|CA|7E|DF|E3|B9|F7|F8|ED|FF|FF|FF|B0|BA|55|A0|AB|31|A0|AB|31|89|97|0|89|97|0|89|97|0|97|A3|1E|CD|D3|93|BC|C4|6E|D9|DD|AC|DC|DF|B2|"
   HexCode = HexCode & "CA|D0|8C|D9|DD|AD|DF|E3|B9|F7|F8|ED|FF|FF|FF|B0|BA|55|AE|B8|51|B4|BD|5E|89|97|0|89|97|0|89|97|0|97|A3|1E|CF|D5|98|AF|B8|53|A4|B0|3D|D3|D8|A0|DC|E0|B3|DC|E0|B3|DE|E1|B7|F6|F7|EA|FF|FF|FF|B0|BA|55|B0|BA|55|CF|D4|96|89|97|0|89|97|0|89|97|0|97|A3|1E|C6|CD|83|B6|BE|62|AD|B7|4F|BF|C6|74|DC|E0|B3|DC|E0|B3|DA|DF|B0|F2|F4|E3|FF|FF|FF|B0|BA|55|B0|BA|55|F1|F3|E0|8A|97|2|89|97|0|89|97|0|91|9E|12|C7|CD|86|BD|C5|71|C9|CF|8B|C6|CD|84|D2|D8|9F|D2|D8|9F|CA|D1|8E|E7|EA|CA|FF|FF|FF|B0|BA|55|B0|BA|55|FF|FF|FE|D9|DE|AD|D8|DD|AB|D8|DD|AB|D8|DD|AB|E8|EB|CD|EB|ED|D2|EB|ED|D2|EB|ED|D2|EB|ED|D2|EB|ED|D2|E6|E9|C9|F2|F4|E3|FF|FF|FF|B0|BA|55|A3|AE|39|D8|DD|AB|D8|DD|AB|D8|DD|AB|D8|DD|AB|D8|DD|AB|D8|DD|AB|D8|DD|AB|D8|DD|AB|D8|DD|AB|D8|DD|AB|DB|E0|B1|FF|FF|FF|FF|FF|FF|FF|FF|FF|B0|BA|55|89|97|0|89|97|0|89|97|0|89|97|0|89|97|0|89|97|0|89|97|0|89|97|0|89|97|0|89|97|0|8B|99|6|B5|BD|5E|FF|FF|FF|FF|FF|FF|FF|FF|FF|C9|D0|8A|8B|99|6|89|97|0|89|97|0|89|97|0|8F|9C|E|C4|CB|80|FF|"
   HexCode = HexCode & "FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|F8|F9|EF|DB|DF|B1|D8|DD|AB|D8|DD|AB|D8|DD|AB|DF|E3|B9|F9|FA|F2|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|FF|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
   GetIconCode3 = HexCode

End Function



