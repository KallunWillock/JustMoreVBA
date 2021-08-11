
    '
    '    Class used to parse Json response
    '    source http://demon.tw/my-work/vbs-json.html
    '
    '    Class VbsJson
    '    Author: Demon
    '    Date: 2012/5/3
    '    Website: http://demon.tw

    Private Whitespace, NumberRegex, StringChunk
    Private B, f, R, n, t

    Private Sub Class_Initialize()
        Whitespace = " " & vbTab & vbCr & vbLf
        B = ChrW(8)
        f = vbFormFeed
        R = vbCr
        n = vbLf
        t = vbTab

        Set NumberRegex = New regexp
        NumberRegex.pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
        NumberRegex.Global = False
        NumberRegex.MultiLine = True
        NumberRegex.IgnoreCase = True

        Set StringChunk = New regexp
        StringChunk.pattern = "([\s\S]*?)([""\\\x00-\x1f])"
        StringChunk.Global = False
        StringChunk.MultiLine = True
        StringChunk.IgnoreCase = True
    End Sub

    'Return a JSON string representation of a VBScript data structure
    'Supports the following objects and types
    '+-------------------+---------------+
    '| VBScript          | JSON          |
    '+===================+===============+
    '| Dictionary        | object        |
    '+-------------------+---------------+
    '| Array             | array         |
    '+-------------------+---------------+
    '| String            | string        |
    '+-------------------+---------------+
    '| Number            | number        |
    '+-------------------+---------------+
    '| True              | true          |
    '+-------------------+---------------+
    '| False             | false         |
    '+-------------------+---------------+
    '| Null              | null          |
    '+-------------------+---------------+
    Public Function Encode(ByRef obj)
        Dim buf, I, c, G
        Set buf = CreateObject("Scripting.Dictionary")
        Select Case VarType(obj)
            Case vbNull
                buf.ADD buf.Count, "null"
            Case vbBoolean
                If obj Then
                    buf.ADD buf.Count, "true"
                Else
                    buf.ADD buf.Count, "false"
                End If
            Case vbInteger, vbLong, vbSingle, vbDouble
                buf.ADD buf.Count, obj
            Case vbString
                buf.ADD buf.Count, """"
                For I = 1 To Len(obj)
                    c = Mid(obj, I, 1)
                    Select Case c
                        Case """"
                            buf.ADD buf.Count, "\"""
                        Case "\"
                            buf.ADD buf.Count, "\\"
                        Case "/"
                            buf.ADD buf.Count, "/"
                        Case B
                            buf.ADD buf.Count, "\b"
                        Case f
                            buf.ADD buf.Count, "\f"
                        Case R
                            buf.ADD buf.Count, "\r"
                        Case n
                            buf.ADD buf.Count, "\n"
                        Case t
                            buf.ADD buf.Count, "\t"
                        Case Else
                            If AscW(c) >= 0 And AscW(c) <= 31 Then
                                c = Right("0" & Hex(AscW(c)), 2)
                                buf.ADD buf.Count, "\u00" & c
                            Else
                                buf.ADD buf.Count, c
                            End If
                    End Select
                Next
                buf.ADD buf.Count, """"
            Case vbArray + vbVariant
                G = True
                buf.ADD buf.Count, "["
                For Each I In obj
                    If G Then G = False Else buf.ADD buf.Count, ","
                    buf.ADD buf.Count, Encode(I)
                Next
                buf.ADD buf.Count, "]"
            Case vbObject
                If TypeName(obj) = "Dictionary" Then
                    G = True
                    buf.ADD buf.Count, "{"
                    For Each I In obj
                        If G Then G = False Else buf.ADD buf.Count, ","
                        buf.ADD buf.Count, """" & I & """" & ":" & Encode(obj(I))
                    Next
                    buf.ADD buf.Count, "}"
                Else
                    err.Raise 8732, , "None dictionary object"
                End If
            Case Else
                buf.ADD buf.Count, """" & CStr(obj) & """"
        End Select
        Encode = Join(buf.Items, "")
    End Function

    '   Return the VBScript representation of ``str(``
    '   Performs the following translations in decoding
    '   +---------------+-------------------+
    '   | JSON          | VBScript          |
    '   +===============+===================+
    '   | object        | Dictionary        |
    '   +---------------+-------------------+
    '   | array         | Array             |
    '   +---------------+-------------------+
    '   | string        | String            |
    '   +---------------+-------------------+
    '   | number        | Double            |
    '   +---------------+-------------------+
    '   | true          | True              |
    '   +---------------+-------------------+
    '   | false         | False             |
    '   +---------------+-------------------+
    '   | null          | Null              |
    '   +---------------+-------------------+

    Public Function Decode(ByRef str)
        Dim idx
        idx = SkipWhitespace(str, 1)

        If Mid(str, idx, 1) = "{" Then
            Set Decode = ScanOnce(str, 1)
        Else
            Decode = ScanOnce(str, 1)
        End If
    End Function

    Private Function ScanOnce(ByRef str, ByRef idx)
        Dim c, ms

        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "{" Then
            idx = idx + 1
            Set ScanOnce = ParseObject(str, idx)
            Exit Function
        ElseIf c = "[" Then
            idx = idx + 1
            ScanOnce = ParseArray(str, idx)
            Exit Function
        ElseIf c = """" Then
            idx = idx + 1
            ScanOnce = ParseString(str, idx)
            Exit Function
        ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = Null
            Exit Function
        ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = True
            Exit Function
        ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
            idx = idx + 5
            ScanOnce = False
            Exit Function
        End If
        
        Set ms = NumberRegex.Execute(Mid(str, idx))
        If ms.Count = 1 Then
            idx = idx + ms(0).Length
            ScanOnce = CDbl(ms(0))
            Exit Function
        End If
        
        err.Raise 8732, , "No JSON object could be ScanOnced"
    End Function

    Private Function ParseObject(ByRef str, ByRef idx)
        Dim c, key, value
        Set ParseObject = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)
        
        If c = "}" Then
            Exit Function
        ElseIf c <> """" Then
            err.Raise 8732, , "Expecting property name"
        End If

        idx = idx + 1
        
        Do
            key = ParseString(str, idx)

            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) <> ":" Then
                err.Raise 8732, , "Expecting : delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            ParseObject.ADD key, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "}" Then
                Exit Do
            ElseIf c <> "," Then
                err.Raise 8732, , "Expecting , delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            c = Mid(str, idx, 1)
            If c <> """" Then
                err.Raise 8732, , "Expecting property name"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
    End Function

    Private Function ParseArray(ByRef str, ByRef idx)
        Dim c, values, value
        Set values = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "]" Then
            idx = idx + 1
            ParseArray = values.Items
            Exit Function
        End If

        Do
            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            values.ADD values.Count, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "]" Then
                Exit Do
            ElseIf c <> "," Then
                err.Raise 8732, , "Expecting , delimiter"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
        ParseArray = values.Items
    End Function

    Private Function ParseString(ByRef str, ByRef idx)
        Dim chunks, content, terminator, ms, esc, char
        Set chunks = CreateObject("Scripting.Dictionary")

        Do
            Set ms = StringChunk.Execute(Mid(str, idx))
            If ms.Count = 0 Then
                err.Raise 8732, , "Unterminated string starting"
            End If
            
            content = ms(0).submatches(0)
            terminator = ms(0).submatches(1)
            If Len(content) > 0 Then
                chunks.ADD chunks.Count, content
            End If
            
            idx = idx + ms(0).Length
            
            If terminator = """" Then
                Exit Do
            ElseIf terminator <> "\" Then
                err.Raise 8732, , "Invalid control character"
            End If
            
            esc = Mid(str, idx, 1)

            If esc <> "u" Then
                Select Case esc
                    Case """"
                        char = """"
                    Case "\"
                        char = "\"
                    Case "/"
                        char = "/"
                    Case "b"
                        char = B
                    Case "f"
                        char = f
                    Case "n"
                        char = n
                    Case "r"
                        char = R
                    Case "t"
                        char = t
                    Case Else
                        err.Raise 8732, , "Invalid escape"
                End Select
                idx = idx + 1
            Else
                char = ChrW("&H" & Mid(str, idx + 1, 4))
                idx = idx + 5
            End If

            chunks.ADD chunks.Count, char
        Loop

        ParseString = Join(chunks.Items, "")
    End Function

    Private Function SkipWhitespace(ByRef str, ByVal idx)
        Do While idx <= Len(str) And _
            InStr(Whitespace, Mid(str, idx, 1)) > 0
            idx = idx + 1
        Loop
        SkipWhitespace = idx
    End Function


