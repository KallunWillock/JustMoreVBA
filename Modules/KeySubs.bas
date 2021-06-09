Sub PAUSE(TimePeriod As Single)
    Dim TimeNow As Single
    TimeNow = Timer
    Do
        DoEvents
    Loop Until TimeNow + TimePeriod < Timer
End Sub

Function RANDOMNUMBER(LRange As Long, URange As Long) As Long
    Randomize
    RANDOMNUMBER = CLng((URange - LRange + 1) * Rnd + LRange)
End Function
