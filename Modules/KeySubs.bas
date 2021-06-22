Public Sub PAUSE(Seconds As Single)

    Dim StartTIme As Long

    StartTIme = Timer

    Do
        DoEvents
    Loop Until StartTIme + Seconds < Timer

End Subs

Function RANDOMNUMBER(LRange As Long, URange As Long) As Long
    Randomize
    RANDOMNUMBER = CLng((URange - LRange + 1) * Rnd + LRange)
End Function
