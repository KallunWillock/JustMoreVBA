Sub PAUSE(TimePeriod As Single)
    Dim TimeNow As Single
    TimeNow = Timer
    Do
        DoEvents
    Loop Until TimeNow + TimePeriod < Timer
End Sub