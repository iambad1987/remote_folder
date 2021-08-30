Option Explicit

Public Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
        
Private Sub VBASound()
'Call Api to play LoadIt.wav witch is in the same folder as
'the active document!
    Dim I As Integer
    For I = 1 To 4
        Call sndPlaySound32("C:\WINDOWS\Media\Alarm02.wav", 0)
        Application.Wait Now + TimeValue("00:00:01")
    Next I
End Sub

Sub SetAlarm()
    Dim alarm_time As Date
    Dim rng As Range
    Dim wks As Worksheet
    
    Set wks = Worksheets("Alarm")
    
    For Each rng In wks.Range("A1:A3")
        alarm_time = rng.Value
        Application.OnTime TimeValue(alarm_time), "VBASound"
    Next
End Sub