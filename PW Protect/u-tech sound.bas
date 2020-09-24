Attribute VB_Name = "utech_sound"

'feel free to change any part of this project
'just give credit where credit is due

'feeback or comments
'email:  u_tech@ excite.com

Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
Public Function RndNum(Min As Long, Max As Long) As Long
Attribute RndNum.VB_Description = "Returns a random number between the min and max variable values.                                                                "
Dim X As Long
100
     Randomize
     X = Int(((Max + 1) * Rnd))
If X < Min Then GoTo 100
     RndNum = X
End Function
Public Function PlayWav(WavFile As String)
Attribute PlayWav.VB_Description = "Simplified function to play a non looping wave file"
Dim Flags As Long
Dim X As Long
Flags = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(WavFile, Flags)
End Function

