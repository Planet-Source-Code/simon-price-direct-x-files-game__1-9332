Attribute VB_Name = "DXmidi"
'DX midi -by Simon Price

'This module I have made to save myself time
'when making games. Now, to play music (midi only)
'in my games I just have to shove in this
'module and call the play and stop subs
'when I want. you can use it in your games too,
'but if you do, leave the module how it is and
'give me due credit

Dim DX As New DirectX7
Dim Perf As DirectMusicPerformance
Dim Perf2 As DirectMusicPerformance
Dim Seg As DirectMusicSegment
Dim SegState As DirectMusicSegmentState
Dim Loader As DirectMusicLoader
Public GetStartTime As Long
Public Offset As Long
Public mtTime As Long
Public mtLength As Double
Public dTempo As Double
Dim TimeSig As DMUS_TIMESIGNATURE
Dim IsPlayingCheck As Boolean
Dim Time As Double
Dim fIsPaused As Boolean

Sub CrankItUp()

Set Loader = DX.DirectMusicLoaderCreate()

Set Perf2 = DX.DirectMusicPerformanceCreate()
Perf2.Init Nothing, 0
Perf2.SetPort -1, 80
Perf2.GetMasterAutoDownload

Set Perf = DX.DirectMusicPerformanceCreate()
Perf.Init Nothing, 0
Perf.SetPort -1, 80
Perf.SetMasterAutoDownload True

End Sub

Sub PlayDaTune(Filename As String)
Set Seg = Loader.LoadSegment(Filename)
Seg.SetStartPoint 0
Set SegState = Perf.PlaySegment(Seg, 0, 0)
End Sub

Sub ShutYerNoise()
  Perf.Stop Seg, SegState, 0, 0
End Sub

Sub SetVolume(Volume As Integer)
  Perf.SetMasterVolume Volume
End Sub
