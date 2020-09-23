Attribute VB_Name = "XfilesMAIN"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Type tUFO
  sx As Integer
  sy As Integer
  x As Integer
  y As Integer
  Style As Byte
End Type

Public Const DEAD = 0
Public Const THETOP = 1
Public Const THELEFT = 2
Public Const THERIGHT = 3
Public Const RANDOM = 4

Public Const MAXUFOs = 5
Public UFO(1 To MAXUFOs) As tUFO

Type tCloud
  x As Integer
  Moving As Boolean
End Type

Public Cloud As tCloud

Public Type tBullet
  x As Integer
  y As Integer
End Type

Const MAXBULLETS = 5
Public Bullet(0 To MAXBULLETS) As tBullet

Public CurBullet As Byte

Public Type tMouldy
  x As Integer
  Pos As Byte
End Type

Public Mouldy As tMouldy
Public Const WALK1 = 0
Public Const WALK2 = 1
Public Const SHOOT = 2

Public Const MANSPEED = 5

Public i, i2, x, y, x2, y2 As Integer
Public curUFO As Byte

Public Score As Integer

Public Const THUNDER = 1
Public Const SHOT = 2
Public Const OUCH = 3
Public Const EXPLODE = 4

Sub Main()
Randomize Timer
curUFO = 1

'show message
MsgBox "This is Direct X-Files, a game bodged together by Simon Price in about 2 hours. Shoot UFO's, watch out for the thunder cloud - fun huh? Note this game has no error handling so if you do not have DirectX 7 you better press Ctrl + Alt + Del now.", vbInformation, "Direct X-Files by Simon Price"

'load stuff
Form1.CrankItUp
ModDX7.SetDisplayMode 640, 480, 16
ModSurfaces.LoadAllPics

'start music
DXmidi.PlayDaTune App.Path & "\resources\x-files.mid"
'do intro screen
DoIntro
'put man in middle
Mouldy.x = 320
'enter main game loop
MainGameLoop

GameOver
End Sub

Sub DoIntro()

'show intro screen
ModDX7.SetRect SrcRect, 0, 0, 640, 480
BackBuffer.BltFast 0, 0, Intro, SrcRect, DDBLTFAST_WAIT
View.Flip Nothing, DDFLIP_WAIT
'wait for keypress
Do
DoEvents
Loop Until Form1.Key
End Sub

Sub CreateUFO(WhereFrom As Byte)
'makes a random ufo from the direction specified
If WhereFrom = RANDOM Then WhereFrom = (Rnd * 3) + 1
Select Case WhereFrom
  Case THETOP
    With UFO(curUFO)
      .sx = Int(Rnd * 540) + 50
      .sy = -Int(Rnd * 100)
      .Style = THETOP
    End With
  Case THERIGHT
    With UFO(curUFO)
      .sx = 670 + Int(Rnd * 100)
      .sy = Int(Rnd * 300) + 50
      .Style = THERIGHT
    End With
  Case THELEFT
    With UFO(curUFO)
      .sx = -30 - Int(Rnd * 100)
      .sy = Int(Rnd * 300) + 50
      .Style = THELEFT
    End With
End Select
With UFO(curUFO)
  .x = .sx
  .y = .sy
End With
End Sub

Sub MainGameLoop()
'send in some UFO's
For i = 1 To 5
  curUFO = i
  CreateUFO RANDOM
Next

Do
DoEvents
Frames = Frames + 1

    ModDX7.SetRect SrcRect, 0, 0, 640, 480
    BackBuffer.BltFast 0, 0, Background, SrcRect, DDBLTFAST_WAIT
        
    BackBuffer.SetForeColor vbGreen
    BackBuffer.DrawText 0, 0, "Score = " & Score, True
    
    DoUFOs
    DoBullets
    DoCloud
    
    BackBuffer.BltFast Mouldy.x - 25, 410, Sprites, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    View.Flip Nothing, DDFLIP_WAIT

Loop Until Key = vbKeyEscape

End Sub

Sub DoUFOs()
'move UFO's
For i = 1 To MAXUFOs
Select Case UFO(i).Style
    Case THETOP
      With UFO(i)
         .y = .y + 1
         If .y > 450 Then
           curUFO = i
           CreateUFO RANDOM
         End If
         .x = .sx + Sin(.y / 7) * 50 - 25
      End With
    Case THELEFT
      With UFO(i)
         .x = .x + 1
         If .x > 660 Then
           curUFO = i
           CreateUFO RANDOM
         End If
         .y = .sy + Sin(.x / 7) * 50 - 25
      End With
    Case THERIGHT
      With UFO(i)
         .x = .x - 1
         If .x < -20 Then
           curUFO = i
           CreateUFO RANDOM
         End If
         .y = .sy + Sin(.x / 7) * 50 - 25
      End With
End Select
Next
'draw UFO's
ModDX7.SetRect SrcRect, 150, 0, 50, 50
For i = 1 To MAXUFOs
  If UFO(i).Style Then
  BackBuffer.BltFast UFO(i).x, UFO(i).y, Sprites, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
  End If
Next
End Sub

Sub DoBullets()
Const BULLETSPEED = 7
For i2 = 0 To MAXBULLETS
  BackBuffer.SetFillColor QBColor(Int(Rnd * 14) + 1)
  If Bullet(i2).y Then
     BackBuffer.DrawCircle Bullet(i2).x, Bullet(i2).y, 5
     Bullet(i2).y = Bullet(i2).y - BULLETSPEED
     For i = 1 To MAXUFOs
       Select Case Bullet(i2).x
         Case UFO(i).x - 25 To UFO(i).x + 25
           Select Case Bullet(i2).y
             Case UFO(i).y - 15 To UFO(i).y + 15
               PlaySound EXPLODE
               curUFO = i
               CreateUFO RANDOM
               Score = Score + 50 + Int(Rnd * 100)
            End Select
       End Select
    Next
  End If
Next
End Sub

Sub DoCloud()
ModDX7.SetRect SrcRect, 200, 0, 100, 50
BackBuffer.BltFast Cloud.x - 50, 25, Sprites, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
ModDX7.SetRect SrcRect, Mouldy.Pos * 50, 0, 50, 50
If Cloud.Moving Then
  If Cloud.x > 690 Then
    Cloud.x = -50
  Else
    Cloud.x = Cloud.x + 1
  End If
Else
  DoLightning
End If

End Sub

Sub DoLightning()
If Int(Rnd * 2) Then
  BackBuffer.SetForeColor vbCyan
Else
  BackBuffer.SetForeColor vbWhite
End If
x2 = Cloud.x
For y = 40 To 460 Step 10
  x = x2 + Int(Rnd * 8) - 4
  BackBuffer.DrawLine x, y, x2, y - 10
  x2 = x
Next
Select Case Cloud.x
  Case Mouldy.x - 35 To Mouldy.x + 35
    If Int(Rnd * 2) Then
    ModDX7.SetRect SrcRect, 300, 0, 50, 50
    Else
    ModDX7.SetRect SrcRect, 350, 0, 50, 50
    End If
    Score = Score - 10
    If Score < 0 Then GameOver
    PlaySound OUCH
End Select
End Sub

Sub GameOver()
  DXmidi.ShutYerNoise
  ModSurfaces.UnloadAllPics
  ModDX7.RestoreDisplayMode
  Form1.ShutItDown
  MsgBox "Thankyou for playing this quick game by Simon Price. If you'd like to see some games which I've spent a bit of time on, then please visit my website : www.VBgames.co.uk"
  End
End Sub

Sub PlaySound(WhatSound As Byte)
Dim FileName As String

FileName = App.Path & "\resources\"
Select Case WhatSound
  Case THUNDER
     FileName = FileName & "thunder.wav"
  Case EXPLODE
     FileName = FileName & "explode.wav"
  Case SHOT
     FileName = FileName & "shoot.wav"
  Case OUCH
     If Form1.OuchT.Enabled = True Then Exit Sub
     Form1.OuchT.Enabled = True
     FileName = FileName & "ouch.wav"
End Select
sndPlaySound FileName, &H1
End Sub
