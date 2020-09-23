VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3468
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5064
   ForeColor       =   &H0000FF00&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   422
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer OuchT 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   2880
   End
   Begin VB.Timer CloudT 
      Interval        =   3000
      Left            =   2640
      Top             =   2640
   End
   Begin VB.Timer GunReload 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   2640
   End
   Begin VB.Timer MusicTime 
      Interval        =   1000
      Left            =   2760
      Top             =   240
   End
   Begin VB.PictureBox SavePic 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1332
      Left            =   2280
      ScaleHeight     =   1332
      ScaleWidth      =   1452
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.PictureBox LoadPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1332
      Left            =   480
      ScaleHeight     =   1332
      ScaleWidth      =   1452
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Timer FramesT 
      Interval        =   1000
      Left            =   480
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Key As Byte
Public Cur As Long
Public MTime As Byte

Sub CrankItUp()
'Cur = ShowCursor(0)
ModDX7.Init Me.hwnd
End Sub

Sub ShutItDown()
'ShowCursor Cur
ModDX7.EndIt Me.hwnd
End Sub

Private Sub CloudT_Timer()
Cloud.Moving = Not Cloud.Moving
If Cloud.Moving = False Then PlaySound THUNDER
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Key = KeyCode

Select Case KeyCode

Case vbKeyEscape
  GameOver
  
Case vbKeyLeft
  If Mouldy.x > 25 Then Mouldy.x = Mouldy.x - MANSPEED
  If Mouldy.Pos Then Mouldy.Pos = 0 Else Mouldy.Pos = 1
  
Case vbKeyRight
  If Mouldy.x < 615 Then Mouldy.x = Mouldy.x + MANSPEED
  If Mouldy.Pos Then Mouldy.Pos = 0 Else Mouldy.Pos = 1

Case vbKeyUp
  If GunReload.Enabled Then Exit Sub
  PlaySound SHOT
  GunReload.Enabled = True
  Mouldy.Pos = SHOOT
  With Bullet(CurBullet)
    .x = Mouldy.x
    .y = 410
  End With
  If CurBullet = MAXBULLETS Then
    CurBullet = 0
  Else
    CurBullet = CurBullet + 1
  End If
End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Key = 0
End Sub
    
Private Sub Form_Load()
DXmidi.CrankItUp
Move 0, 0, 640, 480
ConvertEm
End Sub

Private Sub Form_Unload(Cancel As Integer)
DXmidi.ShutYerNoise
End Sub

Private Sub FramesT_Timer()
'Debug.Print "FPS = " & Frames

Frames = 0
End Sub

Sub ConvertEm()

LoadPic = LoadPicture(App.Path & "\resources\intro.jpg")
SavePic = LoadPic
SavePicture SavePic.Picture, App.Path & "\resources\intro.bmp"

LoadPic = LoadPicture(App.Path & "\resources\background.jpg")
SavePic = LoadPic
SavePicture SavePic.Picture, App.Path & "\resources\background.bmp"

End Sub

Private Sub GunReload_Timer()
GunReload.Enabled = False
Mouldy.Pos = WALK1
End Sub

Private Sub OuchT_Timer()
OuchT.Enabled = False
End Sub
