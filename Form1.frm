VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   12615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14655
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   841
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   977
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFps 
      BackColor       =   &H00000000&
      Caption         =   "FPSDISP"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   7920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code Written by: FireXtol aka DigitaIError
'Copyright © 2003

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type tLife
R As Long
G As Long
B As Long
End Type

Private Type tParticle
x As Single
y As Single
Vx As Single
Vy As Single
Life As tLife
End Type

Private Type tPart
Part() As tParticle
End Type

Private FCnt As Long
Private PCnt As Long

Private Pcles() As tPart
Private inCre As Single

Dim oPt As POINTAPI

Private Sub InitPartAry(ByVal p As Long)
Randomize
Dim Rn As Integer
Dim Rv As Integer
Dim GenVel As Single
Dim rx As Long
Dim ry As Long
Dim A As Long
Dim maxVel As Long
Randomize
rx = (ScaleWidth * Rnd) * 0.8 + (0.1 * ScaleWidth)
ry = ScaleHeight * Rnd / 2.5 + 50
Rn = Int(8 * Rnd)
Rv = Int(14 * Rnd)
A = Int(PCnt * Rnd)
For i = 0 To PCnt

A = A + 1
If A >= PCnt Then A = 0

angle = -3.1416 + (6.2832 * (i / PCnt))

Pcles(p).Part(i).x = rx
Pcles(p).Part(i).y = ry
Select Case Rv
Case Is = 0
Pcles(p).Part(i).Vx = Sin(angle) * Cos(i) * (i / (PCnt / 16))
Pcles(p).Part(i).Vy = Cos(angle) * Cos(i) * (i / (PCnt / 16))
maxVel = 16
Case Is = 1
Pcles(p).Part(i).Vx = Cos(angle) * (8 * Rnd)
Pcles(p).Part(i).Vy = Sin(angle) * (8 * Rnd)
maxVel = 8
Case Is = 2
Pcles(p).Part(i).Vx = Cos(angle) * (2 * Rnd + 7)
Pcles(p).Part(i).Vy = Sin(angle) * (2 * Rnd + 7)
maxVel = 9
Case Is = 3
Pcles(p).Part(i).Vx = Cos(angle) * (6 * Rnd)
Pcles(p).Part(i).Vy = Sin(angle) * (12 * Rnd)
maxVel = 9
Case Is = 4
Pcles(p).Part(i).Vx = Cos(angle) * (12 * Rnd)
Pcles(p).Part(i).Vy = Sin(angle) * (6 * Rnd)
maxVel = 9
Case Is = 5
Pcles(p).Part(i).Vx = Cos(angle) * (4 * Rnd + 4)
Pcles(p).Part(i).Vy = Sin(angle) * (4 * Rnd + 4)
maxVel = 8
Case Is = 6
Pcles(p).Part(i).Vx = Cos(angle) * (4 * Rnd)
Pcles(p).Part(i).Vy = Sin(angle) * (4 * Rnd)
maxVel = 4
Case Is = 7
Pcles(p).Part(i).Vx = Cos(angle) * (4 * Rnd + 10)
Pcles(p).Part(i).Vy = Sin(angle) * (4 * Rnd + 4)
maxVel = 12
Case Is = 8
Pcles(p).Part(i).Vx = Cos(angle) * (4 * Rnd + 4)
Pcles(p).Part(i).Vy = Sin(angle) * (4 * Rnd + 10)
maxVel = 12
Case Is = 9
Pcles(p).Part(i).Vx = Cos(angle) * Cos(A) * (4 * Rnd + 4)
Pcles(p).Part(i).Vy = Sin(angle) * Sin(A) * (2 * Rnd + 8)
maxVel = 9
Case Is = 10
Pcles(p).Part(i).Vx = Cos(angle) * Sin(A) * (2 * Rnd + 8)
Pcles(p).Part(i).Vy = Sin(angle) * Cos(A) * (4 * Rnd + 4)
maxVel = 9
Case Is = 11
Pcles(p).Part(i).Vx = Sin(angle) * Cos(i) * 8
Pcles(p).Part(i).Vy = Cos(angle) * Cos(i) * 8
maxVel = 8
Case Is = 12
Pcles(p).Part(i).Vx = Sin(angle) * Cos(i) * 16
Pcles(p).Part(i).Vy = Sin(angle) * Sin(i) * 16
maxVel = 16
Case Is = 13
Pcles(p).Part(i).Vx = Sin(-3.1416 + (6.2832 * (A / PCnt))) * Cos(-3.1416 + (6.2832 * (A / PCnt))) * 8 * Rnd
Pcles(p).Part(i).Vy = Cos(angle) * Sin(angle) * 8 * Rnd
maxVel = 8
End Select



GenVel = Sqr((Pcles(p).Part(i).Vx * Pcles(p).Part(i).Vx) + (Pcles(p).Part(i).Vy * Pcles(p).Part(i).Vy) + (maxVel * Rnd * maxVel * Rnd))
'Debug.Print "general velocity " & GenVel
Select Case Rn
Case Is = 0
Pcles(p).Part(i).Life.R = 255
Pcles(p).Part(i).Life.G = 255
Pcles(p).Part(i).Life.B = 192 * Rnd + 63
Case Is = 1
Pcles(p).Part(i).Life.R = 128 + 128 * Rnd '192 * Rnd + 63
Pcles(p).Part(i).Life.G = 128 + 128 * Rnd '192 * Rnd + 63
Pcles(p).Part(i).Life.B = 128 + 128 * Rnd
Case Is = 2
Pcles(p).Part(i).Life.R = 255 '192 * Rnd + 63
Pcles(p).Part(i).Life.G = 192 * Rnd + 63
Pcles(p).Part(i).Life.B = 192 * Rnd + 63
Case Is = 3
Pcles(p).Part(i).Life.R = 192 * Rnd + 63
Pcles(p).Part(i).Life.G = 255 '192 * Rnd + 63
Pcles(p).Part(i).Life.B = 192 * Rnd + 63
Case Is = 4
Pcles(p).Part(i).Life.R = 192 * Rnd + 63
Pcles(p).Part(i).Life.G = 192 * Rnd + 63
Pcles(p).Part(i).Life.B = 255
Case Is = 5
Pcles(p).Part(i).Life.R = 192 * Rnd + 63
Pcles(p).Part(i).Life.G = 255
Pcles(p).Part(i).Life.B = 255
Case Is = 6
Pcles(p).Part(i).Life.R = 255
Pcles(p).Part(i).Life.G = 192 * Rnd + 63
Pcles(p).Part(i).Life.B = 255
Case Is = 7
Pcles(p).Part(i).Life.R = 0
Pcles(p).Part(i).Life.G = 128 + 64 * Rnd
Pcles(p).Part(i).Life.B = 192 + 64 * Rnd
End Select
'If GenVel < 10 Then
Pcles(p).Part(i).Life.R = Pcles(p).Part(i).Life.R / maxVel * GenVel
Pcles(p).Part(i).Life.G = Pcles(p).Part(i).Life.G / maxVel * GenVel
Pcles(p).Part(i).Life.B = Pcles(p).Part(i).Life.B / maxVel * GenVel
'End If
Next
End Sub
Private Sub CalcMove(ByVal p As Long, ByVal i As Long)
 Pcles(p).Part(i).x = Pcles(p).Part(i).x + (Pcles(p).Part(i).Vx * inCre)
 Pcles(p).Part(i).y = Pcles(p).Part(i).y + (Pcles(p).Part(i).Vy * inCre)
 End Sub

Private Sub CalcGravity(ByVal p As Long, ByVal i As Long)
 Pcles(p).Part(i).Vx = Pcles(p).Part(i).Vx - ((Rnd * 0.01) * inCre)
 Pcles(p).Part(i).Vy = Pcles(p).Part(i).Vy + (0.05 * inCre)
End Sub

Private Sub CalcFriction(ByVal p As Long, ByVal i As Long)
 Pcles(p).Part(i).Vx = Pcles(p).Part(i).Vx / (1.08 ^ inCre)
 Pcles(p).Part(i).Vy = Pcles(p).Part(i).Vy / (1.08 ^ inCre)
End Sub

Private Sub CalcGandF(ByVal p As Long, ByVal i As Long)
 Pcles(p).Part(i).Vx = (Pcles(p).Part(i).Vx + (((Rnd - 0.5) * 0.01) * inCre)) / (1.08 ^ inCre)
 Pcles(p).Part(i).Vy = (Pcles(p).Part(i).Vy + (0.05 * inCre)) / (1.08 ^ inCre)
End Sub

Private Sub CalcLife(ByVal p As Long, ByVal i As Long)

  Pcles(p).Part(i).Life.R = Pcles(p).Part(i).Life.R - (4 * Rnd * inCre)
  Pcles(p).Part(i).Life.G = Pcles(p).Part(i).Life.G - (4 * Rnd * inCre)
  Pcles(p).Part(i).Life.B = Pcles(p).Part(i).Life.B - (4 * Rnd * inCre)
  If Pcles(p).Part(i).Life.R < 0 Then Pcles(p).Part(i).Life.R = 0
  If Pcles(p).Part(i).Life.G < 0 Then Pcles(p).Part(i).Life.G = 0
  If Pcles(p).Part(i).Life.B < 0 Then Pcles(p).Part(i).Life.B = 0
End Sub

Private Sub chkFps_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If chkFps.Value = 0 Then
 chkFps.Width = 65
Else
 chkFps.Width = 17
End If
End Sub

Private Sub Command1_Click()
Timer1.Enabled = Not Timer1.Enabled
If Timer1.Enabled = False Then
 Cls
 For i = 0 To FCnt
  InitPartAry i
 Next
 Command1.Caption = "-"
Else
 Command1.Caption = "+"
 
End If
End Sub

Private Sub Form_Load()

inCre = 1

FCnt = 19
PCnt = 199
Form1.Picture = Form1.Image
ReDim Pcles(FCnt)
For i = 0 To FCnt
 ReDim Preserve Pcles(i).Part(PCnt)
 DoEvents
 InitPartAry i
Next
'MsgBox UBound(Pcles(1).Part)
End Sub

Private Sub Form_Resize()
Picture1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
Dim Blend As BLENDFUNCTION, BlendLng As Long

Dim pHdc As Long, AvgL As Long, Dead As Long, FPSDisp As String
Static FrameCnt As Long, t As Long
Dim fr As Long, tc As Single
If FrameCnt = 0 Then t = GetTickCount - 1
Randomize
Picture1.Cls
pHdc = Picture1.hdc

            
For p = 0 To FCnt
 Dead = 0
 For i = 0 To PCnt
  'CalcGravity p, i
  'CalcFriction p, i
  CalcGandF p, i
  CalcMove p, i
  CalcLife p, i

  
  If (Pcles(p).Part(i).x > 0 And Pcles(p).Part(i).x < ScaleWidth And Pcles(p).Part(i).y > 0 And Pcles(p).Part(i).y < ScaleHeight) Then
   If (Pcles(p).Part(i).Life.R > 0 Or Pcles(p).Part(i).Life.G > 0 Or Pcles(p).Part(i).Life.B > 0) Then
    AvgL = (Pcles(p).Part(i).Life.R + Pcles(p).Part(i).Life.G + Pcles(p).Part(i).Life.B) \ 3
    SetPixelV hdc, Pcles(p).Part(i).x, Pcles(p).Part(i).y, RGB(Pcles(p).Part(i).Life.R, Pcles(p).Part(i).Life.G, Pcles(p).Part(i).Life.B)
    dc = dc + 1
    If AvgL > 223 Then SetPixelV hdc, Pcles(p).Part(i).x, Pcles(p).Part(i).y - 1, RGB(Pcles(p).Part(i).Life.R, Pcles(p).Part(i).Life.G, Pcles(p).Part(i).Life.B): dc = dc + 1
    If AvgL > 191 Then SetPixelV hdc, Pcles(p).Part(i).x, Pcles(p).Part(i).y + 1, RGB(Pcles(p).Part(i).Life.R, Pcles(p).Part(i).Life.G, Pcles(p).Part(i).Life.B): dc = dc + 1
    If AvgL > 159 Then SetPixelV hdc, Pcles(p).Part(i).x - 1, Pcles(p).Part(i).y, RGB(Pcles(p).Part(i).Life.R, Pcles(p).Part(i).Life.G, Pcles(p).Part(i).Life.B): dc = dc + 1
    If AvgL > 127 Then SetPixelV hdc, Pcles(p).Part(i).x + 1, Pcles(p).Part(i).y, RGB(Pcles(p).Part(i).Life.R, Pcles(p).Part(i).Life.G, Pcles(p).Part(i).Life.B): dc = dc + 1
  'Picture1.PSet (Pcles(p).Part(i).x, Pcles(p).Part(i).y), RGB(Pcles(p).Part(i).Life.R, Pcles(p).Part(i).Life.G, Pcles(p).Part(i).Life.B)
  
   Else
    Dead = Dead + 1
   End If
  End If
 Next
 If Dead > PCnt / 2 Then
  InitPartAry p
  Dead = 0
 End If
Next
FrameCnt = FrameCnt + 1

tc = (GetTickCount - t) / 1000
fr = CLng((FrameCnt) / (tc))
If (GetTickCount - t) > 3000 Then
Caption = FrameCnt / ((GetTickCount - t) / 1000)
t = GetTickCount - 1000 'maintain average
FrameCnt = fr
End If

If CLng(fr * inCre) < 32 Then
inCre = inCre + 0.01
ElseIf CLng(fr * inCre) > 32 Then
inCre = inCre - 0.01
End If
If fr And Abs(30 - (fr * inCre)) > 15 Then inCre = 32 / fr
'420
'If inCre > 2 Then inCre = 2

If chkFps.Value = 1 Then
FPSDisp = dc & " p/f " & Format(dc * fr, "000000") & " p/s inCre: " & Format(Round(inCre, 2), "0.00") & " " & Round(fr * inCre) & " AFPS " & fr & " FPS "
Picture1.CurrentX = Picture1.ScaleWidth - Picture1.TextWidth(FPSDisp)
Picture1.CurrentY = Picture1.ScaleHeight - Picture1.TextHeight(FPSDisp)
Picture1.Print FPSDisp
End If
Blend.SourceConstantAlpha = 64
CopyMemory BlendLng, Blend, 4

'BitBlt hdc, 0, 0, ScaleWidth, ScaleHeight, pHdc, 0, 0, vbSrcCopy
AlphaBlend hdc, 0, 0, ScaleWidth, ScaleHeight, pHdc, 0, 0, ScaleWidth, ScaleHeight, BlendLng

Refresh
End Sub
