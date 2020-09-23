VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HS 
      Height          =   255
      Left            =   120
      Max             =   4
      Min             =   -5
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox A 
      Caption         =   "Antialiasing"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsDrawing%, Antialias As Boolean, nStep!

Private Sub PaintGrad(ByVal bRedraw As Boolean)
Dim cr1 As RGBQUAD, cr2 As RGBQUAD, cr&
cr = Rnd * vbWhite
CopyMemory cr1, cr, 4
cr = Rnd * vbWhite
CopyMemory cr2, cr, 4
FillGradient Picture, cr1, cr2, CInt(Rnd), 1 + CLng(Rnd) * (CLng(Rnd * 25) * 2 + 1), -bRedraw * hWnd
If Not bRedraw Then Refresh
End Sub

Private Sub DrawLines()
Dim X1!, Y1!, X2!, Y2!, I#, cr As RGBQUAD, rc As RECT
Dim hPic&, incR!, incG!, incB!, mult1!, mult2!, mX!, mY!
Caption = "Animations - Drawing ..."
IsDrawing = True
hPic = Picture
mX = ScaleWidth / 600
mY = ScaleHeight / 500
incR = 6 * (Rnd - 0.5!)
incG = 6 * (Rnd - 0.5!)
incB = 6 * (Rnd - 0.5!)
CopyMemory cr, CLng(Rnd * vbWhite), 4
mult1 = 5 * (Rnd - 0.5!)
mult2 = 5 * (Rnd - 0.5!)
Do
I = I + nStep
X1 = mX * (300 + 115 * Sin(mult2 + I / 150) + 125 * Cos(mult2 + I / 190))
X2 = mX * (300 + 115 * Sin(I * mult2 / 180) + 120 * Sin(mult2 + I / 230))
Y1 = mY * (250 + 110 * Sin(mult1 + I / 315))
Y2 = mY * (250 + 200 * Sin(4 * mult1 + I / 249))
DrawLine hPic, X1, Y1, X2, Y2, cr.rgbRed, cr.rgbGreen, cr.rgbBlue, Antialias
If X1 <= X2 Then
rc.Left = Int(X1): rc.Right = Int(X2) + 2
Else
rc.Left = Int(X2): rc.Right = Int(X1) + 2
End If
If Y1 < Y2 Then
rc.Top = Int(Y1): rc.Bottom = Int(Y2) + 2
Else
rc.Top = Int(Y2): rc.Bottom = Int(Y1) + 2
End If
RedrawWindow hWnd, rc, 0, 1
DoEvents
If Not IsDrawing Then Exit Do
Increase cr.rgbBlue, incB
Increase cr.rgbGreen, incG
Increase cr.rgbRed, incR
Loop
If IsDrawing < 1 Then Caption = "Animations - use (Ctrl or Shift or Alt) + Click"
IsDrawing = False
End Sub

Private Sub DrawCircles()
Dim cr As RGBQUAD, rc As RECT, bDib() As Byte
Dim hPic&, CntX&, CntY&, X1&, X2&, Y1&, Y2&, I&, T&, F&, A#, M!, H!, H1!
IsDrawing = True
hPic = Picture
M = IIf(ScaleWidth / 600 < ScaleHeight / 500, ScaleWidth / 600, ScaleHeight / 500)
H = Rnd * 6 - 1
CntX = ScaleWidth / 2
CntY = ScaleHeight / 2
rc.Left = CntX - 250 * M
rc.Right = CntX + 250 * M
rc.Top = CntY - 150 * M
rc.Bottom = CntY + 150 * M
GetPicture hPic, bDib
T = GetTickCount
Do
A = A + pi / 10
H = H - 0.01!
If H < -1 Then H = H + 6
For I = 50 To 1 Step -1
CntY = (ScaleHeight / 2) - Cos(A - I / 50 * pi * 10) * 45 * (1 - I / 50) * M
X1 = CntX - I * 5 * M
X2 = CntX + I * 5 * M - 1
Y1 = CntY - I * 3 * M
Y2 = CntY + I * 3 * M - 1
H1 = H + I / 50
If H1 >= 5 Then H1 = H1 - 6
HUEToRGB H1, cr.rgbRed, cr.rgbGreen, cr.rgbBlue
DrawCircle hPic, X1, Y1, X2, Y2, cr.rgbRed, cr.rgbGreen, cr.rgbBlue, Antialias
Next
RedrawWindow hWnd, rc, 0, 1
DoEvents
If Not IsDrawing Then Exit Do
SetPicture hPic, bDib
F = F + 1
I = GetTickCount
If I - T > 1000 Then
Caption = "Animations - Drawing ... " & FormatNumber(F * 1000 / (I - T), 1) & " fps"
T = I: F = 0
End If
Loop
If IsDrawing < 1 Then Caption = "Animations - use (Ctrl or Shift or Alt) + Click"
IsDrawing = False
End Sub

Private Sub DrawShifting(ByVal Roll%)
Dim bDib() As Byte, hPic&, I&, T&, F&, C&
IsDrawing = True
A.Visible = False
HS.Visible = False
hPic = Picture
If Roll = 0 Then FillNoise hPic, hWnd
GetPicture hPic, bDib
T = GetTickCount
Do
If Roll Then
C = C + (UBound(bDib) + 1) * Roll \ ScaleHeight
If C < 0 Then C = C + UBound(bDib) + 1
If C > UBound(bDib) Then C = C - (UBound(bDib) + 1)
Else
C = 1 + Rnd * (UBound(bDib) - 1)
End If
ShiftPicture hPic, C, bDib
Refresh
DoEvents
If Not IsDrawing Then Exit Do
F = F + 1
I = GetTickCount
If I - T > 1000 Then
Caption = "Animations - Drawing ... " & FormatNumber(F * 1000 / (I - T), 1) & " fps"
T = I: F = 0
End If
Loop
If IsDrawing < 1 Then
Caption = "Animations - use (Ctrl or Shift or Alt) + Click"
IsDrawing = False
A.Visible = True
HS.Visible = True
End If
End Sub

Private Sub A_Click()
Antialias = A.Value
End Sub

Private Sub Form_Load()
Randomize
nStep = 1
Antialias = True
Caption = "Animations - use (Ctrl or Shift or Alt) + Click"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
IsDrawing = False
PaintGrad True
Else
If IsDrawing Then
IsDrawing = False
Else
If Shift = vbShiftMask Then
DrawCircles
ElseIf Shift = vbCtrlMask Then
DrawShifting CInt(Rnd) * 2 - 1
ElseIf Shift = vbAltMask Then
DrawShifting 0
Else
DrawLines
End If
End If
End If
End Sub

Private Sub Form_Resize()
IsDrawing = False
Picture = CreatePicture(ScaleWidth, ScaleHeight, 24)
PaintGrad False
End Sub

Private Sub Form_Unload(Cancel As Integer)
IsDrawing = 1
End Sub

Private Sub HS_Change()
If HS.Value >= 0 Then
nStep = HS.Value + 1
Else
nStep = (10 + HS.Value) / 10
End If
End Sub
