Attribute VB_Name = "modFunc"
Option Explicit

Global Const pi = 3.14159265358979

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(1) As SAFEARRAYBOUND
End Type

Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Type RGBQUAD
    rgbRed As Byte
    rgbGreen As Byte
    rgbBlue As Byte
    rgbReserved As Byte
End Type

Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)

Declare Function GetTickCount Lib "kernel32" () As Long

Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any) As Long
Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC&, pBitmapInfo As BITMAPINFO, ByVal un&, lplpVoid&, ByVal handle&, ByVal dw&) As Long

Declare Function RedrawWindow Lib "user32" (ByVal hWnd&, lprcUpdate As RECT, ByVal hrgnUpdate&, ByVal fuRedraw&) As Long

Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, IPic As IPicture) As Long

Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Sub HUEToRGB(ByVal H!, R As Byte, G As Byte, B As Byte)
Dim rR!, rG!, rB!
If H < 1 Then
rR = 1
If H < 0 Then rB = -H Else rG = H
ElseIf H < 3 Then
rG = 1
If H < 2 Then rR = 2 - H Else rB = H - 2
Else
rB = 1
If H < 4 Then rG = 4 - H Else rR = H - 4
End If
R = rR * 255: G = rG * 255: B = rB * 255
End Sub

Sub Increase(Val, incrBy)
Dim I&
I = Val
I = I + incrBy
If I < 0 Then incrBy = -incrBy: I = 0
If I > 255 Then incrBy = -incrBy: I = 255
Val = I
End Sub

Function CreatePicture(ByVal nWidth&, ByVal nHeight&, ByVal BitDepth&) As Picture
Dim Pic As PicBmp, IID_IDispatch As GUID
Dim BMI As BITMAPINFO
With BMI.bmiHeader
.biSize = Len(BMI.bmiHeader)
.biWidth = nWidth
.biHeight = nHeight
.biPlanes = 1
.biBitCount = BitDepth
End With
Pic.hBmp = CreateDIBSection(0, BMI, 0, 0, 0, 0)
IID_IDispatch.Data1 = &H20400: IID_IDispatch.Data4(0) = &HC0: IID_IDispatch.Data4(7) = &H46
Pic.Size = Len(Pic)
Pic.Type = vbPicTypeBitmap
OleCreatePictureIndirect Pic, IID_IDispatch, 1, CreatePicture
If CreatePicture = 0 Then Set CreatePicture = Nothing
End Function

Function GetPicture(ByVal Pic&, outAry() As Byte) As Boolean
Dim BM As BITMAP
GetObjectAPI Pic, Len(BM), BM
ReDim outAry(BM.bmWidthBytes * BM.bmHeight - 1)
CopyMemory outAry(0), ByVal BM.bmBits, BM.bmWidthBytes * BM.bmHeight
GetPicture = True
End Function

Function SetPicture(ByVal Pic&, inAry() As Byte) As Boolean
Dim BM As BITMAP
GetObjectAPI Pic, Len(BM), BM
If LBound(inAry) <> 0 Or UBound(inAry) <> BM.bmWidthBytes * BM.bmHeight - 1 Then Exit Function
CopyMemory ByVal BM.bmBits, inAry(0), BM.bmWidthBytes * BM.bmHeight
SetPicture = True
End Function

Function ShiftPicture(ByVal Pic&, ByVal ShiftBy&, inAry() As Byte) As Boolean
Dim bDib() As Byte, BM As BITMAP, tSA As SAFEARRAY1D
GetObjectAPI Pic, Len(BM), BM
If LBound(inAry) <> 0 Or UBound(inAry) <> BM.bmWidthBytes * BM.bmHeight - 1 Then Exit Function
With tSA
.cbElements = 1
.cDims = 1
.cElements = BM.bmWidthBytes * BM.bmHeight
.pvData = BM.bmBits
End With
CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
CopyMemory bDib(0), inAry(ShiftBy), UBound(inAry) - ShiftBy + 1
If ShiftBy Then CopyMemory bDib(UBound(inAry) - ShiftBy + 1), inAry(0), ShiftBy
CopyMemory ByVal VarPtrArray(bDib), 0&, 4
ShiftPicture = True
End Function

Function FillGradient(ByVal Pic&, cr1 As RGBQUAD, cr2 As RGBQUAD, ByVal DoVert As Boolean, ByVal Count&, Optional ByVal hRedraw&) As Boolean
Dim bDib() As Byte, BM As BITMAP, tSA As SAFEARRAY2D, rc As RECT
Dim Xmax&, Ymax&, X&, Y&, I&, R As Byte, G As Byte, B As Byte
GetObjectAPI Pic, Len(BM), BM
If BM.bmBitsPixel <> 24 Or BM.bmWidth = 1 And DoVert = True Or BM.bmHeight = 1 And DoVert = False Then Exit Function
Ymax = BM.bmHeight - 1: Xmax = (BM.bmWidth - 1) * 3
With tSA
.cbElements = 1
.cDims = 2
.Bounds(0).cElements = BM.bmHeight
.Bounds(1).cElements = BM.bmWidthBytes
.pvData = BM.bmBits
End With
CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
If DoVert Then
rc.Bottom = Ymax + 1
X = Rnd
For X = IIf(X, 0, Xmax) To IIf(X, Xmax, 0) Step X * 6 - 3
I = Xmax * (Sin(-pi / 2 + X / Xmax * pi * Count) + 1) / 2
B = (1 - I / Xmax) * cr1.rgbBlue + I / Xmax * cr2.rgbBlue
G = (1 - I / Xmax) * cr1.rgbGreen + I / Xmax * cr2.rgbGreen
R = (1 - I / Xmax) * cr1.rgbRed + I / Xmax * cr2.rgbRed
For Y = 0 To Ymax
bDib(X, Y) = B
bDib(X + 1, Y) = G
bDib(X + 2, Y) = R
Next
If hRedraw Then
rc.Left = X \ 3: rc.Right = rc.Left + 1
RedrawWindow hRedraw, rc, 0, &H101
End If
Next
Else
rc.Right = Xmax \ 3 + 1
Y = Rnd
For Y = IIf(Y, 0, Ymax) To IIf(Y, Ymax, 0) Step Y * 2 - 1
I = Ymax * (Sin(-pi / 2 + Y / Ymax * pi * Count) + 1) / 2
B = (1 - I / Ymax) * cr2.rgbBlue + I / Ymax * cr1.rgbBlue
G = (1 - I / Ymax) * cr2.rgbGreen + I / Ymax * cr1.rgbGreen
R = (1 - I / Ymax) * cr2.rgbRed + I / Ymax * cr1.rgbRed
For X = 0 To Xmax Step 3
bDib(X, Y) = B
bDib(X + 1, Y) = G
bDib(X + 2, Y) = R
Next
If hRedraw Then
rc.Top = Ymax - Y: rc.Bottom = rc.Top + 1
RedrawWindow hRedraw, rc, 0, &H101
End If
Next
End If
CopyMemory ByVal VarPtrArray(bDib), 0&, 4
FillGradient = True
End Function

Function FillNoise(ByVal Pic&, ByVal hRedraw&) As Boolean
Dim bDib() As Byte, BM As BITMAP, tSA As SAFEARRAY2D, rc As RECT
Dim Xmax&, Ymax&, X&, Y&
GetObjectAPI Pic, Len(BM), BM
If BM.bmBitsPixel <> 24 Then Exit Function
Ymax = BM.bmHeight - 1: Xmax = (BM.bmWidth - 1) * 3
With tSA
.cbElements = 1
.cDims = 2
.Bounds(0).cElements = BM.bmHeight
.Bounds(1).cElements = BM.bmWidthBytes
.pvData = BM.bmBits
End With
CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
rc.Right = Xmax \ 3 + 1
For Y = Ymax To 0 Step -1
If CInt(Rnd) Then Rnd
For X = 0 To Xmax Step 3
bDib(X, Y) = Rnd * 255
bDib(X + 1, Y) = Rnd * 255
bDib(X + 2, Y) = Rnd * 255
Next
rc.Top = Ymax - Y: rc.Bottom = rc.Top + 1
RedrawWindow hRedraw, rc, 0, &H101
Next
CopyMemory ByVal VarPtrArray(bDib), 0&, 4
FillNoise = True
End Function

Sub DrawLine(ByVal Pic&, ByVal X1!, ByVal Y1!, ByVal X2!, ByVal Y2!, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional Antialiasing As Boolean = True)
Dim bDib() As Byte, BM As BITMAP, tSA As SAFEARRAY2D
Dim Xmax&, Ymax&, dX&, dY&, X!, Y!, C!, C1!
GetObjectAPI Pic, Len(BM), BM
If BM.bmBitsPixel <> 24 Then Exit Sub
Ymax = BM.bmHeight - 1: Xmax = (BM.bmWidth - 1) * 3
Y1 = BM.bmHeight - Y1 - 1: Y2 = BM.bmHeight - Y2 - 1
With tSA
.cbElements = 1
.cDims = 2
.Bounds(0).cElements = BM.bmHeight
.Bounds(1).cElements = BM.bmWidthBytes
.pvData = BM.bmBits
End With
CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
If Antialiasing Then
If X2 < X1 Then X = X1: X1 = X2: X2 = X
If Y2 < Y1 Then Y = Y1: Y1 = Y2: Y2 = Y
If Abs(X2 - X1) > Abs(Y2 - Y1) Then
If X1 <> X2 Then
For X = Int(X1) To Int(X2) + 1
Y = (Y2 - Y1) * (X1 - X) / (X1 - X2) + Y1
dX = X * 3
dY = Int(Y)
If dX >= 0 And dY >= 0 And dX <= Xmax And dY < Ymax Then
C = Y - dY
If X = Int(X1) Then
C1 = 1 - X1 + X
ElseIf X = Int(X2) + 1 Then
C1 = 1 - X + X2
Else
C1 = 1
End If
bDib(dX, dY) = (1 - C) * C1 * B + (C * C1 + 1 - C1) * bDib(dX, dY)
bDib(dX + 1, dY) = (1 - C) * C1 * G + (C * C1 + 1 - C1) * bDib(dX + 1, dY)
bDib(dX + 2, dY) = (1 - C) * C1 * R + (C * C1 + 1 - C1) * bDib(dX + 2, dY)
bDib(dX, dY + 1) = C * C1 * B + ((1 - C) * C1 + 1 - C1) * bDib(dX, dY + 1)
bDib(dX + 1, dY + 1) = C * C1 * G + ((1 - C) * C1 + 1 - C1) * bDib(dX + 1, dY + 1)
bDib(dX + 2, dY + 1) = C * C1 * R + ((1 - C) * C1 + 1 - C1) * bDib(dX + 2, dY + 1)
End If
Next
End If
Else
If Y1 <> Y2 Then
For dY = Int(Y1) To Int(Y2) + 1
X = (X1 - X2) * (Y1 - dY) / (Y2 - Y1) + X1
dX = Int(X)
C = X - dX
dX = dX * 3
If dX >= 0 And dY >= 0 And dX < Xmax And dY <= Ymax Then
If dY = Int(Y1) Then
C1 = 1 - Y1 + dY
ElseIf dY = Int(Y2) + 1 Then
C1 = 1 - dY + Y2
Else
C1 = 1
End If
bDib(dX, dY) = (1 - C) * C1 * B + (C * C1 + 1 - C1) * bDib(dX, dY)
bDib(dX + 1, dY) = (1 - C) * C1 * G + (C * C1 + 1 - C1) * bDib(dX + 1, dY)
bDib(dX + 2, dY) = (1 - C) * C1 * R + (C * C1 + 1 - C1) * bDib(dX + 2, dY)
bDib(dX + 3, dY) = C * C1 * B + ((1 - C) * C1 + 1 - C1) * bDib(dX + 3, dY)
bDib(dX + 4, dY) = C * C1 * G + ((1 - C) * C1 + 1 - C1) * bDib(dX + 4, dY)
bDib(dX + 5, dY) = C * C1 * R + ((1 - C) * C1 + 1 - C1) * bDib(dX + 5, dY)
End If
Next
End If
End If
Else
If Abs(X2 - X1) > Abs(Y2 - Y1) Then
If X1 <> X2 Then
For X = CLng(X1) To CLng(X2) Step Sgn(X2 - X1)
dY = (Y2 - Y1) * (X1 - X) / (X1 - X2) + Y1
dX = X * 3
If dX >= 0 And dY >= 0 And dX <= Xmax And dY <= Ymax Then
bDib(dX, dY) = B: bDib(dX + 1, dY) = G: bDib(dX + 2, dY) = R
End If
Next
End If
Else
If Y1 <> Y2 Then
For dY = CLng(Y1) To CLng(Y2) Step Sgn(Y2 - Y1)
dX = CLng((X1 - X2) * (Y1 - dY) / (Y2 - Y1) + X1) * 3
If dX >= 0 And dY >= 0 And dX <= Xmax And dY <= Ymax Then
bDib(dX, dY) = B: bDib(dX + 1, dY) = G: bDib(dX + 2, dY) = R
End If
Next
End If
End If
End If
CopyMemory ByVal VarPtrArray(bDib), 0&, 4
End Sub

Sub DrawCircle(ByVal Pic&, ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional Antialiasing As Boolean = True)
Dim bDib() As Byte, BM As BITMAP, tSA As SAFEARRAY2D
Dim Xmax&, Ymax&, I&, dX&, dY&, X!, Y!, sX!, sY!, dA!, dB!, C!
GetObjectAPI Pic, Len(BM), BM
If BM.bmBitsPixel <> 24 Then Exit Sub
Ymax = BM.bmHeight - 1: Xmax = (BM.bmWidth - 1) * 3
Y1 = BM.bmHeight - Y1 - 1: Y2 = BM.bmHeight - Y2 - 1
sX = (X1 + X2) / 2: sY = (Y1 + Y2) / 2
dA = Abs(sX - X1): dB = Abs(sY - Y1)
With tSA
.cbElements = 1
.cDims = 2
.Bounds(0).cElements = BM.bmHeight
.Bounds(1).cElements = BM.bmWidthBytes
.pvData = BM.bmBits
End With
CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
If Antialiasing Then
If X1 <> X2 And dA > 0 Then
For X = X1 To X2 Step Sgn(X2 - X1)
Y = Sqr((1 - (X - sX) * (X - sX) / (dA * dA)) * dB * dB) + sY
dX = X * 3
For I = 0 To 1
dY = Int(Y)
If dX >= 0 And dY >= 0 And dX <= Xmax And dY < Ymax Then
C = Y - dY
bDib(dX, dY) = (1 - C) * B + C * bDib(dX, dY)
bDib(dX + 1, dY) = (1 - C) * G + C * bDib(dX + 1, dY)
bDib(dX + 2, dY) = (1 - C) * R + C * bDib(dX + 2, dY)
bDib(dX, dY + 1) = C * B + (1 - C) * bDib(dX, dY + 1)
bDib(dX + 1, dY + 1) = C * G + (1 - C) * bDib(dX + 1, dY + 1)
bDib(dX + 2, dY + 1) = C * R + (1 - C) * bDib(dX + 2, dY + 1)
End If
Y = sY + sY - Y
Next
Next
End If
If Y1 <> Y2 And dB > 0 Then
For dY = Y1 To Y2 Step Sgn(Y2 - Y1)
X = Sqr((1 - (dY - sY) * (dY - sY) / (dB * dB)) * dA * dA) + sX
For I = 0 To 1
dX = Int(X)
C = X - dX
dX = dX * 3
If dX >= 0 And dY >= 0 And dX < Xmax And dY <= Ymax Then
bDib(dX, dY) = (1 - C) * B + C * bDib(dX, dY)
bDib(dX + 1, dY) = (1 - C) * G + C * bDib(dX + 1, dY)
bDib(dX + 2, dY) = (1 - C) * R + C * bDib(dX + 2, dY)
bDib(dX + 3, dY) = C * B + (1 - C) * bDib(dX + 3, dY)
bDib(dX + 4, dY) = C * G + (1 - C) * bDib(dX + 4, dY)
bDib(dX + 5, dY) = C * R + (1 - C) * bDib(dX + 5, dY)
End If
X = sX + sX - X
Next
Next
End If
Else
If X1 <> X2 And dA > 0 Then
For X = X1 To X2 Step Sgn(X2 - X1)
dY = Sqr((1 - (X - sX) * (X - sX) / (dA * dA)) * dB * dB) + sY
dX = X * 3
For I = 0 To 1
If dX >= 0 And dY >= 0 And dX <= Xmax And dY <= Ymax Then
bDib(dX, dY) = B: bDib(dX + 1, dY) = G: bDib(dX + 2, dY) = R
End If
dY = sY + sY - dY
Next
Next
End If
If Y1 <> Y2 And dB > 0 Then
For dY = Y1 To Y2 Step Sgn(Y2 - Y1)
dX = CLng(Sqr((1 - (dY - sY) * (dY - sY) / (dB * dB)) * dA * dA) + sX) * 3
For I = 0 To 1
If dX >= 0 And dY >= 0 And dX <= Xmax And dY <= Ymax Then
bDib(dX, dY) = B: bDib(dX + 1, dY) = G: bDib(dX + 2, dY) = R
End If
dX = sX * 6 - dX
Next
Next
End If
End If
CopyMemory ByVal VarPtrArray(bDib), 0&, 4
End Sub
