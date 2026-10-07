Attribute VB_Name = "FastBlit"
'FastBlit.bas
'by David Finch 11-29-2000
'Updated 12-03-2000
'Updated 01-06-2002 - faster, simpler
'Quickly draws the contents of an array of 32bit,16bit, or
'15bit color values to a picturebox, stretching if necessary.
'Written in VB6, but should work in VB4 and VB5 as well.

'You are free to use FastBlit in your own programs as
'long as you list my name in the "Special Thanks."

DefLng A-Z
Option Explicit

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Type BITMAPINFOHEADER '40 bytes
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
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColor0 As Long
        bmiColor1 As Long
        bmiColor2 As Long
End Type

Public Sub Blit(buffer(), pic As PictureBox, imagewidth, imageheight)
    If imagewidth = 0 Or imageheight = 0 Then Exit Sub
    Dim bi As BITMAPINFO
    With bi.bmiHeader
        .biWidth = imagewidth
        .biHeight = -imageheight
        .biSize = 40
        .biBitCount = 32
        .biPlanes = 1
    End With
    If pic.ScaleMode <> 3 Then pic.ScaleMode = 3
    StretchDIBits pic.hdc, 0, 0, pic.ScaleWidth, pic.ScaleHeight, 0, 0, imagewidth, imageheight, buffer(0), bi, 0, vbSrcCopy
End Sub

Public Sub blit15(buffer() As Integer, pic As PictureBox, imagewidth, imageheight)
    If imagewidth = 0 Or imageheight = 0 Then Exit Sub
    Dim bi As BITMAPINFO
    With bi.bmiHeader
        .biWidth = imagewidth
        .biHeight = -imageheight
        .biSize = 40
        .biBitCount = 16
        .biPlanes = 1
    End With
    If pic.ScaleMode <> 3 Then pic.ScaleMode = 3
    StretchDIBits pic.hdc, 0, 0, pic.ScaleWidth, pic.ScaleHeight, 0, 0, imagewidth, imageheight, buffer(0), bi, 0, vbSrcCopy
End Sub

Public Sub blit16(buffer() As Integer, pic As PictureBox, imagewidth, imageheight)
    If imagewidth = 0 Or imageheight = 0 Then Exit Sub
    Dim bi As BITMAPINFO
    With bi.bmiHeader
        .biWidth = imagewidth
        .biHeight = -imageheight
        .biSize = 40
        .biBitCount = 16
        .biPlanes = 1
        .biCompression = 3
    End With
    bi.bmiColor0 = &HF800&
    bi.bmiColor1 = &H7E0&
    bi.bmiColor2 = &H1F&
    If pic.ScaleMode <> 3 Then pic.ScaleMode = 3
    StretchDIBits pic.hdc, 0, 0, pic.ScaleWidth, pic.ScaleHeight, 0, 0, imagewidth, imageheight, buffer(0), bi, 0, vbSrcCopy
End Sub

'Pretty simple, plots a point and sees how the color is approximated
Public Function getColorDepth(P As PictureBox)
    Static depth As Long
    If depth = 0 Then depth = 16
    P.PSet (0, 0), &H151515
    Select Case P.Point(0, 0)
    Case &H181418
        depth = 16
    Case &H181818
        depth = 15
    Case &H151515
        depth = 32
    End Select
    getColorDepth = depth
End Function

Public Sub CleanUp() 'dummy sub for compatibility
End Sub
