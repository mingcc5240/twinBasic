Attribute VB_Name = "Grfx"
'This is a part of tha BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'(I know the emulator is NOT OPTIMIZED AT ALL)

'Graphic Engine
'Currenlty a bit optimized
'Color Gameboy Part is 95% finished
'Coments will be added with the next release

'Sory for my bad english ...

Option Base 0
Option Explicit

Public FPS As Long, fpsT As Long
Dim colid(128, 128) As Long
Dim colidx(128, 128) As Byte
Dim colid2(2, 128, 128) As Long
'Public DDraw As clsDirectDraw7
'Public scrb As clsDDSurface7
'Dim re2 As DxVBLib.RECT
Public FULLSCREEN As Boolean
Dim TCol As Long, mode1 As Boolean
Dim i As Long, i2 As Long, x As Long, y As Long, tilemap As Long, tileloc As Long, tileptr As Long
Dim xoffset As Long, yoffset As Long, TileData As Long, tileend As Long
Dim LByte As Long, HByte As Long, SpriteY As Long
Dim dy As Long, dx As Long, spat As Long
Dim tmp1 As Long, tmp2 As Long, memptr As Long, tmp3 As Long, rx As Long, ry As Long, lolm As Long, tms As Long
Global bb As BITMAPINFO, desthdc As Long, destW As Long, destH As Long, desthimg As Long
Global Vram(0 To 159, 0 To 143) As Integer, mir(255) As Byte, mv1 As Byte, mv2 As Byte, vf As Boolean, cid2 As Long, tiletmp As Long, xs As Long, ys As Long, xt As Long
Global tobj As Long, thdc As Long, dh As Long, dw As Long, xr As Long, yr As Long
Global fskip As Long, fmode As Long, objp(7, 3) As Integer, bgp(7, 3) As Integer
Global vrm As Long, ccp As Long, ccid(128, 128) As Long, tm2 As Long, tm1 As Long, bgat As Byte, gbcP(32767) As Integer
Global wv As Boolean, bgv As Boolean, objv As Boolean, xflip As Long, yflip As Long, lastline As Long
Global curline As Long, tcls(19) As Integer, curFreq As Currency, curStart As Currency, CurEnd As Currency, dblResult As Double
Global Skipf As Boolean

Public Const HALFTONE As Long = 4
Public Const COLORONCOLOR As Long = 3



' 2¹è È®´ë¿ë ºñÆ®¸Ê Çì´õ ¹× 320x288 ¹öÆÛ
Global bb2x As BITMAPINFO
Global Vram2x(0 To 319, 0 To 287) As Integer
Global Scale2x As Boolean

Public Sub InitScale2xHeader()
    With bb2x.Header
        .biSize = Len(bb2x.Header)
        .biWidth = 320
        ' Top-Down DIB ¹æ½ÄÀ¸·Î »óÇÏ ¹ÝÀüÀ» ¹æÁöÇÏ·Á¸é À½¼ö(-288)¸¦ ÁöÁ¤ÇÕ´Ï´Ù.
        ' ±âÁ¸ bb ±¸Á¶Ã¼ÀÇ biHeight°¡ À½¼ö¶ó¸é -288, ¾ç¼ö¶ó¸é 288·Î ¸ÂÃçÁÝ´Ï´Ù.
        .biHeight = -288
        .biPlanes = 1
        .biBitCount = 16            ' 16-bit RGB
        .biCompression = 0          ' BI_RGB
        .biSizeImage = 320& * 288& * 4&
    End With
    Scale2x = True
End Sub

Public Sub ApplyScale2x()
    Dim x As Long, y As Long
    Dim p As Long, a As Long, b As Long, c As Long, d As Long
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long
    Dim dx As Long, dy As Long

    For y = 0 To 143
        dy = y * 2
        For x = 0 To 159
            dx = x * 2
            p = Vram(x, y)

            ' °æ°è °Ë»ç: »ó(A), ¿ì(B), ÇÏ(C), ÁÂ(D)
            If y > 0 Then a = Vram(x, y - 1) Else a = p
            If x < 159 Then b = Vram(x + 1, y) Else b = p
            If y < 143 Then c = Vram(x, y + 1) Else c = p
            If x > 0 Then d = Vram(x - 1, y) Else d = p

            ' EPX/Scale2x ±ÔÄ¢ ÆÇÁ¤
            If (d = a) And (c <> a) And (d <> b) Then p1 = a Else p1 = p
            If (a = b) And (a <> c) And (b <> d) Then p2 = b Else p2 = p
            If (d = c) And (d <> a) And (c <> b) Then p3 = c Else p3 = p
            If (b = c) And (b <> a) And (c <> d) Then p4 = c Else p4 = p

            ' 2x2 ¼­ºêÇÈ¼¿ ±â·Ï
            Vram2x(dx, dy) = p1
            Vram2x(dx + 1, dy) = p2
            Vram2x(dx, dy + 1) = p3
            Vram2x(dx + 1, dy + 1) = p4
        Next x
    Next y
End Sub

Public Sub DrawScreen() 'Using Api and SetBits/StrechBits
    FPS = FPS + 1
    
    If fmode = 0 Then 'frame skip mode 1(act skip(x1(1),x2(2),x3(3),x4(4),x5(5),x6(6))
    If FPS Mod fskip > 0 Then
    Skipf = True
    If lfp Then
    Do
    QueryPerformanceCounter CurEnd
    dblResult = (CurEnd - curStart) / curFreq
    Loop While dblResult < 16.6
    End If
    Exit Sub
    End If
    Else 'frame skip mode 2(act skip(x1.20(6),x1.25(5),x1,3(4),x1.5(3))
    If FPS Mod fskip = 0 Then
    Skipf = True
    If lfp Then
    Do
    QueryPerformanceCounter CurEnd
    dblResult = (CurEnd - curStart) / curFreq
    Loop While dblResult < 16.6
    End If
    Exit Sub
    End If
    End If
    
  If Scale2x Then
  ' 1. ¿øº» VramÀ» Scale2x ÇÊÅÍ·Î ¿¬»êÇÏ¿© Vram2x ¹öÆÛ »ý¼º
    Call ApplyScale2x
    
    ' 2. 320x288 Å©±âÀÇ Vram2x ¹öÆÛ¸¦ ´ë»ó Ã¢ Å©±â(dw, dh)·Î Ãâ·Â
    Call StretchDIBits(desthdc, 0, 0, dw, dh, 0, 0, 320, 288, Vram2x(0, 0), bb2x, 0, vbSrcCopy)
  Else
    StretchDIBits desthdc, 0, 0, dw, dh, 0, 0, 160, 144, Vram(0, 0), bb, 0, vbSrcCopy
  End If
  
    Form1.Picture1.Refresh
 
     Skipf = False
    If lfp Then
      Do
        QueryPerformanceCounter CurEnd
        dblResult = (CurEnd - curStart) / curFreq
      Loop While dblResult < 16.6
    End If
End Sub
   

Sub ccolid2(col As Byte, target As Long)
    Dim tm1 As Long, tm2 As Long
    Dim idx1 As Long, idx2 As Long

    ' 1. (col And 2, col And 1) -> (0~1, 0~1) ì •ê·œí™”
    idx1 = (col And 2) \ 2
    idx2 = col And 1
    colid2(target, 0, 0) = colid(idx1, idx2)

    ' 2. (col And 8, col And 4) -> (0~1, 0~1) ì •ê·œí™”
    idx1 = (col And 8) \ 8
    idx2 = (col And 4) \ 4
    
    tm2 = 1
    Do While tm2 <= 128
        colid2(target, 0, tm2) = colid(idx1, idx2)
        tm2 = tm2 * 2
    Loop

    ' 3. (col And 32, col And 16) -> (0~1, 0~1) ì •ê·œí™”
    idx1 = (col And 32) \ 32
    idx2 = (col And 16) \ 16
    
    tm1 = 1
    Do While tm1 <= 128
        colid2(target, tm1, 0) = colid(idx1, idx2)
        tm1 = tm1 * 2
    Loop

    ' 4. (col And 128, col And 64) -> (0~1, 0~1) ì •ê·œí™”
    idx1 = (col And 128) \ 128
    idx2 = (col And 64) \ 64
    
    tm1 = 1
    Do While tm1 <= 128
        tm2 = 1
        Do While tm2 <= 128
            colid2(target, tm1, tm2) = colid(idx1, idx2)
            tm2 = tm2 * 2
        Loop
        tm1 = tm1 * 2
    Loop
End Sub


Sub initGxMode2(dest As PictureBox, Siz As Long)
mode1 = False
Form1.Picture1.AutoRedraw = True
Form1.Picture1.BorderStyle = 0
Form1.Picture1.ClipControls = False
Form1.Picture1.ScaleMode = 3
Form1.Picture1.BackColor = RGB(256, 256, 256)
Form1.Picture1.Width = 15 * 160 * Siz
Form1.Picture1.Height = 15 * 144 * Siz
Form1.Picture1.Visible = True
With bb.Header
    .biSize = 40
    .biWidth = 160
    .biHeight = -144
    .biPlanes = 1
    .biBitCount = 16
    .biSizeImage = 46080
End With
destW = dest.ScaleWidth
destH = dest.ScaleHeight
desthdc = dest.hdc
desthimg = dest.Image.Handle
initCol
dh = 144 * Siz
dw = 160 * Siz
End Sub

Sub initCol()
For i = 0 To 32767
gbcP(i) = (i And 31744) \ 32 \ 32 + (i And 992) + (i And 31) * 32 * 32
'bgr rgb
Next i
    initMir
    Dim tm2 As Long, tm1 As Long
    If TGBC Then
    colid(0, 0) = rgb15(31, 31, 31)
    For tm2 = 1 To 128
    colid(0, tm2) = rgb15(21, 21, 21)
    Next tm2
    
    For tm1 = 1 To 128
    colid(tm1, 0) = rgb15(10, 10, 10)
    Next tm1
    
    For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid(tm1, tm2) = rgb15(0, 0, 0)
    Next tm2
    Next tm1
    Else
    colid(0, 0) = rgb15(31, 31, 31)
    For tm2 = 1 To 128
    colid(0, tm2) = rgb15(21, 21, 21)
    Next tm2
    
    For tm1 = 1 To 128
    colid(tm1, 0) = rgb15(10, 10, 10)
    Next tm1
    
    For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid(tm1, tm2) = rgb15(0, 0, 0)
    Next tm2
    Next tm1
    End If
    colidx(0, 0) = 0
    
    For tm2 = 1 To 128
    colidx(0, tm2) = 2
    Next tm2
    
    For tm1 = 1 To 128
    colidx(tm1, 0) = 1
    Next tm1
    
    For tm1 = 1 To 128
    For tm2 = 1 To 128
    colidx(tm1, tm2) = 3
    Next tm2
    Next tm1
    
    colid2(0, 0, 0) = colid(0, 0)
    For tm2 = 1 To 128
    colid2(0, 0, tm2) = colid(0, 1)
    Next tm2
    For tm1 = 1 To 128
    colid2(0, tm1, 0) = colid(1, 0)
    Next tm1
    For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid2(0, tm1, tm2) = colid(1, 1)
    Next tm2
    Next tm1
    colid2(1, 0, 0) = colid(0, 0)
    For tm2 = 1 To 128
    colid2(1, 0, tm2) = colid(0, 1)
    Next tm2: For tm1 = 1 To 128
    colid2(1, tm1, 0) = colid(1, 0)
    Next tm1: For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid2(1, tm1, tm2) = colid(1, 1)
    Next tm2
    Next tm1
    colid2(2, 0, 0) = colid(0, 0)
    For tm2 = 1 To 128
    colid2(2, 0, tm2) = colid(0, 1)
    Next tm2
    For tm1 = 1 To 128
    colid2(2, tm1, 0) = colid(1, 0)
    Next tm1: For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid2(2, tm1, tm2) = colid(1, 1)
    Next tm2: Next tm1
End Sub

Public Function initMir()
For i = 0 To 255
mir(i) = (i And 128) \ 128 + (i And 64) \ 32 + (i And 32) \ 8 + (i And 16) \ 2 + _
         (i And 8) * 2 + (i And 4) * 8 + (i And 2) * 32 + (i And 1) * 128
Next i
End Function

Function rgb15(Red As Byte, Green As Byte, Blue As Byte) As Integer
rgb15 = Red * 1024 + Green * 32 + Blue
End Function
Sub SrceenShot()
Dim flname As String, ni As Long
ret:
flname = App.Path & "\" & rominfo.title & " - " & ni & ".bmp"
If Not fileexist(flname) Then
SavePicture Form1.Picture1.Image, flname
Else
ni = ni + 1
GoTo ret
End If
End Sub

Function fileexist(file As String) As Boolean
On Error Resume Next
If FileLen(file) = 0 Then fileexist = False Else fileexist = True
End Function

Sub Drawline4() 'Using Api and SetBits/StrechBits
curline = RAM(65348, 0)
If curline = lastline Then Exit Sub
lastline = curline
    ' Draw Background
    ' Get BG & window Tile Pattern Data Address
    If RAM(65344, 0) And 16 Then
        TileData = 32768
    Else
        TileData = 34816
    End If
    If bgv Then
    ' Get BG Tile Table Address
    If RAM(65344, 0) And 8 Then
        tilemap = 39936
        tms = 39936
    Else
        tilemap = 38912
        tms = 38912
    End If
    tileend = tilemap + 1023
    xoffset = RAM(65347, 0)
    yoffset = RAM(65346, 0) + curline
    xs = xoffset \ 8
    ys = yoffset \ 8
    yoffset = yoffset And 7
    xoffset = -(xoffset And 7)
    For x = xoffset To 159 Step 8
        tiletmp = tilemap + ys * 32 + xs
        If tiletmp > tileend Then tiletmp = tiletmp - 1024
        If TileData = 32768 Then           ' Tile Data @ &H8800-&h97FF is 128ed
            tileptr = RAM(tiletmp, 0) * 16             'Get pointer to tile
        Else
            tileptr = (RAM(tiletmp, 0) Xor 128) * 16
        End If
        
        xs = (xs + 1) Mod 32
        
        memptr = TileData + tileptr + (yoffset And 7) * 2
        mv1 = RAM(memptr + 1, 0): mv2 = RAM(memptr, 0)
        If x > -1 And x < 153 Then
            Vram(x + 7, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 6, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 5, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 4, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 3, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 2, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 1, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x, curline) = colid2(0, mv1 And 1, mv2 And 1)
        Else
            If x < 153 Then Vram(x + 7, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -7 And x < 154 Then Vram(x + 6, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -6 And x < 155 Then Vram(x + 5, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -5 And x < 156 Then Vram(x + 4, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -4 And x < 157 Then Vram(x + 3, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -3 And x < 158 Then Vram(x + 2, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -2 And x < 159 Then Vram(x + 1, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -1 Then Vram(x, curline) = colid2(0, mv1 And 1, mv2 And 1)
        End If
    Next x
    End If
    
    
        'Draw Window
    
If (RAM(65344, 0) And 32) = 32 And wv And curline >= RAM(65354, 0) And RAM(65355, 0) < 167 Then
    ' Get window Tile Table Address
    If RAM(65344, 0) And 64 Then
        tilemap = 39936
    Else
        tilemap = 38912
    End If
    yoffset = curline - RAM(65354, 0)
    tilemap = tilemap + (yoffset \ 8) * 32
    yoffset = yoffset And 7
    For x = RAM(65355, 0) - 7 To 159 Step 8
        If TileData = 32768 Then           ' Tile Data @ &H8800-&h97FF is 128ed
            tileptr = RAM(tilemap, 0) * 16             'Get pointer to tile
        Else
            tileptr = (RAM(tilemap, 0) Xor 128) * 16
        End If
        memptr = TileData + tileptr + (yoffset And 7) * 2
        mv1 = RAM(memptr + 1, 0): mv2 = RAM(memptr, 0)
        
        If x > -1 And x < 153 Then
            Vram(x + 7, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 6, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 5, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 4, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 3, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 2, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x + 1, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(x, curline) = colid2(0, mv1 And 1, mv2 And 1)
            tilemap = tilemap + 1
        Else
            If x < 153 Then Vram(x + 7, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -7 And x < 154 Then Vram(x + 6, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -6 And x < 155 Then Vram(x + 5, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -5 And x < 156 Then Vram(x + 4, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -4 And x < 157 Then Vram(x + 3, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -3 And x < 158 Then Vram(x + 2, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -2 And x < 159 Then Vram(x + 1, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -1 Then Vram(x, curline) = colid2(0, mv1 And 1, mv2 And 1)
            tilemap = tilemap + 1
        End If
        
    Next x
End If

        'Draw Sprites
If (RAM(65344, 0) And 2) And objv Then
    If (RAM(65344, 0) And 4) = 0 Then
        SpriteY = 7
    Else
        SpriteY = 15
    End If
    For tilemap = 65180 To 65024 Step -4
        y = RAM(tilemap, 0) - 16
        x = RAM(tilemap + 1, 0) - 8
        If y <= curline And y + SpriteY >= curline And x > -8 And x < 160 Then
        
        If SpriteY = 7 Then tileptr = readM(tilemap + 2) * 16 Else tileptr = (RAM(tilemap + 2, 0) And 254) * 16 'Get pointer to tile
        TCol = -((RAM(tilemap + 3, 0) And 128) > 0)
        spat = -((readM(tilemap + 3) And 16) > 0) + 1
        vf = RAM(tilemap + 3, 0) And 32
        'init palete
        memptr = 32768 + tileptr
        
        If (RAM(tilemap + 3, 0) And 64) Then memptr = memptr + SpriteY * 2 - (curline - y) * 2 Else memptr = memptr + (curline - y) * 2
        If vf Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        
        If TCol = 0 Then
            
            If x < 153 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 7, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -7 And x < 154 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 6, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -6 And x < 155 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 5, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -5 And x < 156 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 4, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -4 And x < 157 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 3, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -3 And x < 158 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 2, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -2 And x < 159 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 1, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -1 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            
        Else
            cid2 = colid2(0, 0, 0)
            If x < 153 Then If Vram(x + 7, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 7, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -7 And x < 154 Then If Vram(x + 6, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 6, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -6 And x < 155 Then If Vram(x + 5, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 5, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -5 And x < 156 Then If Vram(x + 4, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 4, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -4 And x < 157 Then If Vram(x + 3, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 3, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -3 And x < 158 Then If Vram(x + 2, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 2, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -2 And x < 159 Then If Vram(x + 1, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 1, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -1 Then If Vram(x, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            
        End If
    End If
    Next tilemap
   End If
   
   
    
End Sub

Sub Drawline() 'Using Api and SetBits/StrechBits
   
curline = RAM(65348, 0)
If curline = lastline Then Exit Sub
lastline = curline
    ' Draw Background
    ' Get BG & window Tile Pattern Data Address
    If RAM(65344, 0) And 16 Then
        TileData = 32768
    Else
        TileData = 34816
    End If

If bgv Then
    ' Get BG Tile Table Address
    If RAM(65344, 0) And 8 Then
        tilemap = 39936
        tms = 39936
    Else
        tilemap = 38912
        tms = 38912
    End If
    
    tileend = tilemap + 1023
    xoffset = RAM(65347, 0)
    yoffset = RAM(65346, 0) + curline
    xs = xoffset \ 8
    ys = yoffset \ 8
    yoffset = yoffset And 7
    xoffset = -(xoffset And 7)
    
    For x = xoffset To 159 Step 8
        tiletmp = tilemap + ys * 32 + xs
        If tiletmp > tileend Then tiletmp = tiletmp - 1024
        If TileData = 32768 Then           ' Tile Data @ &H8800-&h97FF is 128ed
            tileptr = RAM(tiletmp, 0) * 16             'Get pointer to tile
        Else
            tileptr = (RAM(tiletmp, 0) Xor 128) * 16
        End If
        bgat = RAM(tiletmp, 1)
        ccp = bgat And 7
        tcls(x \ 8) = gbcP(bgp(ccp, 0))
        vrm = (bgat And 8) \ 8
        xflip = (bgat And 32) \ 32: yflip = (bgat And 64) \ 64
        ccid(0, 0) = gbcP(bgp(ccp, 0))
        ccid(0, 1) = gbcP(bgp(ccp, 1))
        ccid(1, 0) = gbcP(bgp(ccp, 2))
        ccid(1, 1) = gbcP(bgp(ccp, 3))
        If yflip Then memptr = TileData + tileptr + 14 - (yoffset And 7) * 2 Else memptr = TileData + tileptr + (yoffset And 7) * 2
        If xflip Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        xs = (xs + 1) Mod 32
        
      
        '''''''''''''
        ' È­¸é ¾ÈÂÊ(0 <= x <= 152)Àº °æ°è °Ë»ç(If)°¡ ÀüÇô ÇÊ¿ä ¾øÀ¸¹Ç·Î ¹«Á¶°Ç ´ÙÀÌ·ºÆ® ±â·Ï
        If x >= 0 And x <= 152 Then
            Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
            Vram(x + 6, curline) = ccid((mv1 \ 2) And 1, (mv2 \ 2) And 1)
            Vram(x + 5, curline) = ccid((mv1 \ 4) And 1, (mv2 \ 4) And 1)
            Vram(x + 4, curline) = ccid((mv1 \ 8) And 1, (mv2 \ 8) And 1)
            Vram(x + 3, curline) = ccid((mv1 \ 16) And 1, (mv2 \ 16) And 1)
            Vram(x + 2, curline) = ccid((mv1 \ 32) And 1, (mv2 \ 32) And 1)
            Vram(x + 1, curline) = ccid((mv1 \ 64) And 1, (mv2 \ 64) And 1)
            Vram(x, curline) = ccid((mv1 \ 128) And 1, (mv2 \ 128) And 1)
            
           
        Else
            ' È­¸é ÁÂ/¿ì °æ°è(-7 <= x < 0 ¶Ç´Â x > 152)¿¡ °ÉÄ£ Å¸ÀÏ¸¸ ¾ÈÀüÇÏ°Ô 1ÇÈ¼¿¾¿ Å¬¸®ÇÎ
            Dim px As Long, scrX As Long
            For px = 7 To 0 Step -1
                scrX = x + px
                If scrX >= 0 And scrX < 160 Then
                    Vram(scrX, curline) = ccid(mv1 And 1, mv2 And 1)
                End If
                mv1 = mv1 \ 2
                mv2 = mv2 \ 2
            Next px
        End If
        
        ''''''''''''
      
            
    Next x
    End If
    
    '''''''''''''''''''
    ' ----------------------------------------------------
    ' Draw Window
    ' ----------------------------------------------------
    Dim winY As Long, winX As Long
    Dim wx As Long, wy As Long
    Dim winTileMapBase As Long
    Dim winTileAddr As Long
    
    winY = RAM(65354, 0)        ' WY (Window Y)
    winX = RAM(65355, 0) - 7    ' WX (Window X - 7)
    
    If (RAM(65344, 0) And 32) = 32 And wv And curline >= winY And winX < 160 Then
        ' Window Tile Map Base Address (LCDC bit 6)
        If RAM(65344, 0) And 64 Then
            winTileMapBase = 39936  ' &H9C00
        Else
            winTileMapBase = 38912  ' &H9800
        End If
        
        yoffset = curline - winY
        wy = yoffset \ 8            ' À©µµ¿ì Å¸ÀÏ¸Ê Çà (0 ~ 31)
        yoffset = yoffset And 7     ' Å¸ÀÏ ³» ¶óÀÎ (0 ~ 7)
        
        wx = 0
        For x = winX To 159 Step 8
            ' ÇöÀç Å¸ÀÏÀÇ ¸Ê ÁÖ¼Ò (X ÀÎµ¦½º´Â 0ºÎÅÍ 1Ä­¾¿ Áõ°¡)
            winTileAddr = winTileMapBase + (wy * 32) + (wx Mod 32)
            wx = wx + 1
            
            ' Å¸ÀÏ ¹øÈ£ °¡Á®¿À±â ¹× ÆÐÅÏ µ¥ÀÌÅÍ ÁÖ¼Ò °è»ê
            If TileData = 32768 Then
                ' 0x8000 ¹æ½Ä (Unsigned)
                tileptr = RAM(winTileAddr, 0) * 16
                memptr = 32768 + tileptr
            Else
                ' 0x8800 ¹æ½Ä (Signed: -128 ~ 127, ±âÁØÁ¡ 0x9000 = 36864)
                ' (VBA ¹ÙÀÌÆ®¸¦ Signed Á¤¼ö·Î Ã³¸®)
                Dim rawTile As Long
                rawTile = RAM(winTileAddr, 0)
                If rawTile > 127 Then rawTile = rawTile - 256
                memptr = 36864 + (rawTile * 16)
            End If
            
            bgat = RAM(winTileAddr, 1)
            ccp = bgat And 7
            vrm = (bgat And 8) \ 8
            xflip = (bgat And 32) \ 32
            yflip = (bgat And 64) \ 64
            
            ' ÆÈ·¹Æ® ¼¼ÆÃ
            ccid(0, 0) = gbcP(bgp(ccp, 0))
            ccid(0, 1) = gbcP(bgp(ccp, 1))
            ccid(1, 0) = gbcP(bgp(ccp, 2))
            ccid(1, 1) = gbcP(bgp(ccp, 3))
            
            ' Y-Flip¿¡ µû¸¥ ½ºÄµ¶óÀÎ ¿ÀÇÁ¼Â °è»ê
            If yflip Then
                memptr = memptr + (14 - yoffset * 2)
            Else
                memptr = memptr + (yoffset * 2)
            End If
            
            ' VRAM Bank(vrm)¿¡¼­ 2¹ÙÀÌÆ® ÀÐ±â (X-Flip Àû¿ë)
            If xflip Then
                mv1 = mir(RAM(memptr + 1, vrm))
                mv2 = mir(RAM(memptr, vrm))
            Else
                mv1 = RAM(memptr + 1, vrm)
                mv2 = RAM(memptr, vrm)
            End If
            
            ' ÇÈ¼¿ Âï±â
         
            
            '''''''''''''''''
            If x >= 0 And x <= 152 Then
                Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
                Vram(x + 6, curline) = ccid((mv1 \ 2) And 1, (mv2 \ 2) And 1)
                Vram(x + 5, curline) = ccid((mv1 \ 4) And 1, (mv2 \ 4) And 1)
                Vram(x + 4, curline) = ccid((mv1 \ 8) And 1, (mv2 \ 8) And 1)
                Vram(x + 3, curline) = ccid((mv1 \ 16) And 1, (mv2 \ 16) And 1)
                Vram(x + 2, curline) = ccid((mv1 \ 32) And 1, (mv2 \ 32) And 1)
                Vram(x + 1, curline) = ccid((mv1 \ 64) And 1, (mv2 \ 64) And 1)
                Vram(x, curline) = ccid((mv1 \ 128) And 1, (mv2 \ 128) And 1)
                
               
            
            Else
                For px = 7 To 0 Step -1
                    scrX = x + px
                    If scrX >= 0 And scrX < 160 Then
                        Vram(scrX, curline) = ccid(mv1 And 1, mv2 And 1)
                    End If
                    mv1 = mv1 \ 2
                    mv2 = mv2 \ 2
                Next px
            End If
            
            ''''''''''''''''''''''
            
            
                
        Next x
    End If
    
    
    
    
    ''''''''''''''''''''''
        
    
    'Draw Sprites
If (RAM(65344, 0) And 2) And objv Then
    If (RAM(65344, 0) And 4) = 0 Then
        SpriteY = 7
    Else
        SpriteY = 15
    End If
    For tilemap = 65180 To 65024 Step -4
        y = RAM(tilemap, 0) - 16
        x = RAM(tilemap + 1, 0) - 8
        If y <= curline And y + SpriteY >= curline And x > -8 And x < 160 Then
        
        If SpriteY = 7 Then tileptr = readM(tilemap + 2) * 16 Else tileptr = (RAM(tilemap + 2, 0) And 254) * 16 'Get pointer to tile
        TCol = -((RAM(tilemap + 3, 0) And 128) > 0)
        vf = RAM(tilemap + 3, 0) And 32
        vrm = (RAM(tilemap + 3, 0) And 8) \ 8
        ccp = RAM(tilemap + 3, 0) And 7
        
        'init palete
        ccid(0, 0) = gbcP(objp(ccp, 0))
        ccid(0, 1) = gbcP(objp(ccp, 1))
        ccid(1, 0) = gbcP(objp(ccp, 2))
        ccid(1, 1) = gbcP(objp(ccp, 3))
        memptr = 32768 + tileptr
        
        If (RAM(tilemap + 3, 0) And 64) Then memptr = memptr + SpriteY * 2 - (curline - y) * 2 Else memptr = memptr + (curline - y) * 2
        If vf Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        
      
        
        If TCol = 0 Then
            If x < 153 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -7 And x < 154 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 6, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -6 And x < 155 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 5, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -5 And x < 156 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 4, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -4 And x < 157 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 3, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -3 And x < 158 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 2, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -2 And x < 159 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 1, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -1 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x, curline) = ccid(mv1 And 1, mv2 And 1)
            
        Else
            If x < 153 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 7, curline) = tcls((x + 7) \ 8) Then Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -7 And x < 154 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 6, curline) = tcls((x + 6) \ 8) Then Vram(x + 6, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -6 And x < 155 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 5, curline) = tcls((x + 5) \ 8) Then Vram(x + 5, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -5 And x < 156 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 4, curline) = tcls((x + 4) \ 8) Then Vram(x + 4, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -4 And x < 157 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 3, curline) = tcls((x + 3) \ 8) Then Vram(x + 3, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -3 And x < 158 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 2, curline) = tcls((x + 2) \ 8) Then Vram(x + 2, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -2 And x < 159 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 1, curline) = tcls((x + 1) \ 8) Then Vram(x + 1, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -1 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x, curline) = tcls(x \ 8) Then Vram(x, curline) = ccid(mv1 And 1, mv2 And 1)
            
        End If
        

    End If
               
                
    Next tilemap
   End If
   
   
End Sub



Sub Drawline_0() 'Using Api and SetBits/StrechBits
curline = RAM(65348, 0)
If curline = lastline Then Exit Sub
lastline = curline
    ' Draw Background
    ' Get BG & window Tile Pattern Data Address
    If RAM(65344, 0) And 16 Then
        TileData = 32768
    Else
        TileData = 34816
    End If
If bgv Then
    ' Get BG Tile Table Address
    If RAM(65344, 0) And 8 Then
        tilemap = 39936
        tms = 39936
    Else
        tilemap = 38912
        tms = 38912
    End If
    tileend = tilemap + 1023
    xoffset = RAM(65347, 0)
    yoffset = RAM(65346, 0) + curline
    xs = xoffset \ 8
    ys = yoffset \ 8
    yoffset = yoffset And 7
    xoffset = -(xoffset And 7)
    
    For x = xoffset To 159 Step 8
        tiletmp = tilemap + ys * 32 + xs
        If tiletmp > tileend Then tiletmp = tiletmp - 1024
        If TileData = 32768 Then           ' Tile Data @ &H8800-&h97FF is 128ed
            tileptr = RAM(tiletmp, 0) * 16             'Get pointer to tile
        Else
            tileptr = (RAM(tiletmp, 0) Xor 128) * 16
        End If
        bgat = RAM(tiletmp, 1)
        ccp = bgat And 7
        tcls(x \ 8) = gbcP(bgp(ccp, 0))
        vrm = (bgat And 8) \ 8
        xflip = (bgat And 32) \ 32: yflip = (bgat And 64) \ 64
        ccid(0, 0) = gbcP(bgp(ccp, 0))
        ccid(0, 1) = gbcP(bgp(ccp, 1))
        ccid(1, 0) = gbcP(bgp(ccp, 2))
        ccid(1, 1) = gbcP(bgp(ccp, 3))
        If yflip Then memptr = TileData + tileptr + 14 - (yoffset And 7) * 2 Else memptr = TileData + tileptr + (yoffset And 7) * 2
        If xflip Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        xs = (xs + 1) Mod 32
        
        If x > -1 And x < 153 Then
        Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 6, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 5, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 4, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 3, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 2, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 1, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x, curline) = ccid(mv1 And 1, mv2 And 1)
        Else
        If x < 153 Then Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -7 And x < 154 Then Vram(x + 6, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -6 And x < 155 Then Vram(x + 5, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -5 And x < 156 Then Vram(x + 4, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -4 And x < 157 Then Vram(x + 3, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -3 And x < 158 Then Vram(x + 2, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -2 And x < 159 Then Vram(x + 1, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -1 Then Vram(x, curline) = ccid(mv1 And 1, mv2 And 1)
        End If
        
    Next x
    End If
    
    
        'Draw Window
    
If (RAM(65344, 0) And 32) = 32 And wv And curline >= RAM(65354, 0) And RAM(65355, 0) < 167 Then
    ' Get window Tile Table Address
    If RAM(65344, 0) And 64 Then
        tilemap = 39936
    Else
        tilemap = 38912
    End If
    yoffset = curline - RAM(65354, 0)
    tilemap = tilemap + (yoffset \ 8) * 32
    yoffset = yoffset And 7
    For x = RAM(65355, 0) - 7 To 159 Step 8
        If TileData = 32768 Then           ' Tile Data @ &H8800-&h97FF is 128ed
            tileptr = RAM(tilemap, 0) * 16             'Get pointer to tile
        Else
            tileptr = (RAM(tilemap, 0) Xor 128) * 16
        End If
        bgat = RAM(tilemap, 1)
        ccp = bgat And 7
        vrm = (bgat And 8) \ 8
        xflip = (bgat And 32) \ 32: yflip = (bgat And 64) \ 64
        ccid(0, 0) = gbcP(bgp(ccp, 0))
        ccid(0, 1) = gbcP(bgp(ccp, 1))
        ccid(1, 0) = gbcP(bgp(ccp, 2))
        ccid(1, 1) = gbcP(bgp(ccp, 3))
        If yflip Then memptr = TileData + tileptr + 14 - yoffset * 2 Else memptr = TileData + tileptr + yoffset * 2
        If xflip Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        
        If x > -1 And x < 153 Then
        Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 6, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 5, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 4, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 3, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 2, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x + 1, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(x, curline) = ccid(mv1 And 1, mv2 And 1)
        Else
        If x < 154 Then Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -7 And x < 154 Then Vram(x + 6, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -6 And x < 155 Then Vram(x + 5, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -5 And x < 156 Then Vram(x + 4, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -4 And x < 157 Then Vram(x + 3, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -3 And x < 158 Then Vram(x + 2, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -2 And x < 159 Then Vram(x + 1, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If x > -1 Then Vram(x, curline) = ccid(mv1 And 1, mv2 And 1)
        End If
        
        
        
        tilemap = tilemap + 1
    Next x
End If
    
    'Draw Sprites
If (RAM(65344, 0) And 2) And objv Then
    If (RAM(65344, 0) And 4) = 0 Then
        SpriteY = 7
    Else
        SpriteY = 15
    End If
    For tilemap = 65180 To 65024 Step -4
        y = RAM(tilemap, 0) - 16
        x = RAM(tilemap + 1, 0) - 8
        If y <= curline And y + SpriteY >= curline And x > -8 And x < 160 Then
        
        If SpriteY = 7 Then tileptr = readM(tilemap + 2) * 16 Else tileptr = (RAM(tilemap + 2, 0) And 254) * 16 'Get pointer to tile
        TCol = -((RAM(tilemap + 3, 0) And 128) > 0)
        vf = RAM(tilemap + 3, 0) And 32
        vrm = (RAM(tilemap + 3, 0) And 8) \ 8
        ccp = RAM(tilemap + 3, 0) And 7
        'init palete
        ccid(0, 0) = gbcP(objp(ccp, 0))
        ccid(0, 1) = gbcP(objp(ccp, 1))
        ccid(1, 0) = gbcP(objp(ccp, 2))
        ccid(1, 1) = gbcP(objp(ccp, 3))
        memptr = 32768 + tileptr
        
        If (RAM(tilemap + 3, 0) And 64) Then memptr = memptr + SpriteY * 2 - (curline - y) * 2 Else memptr = memptr + (curline - y) * 2
        If vf Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        
        If TCol = 0 Then
            If x < 153 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -7 And x < 154 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 6, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -6 And x < 155 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 5, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -5 And x < 156 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 4, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -4 And x < 157 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 3, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -3 And x < 158 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 2, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -2 And x < 159 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x + 1, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -1 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(x, curline) = ccid(mv1 And 1, mv2 And 1)
            
        Else
            If x < 153 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 7, curline) = tcls((x + 7) \ 8) Then Vram(x + 7, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -7 And x < 154 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 6, curline) = tcls((x + 6) \ 8) Then Vram(x + 6, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -6 And x < 155 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 5, curline) = tcls((x + 5) \ 8) Then Vram(x + 5, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -5 And x < 156 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 4, curline) = tcls((x + 4) \ 8) Then Vram(x + 4, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -4 And x < 157 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 3, curline) = tcls((x + 3) \ 8) Then Vram(x + 3, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -3 And x < 158 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 2, curline) = tcls((x + 2) \ 8) Then Vram(x + 2, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -2 And x < 159 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x + 1, curline) = tcls((x + 1) \ 8) Then Vram(x + 1, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If x > -1 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(x, curline) = tcls(x \ 8) Then Vram(x, curline) = ccid(mv1 And 1, mv2 And 1)
            
        End If
    End If
    Next tilemap
   End If
   
End Sub
