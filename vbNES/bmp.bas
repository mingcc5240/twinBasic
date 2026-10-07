Attribute VB_Name = "Bitmap"
' BMP File format saving for basicNES 2000.
' By Don Jarrett w/ help from Azimer, 2002.

Public Sub WriteBitmap(filename As String)
Const Header As String = "BM"
Const HeaderSize As Double = &H28
Const Compression As Double = &H0
Dim FileSize As Double, Reserved As Double
Dim Offset As Double, Height As Double, Width As Double
Dim Planes As Long, BPP As Long
Height = 256: Width = 240
Planes = &H1: BPP = &H8
    Close #1
    Open filename For Binary As #1
        
        Put #1, , Header
        Put #1, , FileSize
        Put #1, , Reserved
        Put #1, , Offset
        Put #1, , HeaderSize
        Put #1, , Height
        Put #1, , Width
        Put #1, , Planes
        Put #1, , BPP
        Put #1, , Compression
    Close #1
        
End Sub
