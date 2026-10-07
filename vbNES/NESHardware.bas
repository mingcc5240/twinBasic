Attribute VB_Name = "NESHardware"
'+====================================================+'
'| basicNES 2k                                        |'
'| by Don Jarrett, 2000                               |'
'| Big portions of graphics code done by David Finch. |'
'| Always a source/binary release                     |'
'| http://home.att.net/~r.jarrett/bNES.html           |'
'+====================================================+'

Option Explicit ' Option Explicit is important for avoiding mysterious, crippling bugs

DefLng A-Z 'this ensures that no variants are used unless we choose so

'Public EmphVal As Integer

Public romName As String

Public tmpLatch As Byte, ppuLatch As Byte

Public map24_irqv As Byte

Public Interlace As Long

Public MirrorTypes(4) As String

'DF: array to draw each frame to
Public vBuffer(256& * 241& - 1) As Byte '256*241 to allow for some overflow

Public nt(0 To 3, &H0 To &H3FF) As Byte
Public mirror(3) As Byte

Public vBuffer16(256& * 240& - 1) As Integer
Public vBuffer32(256& * 240& - 1) As Long

Public oldvBuffer16(256& * 240& - 1) As Integer 'used for scaling modes that take advantage of unchanged pixels

Public vBuffer2x16(512& * 480& - 1) As Integer
Public vBuffer2x32(512& * 480& - 1) As Long

Public tLook(65536 * 8 - 1) As Byte

Public newScroll As Boolean

Public SpritesChanged As Boolean

Public Record As Boolean
Public MovieIndex As Integer
Public Playing As Boolean

Public map17_irqon As Boolean
Public map17_irq As Long

Public IRQCounter As Long ' for mapper 6
Public map6_irqon As Byte
Public map225_psize As Byte
Public map225_psel As Byte
Public Train(&H1FF) As Byte
Public latch13 As Byte
Public MMC19_IRQCount As Long
Public MIRQOn As Byte
Public map24_irqon As Byte


'DF: powers of 2
Public pow2(31) As Long

Public tilebased As Boolean

' NES Hardware defines
Public PPU_Control1 As Byte ' $2000
Public PPU_Control2 As Byte ' $2001
Public PPU_Status As Byte ' $2002
Public SpriteAddress As Long ' $2003
Public PPUAddressHi As Long ' $2006, 1st write
Public PPUAddress As Long ' $2006
Public PPU_AddressIsHi As Boolean
Public VRAM(&H3FFF) As Byte, VROM() As Byte  ' Video RAM
Public SpriteRAM(&HFF) As Byte

Public Sound(0 To &H15) As Byte
Public SoundCtrl As Byte

Public VScroll2 As Long

Public PrgCount As Byte, PrgCount2 As Long, ChrCount As Byte, ChrCount2 As Long

Public cmd As Byte, prg As Byte, chr1 As Byte

Public reg8 As Byte
Public regA As Byte
Public regC As Byte
Public regE As Byte

Public NESPal(&HF) As Byte

Public CPal() As Long

Public FrameSkip As Long 'Integer
Public Frames As Long

Public ScrollToggle As Byte
Public HScroll As Byte, VScroll As Long 'Integer ' $2005

Public Map15_BankAddr As Byte
Public Map15_SwapReg As Byte

'DF: these variables were undefined and therefore local:
Public swap As Boolean
Public map15_swapaddr As Long

' MMC3[Mapper #4] infos
Public MMC3_Command As Byte
Public MMC3_PrgAddr As Byte
Public MMC3_ChrAddr As Integer
Public MMC3_IrqVal As Byte
Public MMC3_TmpVal As Byte
Public MMC3_IrqOn As Boolean

Public PatternTable As Long
Public NameTable As Long

Public bank_regs(16) As Byte

Public Const PPU_InVBlank = &H80
Public Const PPU_Sprite0 = &H40
Public Const PPU_SpriteCount = &H20
Public Const PPU_Ignored = &H10

Public reg8000 As Byte ' Needed for mapper #69.

Public AndIt As Byte, AndIt2 As Byte

Public data(4) As Byte, sequence As Long 'Integer
Public accumulator As Long 'Integer

Public render As Boolean

Public MotionBlur As Boolean
Public Smooth2x As Long

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public OpenFile As OPENFILENAME

Public ChrCnt As Byte, PrgCnt As Byte
Public FirstRead As Boolean ' First read to $2007 is invalid
Public Joypad1(7) As Byte, Joypad1_Count As Byte

Public Mapper As Byte, Mirroring As Byte, Trainer As Byte, FourScreen As Byte
Public MirrorXor As Long 'Integer
Public UsesSRAM As Boolean

Public Keyboard(255) As Byte

Public CPURunning As Boolean

Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal cb&)

Public Declare Sub MemFill Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)


Public MMC32_Switch As Byte

Public MMC16_Irq As Long 'Integer
Public MMC16_IrqOn As Byte

Public PPUAddress2 As Long
Public HScroll2 As Long

'outdated code
Public Type Sprite
    X As Integer
    Y As Integer
    tileno As Byte
    attrib As Byte
End Type
Public SpriteAddr As Long 'Integer

Public pal(255) As Long
Public pal16(255) As Integer
Public pal15(255) As Integer

'Public Sprites(63) As Sprite

Public noScroll2006 As Boolean

Public Mapper40_IRQEnabled As Byte, Mapper40_IRQCounter As Byte

Public latch1 As Byte, latch2 As Byte
Public Latch0FD As Byte, Latch0FE As Byte
Public Latch1FD As Byte, Latch1FE As Byte

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Other
Public Const VERSION = "basicNES 2000 v1.5 [debug level 2]"

'Fills color lookup table
Public Sub fillTLook()
    Dim b1, b2, c, X
    For b1 = 0 To 255
    For b2 = 0 To 255
        For X = 0 To 7
            If b1 And pow2(X) Then c = 1 Else c = 0
            If b2 And pow2(X) Then c = c + 2
            tLook(b1 * 2048 + b2 * 8 + X) = c
        Next X
    Next b2, b1
End Sub

Public Sub blitScreen()
    Static mbt As Boolean
    Static n As Long
    mbt = Not mbt
    n = 1 - n
    Static count64 As Long
    count64 = (count64 + 17) And 63
    
    Dim fscan As Boolean
    
    Dim t As Long
        
    'Lets you know what type of mirroring is being used
    Select Case Mirroring
    Case 0
        frmNES.lbMirror.Caption = "|"
    Case 1
        frmNES.lbMirror.Caption = "-"
    Case 2
        frmNES.lbMirror.Caption = "+"
    End Select
    
    Dim p01 As Integer, p10 As Integer, p21 As Integer, p12 As Integer
    
    Dim p01b As Long, p10b As Long, p21b As Long, p12b As Long
    
    'Involves converting to high or true color.
    'Could be sped up if you took time to figure out palette animation in windows.
    Dim X, Y, k, b, c
    
    Dim i As Long
    Static p16(31) As Integer
    Static p32(31) As Long
    Static colorDepth
    If colorDepth = 0 Or (Frames And 63) = 0 Then colorDepth = getColorDepth(frmNES.picScreen)
    Select Case colorDepth
    Case 8
        'first change the bitmap palette
            'not sure how
            
        'then blit
        'blit8 vBuffer, frmNES.picScreen, 256, 240
    Case 16
        If MotionBlur Then
            For i = 0 To 31
                p16(i) = (pal16(VRAM(i + &H3F00)) And &HF7DE&) \ 2
            Next i
            If mbt Then
                For i = 0 To 61439
                    vBuffer16(i) = (vBuffer16(i) And &HF7DE&) \ 2 + p16(vBuffer(i)) + &H20
                Next i
            Else
                For i = 0 To 61439
                    vBuffer16(i) = (vBuffer16(i) And &HF7DE&) \ 2 + p16(vBuffer(i)) + &H801
                Next i
            End If
            
        Else
            For i = 0 To 31
                p16(i) = pal16(VRAM(i + &H3F00))
            Next i
            For i = 0 To 61439 Step 4
                vBuffer16(i) = p16(vBuffer(i))
                vBuffer16(i + 1) = p16(vBuffer(i + 1))
                vBuffer16(i + 2) = p16(vBuffer(i + 2))
                vBuffer16(i + 3) = p16(vBuffer(i + 3))
            Next i
        End If
        
        If Smooth2x Then
            If Smooth2x = 1 Then
                For Y = 0 To 239
                i = Y * 256&
                k = Y * 1024&
                c = 0
                For X = 0 To 255
                    b = vBuffer16(i)
                    vBuffer2x16(k + 1) = b
                    b = (b And &HF7DE&) \ 2
                    vBuffer2x16(k) = b + c
                    c = b
                    i = i + 1
                    k = k + 2
                Next X, Y
                For Y = 0 To 238
                k = Y * 1024 + 512
                For X = 0 To 511
                    vBuffer2x16(k) = (vBuffer2x16(k - 512) And &HF7DE&) \ 2 + (vBuffer2x16(k + 512) And &HF7DE&) \ 2
                    k = k + 1
                Next X, Y
            Else
                For Y = 1 To 238
                i = Y * 256&
                k = Y * 1024&
                c = vBuffer16(i - 1)
                p21 = vBuffer16(i)
                fscan = (Y And 63) = count64
                For X = IIf(X > 1, 0, 1) To 255
                    b = vBuffer16(i + 1)
                    If b <> oldvBuffer16(i + 1) Then
                        If t <= 0 Then
                            c = vBuffer16(i - 1)
                            p21 = vBuffer16(i)
                        End If
                        t = 3
                    End If
                    If t > 0 Or fscan Then
                        t = t - 1
                        p01 = c
                        c = p21
                        p21 = b
                        b = (c And &HF7DE&) \ 2
                        p10 = vBuffer16(i - 256)
                        p12 = vBuffer16(i + 256)
                        If p01 = p10 And p12 = p21 Then
                            If p10 = p12 And vBuffer16(i - 257) = c Then
                                vBuffer2x16(k) = c
                                vBuffer2x16(k + 1) = (p10 And &HF7DE&) \ 2 + b
                                vBuffer2x16(k + 512) = (p12 And &HF7DE&) \ 2 + b
                                vBuffer2x16(k + 513) = c
                            Else
                                vBuffer2x16(k) = (p01 And &HF7DE&) \ 2 + b
                                vBuffer2x16(k + 1) = c
                                vBuffer2x16(k + 512) = c
                                vBuffer2x16(k + 513) = (p21 And &HF7DE&) \ 2 + b
                            End If
                        ElseIf p10 = p21 And p01 = p12 Then
                            vBuffer2x16(k) = c
                            vBuffer2x16(k + 1) = (p10 And &HF7DE&) \ 2 + b
                            vBuffer2x16(k + 512) = (p12 And &HF7DE&) \ 2 + b
                            vBuffer2x16(k + 513) = c
                        Else
                            vBuffer2x16(k) = c
                            vBuffer2x16(k + 1) = c
                            vBuffer2x16(k + 512) = c
                            vBuffer2x16(k + 513) = c
                        End If
                    End If
                    i = i + 1
                    k = k + 2
                Next X, Y
                MemCopy oldvBuffer16(0), vBuffer16(0), 122880
            End If
            
            
            blit16 vBuffer2x16, frmNES.picScreen, 512, 480
        Else
            blit16 vBuffer16, frmNES.picScreen, 256, 240
        End If
    Case 15
        If MotionBlur Then
            For i = 0 To 31
                p16(i) = (pal15(VRAM(i + &H3F00)) And &H7BDE&) \ 2
            Next i
            If mbt Then
                For i = 0 To 61439
                    vBuffer16(i) = (vBuffer16(i) And &H7BDE&) \ 2 + p16(vBuffer(i)) + &H20
                Next i
            Else
                For i = 0 To 61439
                    vBuffer16(i) = (vBuffer16(i) And &H7BDE&) \ 2 + p16(vBuffer(i)) + &H401
                Next i
            End If
        Else
            For i = 0 To 31
                p16(i) = pal15(VRAM(i + &H3F00))
            Next i
            For i = 0 To 61439 Step 4
                vBuffer16(i) = p16(vBuffer(i))
                vBuffer16(i + 1) = p16(vBuffer(i + 1))
                vBuffer16(i + 2) = p16(vBuffer(i + 2))
                vBuffer16(i + 3) = p16(vBuffer(i + 3))
            Next i
        End If
        
        If Smooth2x Then
            If Smooth2x = 1 Then
                For Y = 0 To 239
                i = Y * 256&
                k = Y * 1024&
                c = 0
                For X = 0 To 255
                    b = vBuffer16(i)
                    vBuffer2x16(k + 1) = b
                    b = (b And &H7BDE&) \ 2
                    vBuffer2x16(k) = b + c
                    c = b
                    i = i + 1
                    k = k + 2
                Next X, Y
                For Y = 0 To 238
                k = Y * 1024 + 512
                For X = 0 To 511
                    vBuffer2x16(k) = (vBuffer2x16(k - 512) And &H7BDE&) \ 2 + (vBuffer2x16(k + 512) And &H7BDE&) \ 2
                    k = k + 1
                Next X, Y
            Else
                For Y = 1 To 238
                i = Y * 256&
                k = Y * 1024&
                c = vBuffer16(i - 1)
                p21 = vBuffer16(i)
                fscan = (Y And 63) = count64
                For X = IIf(X > 1, 0, 1) To 255
                    If vBuffer16(i + 2) <> oldvBuffer16(i + 2) Then
                        If t <= 0 Then
                            c = vBuffer16(i - 1)
                            p21 = vBuffer16(i)
                        End If
                        t = 5
                    End If
                    If t > 0 Or fscan Then
                        t = t - 1
                        p01 = c
                        c = p21
                        p21 = vBuffer16(i + 1)
                        b = (c And &H7BDE&) \ 2
                        p10 = vBuffer16(i - 256)
                        p12 = vBuffer16(i + 256)
                        If p01 = p10 And p12 = p21 Then
                            If p10 = p12 And vBuffer16(i - 257) = c Then
                                vBuffer2x16(k) = c
                                vBuffer2x16(k + 1) = (p10 And &H7BDE&) \ 2 + b
                                vBuffer2x16(k + 512) = (p12 And &H7BDE&) \ 2 + b
                                vBuffer2x16(k + 513) = c
                            Else
                                vBuffer2x16(k) = (p01 And &H7BDE&) \ 2 + b
                                vBuffer2x16(k + 1) = c
                                vBuffer2x16(k + 512) = c
                                vBuffer2x16(k + 513) = (p21 And &H7BDE&) \ 2 + b
                            End If
                        ElseIf p10 = p21 And p01 = p12 Then
                            vBuffer2x16(k) = c
                            vBuffer2x16(k + 1) = (p10 And &H7BDE&) \ 2 + b
                            vBuffer2x16(k + 512) = (p12 And &H7BDE&) \ 2 + b
                            vBuffer2x16(k + 513) = c
                        Else
                            vBuffer2x16(k) = c
                            vBuffer2x16(k + 1) = c
                            vBuffer2x16(k + 512) = c
                            vBuffer2x16(k + 513) = c
                        End If
                    End If
                    i = i + 1
                    k = k + 2
                Next X, Y
                MemCopy oldvBuffer16(0), vBuffer16(0), 122880
            End If
            blit15 vBuffer2x16, frmNES.picScreen, 512, 480
        Else
            blit15 vBuffer16, frmNES.picScreen, 256, 240
        End If
    Case Else
        If MotionBlur Then
            For i = 0 To 31
                p32(i) = (pal(VRAM(i + &H3F00)) And &HFEFEFE) \ 2
            Next i
            For i = 0 To 61439
                vBuffer32(i) = (vBuffer32(i) And &HFEFEFE) \ 2 + p32(vBuffer(i))
            Next i
        Else
            For i = 0 To 31
                p32(i) = pal(VRAM(i + &H3F00))
            Next i
            For i = 0 To 61439 Step 4
                vBuffer32(i) = p32(vBuffer(i))
                vBuffer32(i + 1) = p32(vBuffer(i + 1))
                vBuffer32(i + 2) = p32(vBuffer(i + 2))
                vBuffer32(i + 3) = p32(vBuffer(i + 3))
            Next i
        End If
    
        If Smooth2x Then
            If Smooth2x = 1 Then
                For Y = 0 To 239
                i = Y * 256&
                k = Y * 1024&
                c = 0
                For X = 0 To 255
                    b = vBuffer32(i)
                    vBuffer2x32(k + 1) = b
                    b = (b And &HFEFEFE) \ 2
                    vBuffer2x32(k) = b + c
                    c = b
                    i = i + 1
                    k = k + 2
                Next X, Y
                For Y = 0 To 238
                k = Y * 1024 + 512
                For X = 0 To 511
                    vBuffer2x32(k) = (vBuffer2x32(k - 512) And &HFEFEFE) \ 2 + (vBuffer2x32(k + 512) And &HFEFEFE) \ 2
                    k = k + 1
                Next X, Y
            Else
                t = n
                For Y = 1 To 238
                i = Y * 256&
                k = Y * 1024&
                c = vBuffer32(i - 1)
                p21b = vBuffer32(i)
                t = 1 - t
                For X = IIf(X > 1, 0, 1) To 255
                    t = 1 - t
                    p01b = c
                    c = p21b
                    p21b = vBuffer32(i + 1)
                    If t Or c <> vBuffer2x32(k) Then
                        b = (c And &HFEFEFE) \ 2
                        p10b = vBuffer32(i - 256)
                        p12b = vBuffer32(i + 256)
                        If p01b = p10b And p12b = p21b Then
                            If p10b = p12b And vBuffer32(i - 257) = c Then
                                vBuffer2x32(k) = c
                                vBuffer2x32(k + 1) = (p10b And &HFEFEFE) \ 2 + b
                                vBuffer2x32(k + 512) = (p12b And &HFEFEFE) \ 2 + b
                                vBuffer2x32(k + 513) = c
                            Else
                                vBuffer2x32(k) = (p01b And &HFEFEFE) \ 2 + b
                                vBuffer2x32(k + 1) = c
                                vBuffer2x32(k + 512) = c
                                vBuffer2x32(k + 513) = (p21b And &HFEFEFE) \ 2 + b
                            End If
                        ElseIf p10b = p21b And p01b = p12b Then
                            vBuffer2x32(k) = c
                            vBuffer2x32(k + 1) = (p10b And &HFEFEFE) \ 2 + b
                            vBuffer2x32(k + 512) = (p12b And &HFEFEFE) \ 2 + b
                            vBuffer2x32(k + 513) = c
                        Else
                            vBuffer2x32(k) = c
                            vBuffer2x32(k + 1) = c
                            vBuffer2x32(k + 512) = c
                            vBuffer2x32(k + 513) = c
                        End If
                    End If
                    i = i + 1
                    k = k + 2
                Next X, Y
            End If
            
            
            Blit vBuffer2x32, frmNES.picScreen, 512, 480
        Else
            Blit vBuffer32, frmNES.picScreen, 256, 240
        End If
    End Select

End Sub



Public Sub RenderScanline(ByVal Scanline As Long)
DoMirror ' set the mirroring
'scanline based sprite rendering
If Scanline > 239 Then Exit Sub

If Scanline = 0 Then
    If render Then MemFill vBuffer(0), 256& * 240&, 16

    'temporary measure until the mirroring problems can be fixed
    Static pm, pmx
    If MirrorXor <> pmx Then
        If MirrorXor = 0 Then
            Mirroring = 2
        ElseIf MirrorXor = &H400& Then
            Mirroring = 0
        Else
            Mirroring = 1
        End If
    ElseIf pm <> Mirroring Then
        If Mirroring = 0 And MirrorXor <> &H400& Then
                MirrorXor = &H400&
        ElseIf Mirroring = 1 And MirrorXor <> &H800& Then
                MirrorXor = &H800&
        ElseIf Mirroring = 2 And MirrorXor <> 0 Then
                MirrorXor = &H800&
        End If
    End If
    pm = Mirroring
    pmx = MirrorXor
End If

If ((PPU_Control2 And 16) = 0 And newScroll) Or Not render Then
    If Scanline > SpriteRAM(0) + 8 Then PPU_Status = PPU_Status Or 64
    Exit Sub
End If

Dim v As Long
Dim nt2 As Byte
'still some bugs in Little Nemo, Kirby.
If newScroll Then
    If Scanline = 0 Then
        PPUAddress = PPUAddress2
    Else
        PPUAddress = (PPUAddress And &HFBE0&) Or (PPUAddress2 And &H41F&)
    End If
   
    NameTable = &H2000& + (PPUAddress And &HC00)
    nt2 = (NameTable And &HC00&) \ &H400&

    HScroll = (PPUAddress And 31) * 8 + HScroll2
    VScroll = (PPUAddress \ 32 And 31) * 8 Or ((PPUAddress \ &H1000&) And 7)
    
    'If PPUAddress And 8192 Then VScroll = VScroll + 240
    
    VScroll = VScroll - Scanline
    
    v = PPUAddress
    
    ' the following "if" and contents were ported from Nester
    If (v And &H7000&) = &H7000& Then '/* is subtile y offset == 7? */
        v = v And &H8FFF& '/* subtile y offset = 0 */
        If (v And &H3E0&) = &H3A0& Then '/* name_tab line == 29? */
            v = v Xor &H800&   '/* switch nametables (bit 11) */
            v = v And &HFC1F&  '/* name_tab line = 0 */
        Else
            If (v And &H3E0&) = &H3E0& Then  '/* line == 31? */
                v = v And &HFC1F&  '/* name_tab line = 0 */
            Else
                v = v + &H20&
            End If
        End If
    Else
        v = v + &H1000&
    End If
    
    PPUAddress = v And &HFFFF&
End If

If Keyboard(219) And 1 Then VScroll = Frames Mod 240
If Keyboard(221) And 1 Then HScroll = Frames And 255


'If Scanline < 8 Then PPU_Status = PPU_Status And &H3F
If Scanline = 239 Then PPU_Status = PPU_Status Or &H80
If Not render Then 'speed optimization when drawing skipped frame
    If PPU_Status And &H40 Then Exit Sub
    If Scanline > SpriteRAM(0) + 8 Then PPU_Status = PPU_Status Or &H40
    Exit Sub
End If

'If Scanline < 50 Then HScroll = 127


'VScroll = Scanline + (Frames Mod 240)

Dim TileRow As Byte, TileYOffset As Long 'Integer
Dim TileCounter As Long 'Integer
Dim Color As Long
Dim TileIndex As Byte, Byte1 As Byte, Byte2 As Byte
Dim LookUp As Byte, addToCol As Long
Dim pixel As Long 'Integer
Dim X As Long, Aa As Long ', c As Long, pc As Long, px As Long
Dim m As Long
Dim sc As Long
Dim atrtab As Long
Static phs As Long, pvs As Long
Dim Y As Long




If tilebased Then
    Dim h As Long
    If PPU_Control1 And &H20 Then h = 16 Else h = 8
    If (PPU_Status And &H40) = 0 Then If Scanline > SpriteRAM(0) + h Then PPU_Status = PPU_Status Or &H40
    
    If Scanline = 0 Then
        DrawSprites True
    ElseIf Scanline = 236 Then
        DrawSprites False
    End If
Else
    'draw background sprites
    'RenderSprites Scanline, True
End If

sc = Scanline + VScroll
'If newScroll And sc >= 240 Then sc = sc - 1 'quick fix for mysterious problem
If sc > 480 Then sc = sc - 480

'draw background
TileRow = (sc \ 8) Mod 30
TileYOffset = sc And 7

If (Not tilebased) Or VScroll <> pvs Or HScroll <> phs Or TileYOffset = 0 Then
If TileYOffset = 0 Then
    pvs = VScroll
    phs = HScroll
End If

If Not newScroll Then
    If sc < 240 Then
        NameTable = &H2000& + (&H400& * (PPU_Control1 And &H3))
        nt2 = (NameTable And &HC00&) \ &H400&
    Else
        NameTable = &H2000& + (&H400& * (PPU_Control1 And &H3)) Xor &H800
        nt2 = (NameTable And &HC00&) \ &H400&
    End If
End If


atrtab = &H3C0
PatternTable = (PPU_Control1 And &H10) * &H100&
For TileCounter = HScroll \ 8 To 31
    TileIndex = nt(mirror(nt2), TileCounter + TileRow * 32)
    'TileIndex = VRAM(NameTable + TileCounter + TileRow * 32)
    If Mapper = 9 Or Mapper = 10 Then
        If PatternTable = &H0 Then
            map9_latch TileIndex, False
        ElseIf PatternTable = &H1000& Then
            map9_latch TileIndex, True
        End If
    End If
    X = TileCounter * 8 - HScroll + 7
    If X < 7 Then m = X Else m = 7
    X = X + Scanline * 256&
    LookUp = nt(mirror(nt2), (&H3C0& + TileCounter \ 4 + (TileRow \ 4) * &H8&))
    Select Case (TileCounter And 2) Or (TileRow And 2) * 2
        Case 0
            addToCol = LookUp * 4 And 12
        Case 2
            addToCol = LookUp And 12
        Case 4
            addToCol = LookUp \ 4 And 12
        Case 6
            addToCol = LookUp \ 16 And 12
    End Select
    If tilebased And TileYOffset = 0 Then
        For Y = 0 To 7
            Byte1 = VRAM(PatternTable + TileIndex * 16 + Y)
            Byte2 = VRAM(PatternTable + TileIndex * 16 + 8 + Y)
            Aa = Byte1 * 2048& + Byte2 * 8
            For pixel = m To 0 Step -1
                Color = tLook(Aa + pixel)
                If Color Then vBuffer(X - pixel) = Color Or addToCol
            Next pixel
            X = X + 256
        Next Y
    Else
        Byte1 = VRAM(PatternTable + TileIndex * 16 + TileYOffset)
        Byte2 = VRAM(PatternTable + TileIndex * 16 + 8 + TileYOffset)
        Aa = Byte1 * 2048& + Byte2 * 8
        For pixel = m To 0 Step -1
            Color = tLook(Aa + pixel)
            If Color Then vBuffer(X - pixel) = Color Or addToCol
        Next pixel
    End If
Next TileCounter

If newScroll Then
    'NameTable = &H2000 + (PPUAddress And &H800) Xor &H400
    NameTable = NameTable Xor &H400
    nt2 = (NameTable And &HC00&) \ &H400&
Else
    If sc < 240 Then
        NameTable = &H2000& + (&H400& * (PPU_Control1 And &H3)) Xor &H400
        nt2 = (NameTable And &HC00&) \ &H400&
    Else
        NameTable = &H2000& + (&H400& * (PPU_Control1 And &H3)) Xor &HC00
        nt2 = (NameTable And &HC00&) \ &H400&
    End If
End If
atrtab = &H3C0
For TileCounter = 0 To HScroll \ 8
    TileIndex = nt(mirror(nt2), TileCounter + TileRow * 32)
    'TileIndex = VRAM(NameTable + (TileCounter + TileRow * 32))
    If Mapper = 9 Or Mapper = 10 Then
        If PatternTable = &H0 Then
            map9_latch TileIndex, False
        ElseIf PatternTable = &H1000& Then
            map9_latch TileIndex, True
        End If
    End If
    X = TileCounter * 8 + 256 - HScroll + 7
    If X > 255 Then m = X - 255 Else m = 0
    X = X + Scanline * 256&
    LookUp = nt(mirror(nt2), (&H3C0& + TileCounter \ 4 + (TileRow \ 4) * &H8&))
    Select Case (TileCounter And 2) Or (TileRow And 2) * 2
        Case 0
            addToCol = LookUp * 4 And 12
        Case 2
            addToCol = LookUp And 12
        Case 4
            addToCol = LookUp \ 4 And 12
        Case 6
            addToCol = LookUp \ 16 And 12
    End Select
    If tilebased And TileYOffset = 0 Then
        For Y = 0 To 7
            Byte1 = VRAM(PatternTable + TileIndex * 16 + Y)
            Byte2 = VRAM(PatternTable + TileIndex * 16 + 8 + Y)
            Aa = Byte1 * 2048& + Byte2 * 8
            For pixel = 7 To m Step -1
                Color = tLook(Aa + pixel)
                If Color Then vBuffer(X - pixel) = Color Or addToCol
            Next pixel
            X = X + 256
        Next Y
    Else
        Byte1 = VRAM(PatternTable + (TileIndex * 16) + TileYOffset)
        Byte2 = VRAM(PatternTable + (TileIndex * 16) + 8 + TileYOffset)
        Aa = Byte1 * 2048& + Byte2 * 8
        For pixel = 7 To m Step -1
            Color = tLook(Aa + pixel)
            If Color Then vBuffer(X - pixel) = Color Or addToCol
        Next pixel
    End If
Next TileCounter
End If

If Not tilebased Then
''draw foreground sprites
'draw all sprites
RenderSprites Scanline - 1 ', False
End If
End Sub
Public Sub RenderSprites(ByVal Scanline As Long) ' , topLayer As Boolean)
Dim solid(264) As Boolean

If (PPU_Control2 And 16) = 0 Then Exit Sub

Dim TileRow As Byte, TileYOffset As Long 'Integer
Dim TileCounter As Long 'Integer
Dim Color As Byte
Dim TileIndex As Byte, Byte1 As Byte, Byte2 As Byte
Dim addToCol As Long
Dim h As Byte
Dim minX As Long
Dim X1 As Long, y1 As Long
Dim ptable As Long

TileRow = Scanline \ 8
TileYOffset = Scanline And 7
If PPU_Control1 And &H20 Then
    h = 16
Else
    h = 8
End If
If PPU_Control1 And &H8 Then
    If h = 8 Then
        PatternTable = &H1000&
    End If
Else
    If h = 8 Then
        PatternTable = &H0
    End If
End If
If PPU_Control2 And &H8 Then minX = 0 Else minX = 8
Dim spr As Long 'Integer
Dim SpriteAddr As Integer
Dim i As Long, X As Long, Aa As Long, v As Long
Dim ontop As Boolean
Dim attr As Byte
i = Scanline * 256&
For spr = 0 To 63 '63 To 0 Step -1
    SpriteAddr = 4 * spr
    y1 = SpriteRAM(SpriteAddr) + 1
    If y1 <= Scanline And y1 > Scanline - h Then
        attr = SpriteRAM(SpriteAddr + 2)
        ontop = (attr And 32) = 0
        'If (attr And 32) = 0 Xor topLayer Then
            X1 = SpriteRAM(SpriteAddr + 3)
            If X1 >= minX Then
                If render And X1 < 248 Then
                    addToCol = &H10 + (attr And 3) * 4
                    
                    TileIndex = SpriteRAM(SpriteAddr + 1)
                    If Mapper = 9 Or Mapper = 10 Then
                        If PatternTable = &H0 Then
                            map9_latch TileIndex, False
                        ElseIf PatternTable = &H1000& Then
                            map9_latch TileIndex, True
                        End If
                    End If
                    If h = 16 Then
                        If TileIndex And 1 Then
                            PatternTable = &H1000
                            TileIndex = TileIndex Xor 1
                        Else
                            PatternTable = 0
                        End If
                    End If
                    If attr And 128 Then 'vertical flip
                        v = y1 - Scanline - 1
                    Else
                        v = Scanline - y1
                    End If
                    v = v And h - 1
                    If v >= 8 Then v = v + 8
                    Byte1 = VRAM(PatternTable + (TileIndex * 16) + v)
                    Byte2 = VRAM(PatternTable + (TileIndex * 16) + 8 + v)
                    '#If 0 Then
                    'real sprite 0 detection
                    If spr = 0 And (PPU_Status And 64) = 0 Then
                        If attr And 64 Then 'horizontal flip
                            Aa = i + X1
                            For X = 0 To 7
                                If Byte1 And pow2(X) Then Color = 1 Else Color = 0
                                If Byte2 And pow2(X) Then Color = Color + 2
                                If Color Then
                                    'If Not solid(X1 + x) Then
                                        If vBuffer(Aa + X) And 3 And (PPU_Status And 64) = 0 Then
                                            PPU_Status = PPU_Status Or 64
                                            If ontop Then vBuffer(Aa + X) = addToCol Or Color
                                        Else
                                            vBuffer(Aa + X) = addToCol Or Color
                                        End If
                                        solid(X1 + X) = True
                                    'End If
                                End If
                            Next X
                        Else
                            Aa = i + X1 + 7
                            For X = 7 To 0 Step -1
                                If Byte1 And pow2(X) Then Color = 1 Else Color = 0
                                If Byte2 And pow2(X) Then Color = Color + 2
                                If Color Then
                                    'If Not solid(X1 + 7 - x) Then
                                        If vBuffer(Aa - X) And 3 And (PPU_Status And 64) = 0 Then
                                            PPU_Status = PPU_Status Or 64
                                            If ontop Then vBuffer(Aa - X) = addToCol Or Color
                                        Else
                                            vBuffer(Aa - X) = addToCol Or Color
                                        End If
                                        solid(X1 + 7 - X) = True
                                    'End If
                                End If
                            Next X
                        End If
                    Else
                    '#End If
                        If attr And 64 Then 'horizontal flip
                            Aa = i + X1
                            For X = 0 To 7 'draw yellow block for now.
                                If Byte1 And pow2(X) Then Color = 1 Else Color = 0
                                If Byte2 And pow2(X) Then Color = Color + 2
                                If Color Then
                                    If Not solid(X1 + X) Then
                                        If ontop Then
                                            vBuffer(Aa + X) = addToCol Or Color
                                        ElseIf (vBuffer(Aa + X) And 3) = 0 Then
                                            vBuffer(Aa + X) = addToCol Or Color
                                        End If
                                        solid(X1 + X) = True
                                    End If
                                End If
                            Next X
                        Else
                            Aa = i + X1 + 7
                            For X = 7 To 0 Step -1 'draw yellow block for now.
                                If Byte1 And pow2(X) Then Color = 1 Else Color = 0
                                If Byte2 And pow2(X) Then Color = Color + 2
                                If Color Then
                                    If Not solid(X1 + 7 - X) Then
                                        If ontop Then
                                            vBuffer(Aa - X) = addToCol Or Color
                                        ElseIf (vBuffer(Aa - X) And 3) = 0 Then
                                            vBuffer(Aa - X) = addToCol Or Color
                                        End If
                                        solid(X1 + 7 - X) = True
                                    End If
                                End If
                            Next X
                        End If
                    End If
                End If
                'force sprite 0 detection for those games with background glitches
                'otherwise, some crash
                If spr = 0 Then If Scanline = y1 + h - 1 Then PPU_Status = PPU_Status Or &H40 'claim we hit sprite #0
            End If
        'End If
    End If
Next spr
End Sub
Public Sub map6_write(Address As Long, value As Byte)
'WIP: If anyone wants to help me out on this one, thanks...
If Address < &H42FC& Then Exit Sub
    
    Select Case Address
        Case &H42FC& To &H42FD&: ' Unknown
        Case &H42FE& ' Page Select
            If value And &H20 Then
                NameTable = &H2400&
            Else
                NameTable = &H2000&
            End If
        Case &H42FF& ' Mirroring
            Mirroring = (value And &H20) \ &H20
            DoMirror
        Case &H4501&
            map6_irqon = 0
        Case &H4502&
            tmpLatch = value
        Case &H4503&
            IRQCounter = (value * &H100&) + tmpLatch
            map6_irqon = 1
        Case &H8000& To &HFFFF&
            reg8 = (value And &HF) * 2
            regA = reg8 + 1
            Select8KVROM value And &H3
            SetupBanks
    End Select
End Sub
Public Function map6_hblank(Scanline) As Byte
    If (map6_irqon <> 0) Then
        IRQCounter = IRQCounter + 1
        If (IRQCounter >= &HFFFF&) Then
            IRQCounter = 0
            irq6502
        End If
    End If
End Function
Public Sub map5_write(Address As Long, value As Byte)

End Sub
Public Sub map90_write(Address As Long, value As Byte)
    Select Case Address
        Case &H8000&: reg8 = value
        Case &H8001&: regA = value
        Case &H8002&: regC = value
        Case &H8003&: regE = value
    End Select
    If Address >= &H8000& And Address <= &H8003& Then SetupBanks
End Sub
Public Sub map24_write(Address As Long, value As Byte)
    Select Case Address
        Case &H8000&
            reg8 = value * 2
            regA = reg8 + 1
            SetupBanks
        Case &HB003&
            Mirroring = ((value And &HC) \ &H4)
        Case &HC000&
            regC = value
            SetupBanks
        Case &HD000& To &HD003&
            Select1KVROM value, (Address And &H3)
        Case &HE000& To &HE003&
            Select1KVROM value, (Address And &H3) + 4
        Case &HF000&
            map24_irqv = value
        Case &HF001&
            map24_irqon = (value And &H1)
        Case &HF002&
            map24_irqv = 0
    End Select
End Sub
Public Sub map24_irq()
    If map24_irqon = 0 Then Exit Sub
    map24_irqv = map24_irqv + 1
    If (map24_irqv = &HFF) And (map24_irqon) Then
        map24_irqon = 0
        map24_irqv = 0
        irq6502
    End If
End Sub
Public Sub map13_write(Address As Long, value As Byte)
    Dim prg_bank As Byte
    
    prg_bank = (value And &H30) \ 16
            
    reg8 = prg_bank * 4: regA = reg8 + 1: regC = reg8 + 2: regE = reg8 + 3: SetupBanks
    
    Select4KVROM 0, 2
    Select4KVROM (value And &H3), 3
   
    latch13 = value
    
End Sub
Public Sub map16_write(Address As Long, value As Byte)
    Select Case (Address And &HD)
        Case &H0 To &H7: Select1KVROM value, Address And &H7
        Case &H8: reg8 = value * 2: regA = reg8 + 1: SetupBanks
        Case &H9: Mirroring = (value And &H1)
        Case &HA: If value Then MMC16_IrqOn = 1 Else MMC16_IrqOn = 0
        Case &HB: tmpLatch = value
        Case &HC: MMC16_Irq = (value * &H100&) + tmpLatch
        Case &HD: ' Unknown
    End Select
End Sub
Public Sub map65_write(Address As Long, value As Byte)
' Mapper #65 - Irem H-3001
    Select Case Address
        Case &H8000&
            reg8 = value
            SetupBanks
        Case &H9003& ' Mirroring
        Case &H9005& ' IRQ Control 1
        Case &H9006& ' IRQ Control 2
        Case &HA000&
            regA = value
            SetupBanks
        Case &HB000& To &HB007&
            Select1KVROM value, (Address And &H7)
        Case &HC000&
            regC = value
            SetupBanks
    End Select
End Sub
Public Sub map19_write(Address As Long, value As Byte)

If Address < &H5000& Then Exit Sub
    
    Select Case Address
        Case &H5000& To &H57FF&
            tmpLatch = value
        Case &H5800& To &H5FFF&
            If value And &H80 Then MIRQOn = 1 Else MIRQOn = 0
            MMC19_IRQCount = ((value And &H7F) * &H100&) + tmpLatch
        Case &H8000& To &H87FF&
            Select1KVROM value, 0
        Case &H8800& To &H8FFF&
            Select1KVROM value, 1
        Case &H9000& To &H97FF&
            Select1KVROM value, 2
        Case &H9800& To &H9FFF&
            Select1KVROM value, 3
        Case &HA000& To &HA7FF&
            Select1KVROM value, 4
        Case &HA800& To &HAFFF&
            Select1KVROM value, 5
        Case &HB000& To &HB7FF&
            Select1KVROM value, 6
        Case &HB800& To &HBFFF&
            Select1KVROM value, 7
        Case &HC000& To &HC7FF&
            If value < &HE0 Then Select1KVROM value, 8
        Case &HC800& To &HC8FF&
            If value < &HE0 Then Select1KVROM value, 9
        Case &HD000& To &HD7FF&
            If value < &HE0 Then Select1KVROM value, 10
        Case &HD800& To &HD8FF&
            If value < &HE0 Then Select1KVROM value, 11
        Case &HE000& To &HE7FF&
            reg8 = value
            SetupBanks
        Case &HE800& To &HEFFF&
            regA = value
            SetupBanks
        Case &HF000& To &HF7FF&
            regC = value
            SetupBanks
End Select

End Sub
Public Sub map19_irq()
    If MIRQOn = 1 Then
        MMC19_IRQCount = MMC19_IRQCount + 1
        If MMC19_IRQCount = &H7FFF& Then
            irq6502
        End If
    End If
End Sub
Public Sub map64_write(Address As Long, value As Byte)
    ' Tengen RAMBO-1, like MMC3 i guess
    Select Case Address
        Case &H8000&
            cmd = value And &HF
            prg = (value And &H40)
            chr1 = (value And &H80)
        Case &H8001&
            Select Case cmd
                Case 0
                    If (chr1) Then
                        Select1KVROM value, 4
                        Select1KVROM value + 1, 5
                    Else
                        Select1KVROM value, 0
                        Select1KVROM value, 1
                    End If
                Case 1
                    If (chr1) Then
                        Select1KVROM value, 6
                        Select1KVROM value + 1, 7
                    Else
                        Select1KVROM value, 2
                        Select1KVROM value + 1, 3
                    End If
                Case 2: If (chr1) Then Select1KVROM value, 0 Else Select1KVROM value, 4
                Case 3: If (chr1) Then Select1KVROM value, 1 Else: Select1KVROM value, 5
                Case 4: If (chr1) Then Select1KVROM value, 2 Else Select1KVROM value, 6
                Case 5: If (chr1) Then Select1KVROM value, 3 Else Select1KVROM value, 7
                Case 6
                    If (prg) Then
                        regA = value
                    Else
                        reg8 = value
                    End If
                    SetupBanks
                Case 7
                    If (prg) Then
                        regC = value
                    Else
                        reg8 = value
                    End If
                    SetupBanks
                Case 8: Select1KVROM value, 1
                Case 9: Select1KVROM value, 3
                Case &HF
                    If (prg) Then
                        reg8 = value
                    Else
                        regC = value
                    End If
                    SetupBanks
            End Select
        Case &HA000&
            Mirroring = (value And &H1)
            DoMirror
    End Select
End Sub

Public Sub DoMirror()
    MirrorXor = (((Mirroring + 1) Mod 3) * &H400&)
    If Mirroring = 0 Then
        mirror(0) = 0: mirror(1) = 0: mirror(2) = 1: mirror(3) = 1
    ElseIf Mirroring = 1 Then
        mirror(0) = 0: mirror(1) = 1: mirror(2) = 0: mirror(3) = 1
    ElseIf Mirroring = 2 Then
        mirror(0) = 0: mirror(1) = 0: mirror(2) = 0: mirror(3) = 0
    ElseIf Mirroring = 4 Then
        mirror(0) = 0: mirror(1) = 1: mirror(2) = 2: mirror(3) = 3
    End If
End Sub
Public Function Read6502(ByVal Address As Long) As Byte
Dim tmp As Byte
    ' Rewritten 29.03.02
    Select Case Address
        Case &H0 To &H1FFF&: Read6502 = bank0(Address And &H7FF&) ' NES ram 0-7ff mirrored at 800 1000 1800
        Case &H2000& To &H3FFF&
            Select Case (Address And &H7&)
                Case &H0&: Read6502 = ppuLatch
                Case &H1&: Read6502 = ppuLatch
                Case &H2&
                    Dim ret As Byte
                    ScrollToggle = 0
                    PPU_AddressIsHi = True
                    ret = (ppuLatch And &H1F) Or PPU_Status
                    If (ret And &H80) Then PPU_Status = (PPU_Status And &H60)
                    Read6502 = ret
                Case &H4&
                    tmp = ppuLatch
                    ppuLatch = SpriteRAM(SpriteAddress)
                    SpriteAddress = (SpriteAddress + 1) And &HFF
                    Read6502 = tmp
                Case &H5&: Read6502 = ppuLatch
                Case &H6&: Read6502 = ppuLatch
                Case &H7&
                    tmp = ppuLatch
                    If Mapper = 9 Then
                        If PPUAddress < &H2000& Then
                            map9_latch tmp, (PPUAddress And &H1000&)
                        End If
                    End If
                    If PPUAddress >= &H2000& And PPUAddress <= &H2FFF& Then
                        ppuLatch = nt(mirror((PPUAddress And &HC00&) \ &H400&), PPUAddress And &H3FF&)
                    Else
                        ppuLatch = VRAM(PPUAddress And &H3FFF&)
                    End If
                    If (PPU_Control1 And &H4) Then
                        PPUAddress = PPUAddress + 32
                    Else
                        PPUAddress = PPUAddress + 1
                    End If
                    PPUAddress = (PPUAddress And &H3FFF&)
                    Read6502 = tmp
            End Select
        Case &H4000& To &H4013&, &H4015&
            Read6502 = Sound(Address - &H4000&)
        Case &H4016& ' Joypad1
            Read6502 = Joypad1(Joypad1_Count)
            Joypad1_Count = (Joypad1_Count + 1) And 7
        Case &H6000& To &H7FFF&: Read6502 = bank6(Address And &H1FFF&)
        Case &H8000& To &H9FFF&: Read6502 = bank8(Address And &H1FFF&)
        Case &HA000& To &HBFFF&: Read6502 = bankA(Address And &H1FFF&)
        Case &HC000& To &HDFFF&: Read6502 = bankC(Address And &H1FFF&)
        Case &HE000& To &HFFFF&: Read6502 = bankE(Address And &H1FFF&)
    End Select
End Function
Public Sub Write6502(ByVal Address As Long, ByVal value As Byte)
    ' Rewritten 29.03.02
    If Address >= &H2000& And Address <= &H3FFF& Then ppuLatch = value
    Select Case Address
        Case &H0& To &H1FFF&: bank0(Address And &H7FF&) = value
        Case &H2000& To &H3FFF&
            Select Case (Address And &H7)
                Case &H0&
                    PPU_Control1 = value
                    PPUAddress2 = (PPUAddress2 And &HF3FF&) Or (value And 3) * &H400&
                Case &H1&: PPU_Control2 = value
                Case &H2&: ppuLatch = value
                Case &H3&: SpriteAddress = value
                Case &H4&
                    SpriteRAM(SpriteAddress) = value
                    SpriteAddress = (SpriteAddress + 1) And &HFF
                    SpritesChanged = True
                Case &H5&
                    If PPU_AddressIsHi Then
                        HScroll2 = value And 7
                        PPUAddress2 = (PPUAddress2 And &HFFE0&) Or value \ 8
                        If Not newScroll Then
                            HScroll = value
                        End If
                        PPU_AddressIsHi = False
                    Else
                        PPUAddress2 = (PPUAddress2 And &H8C1F&) Or (value And &HF8) * 4 Or (value And 7) * &H1000&
                        If Not newScroll Then
                            VScroll = value
                            If VScroll > 240 Then VScroll = 0
                        End If
                        PPU_AddressIsHi = True
                    End If
                Case &H6&
                    If PPU_AddressIsHi Then
                        PPUAddress = (PPUAddress And &HFF) Or ((value And &H3F) * &H100&)
                        PPU_AddressIsHi = False
                    Else
                        PPUAddress = (PPUAddress And &H7F00&) Or value
                        PPU_AddressIsHi = True
                    End If
                    If Not newScroll Then
                    'seems to work better to ignore when display is disabled.
                        If PPU_Control2 And 16 Then
                            HScroll = (PPUAddress And 31) * 8 + HScroll2
                            VScroll = (PPUAddress \ 32 And 31) * 8 Or (PPUAddress \ &H1000& And 7)
                            If CurrentLine < 240 Then VScroll = VScroll - CurrentLine
                        End If
                    End If
                Case &H7&
                    ppuLatch = value
                    If Mapper = 9 Then
                        If PPUAddress <= &H1FFF& And PPUAddress >= &H0& Then
                            map9_latch value, (PPUAddress And &H1000&)
                        End If
                    End If
                    If PPUAddress >= &H2000& And PPUAddress <= &H2FFF& Then
                        nt(mirror((PPUAddress And &HC00&) \ &H400&), PPUAddress And &H3FF&) = value
                    Else
                        VRAM(PPUAddress) = value
                        If (PPUAddress And &HFFEF&) = &H3F00& Then
                            VRAM(PPUAddress Xor 16) = value
                        End If
                    End If
                    If (PPU_Control1 And &H4) Then
                        PPUAddress = (PPUAddress + 32)
                    Else
                        PPUAddress = (PPUAddress + 1)
                    End If
                    PPU_AddressIsHi = True
                    PPUAddress = (PPUAddress And &H3FFF&)
            End Select
        Case &H4000& To &H4013&
            Sound(Address - &H4000&) = value
            Dim n
            n = (Address - &H4000&) \ 4
            If n < 4 Then ChannelWrite(n) = True
        Case &H4014&
            MemCopy SpriteRAM(0), bank0(value * &H100&), &HFF
            SpritesChanged = True
        Case &H4015&: SoundCtrl = value
        Case &H6000& To &H7FFF&
            If SpecialWrite6000 Then
                MapperWrite Address, value
            Else
                bank6(Address And &H1FFF&) = value
            End If
        Case &H8000& To &HFFFF&: MapperWrite Address, value
    End Select
End Sub
Public Sub map11_write(Address As Long, value As Byte)
    reg8 = 4 * (value And &HF)
    regA = reg8 + 1
    regC = reg8 + 2
    regE = reg8 + 3
    SetupBanks
    Select8KVROM (value And &HF0)
End Sub
Public Sub map15_write(Address As Long, value As Byte)
            Map15_BankAddr = (value And &H3F) * 2
            Select Case (Address And &H3)
                Case &H0
                    map15_swapaddr = (value And &H80)
                    reg8 = Map15_BankAddr
                    regA = Map15_BankAddr + 1
                    regC = Map15_BankAddr + 2
                    regE = Map15_BankAddr + 3
                    SetupBanks
                    Mirroring = (value And &H40) \ &H40
                Case &H1
                    Map15_SwapReg = (value And &H80)
                    regC = Map15_BankAddr
                    regE = Map15_BankAddr + 1
                    SetupBanks
            Case &H2
                If (value And &H80) Then
                    reg8 = Map15_BankAddr + 1
                    regA = Map15_BankAddr + 1
                    regC = Map15_BankAddr + 1
                    regE = Map15_BankAddr + 1
                Else
                    reg8 = Map15_BankAddr
                    regA = reg8
                    regC = regA
                    regE = regC
                End If
            SetupBanks
        Case &H3
            Map15_SwapReg = (value And &H80)
            regC = Map15_BankAddr
            regE = Map15_BankAddr + 1
            SetupBanks
            Mirroring = (value And &H40&) \ &H40&
            ' TODO: Add mirroring
    End Select
End Sub
Public Sub map18_write(Address As Long, value As Byte)
    Address = (Address And &HF003&)
    Select Case Address
        Case &H8000&
    End Select
End Sub
Public Sub map32_write(Address As Long, value As Byte)
    Select Case Address
        Case &H8000& To &H8FFF&
            If (MMC32_Switch And &H2) = 2 Then
                regC = value
                SetupBanks
            Else
                reg8 = value
                SetupBanks
            End If
        Case &H9000& To &H9FFF&
            Mirroring = (value And &H1)
            MMC32_Switch = value
            DoMirror
        Case &HA000& To &HAFFF&
            regA = value
            SetupBanks
        Case &HBFF0&: Select1KVROM value, 0
        Case &HBFF1&: Select1KVROM value, 1
        Case &HBFF2&: Select1KVROM value, 2
        Case &HBFF3&: Select1KVROM value, 3
        Case &HBFF4&: Select1KVROM value, 4
        Case &HBFF5&: Select1KVROM value, 5
        Case &HBFF6&: Select1KVROM value, 6
        Case &HBFF7&: Select1KVROM value, 7
    End Select
End Sub
Public Sub map33_write(Address As Long, value As Byte)
    Select Case Address
        Case &H8000&: reg8 = value: SetupBanks
        Case &H8001&: regA = value: SetupBanks
        Case &H8002&: Select2KVROM value, 0
        Case &H8003&: Select2KVROM value, 2
        Case &HA000&: Select1KVROM value, 4
        Case &HA001&: Select1KVROM value, 5
        Case &HA002&: Select1KVROM value, 6
        Case &HA003&: Select1KVROM value, 7
        Case &HC000, &HC001, &HE000&
    End Select
End Sub
Public Sub map34_write(Address As Long, value As Byte)
    Select Case Address
        Case &H7FFD&
            reg8 = 4 * (value)
            regA = reg8 + 1
            regC = regA + 1
            regE = regC + 1
            SetupBanks
        Case &H7FFE&
            Select4KVROM value, 0
        Case &H7FFF&
            Select4KVROM value, 1
        Case &H8000& To &HFFFF&
            reg8 = 4 * (value)
            regA = reg8 + 1
            regC = regA + 1
            regE = regC + 1
            SetupBanks
    End Select
End Sub
Public Sub map40_write(Address As Long, value As Byte)
    Select Case (Address And &HE000&)
        Case &H8000&
            Mapper40_IRQEnabled = 0
            Mapper40_IRQCounter = 36
        Case &HA000&
            ' IRQ enable
            Mapper40_IRQEnabled = 1
        Case &HE000&
            regC = value
            SetupBanks
    End Select
End Sub
Public Sub map66_write(Address As Long, value As Byte)
    reg8 = (value \ &H10) And &H3
    regA = reg8 + 1
    regC = reg8 + 2
    regE = reg8 + 3
    SetupBanks
    Select8KVROM value And &H3
End Sub
Public Sub map68_write(Address As Long, value As Byte)
    Select Case Address
        Case &H8000&: Select2KVROM value, 0
        Case &H9000&: Select2KVROM value, 1
        Case &HA000&: Select2KVROM value, 2
        Case &HB000&: Select2KVROM value, 3
        Case &HE000&: Mirroring = (value And &H3): DoMirror
        Case &HF000&: reg8 = (value * 2): regA = (value * 2) + 1: SetupBanks
    End Select
End Sub
Public Sub map69_write(Address As Long, value As Byte)
    Select Case Address
        Case &H8000&: reg8000 = value And &HF
        Case &HA000&
            Select Case reg8000
                Case 0 To 7: Select1KVROM value, reg8000
                Case 8: value = MaskBankAddress(value): MemCopy bank6(0), gameImage(value * &H2000&), &H2000&
                Case 9: reg8 = value
                Case 10: regA = value
                Case 11: regC = value
                Case 12: Mirroring = (value And &H3): DoMirror
                Case 13: ' Not sure here either
                Case 14, 15 ' Not sure about these.
            End Select
    End Select
End Sub
Public Sub map71_write(Address As Long, value As Byte)
    Select Case Address
        Case &H8000& To &HBFFF&: ' Unknown
        Case &HC000& To &HFFFF&: reg8 = (value * 2): regA = (value * 2) + 1: SetupBanks
    End Select
End Sub
Public Sub map78_write(Address As Long, value As Byte)
    
    Dim vromptr As Byte, romptr As Byte
    
    vromptr = (value \ &H10&) And &HF
    romptr = value And &HF
    
    Select8KVROM vromptr
    reg8 = (romptr * 2)
    regA = (romptr * 2) + 1
    SetupBanks
    
End Sub
Public Sub map91_write(Address As Long, value As Byte)
    Select Case Address
        Case &H6000& To &H6FFF&: Select2KVROM value, (Address And &H3)
        Case &H7000& To &H7FFB&
            Select Case Address And &H1
                Case 0
                    reg8 = value
                    SetupBanks
                Case 1
                    regA = value
                    SetupBanks
            End Select
    End Select
End Sub
Public Sub MapperWrite(ByVal Address As Long, ByVal value As Byte)
'===================================='
'       MapperWrite(Address,value)   '
' Selects/Switches Chr-ROM and Prg-  '
' ROM depending on the mapper. Based '
' on DarcNES.                        '
'===================================='
    Select Case Mapper
        Case 1: map1_write Address, value
        Case 2: map2_write Address, value
        Case 3: map3_write Address, value
        Case 4: map4_write Address, value
        Case 6: map6_write Address, value
        Case 7: map7_write Address, value
        Case 8: map66_write Address, value
        Case 9: map9_write Address, value
        Case 10: map9_write Address, value
        Case 11: map11_write Address, value
        Case 13: map13_write Address, value
        Case 15: map15_write Address, value
        Case 16: map16_write Address, value
        Case 17: map17_write Address, value
        Case 19: map19_write Address, value
        Case 22: map22_write Address, value
        Case 23: map23_write Address, value
        Case 24: map24_write Address, value
        Case 32: map32_write Address, value
        Case 33: map33_write Address, value
        Case 34: map34_write Address, value
        Case 40: map40_write Address, value
        Case 64: map64_write Address, value
        Case 65: map65_write Address, value
        Case 66: map66_write Address, value
        Case 68: map68_write Address, value
        Case 69: map69_write Address, value
        Case 71: map71_write Address, value
        Case 78: map78_write Address, value
        Case 90: map90_write Address, value
        Case 91: map91_write Address, value
        Case 118: map4_write Address, value
    End Select
End Sub
Public Sub map118_write(Address As Long, value As Byte)

End Sub
Public Sub map22_write(Address As Long, value As Byte)
' Konami VRC2 Type A
' This mapper was a breeze.
    Select Case Address
        Case &H8000&
            reg8 = value
            SetupBanks
        Case &H9000&
            Mirroring = (value And &H3)
        Case &HA000&
            regA = value
            SetupBanks
        Case &HB000&
            Select1KVROM value \ 2, 0
        Case &HB001&
            Select1KVROM value \ 2, 1
        Case &HC000&
            Select1KVROM value \ 2, 2
        Case &HC001&
            Select1KVROM value \ 2, 3
        Case &HD000&
            Select1KVROM value \ 2, 4
        Case &HD001&
            Select1KVROM value \ 2, 5
        Case &HE000&
            Select1KVROM value \ 2, 6
        Case &HE001&
            Select1KVROM value \ 2, 7
    End Select
End Sub
Public Function LoadNES(ByVal filename As String) As Byte
tmpLatch = 0
'===================================='
'           LoadNES(filename)        '
' Used to Load the NES ROM/VROM to   '
' specified arrays, gameImage and    '
' VROM, then figures out what to do  '
' based on the mapper number.        '
'===================================='
Dim Header As String * 3
Dim FileNum As Integer
FileNum = FreeFile

Dim i As Long
Dim ROMCtrl As Byte, ROMCtrl2 As Byte
Erase VRAM: Erase VROM: Erase gameImage: Erase bank8: Erase bankA
Erase bankC: Erase bankE: Erase bank0: Erase bank6
SpecialWrite6000 = False

If Dir$(filename) = "" Then MsgBox "File Not Found.", vbCritical, "basicNES": LoadNES = 0: Exit Function

Close #1

PrgCount = 0: PrgCount2 = 0: ChrCount = 0: ChrCount2 = 0

Open filename For Binary As #FileNum
    Get #FileNum, , Header
    If Header <> "NES" Then
        MsgBox "Invalid Header", vbCritical, "basicNES"
        LoadNES = 0
        Close #FileNum
        Exit Function
    End If
    
    Get #FileNum, 5, PrgCount: PrgCount2 = PrgCount
    Get #FileNum, 6, ChrCount: ChrCount2 = ChrCount
    Debug.Print "[ " & PrgCount & " ] ROM Bank(s)"
    Debug.Print "[ " & ChrCount & " ] CHR Bank(s)"
    
    Get #FileNum, 7, ROMCtrl
    Debug.Print "[ " & ROMCtrl & " ] ROM Control Byte #1"
    Get #FileNum, 8, ROMCtrl2
    Debug.Print "[ " & ROMCtrl2 & " ] ROM Control Byte #2"
    
    Mapper = (ROMCtrl And &HF0) \ 16
    Mapper = Mapper + ROMCtrl2
    
    If Mapper <> 0 And (ChrCount = 1 And PrgCount = 2) Then
        Mapper = 0
    End If
    If Mapper <> 0 And (ChrCount = 1 And PrgCount = 1) Then
        Mapper = 0
    End If
    Debug.Print "[ " & Mapper & " ] Mapper"
    Trainer = ROMCtrl And &H4
    Mirroring = ROMCtrl And &H1
    FourScreen = ROMCtrl And &H8
    If ROMCtrl And &H2 Then UsesSRAM = True
    Debug.Print "Mirroring=" & Mirroring & " Trainer=" & Trainer & " FourScreen=" & FourScreen & " SRAM=" & UsesSRAM
    Dim PrgMark As Long
    PrgMark = (PrgCount2 * &H4000&) - 1
    'Mirroring=1 is vertical
    If Trainer Then
        Get #FileNum, 17, Train
        'MemCopy bank6(&H1000&), Train(0), &H200
    End If
    
    ReDim gameImage(PrgMark) As Byte
    Dim startat As Integer
    If Trainer Then startat = 529 Else startat = 17
    Get #FileNum, startat, gameImage

    ReDim VROM(ChrCount2 * &H2000&) As Byte
    PrgMark = &H4000& * PrgCount2 + startat
    If ChrCount2 Then
        Get #FileNum, PrgMark, VROM
    End If
    MMC_Reset
    Select Case Mapper
        Case 0, 2, 3, 22, 23, 32, 33
            If ChrCount Then Select8KVROM 0
            reg8 = 0: regA = 1: regC = &HFE: regE = &HFF
            SetupBanks
        Case 1
            Select8KVROM 0
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
            sequence = 0: accumulator = 0
            Erase data
            data(0) = &H1F: data(3) = 0
        Case 4, 118
            swap = False
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            map4_sync
            MMC3_IrqVal = 0: MMC3_IrqOn = False: MMC3_TmpVal = 0
            If ChrCount Then
                Select8KVROM 0
            End If
        Case 5
            reg8 = &HFE
            regA = &HFE
            regC = &HFE
            regE = &HFE
            SetupBanks
        Case 6
            reg8 = &H0
            regA = &H1
            regC = &H7
            regE = &H8
            SetupBanks
            If ChrCount Then
                Select8KVROM 0
            End If
        Case 7
            reg8 = 0
            regA = 1
            regC = 2
            regE = 3
            SetupBanks
        Case 8
            reg8 = 0: regA = 1: regC = 2: regE = 3: SetupBanks
            Select8KVROM 0
            SetupBanks
        Case 9, 10 ' MMC2, MMC4
            reg8 = 0
            regA = &HFD
            regC = &HFE
            regE = &HFF
            SetupBanks
            latch1 = &HFE
            latch2 = &HFE
            Select8KVROM 0
        Case 11
            reg8 = 0: regA = 1: regC = 2: regE = 3
            SetupBanks
            Select8KVROM 0
        Case 13
            reg8 = 0: regA = 1: regC = 2: regE = 3: SetupBanks
            Select4KVROM 0, 0
            Select4KVROM 0, 1
            latch13 = 0
        Case 15
            reg8 = 0
            regA = 1
            regC = 2
            regE = 3
            SetupBanks
        Case 16
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
            MMC16_Irq = 0: MMC16_IrqOn = 0
            SpecialWrite6000 = True
        Case 17
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
            map17_irq = 0: map17_irqon = False
        Case 18
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
            map17_irq = 0: map17_irqon = False
        Case 19
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
            Select8KVROM ChrCount - 1
        Case 24
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
            map24_irqv = 0: map24_irqon = 0
        Case 34
            SpecialWrite6000 = True
            reg8 = 0
            regA = 1
            regC = 2
            regE = 3
            SetupBanks
            If ChrCount Then Select8KVROM 0
        Case 40 ' SMB2j [works!]
            Dim ChoseIt As Byte
            ChoseIt = 6
            ChoseIt = MaskBankAddress(ChoseIt)
            UsesSRAM = True
            MemCopy bank6(0), gameImage(ChoseIt * &H2000&), &H2000&
            reg8 = &HFC
            regA = &HFD
            regC = &HFE
            regE = &HFF
            SetupBanks
            Select8KVROM 0
            Mapper40_IRQEnabled = 0
            Mapper40_IRQCounter = 0
        Case 64
            Dim banks As Byte: banks = PrgCount * 2
            reg8 = &HFF: regA = reg8: regC = reg8: regE = reg8
            SetupBanks
            If ChrCount Then Select8KVROM 0
            cmd = 0: chr1 = 0: prg = 0
        Case 65
            reg8 = 0: regA = 1: regC = &HFE: regE = &HFF
            SetupBanks
        Case 66
            reg8 = 0: regA = 1: regC = 2: regE = 3: SetupBanks
            Select8KVROM 0
            SetupBanks
        Case 68
            reg8 = 0: regA = 1: regC = &HFE: regE = &HFF
            Select8KVROM 0
            SetupBanks
            UsesSRAM = True
        Case 69
            reg8 = 0
            regA = 1
            regC = 2
            regE = &HFF
            SetupBanks
            Select8KVROM 0
        Case 71
            reg8 = 0: regA = 1: regC = &HFE: regE = &HFF
            SetupBanks
        Case 78
            regC = &HFE
            regE = &HFF
            SetupBanks
            map78_write 0, 0
        Case 90, 160
            reg8 = &HFC
            regA = &HFD
            regC = &HFE
            regE = &HFF
            SetupBanks
        Case 91
            SpecialWrite6000 = True
            reg8 = &HFE
            regA = &HFF
            regC = &HFE
            regE = &HFF
            Select8KVROM 0
            SetupBanks
        Case Else
            MsgBox "Mapper #" & Mapper & " is not supported.", vbCritical, VERSION
            Erase gameImage: Erase VROM
            Close #FileNum: LoadNES = 0: Exit Function
    End Select
    Debug.Print "Successfully loaded " & filename
    reset6502
    If Mirroring = 1 Then MirrorXor = &H800& Else MirrorXor = &H400&
    If FourScreen Then Mirroring = 4
    DoMirror
Close #FileNum

CurrentLine = 0
For i = 0 To 7
    Joypad1(i) = &H40
Next i

FileNum = FreeFile
If UsesSRAM = True Then ' save the SRAM to a file.
    If Dir(App.Path & "\" & romName & ".wrm") <> "" Then
        Open App.Path & "\" & romName & ".wrm" For Binary Access Read As #FileNum
            Get #FileNum, , bank6
        Close #FileNum
    End If
End If

Frames = 0
CPUPaused = False
ScrollToggle = 1
frmNES.mnuFileRomInfo.Enabled = True
frmNES.mnuFileFree.Enabled = True
LoadNES = 1
End Function

Public Sub map17_write(Address As Long, value As Byte)
    Select Case Address
        Case &H42FE:
        Case &H42FF:
        Case &H4501: map17_irqon = (value And &H1)
        Case &H4502: map17_irq = &HFF00&: map17_irq = map17_irq Or value
        Case &H4503: map17_irq = &HFF: map17_irq = map17_irq Or value * &H100&: map17_irqon = 1
        Case &H4504: reg8 = value: SetupBanks
        Case &H4505: regA = value: SetupBanks
        Case &H4506: regC = value: SetupBanks
        Case &H4507: regE = value: SetupBanks
        Case &H4510 To &H4517: Select1KVROM value, (Address - &H4510)
    End Select
End Sub
Public Sub map17_doirq()
    If map17_irqon Then
        map17_irq = (map17_irq + 1)
        If map17_irq = &H10000 Then
            irq6502
            map17_irqon = False
            map17_irq = 0
        End If
    End If
End Sub
Public Sub map23_write(Address As Long, value As Byte)
    Select Case Address
        Case &H8000&
            reg8 = value
            SetupBanks
        Case &H9000&
             MirrorXor = pow2((value And &H3) + 10) '&H400& * pow2(value And &H3)
        Case &HA000&
            regA = value
            SetupBanks
        'TODO: finish code...
    End Select
End Sub

Public Sub map1_write(Address As Long, value As Byte)

Dim bank_select As Long 'Integer

If (value And &H80) Then
    data(0) = data(0) Or &HC
    accumulator = data((Address \ &H2000&) And 3)
    sequence = 5
Else
    If value And 1 Then accumulator = accumulator Or pow2(sequence)
    sequence = sequence + 1
End If

If (sequence = 5) Then
    data(Address \ &H2000& And 3) = accumulator
    sequence = 0
    accumulator = 0

    'MirrorXor = pow2((data(0) And &H3) + 10)
    
    If (PrgCount = &H20) Then '/* 512k cart */'
        bank_select = (data(1) And &H10) * 2
    Else '/* other carts */'
        bank_select = 0
    End If
    
    If data(0) And 2 Then 'enable panning
       Mirroring = (data(0) And 1) Xor 1
    Else 'disable panning
        Mirroring = 2
    End If
    DoMirror
    Select Case Mirroring
    Case 0
        MirrorXor = &H400
    Case 1
        MirrorXor = &H800
    Case 2
        MirrorXor = 0
    End Select
    
    If (data(0) And 8) = 0 Then 'base boot select $8000?
        reg8 = 4 * (data(3) And 15) + bank_select
        regA = 4 * (data(3) And 15) + bank_select + 1
        regC = 4 * (data(3) And 15) + bank_select + 2
        regE = 4 * (data(3) And 15) + bank_select + 3
        SetupBanks
    ElseIf (data(0) And 4) Then '16k banks
        reg8 = ((data(3) And 15) * 2) + bank_select
        regA = ((data(3) And 15) * 2) + bank_select + 1
        regC = &HFE
        regE = &HFF
        SetupBanks
    Else '32k banks
        reg8 = 0
        regA = 1
        regC = ((data(3) And 15) * 2) + bank_select
        regE = ((data(3) And 15) * 2) + bank_select + 1
        SetupBanks
    End If
    
    If (data(0) And &H10) Then '4k
        Select4KVROM data(1), 0
        Select4KVROM data(2), 1
    Else '8k
        Select8KVROM data(1) \ 2
    End If
End If

End Sub
Public Sub map2_write(Address As Long, value As Byte)
    reg8 = (value * 2)
    regA = reg8 + 1
    SetupBanks
End Sub
Public Sub map3_write(Address As Long, value As Byte)
    Select8KVROM value
End Sub
Public Sub map4_write(Address As Long, value As Byte)
Select Case Address
    Case &H8000&
        MMC3_Command = value And &H7
        If value And &H80 Then MMC3_ChrAddr = &H1000& Else MMC3_ChrAddr = 0
        If value And &H40 Then swap = 1 Else swap = 0
    Case &H8001&
        Select Case MMC3_Command
            Case 0: Select1KVROM value, 0: Select1KVROM value + 1, 1
            Case 1: Select1KVROM value, 2: Select1KVROM value + 1, 3
            Case 2: Select1KVROM value, 4
            Case 3: Select1KVROM value, 5
            Case 4: Select1KVROM value, 6
            Case 5: Select1KVROM value, 7
            Case 6: PrgSwitch1 = value: map4_sync
            Case 7: PrgSwitch2 = value: map4_sync
        End Select
    Case &HA000&
        If value And 1 Then Mirroring = 0 Else Mirroring = 1
        DoMirror
    Case &HA001&: If value Then UsesSRAM = True Else UsesSRAM = False
    Case &HC000&: MMC3_IrqVal = value
    Case &HC001&: MMC3_TmpVal = value
    Case &HE000&: MMC3_IrqOn = False: MMC3_IrqVal = MMC3_TmpVal
    Case &HE001&: MMC3_IrqOn = True
End Select
End Sub
Public Function map4_hblank(Scanline, two As Byte) As Boolean
    
    If Scanline = 0 Then
        MMC3_IrqVal = MMC3_TmpVal
    ElseIf Scanline > 239 Then
        Exit Function
    ElseIf MMC3_IrqOn And (two And &H18) Then
        MMC3_IrqVal = (MMC3_IrqVal - 1) And &HFF
        If (MMC3_IrqVal = 0) Then
            irq6502
            MMC3_IrqVal = MMC3_TmpVal
        End If
    End If
    
End Function
Public Sub map7_write(Address As Long, value As Byte)
    reg8 = 4 * (value And &HF)
    regA = reg8 + 1
    regC = reg8 + 2
    regE = reg8 + 3
    SetupBanks
    Mirroring = 2
    DoMirror
End Sub
Public Sub map9_write(Address As Long, value As Byte)
Static bnk As Long
Select Case (Address And &HF000&)
        Case &HA000&
            If Mapper = 9 Then
                reg8 = value
            ElseIf Mapper = 10 Then
                reg8 = value * 2
                regA = reg8 + 1
            End If
            SetupBanks
        Case &HB000&
            Latch0FD = value
            If latch1 = &HFD Then
                Select4KVROM value, 0
            End If
        Case &HC000&
            Latch0FE = value
            If latch1 = &HFE Then
                Select4KVROM value, 0
            End If
        Case &HD000&
            Latch1FD = value
            If latch2 = &HFD Then
                Select4KVROM value, 1
            End If
        Case &HE000&
            Latch1FE = value
            If latch2 = &HFE Then
                Select4KVROM value, 1
            End If
        Case &HF000&
            If (value And 1) Then
                Mirroring = 0
            ElseIf (value And 1) = 0 Then
                Mirroring = 1
            End If
End Select
End Sub

Public Sub map9_latch(TileNum As Byte, Hi As Boolean)
If Mapper <> 9 Then Exit Sub

If (TileNum = &HFD) Then
    If (Hi = False) Then
        Select4KVROM Latch0FD, 0
        latch1 = &HFD
    ElseIf (Hi = True) Then
        Select4KVROM Latch1FD, 1
        latch2 = &HFD
    End If
ElseIf (TileNum = &HFE) Then
    If (Hi = False) Then
        Select4KVROM Latch0FE, 0
        latch1 = &HFE
    ElseIf (Hi = True) Then
        Select4KVROM Latch1FE, 1
        latch2 = &HFE
    End If
End If

End Sub
'Tile based sprite renderer
Public Sub DrawSprites(ontop As Boolean)
If (PPU_Control2 And 16) = 0 Then Exit Sub
    Dim SpritePattern As Long 'Integer
    SpritePattern = (PPU_Control1 And &H8) * &H200&
    Dim spr As Long 'Integer
    Dim X1 As Long, y1 As Long
    Dim Byte1 As Byte, Byte2 As Byte
    Dim Color As Byte
    Dim sa As Long
    Dim h As Long
    Dim i As Long
    Dim X As Long, Y As Long, attrib As Long, tileno As Long, pal As Long, Aa As Long
    
    If PPU_Control1 And &H20 Then
        h = 16
    Else
        h = 8
    End If
    
    SpriteAddr = 0
    
    For spr = 63 To 0 Step -1
        SpriteAddr = 4 * spr
        attrib = SpriteRAM(SpriteAddr + 2)
        If (attrib And 32) = 0 Xor ontop Then
            X = SpriteRAM(SpriteAddr + 3)
            Y = SpriteRAM(SpriteAddr)
            If Y < 239 And X < 248 Then
                tileno = SpriteRAM(SpriteAddr + 1)
                If h = 16 Then
                    SpritePattern = (tileno And 1) * &H1000
                    tileno = tileno Xor (tileno And 1)
                End If
                sa = SpritePattern + 16 * tileno
                i = Y * 256& + X + 256
                pal = 16 + (attrib And 3) * 4
                
                If attrib And 128 Then
                    If attrib And 64 Then
                        For y1 = h - 1 To 0 Step -1
                            If y1 >= 8 Then
                                Byte1 = VRAM(sa + 8 + y1)
                                Byte2 = VRAM(sa + 16 + y1)
                            Else
                                Byte1 = VRAM(sa + y1)
                                Byte2 = VRAM(sa + y1 + 8)
                            End If
                            Aa = Byte1 * 2048& + Byte2 * 8
                            For X1 = 0 To 7
                                Color = tLook(X1 + Aa)
                                If Color Then vBuffer(i + X1) = Color Or pal
                            Next X1
                            i = i + 256
                            If i >= 256& * 240 Then Exit For
                        Next y1
                    Else
                        i = i + 7
                        For y1 = h - 1 To 0 Step -1
                            If y1 >= 8 Then
                                Byte1 = VRAM(sa + 8 + y1)
                                Byte2 = VRAM(sa + 16 + y1)
                            Else
                                Byte1 = VRAM(sa + y1)
                                Byte2 = VRAM(sa + y1 + 8)
                            End If
                            Aa = Byte1 * 2048& + Byte2 * 8
                            For X1 = 7 To 0 Step -1
                                Color = tLook(X1 + Aa)
                                If Color Then vBuffer(i - X1) = Color Or pal
                            Next X1
                            i = i + 256
                            If i >= 256& * 240 Then Exit For
                        Next y1
                    End If
                Else
                    If attrib And 64 Then
                        For y1 = 0 To h - 1
                            If y1 = 8 Then sa = sa + 8
                            Byte1 = VRAM(sa + y1)
                            Byte2 = VRAM(sa + y1 + 8)
                            Aa = Byte1 * 2048& + Byte2 * 8
                            For X1 = 0 To 7
                                Color = tLook(X1 + Aa)
                                If Color Then vBuffer(i + X1) = Color Or pal
                            Next X1
                            i = i + 256
                            If i >= 256& * 240 Then Exit For
                        Next y1
                    Else
                        i = i + 7
                        For y1 = 0 To h - 1
                            If y1 = 8 Then sa = sa + 8
                            Byte1 = VRAM(sa + y1)
                            Byte2 = VRAM(sa + y1 + 8)
                            Aa = Byte1 * 2048& + Byte2 * 8
                            For X1 = 7 To 0 Step -1
                                Color = tLook(X1 + Aa)
                                If Color Then vBuffer(i - X1) = Color Or pal
                            Next X1
                            i = i + 256
                            If i >= 256& * 240 Then Exit For
                        Next y1
                    End If
                End If
            End If
        End If
    Next spr
End Sub
Public Sub LoadPal(file As String)
On Error Resume Next
Dim n As Long 'Integer
Dim r As Byte, g As Byte, b As Byte
Dim FileNum As Integer
FileNum = FreeFile

Open App.Path & "\" + file For Binary As #FileNum

For n = 0 To 63
    Get #FileNum, , r
    Get #FileNum, , g
    Get #FileNum, , b
    pal(n) = RGB2(r, g, b)
    pal16(n) = rgb16(r, g, b)
    pal15(n) = rgb15(r, g, b)
    pal(n + 64) = pal(n) 'DF: saves us an "and" per pixel
    pal(n + 128) = pal(n)
    pal(n + 192) = pal(n)
    pal16(n + 64) = pal16(n)
    pal16(n + 128) = pal16(n)
    pal16(n + 192) = pal16(n)
    pal15(n + 64) = pal15(n)
    pal15(n + 128) = pal15(n)
    pal15(n + 192) = pal15(n)
Next n
Close #FileNum
End Sub

'DF: had to reverse colors in palette for it to look right with new gfx.
'later changed for 16bit color, then back to 32bit
Function RGB2(ByVal b As Long, ByVal g As Long, ByVal r As Long) As Long
RGB2 = RGB(r, g, b)
End Function
Function rgb16(ByVal b As Long, ByVal g As Long, ByVal r As Long) As Long
    Dim c As Long
    c = (r \ 8) + (g \ 4) * 32 + (b \ 8) * 2048
    If c > 32767 Then c = c - 65536
    rgb16 = c
End Function
Function rgb15(ByVal b As Long, ByVal g As Long, ByVal r As Long) As Long
    rgb15 = (r \ 8) + (g \ 8) * 32 + (b \ 8) * 1024
End Function

