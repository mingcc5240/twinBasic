Attribute VB_Name = "MMC"
Option Explicit

DefLng A-Z

' Functions for emulating MMCs. Select8KVROM and the
' like
' 16.07.00
Public CurrVr As Byte
Public PrgSwitch1 As Byte
Public PrgSwitch2 As Byte
Public SpecialWrite6000 As Boolean

Public bank0(2047) As Byte ' RAM
Public bank6(8191) As Byte ' SaveRAM
Public bank8(8191) As Byte '8-E are PRG-ROM.
Public bankA(8191) As Byte
Public bankC(8191) As Byte
Public bankE(8191) As Byte

Public MapperNames As Variant

Private p8, pA, rC, pE 'addresses of prg-rom banks currently selected.

Private prevBSSrc(7) As Long 'used to ensure that it doesn't bankswitch when the correct bank is already selected


Private allowXor As Boolean

Private Sub CopyBanks(dest, src, count)
    On Error Resume Next
    If Mapper = 4 And allowXor Then
        Dim i
        For i = 0 To count - 1
            MemCopy VRAM(MMC3_ChrAddr Xor (dest + i) * &H400), VROM((src + i) * &H400), &H400
        Next i
    Else
        MemCopy VRAM(dest * &H400), VROM(src * &H400), count * &H400
    End If
End Sub

'doesn't bankswitch when not needed
Private Sub BankSwitch(ByVal dest, ByVal src, ByVal count)
    Dim Aa, b, c
    Aa = 0
    c = 0
    allowXor = count <= 2 'only xor with MMC3_ChrAddr with banks of 1 or 2k
    For b = 0 To count - 1
        If prevBSSrc(dest + b) <> src + b Then
            c = c + 1 'we copy banks in groups, not 1 at a time. a little faster.
            prevBSSrc(dest + b) = src + b
        Else
            If c > 0 Then CopyBanks dest + Aa, src + Aa, c
            Aa = b + 1
            c = 0
        End If
    Next b
    If c > 0 Then CopyBanks dest + Aa, src + Aa, c
End Sub



'resets the info used to decide if a bankswitch is needed.
Public Sub MMC_Reset()
    p8 = -1
    pA = -1
    rC = -1
    pE = -1
    Dim i As Long
    For i = 0 To 7
        prevBSSrc(i) = -1
    Next i
End Sub


Public Sub map4_sync()
If swap Then
    reg8 = &HFE
    regA = PrgSwitch2
    regC = PrgSwitch1
    regE = &HFF
Else
    reg8 = PrgSwitch1
    regA = PrgSwitch2
    regC = &HFE
    regE = &HFF
End If
SetupBanks
End Sub

Public Function MaskBankAddress(bank As Byte)
If bank >= PrgCount * 2 Then
    Dim i As Byte: i = &HFF
    Do While (bank And i) >= PrgCount * 2
        i = i \ 2
    Loop
    MaskBankAddress = (bank And i)
Else
    MaskBankAddress = bank
End If
End Function
Public Function MaskVROM(page As Byte, ByVal mask As Long) As Byte
    Dim i As Long
    If mask = 0 Then mask = 256
    If mask And mask - 1 Then 'if mask is not a power of 2
        i = 1
        Do While i < mask 'find smallest power of 2 >= mask
            i = i + i
        Loop
    Else
        i = mask
    End If
    i = (page And (i - 1))
    If i >= mask Then i = mask - 1
    MaskVROM = i
End Function

'only switches banks when needed
Public Sub SetupBanks()
    reg8 = MaskBankAddress(reg8)
    regA = MaskBankAddress(regA)
    regC = MaskBankAddress(regC)
    regE = MaskBankAddress(regE)
    
    If p8 <> reg8 Then MemCopy bank8(0), gameImage(reg8 * &H2000&), &H2000&
    If pA <> regA Then MemCopy bankA(0), gameImage(regA * &H2000&), &H2000&
    If rC <> regC Then MemCopy bankC(0), gameImage(regC * &H2000&), &H2000&
    If pE <> regE Then MemCopy bankE(0), gameImage(regE * &H2000&), &H2000&
    p8 = reg8
    pA = regA
    rC = regC
    pE = regE
End Sub

Public Sub Select8KVROM(val1 As Byte)
    val1 = MaskVROM(val1, ChrCount)
    BankSwitch 0, val1 * 8, 8
End Sub
Public Sub Select4KVROM(val1 As Byte, bank As Byte)
    val1 = MaskVROM(val1, ChrCount * 2)
    BankSwitch bank * 4, val1 * 4, 4
End Sub
Public Sub Select2KVROM(val1 As Byte, bank As Byte)
    val1 = MaskVROM(val1, ChrCount * 4)
    BankSwitch bank * 2, val1 * 2, 2
End Sub
Public Sub Select1KVROM(val1 As Byte, bank As Byte)
    val1 = MaskVROM(val1, ChrCount * 8)
    BankSwitch bank, val1, 1
End Sub
