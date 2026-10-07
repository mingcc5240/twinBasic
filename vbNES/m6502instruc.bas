Attribute VB_Name = "m6502instruct"
Option Explicit

DefLng A-Z


' This is where all 6502 instructions are kept.
Public Sub adc6502()
    Dim tmp As Long ' Integer
      
    adrmode opcode
    value = Read6502(savepc)
     
    saveflags = (P And &H1)

    sum = a
    sum = (sum + value) And &HFF
    sum = (sum + saveflags) And &HFF
      
    If (sum > &H7F) Or (sum < -&H80) Then
        P = P Or &H40
    Else
        P = (P And &HBF)
    End If
      
    sum = a + (value + saveflags)
    If (sum > &HFF) Then
        P = P Or &H1
    Else
        P = (P And &HFE)
    End If
      
    a = sum And &HFF
    If (P And &H8) Then
        P = (P And &HFE)
        If ((a And &HF) > &H9) Then
            a = (a + &H6) And &HFF
        End If
        If ((a And &HF0) > &H90) Then
            a = (a + &H60) And &HFF
            P = P Or &H1
        End If
    Else
        clockticks6502 = clockticks6502 + 1
    End If
    SetFlags a
End Sub

Public Sub adrmode(opcode As Byte)
Select Case addrmode(opcode)
    Case ADR_ABS: savepc = Read6502(PC) + (Read6502(PC + 1) * &H100&): PC = PC + 2
    Case ADR_ABSX: absx6502
    Case ADR_ABSY: absy6502
    Case ADR_IMP: ' nothing really necessary cause implied6502 = ""
    Case ADR_IMM: savepc = PC: PC = PC + 1
    Case ADR_INDABSX: indabsx6502
    Case ADR_IND: indirect6502
    Case ADR_INDX: indx6502
    Case ADR_INDY: indy6502
    Case ADR_INDZP: indzp6502
    Case ADR_REL: savepc = Read6502(PC): PC = PC + 1: If (savepc And &H80) Then savepc = savepc - &H100&
    Case ADR_ZP: savepc = Read6502(PC): savepc = savepc And &HFF: PC = PC + 1
    Case ADR_ZPX: zpx6502
    Case ADR_ZPY: zpy6502
    Case Else: Debug.Print addrmode(opcode)
End Select
End Sub

Public Sub and6502()
  adrmode opcode
  value = Read6502(savepc)
  a = (a And value)
  SetFlags a
End Sub
Public Sub asl6502()
  adrmode opcode
  value = Read6502(savepc)
  
  P = (P And &HFE) Or ((value \ 128) And &H1)
  value = (value * 2) And &HFF
  
  Write6502 savepc, (value And &HFF)
  SetFlags value
End Sub


Public Sub asla6502()
  P = (P And &HFE) Or ((a \ 128) And &H1)
  a = (a * 2) And &HFF
  SetFlags a
End Sub

Public Sub bcc6502()
  If ((P And &H1) = 0) Then
    adrmode opcode
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    PC = PC + 1
  End If
End Sub

Public Sub bcs6502()
  If (P And &H1) Then
    adrmode opcode
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    PC = PC + 1
  End If
End Sub

Public Sub beq6502()
  If (P And &H2) Then
    adrmode opcode
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    PC = PC + 1
  End If
End Sub

Public Sub bit6502()
    adrmode opcode
    value = Read6502(savepc)
  
    If (value And a) Then
        P = (P And &HFD)
    Else
        P = P Or &H2
    End If
    P = ((P And &H3F) Or (value And &HC0))
End Sub

Public Sub bmi6502()
    If (P And &H80) Then
        adrmode opcode
        PC = PC + savepc
        clockticks6502 = clockticks6502 + 1
    Else
        PC = PC + 1
    End If
End Sub

Public Sub bne6502()
    If ((P And &H2) = 0) Then
        adrmode opcode
        PC = PC + savepc
    Else
        PC = PC + 1
    End If
End Sub

Public Sub bpl6502()
    If ((P And &H80) = 0) Then
        adrmode opcode
        PC = PC + savepc
    Else
        PC = PC + 1
    End If
End Sub

Public Sub brk6502()
    PC = PC + 1
    Write6502 &H100& + s, (PC \ &H100&) And &HFF
    s = (s - 1) And &HFF
    Write6502 &H100& + s, (PC And &HFF)
    s = (s - 1) And &HFF
    Write6502 &H100& + s, P
    s = (s - 1) And &HFF
    P = P Or &H14
    PC = Read6502(&HFFFE&) + (Read6502(&HFFFF&) * &H100&)
End Sub

Public Sub bvc6502()
    If ((P And &H40) = 0) Then
        adrmode opcode
        PC = PC + savepc
        clockticks6502 = clockticks6502 + 1
    Else
        PC = PC + 1
    End If
End Sub

Public Sub bvs6502()
  If (P And &H40) Then
    adrmode opcode
    PC = PC + savepc
    clockticks6502 = clockticks6502 + 1
  Else
    PC = PC + 1
  End If
End Sub

Public Sub clc6502()
  P = P And &HFE
End Sub

Public Sub cld6502()
  P = P And &HF7
End Sub

Public Sub cli6502()
  P = P And &HFB
End Sub

Public Sub clv6502()
  P = P And &HBF
End Sub

Public Sub cmp6502()
  adrmode opcode
  value = Read6502(savepc)
  
  If (a + &H100 - value) > &HFF Then
    P = P Or &H1
  Else
    P = (P And &HFE)
  End If
  
  value = (a + &H100 - value) And &HFF
  SetFlags value
End Sub

Public Sub cpx6502()
  adrmode opcode
  value = Read6502(savepc)
      
  If (X + &H100 - value > &HFF) Then
    P = P Or &H1
  Else
    P = (P And &HFE)
  End If
  
  value = (X + &H100 - value) And &HFF
  SetFlags value
End Sub

Public Sub cpy6502()
  adrmode opcode
  value = Read6502(savepc)
      
  If (Y + &H100 - value > &HFF) Then
    P = (P Or &H1)
  Else
    P = (P And &HFE)
  End If
  value = (Y + &H100 - value) And &HFF
  SetFlags value
End Sub

Public Sub dec6502()
  adrmode opcode
  Write6502 (savepc), (Read6502(savepc) - 1) And &HFF
  value = Read6502(savepc)
  If (value) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (value And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub dex6502()
  X = (X - 1) And &HFF
  If (X) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (X And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub

Public Sub dey6502()
  Y = (Y - 1) And &HFF
  If (Y) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (Y And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub

Public Sub eor6502()
  adrmode opcode
  a = a Xor Read6502(savepc)
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub

Public Sub inc6502()
  adrmode opcode
  Write6502 (savepc), (Read6502(savepc) + 1) And &HFF
  value = Read6502(savepc)
  If (value) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (value And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub

Public Sub inx6502()
  X = (X + 1) And &HFF
  If (X) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (X And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub

Public Sub iny6502()
  Y = (Y + 1) And &HFF
  If (Y) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (Y And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub jmp6502()
  adrmode opcode
  PC = savepc
End Sub
Public Sub jsr6502()
  PC = PC + 1
  Write6502 s + &H100&, (PC \ &H100&)
  s = (s - 1) And &HFF
  Write6502 s + &H100&, (PC And &HFF)
  s = (s - 1) And &HFF
  PC = PC - 1
  adrmode opcode
  PC = savepc
End Sub
Public Sub lda6502()
  adrmode opcode
  
  a = Read6502(savepc)
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub ldx6502()
  adrmode opcode
  X = Read6502(savepc)
  If (X) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (X And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub ldy6502()
  adrmode opcode
  Y = Read6502(savepc)
  If (Y) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (Y And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub lsr6502()
  adrmode opcode
  value = Read6502(savepc)
         
  P = ((P And &HFE) Or (value And &H1))
  
  value = (value \ 2) And &HFF
  Write6502 savepc, (value And &HFF)
  If (value) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (value And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub lsra6502()
  P = (P And &HFE) Or (a And &H1)
  a = (a \ 2) And &HFF
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub nop6502()
'TS: Implemented complex code structure ;)
End Sub
Public Sub ora6502()
  adrmode opcode
  a = a Or Read6502(savepc)
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub pha6502()
  Write6502 &H100& + s, a
  s = (s - 1) And &HFF
End Sub
Public Sub php6502()
  Write6502 &H100& + s, P
  s = (s - 1) And &HFF
End Sub
Public Sub pla6502()
  s = (s + 1) And &HFF
  a = Read6502(s + &H100)
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub plp6502()
  s = (s + 1) And &HFF
  P = Read6502(s + &H100) Or &H20
End Sub
Public Sub rol6502()
  saveflags = (P And &H1)
  adrmode opcode
  value = Read6502(savepc)
      
  P = (P And &HFE) Or ((value \ 128) And &H1)
  
  value = (value * 2) And &HFF
  value = value Or saveflags
  
  Write6502 savepc, (value And &HFF)
  
  If (value) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (value And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub rola6502()
  saveflags = (P And &H1)
  P = (P And &HFE) Or ((a \ 128) And &H1)
  a = (a * 2) And &HFF
  a = a Or saveflags
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub ror6502()
  saveflags = (P And &H1)
  adrmode opcode
  value = Read6502(savepc)
      
  P = (P And &HFE) Or (value And &H1)
  value = (value \ 2) And &HFF
  If (saveflags) Then
    value = value Or &H80
  End If
  Write6502 (savepc), value And &HFF
  If (value) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (value And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub rora6502()
  saveflags = (P And &H1)
  P = (P And &HFE) Or (a And &H1)
  a = (a \ 2) And &HFF
  
  If (saveflags) Then
    a = a Or &H80
  End If
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub rti6502()
  s = (s + 1) And &HFF
  P = Read6502(s + &H100&) Or &H20
  s = (s + 1) And &HFF
  PC = Read6502(s + &H100&)
  s = (s + 1) And &HFF
  PC = PC + (Read6502(s + &H100) * &H100&)
End Sub

Public Sub rts6502()
  s = (s + 1) And &HFF
  PC = Read6502(s + &H100)
  s = (s + 1) And &HFF
  PC = PC + (Read6502(s + &H100) * &H100&)
  PC = PC + 1
End Sub

Public Sub sbc6502()
  adrmode opcode
  value = Read6502(savepc) Xor &HFF
  
  saveflags = (P And &H1)
  
  sum = a
  sum = (sum + value) And &HFF
  sum = (sum + (saveflags * 16)) And &HFF
  
  If ((sum > &H7F) Or (sum <= -&H80)) Then
    P = P Or &H40
  Else
    P = P And &HBF
  End If
  
  sum = a + (value + saveflags)
  
  If (sum > &HFF) Then
    P = P Or &H1
  Else
    P = P And &HFE
  End If
  
  a = sum And &HFF
  If (P And &H8) Then
        a = (a - &H66) And &HFF
        P = P And &HFE
    If ((a And &HF) > &H9) Then
      a = (a + &H6) And &HFF
    End If
    If ((a And &HF0) > &H90) Then
      a = (a + &H60) And &HFF
      P = P Or &H1
    End If
  Else
    clockticks6502 = clockticks6502 + 1
  End If
  'Debug.Print "sbc6502"
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub sec6502()
  P = P Or &H1
End Sub
Public Sub sed6502()
  P = P Or &H8
End Sub
Public Sub sei6502()
  P = P Or &H4
End Sub
Public Sub sta6502()
  adrmode opcode
  Write6502 (savepc), a
End Sub
Public Sub stx6502()
  adrmode opcode
  Write6502 (savepc), X
End Sub
Public Sub sty6502()
  adrmode opcode
  Write6502 (savepc), Y
End Sub
Public Sub tax6502()
  X = a
  If (X) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (X And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub

Public Sub tay6502()
  Y = a
  If (Y) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (Y And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub tsx6502()
  X = s
  If (X) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (X And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub

Public Sub txa6502()
  a = X
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub txs6502()
  s = X
End Sub

Public Sub tya6502()
  a = Y
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub bra6502()
  adrmode opcode
  PC = PC + savepc
  clockticks6502 = clockticks6502 + 1
End Sub
Public Sub dea6502()
  a = (a - 1) And &HFF
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub ina6502()
  a = (a + 1) And &HFF
  If (a) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (a And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub phx6502()
  Write6502 &H100 + s, X
  s = (s - 1) And &HFF
End Sub

Public Sub plx6502()
  s = (s + 1) And &HFF
  X = Read6502(s + &H100)
  If (X) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (X And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
Public Sub phy6502()
  Write6502 &H100 + s, Y
  s = (s - 1) And &HFF
End Sub
Public Sub ply6502()
  s = (s + 1) And &HFF
  
  Y = Read6502(s + &H100)
  If (Y) Then
    P = P And &HFD
  Else
    P = P Or &H2
  End If
  If (Y And &H80) Then
    P = P Or &H80
  Else
    P = P And &H7F
  End If
End Sub
