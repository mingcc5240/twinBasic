Attribute VB_Name = "M6502AddrModes"
Option Explicit

DefLng A-Z


' M6502 CPU Implementation for basicNES 2000.
' By Don Jarrett and Tobias Strömstedt, 1997-2003.
' If you use this file commercially please drop me a mail
' at r.jarrett@worldnet.att.net or d.jarrett@worldnet.att.net.
' basicNES Copyright (C) 1996-2002 Don Jarrett.

Public CurrentLine As Long 'Integer

'Registers and tempregisters
'DF: Be careful. Anything, anywhere that uses a variable of the same name without declaring it will be using these:
Public A As Byte
Public X As Byte
Public Y As Byte
Public S As Byte
Public P As Byte

'32bit instructions are faster in protected mode than 16bit
Public PC As Long
Public savepc As Long
Public value As Long 'Integer
Public value2 As Long 'Integer
Public sum As Long 'Integer
Public saveflags As Long 'Integer
Public help As Long

Public opcode As Byte
Public clockticks6502 As Long

'Private opcount(255) As Long 'comment if not debugging

' arrays
Public Ticks(0 To &H100&) As Byte
Public addrmode(0 To &H100&) As Byte
Public instruction(0 To &H100&) As Byte
Public gameImage() As Byte

Public maxIdle As Long

Public CPUPaused As Boolean

Public addrmodeBase As Long

Public maxCycles1 As Long 'max cycles per scanline from scanlines 0-239
Public maxCycles As Long 'max cycles until next scanline
Public SmartExec As Boolean
Public realframes As Long 'actual # of frames rendered

Public IdleDetect As Boolean
'Private IdleAddr As Long
Public idleCheck(65535) As Byte

Public autospeed As Boolean

Public KeyCodes(7) As Long

Public rCycles As Long 'real number of cycles executed
Public nCycles As Long 'number that should be executed
Public SaveCPU As Boolean ' Pause CPU when basicNES loses focus?

Public Const M6502_INTERNAL_REVISION = "v.18"
Public Sub implied6502()

End Sub
Public Sub indabsx6502()
  help = Read6502(PC) + (Read6502(PC + 1) * &H100&) + X
  savepc = Read6502(help) + (Read6502(help + 1) * &H100&)
End Sub
Public Sub indx6502()
'TS: Changed PC++ and removed ' (?)
  value = Read6502(PC) And &HFF
  value = (value + X) And &HFF
  PC = PC + 1
  savepc = Read6502(value) + (Read6502(value + 1) * &H100&)
End Sub
Public Sub indy6502()
'TS: Changed PC++ and == to != (If then else)
  value = Read6502(PC)
  PC = PC + 1
      
  savepc = Read6502(value) + (Read6502(value + 1) * &H100&)
  If (Ticks(opcode) = 5) Then
    If ((savepc \ &H100&) = ((savepc + Y) \ &H100&)) Then
    Else
      clockticks6502 = clockticks6502 + 1
    End If
  End If
  savepc = savepc + Y
End Sub

Public Sub zpx6502()
'TS: Rewrote everything!
'Overflow stupid check
  savepc = Read6502(PC)
  savepc = savepc + X
  PC = PC + 1
  savepc = savepc And &HFF
End Sub
Public Sub exec6502()

#If False Then 'used in deciding which opcodes to optimize.
    Static fc As Long
    fc = fc + 1
    If fc = 1000 Then
        Open "c:\perf.txt" For Output As #1
        Dim u As Long, mu As Long
        mu = 0
        Do
        For u = 0 To 255
            If opcount(u) > opcount(mu) Then mu = u
        Next u
        If opcount(mu) > 0 Then
            Print #1, mu&; "  " & opcount(mu)
            opcount(mu) = 0
        Else
            Exit Do
        End If
        Loop Until False
        Close #1
        MsgBox "Done"
    End If
#End If
    


  Dim f As Long
  f = Frames
  While CPUPaused
        DoEvents
  Wend
  While Frames = f And CPURunning
  opcode = Read6502(PC)  ' Fetch Next Operation
  PC = PC + 1
  If IdleDetect Then
    If idleCheck(PC) > 8 Then
        If CurrentLine > 240 Or CurrentLine < maxIdle Then
            idleCheck(PC) = idleCheck(PC) - 8
        Else
            clockticks6502 = clockticks6502 + CurrentLine \ 2: rCycles = rCycles - CurrentLine \ 2
        End If
    End If
    If CurrentLine >= 231 And CurrentLine < 238 And idleCheck(PC) < 240 Then
        idleCheck(PC) = idleCheck(PC) + 1
    End If
  End If

  clockticks6502 = clockticks6502 + Ticks(opcode)

'opcount(instruction(opcode)) = opcount(instruction(opcode)) + 1 'comment if not debugging



Select Case instruction(opcode)
    Case INS_JMP: ' jmp6502
        adrmode opcode
        PC = savepc
    Case INS_LDA: ' lda6502
        adrmode opcode
        A = Read6502(savepc)
        SetFlags A
    Case INS_LDX:
        adrmode (opcode)
        X = Read6502(savepc)
        SetFlags X
    Case INS_LDY
        adrmode (opcode)
        Y = Read6502(savepc)
        SetFlags Y
    Case INS_BNE: bne6502
    Case INS_CMP: cmp6502
    Case INS_STA
        adrmode (opcode)
        Write6502 savepc, A
    Case INS_BIT: bit6502
    Case INS_BVC: bvc6502
    Case INS_BEQ: beq6502
    Case INS_INY: iny6502
    Case INS_BPL: bpl6502
    Case INS_DEX: dex6502
    Case INS_INC: inc6502
    Case INS_DEC: dec6502
    Case INS_JSR: jsr6502
    Case INS_AND: and6502
    Case INS_NOP:
    
    Case INS_BRK: brk6502
    Case INS_ADC: adc6502
    Case INS_EOR: eor6502
    Case INS_ASL: asl6502
    Case INS_ASLA: asla6502
    Case INS_BCC: bcc6502
    Case INS_BCS: bcs6502
    Case INS_BMI: bmi6502
    Case INS_BVS: bvs6502
    Case INS_CLC: P = P And &HFE
    Case INS_CLD: P = P And &HF7
    Case INS_CLI: P = P And &HFB
    Case INS_CLV: P = P And &HBF
    Case INS_CPX: cpx6502
    Case INS_CPY: cpy6502
    Case INS_DEA: dea6502
    Case INS_DEY: dey6502
    Case INS_INA: ina6502
    Case INS_INX: inx6502
    Case INS_LSR: lsr6502
    Case INS_LSRA: lsra6502
    Case INS_ORA
        adrmode opcode
        A = A Or Read6502(savepc)
        SetFlags A
    Case INS_PHA: pha6502
    Case INS_PHX: phx6502
    Case INS_PHP: php6502
    Case INS_PHY: phy6502
    Case INS_PLA: pla6502
    Case INS_PLP: plp6502
    Case INS_PLX: plx6502
    Case INS_PLY: ply6502
    Case INS_ROL: rol6502
    Case INS_ROLA: rola6502
    Case INS_ROR: ror6502
    Case INS_RORA: rora6502
    Case INS_RTI: rti6502
    Case INS_RTS: rts6502
    Case INS_SBC: sbc6502
    Case INS_SEC: P = P Or &H1
    Case INS_SED: P = P Or &H8
    Case INS_SEI: P = P Or &H4
    Case INS_STX
        adrmode (opcode)
        Write6502 savepc, X
    Case INS_STY
        adrmode (opcode)
        Write6502 savepc, Y
    Case INS_TAX: tax6502
    Case INS_TAY: tay6502
    Case INS_TXA: txa6502
    Case INS_TYA: tya6502
    Case INS_TXS: txs6502
    Case INS_TSX: tsx6502
    Case INS_BRA: bra6502
    Case Else: MsgBox "Invalid opcode - " & Hex$(opcode)
End Select
  
  If clockticks6502 > maxCycles Then
        nCycles = nCycles + 114
        rCycles = rCycles + maxCycles
        If Mapper = 4 Then
            map4_hblank CurrentLine, PPU_Control2
        ElseIf Mapper = 6 Then
            map6_hblank CurrentLine
        ElseIf Mapper = 16 Then
            If (MMC16_IrqOn <> 0) Then
                MMC16_Irq = MMC16_Irq - 1
                If (MMC16_Irq = 0) Then
                    irq6502
                    MMC16_IrqOn = 0
                End If
            End If
        ElseIf Mapper = 17 Then
            map17_doirq
        ElseIf Mapper = 19 Then
            map19_irq
        ElseIf Mapper = 24 Then
            map24_irq
        End If
        RenderScanline CurrentLine
        If CurrentLine >= 240 Then
            If CurrentLine = 240 Then
                If render Then
                    blitScreen
                End If
                Frames = Frames + 1
                If render Then realframes = realframes + 1
                DoEvents 'ensure most recent keyboard input
                
                'hold down 'a' key for autofire. now 15hz, not 30
                
                If Keyboard(65) <> &H41 Then
                    Joypad1(0) = Keyboard(nes_ButA) 'A
                    Joypad1(1) = Keyboard(nes_ButB) 'B
                ElseIf Frames And 2 Then
                    Joypad1(0) = Keyboard(nes_ButA) 'A
                    Joypad1(1) = &H40
                Else
                    Joypad1(0) = &H40
                    Joypad1(1) = Keyboard(nes_ButB) 'B
                End If
                
                
                Joypad1(2) = Keyboard(nes_ButSel) 'Sl
                Joypad1(3) = Keyboard(nes_ButSta)
                
                Joypad1(4) = Keyboard(nes_ButUp)
                Joypad1(5) = Keyboard(nes_ButDn)
                Joypad1(6) = Keyboard(nes_ButLt)
                Joypad1(7) = Keyboard(nes_ButRt)
                If Record = True Then
                    Put #1, , Joypad1
                ElseIf Playing = True Then
                    Get #1, , Joypad1
                    If EOF(1) Then
                        StopPlaying
                        frmNES.mnuPlayMovie.Caption = "&Play"
                        frmNES.lblStatus.Caption = "Stopped playing."
                        Playing = False
                    End If
                End If
            End If
            PPU_Status = &H80
            
            If CurrentLine = 240 Then
                If (PPU_Control1 And &H80) Then
                    If IdleDetect Then idleCheck(PC) = 1
                    nmi6502
                End If
            End If
            If Mapper = 40 Then
                If Mapper40_IRQEnabled = 1 Then
                    Mapper40_IRQCounter = (Mapper40_IRQCounter - 1) And 36
                    If Mapper40_IRQCounter = 0 Then irq6502
                End If
            End If
        End If
        
        If CurrentLine = 0 Or CurrentLine = 131 Then updateSounds
        If CurrentLine = 258 Then PPU_Status = &H0

        If CurrentLine = 262 Then
            If maxIdle < 240 And Not SpritesChanged Then
                maxIdle = maxIdle + 16
            Else
                If maxIdle > 8 Then maxIdle = maxIdle - 8
                SpritesChanged = False
            End If
            
            'DF: repaint now when currentline=240
            
            CurrentLine = 0
            
            'DF: delays if too fast, skips frames if too slow.
            If Keyboard(192) And 1 Then
                render = (Frames Mod 3) = 0
            Else
                If autospeed Then
                    Dim delay As Long, ti As Long, ti2 As Long
                    Static tmr As Double
                    Static ptime As Double
                    Static facc As Double
                    Static pframe As Long
                    Dim tme As Double
                    tme = Timer
                    If tme - ptime < 0.2 And ptime < tme Then
                        tmr = tmr - tme + ptime
                    End If
                    tmr = tmr + 0.01667
                    ptime = tme
                    
                    render = True
                    If tmr > 0 Then
                        delay = 10000 * tmr * tmr
                        For ti = 0 To delay
                            For ti2 = 0 To 1000
                                vBuffer32(ti And 255) = ti2 'slow things down
                            Next ti2
                        Next ti
                    Else
                        If tmr < -1 Then
                            tmr = -0.2
                            If FrameSkip < 2 Then
                                frmNES.mnuFS_Click (FrameSkip) '+1 not needed
                            End If
                        End If
                    End If
                End If
                render = (Frames Mod FrameSkip = 0)
            End If
            
            PPU_Status = &H0
        Else
                CurrentLine = CurrentLine + 1
        End If
        clockticks6502 = clockticks6502 - maxCycles
        If Not SmartExec Then
            If CurrentLine < 218 Or maxCycles1 > 114 Then
                maxCycles = maxCycles1
            Else
                maxCycles = 114
            End If
        End If
  End If
  Wend
End Sub
Public Sub SetFlags(ByVal value As Byte)
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
Public Sub indzp6502()
'Added pc=pc+1, and (value+1) (Why Don?)
  value = Read6502(PC)
  PC = PC + 1
  savepc = Read6502(value) + (Read6502(value + 1) * &H100&)
End Sub

Public Sub zpy6502()
'TS: Added PC=PC+1
      savepc = Read6502(PC)
      savepc = savepc + Y
      PC = PC + 1
      'savepc = savepc And &HFF
End Sub

Public Sub absy6502()
'TS: Changed to != instead of == (Look at absx for more details)
  savepc = Read6502(PC) + (Read6502(PC + 1) * &H100&)
  PC = PC + 2

  If (Ticks(opcode) = 4) Then
    If ((savepc \ &H100&) = ((savepc + Y) \ &H100&)) Then
    Else
      clockticks6502 = clockticks6502 + 1
    End If
  End If
  savepc = savepc + Y
End Sub
Public Sub immediate6502()
  savepc = PC
  PC = PC + 1
End Sub
Public Sub indirect6502()
  help = Read6502(PC) + (Read6502(PC + 1) * &H100&)
  savepc = Read6502(help) + (Read6502(help + 1) * &H100&)
  PC = PC + 2
End Sub

Public Sub absx6502()
'TS: Changed to if then else instead of if then (!= instead of ==)
  savepc = Read6502(PC)
  savepc = savepc + (Read6502(PC + 1) * &H100&)
  PC = PC + 2

  If (Ticks(opcode) = 4) Then
    If ((savepc \ &H100&) = ((savepc + X) \ &H100&)) Then
    Else
      clockticks6502 = clockticks6502 + 1
    End If
  End If
  savepc = savepc + X
End Sub
Public Sub abs6502()
  savepc = Read6502(PC) + (Read6502(PC + 1) * &H100&)
  PC = PC + 2
End Sub
Public Sub relative6502()
'Changed to PC++ and == to != (If then else)
  
    savepc = Read6502(PC)
    PC = PC + 1

    If (savepc And &H80) Then savepc = (savepc - &H100&)

End Sub
Public Sub reset6502()
Dim i As Long
For i = 0 To 65535
    idleCheck(i) = 0
Next i

    A = 0: X = 0: Y = 0: P = &H20
    S = &HFF
      
    PC = Read6502(&HFFFC&) + (Read6502(&HFFFD&) * &H100&)
    Debug.Print "Reset to $" & Hex$(PC) & "[" & PC & "]"
End Sub
Public Sub zp6502()
  savepc = Read6502(PC)
  PC = PC + 1
End Sub
Public Sub irq6502()
' Maskable interrupt
If (P And &H4) = 0 Then
    Write6502 &H100& + S, (PC \ &H100&)
    S = (S - 1) And &HFF
    Write6502 &H100& + S, (PC And &HFF)
    S = (S - 1) And &HFF
    Write6502 &H100& + S, P
    S = (S - 1) And &HFF
    P = P Or &H4
    PC = Read6502(&HFFFE&) + (Read6502(&HFFFF&) * &H100&)
    clockticks6502 = clockticks6502 + 7
End If
End Sub
Public Sub nmi6502()
'TS: Changed PC>>8 to / not *
    Write6502 (S + &H100&), (PC \ &H100&)
    S = (S - 1) And &HFF
    Write6502 (S + &H100&), (PC And &HFF)
    S = (S - 1) And &HFF
    Write6502 (S + &H100&), P
    P = P Or &H4
    S = (S - 1) And &HFF
    PC = Read6502(&HFFFA&) + (Read6502(&HFFFB&) * &H100&)
    clockticks6502 = clockticks6502 + 7
End Sub
