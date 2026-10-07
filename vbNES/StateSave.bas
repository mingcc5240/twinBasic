Attribute VB_Name = "StateSave"

' State Saving/Movie Saving functions.
' Moved to StateSave.bas, 11/23/2002
' Don Jarrett, 1996-2003
Public Sub saveState(Index As Long)
    ' 12/3/01 - standardized for mapper support.
    ChDir App.Path
    Dim f As String
    f = romName + ".sv" + CStr(Index)
    Open App.Path & "\nessave.tmp" For Binary As #1
    'delete f
    'Open f For Binary As #1
        Put #1, , a
        Put #1, , X
        Put #1, , Y
        Put #1, , PC
        Put #1, , savepc
        Put #1, , P
        Put #1, , s
        Put #1, , value
        Put #1, , sum
        Put #1, , FirstRead
        Put #1, , Joypad1
        Put #1, , Joypad1_Count
        Put #1, , Mirroring
        Put #1, , Mapper
        Put #1, , Trainer
        Put #1, , MirrorXor
        Put #1, , HScroll
        Put #1, , VScroll
        Put #1, , bank_regs
        Put #1, , PPU_Control1
        Put #1, , PPU_Control2
        Put #1, , PPU_Status
        Put #1, , SpriteAddress
        Put #1, , PPUAddressHi
        Put #1, , PPUAddress
        Put #1, , PPU_AddressIsHi
        Put #1, , PatternTable
        Put #1, , NameTable
        Put #1, , reg8
        Put #1, , regA
        Put #1, , regC
        Put #1, , regE
        Put #1, , CurrentLine
        Put #1, , bank0
        Put #1, , bank6
        Put #1, , VRAM
        Put #1, , SpriteRAM
        
        Select Case Mapper
            Case 0, 2, 3, 5, 6, 7, 8, 11, 66, 68, 71, 78, 91: 'nothing to do here
            Case 1
                Put #1, , data
                Put #1, , accumulator
                Put #1, , sequence
            Case 4
                Put #1, , MMC3_Command
                Put #1, , MMC3_PrgAddr
                Put #1, , MMC3_ChrAddr
                Put #1, , MMC3_IrqVal
                Put #1, , MMC3_TmpVal
                Put #1, , MMC3_IrqOn
                Put #1, , swap
                Put #1, , PrgSwitch1
                Put #1, , PrgSwitch2
            Case 9, 10
                Put #1, , Latch0FD
                Put #1, , Latch0FE
                Put #1, , Latch1FD
                Put #1, , Latch1FE
            Case 13
                Put #1, , latch13
            Case 15
                Put #1, , Map15_BankAddr
                Put #1, , Map15_SwapReg
            Case 16
                Put #1, , tmpLatch
                Put #1, , MMC16_IrqOn
                Put #1, , MMC16_Irq
            Case 17
                Put #1, , map17_irqon
                Put #1, , map17_irq
            Case 19
                Put #1, , tmpLatch
                Put #1, , MIRQOn
                Put #1, , MMC19_IRQCount
            Case 24
                Put #1, , map24_irqv
                Put #1, , map24_irqon
                Put #1, , map24_irqv
            Case 32
                Put #1, , MMC32_Switch
            Case 40
                Put #1, , Mapper40_IRQEnabled
                Put #1, , Mapper40_IRQCounter
            Case 64
                Put #1, , cmd
                Put #1, , prg
                Put #1, , chr1
            Case 69
                Put #1, , reg8000
        End Select
        'added
        Put #1, , nt
    Close #1
    RLECompress "nessave.tmp", f
    delete "nessave.tmp"
End Sub
Public Sub loadState(Index As Long)
    ChDir App.Path
    Dim f As String
    f = romName + ".sv" + CStr(Index)
    If Dir$(App.Path & "\" & f) = "" Then
        frmNES.lblStatus.Caption = "There is no " & Index & " saved state."
        Exit Sub
    End If
    RLEDecompress f, "nesload.tmp"
    Open "nesload.tmp" For Binary As #1
    'Open f For Binary As #1
        Get #1, , a
        Get #1, , X
        Get #1, , Y
        Get #1, , PC
        Get #1, , savepc
        Get #1, , P
        Get #1, , s
        Get #1, , value
        Get #1, , sum
        Get #1, , FirstRead
        Get #1, , Joypad1
        Get #1, , Joypad1_Count
        Get #1, , Mirroring
        Get #1, , Mapper
        Get #1, , Trainer
        Get #1, , MirrorXor
        Get #1, , HScroll
        Get #1, , VScroll
        Get #1, , bank_regs
        Get #1, , PPU_Control1
        Get #1, , PPU_Control2
        Get #1, , PPU_Status
        Get #1, , SpriteAddress
        Get #1, , PPUAddressHi
        Get #1, , PPUAddress
        Get #1, , PPU_AddressIsHi
        Get #1, , PatternTable
        Get #1, , NameTable
        Get #1, , reg8
        Get #1, , regA
        Get #1, , regC
        Get #1, , regE
        Get #1, , CurrentLine
        Get #1, , bank0
        Get #1, , bank6
        Get #1, , VRAM
        Get #1, , SpriteRAM
        
        Select Case Mapper
            Case 0: 'nothing to do here
            Case 1
                Get #1, , data
                Get #1, , accumulator
                Get #1, , sequence
            Case 4
                Get #1, , MMC3_Command
                Get #1, , MMC3_PrgAddr
                Get #1, , MMC3_ChrAddr
                Get #1, , MMC3_IrqVal
                Get #1, , MMC3_TmpVal
                Get #1, , MMC3_IrqOn
                Get #1, , swap
                Get #1, , PrgSwitch1
                Get #1, , PrgSwitch2
            Case 9, 10
                Get #1, , Latch0FD
                Get #1, , Latch0FE
                Get #1, , Latch1FD
                Get #1, , Latch1FE
            Case 13
                Get #1, , latch13
            Case 15
                Get #1, , Map15_BankAddr
                Get #1, , Map15_SwapReg
            Case 16
                Get #1, , tmpLatch
                Get #1, , MMC16_IrqOn
                Get #1, , MMC16_Irq
            Case 17
                Get #1, , map17_irqon
                Get #1, , map17_irq
            Case 19
                Get #1, , tmpLatch
                Get #1, , MIRQOn
                Get #1, , MMC19_IRQCount
            Case 24
                Get #1, , map24_irqv
                Get #1, , map24_irqon
                Get #1, , map24_irqv
            Case 32
                Get #1, , MMC32_Switch
            Case 40
                Get #1, , Mapper40_IRQEnabled
                Get #1, , Mapper40_IRQCounter
            Case 64
                Get #1, , cmd
                Get #1, , prg
                Get #1, , chr1
            Case 69
                Get #1, , reg8000
        End Select
        Get #1, , nt
    Close #1
    SetupBanks
    delete "nesload.tmp"
End Sub
Public Sub PlayMovie(Index As Integer)
    ' 12/3/01 - standardized for mapper support.
    ChDir App.Path
    Dim f As String
    f = romName + ".mv" + CStr(Index)
    Open f For Binary As #1
    'delete f
    'Open f For Binary As #1
        Get #1, , a
        Get #1, , X
        Get #1, , Y
        Get #1, , PC
        Get #1, , savepc
        Get #1, , P
        Get #1, , s
        Get #1, , value
        Get #1, , sum
        Get #1, , FirstRead
        Get #1, , Joypad1
        Get #1, , Joypad1_Count
        Get #1, , Mirroring
        Get #1, , Mapper
        Get #1, , Trainer
        Get #1, , MirrorXor
        Get #1, , HScroll
        Get #1, , VScroll
        Get #1, , bank_regs
        Get #1, , PPU_Control1
        Get #1, , PPU_Control2
        Get #1, , PPU_Status
        Get #1, , SpriteAddress
        Get #1, , PPUAddressHi
        Get #1, , PPUAddress
        Get #1, , PPU_AddressIsHi
        Get #1, , PatternTable
        Get #1, , NameTable
        Get #1, , reg8
        Get #1, , regA
        Get #1, , regC
        Get #1, , regE
        Get #1, , CurrentLine
        Get #1, , bank0
        Get #1, , bank6
        Get #1, , VRAM
        Get #1, , SpriteRAM
        
        Select Case Mapper
            Case 0: 'nothing to do here
            Case 1
                Get #1, , data
                Get #1, , accumulator
                Get #1, , sequence
            Case 4
                Get #1, , MMC3_Command
                Get #1, , MMC3_PrgAddr
                Get #1, , MMC3_ChrAddr
                Get #1, , MMC3_IrqVal
                Get #1, , MMC3_TmpVal
                Get #1, , MMC3_IrqOn
                Get #1, , swap
                Get #1, , PrgSwitch1
                Get #1, , PrgSwitch2
            Case 9, 10
                Get #1, , Latch0FD
                Get #1, , Latch0FE
                Get #1, , Latch1FD
                Get #1, , Latch1FE
            Case 13
                Get #1, , latch13
            Case 15
                Get #1, , Map15_BankAddr
                Get #1, , Map15_SwapReg
            Case 16
                Get #1, , tmpLatch
                Get #1, , MMC16_IrqOn
                Get #1, , MMC16_Irq
            Case 17
                Get #1, , map17_irqon
                Get #1, , map17_irq
            Case 19
                Get #1, , tmpLatch
                Get #1, , MIRQOn
                Get #1, , MMC19_IRQCount
            Case 24
                Get #1, , map24_irqv
                Get #1, , map24_irqon
                Get #1, , map24_irqv
            Case 32
                Get #1, , MMC32_Switch
            Case 40
                Get #1, , Mapper40_IRQEnabled
                Get #1, , Mapper40_IRQCounter
            Case 64
                Get #1, , cmd
                Get #1, , prg
                Get #1, , chr1
        End Select
        Get #1, , nt
        Playing = True

End Sub
Public Sub StopPlaying()
    Playing = False
    frmNES.lblStatus.Caption = "Stopped playing"
    Close #1
End Sub
Public Sub RecordMovie(Index As Long)
    ' 12/3/01 - standardized for mapper support.
    ChDir App.Path
    Dim f As String
    f = romName + ".mv" + CStr(Index)
    Open f For Binary As #1
    'delete f
    'Open f For Binary As #1
        Put #1, , a
        Put #1, , X
        Put #1, , Y
        Put #1, , PC
        Put #1, , savepc
        Put #1, , P
        Put #1, , s
        Put #1, , value
        Put #1, , sum
        Put #1, , FirstRead
        Put #1, , Joypad1
        Put #1, , Joypad1_Count
        Put #1, , Mirroring
        Put #1, , Mapper
        Put #1, , Trainer
        Put #1, , MirrorXor
        Put #1, , HScroll
        Put #1, , VScroll
        Put #1, , bank_regs
        Put #1, , PPU_Control1
        Put #1, , PPU_Control2
        Put #1, , PPU_Status
        Put #1, , SpriteAddress
        Put #1, , PPUAddressHi
        Put #1, , PPUAddress
        Put #1, , PPU_AddressIsHi
        Put #1, , PatternTable
        Put #1, , NameTable
        Put #1, , reg8
        Put #1, , regA
        Put #1, , regC
        Put #1, , regE
        Put #1, , CurrentLine
        Put #1, , bank0
        Put #1, , bank6
        Put #1, , VRAM
        Put #1, , SpriteRAM
        Select Case Mapper
            Case 0: 'nothing to do here
            Case 1
                Put #1, , data
                Put #1, , accumulator
                Put #1, , sequence
            Case 4
                Put #1, , MMC3_Command
                Put #1, , MMC3_PrgAddr
                Put #1, , MMC3_ChrAddr
                Put #1, , MMC3_IrqVal
                Put #1, , MMC3_TmpVal
                Put #1, , MMC3_IrqOn
                Put #1, , swap
                Put #1, , PrgSwitch1
                Put #1, , PrgSwitch2
            Case 9, 10
                Put #1, , Latch0FD
                Put #1, , Latch0FE
                Put #1, , Latch1FD
                Put #1, , Latch1FE
            Case 13
                Put #1, , latch13
            Case 15
                Put #1, , Map15_BankAddr
                Put #1, , Map15_SwapReg
            Case 16
                Put #1, , tmpLatch
                Put #1, , MMC16_IrqOn
                Put #1, , MMC16_Irq
            Case 17
                Put #1, , map17_irqon
                Put #1, , map17_irq
            Case 19
                Put #1, , tmpLatch
                Put #1, , MIRQOn
                Put #1, , MMC19_IRQCount
            Case 24
                Put #1, , map24_irqv
                Put #1, , map24_irqon
                Put #1, , map24_irqv
            Case 32
                Put #1, , MMC32_Switch
            Case 40
                Put #1, , Mapper40_IRQEnabled
                Put #1, , Mapper40_IRQCounter
            Case 64
                Put #1, , cmd
                Put #1, , prg
                Put #1, , chr1
        End Select
        Put #1, , nt
        Record = True
    'don't close the file till the movie is stopped.
End Sub
Public Sub StopRecording()
Close #1
frmNES.lblStatus.Caption = "Stopped Recording"
Record = False
End Sub
