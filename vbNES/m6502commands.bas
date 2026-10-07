Attribute VB_Name = "m6502Commands"
Option Explicit

DefLng A-Z

' Declarations for M6502
' Addressing Modes
Public Const ADR_ABS As Long = 0
Public Const ADR_ABSX As Long = 1
Public Const ADR_ABSY  As Long = 2
Public Const ADR_IMM As Long = 3
Public Const ADR_IMP As Long = 4
Public Const ADR_INDABSX As Long = 5
Public Const ADR_IND As Long = 6
Public Const ADR_INDX As Long = 7
Public Const ADR_INDY As Long = 8
Public Const ADR_INDZP As Long = 9
Public Const ADR_REL As Long = 10
Public Const ADR_ZP As Long = 11
Public Const ADR_ZPX As Long = 12
Public Const ADR_ZPY As Long = 13

' Opcodes
Public Const INS_ADC As Long = 0
Public Const INS_AND As Long = 1
Public Const INS_ASL As Long = 2
Public Const INS_ASLA As Long = 3
Public Const INS_BCC As Long = 4
Public Const INS_BCS As Long = 5
Public Const INS_BEQ As Long = 6
Public Const INS_BIT As Long = 7
Public Const INS_BMI As Long = 8
Public Const INS_BNE As Long = 9
Public Const INS_BPL As Long = 10
Public Const INS_BRK As Long = 11
Public Const INS_BVC As Long = 12
Public Const INS_BVS As Long = 13
Public Const INS_CLC As Long = 14
Public Const INS_CLD As Long = 15
Public Const INS_CLI As Long = 16
Public Const INS_CLV As Long = 17
Public Const INS_CMP As Long = 18
Public Const INS_CPX As Long = 19
Public Const INS_CPY As Long = 20
Public Const INS_DEC As Long = 21
Public Const INS_DEA As Long = 22
Public Const INS_DEX As Long = 23
Public Const INS_DEY As Long = 24
Public Const INS_EOR As Long = 25
Public Const INS_INC As Long = 26
Public Const INS_INX As Long = 27
Public Const INS_INY As Long = 28
Public Const INS_JMP As Long = 29
Public Const INS_JSR As Long = 30
Public Const INS_LDA As Long = 31
Public Const INS_LDX As Long = 32
Public Const INS_LDY As Long = 33
Public Const INS_LSR As Long = 34
Public Const INS_LSRA As Long = 35
Public Const INS_NOP As Long = 36
Public Const INS_ORA As Long = 37
Public Const INS_PHA As Long = 38
Public Const INS_PHP As Long = 39
Public Const INS_PLA As Long = 40
Public Const INS_PLP As Long = 41
Public Const INS_ROL As Long = 42
Public Const INS_ROLA As Long = 43
Public Const INS_ROR As Long = 44
Public Const INS_RORA As Long = 45
Public Const INS_RTI As Long = 46
Public Const INS_RTS As Long = 47
Public Const INS_SBC As Long = 48
Public Const INS_SEC As Long = 49
Public Const INS_SED As Long = 50
Public Const INS_SEI As Long = 51
Public Const INS_STA As Long = 52
Public Const INS_STX As Long = 53
Public Const INS_STY As Long = 54
Public Const INS_TAX As Long = 55
Public Const INS_TAY As Long = 56
Public Const INS_TSX As Long = 57
Public Const INS_TXA As Long = 58
Public Const INS_TXS As Long = 59
Public Const INS_TYA As Long = 60
Public Const INS_BRA As Long = 61
Public Const INS_INA As Long = 62
Public Const INS_PHX As Long = 63
Public Const INS_PLX As Long = 64
Public Const INS_PHY As Long = 65
Public Const INS_PLY As Long = 66
Public Function init6502()
      Ticks(&H0) = 7: instruction(&H0) = INS_BRK: addrmode(&H0) = ADR_IMP
      Ticks(&H1) = 6: instruction(&H1) = INS_ORA: addrmode(&H1) = ADR_INDX
      Ticks(&H2) = 2: instruction(&H2) = INS_NOP: addrmode(&H2) = ADR_IMP
      Ticks(&H3) = 2: instruction(&H3) = INS_NOP: addrmode(&H3) = ADR_IMP
      Ticks(&H4) = 3: instruction(&H4) = INS_NOP: addrmode(&H4) = ADR_ZP
      Ticks(&H5) = 3: instruction(&H5) = INS_ORA: addrmode(&H5) = ADR_ZP
      Ticks(&H6) = 5: instruction(&H6) = INS_ASL: addrmode(&H6) = ADR_ZP
      Ticks(&H7) = 2: instruction(&H7) = INS_NOP: addrmode(&H7) = ADR_IMP
      Ticks(&H8) = 3: instruction(&H8) = INS_PHP: addrmode(&H8) = ADR_IMP
      Ticks(&H9) = 3: instruction(&H9) = INS_ORA: addrmode(&H9) = ADR_IMM
      Ticks(&HA) = 2: instruction(&HA) = INS_ASLA: addrmode(&HA) = ADR_IMP
      Ticks(&HB) = 2: instruction(&HB) = INS_NOP: addrmode(&HB) = ADR_IMP
      Ticks(&HC) = 4: instruction(&HC) = INS_NOP: addrmode(&HC) = ADR_ABS
      Ticks(&HD) = 4: instruction(&HD) = INS_ORA: addrmode(&HD) = ADR_ABS
      Ticks(&HE) = 6: instruction(&HE) = INS_ASL: addrmode(&HE) = ADR_ABS
      Ticks(&HF) = 2: instruction(&HF) = INS_NOP: addrmode(&HF) = ADR_IMP
      Ticks(&H10) = 2: instruction(&H10) = INS_BPL: addrmode(&H10) = ADR_REL
      Ticks(&H11) = 5: instruction(&H11) = INS_ORA: addrmode(&H11) = ADR_INDY
      Ticks(&H12) = 3: instruction(&H12) = INS_ORA: addrmode(&H12) = ADR_INDZP
      Ticks(&H13) = 2: instruction(&H13) = INS_NOP: addrmode(&H13) = ADR_IMP
      Ticks(&H14) = 3: instruction(&H14) = INS_NOP: addrmode(&H14) = ADR_ZP
      Ticks(&H15) = 4: instruction(&H15) = INS_ORA: addrmode(&H15) = ADR_ZPX
      Ticks(&H16) = 6: instruction(&H16) = INS_ASL: addrmode(&H16) = ADR_ZPX
      Ticks(&H17) = 2: instruction(&H17) = INS_NOP: addrmode(&H17) = ADR_IMP
      Ticks(&H18) = 2: instruction(&H18) = INS_CLC: addrmode(&H18) = ADR_IMP
      Ticks(&H19) = 4: instruction(&H19) = INS_ORA: addrmode(&H19) = ADR_ABSY
      Ticks(&H1A) = 2: instruction(&H1A) = INS_INA: addrmode(&H1A) = ADR_IMP
      Ticks(&H1B) = 2: instruction(&H1B) = INS_NOP: addrmode(&H1B) = ADR_IMP
      Ticks(&H1C) = 4: instruction(&H1C) = INS_NOP: addrmode(&H1C) = ADR_ABS
      Ticks(&H1D) = 4: instruction(&H1D) = INS_ORA: addrmode(&H1D) = ADR_ABSX
      Ticks(&H1E) = 7: instruction(&H1E) = INS_ASL: addrmode(&H1E) = ADR_ABSX
      Ticks(&H1F) = 2: instruction(&H1F) = INS_NOP: addrmode(&H1F) = ADR_IMP
      Ticks(&H20) = 6: instruction(&H20) = INS_JSR: addrmode(&H20) = ADR_ABS
      Ticks(&H21) = 6: instruction(&H21) = INS_AND: addrmode(&H21) = ADR_INDX
      Ticks(&H22) = 2: instruction(&H22) = INS_NOP: addrmode(&H22) = ADR_IMP
      Ticks(&H23) = 2: instruction(&H23) = INS_NOP: addrmode(&H23) = ADR_IMP
      Ticks(&H24) = 3: instruction(&H24) = INS_BIT: addrmode(&H24) = ADR_ZP
      Ticks(&H25) = 3: instruction(&H25) = INS_AND: addrmode(&H25) = ADR_ZP
      Ticks(&H26) = 5: instruction(&H26) = INS_ROL: addrmode(&H26) = ADR_ZP
      Ticks(&H27) = 2: instruction(&H27) = INS_NOP: addrmode(&H27) = ADR_IMP
      Ticks(&H28) = 4: instruction(&H28) = INS_PLP: addrmode(&H28) = ADR_IMP
      Ticks(&H29) = 3: instruction(&H29) = INS_AND: addrmode(&H29) = ADR_IMM
      Ticks(&H2A) = 2: instruction(&H2A) = INS_ROLA: addrmode(&H2A) = ADR_IMP
      Ticks(&H2B) = 2: instruction(&H2B) = INS_NOP: addrmode(&H2B) = ADR_IMP
      Ticks(&H2C) = 4: instruction(&H2C) = INS_BIT: addrmode(&H2C) = ADR_ABS
      Ticks(&H2D) = 4: instruction(&H2D) = INS_AND: addrmode(&H2D) = ADR_ABS
      Ticks(&H2E) = 6: instruction(&H2E) = INS_ROL: addrmode(&H2E) = ADR_ABS
      Ticks(&H2F) = 2: instruction(&H2F) = INS_NOP: addrmode(&H2F) = ADR_IMP
      Ticks(&H30) = 2: instruction(&H30) = INS_BMI: addrmode(&H30) = ADR_REL
      Ticks(&H31) = 5: instruction(&H31) = INS_AND: addrmode(&H31) = ADR_INDY
      Ticks(&H32) = 3: instruction(&H32) = INS_AND: addrmode(&H32) = ADR_INDZP
      Ticks(&H33) = 2: instruction(&H33) = INS_NOP: addrmode(&H33) = ADR_IMP
      Ticks(&H34) = 4: instruction(&H34) = INS_BIT: addrmode(&H34) = ADR_ZPX
      Ticks(&H35) = 4: instruction(&H35) = INS_AND: addrmode(&H35) = ADR_ZPX
      Ticks(&H36) = 6: instruction(&H36) = INS_ROL: addrmode(&H36) = ADR_ZPX
      Ticks(&H37) = 2: instruction(&H37) = INS_NOP: addrmode(&H37) = ADR_IMP
      Ticks(&H38) = 2: instruction(&H38) = INS_SEC: addrmode(&H38) = ADR_IMP
      Ticks(&H39) = 4: instruction(&H39) = INS_AND: addrmode(&H39) = ADR_ABSY
      Ticks(&H3A) = 2: instruction(&H3A) = INS_DEA: addrmode(&H3A) = ADR_IMP
      Ticks(&H3B) = 2: instruction(&H3B) = INS_NOP: addrmode(&H3B) = ADR_IMP
      Ticks(&H3C) = 4: instruction(&H3C) = INS_BIT: addrmode(&H3C) = ADR_ABSX
      Ticks(&H3D) = 4: instruction(&H3D) = INS_AND: addrmode(&H3D) = ADR_ABSX
      Ticks(&H3E) = 7: instruction(&H3E) = INS_ROL: addrmode(&H3E) = ADR_ABSX
      Ticks(&H3F) = 2: instruction(&H3F) = INS_NOP: addrmode(&H3F) = ADR_IMP
      Ticks(&H40) = 6: instruction(&H40) = INS_RTI: addrmode(&H40) = ADR_IMP
      Ticks(&H41) = 6: instruction(&H41) = INS_EOR: addrmode(&H41) = ADR_INDX
      Ticks(&H42) = 2: instruction(&H42) = INS_NOP: addrmode(&H42) = ADR_IMP
      Ticks(&H43) = 2: instruction(&H43) = INS_NOP: addrmode(&H43) = ADR_IMP
      Ticks(&H44) = 2: instruction(&H44) = INS_NOP: addrmode(&H44) = ADR_IMP
      Ticks(&H45) = 3: instruction(&H45) = INS_EOR: addrmode(&H45) = ADR_ZP
      Ticks(&H46) = 5: instruction(&H46) = INS_LSR: addrmode(&H46) = ADR_ZP
      Ticks(&H47) = 2: instruction(&H47) = INS_NOP: addrmode(&H47) = ADR_IMP
      Ticks(&H48) = 3: instruction(&H48) = INS_PHA: addrmode(&H48) = ADR_IMP
      Ticks(&H49) = 3: instruction(&H49) = INS_EOR: addrmode(&H49) = ADR_IMM
      Ticks(&H4A) = 2: instruction(&H4A) = INS_LSRA: addrmode(&H4A) = ADR_IMP
      Ticks(&H4B) = 2: instruction(&H4B) = INS_NOP: addrmode(&H4B) = ADR_IMP
      Ticks(&H4C) = 3: instruction(&H4C) = INS_JMP: addrmode(&H4C) = ADR_ABS
      Ticks(&H4D) = 4: instruction(&H4D) = INS_EOR: addrmode(&H4D) = ADR_ABS
      Ticks(&H4E) = 6: instruction(&H4E) = INS_LSR: addrmode(&H4E) = ADR_ABS
      Ticks(&H4F) = 2: instruction(&H4F) = INS_NOP: addrmode(&H4F) = ADR_IMP
      Ticks(&H50) = 2: instruction(&H50) = INS_BVC: addrmode(&H50) = ADR_REL
      Ticks(&H51) = 5: instruction(&H51) = INS_EOR: addrmode(&H51) = ADR_INDY
      Ticks(&H52) = 3: instruction(&H52) = INS_EOR: addrmode(&H52) = ADR_INDZP
      Ticks(&H53) = 2: instruction(&H53) = INS_NOP: addrmode(&H53) = ADR_IMP
      Ticks(&H54) = 2: instruction(&H54) = INS_NOP: addrmode(&H54) = ADR_IMP
      Ticks(&H55) = 4: instruction(&H55) = INS_EOR: addrmode(&H55) = ADR_ZPX
      Ticks(&H56) = 6: instruction(&H56) = INS_LSR: addrmode(&H56) = ADR_ZPX
      Ticks(&H57) = 2: instruction(&H57) = INS_NOP: addrmode(&H57) = ADR_IMP
      Ticks(&H58) = 2: instruction(&H58) = INS_CLI: addrmode(&H58) = ADR_IMP
      Ticks(&H59) = 4: instruction(&H59) = INS_EOR: addrmode(&H59) = ADR_ABSY
      Ticks(&H5A) = 3: instruction(&H5A) = INS_PHY: addrmode(&H5A) = ADR_IMP
      Ticks(&H5B) = 2: instruction(&H5B) = INS_NOP: addrmode(&H5B) = ADR_IMP
      Ticks(&H5C) = 2: instruction(&H5C) = INS_NOP: addrmode(&H5C) = ADR_IMP
      Ticks(&H5D) = 4: instruction(&H5D) = INS_EOR: addrmode(&H5D) = ADR_ABSX
      Ticks(&H5E) = 7: instruction(&H5E) = INS_LSR: addrmode(&H5E) = ADR_ABSX
      Ticks(&H5F) = 2: instruction(&H5F) = INS_NOP: addrmode(&H5F) = ADR_IMP
      Ticks(&H60) = 6: instruction(&H60) = INS_RTS: addrmode(&H60) = ADR_IMP
      Ticks(&H61) = 6: instruction(&H61) = INS_ADC: addrmode(&H61) = ADR_INDX
      Ticks(&H62) = 2: instruction(&H62) = INS_NOP: addrmode(&H62) = ADR_IMP
      Ticks(&H63) = 2: instruction(&H63) = INS_NOP: addrmode(&H63) = ADR_IMP
      Ticks(&H64) = 3: instruction(&H64) = INS_NOP: addrmode(&H64) = ADR_ZP
      Ticks(&H65) = 3: instruction(&H65) = INS_ADC: addrmode(&H65) = ADR_ZP
      Ticks(&H66) = 5: instruction(&H66) = INS_ROR: addrmode(&H66) = ADR_ZP
      Ticks(&H67) = 2: instruction(&H67) = INS_NOP: addrmode(&H67) = ADR_IMP
      Ticks(&H68) = 4: instruction(&H68) = INS_PLA: addrmode(&H68) = ADR_IMP
      Ticks(&H69) = 3: instruction(&H69) = INS_ADC: addrmode(&H69) = ADR_IMM
      Ticks(&H6A) = 2: instruction(&H6A) = INS_RORA: addrmode(&H6A) = ADR_IMP
      Ticks(&H6B) = 2: instruction(&H6B) = INS_NOP: addrmode(&H6B) = ADR_IMP
      Ticks(&H6C) = 5: instruction(&H6C) = INS_JMP: addrmode(&H6C) = ADR_IND
      Ticks(&H6D) = 4: instruction(&H6D) = INS_ADC: addrmode(&H6D) = ADR_ABS
      Ticks(&H6E) = 6: instruction(&H6E) = INS_ROR: addrmode(&H6E) = ADR_ABS
      Ticks(&H6F) = 2: instruction(&H6F) = INS_NOP: addrmode(&H6F) = ADR_IMP
      Ticks(&H70) = 2: instruction(&H70) = INS_BVS: addrmode(&H70) = ADR_REL
      Ticks(&H71) = 5: instruction(&H71) = INS_ADC: addrmode(&H71) = ADR_INDY
      Ticks(&H72) = 3: instruction(&H72) = INS_ADC: addrmode(&H72) = ADR_INDZP
      Ticks(&H73) = 2: instruction(&H73) = INS_NOP: addrmode(&H73) = ADR_IMP
      Ticks(&H74) = 4: instruction(&H74) = INS_NOP: addrmode(&H74) = ADR_ZPX
      Ticks(&H75) = 4: instruction(&H75) = INS_ADC: addrmode(&H75) = ADR_ZPX
      Ticks(&H76) = 6: instruction(&H76) = INS_ROR: addrmode(&H76) = ADR_ZPX
      Ticks(&H77) = 2: instruction(&H77) = INS_NOP: addrmode(&H77) = ADR_IMP
      Ticks(&H78) = 2: instruction(&H78) = INS_SEI: addrmode(&H78) = ADR_IMP
      Ticks(&H79) = 4: instruction(&H79) = INS_ADC: addrmode(&H79) = ADR_ABSY
      Ticks(&H7A) = 4: instruction(&H7A) = INS_PLY: addrmode(&H7A) = ADR_IMP
      Ticks(&H7B) = 2: instruction(&H7B) = INS_NOP: addrmode(&H7B) = ADR_IMP
      Ticks(&H7C) = 6: instruction(&H7C) = INS_JMP: addrmode(&H7C) = ADR_INDABSX
      Ticks(&H7D) = 4: instruction(&H7D) = INS_ADC: addrmode(&H7D) = ADR_ABSX
      Ticks(&H7E) = 7: instruction(&H7E) = INS_ROR: addrmode(&H7E) = ADR_ABSX
      Ticks(&H7F) = 2: instruction(&H7F) = INS_NOP: addrmode(&H7F) = ADR_IMP
      Ticks(&H80) = 2: instruction(&H80) = INS_BRA: addrmode(&H80) = ADR_REL
      Ticks(&H81) = 6: instruction(&H81) = INS_STA: addrmode(&H81) = ADR_INDX
      Ticks(&H82) = 2: instruction(&H82) = INS_NOP: addrmode(&H82) = ADR_IMP
      Ticks(&H83) = 2: instruction(&H83) = INS_NOP: addrmode(&H83) = ADR_IMP
      Ticks(&H84) = 2: instruction(&H84) = INS_STY: addrmode(&H84) = ADR_ZP
      Ticks(&H85) = 2: instruction(&H85) = INS_STA: addrmode(&H85) = ADR_ZP
      Ticks(&H86) = 2: instruction(&H86) = INS_STX: addrmode(&H86) = ADR_ZP
      Ticks(&H87) = 2: instruction(&H87) = INS_NOP: addrmode(&H87) = ADR_IMP
      Ticks(&H88) = 2: instruction(&H88) = INS_DEY: addrmode(&H88) = ADR_IMP
      Ticks(&H89) = 2: instruction(&H89) = INS_BIT: addrmode(&H89) = ADR_IMM
      Ticks(&H8A) = 2: instruction(&H8A) = INS_TXA: addrmode(&H8A) = ADR_IMP
      Ticks(&H8B) = 2: instruction(&H8B) = INS_NOP: addrmode(&H8B) = ADR_IMP
      Ticks(&H8C) = 4: instruction(&H8C) = INS_STY: addrmode(&H8C) = ADR_ABS
      Ticks(&H8D) = 4: instruction(&H8D) = INS_STA: addrmode(&H8D) = ADR_ABS
      Ticks(&H8E) = 4: instruction(&H8E) = INS_STX: addrmode(&H8E) = ADR_ABS
      Ticks(&H8F) = 2: instruction(&H8F) = INS_NOP: addrmode(&H8F) = ADR_IMP
      Ticks(&H90) = 2: instruction(&H90) = INS_BCC: addrmode(&H90) = ADR_REL
      Ticks(&H91) = 6: instruction(&H91) = INS_STA: addrmode(&H91) = ADR_INDY
      Ticks(&H92) = 3: instruction(&H92) = INS_STA: addrmode(&H92) = ADR_INDZP
      Ticks(&H93) = 2: instruction(&H93) = INS_NOP: addrmode(&H93) = ADR_IMP
      Ticks(&H94) = 4: instruction(&H94) = INS_STY: addrmode(&H94) = ADR_ZPX
      Ticks(&H95) = 4: instruction(&H95) = INS_STA: addrmode(&H95) = ADR_ZPX
      Ticks(&H96) = 4: instruction(&H96) = INS_STX: addrmode(&H96) = ADR_ZPY
      Ticks(&H97) = 2: instruction(&H97) = INS_NOP: addrmode(&H97) = ADR_IMP
      Ticks(&H98) = 2: instruction(&H98) = INS_TYA: addrmode(&H98) = ADR_IMP
      Ticks(&H99) = 5: instruction(&H99) = INS_STA: addrmode(&H99) = ADR_ABSY
      Ticks(&H9A) = 2: instruction(&H9A) = INS_TXS: addrmode(&H9A) = ADR_IMP
      Ticks(&H9B) = 2: instruction(&H9B) = INS_NOP: addrmode(&H9B) = ADR_IMP
      Ticks(&H9C) = 4: instruction(&H9C) = INS_NOP: addrmode(&H9C) = ADR_ABS
      Ticks(&H9D) = 5: instruction(&H9D) = INS_STA: addrmode(&H9D) = ADR_ABSX
      Ticks(&H9E) = 5: instruction(&H9E) = INS_NOP: addrmode(&H9E) = ADR_ABSX
      Ticks(&H9F) = 2: instruction(&H9F) = INS_NOP: addrmode(&H9F) = ADR_IMP
      Ticks(&HA0) = 3: instruction(&HA0) = INS_LDY: addrmode(&HA0) = ADR_IMM
      Ticks(&HA1) = 6: instruction(&HA1) = INS_LDA: addrmode(&HA1) = ADR_INDX
      Ticks(&HA2) = 3: instruction(&HA2) = INS_LDX: addrmode(&HA2) = ADR_IMM
      Ticks(&HA3) = 2: instruction(&HA3) = INS_NOP: addrmode(&HA3) = ADR_IMP
      Ticks(&HA4) = 3: instruction(&HA4) = INS_LDY: addrmode(&HA4) = ADR_ZP
      Ticks(&HA5) = 3: instruction(&HA5) = INS_LDA: addrmode(&HA5) = ADR_ZP
      Ticks(&HA6) = 3: instruction(&HA6) = INS_LDX: addrmode(&HA6) = ADR_ZP
      Ticks(&HA7) = 2: instruction(&HA7) = INS_NOP: addrmode(&HA7) = ADR_IMP
      Ticks(&HA8) = 2: instruction(&HA8) = INS_TAY: addrmode(&HA8) = ADR_IMP
      Ticks(&HA9) = 3: instruction(&HA9) = INS_LDA: addrmode(&HA9) = ADR_IMM
      Ticks(&HAA) = 2: instruction(&HAA) = INS_TAX: addrmode(&HAA) = ADR_IMP
      Ticks(&HAB) = 2: instruction(&HAB) = INS_NOP: addrmode(&HAB) = ADR_IMP
      Ticks(&HAC) = 4: instruction(&HAC) = INS_LDY: addrmode(&HAC) = ADR_ABS
      Ticks(&HAD) = 4: instruction(&HAD) = INS_LDA: addrmode(&HAD) = ADR_ABS
      Ticks(&HAE) = 4: instruction(&HAE) = INS_LDX: addrmode(&HAE) = ADR_ABS
      Ticks(&HAF) = 2: instruction(&HAF) = INS_NOP: addrmode(&HAF) = ADR_IMP
      Ticks(&HB0) = 2: instruction(&HB0) = INS_BCS: addrmode(&HB0) = ADR_REL
      Ticks(&HB1) = 5: instruction(&HB1) = INS_LDA: addrmode(&HB1) = ADR_INDY
      Ticks(&HB2) = 3: instruction(&HB2) = INS_LDA: addrmode(&HB2) = ADR_INDZP
      Ticks(&HB3) = 2: instruction(&HB3) = INS_NOP: addrmode(&HB3) = ADR_IMP
      Ticks(&HB4) = 4: instruction(&HB4) = INS_LDY: addrmode(&HB4) = ADR_ZPX
      Ticks(&HB5) = 4: instruction(&HB5) = INS_LDA: addrmode(&HB5) = ADR_ZPX
      Ticks(&HB6) = 4: instruction(&HB6) = INS_LDX: addrmode(&HB6) = ADR_ZPY
      Ticks(&HB7) = 2: instruction(&HB7) = INS_NOP: addrmode(&HB7) = ADR_IMP
      Ticks(&HB8) = 2: instruction(&HB8) = INS_CLV: addrmode(&HB8) = ADR_IMP
      Ticks(&HB9) = 4: instruction(&HB9) = INS_LDA: addrmode(&HB9) = ADR_ABSY
      Ticks(&HBA) = 2: instruction(&HBA) = INS_TSX: addrmode(&HBA) = ADR_IMP
      Ticks(&HBB) = 2: instruction(&HBB) = INS_NOP: addrmode(&HBB) = ADR_IMP
      Ticks(&HBC) = 4: instruction(&HBC) = INS_LDY: addrmode(&HBC) = ADR_ABSX
      Ticks(&HBD) = 4: instruction(&HBD) = INS_LDA: addrmode(&HBD) = ADR_ABSX
      Ticks(&HBE) = 4: instruction(&HBE) = INS_LDX: addrmode(&HBE) = ADR_ABSY
      Ticks(&HBF) = 2: instruction(&HBF) = INS_NOP: addrmode(&HBF) = ADR_IMP
      Ticks(&HC0) = 3: instruction(&HC0) = INS_CPY: addrmode(&HC0) = ADR_IMM
      Ticks(&HC1) = 6: instruction(&HC1) = INS_CMP: addrmode(&HC1) = ADR_INDX
      Ticks(&HC2) = 2: instruction(&HC2) = INS_NOP: addrmode(&HC2) = ADR_IMP
      Ticks(&HC3) = 2: instruction(&HC3) = INS_NOP: addrmode(&HC3) = ADR_IMP
      Ticks(&HC4) = 3: instruction(&HC4) = INS_CPY: addrmode(&HC4) = ADR_ZP
      Ticks(&HC5) = 3: instruction(&HC5) = INS_CMP: addrmode(&HC5) = ADR_ZP
      Ticks(&HC6) = 5: instruction(&HC6) = INS_DEC: addrmode(&HC6) = ADR_ZP
      Ticks(&HC7) = 2: instruction(&HC7) = INS_NOP: addrmode(&HC7) = ADR_IMP
      Ticks(&HC8) = 2: instruction(&HC8) = INS_INY: addrmode(&HC8) = ADR_IMP
      Ticks(&HC9) = 3: instruction(&HC9) = INS_CMP: addrmode(&HC9) = ADR_IMM
      Ticks(&HCA) = 2: instruction(&HCA) = INS_DEX: addrmode(&HCA) = ADR_IMP
      Ticks(&HCB) = 2: instruction(&HCB) = INS_NOP: addrmode(&HCB) = ADR_IMP
      Ticks(&HCC) = 4: instruction(&HCC) = INS_CPY: addrmode(&HCC) = ADR_ABS
      Ticks(&HCD) = 4: instruction(&HCD) = INS_CMP: addrmode(&HCD) = ADR_ABS
      Ticks(&HCE) = 6: instruction(&HCE) = INS_DEC: addrmode(&HCE) = ADR_ABS
      Ticks(&HCF) = 2: instruction(&HCF) = INS_NOP: addrmode(&HCF) = ADR_IMP
      Ticks(&HD0) = 2: instruction(&HD0) = INS_BNE: addrmode(&HD0) = ADR_REL
      Ticks(&HD1) = 5: instruction(&HD1) = INS_CMP: addrmode(&HD1) = ADR_INDY
      Ticks(&HD2) = 3: instruction(&HD2) = INS_CMP: addrmode(&HD2) = ADR_INDZP
      Ticks(&HD3) = 2: instruction(&HD3) = INS_NOP: addrmode(&HD3) = ADR_IMP
      Ticks(&HD4) = 2: instruction(&HD4) = INS_NOP: addrmode(&HD4) = ADR_IMP
      Ticks(&HD5) = 4: instruction(&HD5) = INS_CMP: addrmode(&HD5) = ADR_ZPX
      Ticks(&HD6) = 6: instruction(&HD6) = INS_DEC: addrmode(&HD6) = ADR_ZPX
      Ticks(&HD7) = 2: instruction(&HD7) = INS_NOP: addrmode(&HD7) = ADR_IMP
      Ticks(&HD8) = 2: instruction(&HD8) = INS_CLD: addrmode(&HD8) = ADR_IMP
      Ticks(&HD9) = 4: instruction(&HD9) = INS_CMP: addrmode(&HD9) = ADR_ABSY
      Ticks(&HDA) = 3: instruction(&HDA) = INS_PHX: addrmode(&HDA) = ADR_IMP
      Ticks(&HDB) = 2: instruction(&HDB) = INS_NOP: addrmode(&HDB) = ADR_IMP
      Ticks(&HDC) = 2: instruction(&HDC) = INS_NOP: addrmode(&HDC) = ADR_IMP
      Ticks(&HDD) = 4: instruction(&HDD) = INS_CMP: addrmode(&HDD) = ADR_ABSX
      Ticks(&HDE) = 7: instruction(&HDE) = INS_DEC: addrmode(&HDE) = ADR_ABSX
      Ticks(&HDF) = 2: instruction(&HDF) = INS_NOP: addrmode(&HDF) = ADR_IMP
      Ticks(&HE0) = 3: instruction(&HE0) = INS_CPX: addrmode(&HE0) = ADR_IMM
      Ticks(&HE1) = 6: instruction(&HE1) = INS_SBC: addrmode(&HE1) = ADR_INDX
      Ticks(&HE2) = 2: instruction(&HE2) = INS_NOP: addrmode(&HE2) = ADR_IMP
      Ticks(&HE3) = 2: instruction(&HE3) = INS_NOP: addrmode(&HE3) = ADR_IMP
      Ticks(&HE4) = 3: instruction(&HE4) = INS_CPX: addrmode(&HE4) = ADR_ZP
      Ticks(&HE5) = 3: instruction(&HE5) = INS_SBC: addrmode(&HE5) = ADR_ZP
      Ticks(&HE6) = 5: instruction(&HE6) = INS_INC: addrmode(&HE6) = ADR_ZP
      Ticks(&HE7) = 2: instruction(&HE7) = INS_NOP: addrmode(&HE7) = ADR_IMP
      Ticks(&HE8) = 2: instruction(&HE8) = INS_INX: addrmode(&HE8) = ADR_IMP
      Ticks(&HE9) = 3: instruction(&HE9) = INS_SBC: addrmode(&HE9) = ADR_IMM
      Ticks(&HEA) = 2: instruction(&HEA) = INS_NOP: addrmode(&HEA) = ADR_IMP
      Ticks(&HEB) = 2: instruction(&HEB) = INS_NOP: addrmode(&HEB) = ADR_IMP
      Ticks(&HEC) = 4: instruction(&HEC) = INS_CPX: addrmode(&HEC) = ADR_ABS
      Ticks(&HED) = 4: instruction(&HED) = INS_SBC: addrmode(&HED) = ADR_ABS
      Ticks(&HEE) = 6: instruction(&HEE) = INS_INC: addrmode(&HEE) = ADR_ABS
      Ticks(&HEF) = 2: instruction(&HEF) = INS_NOP: addrmode(&HEF) = ADR_IMP
      Ticks(&HF0) = 2: instruction(&HF0) = INS_BEQ: addrmode(&HF0) = ADR_REL
      Ticks(&HF1) = 5: instruction(&HF1) = INS_SBC: addrmode(&HF1) = ADR_INDY
      Ticks(&HF2) = 3: instruction(&HF2) = INS_SBC: addrmode(&HF2) = ADR_INDZP
      Ticks(&HF3) = 2: instruction(&HF3) = INS_NOP: addrmode(&HF3) = ADR_IMP
      Ticks(&HF4) = 2: instruction(&HF4) = INS_NOP: addrmode(&HF4) = ADR_IMP
      Ticks(&HF5) = 4: instruction(&HF5) = INS_SBC: addrmode(&HF5) = ADR_ZPX
      Ticks(&HF6) = 6: instruction(&HF6) = INS_INC: addrmode(&HF6) = ADR_ZPX
      Ticks(&HF7) = 2: instruction(&HF7) = INS_NOP: addrmode(&HF7) = ADR_IMP
      Ticks(&HF8) = 2: instruction(&HF8) = INS_SED: addrmode(&HF8) = ADR_IMP
      Ticks(&HF9) = 4: instruction(&HF9) = INS_SBC: addrmode(&HF9) = ADR_ABSY
      Ticks(&HFA) = 4: instruction(&HFA) = INS_PLX: addrmode(&HFA) = ADR_IMP
      Ticks(&HFB) = 2: instruction(&HFB) = INS_NOP: addrmode(&HFB) = ADR_IMP
      Ticks(&HFC) = 2: instruction(&HFC) = INS_NOP: addrmode(&HFC) = ADR_IMP
      Ticks(&HFD) = 4: instruction(&HFD) = INS_SBC: addrmode(&HFD) = ADR_ABSX
      Ticks(&HFE) = 7: instruction(&HFE) = INS_INC: addrmode(&HFE) = ADR_ABSX
      Ticks(&HFF) = 2: instruction(&HFF) = INS_NOP: addrmode(&HFF) = ADR_IMP
End Function
