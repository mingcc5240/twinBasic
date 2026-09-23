Attribute VB_Name = "Sound"
'Attribute VB_Name = "Sound"
' =========================================================================
' BasicBoy Emulator Sound Module for Visual Basic 6.0
' 16-bit Stereo Synthesizer with Circular Ring Buffer
' =========================================================================

Option Explicit

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    reserved As Long
End Type

Private Const WHDR_DONE As Long = &H1
Private Const WHDR_PREPARED As Long = &H2

Private Declare Function waveOutOpen Lib "winmm.dll" ( _
    ByRef phwo As Long, ByVal uDeviceID As Long, ByRef pwfx As WAVEFORMATEX, _
    ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long

Private Declare Function waveOutPrepareHeader Lib "winmm.dll" ( _
    ByVal hwo As Long, ByRef pwh As WAVEHDR, ByVal cbwh As Long) As Long

Private Declare Function waveOutWrite Lib "winmm.dll" ( _
    ByVal hwo As Long, ByRef pwh As WAVEHDR, ByVal cbwh As Long) As Long

Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" ( _
    ByVal hwo As Long, ByRef pwh As WAVEHDR, ByVal cbwh As Long) As Long

Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hwo As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hwo As Long) As Long

Private Const SAMPLE_RATE As Long = 44100
' 1개 버퍼 블록 크기: 1024 샘플 (약 23.2ms)
Private Const BUFFER_SAMPLES As Long = 3 * 1024 + 256 '1024
' 16비트 Stereo = 샘플당 4바이트 (Left 2Byte + Right 2Byte)
Private Const BUFFER_BYTES As Long = BUFFER_SAMPLES * 4

' ★ 링 버퍼 슬롯 개수 (8개 슬롯 = 약 185ms 오디오 큐 윈도우 제공)
Private Const NUM_BUFFERS As Long = 2

Private hWaveOut As Long
Private WaveHeader(0 To NUM_BUFFERS - 1) As WAVEHDR
Private Buffer(0 To NUM_BUFFERS - 1, 0 To (BUFFER_SAMPLES * 2) - 1) As Integer
Private CurrentBufferIdx As Long

' --- APU Registers & States ---
Private DutyPatterns(0 To 3, 0 To 7) As Byte

' Channel 1
Private Ch1_Enable As Boolean
Private Ch1_Freq As Long, Ch1_Vol As Byte, Ch1_DutyIdx As Long, Ch1_Phase As Double, Ch1_RawFreq As Long

' Channel 2
Private Ch2_Enable As Boolean
Private Ch2_Freq As Long, Ch2_Vol As Byte, Ch2_DutyIdx As Long, Ch2_Phase As Double, Ch2_RawFreq As Long

' Channel 3 (Wave)
Private Ch3_Enable As Boolean
Private Ch3_Freq As Long, Ch3_VolShift As Byte, Ch3_Phase As Single, Ch3_RawFreq As Long
Private WaveRAM(0 To 15) As Byte

' Channel 4 (Noise)
Private Ch4_Enable As Boolean
Private Ch4_Vol As Byte, Ch4_Freq As Single, Ch4_Phase As Single, Ch4_LFSR As Long, Ch4_Step7 As Boolean

' Panning & Master Control (NR50, NR51, NR52)
Private SoundEnabled As Boolean
Private NR50_VolL As Byte, NR50_VolR As Byte
Private NR51_Pan As Byte

Public Sub InitSound()
    Dim wfx As WAVEFORMATEX
    Dim k As Long
    
    With wfx
        .wFormatTag = 1          ' PCM
        .nChannels = 2            ' Stereo
        .nSamplesPerSec = SAMPLE_RATE
        .wBitsPerSample = 16      ' 16-bit
        .nBlockAlign = 4          ' 2채널 * 2바이트
        .nAvgBytesPerSec = SAMPLE_RATE * 4
        .cbSize = 0
    End With

    If hWaveOut <> 0 Then CloseSound

    If waveOutOpen(hWaveOut, -1, wfx, 0, 0, 0) = 0 Then
        ' 링 버퍼 헤더 초기화
        For k = 0 To NUM_BUFFERS - 1
            With WaveHeader(k)
                .lpData = VarPtr(Buffer(k, 0))
                .dwBufferLength = BUFFER_BYTES
                .dwBytesRecorded = 0
                .dwUser = 0
                .dwFlags = WHDR_DONE ' 초기 상태를 쓰기 가능한 상태(완료됨)로 마킹
                .dwLoops = 0
            End With
        Next k
        
        CurrentBufferIdx = 0
        SoundEnabled = True
        NR51_Pan = &HFF
        NR50_VolL = 7: NR50_VolR = 7
        Ch4_LFSR = &H7FFF
        
        ' GB APU 하드웨어 정밀 Duty Patterns
        DutyPatterns(0, 0) = 0: DutyPatterns(0, 1) = 0: DutyPatterns(0, 2) = 0: DutyPatterns(0, 3) = 0
        DutyPatterns(0, 4) = 0: DutyPatterns(0, 5) = 0: DutyPatterns(0, 6) = 0: DutyPatterns(0, 7) = 1
        
        DutyPatterns(1, 0) = 1: DutyPatterns(1, 1) = 0: DutyPatterns(1, 2) = 0: DutyPatterns(1, 3) = 0
        DutyPatterns(1, 4) = 0: DutyPatterns(1, 5) = 0: DutyPatterns(1, 6) = 0: DutyPatterns(1, 7) = 1
        
        DutyPatterns(2, 0) = 1: DutyPatterns(2, 1) = 0: DutyPatterns(2, 2) = 0: DutyPatterns(2, 3) = 0
        DutyPatterns(2, 4) = 0: DutyPatterns(2, 5) = 1: DutyPatterns(2, 6) = 1: DutyPatterns(2, 7) = 1
        
        DutyPatterns(3, 0) = 0: DutyPatterns(3, 1) = 1: DutyPatterns(3, 2) = 1: DutyPatterns(3, 3) = 1
        DutyPatterns(3, 4) = 1: DutyPatterns(3, 5) = 1: DutyPatterns(3, 6) = 1: DutyPatterns(3, 7) = 0
    End If
End Sub

Public Sub CloseSound()
    Dim k As Long
    If hWaveOut <> 0 Then
        Call waveOutReset(hWaveOut)
        For k = 0 To NUM_BUFFERS - 1
            If (WaveHeader(k).dwFlags And WHDR_PREPARED) <> 0 Then
                Call waveOutUnprepareHeader(hWaveOut, WaveHeader(k), LenB(WaveHeader(k)))
            End If
        Next k
        Call waveOutClose(hWaveOut)
        hWaveOut = 0
    End If
End Sub

Public Sub updatesnd(clc As Long)
    If Not SoundEnabled Or hWaveOut = 0 Then Exit Sub

    ' ★ [링 버퍼 슬롯 가용성 검증]
    ' 타깃 버퍼가 아직 오디오 장치에서 재생 중(WHDR_DONE 플래그 부재)이면 쓰기를 건너뜁니다.
    If (WaveHeader(CurrentBufferIdx).dwFlags And WHDR_DONE) = 0 Then Exit Sub

    ' 이전 재생 완료 헤더 정리
    If (WaveHeader(CurrentBufferIdx).dwFlags And WHDR_PREPARED) <> 0 Then
        Call waveOutUnprepareHeader(hWaveOut, WaveHeader(CurrentBufferIdx), LenB(WaveHeader(CurrentBufferIdx)))
    End If

    Dim i As Long
    Dim ch1Sig As Double, ch2Sig As Double, ch3Sig As Double, ch4Sig As Double
    Dim outL As Double, outR As Double
    Dim sampleRateF As Double: sampleRateF = CDbl(SAMPLE_RATE)
    Dim bufIdx As Long: bufIdx = 0

    For i = 0 To BUFFER_SAMPLES - 1
        ch1Sig = 0#: ch2Sig = 0#: ch3Sig = 0#: ch4Sig = 0#

        ' --- Channel 1 (Square 1) ---
        If Ch1_Enable And (Ch1_Freq > 0) And (Ch1_Vol > 0) Then
            Ch1_Phase = Ch1_Phase + (CDbl(Ch1_Freq) / sampleRateF)
            If Ch1_Phase >= 1# Then Ch1_Phase = Ch1_Phase - Int(Ch1_Phase)
            
            Dim step1 As Long
            step1 = Int(Ch1_Phase * 8#) And 7
            
            If DutyPatterns(Ch1_DutyIdx, step1) <> 0 Then
                ch1Sig = CDbl(Ch1_Vol) * 400#
            Else
                ch1Sig = -CDbl(Ch1_Vol) * 400#
            End If
        End If

        ' --- Channel 2 (Square 2) ---
        If Ch2_Enable And (Ch2_Freq > 0) And (Ch2_Vol > 0) Then
            Ch2_Phase = Ch2_Phase + (CDbl(Ch2_Freq) / sampleRateF)
            If Ch2_Phase >= 1# Then Ch2_Phase = Ch2_Phase - Int(Ch2_Phase)
            
            Dim step2 As Long
            step2 = Int(Ch2_Phase * 8#) And 7
            
            If DutyPatterns(Ch2_DutyIdx, step2) <> 0 Then
                ch2Sig = CDbl(Ch2_Vol) * 400#
            Else
                ch2Sig = -CDbl(Ch2_Vol) * 400#
            End If
        End If

        ' --- Channel 3 (Wave) ---
        If Ch3_Enable And (Ch3_Freq > 0) And (Ch3_VolShift > 0) Then
            Ch3_Phase = Ch3_Phase + (CDbl(Ch3_Freq) / sampleRateF)
            If Ch3_Phase >= 1# Then Ch3_Phase = Ch3_Phase - Int(Ch3_Phase)
            
            Dim wavePos As Double
            wavePos = Ch3_Phase * 32#
            
            Dim idx0 As Long, idx1 As Long
            Dim frac As Double
            idx0 = Int(wavePos) And 31
            idx1 = (idx0 + 1) And 31
            frac = wavePos - CDbl(Int(wavePos))

            Dim s0 As Double, s1 As Double
            s0 = CDbl(GetWaveSample(idx0)) - 7.5
            s1 = CDbl(GetWaveSample(idx1)) - 7.5
            
            ch3Sig = s0 + (s1 - s0) * frac

            Select Case Ch3_VolShift
                Case 1: ch3Sig = ch3Sig * 300#
                Case 2: ch3Sig = ch3Sig * 150#
                Case 3: ch3Sig = ch3Sig * 75#
                Case Else: ch3Sig = 0#
            End Select
        End If

        ' --- Channel 4 (Noise) ---
        If Ch4_Enable And (Ch4_Freq > 0#) Then
            Ch4_Phase = Ch4_Phase + (CDbl(Ch4_Freq) / sampleRateF)
            
            If Ch4_Phase >= 1# Then
                Dim shifts As Long
                shifts = Int(Ch4_Phase)
                Ch4_Phase = Ch4_Phase - CDbl(shifts)
                
                Dim sCount As Long
                For sCount = 1 To shifts
                    Dim bit0 As Long, bit1 As Long, resultBit As Long
                    bit0 = Ch4_LFSR And 1
                    bit1 = (Ch4_LFSR \ 2) And 1
                    resultBit = bit0 Xor bit1
                    
                    Ch4_LFSR = (Ch4_LFSR \ 2) And &H3FFF
                    Ch4_LFSR = Ch4_LFSR Or (resultBit * 16384)
                    
                    If Ch4_Step7 Then
                        Ch4_LFSR = (Ch4_LFSR And &HFFBF) Or (resultBit * 64)
                    End If
                Next sCount
            End If
            
            If (Ch4_LFSR And 1) = 0 Then
                ch4Sig = CDbl(Ch4_Vol) * 100#
            Else
                ch4Sig = -CDbl(Ch4_Vol) * 100#
            End If
        End If

        ' --- Stereo Panning Matrix (NR51) ---
        outL = 0#: outR = 0#

        If (NR51_Pan And &H10) <> 0 Then outL = outL + ch1Sig
        If (NR51_Pan And &H20) <> 0 Then outL = outL + ch2Sig
        If (NR51_Pan And &H40) <> 0 Then outL = outL + ch3Sig
        If (NR51_Pan And &H80) <> 0 Then outL = outL + ch4Sig

        If (NR51_Pan And &H1) <> 0 Then outR = outR + ch1Sig
        If (NR51_Pan And &H2) <> 0 Then outR = outR + ch2Sig
        If (NR51_Pan And &H4) <> 0 Then outR = outR + ch3Sig
        If (NR51_Pan And &H8) <> 0 Then outR = outR + ch4Sig

        ' Master Volume Apply
        outL = outL * (CDbl(NR50_VolL + 1) / 8#)
        outR = outR * (CDbl(NR50_VolR + 1) / 8#)

        ' 16-bit PCM Clamping
        If outL > 32767# Then outL = 32767#
        If outL < -32768# Then outL = -32768#
        If outR > 32767# Then outR = 32767#
        If outR < -32768# Then outR = -32768#

        Buffer(CurrentBufferIdx, bufIdx) = CInt(outL)
        Buffer(CurrentBufferIdx, bufIdx + 1) = CInt(outR)
        bufIdx = bufIdx + 2
    Next i

    ' ★ [링 큐 등록 및 다음 슬롯 이동]
    Call waveOutPrepareHeader(hWaveOut, WaveHeader(CurrentBufferIdx), LenB(WaveHeader(CurrentBufferIdx)))
    Call waveOutWrite(hWaveOut, WaveHeader(CurrentBufferIdx), LenB(WaveHeader(CurrentBufferIdx)))
    CurrentBufferIdx = (CurrentBufferIdx + 1) Mod NUM_BUFFERS
End Sub

Private Function GetWaveSample(ByVal sampleIdx As Long) As Single
    Dim byteVal As Byte
    byteVal = WaveRAM(sampleIdx \ 2)
    If (sampleIdx Mod 2) = 0 Then
        GetWaveSample = CSng((byteVal \ 16) And &HF)
    Else
        GetWaveSample = CSng(byteVal And &HF)
    End If
End Function

' =========================================================================
' Sound Registers Handlers
' =========================================================================

Public Sub setNR10(Val As Long): End Sub

Public Sub setNR11(Val As Long)
    Ch1_DutyIdx = (Val \ 64) And &H3
End Sub

Public Sub setNR12(Val As Long)
    Ch1_Vol = (Val \ 16) And &HF
    If Ch1_Vol = 0 Then Ch1_Enable = False
End Sub

Public Sub setNR13(Val As Long)
    Ch1_RawFreq = (Ch1_RawFreq And &H700) Or (Val And &HFF)
    If Ch1_RawFreq < 2048 Then Ch1_Freq = 131072 \ (2048 - Ch1_RawFreq)
End Sub

Public Sub setNR14(Val As Long)
    Ch1_RawFreq = (Ch1_RawFreq And &HFF) Or ((Val And &H7) * 256)
    If Ch1_RawFreq < 2048 Then Ch1_Freq = 131072 \ (2048 - Ch1_RawFreq)
    If (Val And &H80) <> 0 Then
        Ch1_Enable = True
        Ch1_Phase = 0#
    End If
End Sub

Public Sub setNR21(Val As Long)
    Ch2_DutyIdx = (Val \ 64) And &H3
End Sub

Public Sub setNR22(Val As Long)
    Ch2_Vol = (Val \ 16) And &HF
    If Ch2_Vol = 0 Then Ch2_Enable = False
End Sub

Public Sub setNR23(Val As Long)
    Ch2_RawFreq = (Ch2_RawFreq And &H700) Or (Val And &HFF)
    If Ch2_RawFreq < 2048 Then Ch2_Freq = 131072 \ (2048 - Ch2_RawFreq)
End Sub

Public Sub setNR24(Val As Long)
    Ch2_RawFreq = (Ch2_RawFreq And &HFF) Or ((Val And &H7) * 256)
    If Ch2_RawFreq < 2048 Then Ch2_Freq = 131072 \ (2048 - Ch2_RawFreq)
    If (Val And &H80) <> 0 Then
        Ch2_Enable = True
        Ch2_Phase = 0#
    End If
End Sub

Public Sub setNR30(Val As Long)
    Ch3_Enable = ((Val And &H80) <> 0)
End Sub

Public Sub setNR31(Val As Long): End Sub

Public Sub setNR32(Val As Long)
    Select Case ((Val \ 32) And &H3)
        Case 0: Ch3_VolShift = 0
        Case 1: Ch3_VolShift = 1
        Case 2: Ch3_VolShift = 2
        Case 3: Ch3_VolShift = 3
    End Select
End Sub

Public Sub setNR33(Val As Long)
    Ch3_RawFreq = (Ch3_RawFreq And &H700) Or (Val And &HFF)
    If Ch3_RawFreq < 2048 Then Ch3_Freq = 65536 \ (2048 - Ch3_RawFreq)
End Sub

Public Sub setNR34(Val As Long)
    Ch3_RawFreq = (Ch3_RawFreq And &HFF) Or ((Val And &H7) * 256)
    If Ch3_RawFreq < 2048 Then Ch3_Freq = 65536 \ (2048 - Ch3_RawFreq)
    If (Val And &H80) <> 0 Then Ch3_Enable = True: Ch3_Phase = 0!
End Sub

Public Sub setWaveRAM(ByVal Index As Long, ByVal Val As Byte)
    If Index >= 0 And Index <= 15 Then WaveRAM(Index) = Val
End Sub

Public Sub setNR41(Val As Long): End Sub

Public Sub setNR42(Val As Long)
    Ch4_Vol = (Val \ 16) And &HF
    If Ch4_Vol = 0 Then Ch4_Enable = False
End Sub

Public Sub setNR43(Val As Long)
    Dim s As Long: s = (Val \ 16) And &HF
    Dim r As Long: r = Val And &H7
    Ch4_Step7 = ((Val And &H8) <> 0)
    
    Dim d As Double
    If r = 0 Then d = 0.5 Else d = CDbl(r)
    
    Ch4_Freq = 524288! / (d * CDbl(2 ^ (s + 1)))
End Sub

Public Sub setNR44(Val As Long)
    If (Val And &H80) <> 0 Then
        Ch4_Enable = True
        Ch4_LFSR = &H7FFF
    End If
End Sub

' Stereo Routing / Volume Control
Public Sub setNR50(Val As Long)
    NR50_VolR = Val And &H7
    NR50_VolL = (Val \ 16) And &H7
End Sub

Public Sub setNR51(Val As Long)
    NR51_Pan = Val And &HFF
End Sub

Public Sub setNR52(Val As Long)
    SoundEnabled = ((Val And &H80) <> 0)
End Sub

