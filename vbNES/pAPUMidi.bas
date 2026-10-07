Attribute VB_Name = "pAPUMidi"
'Partially emulates the Audio Processing Unit (APU). Sometimes it works.

Option Explicit
DefLng A-Z

Private tones(3)
Private volume(3)
Private lastFrame(3)
Private stopTones(3)


Public doSound As Boolean

Private vlengths(31) As Long

Public ChannelWrite(3) As Boolean


Private Sub fillArray(a() As Long, ParamArray b() As Variant)
    Dim i
    For i = 0 To UBound(a)
        a(i) = b(i)
    Next i
End Sub

Public Sub pAPUinit()
    'Lookup table used by nester.
    fillArray vlengths, 5, 127, 10, 1, 19, 2, 40, 3, 80, 4, 30, 5, 7, 6, 13, 7, 6, 8, 12, 9, 24, 10, 48, 11, 96, 12, 36, 13, 8, 14, 16, 15
    
    ' known problems:
    '   no instantaneous volume with midi. Really short notes can't be generated and heard.
    '   couldn't find an adequate noise generator.
    '   can't control the shape of the square wave.
    
    SelectInstrument 0, 80 'Square wave
    SelectInstrument 1, 80 'Square wave
    SelectInstrument 2, 74 'Triangle wave. Used recorder (like a flute
    
    SelectInstrument 3, 127 'Noise. Used gunshot. Sometimes inadequate.
End Sub


Public Sub playTone(channel, tone, v)
    If tone <> tones(channel) Or v < volume(channel) - 3 Or v > volume(channel) Or v = 0 Then
        If tones(channel) <> 0 Then
            ToneOff channel, tones(channel)
            tones(channel) = 0
            volume(channel) = 0
        End If
        If doSound And tone > 0 And tone <= 127 And v > 0 Then
            volume(channel) = v
            tones(channel) = tone
            ToneOn channel, tone, v * 8
        End If
    End If
End Sub

Public Sub stopTone(channel)
    If tones(channel) <> 0 Then
        stopTones(channel) = tones(channel)
        tones(channel) = 0
        volume(channel) = 0
    End If
End Sub

Public Sub ReallyStopTones()
    Dim i
    For i = 0 To 3
        If stopTones(i) <> 0 And stopTones(i) <> tones(i) Then
            ToneOff i, stopTones(i)
            stopTones(i) = 0
        End If
    Next i
End Sub


'Calculates a midi tone given an nes frequency.
'Frequency passed is actual interval in 1/65536's of a second (I hope). nope.
Public Function getTone(ByVal freq) As Long
    If freq <= 0 Then Exit Function
    
    Dim t As Long
    
    ' Hopefully this is correct. Convert period to frequency
'    freq = 65536 / freq
    ' wow. I was way off. Almost an entire octave. -DF
    freq = 111861 / (freq + 1)
    
    'convert to frequency to closest note
    t = CLng(Log(freq / 8.176) * 17.31234)
    
    getTone = t
End Function

Public Sub PlayRect(ch)
    Dim f, l, v
    If SoundCtrl And pow2(ch) Then
        v = (Sound(ch * 4 + 0) And 15) 'Get volume
        l = vlengths(Sound(ch * 4 + 3) \ 8) 'Get length
        If v > 0 Then
            f = Sound(ch * 4 + 2) + (Sound(ch * 4 + 3) And 7) * 256 'Get frequency
            If f > 1 Then
                If ChannelWrite(ch) Then 'Ensures that a note doesn't replay unless memory written
                    ChannelWrite(ch) = False
                    lastFrame(ch) = Frames + l
                    playTone ch, getTone(f), v
                End If
            Else
                stopTone ch
            End If
        Else
            stopTone ch
        End If
    Else
        ChannelWrite(ch) = True
        stopTone ch
    End If
    If Frames >= lastFrame(ch) Then
        stopTone ch
    End If
End Sub

Public Sub PlayTriangle(ch)
    Dim f, l, v
    If SoundCtrl And pow2(ch) Then
        v = 6 'triangle
        l = vlengths(Sound(ch * 4 + 3) \ 8)
        If v > 0 Then
            f = Sound(ch * 4 + 2) + (Sound(ch * 4 + 3) And 7) * 256
            If f > 1 Then
                If ChannelWrite(ch) Then
                    ChannelWrite(ch) = False
                    lastFrame(ch) = Frames + l
                    playTone ch, getTone(f), v
                End If
            Else
                stopTone ch
            End If
        Else
            stopTone ch
        End If
    Else
        ChannelWrite(ch) = True
        stopTone ch
    End If
    If Frames >= lastFrame(ch) Then
        stopTone ch
    End If
End Sub

Public Sub PlayNoise(ch)
    Dim f, l, v
    If SoundCtrl And pow2(ch) Then
        v = 6
        l = vlengths(Sound(ch * 4 + 3) \ 8)
        If v > 0 Then
            f = (Sound(ch * 4 + 2) And 15) * 128
            If f > 1 Then
                If ChannelWrite(ch) Then
                    ChannelWrite(ch) = False
                    lastFrame(ch) = Frames + l
                    playTone ch, getTone(f), v
                End If
            Else
                stopTone ch
            End If
        Else
            stopTone ch
        End If
    Else
        ChannelWrite(ch) = True
        stopTone ch
    End If
    If Frames >= lastFrame(ch) Then
        stopTone ch
    End If
End Sub

Public Sub StopSound()
    stopTone 0
    stopTone 1
    stopTone 2
    stopTone 3
    ReallyStopTones
End Sub



Public Sub updateSounds()
    If doSound Then
        ReallyStopTones
        PlayRect 0
        PlayRect 1
        PlayTriangle 2
        PlayNoise 3
    Else
        StopSound
    End If
End Sub


