Attribute VB_Name = "MidiLib"
'Written by David Finch
Option Explicit
DefLng A-Z

Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long


'#define NoteOnCmd       0x90
'#define NoteOffCmd      0x80
'#define PgmChngCmd      0xC0
'#define ControlCmd      0xB0
'#define PolyPressCmd    0xA0
'#define ChanPressCmd    0xD0
'#define PchWheelCmd     0xE0
'#define SysExCmd        0xF0

Private mdh As Long
Private midiOpened As Boolean

Public Sub MidiOpen()
    If midiOutOpen(mdh, 0, 0, 0, 0) Then
        On Error Resume Next
        Open App.Path + "\pmidi.tmp" For Input As #4
        If Err.Number = 0 Then
            Dim i As Long
            Dim s As String
            Line Input #4, s
            Input #4, i
            midiOutClose i
            Close #4
        End If
        If midiOutOpen(mdh, 0, 0, 0, 0) Then
            frmNES.lblStatus.Caption = "Cannot open midi. Either you have no sound card or another program is hogging the midi."
            Exit Sub
        End If
    End If
    'allows for closing midi after a crash
    Open App.Path + "\pmidi.tmp" For Output As #4
    Print #4, "Previous midi handle: (used in case it crashed last time)"
    Print #4, mdh
    Close #4
    midiOpened = True
End Sub

Public Sub MidiClose()
    If midiOpened Then
        midiOutClose mdh
        midiOpened = False
    End If
End Sub

Public Sub SelectInstrument(ByVal channel As Long, ByVal patch As Long)
    If midiOpened Then midiOutShortMsg mdh, &HC0 Or patch * 256 Or channel
End Sub

Public Sub ToneOn(ByVal channel As Long, ByVal tone As Long, ByVal volume As Long)
    If midiOpened Then
        If tone < 0 Then tone = 0
        If tone > 127 Then tone = 127
        midiOutShortMsg mdh, &H90 Or tone * 256 Or channel Or volume * 65536
    End If
End Sub

Public Sub ToneOff(ByVal channel As Long, ByVal tone As Long)
    If midiOpened Then
        If tone < 0 Then tone = 0
        If tone > 127 Then tone = 127
        midiOutShortMsg mdh, &H80 Or tone * 256 Or channel
    End If
End Sub



#If 0 Then
MIDI instrument list. Ripped off some website I've forgotten which

0=Acoustic Grand Piano
1=Bright Acoustic Piano
2=Electric Grand Piano
3=Honky-tonk Piano
4=Rhodes Piano
5=Chorus Piano
6=Harpsi -chord
7=Clavinet
8=Celesta
9=Glocken -spiel
10=Music Box
11=Vibra -phone
12=Marimba
13=Xylo-phone
14=Tubular Bells
15=Dulcimer
16=Hammond Organ
17=Percuss. Organ
18=Rock Organ
19=Church Organ
20=Reed Organ
21=Accordion
22=Harmonica
23=Tango Accordion
24=Acoustic Guitar (nylon)
25=Acoustic Guitar (steel)
26=Electric Guitar (jazz)
27=Electric Guitar (clean)
28=Electric Guitar (muted)
29=Overdriven Guitar
30=Distortion Guitar
31=Guitar Harmonics
32=Acoustic Bass
33=Electric Bass (finger)
34=Electric Bass (pick)
35=Fretless Bass
36=Slap Bass 1
37=Slap Bass 2
38=Synth Bass 1
39=Synth Bass 2
40=Violin
41=Viola
42=Cello
43=Contra Bass
44=Tremolo Strings
45=Pizzicato Strings
46=Orchestral Harp
47=Timpani
48=String Ensemble 1
49=String Ensemble 2
50=Synth Strings 1
51=Synth Strings 2
52=Choir Aahs
53=Voice Oohs
54=Synth Voice
55=Orchestra Hit
56=Trumpet
57=Trombone
58=Tuba
59=Muted Trumpet
60=French Horn
61=Brass Section
62=Synth Brass 1
63=Synth Brass 2
64=Soprano Sax
65=Alto Sax
66=Tenor Sax
67=Baritone Sax
68=Oboe
69=English Horn
70=Bassoon
71=Clarinet
72=Piccolo
73=Flute
74=Recorder
75=Pan Flute
76=Bottle Blow
77=Shaku
78=Whistle
79=Ocarina
80=Lead 1 (square)
81=Lead 2 (saw tooth)
82=Lead 3 (calliope lead)
83=Lead 4 (chiff lead)
84=Lead 5 (charang)
85=Lead 6 (voice)
86=Lead 7 (fifths)
87=Lead 8 (bass + lead)
88=Pad 1 (new age)
89=Pad 2 (warm)
90=Pad 3 (poly synth)
91=Pad 4 (choir)
92=Pad 5 (bowed)
93=Pad 6 (metallic)
94=Pad 7 (halo)
95=Pad 8 (sweep)
96=FX 1 (rain)
97=FX 2 (sound track)
98=FX 3 (crystal)
99=FX 4 (atmo - sphere)
100=FX 5 (bright)
101=FX 6 (goblins)
102=FX 7 (echoes)
103=FX 8 (sci-fi)
104=Sitar
105=Banjo
106=Shamisen
107=Koto
108=Kalimba
109=Bagpipe
110=Fiddle
111=Shanai
112=Tinkle Bell
113=Agogo
114=Steel Drums
115=Wood block
116=Taiko Drum
117=Melodic Tom
118=Synth Drum
119=Reverse Cymbal
120=Guitar Fret Noise
121=Breath Noise
122=Seashore
123=Bird Tweet
124=Telephone Ring
125=Helicopter
126=Applause
127=Gunshot
#End If
