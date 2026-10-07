VERSION 5.00
Begin VB.Form frmNES 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "basicNES"
   ClientHeight    =   4365
   ClientLeft      =   3495
   ClientTop       =   2280
   ClientWidth     =   4440
   Icon            =   "frmNES.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   480
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Interval        =   467
      Left            =   0
      Top             =   3960
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   210
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   264
      TabIndex        =   0
      Top             =   240
      Width           =   3990
   End
   Begin VB.Label lbMirror 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   240
   End
   Begin VB.Label lbSpeed 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   1320
      TabIndex        =   2
      Top             =   4005
      Width           =   3030
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   4050
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Load ROM"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileFree 
         Caption         =   "&Free ROM"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuHyphen20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRomInfo 
         Caption         =   "&ROM Info"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEmulation 
      Caption         =   "&Emulation"
      Begin VB.Menu mnuEmuConfgKeys 
         Caption         =   "Configure Keys"
      End
      Begin VB.Menu mfast 
         Caption         =   "Make it FAST (unsafe)"
      End
      Begin VB.Menu mbarfgjsdhfg 
         Caption         =   "-"
      End
      Begin VB.Menu msaveslt 
         Caption         =   "Save slot"
         Begin VB.Menu msaveslots 
            Caption         =   "0"
            Index           =   0
         End
         Begin VB.Menu msaveslots 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu msaveslots 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu msaveslots 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu msaveslots 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu msaveslots 
            Caption         =   "5"
            Index           =   5
         End
         Begin VB.Menu msaveslots 
            Caption         =   "6"
            Index           =   6
         End
         Begin VB.Menu msaveslots 
            Caption         =   "7"
            Index           =   7
         End
         Begin VB.Menu msaveslots 
            Caption         =   "8"
            Index           =   8
         End
         Begin VB.Menu msaveslots 
            Caption         =   "9"
            Index           =   9
         End
      End
      Begin VB.Menu msave 
         Caption         =   "Save state"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mrestore 
         Caption         =   "Restore state"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuMovies 
         Caption         =   "&Record NES Movie"
         Begin VB.Menu mnuMvSlot 
            Caption         =   "Slot Number"
            Begin VB.Menu mnu0 
               Caption         =   "0"
               Index           =   0
            End
            Begin VB.Menu mnu0 
               Caption         =   "1"
               Index           =   1
            End
            Begin VB.Menu mnu0 
               Caption         =   "2"
               Index           =   2
            End
            Begin VB.Menu mnu0 
               Caption         =   "3"
               Index           =   3
            End
            Begin VB.Menu mnu0 
               Caption         =   "4"
               Index           =   4
            End
            Begin VB.Menu mnu0 
               Caption         =   "5"
               Index           =   5
            End
            Begin VB.Menu mnu0 
               Caption         =   "6"
               Index           =   6
            End
            Begin VB.Menu mnu0 
               Caption         =   "7"
               Index           =   7
            End
            Begin VB.Menu mnu0 
               Caption         =   "8"
               Index           =   8
            End
            Begin VB.Menu mnu0 
               Caption         =   "9"
               Index           =   9
            End
         End
         Begin VB.Menu mnuStartRecord 
            Caption         =   "&Start"
         End
         Begin VB.Menu mnuPlayMovie 
            Caption         =   "&Play"
         End
      End
      Begin VB.Menu mbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFrameSkip 
         Caption         =   "Frame Skip"
         Begin VB.Menu mnuFS 
            Caption         =   "Draw All Frames"
            Index           =   0
         End
         Begin VB.Menu mnuFS 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu mnuFS 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu mnuFS 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu hyphen3457 
            Caption         =   "-"
         End
         Begin VB.Menu mAutoSpeed 
            Caption         =   "Limit to 60 fps"
         End
      End
      Begin VB.Menu mexec 
         Caption         =   "Execute %"
         Begin VB.Menu mexecv 
            Caption         =   "Auto adjust"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mexecv 
            Caption         =   "150% = Overclocked"
            Index           =   1
         End
         Begin VB.Menu mexecv 
            Caption         =   "100% = safest"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mexecv 
            Caption         =   "75% = safer"
            Index           =   3
         End
         Begin VB.Menu mexecv 
            Caption         =   "50% = faster"
            Index           =   4
         End
         Begin VB.Menu mexecv 
            Caption         =   "25% = fastest"
            Index           =   5
         End
         Begin VB.Menu mbar234 
            Caption         =   "-"
         End
         Begin VB.Menu mIdle 
            Caption         =   "Idle detection"
         End
      End
      Begin VB.Menu mnuCPUPause 
         Caption         =   "&Pause CPU"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEmuReset 
         Caption         =   "&Reset CPU"
         Shortcut        =   ^R
      End
      Begin VB.Menu mMute 
         Caption         =   "Mute sound"
      End
      Begin VB.Menu mbarg348 
         Caption         =   "-"
      End
      Begin VB.Menu mtiled 
         Caption         =   "Tilebased"
      End
      Begin VB.Menu mNewScroll 
         Caption         =   "New scroll code"
         Checked         =   -1  'True
         Shortcut        =   {F11}
      End
      Begin VB.Menu mPalette 
         Caption         =   "Color palette"
         Begin VB.Menu mSelPalette 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mMotionBlur 
         Caption         =   "Motion blur"
         Shortcut        =   ^M
      End
      Begin VB.Menu mSmoothTop 
         Caption         =   "Zoom quality"
         Begin VB.Menu mSmooth 
            Caption         =   "Normal, fast"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mSmooth 
            Caption         =   "Interpolated, slow"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mSmooth 
            Caption         =   "Edge-Finding, slow"
            Enabled         =   0   'False
            Index           =   2
         End
      End
   End
   Begin VB.Menu mzm 
      Caption         =   "&Zoom"
      Begin VB.Menu mZoom 
         Caption         =   "1x"
         Index           =   0
      End
      Begin VB.Menu mZoom 
         Caption         =   "2x"
         Index           =   1
      End
      Begin VB.Menu mZoom 
         Caption         =   "3x"
         Index           =   2
      End
      Begin VB.Menu mnuFull 
         Caption         =   "Full Screen"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmNES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

DefLng A-Z

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private SlotIndex As Long 'for the save-state

Private StatusTimeout As Long 'countdown to clear statusbar
'stores b to a and returns b
Private Function st(Aa As Long, b As Long) As Long
        Aa = b
    st = b
End Function
'Takes a path and returns the filename
Private Function extractFileName(fi As String) As String
    Dim s As String
    Dim i As Long
    s = fi
    Do While InStr(s, "\")
        s = Mid$(s, InStr(s, "\") + 1)
    Loop
    If InStr(s, ".") Then
        s = Left$(s, InStr(s, ".") - 1)
    End If
    extractFileName = s
End Function
Private Function open_file() As String
    Dim lReturn As Long
    Dim sFilter As String
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = Me.hwnd
    OpenFile.hInstance = App.hInstance
    sFilter = "NES ROMs (*.nes)" & Chr(0) & "*.nes" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    
    'DF: Changed to app.path
    OpenFile.lpstrInitialDir = App.Path & "\"
    
    OpenFile.lpstrTitle = "Open *.NES ROM"
    OpenFile.flags = 4
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
       open_file = ""
    Else
       open_file = OpenFile.lpstrFile
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode >= Asc("0") And KeyCode <= Asc("9") Then
    'select save slot
    mSaveSlots_Click KeyCode - Asc("0")
Else
    Select Case KeyCode 'opposite arrow keys can't both be down. Many games have problems if they are.
        Case 37
            Keyboard(nes_ButRt) = &H40
        Case 38
            Keyboard(nes_ButDn) = &H40
        Case 39
            Keyboard(nes_ButLt) = &H40
        Case 40
            Keyboard(nes_ButUp) = &H40
    End Select
    Keyboard(KeyCode) = &H41
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Keyboard(KeyCode) = &H40
End Sub

Private Sub GetPaletteList()
    Dim s As String
    Dim i As Long
    s = Dir$(App.Path + "\*.pal")
    Do While s <> ""
        If i Then Load mSelPalette(i)
        mSelPalette(i).Caption = s
        i = i + 1
        s = Dir$
    Loop
End Sub

Private Sub Form_Load()

newScroll = True
mAutoSpeed_Click
mnuFS_Click 0

GetPaletteList
MapperNames = Array("NROM", "MMC1", "UNROM", "CNROM", "MMC3", "MMC5", "FFE F4xxx", "AOROM", "FFE F3xxx", "MMC2", "MMC4", "Color Dreams", 0, "Videomation", 0, "100-In-1", "Bandai", "FFE F8xxx", "Jaleco SS8806", "Namcot 106", "Konami VRC4", "Konami VRC2 type A", "Konami VRC2 type B", "Konami VRC6")
doSound = True
MidiOpen
SelectInstrument 5, 123
'SelectInstrument 6, 30
ToneOn 5, 100, 64
'ToneOn 6, 31, 127
pAPUinit
tilebased = False
FrameSkip = 1
maxCycles1 = 114
autospeed = True
mAutoSpeed.Checked = True
mSaveSlots_Click 0
mnu0_Click 0

LoadConfig

Caption = VERSION
CPUPaused = False
frmNES.mnuCPUPause.Checked = False
Show
If palName = "" Then palName = "bnes.pal"
If Dir(palName) = "" Then palName = "bnes.pal"

LoadPal palName

Dim i As Long
For i = 0 To 30
    pow2(i) = 2 ^ i
Next i
pow2(31) = -2147483648#

'Fill our half mb color lookup table
fillTLook

msave.Enabled = False
mrestore.Enabled = False
mnuStartRecord.Enabled = False
mnuCPUPause.Enabled = False
mnuEmuReset.Enabled = False
mnuPlayMovie.Enabled = False
mZoom_Click 0
End Sub

Private Sub Form_Terminate()
    MidiClose
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MidiClose 'Otherwise midi will be unavailable until reboot
    CleanUp 'Needed by fastblit

    Cancel = True
    mnuFileExit_Click
End Sub


Private Sub lblStatus_Change()
    StatusTimeout = 36
End Sub



Private Sub mAutoSpeed_Click()
    Dim i As Long
'    FrameSkip = 1
    autospeed = Not autospeed
    mAutoSpeed.Checked = autospeed
'    For i = 0 To mnuFS.UBound
        'mnuFS(i).Checked = False
    'Next i
End Sub

Private Sub mexecv_Click(Index As Integer)
    If Index = 0 Then
        SmartExec = True
    ElseIf Index = 1 Then
        maxCycles1 = 171
    Else
        SmartExec = False
        ' Last 22 scanlines are 100% exec, so we have to adjust the exec for the rest to average out
        maxCycles1 = (114& * 262& * (6& - Index) \ 4& - 114& * 44&) \ 218&
    End If

    Dim i As Long

    For i = 0 To mexecv.UBound
        mexecv(i).Checked = False
    Next i
    mexecv(Index).Checked = True
End Sub

Private Sub mfast_Click()
    mIdle.Checked = True
    IdleDetect = True
    tilebased = True
    mtiled.Checked = True
End Sub

Private Sub mIdle_Click()
    mIdle.Checked = Not mIdle.Checked
    IdleDetect = mIdle.Checked
End Sub

Private Sub mMotionBlur_Click()
    MotionBlur = Not MotionBlur
    mMotionBlur.Checked = MotionBlur
End Sub

Private Sub mMute_Click()
    mMute.Checked = doSound
    doSound = Not doSound
End Sub

Private Sub mNewScroll_Click()
    newScroll = Not newScroll
    mNewScroll.Checked = newScroll
End Sub

Private Sub mnu0_Click(Index As Integer)
lblStatus.Caption = "Movie slot is " & CStr(Index)
Dim iiii As Byte
For iiii = 0 To 4
    mnu0(iiii).Checked = False
Next iiii
mnu0(Index).Checked = True
MovieIndex = Index
End Sub

Public Sub mnuCPUPause_Click()
If CPURunning = False Then Exit Sub
    If CPUPaused = False Then
        StopSound
        CPUPaused = True
        mnuCPUPause.Checked = True
    Else
        CPUPaused = False
        mnuCPUPause.Checked = False
    End If
End Sub

Private Sub mnuEmuConfgKeys_Click()
If CPUPaused = False Then mnuCPUPause_Click
Load frmConfig
End Sub

Private Sub mnuEmuReset_Click()
    Select Case Mapper
        Case 0, 3
            Select8KVROM 0
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
        Case 2
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
        Case 4
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
            MMC3_IrqOn = False
            MMC3_IrqVal = 0
            MMC3_TmpVal = 0
            If ChrCount Then Select8KVROM 0
        Case 7
            reg8 = 0
            regA = 1
            regC = 2
            regE = 3
            SetupBanks
        Case 8 'FFE F3xxx
            reg8 = 0: regA = 1
            regC = &HFE: regE = &HFF
            SetupBanks
            Select8KVROM 0
        Case 9 ' MMC2
            reg8 = 0
            regA = &HFD
            regC = &HFE
            regE = &HFF
            SetupBanks
            'latch1.hi_bank = 0: latch1.lo_bank = 0: latch1.state = 0
            'latch2.hi_bank = 0: latch2.lo_bank = 0: latch2.state = 0
         Case 11
            reg8 = 0: regA = 1: regC = 2: regE = 3
            SetupBanks
            Select8KVROM 0
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
        Case 32
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
        Case 33
            reg8 = 0
            regA = 1
            regC = &HFE
            regE = &HFF
            SetupBanks
        Case 34
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
            Debug.Print ChoseIt
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
        Case 66
            reg8 = 0: regA = 1: regC = 2: regE = 3: SetupBanks
            Select8KVROM 0
        Case 91
            reg8 = &HFE
            regA = &HFF
            regC = &HFE
            regE = &HFF
            Select8KVROM 0
            SetupBanks
    End Select
    'Debug.Print "Successfully loaded " & filename
    reset6502
    If Mirroring = 1 Then MirrorXor = &H800& Else MirrorXor = &H400&
    Debug.Print "Reset from menu"
End Sub

Private Sub mnuFileExit_Click()
Dim FileNum As Integer
FileNum = FreeFile
Erase gameImage, VROM, VRAM, bank0, bank6, bank8, bankA, bankC, bankE
SaveConfig
If UsesSRAM = True Then ' save the SRAM to a file.
    Open App.Path & "\" & romName & ".wrm" For Binary As #FileNum
        Put #FileNum, , bank6
    Close #FileNum
End If
End
End Sub

Private Sub mnuFileFree_Click()
' Erase all known content of rom.
Erase VRAM: Erase gameImage: Erase bank0: Erase bank6
Erase VROM: Erase SpriteRAM: Erase Joypad1
Erase bank8: Erase bankA: Erase bankC: Erase bankE
picScreen.Cls
CPURunning = False
msave.Enabled = False
mrestore.Enabled = False
mnuStartRecord.Enabled = False
mnuPlayMovie.Enabled = False
mnuCPUPause.Enabled = False
mnuEmuReset.Enabled = False
mnuFileRomInfo.Enabled = False
mnuFileFree.Enabled = False
If UsesSRAM = True Then ' save the SRAM to a file.
    Open App.Path & "\" & romName & ".wrm" For Binary Access Write As #11
        Put #11, , bank6
    Close #11
End If
UsesSRAM = False
End Sub

Private Function hasAny(ByVal s As String, ParamArray P() As Variant) As Boolean
    Dim i As Long
    s = Replace(LCase(s), " ", "")
    For i = 0 To UBound(P)
        If InStr(s, P(i)) Then
            hasAny = True
            Exit Function
        End If
    Next i
    hasAny = False
End Function

Private Sub presetfixes(ByVal file As String)
maxCycles1 = 114
'If hasAny(file, "smb", "mario", "kirby", "zelda", "dragonwarrior2", "dragonw2", "dwarrior2") Then
    'roms broken by the new scroll code
    'newScroll = False
'ElseIf hasAny(file, "megaman", "rockman") Then
    'maxCycles1 = 145 'megaman seems to benefit greatly from just a few extra cycles
    'newScroll = True
'ElseIf hasAny(file, "cmc", "spiderman", "finalf3", "ff3", "ducktales") Then
    'roms that require the new scroll code
'    newScroll = True
'Else
If hasAny(file, "demonhead", "elite") Then
    'Clash at demonhead must be run at 150% execute, not sure why.
    'Also, Elite benefits greatly from the extra cycles.
    'Elite does not get along with the new scrolling.
    '  Many emulators share the same problem.
    newScroll = False
    mexecv_Click 1
End If
mNewScroll.Checked = newScroll
End Sub

Private Sub mnuFileLoad_Click()
StopSound

Dim filename As String
filename = open_file()
CPUPaused = False
If filename = "" Then Exit Sub
romName = extractFileName(filename)

If LoadNES(filename) = 0 Then Exit Sub
lblStatus = romName + " loaded"
lblStatus.Refresh
CPURunning = True
FirstRead = True
PPU_AddressIsHi = True
PPUAddress = 0: SpriteAddress = 0: PPU_Status = 0
PPU_Control1 = 0: PPU_Control2 = 0
init6502
lblStatus = romName + " initialized"
msave.Enabled = True
mrestore.Enabled = True
mnuStartRecord.Enabled = True
mnuPlayMovie.Enabled = True
mnuCPUPause.Enabled = True
mnuEmuReset.Enabled = True

presetfixes romName

Do Until CPURunning = False
    exec6502
Loop
End Sub

Private Sub mnuFileRomInfo_Click()
#If 0 Then
    Dim Message As String
    Message = "ROM Info for " & romName & vbCrLf & vbCrLf
    Message = Message & "PRG Banks: " & PrgCount & vbCrLf
    Message = Message & "CHR Banks: " & ChrCount & vbCrLf
    Message = Message & "Mirroring Type: " & IIf(Mirroring And &H1, "Vertical", "Horizontal") & vbCrLf
    Message = Message & "Trainer: " & IIf(Trainer, "Yes", "No") & vbCrLf
    Message = Message & "Mapper: " & Mapper & "(" & MapperNames(Mapper) & ")" & vbCrLf
    MsgBox Message, 64, VERSION
#End If
Load frmROMInfo
End Sub

Public Sub mnuFS_Click(Index As Integer)
'autospeed = False
'mAutoSpeed.Checked = False
Dim i As Long

lblStatus.Caption = "Set frameskip to " & Index

FrameSkip = Index + 1
For i = 0 To mnuFS.UBound
    mnuFS(i).Checked = False
Next i
On Error Resume Next
mnuFS(Index).Checked = True
End Sub

Private Sub mnuHelpAbout_Click()
Load frmAbout
End Sub

Private Sub mnuPlayMovie_Click()
If mnuPlayMovie.Caption = "&Play" Then
    PlayMovie CLng(MovieIndex)
    Record = False
    Playing = True
    mnuStartRecord.Caption = "&Stop"
    lblStatus.Caption = "Playing from slot " & MovieIndex
Else
    StopPlaying
    mnuStartRecord.Caption = "&Start"
End If
End Sub
Private Sub mnuStartRecord_Click()
If mnuStartRecord.Caption = "&Start" Then
    RecordMovie CLng(MovieIndex)
    Playing = False
    mnuStartRecord.Caption = "&Stop"
    lblStatus.Caption = "Recording to slot " & MovieIndex
Else
    StopRecording
    mnuStartRecord.Caption = "&Start"
End If
End Sub

Private Sub mrestore_Click()
loadState SlotIndex
lblStatus = "Loaded from slot " + CStr(SlotIndex)
End Sub

Private Sub msave_Click()
saveState SlotIndex
lblStatus = "Saved to slot " + CStr(SlotIndex)
End Sub


Private Sub mSaveSlots_Click(Index As Integer)
SlotIndex = Index
Dim i As Long
For i = 0 To 9
msaveslots(i).Checked = False
Next i
msaveslots(Index).Checked = True
lblStatus = "Selected save slot " + CStr(SlotIndex)
End Sub


Private Sub mSelPalette_Click(Index As Integer)
palName = mSelPalette(Index).Caption
LoadPal mSelPalette(Index).Caption
End Sub



Private Sub mSmooth_Click(Index As Integer)
mSmooth(Smooth2x).Checked = False
Smooth2x = Index
mSmooth(Smooth2x).Checked = True
End Sub

Private Sub mtiled_Click()
    tilebased = Not tilebased
    mtiled.Checked = tilebased
End Sub

Private Sub mZoom_Click(Index As Integer)
If Index >= 1 Then
    If mSmooth(0).Checked Then
        Smooth2x = 0
    ElseIf mSmooth(1).Checked Then
        Smooth2x = 1
    ElseIf mSmooth(2).Checked Then
        Smooth2x = 2
    End If
    mSmooth(1).Enabled = True
    mSmooth(2).Enabled = True
Else
    Smooth2x = 0
    mSmooth(1).Enabled = False
    mSmooth(2).Enabled = False
End If
picScreen.Move picScreen.Left, picScreen.Top, (256 * (Index + 1) + 2) * 15, (240 * (Index + 1) + 2) * 15
Move Left, Top, picScreen.Width + 38 * 15, picScreen.Height + 96 * 15
lblStatus.Top = picScreen.Top + picScreen.Height + 4 * 15
lbSpeed.Top = picScreen.Top + picScreen.Height + 4 * 15
End Sub



Private Sub picScreen_Paint()
    If CPUPaused Then blitScreen
End Sub

'DF: Measures speed in fps
Private Sub Timer1_Timer()
On Error Resume Next

Static P As Long, pr As Long
Static ptime As Double
Dim ctime As Double

StatusTimeout = StatusTimeout - 1
If StatusTimeout = 0 Then lblStatus = ""
Dim s As String
s = ""
If (maxCycles1 < 114 Or IdleDetect) And nCycles > 1 Then
    s = " " + CStr((rCycles * 100 + nCycles \ 2) \ nCycles) + "% execution"
End If
rCycles = 0
nCycles = 0

ctime = Timer
lbSpeed = CStr(CLng((realframes - pr) / (ctime - ptime))) + " fps (" + CStr(CLng((Frames - P) / (ctime - ptime))) + " virtual)" + s
ptime = ctime
P = Frames
pr = realframes
End Sub

