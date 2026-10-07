Attribute VB_Name = "modConfig"
' config.bas
' april 14th, 2002
' by don jarrett

Public Const ConfigFile = "bNES.cfg" ' basicNES configuration file.

' NES button constants, to be saved in this order.
Public nes_ButA As Byte
Public nes_ButB As Byte
Public nes_ButSel As Byte
Public nes_ButSta As Byte

Public nes_ButUp As Byte
Public nes_ButDn As Byte
Public nes_ButLt As Byte
Public nes_ButRt As Byte
Public Function LoadConfig() As Boolean
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Dim tmp As String
    
    ChDir App.Path & "\"
    If Dir(ConfigFile) = "" Then
        MsgBox "Config file not found! Using default keys!", vbExclamation, VERSION
        nes_ButA = vbKeyX
        nes_ButB = vbKeyZ
        nes_ButSel = vbKeyC
        nes_ButSta = vbKeyV
        nes_ButUp = vbKeyUp
        nes_ButDn = vbKeyDown
        nes_ButLt = vbKeyLeft
        nes_ButRt = vbKeyRight
        LoadConfig = False
        Exit Function
    End If
    
    Open ConfigFile For Input As #FileNum
        Line Input #FileNum, tmp: nes_ButA = CByte(tmp)
        Line Input #FileNum, tmp: nes_ButB = CByte(tmp)
        Line Input #FileNum, tmp: nes_ButSel = CByte(tmp)
        Line Input #FileNum, tmp: nes_ButSta = CByte(tmp)
        Line Input #FileNum, tmp: nes_ButUp = CByte(tmp)
        Line Input #FileNum, tmp: nes_ButDn = CByte(tmp)
        Line Input #FileNum, tmp: nes_ButLt = CByte(tmp)
        Line Input #FileNum, tmp: nes_ButRt = CByte(tmp)
        Line Input #FileNum, tmp: SaveCPU = CBool(tmp)
        If EOF(FileNum) = False Then
            Line Input #FileNum, palName
        End If
    Close #FileNum
    
    LoadConfig = True
    
End Function
Public Sub SaveConfig()
    
    ChDir App.Path & "\"
    Dim FileNum As Integer
    FileNum = FreeFile
    
    Open ConfigFile For Output As #FileNum
        Print #FileNum, nes_ButA
        Print #FileNum, nes_ButB
        Print #FileNum, nes_ButSel
        Print #FileNum, nes_ButSta
        
        Print #FileNum, nes_ButUp
        Print #FileNum, nes_ButDn
        Print #FileNum, nes_ButLt
        Print #FileNum, nes_ButRt
        Print #FileNum, SaveCPU
        Print #FileNum, palName
    Close #FileNum

End Sub
