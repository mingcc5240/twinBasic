VERSION 5.00
Begin VB.Form frmROMInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rom Information"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   3375
   End
   Begin VB.ListBox lstInfo 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmROMInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LAdd(ByVal textmsg As String)
    lstInfo.AddItem textmsg
End Sub
Private Sub cmdOK_Click()
Unload Me
End Sub
Private Sub Form_Load()
Dim MType As String
LAdd "ROM Information"
LAdd ""
LAdd "Filename: " & romName & ".nes"
LAdd "PRG-ROM: " & (PrgCount * 16) & "k"
LAdd "CHR-ROM: " & (ChrCount * 8) & "k"
LAdd "Mapper: " & Mapper '& "(" & MapperNames(Mapper) & ")"
If Mirroring = 0 Then MType = "Horizontal"
If Mirroring = 1 Then MType = "Vertical"
If Mirroring = 2 Then MType = "One-Screen"
If Mirroring = 4 Then MType = "Four-Screen"
LAdd "Mirroring: " & MType
LAdd "Trainer: " & IIf(Trainer, "Yes", "No")
Show
End Sub

