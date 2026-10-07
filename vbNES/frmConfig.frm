VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSta 
      Height          =   285
      Left            =   2880
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtSel 
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtB 
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtA 
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtRight 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtLeft 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtDown 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtUp 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Start"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Select"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "B"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "A"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Configure the basicNES emulator here. More options soon."
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "Right"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Left"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Down"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Up"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
' don't save the stuff.
If CPUPaused = True Then frmNES.mnuCPUPause_Click
Unload frmConfig
End Sub

Private Sub cmdOK_Click()
nes_ButA = CByte(txtA.Text): nes_ButB = CByte(txtB.Text): nes_ButSel = CByte(txtSel.Text): nes_ButSta = CByte(txtSta.Text)
nes_ButUp = CByte(txtUp.Text): nes_ButDn = CByte(txtDown.Text): nes_ButLt = CByte(txtLeft.Text): nes_ButRt = CByte(txtRight.Text)
If CPUPaused = True Then frmNES.mnuCPUPause_Click
Unload frmConfig
End Sub

Private Sub Form_Load()
Caption = "Configure " & VERSION
txtUp.Text = nes_ButUp: txtDown.Text = nes_ButDn: txtLeft.Text = nes_ButLt: txtRight.Text = nes_ButRt
txtA.Text = nes_ButA: txtB.Text = nes_ButB: txtSel.Text = nes_ButSel: txtSta.Text = nes_ButSta
Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
If CPUPaused = True Then frmNES.mnuCPUPause_Click
Unload frmConfig
End Sub

Private Sub txtA_KeyDown(KeyCode As Integer, Shift As Integer)
    txtA.Text = KeyCode
End Sub

Private Sub txtA_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtB_KeyDown(KeyCode As Integer, Shift As Integer)
    txtB.Text = KeyCode
End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtDown_KeyDown(KeyCode As Integer, Shift As Integer)
    txtDown.Text = KeyCode
End Sub

Private Sub txtDown_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtLeft_KeyDown(KeyCode As Integer, Shift As Integer)
    txtLeft.Text = KeyCode
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtRight_KeyDown(KeyCode As Integer, Shift As Integer)
    txtRight.Text = KeyCode
End Sub

Private Sub txtRight_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtSel_KeyDown(KeyCode As Integer, Shift As Integer)
    txtSel.Text = KeyCode
End Sub

Private Sub txtSel_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtSta_KeyDown(KeyCode As Integer, Shift As Integer)
    txtSta.Text = KeyCode
End Sub

Private Sub txtSta_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtUp_KeyDown(KeyCode As Integer, Shift As Integer)
    txtUp.Text = KeyCode
End Sub

Private Sub txtUp_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
