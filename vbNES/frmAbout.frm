VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   5055
   End
   Begin VB.TextBox txtAbout 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "http://home.att.net/~r.jarrett/bNES.html"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1680
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label lblHome 
      Caption         =   "basicNES Homepage:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmAbout
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Font.Underline = False
End Sub

Private Sub Form_Load()
Dim Message As String
Message = VERSION & vbCrLf
Message = Message & vbCrLf & "By Don Jarrett, 1996-2003" & vbCrLf & "Optimization, Sprites, and scrolling, etc. by David Finch, 2000-2003." & vbCrLf
Message = Message & "http://home.att.net/~r.jarrett/bNES.html" & vbCrLf & vbCrLf
Message = Message & "Some mapper code is courtesy of DarcNES and Pretendo." & vbCrLf & "Thanks to both nyef and Delta for the information." & vbCrLf
Message = Message & "Open File Dialog code courtesy Matthew Leverton." & vbCrLf
Message = Message & "Line based rendering ideas courtesy of Lothos." & vbCrLf
Message = Message & "Palette code ideas by Bero of FCE." & vbCrLf & "Thanks to The Quietust for mapper information and help." & vbCrLf & "Thanks to Kevin Horton for his help and mad support." & vbCrLf
Message = Message & "Big time thanks to squaredot for the URL link code, support and his friendship. To Chris Cowley for the interface ideas and much help with my other emulators." & vbCrLf
Message = Message & vbCrLf & "Check out vbSpec, a Spectrum emulator in VB, at http://freestuff.grok.co.uk/vbspec/" & vbCrLf & vbCrLf & "Grab my other emulators at http://home.att.net/~r.jarrett/prod.html" & vbCrLf & vbCrLf
Message = Message & "I do not advocate the piracy of Nintendo ROMs." & vbCrLf & vbCrLf
Message = Message & "basicNES is Copyright (C) 1996-2003 Don Jarrett." & vbCrLf
Message = Message & "Portions of code Copyright (C) 2000-2003 David Finch." & vbCrLf
Message = Message & "M6502 is Copyright (C) 1997-2003 Don Jarrett & Tobias Stromstedt." & vbCrLf
txtAbout.Text = Message
Caption = "About basicNES 2000"
Icon = frmNES.Icon
Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Font.Underline = False
End Sub

Private Sub Label1_Click()
ShellExecute hwnd, "open", "http://home.att.net/~r.jarrett/bNES.html", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Font.Underline = True
End Sub

Private Sub lblHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Font.Underline = False
End Sub

Private Sub txtAbout_GotFocus()
Command1.SetFocus
End Sub

Private Sub txtAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Font.Underline = False
End Sub
