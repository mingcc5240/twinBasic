VERSION 5.00
Begin VB.Form frmKeys 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Define Keys "
   ClientHeight    =   3510
   ClientLeft      =   645
   ClientTop       =   2400
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Keys for Joypad 2:"
      Height          =   1245
      Left            =   4830
      TabIndex        =   16
      Top             =   1800
      Width           =   4680
      Begin VB.TextBox txtStart2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   435
         Width           =   435
      End
      Begin VB.TextBox txtSelect2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   750
         Width           =   435
      End
      Begin VB.TextBox txtA2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   765
         Width           =   435
      End
      Begin VB.TextBox txtB2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   450
         Width           =   435
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         X1              =   90
         X2              =   4590
         Y1              =   255
         Y2              =   255
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   90
         X2              =   4590
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Label Label9 
         Caption         =   "Start - "
         Height          =   225
         Left            =   210
         TabIndex        =   24
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "Select - "
         Height          =   225
         Left            =   105
         TabIndex        =   23
         Top             =   765
         Width           =   585
      End
      Begin VB.Label Label6 
         Caption         =   "A - "
         Height          =   225
         Left            =   1485
         TabIndex        =   22
         Top             =   765
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "B - "
         Height          =   225
         Left            =   1485
         TabIndex        =   21
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   930
      TabIndex        =   4
      Top             =   3105
      Width           =   885
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
      Height          =   345
      Left            =   15
      TabIndex        =   3
      Top             =   3105
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Keys for Joypad 1:"
      Height          =   1245
      Left            =   15
      TabIndex        =   0
      Top             =   1800
      Width           =   4680
      Begin VB.TextBox txtB 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   450
         Width           =   435
      End
      Begin VB.TextBox txtA 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   765
         Width           =   435
      End
      Begin VB.TextBox txtSelect 
         Enabled         =   0   'False
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   750
         Width           =   435
      End
      Begin VB.TextBox txtStart 
         Enabled         =   0   'False
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   435
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "B - "
         Height          =   225
         Left            =   1485
         TabIndex        =   15
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "A - "
         Height          =   225
         Left            =   1485
         TabIndex        =   14
         Top             =   765
         Width           =   300
      End
      Begin VB.Label lblChrA 
         Height          =   285
         Left            =   1230
         TabIndex        =   8
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "Select - "
         Height          =   225
         Left            =   105
         TabIndex        =   2
         Top             =   765
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Start - "
         Height          =   225
         Left            =   225
         TabIndex        =   1
         Top             =   480
         Width           =   465
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   90
         X2              =   4590
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   90
         X2              =   4590
         Y1              =   255
         Y2              =   255
      End
   End
   Begin VB.Label trigStart2 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   7050
      TabIndex        =   28
      Top             =   1305
      Width           =   450
   End
   Begin VB.Label trigSelect2 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   6345
      TabIndex        =   27
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label trigB2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8025
      TabIndex        =   26
      Top             =   1215
      Width           =   330
   End
   Begin VB.Label trigA2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8610
      TabIndex        =   25
      Top             =   1215
      Width           =   330
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   4785
      Picture         =   "frmKeys.frx":0000
      Top             =   0
      Width           =   4800
   End
   Begin VB.Label trigA 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3765
      TabIndex        =   11
      Top             =   1215
      Width           =   330
   End
   Begin VB.Label trigB 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3180
      TabIndex        =   10
      Top             =   1215
      Width           =   330
   End
   Begin VB.Label trigSelect 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   1500
      TabIndex        =   9
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label trigStart 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   2205
      TabIndex        =   7
      Top             =   1305
      Width           =   450
   End
   Begin VB.Image imgGamepad 
      Height          =   1800
      Left            =   -45
      Picture         =   "frmKeys.frx":16CF
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "frmKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Key Defining Added by Factor-1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'sk short for selectkey
Public Function sk(i As Long, a)
    keycodes(i) = CLng(a)
    sk = a
End Function

Private Sub cmdSet_Click()
    Call WriteINI("JOY1", "START", sk(0, txtStart))
    Call WriteINI("JOY1", "SELECT", sk(1, txtSelect))
    Call WriteINI("JOY1", "A", sk(2, txtA))
    Call WriteINI("JOY1", "B", sk(3, txtB))
    Call WriteINI("JOY2", "START", sk(4, txtStart2))
    Call WriteINI("JOY2", "SELECT", sk(5, txtSelect2))
    Call WriteINI("JOY2", "A", sk(6, txtA2))
    Call WriteINI("JOY2", "B", sk(7, txtB2))
    Unload Me
End Sub

Private Sub Form_Load()

    txtStart = sk(0, ReadINI("JOY1", "START"))
    txtSelect = sk(1, ReadINI("JOY1", "SELECT"))
    txtA = sk(2, ReadINI("JOY1", "A"))
    txtB = sk(3, ReadINI("JOY1", "B"))
    txtStart2 = sk(4, ReadINI("JOY2", "START"))
    txtSelect2 = sk(5, ReadINI("JOY2", "SELECT"))
    txtA2 = sk(6, ReadINI("JOY2", "A"))
    txtB2 = sk(7, ReadINI("JOY2", "B"))
    
End Sub

Private Sub trigA_Click()
    txtA.Enabled = True
    txtA.SetFocus
End Sub

Private Sub trigA2_Click()
    txtA2.Enabled = True
    txtA2.SetFocus
End Sub

Private Sub trigB_Click()
    txtB.Enabled = True
    txtB.SetFocus
End Sub

Private Sub trigB2_Click()
    txtB2.Enabled = True
    txtB2.SetFocus
End Sub

Private Sub trigSelect_Click()
    txtSelect.Enabled = True
End Sub

Private Sub trigSelect2_Click()
    txtSelect2.Enabled = True
End Sub

Private Sub trigStart_Click()
    txtStart.Enabled = True
End Sub

Private Sub trigStart2_Click()
    txtStart2.Enabled = True
End Sub

Private Sub txtA_KeyDown(KeyCode As Integer, Shift As Integer)
    txtA.Text = KeyCode
End Sub

Private Sub txtB_KeyDown(KeyCode As Integer, Shift As Integer)
    txtB.Text = KeyCode
End Sub

Private Sub txtSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    txtSelect.Text = KeyCode
End Sub

Private Sub txtStart_KeyDown(KeyCode As Integer, Shift As Integer)
    txtStart.Text = KeyCode
End Sub

Private Sub txtStart2_KeyDown(KeyCode As Integer, Shift As Integer)
    txtStart2.Text = KeyCode
End Sub

Private Sub txtselect2_KeyDown(KeyCode As Integer, Shift As Integer)
    txtSelect2.Text = KeyCode
End Sub

Private Sub txtA2_KeyDown(KeyCode As Integer, Shift As Integer)
    txtA2.Text = KeyCode
End Sub

Private Sub txtB2_KeyDown(KeyCode As Integer, Shift As Integer)
    txtB2.Text = KeyCode
End Sub
