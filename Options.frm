VERSION 5.00
Begin VB.Form Options 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   12960
   ClientLeft      =   450
   ClientTop       =   -465
   ClientWidth     =   17280
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12960
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   13200
      Top             =   2040
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1080
      Picture         =   "Options.frx":0000
      ScaleHeight     =   1695
      ScaleWidth      =   15135
      TabIndex        =   1
      Top             =   10800
      Width           =   15135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1080
      Picture         =   "Options.frx":5F3B2
      ScaleHeight     =   1695
      ScaleWidth      =   15135
      TabIndex        =   0
      Top             =   240
      Width           =   15135
   End
   Begin VB.Label LBL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gameplay Options"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1095
      Index           =   3
      Left            =   4800
      TabIndex        =   5
      Tag             =   "16711935"
      Top             =   5160
      Width           =   7215
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   9
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   8
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   7
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   6
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   5
      Left            =   360
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LBL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1095
      Index           =   2
      Left            =   7560
      TabIndex        =   4
      Tag             =   "49152"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label LBL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Define Keys"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1095
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Tag             =   "192"
      Top             =   6720
      Width           =   4815
   End
   Begin VB.Label LBL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   975
      Index           =   0
      Left            =   7560
      TabIndex        =   2
      Tag             =   "12582912"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H00000000&
      Height          =   5895
      Index           =   4
      Left            =   3480
      Tag             =   "16711935"
      Top             =   3600
      Width           =   10335
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H00000040&
      Height          =   6375
      Index           =   3
      Left            =   3240
      Top             =   3360
      Width           =   10815
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H00000080&
      Height          =   6855
      Index           =   2
      Left            =   3000
      Top             =   3120
      Width           =   11295
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H000000C0&
      Height          =   7335
      Index           =   1
      Left            =   2760
      Top             =   2880
      Width           =   11775
   End
   Begin VB.Shape SQ 
      BorderColor     =   &H000000FF&
      Height          =   7815
      Index           =   0
      Left            =   2520
      Top             =   2640
      Width           =   12255
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'49152
'192
'12582912
'16711935

Private Sub Form_Load()
    For a = 0 To 7
        Colors(a) = SQ(a).BorderColor
        SQ(a).Tag = a
    Next
    
    For E = 0 To 5
        InString = "_______________________"
        If E = 0 Then
            ISLength = GetPrivateProfileString("Keys", "Up", "Up Arrow", InString, 13, App.Path & "\Stuff.ini")
        ElseIf E = 1 Then
            ISLength = GetPrivateProfileString("Keys", "Down", "Down Arrow", InString, 13, App.Path & "\Stuff.ini")
            KeyUp = tempKey
        ElseIf E = 2 Then
            ISLength = GetPrivateProfileString("Keys", "Left", "Left Arrow", InString, 13, App.Path & "\Stuff.ini")
            KeyDown = tempKey
        ElseIf E = 3 Then
            ISLength = GetPrivateProfileString("Keys", "Right", "Right Arrow", InString, 13, App.Path & "\Stuff.ini")
            KeyLeft = tempKey
        ElseIf E = 4 Then
            ISLength = GetPrivateProfileString("Keys", "Guns", "Shift", InString, 13, App.Path & "\Stuff.ini")
            KeyRight = tempKey
        Else
            ISLength = GetPrivateProfileString("Keys", "Bombs", "Control", InString, 13, App.Path & "\Stuff.ini")
            KeyGuns = tempKey
        End If
        InString = Left(InString, ISLength)
        DKeys.KeyC(E).Caption = InString
        If InString = "Up Arrow" Then
            tempKey = vbKeyUp
        ElseIf InString = "Down Arrow" Then
            tempKey = vbKeyDown
        ElseIf InString = "Left Arrow" Then
            tempKey = vbKeyLeft
        ElseIf InString = "Right Arrow" Then
            tempKey = vbKeyRight
        ElseIf InString = "Space Bar" Then
            tempKey = vbKeySpace
        ElseIf InString = "Control" Then
            tempKey = vbKeyControl
        ElseIf InString = "Shift" Then
            tempKey = vbKeyShift
        ElseIf InString = "Caps Lock" Then
            tempKey = vbKeyCapital
        ElseIf InString = "Backspace" Then
            tempKey = vbKeyBack
        ElseIf InString = "Enter" Then
            tempKey = vbKeyReturn
        ElseIf InString = "Tab" Then
            tempKey = vbKeyTab
        ElseIf InString = "Delete" Then
            tempKey = vbKeyDelete
        ElseIf InString = "Insert" Then
            tempKey = vbKeyInsert
        ElseIf InString = "Home" Then
            tempKey = vbKeyHome
        ElseIf InString = "End" Then
            tempKey = vbKeyEnd
        ElseIf InString = "Page Up" Then
            tempKey = vbKeyPageUp
        ElseIf InString = "Page Down" Then
            tempKey = vbKeyPageDown
        ElseIf InString = "Numlock" Then
            tempKey = vbKeyNumlock
        ElseIf InString = "/" Then
            tempKey = vbKeyDivide
        ElseIf InString = "*" Then
            tempKey = vbKeyMultiply
        ElseIf InString = "-" Then
            tempKey = vbKeySubtract
        ElseIf InString = "+" Then
            tempKey = vbKeyAdd
        ElseIf InString = "Enter(numpad)" Then
            tempKey = vbKeySelect
        ElseIf InString = "." Then
            tempKey = vbKeyDecimal
        ElseIf InString = "Menu Key" Then
            tempKey = vbKeyMenu
        ElseIf InString = "Keypad 0" Then
            tempKey = vbKeyNumpad0
        ElseIf InString = "Keypad 1" Then
            tempKey = vbKeyNumpad1
        ElseIf InString = "Keypad 2" Then
            tempKey = vbKeyNumpad2
        ElseIf InString = "Keypad 3" Then
            tempKey = vbKeyNumpad3
        ElseIf InString = "Keypad 4" Then
            tempKey = vbKeyNumpad4
        ElseIf InString = "Keypad 5" Then
            tempKey = vbKeyNumpad5
        ElseIf InString = "Keypad 6" Then
            tempKey = vbKeyNumpad6
        ElseIf InString = "Keypad 7" Then
            tempKey = vbKeyNumpad7
        ElseIf InString = "Keypad 8" Then
            tempKey = vbKeyNumpad8
        ElseIf InString = "Keypad 9" Then
            tempKey = vbKeyNumpad9
        Else
            tempKey = Asc(InString)
        End If
    Next
    KeyBombs = tempKey
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    For c = 0 To 3
        LBL(c).ForeColor = &H40&
    Next
End Sub

Private Sub LBL_Click(Index As Integer)
    If Index = 0 Then
        End
    ElseIf Index = 1 Then
        DKeys.Visible = True
        LBL(1).ForeColor = &H40&
    ElseIf Index = 2 Then
        Difficulty = GamePlay.ProgBar1.Value - 1
        AltStart.Visible = True
        Options.Visible = False
    ElseIf Index = 3 Then
        GamePlay.Visible = True
        LBL(3).ForeColor = &H40&
    End If
End Sub

Private Sub LBL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    LBL(Index).ForeColor = LBL(Index).Tag
End Sub

Private Sub Timer1_Timer()
    For B = 0 To 7
        SQ(B).Tag = SQ(B).Tag + 1
        If SQ(B).Tag = 8 Then SQ(B).Tag = 0
        SQ(B).BorderColor = Colors(SQ(B).Tag)
    Next
End Sub
