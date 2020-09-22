VERSION 5.00
Begin VB.Form DKeys 
   Appearance      =   0  'Flat
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   5070
   ClientLeft      =   6825
   ClientTop       =   3240
   ClientWidth     =   4470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton EQ 
      BackColor       =   &H0000C000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton EQ 
      BackColor       =   &H0000C000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton EQ 
      BackColor       =   &H0000C000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton EQ 
      BackColor       =   &H0000C000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton EQ 
      BackColor       =   &H0000C000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton EQ 
      BackColor       =   &H0000C000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   -480
      Tag             =   "4"
      Top             =   6600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "DONE"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   21.75
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1320
      TabIndex        =   24
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      Index           =   6
      X1              =   0
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label KeyC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   23
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label KeyC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   22
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label KeyC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Right Arrow"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   21
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label KeyC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Left Arrow"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   20
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label KeyC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Down Arrow"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   19
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      Index           =   5
      X1              =   0
      X2              =   4560
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      Index           =   4
      X1              =   0
      X2              =   4560
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      Index           =   3
      X1              =   -120
      X2              =   4440
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      Index           =   2
      X1              =   0
      X2              =   4560
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      Index           =   1
      X1              =   -120
      X2              =   4440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      Index           =   0
      X1              =   0
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label KeyC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Up arrow"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   18
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label KeyDB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Bombs"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label KeyDB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label KeyDB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label KeyDB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label KeyDB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label KeyDB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Guns"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label KeyD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Bombs"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label KeyD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label KeyD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label KeyD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Guns"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label KeyD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label KeyD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FF80&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H00008000&
      FillStyle       =   6  'Cross
      Height          =   5055
      Left            =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "DKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EQ_Click(Index As Integer)
    If KeyInd <> -1 Then KeyD(KeyInd).BackStyle = 0
    KeyInd = Index
    KeyC(KeyInd).BackStyle = 1
    For D = 0 To 5
        EQ(D).Enabled = False
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyInd = 0 Then KeyUp = KeyCode
    If KeyInd = 1 Then KeyDown = KeyCode
    If KeyInd = 2 Then KeyLeft = KeyCode
    If KeyInd = 3 Then KeyRight = KeyCode
    If KeyInd = 4 Then KeyGuns = KeyCode
    If KeyInd = 5 Then KeyBombs = KeyCode
    
    If KeyInd <> -1 Then
        If KeyCode = vbKeyUp Then
            KeyC(KeyInd).Caption = "Up Arrow"
        ElseIf KeyCode = vbKeyDown Then
            KeyC(KeyInd).Caption = "Down Arrow"
        ElseIf KeyCode = vbKeyLeft Then
            KeyC(KeyInd).Caption = "Left Arrow"
        ElseIf KeyCode = vbKeyRight Then
            KeyC(KeyInd).Caption = "Right Arrow"
        ElseIf KeyCode = vbKeySpace Then
            KeyC(KeyInd).Caption = "Space Bar"
        ElseIf KeyCode = vbKeyControl Then
            KeyC(KeyInd).Caption = "Control"
        ElseIf KeyCode = vbKeyShift Then
            KeyC(KeyInd).Caption = "Shift"
        ElseIf KeyCode = vbKeyCapital Then
            KeyC(KeyInd).Caption = "Caps Lock"
        ElseIf KeyCode = vbKeyBack Then
            KeyC(KeyInd).Caption = "Backspace"
        ElseIf KeyCode = vbKeyReturn Then
            KeyC(KeyInd).Caption = "Enter"
        ElseIf KeyCode = vbKeyTab Then
            KeyC(KeyInd).Caption = "Tab"
        ElseIf KeyCode = vbKeyCelete Then
            KeyC(KeyInd).Caption = "Delete"
        ElseIf KeyCode = vbKeyInsert Then
            KeyC(KeyInd).Caption = "Insert"
        ElseIf KeyCode = vbKeyHome Then
            KeyC(KeyInd).Caption = "Home"
        ElseIf KeyCode = vbKeyEnd Then
            KeyC(KeyInd).Caption = "End"
        ElseIf KeyCode = vbKeyPageUp Then
            KeyC(KeyInd).Caption = "Page Up"
        ElseIf KeyCode = vbKeyPageDown Then
            KeyC(KeyInd).Caption = "Page Down"
        ElseIf KeyCode = vbKeyNumlock Then
            KeyC(KeyInd).Caption = "Numlock"
        ElseIf KeyCode = vbKeyCivide Then
            KeyC(KeyInd).Caption = "/"
        ElseIf KeyCode = vbKeyMultiply Then
            KeyC(KeyInd).Caption = "*"
        ElseIf KeyCode = vbKeySubtract Then
            KeyC(KeyInd).Caption = "-"
        ElseIf KeyCode = vbKeyAdd Then
            KeyC(KeyInd).Caption = "+"
        ElseIf KeyCode = vbKeySelect Then
            KeyC(KeyInd).Caption = "Enter(numpad)"
        ElseIf KeyCode = vbKeyCecimaL Then
            KeyC(KeyInd).Caption = "."
        ElseIf KeyCode = vbKeyMenu Then
            KeyC(KeyInd).Caption = "Menu Key"
        ElseIf KeyCode = vbKeyNumpad0 Then
            KeyC(KeyInd).Caption = "Keypad 0"
        ElseIf KeyCode = vbKeyNumpad1 Then
            KeyC(KeyInd).Caption = "Keypad 1"
        ElseIf KeyCode = vbKeyNumpad2 Then
            KeyC(KeyInd).Caption = "Keypad 2"
        ElseIf KeyCode = vbKeyNumpad3 Then
            KeyC(KeyInd).Caption = "Keypad 3"
        ElseIf KeyCode = vbKeyNumpad4 Then
            KeyC(KeyInd).Caption = "Keypad 4"
        ElseIf KeyCode = vbKeyNumpad5 Then
            KeyC(KeyInd).Caption = "Keypad 5"
        ElseIf KeyCode = vbKeyNumpad6 Then
            KeyC(KeyInd).Caption = "Keypad 6"
        ElseIf KeyCode = vbKeyNumpad7 Then
            KeyC(KeyInd).Caption = "Keypad 7"
        ElseIf KeyCode = vbKeyNumpad8 Then
            KeyC(KeyInd).Caption = "Keypad 8"
        ElseIf KeyCode = vbKeyNumpad9 Then
            KeyC(KeyInd).Caption = "Keypad 9"
        Else
            KeyC(KeyInd).Caption = Chr(KeyCode)
        End If
        KeyC(KeyInd).BackStyle = 0
    End If
    For D = 0 To 5
        EQ(D).Enabled = True
    Next
    
    
    KeyInd = -1

End Sub

Private Sub Form_Load()
    KeyInd = -1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.BackColor = &H4000&
End Sub

Private Sub Label1_Click()
    DKeys.Visible = False
    ISLength = WritePrivateProfileString("Keys", "Up", KeyC(0).Caption & Chr$(0), App.Path & "\Stuff.ini")
    ISLength = WritePrivateProfileString("Keys", "Down", KeyC(1).Caption & Chr$(0), App.Path & "\Stuff.ini")
    ISLength = WritePrivateProfileString("Keys", "Left", KeyC(2).Caption & Chr$(0), App.Path & "\Stuff.ini")
    ISLength = WritePrivateProfileString("Keys", "Right", KeyC(3).Caption & Chr$(0), App.Path & "\Stuff.ini")
    ISLength = WritePrivateProfileString("Keys", "Guns", KeyC(4).Caption & Chr$(0), App.Path & "\Stuff.ini")
    ISLength = WritePrivateProfileString("Keys", "Bombs", KeyC(5).Caption & Chr$(0), App.Path & "\Stuff.ini")
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.BackColor = &HFF00&
End Sub

Private Sub Timer1_Timer()
    Shape1.BorderStyle = Timer1.Tag
    Timer1.Tag = Timer1.Tag + 1
    If Timer1.Tag = 6 Then Timer1.Tag = 4
End Sub
