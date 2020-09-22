VERSION 5.00
Begin VB.Form GamePlay 
   BackColor       =   &H00004000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.ProgBar ProgBar1 
      Height          =   2175
      Left            =   240
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   3836
      BackColour      =   16384
      BarStartColour  =   49152
      BarEndColour    =   255
      BorderStyle     =   0
      FillDirection   =   0
      Max             =   5
      Percent         =   40
      Value           =   2
      BarStyle        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BarEndColour    =   255
   End
   Begin VB.Label Label2 
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
      Left            =   5520
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "DIFFICULTY"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   1560
      Top             =   1800
      Width           =   735
   End
   Begin VB.Line Ln 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   15
      Index           =   4
      X1              =   1680
      X2              =   1920
      Y1              =   2160
      Y2              =   2400
   End
   Begin VB.Line Ln 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   15
      Index           =   5
      X1              =   1920
      X2              =   2160
      Y1              =   2400
      Y2              =   2160
   End
   Begin VB.Line Ln 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   15
      Index           =   3
      X1              =   1920
      X2              =   1920
      Y1              =   2280
      Y2              =   1920
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   1560
      Top             =   720
      Width           =   735
   End
   Begin VB.Line Ln 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   15
      Index           =   2
      X1              =   1920
      X2              =   2160
      Y1              =   840
      Y2              =   1080
   End
   Begin VB.Line Ln 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   15
      Index           =   1
      X1              =   1680
      X2              =   1920
      Y1              =   1080
      Y2              =   840
   End
   Begin VB.Line Ln 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   15
      Index           =   0
      X1              =   1920
      X2              =   1920
      Y1              =   1320
      Y2              =   840
   End
   Begin VB.Shape SHP 
      BorderColor     =   &H0080FF80&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H00008000&
      FillStyle       =   6  'Cross
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "GamePlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    If ProgBar1.Value <> ProgBar1.Max Then ProgBar1.Value = ProgBar1.Value + 1
End Sub
Private Sub Image2_Click()
    If ProgBar1.Value <> 1 Then ProgBar1.Value = ProgBar1.Value - 1
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    For f = 0 To 2
        Ln(f).DrawMode = 12
    Next
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    For f = 0 To 2
        Ln(f).DrawMode = 13
    Next
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    For f = 3 To 5
        Ln(f).DrawMode = 12
    Next
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    For f = 3 To 5
        Ln(f).DrawMode = 13
    Next
End Sub

Private Sub Label2_Click()
    GamePlay.Visible = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label2.BackColor = &HFF00&
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label2.BackColor = &H4000&
End Sub
