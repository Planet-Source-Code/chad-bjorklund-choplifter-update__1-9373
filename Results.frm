VERSION 5.00
Begin VB.Form Results 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4710
   ClientLeft      =   2520
   ClientTop       =   4110
   ClientWidth     =   12105
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   4320
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      Picture         =   "Results.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      Picture         =   "Results.frx":0E52
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   9720
      Picture         =   "Results.frx":12E4
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "Results.frx":2136
      ScaleHeight     =   345
      ScaleWidth      =   1305
      TabIndex        =   2
      Top             =   3960
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "Results.frx":3438
      ScaleHeight     =   345
      ScaleWidth      =   1305
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1200
      Picture         =   "Results.frx":473A
      ScaleHeight     =   1185
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   240
      Width           =   10095
   End
   Begin VB.Label Died 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Died"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Index           =   3
      Left            =   6360
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Died 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Died"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   2
      Left            =   6960
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Done 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level Finished"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   3000
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Caption         =   "People Killed"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "People Saved"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Died 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Died"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   30.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Index           =   1
      Left            =   7200
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Died 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Died"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   0
      Left            =   7320
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Counts As Integer

Dim Atemp
Private Sub Form_Load()
If Label2.Caption + Label4.Caption <> 51 Then
    While Atemp < 4
        Died(Atemp).Visible = True
        Sleep (500)
        Atemp = Atemp + 1
    Wend
Else
    Done.Visible = True
    Picture4.Visible = True
    While Picture4.Left < 640
        Picture4.Left = Picture4.Left + 2
        Results.Refresh
    Wend
End If
Timer2.Enabled = True
End Sub


Private Sub Timer2_Timer()
Results.Refresh
Results.Visible = True
'Sleep (7000)
If TWidth <> 1152 And THeight <> 864 Then
            Dim DevM As DEVMODE
            erg& = EnumDisplaySettings(0&, 0&, DevM)
            DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
            DevM.dmPelsWidth = TWidth 'ScreenWidth
            DevM.dmPelsHeight = THeight 'ScreenHeight
            Select Case erg&
                Case DISP_CHANGE_RESTART
                    'an = MsgBox("You have to reboot(need 1152x864 res)", vbYesNo + vbSystemModal, "Info")
                    erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
                    If an = vbYes Then
                        erg& = ExitWindowsEx(EWX_REBOOT, 0&)
                    End If
                Case DISP_CHANGE_SUCCESSFUL
                    erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
                    'MsgBox "Everything's ok", vbOKOnly + vbSystemModal, "It worked!"
                Case Else
                    MsgBox "Mode Not supported(need 1152x864 res)", vbOKOnly + vbSystemModal, "Error"
            End Select
        End If
        '''''''''''''''''''''''''''''''''''''
    
    DeleteGeneratedDC PlaneB
        DeleteGeneratedDC PlaneW
            DeleteGeneratedDC LPlaneB
        DeleteGeneratedDC LPlaneW
    DeleteGeneratedDC RPlaneB
        DeleteGeneratedDC RPlaneW
            DeleteGeneratedDC RopeB
        DeleteGeneratedDC RopeW
    DeleteGeneratedDC BarrelB
        DeleteGeneratedDC BarrelW
            DeleteGeneratedDC BoomB
        DeleteGeneratedDC BoomW
    DeleteGeneratedDC GuyB
        DeleteGeneratedDC GuyW
            DeleteGeneratedDC DirtB
        DeleteGeneratedDC DirtW
    DeleteGeneratedDC TankB
        DeleteGeneratedDC TankW
            DeleteGeneratedDC GraveBW
        DeleteGeneratedDC WBoomW
    DeleteGeneratedDC WBoomB
        DeleteGeneratedDC SandB
            DeleteGeneratedDC WaterB
        DeleteGeneratedDC BG
            
            Unload Me
            Set Choplifter = Nothing
        End
End Sub
