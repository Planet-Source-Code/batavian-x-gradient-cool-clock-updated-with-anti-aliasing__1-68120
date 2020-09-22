VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   2430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOption 
      BackColor       =   &H00000000&
      Caption         =   "Anti-Aliasing                            (Robert Rayment (RRPaint))"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   7
      Left            =   120
      TabIndex        =   26
      Top             =   4605
      Width           =   2940
   End
   Begin ClockOCX_Test.AnalogClock Clock 
      Height          =   2205
      Left            =   120
      TabIndex        =   25
      ToolTipText     =   "Click Me!"
      Top             =   135
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   3889
      MinuteOutline   =   14737632
      HourOutline     =   14737632
      MajorPoint      =   0
      MinorPoint      =   0
      SecondPointer   =   0
      CircleBorder    =   0
      DrawShadow      =   0   'False
   End
   Begin VB.PictureBox picColor 
      Height          =   345
      Index           =   8
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   24
      Top             =   8400
      Width           =   360
   End
   Begin VB.PictureBox picColor 
      Height          =   345
      Index           =   7
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   21
      Top             =   7590
      Width           =   360
   End
   Begin VB.PictureBox picColor 
      Height          =   345
      Index           =   6
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   20
      Top             =   7995
      Width           =   360
   End
   Begin VB.PictureBox picColor 
      Height          =   345
      Index           =   5
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   18
      Top             =   7185
      Width           =   360
   End
   Begin VB.PictureBox picColor 
      Height          =   345
      Index           =   4
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   16
      Top             =   6780
      Width           =   360
   End
   Begin VB.PictureBox picColor 
      Height          =   345
      Index           =   3
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   14
      Top             =   6375
      Width           =   360
   End
   Begin VB.PictureBox picColor 
      Height          =   345
      Index           =   2
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   12
      Top             =   5970
      Width           =   360
   End
   Begin VB.PictureBox picColor 
      Height          =   345
      Index           =   1
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   5565
      Width           =   360
   End
   Begin VB.PictureBox picColor 
      Height          =   345
      Index           =   0
      Left            =   105
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   8
      Top             =   5160
      Width           =   360
   End
   Begin VB.CheckBox chkOption 
      BackColor       =   &H00000000&
      Caption         =   "Draw Pointers Shadow"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   4215
      Width           =   2940
   End
   Begin VB.CheckBox chkOption 
      BackColor       =   &H00000000&
      Caption         =   "Draw Minor Point"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   3915
      Value           =   1  'Checked
      Width           =   2940
   End
   Begin VB.CheckBox chkOption 
      BackColor       =   &H00000000&
      Caption         =   "Draw Major Point"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   3615
      Value           =   1  'Checked
      Width           =   2940
   End
   Begin VB.CheckBox chkOption 
      BackColor       =   &H00000000&
      Caption         =   "Draw Second Pointer"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   3315
      Value           =   1  'Checked
      Width           =   2940
   End
   Begin VB.CheckBox chkOption 
      BackColor       =   &H00000000&
      Caption         =   "Draw Minute Outline"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   3015
      Value           =   1  'Checked
      Width           =   2940
   End
   Begin VB.CheckBox chkOption 
      BackColor       =   &H00000000&
      Caption         =   "Draw Hour Outline"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2715
      Value           =   1  'Checked
      Width           =   2940
   End
   Begin VB.CheckBox chkOption 
      BackColor       =   &H00000000&
      Caption         =   "Draw Clock Outline"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2415
      Value           =   1  'Checked
      Width           =   2940
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clock Outline Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   555
      TabIndex        =   0
      Top             =   8460
      Width           =   1575
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hour Outline Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   555
      TabIndex        =   23
      Top             =   7650
      Width           =   1530
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minute Outline Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   555
      TabIndex        =   22
      Top             =   8055
      Width           =   1710
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clock Body Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   555
      TabIndex        =   19
      Top             =   7245
      Width           =   1395
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minor Point Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   555
      TabIndex        =   17
      Top             =   6840
      Width           =   1440
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Major Point Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   555
      TabIndex        =   15
      Top             =   6435
      Width           =   1455
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Second Pointer Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   555
      TabIndex        =   13
      Top             =   6030
      Width           =   1755
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minute Pointer Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   555
      TabIndex        =   11
      Top             =   5625
      Width           =   1725
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hour Pointer Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   555
      TabIndex        =   9
      Top             =   5220
      Width           =   1545
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUp 
         Caption         =   "Load Form2"
         Index           =   0
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "Exit"
         Index           =   2
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "About"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOption_Click(Index As Integer)
   Select Case Index
   Case 0: Clock.DrawBodyOutline = chkOption(0).Value
   Case 1: Clock.DrawHourOutline = chkOption(1).Value
   Case 2: Clock.DrawMinuteOutline = chkOption(2).Value
   Case 3: Clock.DrawSecond = chkOption(3).Value
   Case 4: Clock.ShowMajorPoint = chkOption(4).Value
   Case 5: Clock.ShowMinorPoint = chkOption(5).Value
   Case 6: Clock.DrawShadow = chkOption(6).Value
   Case 7: Clock.AntiAliasing = chkOption(7).Value
   End Select
End Sub

Private Sub Clock_Click()
   Me.PopupMenu mnuHidden
End Sub

Private Sub Form_Load()
   picColor(0).BackColor = Clock.HourPointer
   picColor(1).BackColor = Clock.MinutePointer
   picColor(2).BackColor = Clock.SecondPointer
   picColor(3).BackColor = Clock.MajorPoint
   picColor(4).BackColor = Clock.MinorPoint
   picColor(5).BackColor = Clock.ClockBody
   picColor(6).BackColor = Clock.HourOutline
   picColor(7).BackColor = Clock.MinuteOutline
   picColor(8).BackColor = Clock.CircleBorder
End Sub

Private Sub mnuPopUp_Click(Index As Integer)
   Select Case Index
      Case 0: Form2.Show vbModeless, Me
      Case 2: Unload Me
      Case 3: Clock.About
   End Select
End Sub

Private Sub picColor_Click(Index As Integer)
   Dim CDSelColor As SelectedColor
   
   CDSelColor = ShowColor(Me.hwnd, False, picColor(Index).BackColor)
   
   If CDSelColor.bCanceled = False Then
      picColor(Index).BackColor = CDSelColor.oSelectedColor
   End If
   
   Select Case Index
   Case 0: Clock.HourPointer = picColor(0).BackColor
   Case 1: Clock.MinutePointer = picColor(1).BackColor
   Case 2: Clock.SecondPointer = picColor(2).BackColor
   Case 3: Clock.MajorPoint = picColor(3).BackColor
   Case 4: Clock.MinorPoint = picColor(4).BackColor
   Case 5: Clock.ClockBody = picColor(5).BackColor
   Case 6: Clock.HourOutline = picColor(6).BackColor
   Case 7: Clock.MinuteOutline = picColor(7).BackColor
   Case 8: Clock.CircleBorder = picColor(8).BackColor
   End Select
End Sub
