VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Crash at maximized in your system?"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3720
   LinkTopic       =   "Form2"
   ScaleHeight     =   3330
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Circle"
      Height          =   330
      Left            =   30
      TabIndex        =   1
      Top             =   -30
      Width           =   720
   End
   Begin ClockOCX_Test.AnalogClock Clock 
      Height          =   2160
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   3810
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
   Call Form_Resize
End Sub

Private Sub Form_Load()
   Me.WindowState = vbMaximized
End Sub

Private Sub Form_Resize()
   If Me.WindowState <> vbMinimized Then
      If Check1.Value = vbChecked Then
         If Me.ScaleHeight >= Me.ScaleWidth Then
            Clock.Width = Me.ScaleWidth
            Clock.Height = Me.ScaleWidth
         Else
            Clock.Width = Me.ScaleHeight
            Clock.Height = Me.ScaleHeight
         End If
         
         Clock.Left = (Me.ScaleWidth - Clock.Width) / 2
         Clock.Top = (Me.ScaleHeight - Clock.Height) / 2
      Else
         Clock.Width = Me.ScaleWidth
         Clock.Height = Me.ScaleHeight
         Clock.Left = 0
         Clock.Top = 0
      End If
   End If
End Sub
