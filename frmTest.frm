VERSION 5.00
Object = "{89B71FA8-59F5-4844-8E2E-FBCAAF44893E}#17.0#0"; "ClockOCX.ocx"
Begin VB.Form Form1 
   Caption         =   "  "
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin ClockOCX.AnalogClock AnalogClock1 
      Height          =   9840
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   17357
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AnalogClock1_DblClick(Index As Integer)
   AnalogClock1(Index).About
End Sub

Private Sub Form_Load()

End Sub
