VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin GradientClockOCX.GradientClock GradientClock1 
      Height          =   2190
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   3863
      MajorPoint      =   0
      MinorPoint      =   0
      SecondPointer   =   0
      CircleBorder    =   0
      AntiAliasing    =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   2340
      X2              =   2340
      Y1              =   45
      Y2              =   2340
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2310
      X2              =   2310
      Y1              =   45
      Y2              =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTERED VERSION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   2565
      TabIndex        =   2
      Top             =   555
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   3765
      TabIndex        =   1
      Top             =   1830
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gradient Analog Clock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   2565
      TabIndex        =   0
      Top             =   255
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   330
      Left            =   3765
      Shape           =   4  'Rounded Rectangle
      Top             =   1770
      Width           =   1410
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   3810
      Shape           =   4  'Rounded Rectangle
      Top             =   1815
      Width           =   1410
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function GetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim myLS As Long, myTS As Long, myTL As Long

Private Sub Form_Load()
On Error Resume Next

   Dim hRgn As Long
   Dim hRgnRect As Long

   Label1(2) = "REGISTERED VERSION" & vbCrLf & "Made by. BatavianX" & vbCrLf & "Jakarta - Indonesia" & vbCrLf & "batavian.x@gmail.com"
   Me.ScaleMode = vbPixels
   
   hRgn = CreateEllipticRgn(0, 0, 157, 160)
   hRgnRect = CreateRectRgn(144, 0, Me.ScaleWidth, Me.ScaleHeight)
   
   CombineRgn hRgn, hRgn, hRgnRect, 2
   
   SetWindowRgn Me.hwnd, hRgn, False
   
   DeleteObject hRgn
   DeleteObject hRgnRect
   
   myLS = Shape1.Left
   myTS = Shape1.Top
   myTL = Label1(1).Top

On Error GoTo 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shape1.BackColor = vbRed Then
      Shape1.BackColor = vbBlack
   End If
End Sub

Private Sub Form_Paint()
   Dim hRgn As Long
   Dim hBrush As Long
   Dim hRgnRect As Long

   hRgn = CreateEllipticRgn(0, 0, 157, 160)
   hRgnRect = CreateRectRgn(144, 0, Me.ScaleWidth, Me.ScaleHeight)
   CombineRgn hRgn, hRgn, hRgnRect, 2
   hBrush = CreateSolidBrush(vbWhite)
   FrameRgn Me.hDC, hRgn, hBrush, 1, 1
   
   DeleteObject hRgn
   DeleteObject hRgnRect
   DeleteObject hBrush
End Sub

Private Sub Label1_Click(Index As Integer)
   If Index = 1 Then
      Unload Me
   End If
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 1 Then
      Shape1.Move Shape2.Left, Shape2.Top
      Label1(1).Move Shape2.Left, Shape1.Top + (60 / Screen.TwipsPerPixelY)
   End If
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 1 Then
      Shape1.BackColor = vbRed
   End If
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 1 Then
      Shape1.Move myLS, myTS
      Label1(1).Move myLS, myTL
   End If
End Sub
