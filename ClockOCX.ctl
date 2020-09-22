VERSION 5.00
Begin VB.UserControl AnalogClock 
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PaletteMode     =   2  'Custom
   PropertyPages   =   "ClockOCX.ctx":0000
   ScaleHeight     =   375
   ScaleWidth      =   360
   ToolboxBitmap   =   "ClockOCX.ctx":0014
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2595
      Top             =   2625
   End
End
Attribute VB_Name = "AnalogClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Const ALTERNATE As Long = 1
Private Const Pi As Double = 3.14159265358979
Private Const WINDING As Long = 2

Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private m_DrawHourOutline As Boolean
Private m_DrawMinuteOutline As Boolean
Private m_DrawShadow As Boolean
Private m_DrawBodyOutline As Boolean
Private m_DrawSecond As Boolean
Private m_ShowMinorPoint As Boolean
Private m_ShowMajorPoint As Boolean
Private m_MinuteOutline As OLE_COLOR
Private m_HourOutline As OLE_COLOR
Private m_MajorPoint As OLE_COLOR
Private m_MinorPoint As OLE_COLOR
Private m_SecondPointer As OLE_COLOR
Private m_MinutePointer As OLE_COLOR
Private m_HourPointer As OLE_COLOR
Private m_CircleBorder As OLE_COLOR
Private m_ClockBody As OLE_COLOR
Private m_AntiAliasing As Boolean

Private m_button As Integer

Private CenterX As Long
Private CenterY As Long

Public Event Click()
Public Event DblClick()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove

Private Function Dec2Rad(ByVal dblDec As Double) As Double 'Convert Decimal To Radian, I don't know math that well, so I got it from some where
   Dim dRad As Double

   dRad = Pi / 180
   Dec2Rad = dblDec * dRad
End Function

Private Sub ShowTime()
   Dim dH As Long, dH1 As Long, dH2 As Long, iH As Integer 'Hour Variables
   Dim dM As Long, dM1 As Long, dM2 As Long, iM As Integer 'Minute Variables
   Dim dS As Long 'Second Variable
   
   Dim dHX As Double, dHX1 As Double, dHX2 As Double 'Hour Variables
   Dim dHY As Double, dHY1 As Double, dHY2 As Double
   
   Dim dMX As Double, dMX1 As Double, dMX2 As Double 'Minute Variables
   Dim dMY As Double, dMY1 As Double, dMY2 As Double
   
   Dim dSX As Double, dSY As Double 'Second Variables
   
   Dim pP(1 To 4) As POINTAPI 'The Polygons Points
   Dim hRgn As Long, hBrush As Long 'Fill Color
   
   Dim fW As Long, fH As Long
   Dim iCircle As Integer, r As Integer, g As Integer, B As Integer
   
   fW = UserControl.ScaleWidth
   fH = UserControl.ScaleHeight
   
   iH = Hour(Time)                                    '| Get the Current Hour
   If iH > 12 Then iH = iH - 12                       '| Make it 12 Hour Format
   If iH = 0 Then iH = 12                             '|
   dH = (iH * 30) + (Int(Minute(Time) / 12) * 6)      '| Hour Original
   dH1 = dH - 40                                      '| Hour Outer
   dH2 = dH + 40                                      '| Hour Inner
   
   iM = Minute(Time)                                  '| Get the Current Minute
   If iM = 0 Then iM = 60                             '|
   dM = iM * 6                                        '| Minute Original
   dM1 = dM - 40                                      '| Minute Outer
   dM2 = dM + 40                                      '| Minute Inner
   
   dS = Int(Timer) * 6                                '| Second Code

   dHX = Sin(Dec2Rad(dH))                             '| Hour Point.X
   dHY = -Cos(Dec2Rad(dH))                            '| Hour Point.Y
   dHX1 = Sin(Dec2Rad(dH1))                           '| Hour Left Point.X
   dHY1 = -Cos(Dec2Rad(dH1))                          '| Hour Left Point.Y
   dHX2 = Sin(Dec2Rad(dH2))                           '| Hour Right Point.X
   dHY2 = -Cos(Dec2Rad(dH2))                          '| Hour Right Point.Y
                                                            
   dMX = Sin(Dec2Rad(dM))                             '| Minute Point.X
   dMY = -Cos(Dec2Rad(dM))                            '| Minute Point.Y
   dMX1 = Sin(Dec2Rad(dM1))                           '| Minute Left Point.X
   dMY1 = -Cos(Dec2Rad(dM1))                          '| Minute Left Point.Y
   dMX2 = Sin(Dec2Rad(dM2))                           '| Minute Right Point.X
   dMY2 = -Cos(Dec2Rad(dM2))                          '| Minute Right Point.Y
   
   dSX = Sin(Dec2Rad(dS))                             '| The Second
   dSY = -Cos(Dec2Rad(dS))                            '|
   
   UserControl.Cls                                    '| Clear the Form
   UserControl.DrawStyle = 5                          '| Set to Transparent
   UserControl.FillStyle = 0                          '| Set to Solid
   
   hRgn = CreateEllipticRgn(2, 2, CenterX + CenterX, CenterY + CenterY)   '| Clock's Region >>------\
   hBrush = CreateSolidBrush(m_ClockBody)             '|                                            |
   FillRgn UserControl.hdc, hRgn, hBrush              '| Fill the Clock, I think it's not necessary |
   DeleteObject hBrush                                '|                                            |
   Draw_GradientCircle m_ClockBody                    '| Draw circle gradient style                 |
   UserControl.DrawStyle = 0                          '| Set Back to Solid                          |
   UserControl.FillStyle = 1                          '| Set Back to Transparent                    |
                                                      '|                                            |
   hBrush = CreateSolidBrush(IIf(m_DrawBodyOutline, m_CircleBorder, m_ClockBody))                  '|
   FrameRgn UserControl.hdc, hRgn, hBrush, 1, 1       '| Draw the Frame of the Region <<------------/
   DeleteObject hBrush
   DeleteObject hRgn
      
   '===========================================
   '=== The Minute's Pointer Polygon Points ===
   '===========================================
   pP(1).X = (dMX * -Round(fW / 19)) + CenterX
   pP(1).Y = (dMY * -Round(fH / 19)) + CenterY
   pP(2).X = (dMX1 * Round(fW / 19)) + pP(1).X
   pP(2).Y = (dMY1 * Round(fH / 19)) + pP(1).Y
   pP(3).X = (dMX * Round(fW / 2.2)) + pP(1).X
   pP(3).Y = (dMY * Round(fH / 2.2)) + pP(1).Y
   pP(4).X = (dMX2 * Round(fW / 19)) + pP(1).X
   pP(4).Y = (dMY2 * Round(fH / 19)) + pP(1).Y
   '===========================================
   
   hRgn = CreatePolygonRgn(pP(1), 4, WINDING)         '| Create the Minute Region
   
   If m_DrawShadow Then
      OffsetRgn hRgn, 2, 2                            '| Shadow First  <<---------------------------\
      hBrush = CreateSolidBrush(RGB(127, 127, 127))   '| Create a Brush Handle with specified color |
      FillRgn UserControl.hdc, hRgn, hBrush           '| Fill the Shadow                            |
      OffsetRgn hRgn, -2, -2                          '| Then the Pointer  <<-----------------------/
      DeleteObject hBrush                             '| RELEASE THE MEMORY HANDLE, TO AVOID GDI MEMORY LEAK *) That's what they always said! :p
      
      If m_AntiAliasing Then
         AALINE pP(1).X + 2, pP(1).Y + 2, pP(2).X + 2, pP(2).Y + 2, RGB(127, 127, 127)
         AALINE pP(2).X + 2, pP(2).Y + 2, pP(3).X + 2, pP(3).Y + 2, RGB(127, 127, 127)
         AALINE pP(3).X + 2, pP(3).Y + 2, pP(4).X + 2, pP(4).Y + 2, RGB(127, 127, 127)
         AALINE pP(4).X + 2, pP(4).Y + 2, pP(1).X + 2, pP(1).Y + 2, RGB(127, 127, 127)
      End If
   End If
   
   hBrush = CreateSolidBrush(m_MinutePointer)
   FillRgn UserControl.hdc, hRgn, hBrush              '| Fill the Minute Region
   DeleteObject hBrush
   
   If m_AntiAliasing Then
      AALINE pP(1).X, pP(1).Y, pP(2).X, pP(2).Y, IIf(m_DrawMinuteOutline, m_MinuteOutline, m_MinutePointer)
      AALINE pP(2).X, pP(2).Y, pP(3).X, pP(3).Y, IIf(m_DrawMinuteOutline, m_MinuteOutline, m_MinutePointer)
      AALINE pP(3).X, pP(3).Y, pP(4).X, pP(4).Y, IIf(m_DrawMinuteOutline, m_MinuteOutline, m_MinutePointer)
      AALINE pP(4).X, pP(4).Y, pP(1).X, pP(1).Y, IIf(m_DrawMinuteOutline, m_MinuteOutline, m_MinutePointer)
   Else
      If m_DrawMinuteOutline Then
         hBrush = CreateSolidBrush(m_MinuteOutline)
         FrameRgn UserControl.hdc, hRgn, hBrush, 1, 1    '| Draw the Frame of the Minute Region
         DeleteObject hBrush
      End If
   End If
   
   DeleteObject hRgn
   
   '===========================================
   '==== The Hour's Pointer Polygon Points ====
   '===========================================
   pP(1).X = (dHX * -Round(fW / 19)) + CenterX
   pP(1).Y = (dHY * -Round(fH / 19)) + CenterY
   pP(2).X = (dHX1 * Round(fW / 19)) + pP(1).X
   pP(2).Y = (dHY1 * Round(fH / 19)) + pP(1).Y
   pP(3).X = (dHX * Round(fW / 2.8)) + pP(1).X
   pP(3).Y = (dHY * Round(fH / 2.8)) + pP(1).Y
   pP(4).X = (dHX2 * Round(fW / 19)) + pP(1).X
   pP(4).Y = (dHY2 * Round(fH / 19)) + pP(1).Y
   '===========================================

   hRgn = CreatePolygonRgn(pP(1), 4, WINDING)          '| Create the Hour's Region
   
   If m_DrawShadow Then
      OffsetRgn hRgn, 2, 2                             '| Shadow First  <<---------------------------\
      hBrush = CreateSolidBrush(RGB(127, 127, 127))    '| Create a Brush Handle with specified color |
      FillRgn UserControl.hdc, hRgn, hBrush            '| Fill the Shadow                            |
      OffsetRgn hRgn, -2, -2                           '| Then the Pointer  <<-----------------------/
      DeleteObject hBrush                              '| RELEASE THE MEMORY HANDLE, TO AVOID GDI MEMORY LEAK *) That's what they always said! :p
      
      If m_AntiAliasing Then
         AALINE pP(1).X + 2, pP(1).Y + 2, pP(2).X + 2, pP(2).Y + 2, RGB(127, 127, 127)
         AALINE pP(2).X + 2, pP(2).Y + 2, pP(3).X + 2, pP(3).Y + 2, RGB(127, 127, 127)
         AALINE pP(3).X + 2, pP(3).Y + 2, pP(4).X + 2, pP(4).Y + 2, RGB(127, 127, 127)
         AALINE pP(4).X + 2, pP(4).Y + 2, pP(1).X + 2, pP(1).Y + 2, RGB(127, 127, 127)
      End If
   End If
   
   hBrush = CreateSolidBrush(m_HourPointer)
   FillRgn UserControl.hdc, hRgn, hBrush               '| Fill the Hour Region
   DeleteObject hBrush
   
   If m_AntiAliasing Then
      AALINE pP(1).X, pP(1).Y, pP(2).X, pP(2).Y, IIf(m_DrawHourOutline, m_HourOutline, m_HourPointer)
      AALINE pP(2).X, pP(2).Y, pP(3).X, pP(3).Y, IIf(m_DrawHourOutline, m_HourOutline, m_HourPointer)
      AALINE pP(3).X, pP(3).Y, pP(4).X, pP(4).Y, IIf(m_DrawHourOutline, m_HourOutline, m_HourPointer)
      AALINE pP(4).X, pP(4).Y, pP(1).X, pP(1).Y, IIf(m_DrawHourOutline, m_HourOutline, m_HourPointer)
   Else
      If m_DrawHourOutline Then
         hBrush = CreateSolidBrush(m_HourOutline)
         FrameRgn UserControl.hdc, hRgn, hBrush, 1, 1  '| Draw the Frame of the Hour Region
         DeleteObject hBrush
      End If
   End If
   
   DeleteObject hRgn
   
   If m_DrawSecond Then
      If m_DrawShadow Then
         If m_AntiAliasing Then
            AALINE (dSX * -Round(fW / 10)) + CenterX + 2, (dSY * -Round(fH / 10)) + CenterY + 2, _
                   (dSX * Round(fW / 2.5)) + CenterX + 2, (dSY * Round(fH / 2.5)) + CenterY + 2, _
                   RGB(127, 127, 127)
         Else
            UserControl.Line ((dSX * -Round(fW / 10)) + CenterX + 2, (dSY * -Round(fH / 10)) + CenterY + 2)- _
                             ((dSX * Round(fW / 2.5)) + CenterX + 2, (dSY * Round(fH / 2.5)) + CenterY + 2), _
                             RGB(127, 127, 127)              '| Create a Shadow of the Second Pointer
         End If
      End If
                             
      If m_AntiAliasing Then
         AALINE (dSX * -Round(fW / 10)) + CenterX, (dSY * -Round(fH / 10)) + CenterY, _
                (dSX * Round(fW / 2.5)) + CenterX, (dSY * Round(fH / 2.5)) + CenterY, _
                m_SecondPointer
      Else
         UserControl.Line ((dSX * -Round(fW / 10)) + CenterX, (dSY * -Round(fH / 10)) + CenterY)- _
                          ((dSX * Round(fW / 2.5)) + CenterX, (dSY * Round(fH / 2.5)) + CenterY), _
                          m_SecondPointer                    '| Now Create the Simple Second Pointer
      End If
   End If
   
   UserControl.Circle (CenterX, CenterY), 0, m_HourPointer   '| Draw the Pointers Axis
   UserControl.Circle (CenterX, CenterY), 1, m_HourPointer   '| Draw the Pointers Axis, again...?
   
   For iCircle = 6 To 360 Step 6                             '| Draw the Points
      dHX = Sin(Dec2Rad(iCircle))                            '| Zzz...Zzz...Zzz...
      dHY = -(Cos(Dec2Rad(iCircle)))
      dHX1 = Sin(Dec2Rad(iCircle - 0.25))
      dHY1 = -(Cos(Dec2Rad(iCircle - 0.25)))
      dHX2 = Sin(Dec2Rad(iCircle + 0.25))
      dHY2 = -(Cos(Dec2Rad(iCircle + 0.25)))
      
      If m_ShowMajorPoint Then
         If iCircle Mod 30 = 0 Then
            If iCircle Mod 90 = 0 Then
               If m_AntiAliasing Then
                  AALINE (dHX1 * Round(fW / 2.1)) + CenterX, (dHY1 * Round(fH / 2.1)) + CenterY, _
                         (dHX1 * Round(fW / 2.3)) + CenterX, (dHY1 * Round(fH / 2.3)) + CenterY, m_MajorPoint
               Else
                  UserControl.Line ((dHX1 * Round(fW / 2.1)) + CenterX, (dHY1 * Round(fH / 2.1)) + CenterY)- _
                                   ((dHX1 * Round(fW / 2.3)) + CenterX, (dHY1 * Round(fH / 2.3)) + CenterY), m_MajorPoint
               End If
            Else
               If m_AntiAliasing Then
                  AALINE (dHX * Round(fW / 2.1)) + CenterX, (dHY * Round(fH / 2.1)) + CenterY, _
                         (dHX * Round(fW / 2.3)) + CenterX, (dHY * Round(fH / 2.3)) + CenterY, m_MajorPoint
               Else
                  UserControl.Line ((dHX * Round(fW / 2.1)) + CenterX, (dHY * Round(fH / 2.1)) + CenterY)- _
                                   ((dHX * Round(fW / 2.3)) + CenterX, (dHY * Round(fH / 2.3)) + CenterY), m_MajorPoint
               End If
            End If
         Else
            If m_ShowMinorPoint Then
               UserControl.Circle ((dHX * Round(fW / 2.2)) + CenterX, (dHY * Round(fH / 2.2)) + CenterY), 0, m_MinorPoint
            End If
         End If
      Else
         If m_ShowMinorPoint Then
            UserControl.Circle ((dHX * Round(fW / 2.2)) + CenterX, (dHY * Round(fH / 2.2)) + CenterY), 0, m_MinorPoint
         End If
      End If
   Next iCircle
   
   If m_AntiAliasing Then
      AAELLIPSE CenterX, CenterY, CenterX - 3, CenterY - 3, IIf(m_DrawBodyOutline, m_CircleBorder, m_ClockBody)
   End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
   UserControl.Refresh
End Sub

Private Sub Timer1_Timer()
   Static lngSecond As Long
   
   If lngSecond <> Second(Time) Then
      ShowTime
      lngSecond = Second(Time)
   End If
End Sub

Private Sub UserControl_Initialize()
   UserControl.ScaleMode = vbPixels
   UserControl.Refresh
   ShowTime
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_button = Button
End Sub

Private Sub UserControl_Resize()
   Dim hRgn As Long

   CenterX = UserControl.ScaleWidth / 2
   CenterY = UserControl.ScaleHeight / 2
   
   hRgn = CreateEllipticRgn(2, 2, CenterX + CenterX, CenterY + CenterY)   'Clock's Face
   SetWindowRgn UserControl.hwnd, hRgn, True                              'Set to a Circular Form (Not a Rectangle :)

   ShowTime
End Sub

Private Sub UserControl_Show()
   ShowTime
End Sub

Private Sub UserControl_InitProperties()
   Timer1.Enabled = Ambient.UserMode
   m_ShowMajorPoint = True
   m_MinuteOutline = vbWhite
   m_HourOutline = vbWhite
   m_MajorPoint = vbWhite
   m_MinorPoint = vbWhite
   m_SecondPointer = vbWhite
   m_MinutePointer = vbBlack
   m_HourPointer = vbBlack
   m_CircleBorder = vbWhite
   m_ClockBody = vbBlack
   m_ShowMinorPoint = True
   m_DrawHourOutline = True
   m_DrawMinuteOutline = True
   m_DrawShadow = True
   m_DrawBodyOutline = True
   m_DrawSecond = True
   m_AntiAliasing = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Timer1.Enabled = Ambient.UserMode
   m_ShowMajorPoint = PropBag.ReadProperty("ShowMajorPoint", True)
   m_MinuteOutline = PropBag.ReadProperty("MinuteOutline", vbWhite)
   m_HourOutline = PropBag.ReadProperty("HourOutline", vbWhite)
   m_MajorPoint = PropBag.ReadProperty("MajorPoint", vbWhite)
   m_MinorPoint = PropBag.ReadProperty("MinorPoint", vbWhite)
   m_SecondPointer = PropBag.ReadProperty("SecondPointer", vbWhite)
   m_MinutePointer = PropBag.ReadProperty("MinutePointer", vbBlack)
   m_HourPointer = PropBag.ReadProperty("HourPointer", vbBlack)
   m_CircleBorder = PropBag.ReadProperty("CircleBorder", vbWhite)
   m_ClockBody = PropBag.ReadProperty("ClockBody", vbBlack)
   m_ShowMinorPoint = PropBag.ReadProperty("ShowMinorPoint", True)
   m_DrawHourOutline = PropBag.ReadProperty("DrawHourOutline", True)
   m_DrawMinuteOutline = PropBag.ReadProperty("DrawMinuteOutline", True)
   m_DrawShadow = PropBag.ReadProperty("DrawShadow", True)
   m_DrawBodyOutline = PropBag.ReadProperty("DrawBodyOutline", True)
   m_DrawSecond = PropBag.ReadProperty("DrawSecond", True)
   m_AntiAliasing = PropBag.ReadProperty("AntiAliasing", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "ShowMajorPoint", m_ShowMajorPoint, True
      .WriteProperty "MinuteOutline", m_MinuteOutline, vbWhite
      .WriteProperty "HourOutline", m_HourOutline, vbWhite
      .WriteProperty "MajorPoint", m_MajorPoint, vbWhite
      .WriteProperty "MinorPoint", m_MinorPoint, vbWhite
      .WriteProperty "SecondPointer", m_SecondPointer, vbWhite
      .WriteProperty "MinutePointer", m_MinutePointer, vbBlack
      .WriteProperty "HourPointer", m_HourPointer, vbBlack
      .WriteProperty "CircleBorder", m_CircleBorder, vbWhite
      .WriteProperty "ClockBody", m_ClockBody, vbBlack
      .WriteProperty "ShowMinorPoint", m_ShowMinorPoint, True
      .WriteProperty "DrawHourOutline", m_DrawHourOutline, True
      .WriteProperty "DrawMinuteOutline", m_DrawMinuteOutline, True
      .WriteProperty "DrawShadow", m_DrawShadow, True
      .WriteProperty "DrawBodyOutline", m_DrawBodyOutline, True
      .WriteProperty "DrawSecond", m_DrawSecond, True
      .WriteProperty "AntiAliasing", m_AntiAliasing, False
   End With
End Sub

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
   hwnd = UserControl.hwnd
End Property

Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
   hdc = UserControl.hdc
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
   If m_button = vbLeftButton Then
      RaiseEvent Click
   End If
End Sub

Private Sub UserControl_DblClick()
   If m_button = vbLeftButton Then
      RaiseEvent DblClick
   End If
End Sub

Public Property Get ShowMajorPoint() As Boolean
Attribute ShowMajorPoint.VB_Description = "Returns/sets whether major points are drawn."
   ShowMajorPoint = m_ShowMajorPoint
End Property

Public Property Let ShowMajorPoint(ByVal New_ShowMajorPoint As Boolean)
   m_ShowMajorPoint = New_ShowMajorPoint
   PropertyChanged "ShowMajorPoint"
   ShowTime
End Property

Public Property Get ShowMinorPoint() As Boolean
Attribute ShowMinorPoint.VB_Description = "Returns/sets whether minor points are drawn."
   ShowMinorPoint = m_ShowMinorPoint
End Property

Public Property Let ShowMinorPoint(ByVal New_ShowMinorPoint As Boolean)
   m_ShowMinorPoint = New_ShowMinorPoint
   PropertyChanged "ShowMinorPoint"
   ShowTime
End Property

Public Property Get MajorPoint() As OLE_COLOR
Attribute MajorPoint.VB_Description = "Returns/sets the color of major points."
   MajorPoint = m_MajorPoint
End Property

Public Property Let MajorPoint(ByVal New_MajorPoint As OLE_COLOR)
   m_MajorPoint = New_MajorPoint
   PropertyChanged "MajorPoint"
   ShowTime
End Property

Public Property Get MinorPoint() As OLE_COLOR
Attribute MinorPoint.VB_Description = "Returns/sets the color of minor points."
   MinorPoint = m_MinorPoint
End Property

Public Property Let MinorPoint(ByVal New_MinorPoint As OLE_COLOR)
   m_MinorPoint = New_MinorPoint
   PropertyChanged "MinorPoint"
   ShowTime
End Property

Public Property Get SecondPointer() As OLE_COLOR
   SecondPointer = m_SecondPointer
End Property

Public Property Let SecondPointer(ByVal New_SecondPointer As OLE_COLOR)
Attribute SecondPointer.VB_Description = "Returns/sets whether major points are drawn."
   m_SecondPointer = New_SecondPointer
   PropertyChanged "SecondPointer"
   ShowTime
End Property

Public Property Get MinutePointer() As OLE_COLOR
Attribute MinutePointer.VB_Description = "Returns/sets the minute pointer's color."
   MinutePointer = m_MinutePointer
End Property

Public Property Let MinutePointer(ByVal New_MinutePointer As OLE_COLOR)
   m_MinutePointer = New_MinutePointer
   PropertyChanged "MinutePointer"
   ShowTime
End Property

Public Property Get HourPointer() As OLE_COLOR
Attribute HourPointer.VB_Description = "Returns/sets the hour pointer's color."
   HourPointer = m_HourPointer
End Property

Public Property Let HourPointer(ByVal New_HourPointer As OLE_COLOR)
   m_HourPointer = New_HourPointer
   PropertyChanged "HourPointer"
   ShowTime
End Property

Public Property Get CircleBorder() As OLE_COLOR
Attribute CircleBorder.VB_Description = "Returns/sets the clock outline's color."
   CircleBorder = m_CircleBorder
End Property

Public Property Let CircleBorder(ByVal New_CircleBorder As OLE_COLOR)
   m_CircleBorder = New_CircleBorder
   PropertyChanged "CircleBorder"
   ShowTime
End Property

Public Property Get ClockBody() As OLE_COLOR
Attribute ClockBody.VB_Description = "Returns/sets the clock body's color."
   ClockBody = m_ClockBody
End Property

Public Property Let ClockBody(ByVal New_ClockBody As OLE_COLOR)
   m_ClockBody = New_ClockBody
   PropertyChanged "ClockBody"
   ShowTime
End Property

Public Property Get HourOutline() As OLE_COLOR
Attribute HourOutline.VB_Description = "Returns/sets hour pointer's outline color."
   HourOutline = m_HourOutline
End Property

Public Property Let HourOutline(ByVal New_HourOutline As OLE_COLOR)
   m_HourOutline = New_HourOutline
   PropertyChanged "HourOutline"
   ShowTime
End Property

Public Property Get MinuteOutline() As OLE_COLOR
Attribute MinuteOutline.VB_Description = "Returns/sets minute pointer's outline color."
   MinuteOutline = m_MinuteOutline
End Property

Public Property Let MinuteOutline(ByVal New_MinuteOutline As OLE_COLOR)
   m_MinuteOutline = New_MinuteOutline
   PropertyChanged "MinuteOutline"
   ShowTime
End Property

Public Property Get DrawHourOutline() As Boolean
Attribute DrawHourOutline.VB_Description = "Returns/sets whether outline on hour pointer is drawn."
   DrawHourOutline = m_DrawHourOutline
End Property

Public Property Let DrawHourOutline(ByVal New_DrawHourOutline As Boolean)
   m_DrawHourOutline = New_DrawHourOutline
   PropertyChanged "DrawHourOutline"
   ShowTime
End Property

Public Property Get DrawMinuteOutline() As Boolean
Attribute DrawMinuteOutline.VB_Description = "Returns/sets whether outline on minute pointer is drawn."
   DrawMinuteOutline = m_DrawMinuteOutline
End Property

Public Property Let DrawMinuteOutline(ByVal New_DrawMinuteOutline As Boolean)
   m_DrawMinuteOutline = New_DrawMinuteOutline
   PropertyChanged "DrawMinuteOutline"
   ShowTime
End Property

Public Property Get DrawShadow() As Boolean
Attribute DrawShadow.VB_Description = "Returns/sets whether the shadow of the pointers is drawn."
   DrawShadow = m_DrawShadow
End Property

Public Property Let DrawShadow(ByVal New_DrawShadow As Boolean)
   m_DrawShadow = New_DrawShadow
   PropertyChanged "DrawShadow"
   ShowTime
End Property

Public Property Get DrawBodyOutline() As Boolean
Attribute DrawBodyOutline.VB_Description = "Returns/sets whether the clock's outline is drawn."
   DrawBodyOutline = m_DrawBodyOutline
End Property

Public Property Let DrawBodyOutline(ByVal New_DrawBodyOutline As Boolean)
   m_DrawBodyOutline = New_DrawBodyOutline
   PropertyChanged "DrawBodyOutline"
   ShowTime
End Property

Public Property Get DrawSecond() As Boolean
Attribute DrawSecond.VB_Description = "Returns/sets whether the second pointer is drawn."
   DrawSecond = m_DrawSecond
End Property

Public Property Let DrawSecond(ByVal New_DrawSecond As Boolean)
   m_DrawSecond = New_DrawSecond
   PropertyChanged "DrawSecond"
   ShowTime
End Property

Public Property Get AntiAliasing() As Boolean
Attribute AntiAliasing.VB_Description = "Returns/sets whether anti-aliasing is used to draw the clock. (Procedures by Robert Rayment on (RRPaint)"
   AntiAliasing = m_AntiAliasing
End Property

Public Property Let AntiAliasing(ByVal New_AntiAliasing As Boolean)
   m_AntiAliasing = New_AntiAliasing
   PropertyChanged "AntiAliasing"
   ShowTime
End Property

Public Sub About()
Attribute About.VB_Description = "Show about window of this OCX."
Attribute About.VB_UserMemId = -552
   Dim frmX As Form

   For Each frmX In Forms
      If frmX.Name = "frmAbout" Then Unload frmX
   Next frmX
   
   frmAbout.Show vbModeless, UserControl.Parent
End Sub

Private Sub Draw_GradientCircle(lngColor1 As Long, Optional lngColor2 As Long = &HFFFFFF)
   Dim SQNum As Double
   Dim tmpDir As Integer
   
   Dim eScale As ScaleModeConstants
   Dim eDraw As DrawModeConstants
   Dim lngDrawWidth As Long
   
   Dim lngX As Long
   Dim lngY As Long
   
   Dim tmpR1 As Long
   Dim tmpG1 As Long
   Dim tmpB1 As Long
   
   Dim tmpR2 As Long
   Dim tmpG2 As Long
   Dim tmpB2 As Long
   
   Dim FinalR As Long
   Dim FinalG As Long
   Dim FinalB As Long
   
   Dim FinalRGB As Single
   
   Dim lngR
   Dim lngG
   Dim lngB
   
   Dim lngCounter As Integer
   
   eScale = UserControl.ScaleMode
   UserControl.ScaleMode = vbPixels
   
   lngX = UserControl.ScaleWidth / 4
   lngY = UserControl.ScaleHeight / 4
   
   If lngX > (UserControl.ScaleWidth / 2) Then
      If lngY > (UserControl.ScaleHeight / 2) Then
         SQNum = (lngX * lngX) + (lngY * lngY)
         tmpDir = Sqr(SQNum)
      Else
         SQNum = (lngX * lngX) + ((UserControl.ScaleHeight - lngY) * (UserControl.ScaleHeight - lngY))
         tmpDir = Sqr(SQNum)
      End If
   Else
      If lngY > (UserControl.ScaleHeight / 2) Then
         SQNum = ((UserControl.ScaleWidth - lngX) * (UserControl.ScaleWidth - lngX)) + (lngY * lngY)
         tmpDir = Sqr(SQNum)
      Else
         SQNum = ((UserControl.ScaleWidth - lngX) * (UserControl.ScaleWidth - lngX)) + ((UserControl.ScaleHeight - lngY) * (UserControl.ScaleHeight - lngY))
         tmpDir = Sqr(SQNum)
      End If
   End If
   
   tmpR1 = Get_RGB(lngColor2, 1)
   tmpG1 = Get_RGB(lngColor2, 2)
   tmpB1 = Get_RGB(lngColor2, 3)
   tmpR2 = Get_RGB(lngColor1, 1)
   tmpG2 = Get_RGB(lngColor1, 2)
   tmpB2 = Get_RGB(lngColor1, 3)
   
   lngR = (tmpR2 - tmpR1) / tmpDir
   lngG = (tmpG2 - tmpG1) / tmpDir
   lngB = (tmpB2 - tmpB1) / tmpDir
   
   eDraw = UserControl.DrawMode
   lngDrawWidth = UserControl.DrawWidth
   
   UserControl.DrawWidth = 2
   UserControl.DrawMode = 13
   
   For lngCounter = tmpDir - 1 To 0 Step -1
      FinalR = tmpR1 + (lngR * lngCounter)
      FinalG = tmpG1 + (lngG * lngCounter)
      FinalB = tmpB1 + (lngB * lngCounter)
      
      FinalRGB = RGB(FinalR, FinalG, FinalB)
      
      UserControl.FillColor = RGB(FinalR, FinalG, FinalB)
      UserControl.Circle (lngX, lngY), lngCounter
   Next lngCounter
   
   UserControl.ScaleMode = eScale
   UserControl.DrawWidth = lngDrawWidth
   UserControl.DrawMode = eDraw
End Sub

Private Function Get_RGB(RGBValue As Long, val As Integer) As Long
   If RGBValue > -1 And val > 0 And val < 4 Then
      Select Case val
         Case 1
            Get_RGB = (RGBValue And &HFF&)
         Case 2
            Get_RGB = (RGBValue And &HFF00&) / &H100
         Case 3
            Get_RGB = (RGBValue And &HFF0000) / &H10000
      End Select
   End If
End Function

'========================================================
'Anti-Aliasing Procedures By Robert Rayment on (RRPaint)
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1
'========================================================

Private Sub LngToRGB(LCul As Long, r As Byte, g As Byte, B As Byte)
   r = LCul And &HFF&
   g = (LCul And &HFF00&) \ &H100&
   B = (LCul And &HFF0000) \ &H10000
End Sub

Private Sub AALINE(ByVal ix1 As Integer, ByVal iy1 As Integer, ByVal ix2 As Integer, ByVal iy2 As Integer, Cul As Long)
   Dim zm As Single, zc As Single
   Dim Xs As Single, Ys As Single
   Dim i As Integer
   
   If ix1 = ix2 Or iy1 = iy2 Then
      UserControl.Line (ix1, iy1)-(ix2, iy2), Cul
      Exit Sub
   End If
   
   If Abs(ix2 - ix1) < Abs(iy2 - iy1) Then
      If iy2 < iy1 Then
         i = ix1
         ix1 = ix2
         ix2 = i
         i = iy1
         iy1 = iy2
         iy2 = i
      End If
      
      zm = (iy2 - iy1) / (ix2 - ix1)
      zc = iy1 - zm * ix1
      
      For Ys = iy1 To iy2
         Xs = (Ys - zc) / zm
         CalcAA Xs, Ys, Cul
      Next Ys
   Else
      If ix2 < ix1 Then
         i = ix1
         ix1 = ix2
         ix2 = i
         i = iy1
         iy1 = iy2
         iy2 = i
      End If
      
      zm = (iy2 - iy1) / (ix2 - ix1)
      zc = iy1 - zm * ix1
      
      For Xs = ix1 To ix2
         Ys = zm * Xs + zc
         CalcAA Xs, Ys, Cul
      Next Xs
   End If
End Sub

'Based on:-
'http://www.eclipzer.com/tutorials/subpixel/subpixel.html
'Web Ref frm Aleksander Ruzicic at PSC CodeId=66836
Private Sub CalcAA(Xs As Single, Ys As Single, Cul As Long)
   Dim ix As Single, iy As Single
   Dim a1 As Single, a2 As Single, a3 As Single, a4 As Single
   Dim r1 As Byte, g1 As Byte, b1 As Byte
   Dim r2 As Byte, g2 As Byte, b2 As Byte
   Dim r3 As Byte, g3 As Byte, b3 As Byte
   Dim r4 As Byte, g4 As Byte, b4 As Byte
   Dim rc As Byte, gc As Byte, bc As Byte
   Dim cul1 As Long, cul2 As Long, cul3 As Long, cul4 As Long
   
   If Xs = Int(Xs) Then
      Xs = Xs + 0.07
   End If
   
   If Ys = Int(Ys) Then
      Ys = Ys + 0.07
   End If
   
   ix = Int(Xs)
   iy = Int(Ys)
   a1 = (ix + 1 - Xs) * (iy + 1 - Ys)
   a2 = (Xs - ix) * (iy + 1 - Ys)
   a3 = (ix + 1 - Xs) * (Ys - iy)
   a4 = (Xs - ix) * (Ys - iy)
   
   LngToRGB GetPixel(UserControl.hdc, ix, iy), r1, b1, g1
   LngToRGB GetPixel(UserControl.hdc, ix + 1, iy), r2, b2, g2
   LngToRGB GetPixel(UserControl.hdc, ix, iy + 1), r3, b3, g3
   LngToRGB GetPixel(UserControl.hdc, ix + 1, iy + 1), r4, b4, g4
   LngToRGB Cul, rc, gc, bc
   
   cul1 = RGB(a1 * (1& * rc - r1) + r1, a1 * (1& * gc - g1) + g1, a1 * (1& * bc - b1) + b1)
   cul2 = RGB(a2 * (1& * rc - r2) + r2, a2 * (1& * gc - g2) + g2, a2 * (1& * bc - b2) + b2)
   cul3 = RGB(a3 * (1& * rc - r3) + r3, a3 * (1& * gc - g3) + g3, a3 * (1& * bc - b3) + b3)
   cul4 = RGB(a4 * (1& * rc - r4) + r4, a4 * (1& * gc - g4) + g4, a4 * (1& * bc - b4) + b4)

   SetPixelV UserControl.hdc, ix, iy, cul1
   SetPixelV UserControl.hdc, ix + 1, iy, cul2
   SetPixelV UserControl.hdc, ix, iy + 1, cul3
   SetPixelV UserControl.hdc, ix + 1, iy + 1, cul4
End Sub

Private Sub AAELLIPSE(ByVal ix1 As Integer, ByVal iy1 As Integer, ByVal zradx As Single, ByVal zrady As Single, Cul As Long)
   Dim TAlpha As Double
   Dim zxc As Single, zyc As Single
   Dim zStep As Double
   
   If zradx = 0 Then zradx = 0.001
   If zrady = 0 Then zrady = 0.001
   
   zStep = 2 / zradx
   
   If zrady > zradx Then zStep = 2 / zrady
   
   For TAlpha = 0 To 2 * Pi Step zStep
      zxc = ix1 + zradx * Cos(TAlpha)
      zyc = iy1 + zrady * Sin(TAlpha)
      CalcAA zxc, zyc, Cul
   Next TAlpha
   
   For TAlpha = 0 To 2 * Pi Step zStep
      zxc = ix1 + zradx * Cos(TAlpha)
      zyc = iy1 + zrady * Sin(TAlpha)
      SetPixelV UserControl.hdc, zxc, zyc, Cul
   Next TAlpha
End Sub

'========================================================

