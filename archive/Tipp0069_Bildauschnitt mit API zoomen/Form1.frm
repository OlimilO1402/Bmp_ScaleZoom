VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Ende"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   1440
   End
   Begin VB.PictureBox Zoom 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   1440
      ScaleHeight     =   2535
      ScaleWidth      =   2175
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox Source 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1320
      ScaleWidth      =   1125
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

Private Declare Function StretchBlt Lib "gdi32" _
        (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hSrcDC As Long, ByVal xSrc As Long, _
        ByVal ySrc As Long, ByVal nSrcWidth As Long, _
        ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
        
Private Declare Function SetStretchBltMode Lib "gdi32" _
        (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
    
Const SRCCOPY = &HCC0020
Const DELETESCANS = 3

Dim Wert%, Schritt%

Private Sub Form_Load()
  Wert = 90
  Schritt = 5
End Sub

Private Sub Command1_Click()
  Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Call ZoomSub(Wert, Zoom)
  
  If Wert >= 200 Then Schritt = -Schritt
  If Wert <= 10 Then Schritt = -Schritt
  
  If Wert < 50 Then Schritt = Sgn(Schritt) * 2
  If Wert >= 50 Then Schritt = Sgn(Schritt) * 5
  
  Wert = Wert + Schritt
End Sub

Private Sub ZoomSub(ByVal Prozent, Bild As PictureBox)
  Dim dx1&, dy1&, dx2&, dy2&, tx&, ty&, SHdc&, ZHdc&

    tx = Screen.TwipsPerPixelX
    ty = Screen.TwipsPerPixelY
    
    dx1 = Source.Width / tx
    dy1 = Source.Height / ty
    
    dx2 = dx1 * Prozent / 100
    dy2 = dy1 * Prozent / 100
    
    ZHdc = Zoom.hdc
    SHdc = Source.hdc
    Zoom.AutoRedraw = False
    
    Call SetStretchBltMode(ZHdc, DELETESCANS)
    StretchBlt ZHdc, 0, 0, dx2, dy2, _
               SHdc, 0, 0, dx1, dy1, SRCCOPY
    Zoom.Refresh
    Zoom.AutoRedraw = True
End Sub
