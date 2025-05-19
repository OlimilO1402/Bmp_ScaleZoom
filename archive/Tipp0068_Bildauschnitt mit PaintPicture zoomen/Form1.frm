VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ende"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox Zoom 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
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
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
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
      TabIndex        =   0
      Top             =   120
      Width           =   1125
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
  Call ZoomSub(Wert)
  
  If Wert >= 200 Then Schritt = -Schritt
  If Wert <= 10 Then Schritt = -Schritt
  
  If Wert < 50 Then Schritt = Sgn(Schritt) * 2
  If Wert >= 50 Then Schritt = Sgn(Schritt) * 5
  
  Wert = Wert + Schritt
End Sub

Private Sub ZoomSub(ByVal Prozent)
  Dim dx1&, dy1&, dx2&, dy2&
    
    dx1 = Source.Width
    dy1 = Source.Height
    
    dx2 = dx1 * Prozent / 100
    dy2 = dy1 * Prozent / 100
    
    Zoom.PaintPicture Source, 0, 0, dx2, dy2
End Sub
