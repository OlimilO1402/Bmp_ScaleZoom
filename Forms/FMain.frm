VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbZoom 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   360
      Width           =   8895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Zoom:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   555
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewZoomNormal 
         Caption         =   "Normal 1:1"
      End
      Begin VB.Menu mnuViewZoomIn_ 
         Caption         =   "Zoom In"
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "2:1"
            Index           =   2
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "3:1"
            Index           =   3
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "4:1"
            Index           =   4
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "5:1"
            Index           =   5
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "6:1"
            Index           =   6
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "7:1"
            Index           =   7
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "8:1"
            Index           =   8
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "9:1"
            Index           =   9
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "10:1"
            Index           =   10
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "11:1"
            Index           =   11
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "12:1"
            Index           =   12
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "13:1"
            Index           =   13
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "14:1"
            Index           =   14
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "15:1"
            Index           =   15
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "16:1"
            Index           =   16
         End
      End
      Begin VB.Menu mnuViewZoomOut_ 
         Caption         =   "Zoom Out"
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:2"
            Index           =   2
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:3"
            Index           =   3
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:4"
            Index           =   4
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:5"
            Index           =   5
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:6"
            Index           =   6
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:7"
            Index           =   7
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:8"
            Index           =   8
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:9"
            Index           =   9
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:10"
            Index           =   10
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:11"
            Index           =   11
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:12"
            Index           =   12
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:13"
            Index           =   13
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:14"
            Index           =   14
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:15"
            Index           =   15
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:16"
            Index           =   16
         End
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Image    As IPictureDisp
Private m_PBZoom As PictureBoxZoom

Private Sub Form_Load()
    Set m_Image = LoadPicture(App.Path & "\Bild4.bmp")
    'Debug.Print m_Image.Width & " x " & m_Image.Height
    InitZoomCombo 15
    Set m_PBZoom = MNew.PictureBoxZoom(Me.Picture1, m_Image)
End Sub

Private Function HiMetric_ToPixel(ByVal Value_inHiM As Long) As Long
    HiMetric_ToPixel = Value_inHiM / 2.54 * 96 / 1000
End Function

Private Function Pixel_ToHiMetric(ByVal Value_pix As Long) As Long
    Pixel_ToHiMetric = Value_pix * 2.54 / 96 * 1000
End Function

Private Sub Form_Resize()
    Dim L As Single: L = Picture1.Left
    Dim T As Single: T = Picture1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight
    If W > 0 And H > 0 Then Picture1.Move L, T, W, H
End Sub

Sub InitZoomCombo(ByVal ListIndex As Integer)
    Dim i As Long
    For i = 16 To 1 Step -1: CmbZoom.AddItem CStr(i) & ":1": Next
    For i = 2 To 16:         CmbZoom.AddItem "1:" & CStr(i): Next
    CmbZoom.ListIndex = ListIndex
End Sub

Private Sub mnuEditPaste_Click()
    If Not Clipboard.GetFormat(ClipBoardConstants.vbCFBitmap) Then Exit Sub
    Set m_Image = Clipboard.GetData(ClipBoardConstants.vbCFBitmap)
    Debug.Print m_Image.Width & " x " & m_Image.Height
    'If m_PBZoom Is Nothing Then
        Set m_PBZoom = MNew.PictureBoxZoom(Me.Picture1, m_Image)
    'Else
    '    Set m_PBZoom.Image = m_Image
    'End If
End Sub

Private Sub mnuViewZoomNormal_Click()
    CmbZoom.ListIndex = 15
End Sub

Sub UnCheckAllMnuViewZoom()
    Dim i As Integer
    For i = mnuViewZoomIn.LBound To mnuViewZoomIn.UBound:   mnuViewZoomIn(i).Checked = False:  Next
    For i = mnuViewZoomOut.LBound To mnuViewZoomOut.UBound: mnuViewZoomOut(i).Checked = False: Next
End Sub

Private Sub mnuViewZoomIn_Click(Index As Integer)
    UnCheckAllMnuViewZoom
    CmbZoom.ListIndex = 16 - Index
    mnuViewZoomIn(Index).Checked = True
End Sub

Private Sub mnuViewZoomOut_Click(Index As Integer)
    UnCheckAllMnuViewZoom
    CmbZoom.ListIndex = 14 + Index
    mnuViewZoomOut(Index).Checked = True
End Sub

Private Sub CmbZoom_Click()
    If m_PBZoom Is Nothing Then Exit Sub
    Dim li As Long:     li = CmbZoom.ListIndex
    If li < 0 Then Exit Sub
    Dim s As String:     s = CmbZoom.List(li)
    Dim sa() As String: sa = Split(s, ":")
    m_PBZoom.ZoomFactor = CDbl(sa(0)) / CDbl(sa(1))
End Sub

Sub UpdateView()
    m_PBZoom.Refresh
End Sub
