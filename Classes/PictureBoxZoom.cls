VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictureBoxZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Canvas     As PictureBox
Private m_Image      As IPictureDisp
Private m_Window     As Form
Private m_OrigWidth  As Double
Private m_OrigHeight As Double
Private m_ZoomFactor As Double

Private Sub Class_Initialize()
    m_ZoomFactor = 1
End Sub

Friend Sub New_(Window As Form, Canvas As PictureBox, Image As IPictureDisp)
    Set m_Window = Window: Set m_Canvas = Canvas: Set m_Image = Image
    m_Canvas.AutoSize = True
    Set m_Canvas.Picture = m_Image
    m_OrigWidth = m_Canvas.ScaleWidth
    m_OrigHeight = m_Canvas.ScaleHeight
End Sub
'Copy to MNew
'Public Function PictureBoxZoom(Window As Form, Canvas As PictureBox, aImage As StdPicture) As PictureBoxZoom
'    Set PictureBoxZoom = New PictureBoxZoom: PictureBoxZoom.New_ Window, Canvas, aImage
'End Function
'how to use it: in a form do the following:
'Private m_PBZoom As PictureBoxZoom
'Private Sub Form_Load()
'    Dim pic As StdPicture: Set pic = LoadPicture(App.Path & "\Bild1.bmp")
'    Set m_PBZoom = MNew.PictureBoxZoom(Me, Me.Picture2, pic)
'End Sub

Public Property Get ZoomFactor() As Double
    ZoomFactor = m_ZoomFactor
End Property
Public Property Let ZoomFactor(ByVal Value As Double)
    m_ZoomFactor = Value
    UpdateView
End Property

Public Function PropZoomToStr() As String
    PropZoomToStr = IIf(m_ZoomFactor <= 1, "1:" & 1 / m_ZoomFactor, m_ZoomFactor & ":1")
    Debug.Print CmbZoomListIndex
End Function

Public Function CmbZoomListIndex() As Integer
    CmbZoomListIndex = IIf(m_ZoomFactor <= 1, 1 / m_ZoomFactor + 15, m_ZoomFactor - 16)
End Function

Public Property Get Image() As IPictureDisp
    Set Image = m_Image
End Property
Public Property Set Image(ByVal Value As IPictureDisp)
    Set m_Image = Value
    m_Canvas.Cls
    m_Canvas.AutoSize = True
    Set m_Canvas.Picture = m_Image
    m_OrigWidth = m_Canvas.ScaleWidth
    m_OrigHeight = m_Canvas.ScaleHeight
    UpdateView
End Property

Private Sub UpdateView()
    Dim L As Single: L = m_Canvas.Left
    Dim T As Single: T = m_Canvas.Top
    Dim W As Single: W = m_OrigWidth * m_ZoomFactor
    Dim H As Single: H = m_OrigHeight * m_ZoomFactor
    If 0 < W And 0 < H Then
        W = Min(m_Window.ScaleWidth, W)
        H = Min(m_Window.ScaleHeight, H)
        m_Canvas.Move L, T, W, H
        'If W < m_Window.ScaleWidth And H < m_Window.ScaleHeight Then m_Canvas.Move L, T, W, H
    End If
    If m_Image Is Nothing Then Exit Sub
    W = m_Image.Width: H = m_Image.Height
    m_Canvas.Cls
    m_Canvas.PaintPicture m_Image, 0, 0, W * m_ZoomFactor, H * m_ZoomFactor, 0, 0, W, H
End Sub
Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Public Sub Refresh()
    UpdateView
End Sub
