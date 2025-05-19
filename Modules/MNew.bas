Attribute VB_Name = "MNew"
Option Explicit

Public Function PictureBoxZoom(Canvas As PictureBox, aImage As StdPicture) As PictureBoxZoom
    Set PictureBoxZoom = New PictureBoxZoom: PictureBoxZoom.New_ Canvas, aImage
End Function

