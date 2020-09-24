<div align="center">

## Image Resizing


</div>

### Description

To automatically resize an image control in a frame control to view at an acceptable size. The full image is on screen even if the image is bigger than the screen.
 
### More Info
 
What an Image and Frame control is. How pixels and functions work.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[WalkerBro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/walkerbro.md)
**Level**          |Intermediate
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/walkerbro-image-resizing__1-5625/archive/master.zip)





### Source Code

```
Option Explicit
Public Sub PicShow(ByVal PixPath As String, fForm As Form)
On Error GoTo noshow
Dim dHeight, dIHeight
Dim dWidth, dIWidth
Dim dPercent
With fForm
  .ViewImage.Visible = False
  .ViewImage.Stretch = False
  .Caption = App.Title & " - " & UCase(PixPath)
  .ViewImage.Picture = LoadPicture(PixPath)
    If .ViewImage.Height < .PicBack.Height And .ViewImage.Width < .PicBack.Width Then
      .ViewImage.Visible = True
      Exit Sub
    End If
  dHeight = .ViewImage.Height
  dWidth = .ViewImage.Width
  dIHeight = .PicBack.Height - 1
  dIWidth = .PicBack.Width - 1
  .ViewImage.Stretch = True
  .ViewImage.Height = .PicBack.Height - 2
  dPercent = (.PicBack.Height - 2) / dHeight * 100
  .ViewImage.Width = dWidth / 100 * dPercent
    If .ViewImage.Width > (.PicBack.Width - 2) Then
      .ViewImage.Stretch = False
      dHeight = .ViewImage.Height
      dWidth = .ViewImage.Width
      dIHeight = .PicBack.Height - 1
      dIWidth = .PicBack.Width - 1
      .ViewImage.Stretch = True
      .ViewImage.Width = .PicBack.Width - 1
      dPercent = (.PicBack.Width - 1) / dWidth * 100
      .ViewImage.Height = dHeight / 100 * dPercent
    End If
  .ViewImage.Visible = True
  MidPic frmMain2000
End With
Exit Sub
noshow:
Resume noshow1
noshow1:
End Sub
Public Sub MidPic(ByVal fForm As Form)
  fForm.ViewImage.Move (fForm.PicBack.Width - fForm.ViewImage.Width) / 2, (fForm.ViewImage.Height - fForm.ViewImage.Height) / 2
End Sub
'How to call the function
Call PicShow("c:\image.jpg", frmName)
```

