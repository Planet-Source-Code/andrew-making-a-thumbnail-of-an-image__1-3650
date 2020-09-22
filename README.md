<div align="center">

## Making a thumbnail of an image


</div>

### Description

Being optimistic, I have made a code that will resize an image, allowing the user to save the image as a thumbnail. Here's hoping I've solved problems.

Even if it doesn't work for you, it should give you a small idea of how it's done.
 
### More Info
 
All provided

Source and dest are two picturebox objects.

I haven't had much time to test it, but BE WARNED; it might not work if the source image is smaller than your thumbnail size. I tested this at 60 x 80.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andrew](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew.md)
**Level**          |Unknown
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrew-making-a-thumbnail-of-an-image__1-3650/archive/master.zip)





### Source Code

```
Sub thumbnail(width As Integer, height As Integer, source As PictureBox, dest As PictureBox)
 'This should help me to create a thumbnail of an image.
 'ix and iy help to grab the pixels from the relative positions
 'of the thumbnail from the image.
 Dim ix As Single, iy As Single
 'x and y are just For...Next variables and xcounter/ycounter
 'are used for reference to the thumbnail.
 Dim x As Single, y As Single, xcounter As Integer, ycounter As Integer
 'These are a few safety precautions that you should take to
 'make sure that the code works. The ScaleMode of the
 'pictureboxes and their parents must be pixels.
 source.Parent.ScaleMode = vbPixels
 dest.Parent.ScaleMode = vbPixels
 source.ScaleMode = vbPixels
 dest.ScaleMode = vbPixels
 'Calculate ix and iy, which are the 'steps' from which to grab
 'pixels. Think of it as a fixed grid.
 ix = source.ScaleWidth / width
 iy = source.ScaleHeight / height
 'Resize the thumbnail picturebox to accomodate the new
 'thumbnail. There's a trap here; the thumbnail may not be
 'exactly the size required.
 'If you simply put dest.height = height and so on for the
 'width, you might get the extra border on the right and
 'bottom of the thumbnail.
 dest.height = source.ScaleHeight / iy
 dest.width = source.ScaleWidth / ix
 'Now we make the thumbnail.
 For y = 0 To source.ScaleHeight - 1 Step iy
 For x = 0 To source.ScaleWidth - 1 Step ix
  'Grab the image from the source and place it in the
  'right spot in the thumbnail picture box.
  dest.PSet (xcounter, ycounter), source.Point(x, y)
  xcounter = xcounter + 1
 Next
 ycounter = ycounter + 1
 xcounter = 0
 Next
 'The next line is not mandatory, except if you want the
 'thumbnail to become a picture object.
 Set dest.Picture = dest.Image
End Sub
'To save the thumbnail you would then write a line such as
'SavePicture dest.picture, "thumbnail.bmp" (or
'SavePicture dest.image), remembering that the result is a
'bitmap picture.
```

