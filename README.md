<div align="center">

## Circle Drawing Example


</div>

### Description

This code is simply to illustrate the mathematical operations necessary to draw a circle. Of course I realize VB has built in functions for this, but thought this might be useful for some to understand. In my example, I use a picture box with the PSet method to do the drawing.
 
### More Info
 
Simply create a form with a decent size picture box (picture1), a text box we will enter the radius in (text1), a command button to draw the circle (command1), and a command button for exit (command2). Paste the code I have supplied and run it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[G\-Man](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/g-man.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/g-man-circle-drawing-example__1-13613/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  Dim intRadius, intDegree As Integer
  Dim sngX, sngY As Single
  If IsNumeric(Text1.Text) Then
    intRadius = Text1.Text
    For intDegree = 1 To 360
      sngX = (Cos(intDegree) * intRadius) + intRadius
      sngY = intRadius - (Sin(intDegree) * intRadius)
      Picture1.PSet (sngX, sngY), vbBlack
    Next
  Else
    MsgBox "Please enter a numeric value for the radius."
    Text1.SetFocus
  End If
End Sub
Private Sub Command2_Click()
  Unload Form1
End Sub
```

