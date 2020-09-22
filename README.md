<div align="center">

## Moving StarField Form BG \(3D Version\)


</div>

### Description

This is a really cool form effect, it looks like a starfield moving from the middle of the screen out at you, it kind of looks like the Windows screen saver. This is not originally my code, but i like it so much that i have made a better version of it. You can change the speed and color of the field. This is an update of my last submission (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=40900&lngWId=1). Enjoy
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Lane](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-lane.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-lane-moving-starfield-form-bg-3d-version__1-40953/archive/master.zip)





### Source Code

<p><font face="Verdana"><b><u><small><small>Moving StarField Form BG (3D
Version)<br>
</small></small></u><small><small>by Adam Lane</small></small><br>
<br>
</b><small><small>1) Create a form and a timer<br>
2) Form1 and Timer1<br>
3) Copy this code into your form</small></small></font><small><small><font face="Courier New"><small><small><br>
<br>
</small></small></font></small></small><font face="Courier New"><small><small><small><small>Dim X(50), Y(50), xSpeed(50), ySpeed(50) As Integer<br>
<br>
Private Sub Form_Load()<br>
&nbsp;&nbsp;&nbsp; Dim i As Integer<br>
&nbsp;&nbsp;&nbsp; Timer1.Interval = 4<br>
&nbsp;&nbsp;&nbsp; Form1.BackColor = vbBlack<br>
&nbsp;&nbsp;&nbsp; Form1.ForeColor = vbBlack<br>
&nbsp;&nbsp;&nbsp; Form1.FillColor = vbBlack<br>
&nbsp;&nbsp;&nbsp; For i = 0 To 49<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; X(i) = -1<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Y(i) = -1<br>
&nbsp;&nbsp;&nbsp; Next i<br>
&nbsp;&nbsp;&nbsp; Randomize<br>
&nbsp;&nbsp;&nbsp; Timer1.Enabled = True<br>
End Sub<br>
<br>
Private Sub Timer1_Timer()<br>
&nbsp;&nbsp;&nbsp; Dim i As Integer<br>
&nbsp;&nbsp;&nbsp; For i = 0 To 49<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; PSet (X(i), Y(i)), &amp;H0&amp;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; If X(i) &lt; 0 Or X(i) > Form1.ScaleWidth Or Y(i) &lt; 0 Or Y(i) > Form1.ScaleHeight Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; X(i) = Form1.ScaleWidth \ 2<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Y(i) = Form1.ScaleHeight \ 2<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; xSpeed(i) = Int(Rnd(1) * 200) - 100<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ySpeed(i) = Int(Rnd(1) * 200) - 100<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End If<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; X(i) = X(i) + xSpeed(i)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Y(i) = Y(i) + ySpeed(i)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; PSet (X(i), Y(i)), &amp;HFFFFFF<br>
&nbsp;&nbsp;&nbsp; Next i<br>
End Sub</small></small></small></small></font></p>

