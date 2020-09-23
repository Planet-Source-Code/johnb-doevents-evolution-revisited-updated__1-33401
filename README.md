<div align="center">

## DoEvents Evolution Revisited \(Updated\)


</div>

### Description

How to use John Galanopoulos' article in a class module to override the standard DoEvents method.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-04-04 09:56:02
**By**             |[JohnB](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/johnb.md)
**Level**          |Intermediate
**User Rating**    |5.0 (60 globes from 12 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[DoEvents\_E68736442002\.zip](https://github.com/Planet-Source-Code/johnb-doevents-evolution-revisited-updated__1-33401/archive/master.zip)





### Source Code

I added/changed some code to further optimize my code which basically is an
optimization of the DoEvents method. Thanks goes to <a href="http://www.planet-source-code.com/vb/feedback/EmailUser.asp?lngWId=1&lngToPersonId=272887&txtReferralPage=http%3A%2F%2Fwww%2Eplanet%2Dsource%2Dcode%2Ecom%2Fvb%2Fscripts%2FShowCode%2Easp%3FlngWId%3D1%26txtCodeId%3D33401">Marzo
Sette Torres Junior</a> for his help!
<p>
It is best to view the article located here: <a href="http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=29735&lngWId=1">http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=29735&lngWId=1</a>
</p>
<p>
Next download the code attached, which is just a class module. I included the Word DOC file from John G. with the class module.<br>
<br>
Simply add the module to your project and declare an object of the clsDoEvents type:<br>
  <font color="#0000FF">Dim oDoEvents as clsDoEvents</font>
<br>   <font color="#0000FF">Set oDoEvents = New clsDoEvents</font></p>
<p>
Then set its only property to any of the enumerated values. <br>
  <font color="#0000FF">oDoEvents.QueueUsed = Standard</font><br>
<br>
Valid enumerated values for this property<br>
<font color="#0000FF"> All_Inputs = QS_ALLINPUT<br>
 All_Events = QS_ALLEVENTS<br>
 Standard = QS_STANDARD<br>
 Messages = QS_MESSAGES<br>
 InputOnly = QS_INPUT<br>
 Mouse = QS_MOUSE<br>
 MouseMove = QS_MOUSEMOVE<br>
 Timer = QS_TIMER<br>
<br>
</font>Constants relating to the above enumerated values:<br>
<font color="#0000FF"> </font>  <font color="#0000FF">QS_HOTKEY = &H80<br>
  QS_KEY = &H1<br>
  QS_MOUSEBUTTON = &H4<br>
  QS_MOUSEMOVE = &H2<br>
  QS_PAINT = &H20<br>
  QS_POSTMESSAGE = &H8<br>
  QS_SENDMESSAGE = &H40<br>
  QS_TIMER = &H10<br>
  QS_MOUSE = (QS_MOUSEMOVE Or
QS_MOUSEBUTTON)<br>
  QS_INPUT = (QS_MOUSE Or QS_KEY)<br>
  QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or _<br>
  QS_HOTKEY)<br>
  QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or _<br>
  QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)<br>
</font><font color="#0000FF">QS_MESSAGES = (QS_POSTMESSAGE Or QS_SENDMESSAGE) </font><font color="#008000">      ' Not MS standard constant</font><font color="#0000FF"><br>
 QS_STANDARD = (QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT) </font><font color="#008000">  ' Not MS standard constant</font><font color="#0000FF"><br>
</font><br>
Now, in any long winded loop, simple call the only method: <br>
  <font color="#0000FF">oDoEvents.GetInputState</font></p>
<p>
API function that will determine if we need to DoEvents:<br>
<font color="#0000FF">Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long</font></p>
<p>Don't forget to destroy your object when you are through with it!<br>
  <font color="#0000FF">Set oDoevents =
Nothing</font></p>
<p>(I chose this method name to "honor" John for his article.)</p>
<p> Thanks and good luck!</p>

