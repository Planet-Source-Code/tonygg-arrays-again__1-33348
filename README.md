<div align="center">

## Arrays Again


</div>

### Description

Clearing out the digital cupboards, here is a tutorial on Arrays.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[TonyGG](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tonygg.md)
**Level**          |Beginner
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tonygg-arrays-again__1-33348/archive/master.zip)





### Source Code

```
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Visual Basic Tutorial - Arrays</TITLE>
</head>
<BODY bgColor=#ffffff>
<P align=center><STRONG><FONT face=Verdana color=#0080ff><BIG><BIG><BIG>Visual
Basic Tutorial:</BIG></BIG></BIG></FONT><FONT face=Verdana
color=#0000a0><BR></FONT><FONT face=Arial color=#000000><BIG><BIG>Form and App
Objects</BIG></BIG></FONT><FONT face="Comic Sans MS"
color=#008000><BR></FONT></STRONG><FONT face=Verdana color=#000000>By: <A
href="mailto:metadrummer@hotmail.com">Robert Klieger</A></FONT></P>
<HR>
<P><FONT face=Arial color=#000000 size=4><STRONG>Table of
Contents</STRONG></FONT>
<UL>
 <LI><FONT
 face=Verdana color=#0000a0 size=3>Declaring Arrays</FONT>
 <LI><FONT
 face=Verdana color=#000080 size=3>Accessing Elements of an Array for
 Input/Output</FONT>
 <LI><FONT
 face=Verdana color=#000080 size=3>Getting the Size of an Array</FONT>
 <LI><FONT
 face=Verdana color=#000080 size=3>Multidimensional Arrays</FONT>
 <LI><FONT
 face=Verdana color=#000080 size=3>Dynamic Arrays</FONT></LI></UL>
<HR>
<H1><FONT face=Arial color=#000000 size=4>Declaring Arrays</FONT></H1>
<P><FONT face=Verdana size=3>When declaring an Array in VB, you are telling it
three things. First, the name of the array. This is what you will use to access
the array later in your program. Second, how many elements it can store, this is
inside the parenthesis. And last, the data type, Integer, String, etc... If you
declare the array in the general declarations section of a form, the array is
visible to all procedures and functions in that form. If you declare the array
as Public in a Module it will be visible to every procedure and function in your
project. To declare an array use the following syntax:</FONT></P>
<P><FONT face="Courier New"><FONT color=#000080 size=3>Dim</FONT><FONT
size=3> ArrayName(LowerValue To HigherValue) </FONT><FONT color=#000080
size=3>As DataType</FONT></FONT></P>
<P><FONT face=Verdana size=3>If you declare an array in a Form, use Dim or
Private, if you declare it in a Module and you want every procedure to access
it, declare it as Public. If you declare the array in a procedure use Dim.
ArrayName is the name of the array. LowerValue is the first element in the
array. HigherValue is the last element in the array. DataType is a valid
datatype (ie String)/ You don't have to specify the LowerValue, but if you
don't, the array will start from 0. For example,</FONT></P>
<P><FONT face="Courier New"><FONT color=#000080>Dim</FONT> sTestArray(0 To
10) <FONT color=#000080>As String</FONT></FONT></P>
<P><FONT face=Verdana size=3>is the same as...</FONT></P>
<P><FONT face="Courier New"><FONT color=#000080>Dim</FONT> sTestArray(10)
<FONT color=#000080>As String</FONT></FONT></P>
<P><FONT face=Verdana size=3>The code below declares nArray, with 1 as the first
element, 30 as the last element, all of which stores Integers.</FONT></P>
<P><FONT face="Courier New"><FONT color=#000080>Dim</FONT> nArray(1 To 30)
<FONT color=#000080>As Integer</FONT></FONT></P><HR>
<H1><FONT color=#000000><FONT face=Arial><A name=ReadWrite></A></FONT><FONT
face=Arial size=4>Reading/Writing to Arrays</FONT></FONT></H1>
<P><FONT face=Verdana size=3>Assigning a value to an element of an array is as
easy as filling the value of a normal integer or string. Here's an example of
the syntax:</FONT></P>
<P><FONT face=Verdana>ArrayName(Element) = Value</FONT></P>
<P><FONT face=Verdana size=3>Reading from an array is just as easy. You can use
it in several different ways. Here are a few examples:</FONT></P>
<P><FONT face="Courier New">iArray(3) = iArray(7) <FONT color=#008000>'Takes
whatever is in element (3) and copys it into (7)</FONT><BR>Msgbox iArray(5)
<FONT color=#008000>'Takes whatever is in element (5) and puts it in a message
box</FONT></FONT></P>
<P><FONT face=Verdana size=3>So what data types can you use with an array? With
Visual Basic, you can use any data type you could use in a regular variable.
This means Boolean values will work to. With many previous programing vanguages,
this is not possible. Just one more reason I perfer using VB for many
tasks.</FONT></P>
<P><FONT face=Verdana size=3>One great thing about arrays in any programing
language, and this was the whole idea behind them, is that you only need to
specify one name for many values. Each value then has an index number assigned
to it, or more commonly know, an element number. Say you wanted to collect three
names. You could either define three variables:</FONT></P>
<P><FONT face="Courier New"><FONT color=#000080>Dim</FONT> Name1 <FONT
color=#000080>As String</FONT><BR><FONT color=#000080>Dim</FONT> Name2 <FONT
color=#000080>As String<BR>Dim</FONT> Name3 <FONT color=#000080>As
String</FONT></FONT></P>
<P><FONT face=Verdana size=3>The problem with this is, you would have to write
<STRONG>three</STRONG> diffrent codes to access them. If you wanted to do same
task for each name, you would probley have to use a <EM>for</EM> statement and a
<EM>select case </EM>statement<EM>. </EM>Imagine if you had 500 names! Writing
that much code would be a pain in the butt. Plus you would have to declare that
many strings. You are limited to what you declare. For example:</FONT></P>
<P><FONT face="Courier New"><CODE>Msgbox (Name & iNum)</CODE></FONT></P>
<P><FONT face=Verdana size=3>If you were atempting to access for Name1, and iNum
= 1, using Name & iNum would not work. This would simply show whatever was
in Name followed by whatever iNum equaled. With using an Array, you can simply
do this:</FONT></P>
<P><FONT face="Courier New">Msgbox Names(iNum)</FONT></P>
<P><CODE><FONT face=Verdana>This would access the element of <EM>Names</EM>
equaled to <EM>iNum</EM>.</FONT></CODE></P>
<HR>
<H1><FONT color=#000000><FONT face=Arial><A name=GetSize></A></FONT><FONT
face=Arial size=4>Getting the Size of an Array</FONT></FONT></H1>
<P><FONT face=Verdana size=3>By refering to size, we mean the upper and lower
limits of an array in terms of elements. Anotherwords, how many variables are in
an array. This is simply done using the LBound and UBound functions (upper &
lower boundrys). Why would you want to know this? If you wanted to fill a list
box with all the names in an array but are not sure how many names their are.
You would use these functions. This is especially useful if you have used the
ReDim statment (see <A
href="http://www.geocities.com/alpha_productions2/vb_tutorial5.htm#Dynaming">dynamic
arrays</A>). Here is the syntax for finding the upper boundries.</FONT></P>
<P><FONT face="Courier New"><FONT color=#000080>UBound(</FONT>ArrayName,
Dimension<FONT color=#000080>)</FONT> <FONT color=#008000>'ArrayName = the name
of the array</FONT></FONT></P>
<P><FONT face=Verdana size=3>Dimension is an optional integer representing the
dimension number for use when using <A
href="http://www.geocities.com/alpha_productions2/vb_tutorial5.htm#Mulitidimen">multidimensional
arrays</A> found towards the end of this tutorial. To find the lower boundries
of an array use the LBound function syntax:</FONT></P>
<P><FONT face="Courier New"><FONT color=#000080>LBound(</FONT>ArrayName,
Dimension<FONT color=#000080>)</FONT></FONT></P>
<P><FONT face=Verdana size=3>Pretty simple, huh? Heres an example that fills a
list box with the values of all the items in Names.</FONT></P>
<P><FONT face="Courier New">LowerVal = <FONT
color=#000080>LBound(</FONT>Names<FONT color=#000080>)</FONT> <FONT
color=#008000>'Get the lower boundry number.</FONT><BR>UpperVal = <FONT
color=#000080>UBound(</FONT>Names<FONT color=#000080>)</FONT><FONT
color=#008000> 'Get the upper boundry number.</FONT><BR><BR><FONT
color=#000080>For</FONT> i = LowerVal <FONT color=#000080>To</FONT>
UpperVal<BR>    List1.AddItem Names(i) <FONT color=#008000>'Add
each name from array according to how many stored in the Array</FONT><BR><FONT
color=#000080>Next</FONT></FONT></P><HR>
<H1><FONT color=#000000><FONT face=Arial><A name=Mulitidimen></A></FONT><FONT
face=Arial size=4>Multidimensional Arrays</FONT></FONT></H1>
<P><FONT face=Verdana size=3>Now that you know what an array is, figuring out
what a <EM>multi</EM>dimensional array is will be a little easier. A
multidimensional array is an array that looks like a table. If you have ever
used Microsoft Excel, you would know what I'm talking about. It can also be
refered to an array of arrays. That is many arrays using the same name. Here is
an example of a two-dimensional array.</FONT></P>
<P><FONT face="Courier New"><CODE>Static iArray(1 To 2, 1 To
3)</CODE></FONT></P>
<P><FONT face=Verdana>All the elements list out are:<BR>iArray(1,1),
iArray(1,2), iArray(1,3), iArray(2,1), iArray(2,2), iArray(2,3)<BR>If
it were in a table it would look like this:</FONT></P>
<DIV align=left>
<TABLE border=1>
 <TBODY>
 <TR>
 <TD align=middle width="25%"> </TD>
 <TD align=middle width="25%">
  <P align=center><SMALL><FONT face=Verdana>1</FONT></SMALL></P></TD>
 <TD align=middle width="25%">
  <P align=center><SMALL><FONT face=Verdana>2</FONT></SMALL></P></TD>
 <TD align=middle width="25%">
  <P align=center><SMALL><FONT face=Verdana>3</FONT></SMALL></P></TD></TR>
 <TR>
 <TD align=middle width="25%">
  <P align=center><SMALL><FONT face=Verdana>1</FONT></SMALL></P></TD>
 <TD align=middle width="25%"><SMALL><FONT
  face=Verdana>iArray(1,1)</FONT></SMALL></TD>
 <TD align=middle width="25%"><SMALL><FONT
  face=Verdana>iArray(1,2)</FONT></SMALL></TD>
 <TD align=middle width="25%"><SMALL><FONT
  face=Verdana>iArray(1,3)</FONT></SMALL></TD></TR>
 <TR>
 <TD align=middle width="25%"><SMALL><FONT face=Verdana>2</FONT></SMALL></TD>
 <TD align=middle width="25%"><SMALL><FONT
  face=Verdana>iArray(2,1)</FONT></SMALL></TD>
 <TD align=middle width="25%"><SMALL><FONT
  face=Verdana>iArray(2,2)</FONT></SMALL></TD>
 <TD align=middle width="25%"><SMALL><FONT
  face=Verdana>iArray(2,3)</FONT></SMALL></TD></TR></TBODY></TABLE></DIV>
<P><FONT face=Verdana size=3>Most of the time you will be using one dimensional
arrays but some of the time it will be helpful to use 2 or even 3. More then 3
is not always a good idea because it tends to be hard to debug. The usefulness
of VB will allow you to use as many as you want, each seperated by commas. Keep
in mind the more you use, the more computer memory you use up. This can get real
dangerous real fast. If you have some free time someday go ahead and make and
14-15 dimensional array. Make sure you close out and save everything else before
hand. See how well your computer runs. Be prepared to have to
restart.</FONT></P>
<HR>
<H1></FONT><FONT face=Arial><A name=Dynaming></A></FONT><FONT face=Arial
color=#000000 size=4>Dynamic Arrays</FONT><FONT face=Verdana></H1>
<P></FONT><FONT face=Verdana size=3>Arrays because of their uniqeness can consum
alot of memory if your not carefull. For example:</FONT><FONT face=Verdana></P>
<P><FONT face="Courier New"><FONT color=#000080>Dim</FONT> MyArray (10000)
</FONT><FONT face="Courier New" color=#000080>As</FONT><FONT face="Courier New">
<FONT color=#000080>Long</FONT></FONT></P>
<P><FONT face=Arial>Will require 40,004 bytes of memory. That is 10 x 10,001
because long variables take up 4 bytes of memory. This might not be alot now,
but if you start using 10 of these array, they consume 4,000,400 bytes. This is
why most of the time it is wise you set large arrays to minimum at the start,
and later resize them during runtime. Yes, you can do this! This can be done
using the ReDim function. Arrays that change size during runtime are called
dynamic arrays. Its that simple.</FONT></P>
<P></FONT><FONT face=Verdana size=3>When you declare a dynamic array, you don't
need to declare it like a fixed array. When you declare a dynamic array the size
is not specified. Instead you use the following syntax:</FONT><FONT
face=Verdana></P>
<P><FONT face="Courier New" color=#000080>Dim</FONT><FONT
face="Courier New"> ArrayName() </FONT><FONT face="Courier New"
color=#000080>As</FONT><FONT face="Courier New"> <FONT
color=#000080>DataType</FONT></FONT></P>
<P></FONT><FONT face=Verdana size=3>Every thing here is the same as before, Dim
the array (can set to Global if using in a module, or public if in procedure),
ArrayName is the name, and DataType is as said before. Heres how to use the
ReDim function:</FONT><FONT face=Verdana></P>
<P><FONT face="Courier New"><FONT
color=#000080>ReDim</FONT> ArrayName(LowerValue <FONT
color=#000080>To</FONT> HigherValue)</FONT></P>
<P></FONT><FONT face=Verdana size=3>You can use this either in a procedure or
function. Its almost like when you declare it except you use ReDim, don't use a
datatype, and are not declareing it. That is, it must already be
declared.</FONT><FONT face=Verdana><FONT size=3> <FONT face=Verdana>Here is an
example of how to do this:</FONT></FONT></P>
<P><FONT face="Courier New"><FONT color=#000080>Dim</FONT> Names() <FONT
color=#000080>As String</FONT><BR><BR><FONT color=#000080>Sub</FONT>
Form1_Load()<BR>    <FONT color=#000080>ReDim</FONT> Names(1 To
10)<BR><FONT color=#000080>End Sub</FONT></FONT></P>
<P></FONT><FONT face=Verdana size=3>This code assigns 10 elements to Names when
Form1 loads. Note: When using dynamic arrays, you must set the size of an array
using the ReDim statement, before filling the array.</FONT><FONT
face=Verdana></P>
<P></FONT><FONT face=Verdana size=3>However, when using the ReDim statement, any
values already in the array (if it has been resized previously), will be
deleted. In some cases, this is not what you would want! So, you use the
Preserve keyword:</FONT><FONT face=Verdana></P>
<P><FONT face="Courier New"><FONT color=#000080>ReDim</FONT> <FONT
color=#000080>Preserve</FONT> ArrayName(LowerValue To HigherValue)</FONT></P>
<P></FONT><FONT face=Verdana size=3>If the array has grown, there will be a
number of blank array spaces at the end of the array. If the array has shrunk,
you will lose the end items.</FONT><FONT face=Verdana></P>
<HR>
<P>Thanks for spending the time to read this Visual Basic Tutorial, since all of
these tutorials were hand-written, there might possibly be some errors in our
technical accuracy, if you have any comments or find one of the errors, please
contact us at <A href="mailto:admin@alphavb.com">admin@alphavb.com</A>. We stand
by our tutorials and will help you the best we can with any Visual Basic
questions also.</P></BODY></HTML>
```

