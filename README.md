<div align="center">

## Create functions with a variable amount of parameters


</div>

### Description

Explains how to create a function that will accept a varying amount of parameters
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ciarán Walsh](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ciar-n-walsh.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ciar-n-walsh-create-functions-with-a-variable-amount-of-parameters__1-35986/archive/master.zip)





### Source Code

Sometimes it can be very useful to have a function that will accept any number of parameters. For example, say you wanted to join a number of strings into one string, but with a semicolon in between each one. You could do it like this:<BR>
<BR>
JoinedString = String1 & “;” & String2 & “;” & String3<BR>
<BR>
However, if you were doing the same thing more than once, it would be easier to use a function to do the job, like this: <BR>
<BR>
Function Join(Seperator As String, String1 As String, String2 As String, String3 As String) As String<BR>
  Join = String1 & Seperator & String2 & Seperator & String3 & Seperator<BR>
End Function<BR>
<BR>
Unfortunately, this won’t do the job if you want to join a different number of strings, since you can only pass it three strings to join. You could pass the function an array with the strings you want to join, but then you would need to build up the array before you could call the function. <BR>
<BR>
To create a function that can have a varying number of parameters, you use the ParamArray keyword when you declare a function. The syntax is like this: <BR>
<BR>
Public Function Join(Seperator As String, <B>ParamArray</B> Strings()) As String<BR>
<BR>
This makes Strings an array which can hold any number of parameters. When you use the ParamArray keyword, the variable following it must always be a variant array. It is also optional, so you don’t need to pass any parameters for it if you don’t need to. <BR>
Here’s an example for the Join function from before: <BR>
<BR>
Public Function Join(Seperator As String, ParamArray Strings()As Variant) As String<BR>
  Dim aString As Variant<BR>
  For Each aString In Strings<BR>
    Join = Join & aString & Seperator<BR>
  Next<BR>
End Function<BR>
<BR>
This can be called like so: <BR>
<BR>
Joined = Join(";", "a", "b", "c")<BR>
<BR>
This passes the Join function a semicolon as the Seperator variable, and “a”, “b”, and “c” will populate <BR>the Strings array. Then the join function loops through the array using a For Each loop, and adds the Seperator and a string from the array. Unfortunately, since the ParamArray variable is a Variant, the string used to loop round must also be a Variant. <BR>

