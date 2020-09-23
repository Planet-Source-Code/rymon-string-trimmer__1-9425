<div align="center">

## String trimmer


</div>

### Description

Just a little snip of code that lets you trim a specified number of characters from either side of a string.
 
### More Info
 
Inputs are (in order):

data to trim from,

number of characters from left side,

number of characters from right side

The new string, w/o trimmed characters


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[RyMon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rymon.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rymon-string-trimmer__1-9425/archive/master.zip)





### Source Code

```
Function trim_data(data As String, from_left As Integer, from_right As Integer) As String
  'If you try to trim to much, returns an empty string
  If Len(data) <= from_left + from_right Then
    trim_data = ""
  'If not, trim text from sides and return
  Else
    trim_data = Mid(data, from_left + 1, Len(data) - from_left - from_right)
  End If
End Function
```

