<div align="center">

## Enhanced Split Function \(2D\) \(Fixed\)


</div>

### Description

This code is an addon to the "split" function in VB, it takes an string and will split it into a 2D array. It's pretty easy to figure out. I wasn't sure whether to put this as a beginner or intermediate so it is kinda inbetween. If you're using VB5 or before you'll need to download a replacement split function from this site somewhere.

Latest - Fixed a small problem with different sizes of one dimension (could have been fixed with on error resume next but this is more effective)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Robins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-robins.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-robins-enhanced-split-function-2d-fixed__1-14965/archive/master.zip)





### Source Code

```
Private Function Split2D(StringToSplit, FirstDelimiter, SecondDelimiter)
Dim X As Integer, _
  Y As Integer, _
  FirstBound As Integer, _
  SecondBound As Integer, _
  ResultArray()
temparray = Split(StringToSplit, FirstDelimiter)
FirstBound = UBound(temparray)
For X = 0 To FirstBound
  temparray2 = Split(temparray(X), SecondDelimiter)
  If UBound(temparray2) > SecondBound Then SecondBound = UBound(temparray2)
Next
ReDim ResultArray(FirstBound, SecondBound)
For X = 0 To FirstBound
    temparray2 = Split(temparray(X), SecondDelimiter)
  For Y = 0 To UBound(temparray2)
    ResultArray(X, Y) = temparray2(Y)
  Next
Next
Split2D = ResultArray
End Function
```

