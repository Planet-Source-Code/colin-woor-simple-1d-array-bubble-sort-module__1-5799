<div align="center">

## Simple 1d array bubble sort module


</div>

### Description

Simply sorts a 1 dimensional array using a bubble sort algorythm.
 
### More Info
 
Array to be sorted

A sorted array


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Colin Woor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/colin-woor.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/colin-woor-simple-1d-array-bubble-sort-module__1-5799/archive/master.zip)





### Source Code

```
'
'Use:
'
'Sort Array
'
'to sort (A-Z / 1-10, Accending)
'Pretty easy to update it to sort 2 or 3 dimensional arrays
'Or to sort decending
'
'Comments or any info email: col@woor.co.uk
'
Public Sub sort(tmparray)
Dim SortedArray As Boolean
Dim start, Finish As Integer
SortedArray = True
start = LBound(tmparray)
Finish = UBound(tmparray)
Do
  SortedArray = True
  For loopcount = start To Finish - 1
    If tmparray(loopcount) > tmparray(loopcount + 1) Then
      SortedArray = False
      Call swap(tmparray, loopcount, loopcount + 1)
    End If
  Next loopcount
Loop Until SortedArray = True
End Sub
Sub swap(swparray, fpos, spos)
Dim temp As Variant
temp = swparray(fpos)
swparray(fpos) = swparray(spos)
swparray(spos) = temp
End Sub
```

