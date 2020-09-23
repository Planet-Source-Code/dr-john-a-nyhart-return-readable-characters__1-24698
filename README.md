<div align="center">

## Return Readable Characters


</div>

### Description

This small function will strip out the unreadable

characters and return readable characters in a string.

This works great for stripping thoes little boxes

out of the text that you just snagged from a web page.

I know that a lot of you have used my code,

but have not voted. It would be nice to see

some feedback in the form of a vote. Thanks.
 
### More Info
 
Text


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dr\. John A\. Nyhart](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dr-john-a-nyhart.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dr-john-a-nyhart-return-readable-characters__1-24698/archive/master.zip)





### Source Code

```
Function ReturnReadableChrsOnly(sString As String) As String
  Dim lCount As Long
  Dim lPoint As Long
  Dim sTmp As String
  Dim sChr As String
  lCount = Len(sString)           ' Get a count of chars
  For lPoint = 1 To lCount          ' Loop through that count
    sChr = Mid(sString, lPoint, 1)     ' Get one chr
    Select Case Asc(sChr)         ' Set a case for the ASCII value of char
      Case 32 To 126           ' Is this an Readable Char?
        sTmp = sTmp & sChr       ' If yes then append to list
    End Select '»Select Case Asc(sChr)
  Next '»For lPoint = 1 To lCount
  ReturnReadableChrsOnly = sTmp
End Function
```

