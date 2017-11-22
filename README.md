# LotusScript-BigStringWalker
Very fast Mid$() alternative for LotusScript for working with large Strings

Superior speed! Tested with strings up to 60 Million characters in size. Estimated to handle up to 100 Million characters before it becomes noticeably slow on today's hardware (as of 2017).

This class depends on Julian Robichaux's StringBuffer class (named here as "Buffer"). It can be obtained from here: http://www.nsftools.com/tips/NotesTips.htm#stringbuffer

Running the following sample code using Mid$() to extract indivdual characters can take 10s of magnitude more time to complete.


Sample code:
```vb
'This code loads a big string, then processes it character by character:
Dim BSW As New BigStringWalker

BSW.Load(MyBigString)

Dim x As Long
Dim Char As String

For x = 0 To BSW.Size-1
  Char = BSW.CharAt(x)
  Call ProcessChar(Char)
Next
```
