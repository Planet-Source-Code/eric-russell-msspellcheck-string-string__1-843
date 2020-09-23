<div align="center">

## MsSpellCheck\( string \) : string


</div>

### Description

This short and sweet function accepts a string containing text to be

spell checked, checks the text for spelling using MS Word automation,

and then returns the processed text as a string. The familiar

MS Word spelling dialog will allow the user to perform actions such

as selecting from suggested spellings, ignore, adding the word to a

customized dictionary, etc.
 
### More Info
 
String - Text to be checked for spelling

You need to have Microsoft Word95 or higher installed on the PC. Just place the function in a project module or the general declaration section of a form.

String - Text after modification by user from the Word spell checking dialog.

There are no known side effects.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Eric Russell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eric-russell.md)
**Level**          |Unknown
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eric-russell-msspellcheck-string-string__1-843/archive/master.zip)





### Source Code

```
' Description: This function accepts a string containing text to be
' spell checked, checks the text for spelling using MS Word automation,
' and then returns the processed text as a string. The familiar
' MS Word spelling dialog will allow the user to perform actions such
' as selecting from suggested spellings, ignore, adding the word to a
' customized dictionary, etc.
'    Syntax: MsSpellCheck( String ) : String
'    Author: Eric Russell
'    E-Mail: erussell@cris.com
'   WEB Site: http://cris.com/~erussell/VisualBasic
'   Created: 1998-13-14
'   Revised: 1998-04-03
'Compatibility: VB 5.0, VB 4.0(32bit)
' Assumptions: The user must have MS Word95 or higher installed on
'their PC.
'  References: Visual Basic For Applications, Visual Basic runtime
'objects and procedures, Visual Basic objects and procedures.
'
Function MsSpellCheck(strText As String) As String
Dim oWord As Object
Dim strSelection As String
Set oWord = CreateObject("Word.Basic")
oWord.AppMinimize
MsSpellCheck = strText
oWord.FileNewDefault
oWord.EditSelectAll
oWord.EditCut
oWord.Insert strText
oWord.StartOfDocument
On Error Resume Next
oWord.ToolsSpelling
On Error GoTo 0
oWord.EditSelectAll
strSelection = oWord.Selection$
If Mid(strSelection, Len(strSelection), 1) = Chr(13) Then
 strSelection = Mid(strSelection, 1, Len(strSelection) - 1)
End If
If Len(strSelection) > 1 Then
 MsSpellCheck = strSelection
End If
oWord.FileCloseAll 2
oWord.AppClose
Set oWord = Nothing
End Function
```

