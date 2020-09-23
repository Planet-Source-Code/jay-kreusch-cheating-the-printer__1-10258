<div align="center">

## Cheating the Printer


</div>

### Description

Why mess around with the printer object if you don't have to? In my example, I print the contents of a richtextbox control to the printer with only a couple of lines of code. PERFECTLY formatted. Readily applies to just about any control or string, though.
 
### More Info
 
Sometimes displays the splash screen of another program or a print dialog box


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jay Kreusch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jay-kreusch.md)
**Level**          |Beginner
**User Rating**    |4.7 (33 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jay-kreusch-cheating-the-printer__1-10258/archive/master.zip)

### API Declarations

```
'Used for the shell printing
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
  "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, ByVal lpParameters As String, ByVal _
  lpDirectory As String, ByVal nShowCmd As Long) As Long
'Used to come up with the temp file directory
Private Declare Function GetTempPath Lib "kernel32" _
  Alias "GetTempPathA" (ByVal nBufferLength As Long, _
  ByVal lpBuffer As String) As Long
'used to come up with the temp file name
Private Declare Function GetTempFileName Lib "kernel32" _
  Alias "GetTempFileNameA" (ByVal lpszPath As String, _
  ByVal lpPrefixString As String, ByVal wUnique As Long, _
  ByVal lpTempFileName As String) As Long
```


### Source Code

```
'All you need to provide is a prefix if desired, and the file extention
Private Function CreateTempFile(sPrefix As String, sSuffix As String) As String
  Dim sTmpPath As String * 512
  Dim sTmpName As String * 576
  Dim nRet As Long
  'Some API and string manipulation to get the temp file created
  nRet = GetTempPath(512, sTmpPath)
  If (nRet > 0 And nRet < 512) Then
   nRet = GetTempFileName(sTmpPath, sPrefix, 0, sTmpName)
   If nRet <> 0 Then
     sTmpName = Left$(sTmpName, _
      InStr(sTmpName, vbNullChar) - 1)
      CreateTempFile = Left(Trim(sTmpName), Len(Trim(sTmpName)) - 3) & sSuffix
   End If
  End If
End Function
Private Sub Command1_Click()
  Dim sTmpFile As String
  Dim sMsg As String
  Dim hFile As Long
  'We're trying to print a richtextbox, so give it something to name
  'it by, and make sure you set the extention to rtf.
  'You could print a textbox by using txt, etc.
  sTmpFile = CreateTempFile("jTmp", "rtf")
  'Gets the next available open number
  hFile = FreeFile
  'open the file and give it the textRTF of the richtextbox
  'if you don't want to use boxed, you could just pass a string here
  Open sTmpFile For Binary As hFile
   Put #hFile, , RichTextBox1.TextRTF
  Close hFile
  'shell print it
  Call ShellExecute(0&, "Print", sTmpFile, vbNullString, vbNullString, vbHide)
  'delete it.
  Kill sTmpFile
End Sub
```

