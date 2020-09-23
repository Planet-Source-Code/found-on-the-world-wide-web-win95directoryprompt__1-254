<div align="center">

## Win95DirectoryPrompt


</div>

### Description

Prompting the User for a Directory in Win95. Windows' common dialogs are great if you want the user to select a file, but what if you want them to select a directory? Call the following function, which relies on Win32's new SHBrowseForFolder function:
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-win95directoryprompt__1-254/archive/master.zip)

### API Declarations

```
Private Type BrowseInfo
  hWndOwner   As Long
  pIDLRoot    As Long
  pszDisplayName As Long
  lpszTitle   As Long
  ulFlags    As Long
  lpfnCallback  As Long
  lParam     As Long
  iImage     As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, ByVal lpBuffer As String) As Long
```


### Source Code

```
Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String
  Dim iNull As Integer
  Dim lpIDList As Long
  Dim lResult As Long
  Dim sPath As String
  Dim udtBI As BrowseInfo
  With udtBI
    .hWndOwner = hWndOwner
    .lpszTitle = lstrcat(sPrompt, "")
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With
  lpIDList = SHBrowseForFolder(udtBI)
  If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    lResult = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    iNull = InStr(sPath, vbNullChar)
    If iNull Then
      sPath = Left$(sPath, iNull - 1)
    End If
  End If
  BrowseForFolder = sPath
End Function
```

