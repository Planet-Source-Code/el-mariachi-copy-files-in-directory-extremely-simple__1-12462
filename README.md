<div align="center">

## Copy files in directory \- Extremely simple\!\!


</div>

### Description

With this code you can copy all files within a directory with a particular extension to another directory. Actually it will give you a progressive listing of a file with a wildcard and then you can do whatever you want with the file. Extremely simple and straightforward. I dont understand why the other codes are so huge. If you find an error or if you would like to comment please leave a messge.
 
### More Info
 
Create a button on your form and call it Command1. Then create a folder "c:\tempo" and cretae a bunch of files in it with , say, a .doc extension or whatever you want. Then create another folder "c:\tempx" and dont put anything in there.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[El Mariachi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/el-mariachi.md)
**Level**          |Intermediate
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/el-mariachi-copy-files-in-directory-extremely-simple__1-12462/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
 Dim fName As String
 fName = Dir("c:\tempo\*.doc") ' Retrieve the first entry.
 Do While fName <> "" ' Start the loop.
 If GetAttr("c:\tempo\" & fName) <> vbDirectory Then 'only files
  FileCopy "c:\tempo\" & fname, "c:\tempx\" & fName 'copies the file
  'Kill "c:\tempo\" & fname 'deletes the original - optional
 End If
 fName = Dir ' Get next entry.
 Loop
End Sub
```

