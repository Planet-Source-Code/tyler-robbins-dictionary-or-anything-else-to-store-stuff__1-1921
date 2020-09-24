<div align="center">

## Dictionary, or Anything Else to store stuff


</div>

### Description

you enter a word, or whatever else you want and it brings up a definition, or something else.Note: Saves into the registry key: HKEY_Current_User\Software\VB and VBA Program Setting\ whatever you set the path to
 
### More Info
 
create 4 textboxes name them:

AddName, AddDefine, definition, Word

create 1 label, label3 (make its caption blank)

make 2 command buttons

name them:

LookUp, AddWord


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tyler Robbins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tyler-robbins.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tyler-robbins-dictionary-or-anything-else-to-store-stuff__1-1921/archive/master.zip)





### Source Code

```
Private Sub AddWord_Click()
SaveSetting "Dictionary", "Definitions", AddName, AddDefine 'Saves Your Entry In The Registry
AddName = ""
AddDefine = ""
MsgBox ("Entry Saved")
End Sub
Private Sub LookUp_Click()
Label3.Caption = Word & " Means:"
definition = GetSetting("Dictionary", "Definitions", Word) 'Gets the entry from the registry
If definition = "" Then definition = "No Entry Found" 'if no entry found then it tells you
End Sub
```

