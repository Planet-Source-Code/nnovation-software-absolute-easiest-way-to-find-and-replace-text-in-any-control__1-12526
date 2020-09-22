<div align="center">

## Absolute EASIEST Way To Find And Replace Text in Any Control\!\!


</div>

### Description

Absolute Easiest Way To Find And Replace Text in Any Control - ie: labels, textboxes, basically ANY control that supports text editing!!
 
### More Info
 
just make 3 textbox's and a command button, leave default names.

uses the Replace function in VB6.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\!nnovation Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nnovation-software.md)
**Level**          |Intermediate
**User Rating**    |4.3 (30 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nnovation-software-absolute-easiest-way-to-find-and-replace-text-in-any-control__1-12526/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
Text1.Text = Replace(Text1.Text, Text2.Text, Text3.Text, 1, , vbTextCompare)
'here's how it works:
  ' where text1.text is , thats the source of what ur looking in, ex: a label or text box
  ' where text2.text is , that's what u are looking for
  ' where text3.text is , thats what u want to replace what u find with
  ' leave everything else alone after that
Text2.Text = "Find What?"
Text3.Text = "Replace With What?"
End Sub
Private Sub Form_Load()
Text2.Text = "Find What?"
Text3.Text = "Replace With What?"
Text1.Text = "Type Text in Here"
End Sub
```

