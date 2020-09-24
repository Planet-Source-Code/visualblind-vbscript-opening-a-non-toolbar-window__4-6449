<div align="center">

## VBScript opening a non\-toolbar window


</div>

### Description

Have you ever wanted to open a new window without a toolbar at the top and a scrollbar on the side? Well this is exactly what this vbscript code does. (please vote for this code if you like it)
 
### More Info
 
Open a new non-toolbar window with vbscript.

nothing at all


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VisualBlind](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/visualblind.md)
**Level**          |Intermediate
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/visualblind-vbscript-opening-a-non-toolbar-window__4-6449/archive/master.zip)

### API Declarations

This code is not copyrighted therefore you may use it at your own will.


### Source Code

```
<Script Language = "vbscript">
Sub B1_OnClick()
Window.open "access.html","fastftp","toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=403,height=250"
B1.Value = " File Downloaded! "
end sub
</script>
```

