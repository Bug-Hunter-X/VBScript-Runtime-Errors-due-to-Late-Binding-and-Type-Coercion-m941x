Late Binding: VBScript's late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where the expected interface might not be available.  Example:
```vbscript
Set objExcel = CreateObject("Excel.Application")
' ... code that assumes Excel is installed ...
```
If Excel isn't installed, this will fail at runtime. 