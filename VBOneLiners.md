# VBScript One-Liners
## For use with the ActiveScript plugin for TCC

```dos
vb 140.97*52
7330.44
```


```dos
vb DateAdd("m", 708, "03-Aug-1961")
2020-08-03
```


```dos
vb DateDiff("d", "08/03/1961", Now) / 365.25
61.9028062970568
```
The following retrieves text that is on the Windows Clipboard;
```dos
echo Stuff on the clipboard > clip:

vb CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
Stuff on the clipboard
```

The following works with my *jlcUtils.dll* COM Server,<br>
which I developed using Visual Basic 6.0 32-bit;

```dos
vb CreateObject("jlcutils.clsmath").ppp(6.59)
2.99186
```
The following writes the current week number to StdOut;
```dos
vb CreateObject("scripting.filesystemobject").GetStandardStream(1).writeline(DatePart("ww",Now()))
26
```
More...
```dos
vb CreateObject("Shell.Application").ToggleDesktop
```

```dos
vb CreateObject("Win32api.Kernel32").GetWindowsDirectory()
```
The following gets the value of the *username* environment variable;
```dos
vb CreateObject("WScript.Shell").Environment("Process")("username")
jlcav
```
Even More...
```dos
vb MonthName(6)
June
```

```dos
vb FormatPercent(2/32)
6.25%
```

```dos
vb FormatCurrency(1000)
$1,000.00
```


<br>
