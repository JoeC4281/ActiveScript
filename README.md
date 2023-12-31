# ActiveScript
ActiveScript Plugin for Take Command Console

```dos
plugin /i ActiveScript
Module:      E:\Documents\PureBasic\ActiveScript\plugin\ActiveScript.dll
Name:        ActiveScript
Author:      Joe Caverly
Email:       jlcaverlyca@yahoo.ca
Web:         https://www.twitter.com/JoeC4281
Description: ActiveScript - TCC Plugin written using Purebasic
Implements:  AScript, DateDiffd, DateDiffm, DateDiffy, VB, VBR
Version:     2023.7  Build 8
```
## Commands
- [AScript][1]
- [DateDiffd][2]
- [DateDiffm][3]
- [DateDiffy][4]
- [VB][5]
- [VBR][6]

## AScript
Unlike the internal TCC Script command,
AScript works with GetObject.

Examples:

```dos
TestIt()

Function TestIt
Set fso = CreateObject ("Scripting.FileSystemObject")
Set stdout = fso.GetStandardStream (1)

Set iWMI = GetObject("WinMgmts:Root\Cimv2")
Set colItems = iWMI.ExecQuery("SELECT * FROM Win32_Process")

If Err.number <> 0 Then
  stdout.WriteLine "Error"
End if

Results = ""

For Each objItem in colItems
  Results = Results + objItem.Name + vbCrLf
Next
TestIt = Results
stdout.WriteLine TestIt
End Function
```

Now, run the above script;

```dos
AScript e:\utils\vince.vbs

System Idle Process
System
Secure System
Registry
smss.exe
...
```

AScript allows one to use Windows Scripting Components;

```dos
Set oMath = GetObject("script:e:\utils\math.wsc")
```

## DateDiffd
`DateDiffd` calculates the number of days from your date until today.

Examples of using `datediffd`;
```dos
datediffd 1960/07/08 & datediffd July 8, 1960 && datediffd 07/08/1960
23010
23010
23010
```

---
## DateDiffm
`DateDiffm` calculates the number of months from your date until today.

Examples of using `datediffm`;
```dos
datediffm 1960/07/08 & datediffm July 8, 1960 && datediffm 07/08/1960
756
756
756
```

---
## DateDiffy
`DateDiffy` calculates the number of years from your date until today.

Examples of using `datediffy`;
```dos
datediffy 1960/07/08 & datediffy July 8, 1960 && datediffy 07/08/1960
63
63
63
```

---
## VB
The VB Command allows you to evaluate a VBScript function.

Examples:

```dos
R:\>vb 2022-1957
65
```

```dos
R:\>vb CreateObject("Shell.Application").ToggleDesktop
```

```dos
R:\>vb DateDiff("d", "03/11/1957", Now) / 365.25
65.8069815195072
```
Click on the link for some more [VB OneLiners][7]

The VB command takes what you enter on the command line,
and wraps it into the following script;

```dos
vbs = ~""
vbs + ~"dim fso:"
vbs + ~"set fso=CreateObject(\"Scripting.FileSystemObject\"):"
vbs + ~"set stdout=fso.GetStandardStream(1):"
vbs + ~"Function OutStd(txt):"
vbs + ~"stdout.WriteLine txt:"
vbs + ~"End Function:"
vbs + ~"OutStd("
vbs = vbs + theScript
vbs + ~")"
```

The generated script then simply calculates your command line argument,
and sends the result to STDOUT.

---
## VBR
VBR allows the running of one or more VBScript commands or functions.

In order to output to the console,
VBR begins with the following internal code;

```dos
dim fso:set fso=CreateObject("Scripting.FileSystemObject"):set stdout=fso.GetStandardStream(1):
```

You can now write output to the console;

```dos
vbr stdout.Write "Test" + Chr(9):a=1973:stdout.WriteLine CStr(a+1)
Test    1974
```
---

  [1]: #ascript
  [2]: #datediffd
  [3]: #datediffm
  [4]: #datediffy
  [5]: #vb
  [6]: #vbr
  [7]: VBOneLiners.md
