# ActiveScript
ActiveScript Plugin for Take Command Console

```dos
plugin /i ActiveScript

Module:      e:\utils\ActiveScript.dll
Name:        ActiveScript
Author:      Joe Caverly
Email:       jlcaverlyca@yahoo.ca
Web:         https://www.twitter.com/JoeC4281
Description: ActiveScript - TCC Plugin written using Purebasic
Implements:  AScript, VB
Version:     2022.12  Build 31
```


Version 2022-12-31

Added VB command.

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

----------------------------------------------------------------------
Version 2022-12-29

Initial Release.

AScript is the only command at this time.

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


