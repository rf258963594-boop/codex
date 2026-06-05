Set shell = CreateObject("WScript.Shell")
If WScript.Arguments.Count < 5 Then
  WScript.Quit 1
End If

Function Q(value)
  Q = Chr(34) & value & Chr(34)
End Function

pythonExe = WScript.Arguments(0)
appPath = WScript.Arguments(1)
port = WScript.Arguments(2)
outLog = WScript.Arguments(3)
errLog = WScript.Arguments(4)
comspec = shell.ExpandEnvironmentStrings("%ComSpec%")

cmd = Q(comspec) & " /c " & Chr(34) & Q(pythonExe) & " " & Q(appPath) & " " & port & " > " & Q(outLog) & " 2> " & Q(errLog) & Chr(34)
shell.Run cmd, 0, False
