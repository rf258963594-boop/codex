Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
toolsDir = fso.GetParentFolderName(WScript.ScriptFullName)
projectDir = fso.GetParentFolderName(toolsDir)
python = "C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
serverPath = projectDir & "\app\server.py"
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ""Set-Location '" & projectDir & "'; $env:PYTHONWARNINGS='ignore::DeprecationWarning'; & '" & python & "' '" & serverPath & "'"""
shell.Run cmd, 0, False
