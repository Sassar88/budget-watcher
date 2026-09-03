Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
strPython = strDir & "\.venv\Scripts\python.exe"
If Not objFSO.FileExists(strPython) Then
    strPython = "python"
End If

objShell.Run "cmd /c cd /d """ & strDir & """ && """ & strPython & """ start.py", 0, True
