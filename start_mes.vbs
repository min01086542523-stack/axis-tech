' 이 스크립트가 있는 폴더에서 MES를 실행합니다. 폴더를 다른 PC로 옮겨도 동작합니다.
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

folder = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = folder
mainPy = folder & "\main.py"

pythonw = ""
localBase = sh.ExpandEnvironmentStrings("%LocalAppData%") & "\Programs\Python"
If fso.FolderExists(localBase) Then
  For Each dir In fso.GetFolder(localBase).SubFolders
    candidate = dir.Path & "\pythonw.exe"
    If fso.FileExists(candidate) Then
      pythonw = candidate
      Exit For
    End If
  Next
End If

If pythonw = "" Then
  pythonw = "pythonw.exe"
End If

sh.Run """" & pythonw & """ """ & mainPy & """", 0, False
