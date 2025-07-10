On Error Resume Next

Dim shell, result
Set shell = CreateObject("WScript.Shell")

' Run run.bat with no window (0), do not wait for it to finish (False)
result = shell.Run("run.bat", 0, False)

If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
End If

Set shell = Nothing
