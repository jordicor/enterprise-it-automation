Option Explicit
Dim fso, f, n, i, WshShell , WshEnv, WinPath, StrComputer, objWMIService, colAdapters, objAdapter
Set fso = CreateObject("Scripting.FileSystemObject")
set WshShell = CreateObject("WScript.Shell")
set WshEnv = WshShell.Environment("Process")
WinPath = WshEnv("SystemRoot")&"\System32\oeminfo.ini"

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\"& strComputer & "\root\cimv2")
Set colAdapters = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")


Set f = fso.OpenTextFile(WinPath, 2, True)
f.Write "[General]" & chr(13) & chr(10)
f.Write "Manufacturer=Example Corp International" & chr(13) & chr(10)
 
n = 1
 
For Each objAdapter in colAdapters
 
   If Not IsNull(objAdapter.IPAddress) Then
      For i = 0 To UBound(objAdapter.IPAddress)
         f.Write "Model=Direccion IP:" & objAdapter.IPAddress(i) & chr(13) & chr(10)
      Next
   End If

  n = n + 1
 
Next

f.Close