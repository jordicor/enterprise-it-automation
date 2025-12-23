const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
Set StdOut = WScript.StdOut
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Policies\Microsoft\Windows\WindowsUpdate"
strValueName = "DonotallowXPSP2"
strValue = "1"

oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue