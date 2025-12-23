Set wshNetwork = wscript.CreateObject("WScript.Network") 
set WShell = CreateObject("WScript.Shell") 
dim ruta
Set ADSysInfo = CreateObject("ADSystemInfo") 
Set CurrentUser = GetObject("LDAP://" & ADSysInfo.UserName) 
strGroups = LCase(Join(CurrentUser.MemberOf)) 
server = inputbox("Numero en el que acaba el server? ~(o`-´o)~ PIKAA ")
dpto = mid(strGroups,4,5)
planta = mid(strGroups,4,3)
ruta = "\\F250210" & server & "\" & planta & "\" 

wshNetwork.MapNetworkDrive "F:",ruta & dpto & "\" & wshNetwork.username
wshNetwork.MapNetworkDrive "G:",ruta & dpto & "\" & dpto
wshNetwork.MapNetworkDrive "H:",ruta & planta

wScript.sleep 5000
wShell.sendkeys "%{TAB}"
wShell.sendkeys "F"
wShell.sendkeys "{F2}"
wShell.sendkeys "PERSONAL {ENTER}"

wShell.sendkeys "{DOWN}"
wShell.sendkeys "{F2}"
wShell.sendkeys "DEPARTAMENTO {ENTER}"

wShell.sendkeys "{DOWN}"
wShell.sendkeys "{F2}"
wShell.sendkeys "PLANTA {ENTER}"

wShell.sendkeys "%{F4}"
wShell.Appactivate "C:\WINNT\System32\cmd.exe"
wShell.sendkeys "exit {ENTER}"

wShell.sendkeys "^{ESC}"
wShell.sendkeys "{UP}"
wShell.sendkeys "{UP}"
wShell.sendkeys "{UP}"
wShell.sendkeys "{UP}"
wShell.sendkeys "{UP}"
wShell.sendkeys "{RIGHT}"
wShell.sendkeys "{DOWN}"
wShell.sendkeys "{ENTER}"

