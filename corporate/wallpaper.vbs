' ##################### Desactiva servicio de audio de windows #################
' ################## copia logo corporativo y lo pone como fondo ###############
' # Jordi Corrales #
' ##################

On error resume next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

'Set colServiceList = objWMIService.ExecQuery ("Select * from Win32_Service where Name = 'AudioSrv'")
'For Each objService in colServiceList
'    errReturnCode = objService.Change( , , , , "Disabled")   
'Next

Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
 strso = objOperatingSystem.Caption
next

if strso = "Microsoft Windows XP Professional" then
 strValue = "c:\windows\zlogo.bmp"
else
 strValue = "c:\winnt\zlogo.bmp"
end if 

Set Fso = CreateObject("Scripting.FileSystemObject")
strFile = "\\172.28.17.34\temp\zlogo.bmp"
set f1 = Fso.GetFile(strFile)
f1.Copy strValue,true


Set WshShell = WScript.CreateObject("Wscript.Shell")

WshShell.RegWrite "HKCU\Control Panel\colors\Background","0 0 0"
WshShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper",strValue
WshShell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle","0","REG_DWORD"
WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System\NoDispBackgroundPage","1","REG_DWORD"
WshShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters"