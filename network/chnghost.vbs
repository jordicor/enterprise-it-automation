'#################################################################
'############## Script VBS para cambiar ips de hosts y lmhosts. ##
'#################################################################
'## Jordi Corrales #
'###################

Option explicit
On error resume next

Dim fso, fl, tx, rd, wrt, ent, lmh, TmpDir, WshShell , WshEnv, WinPath, nueva, ficog, SistOp
set fso = CreateObject("Scripting.FileSystemObject")
set WshShell = CreateObject("WScript.Shell")
set WshEnv = WshShell.Environment("Process")
SistOp = WshEnv("ComSpec")
TmpDir = WshEnv("temp")

if inStr(LCase(SistOp),"cmd.exe") then
 WinPath = WshEnv("SystemRoot")&"\system32\drivers\etc\"
else
 WinPath = "c:\windows\"
end if

ent = Winpath & "hosts."
lmh = winpath & "lmhosts."

nueva = "172.28.1.1	I58DV734	i58dv734"
ficog = "172.28.1.1	CORPSERVER	corpserver"

If fso.FileExists(ent) then
 set tx = fso.OpenTextFile(ent,1)
  fl = TmpDir & "\hosts.bk"
    Set wrt = fso.CreateTextFile(fl, True)
    Do While tx.AtEndOfStream <> True
      rd = tx.ReadLine
	if inStr(UCase(rd),"CORPSERVER") then
	  wrt.writeLine(ficog)
	else
	  wrt.write(rd & vbCRLF)
	end if
    Loop
     wrt.writeLine(nueva)
    wrt.Close
 tx.Close

 fso.DeleteFile(ent)
 fso.MoveFile fl , ent
 fso.CopyFile ent , lmh , 1 
end if