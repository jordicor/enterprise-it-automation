' Monitorizacion de fichero stock_moviles_new.xls
' Jordi Corrales
' IT Dept Example Corp International
' ---------------------------------------------------

Option Explicit
' Se declaran variables y objetos de vbs para manejar archivos y shell
Dim i, l, oShell, objArgs, objFso, objFile, objFileH, strLine, strPath, strRun, Fso, flcorrct
Dim strpathxls, strpathxls2, strlcxls, flsmtp, xlsnew, strbody
i = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("Wscript.Shell")
xlsnew = "stock_moviles_" & day(now) & month(now) & year(now) & hour(now) & minute(now) & second(now) & ".xls"
strpathxls = "\\172.28.2.108\f80\F80SI\F80MC"
strpathxls2 = strpathxls & "\stock_moviles_new.xls"

' Si se ejecuta sin argumentos (doble click en windows o directamente por cmd)
' se llama a si mismo mediante cscript con argumentos para que asi no se
' vean los cmd al ejecutarse

If WScript.Arguments.Count = 0 Then
 oShell.Run "cmd /k cscript.exe " & wscript.ScriptFullName & " 1", 0, False
 Wscript.quit
End if

' Coje la ruta actual
Set objFileH = objFSO.GetFile(wscript.ScriptFullName)
strPath = objFileH.ParentFolder & "\"
strlcxls = strpath & "stock_moviles_new.xls"

' Si existe hash.dat ya le ha hecho el hash por lo que lo que hace 
' es comprobar si el hash es correcto o no

If (objFso.FileExists("hash.dat")) Then
 strRun = "cmd /c fsum.exe -c -d" & strpathxls & " " & strPath & "hash.dat > chkhsh.tmp"
 oShell.Run strRun, 0, True
 Set objFile = objFSO.OpenTextFile("chkhsh.tmp", 1)
 Do Until objFile.AtEndOfStream
  strLine = objFile.ReadLine
  i = i + 1
  If (instr(1,strLine,"checksums failed")) then
   ' comprobacion fallida, realiza backup del xls, copia el nuevo y avisa por email
   hashea
   sndmail
  End if
 Loop
 objFile.Close
 objFso.DeleteFile("chkhsh.tmp")
else
' Si no existe hash.dat lo genera y copia el xls original
 hashea
End if

wscript.sleep 300000
oShell.Run "cmd /k cscript.exe " & wscript.ScriptFullName & " 1", 0, False
Wscript.quit


' - Fin del programa -

' backup
' Si existe el fichero excel en local se realiza backup
' cambiandole el nombre por la fecha y hora actual
Function backup
 If (objFso.FileExists(strlcxls)) Then
  objFso.MoveFile strlcxls, xlsnew
 End if
 objFso.CopyFile strpathxls2,strpath,1
End Function

' hashea
' Crea el hash del archivo a monitorizar
' Si ha sido modificado o si no existe hash.dat
Function hashea
' mira si existe el fichero en la red antes de generar el hash
 If (objFso.FileExists(strpathxls2)) Then
  strRun = "cmd /c fsum.exe -d" & strpathxls & " stock_moviles_new.xls > hash.dat"
  oShell.Run strRun, 0, True
  backup
 else
  oShell.Run "cmd /c echo No existe stock_moviles_new.xls > err.dmp", 0, True
 End if
End Function

' sndmail
' envia email notificando que el fichero excel ha sido modificado
Function sndmail
 Set flsmtp= objFSO.CreateTextFile("smtp.dat", True)
 flsmtp.WriteLine("HELO stockmon")
 flsmtp.WriteLine("MAIL FROM: stockmon@example-corp.com")
 flsmtp.WriteLine("RCPT TO: jjdcj@example-corp.com")
 flsmtp.WriteLine("DATA")
 flsmtp.WriteLine("To: IT Manager <itmanager@example-corp.com>,Jordi Corrales <jordi.corrales@example-corp.com>")
 flsmtp.WriteLine("Subject: Modificado stock_moviles_new.xls")
 flsmtp.WriteLine("MIME-Version: 1.0")
 flsmtp.WriteLine("X-Mailer: CwT VBS Script 0.69b")
 flsmtp.WriteLine("From: Monitor Stock Moviles <stockmon@example-corp.com>")
 flsmtp.WriteLine(VbCrlf)
 strbody = "Modificacion de fichero " & strpathxls2 & vbcrlf & vbcrlf &_ 
            "Fecha comprobacion: " & now & vbcrlf & "Backup fichero anterior: " & xlsnew & vbcrlf & "." & vbcrlf
 flsmtp.WriteLine(strbody)
 flsmtp.Close
 oShell.Run "cmd /c nc.exe -w 1 172.28.2.5 25<smtp.dat", 0, True
 objFso.DeleteFile("smtp.dat")
End Function
