'############## Script inicio aplicacion gestion y control #######
'#################################################################
'# Jordi Corrales #
'##################

' --- Funcion para copiarse el rar.exe y copiarse el mdb actualizado ---

Option explicit

Function Update( source, target )

Dim oNet , Fso, shell, f1,f2,f3,d1,d2,c1,c2
Set Shell = CreateObject("WScript.Shell") 
Set Fso = CreateObject("Scripting.FileSystemObject")
Set oNet = CreateObject("WScript.Network")
On error resume next

oNet.MapNetworkDrive "Q:", "\\172.28.2.105\F80\F80GC", False
If Not Fso.FolderExists("C:\GESCONTROL") Then 
  Fso.CreateFolder "C:\GESCONTROL"
End If

If fso.FileExists("Q:\rar.exe") then
  if Not fso.FileExists("C:\GESCONTROL\rar.exe") then
    set f3 = fso.GetFile("Q:\rar.exe")
    f3.Copy("c:\gescontrol\")
  end if
End If

if fso.FileExists( source ) then
  set f1 = fso.GetFile( source )
  d1 = f1.DateLastModified
  c1 = Year(d1) * 10000 + Month(d1) * 100 + Day(d1)

  if fso.FileExists( target ) then
     set f2 = fso.GetFile( target )
     d2 = f2.DateLastModified
     c2 = Year(d2) * 10000 + Month(d2) * 100 + Day(d2)
  else
     c2 = 0
  end if

   if c1 > c2 then
      f1.Copy target,True
   end if
end if

End Function


' --- Aqui empieza el programa, llama a la funcion , descomprime y ejecuta la mdb ---
Dim shell,fso
Set Shell = CreateObject("WScript.Shell") 
Set fso = CreateObject("Scripting.FileSystemObject")

Update "Q:\CLTPRYAPL.rar","C:\GESCONTROL\CLTPRYAPL.rar"

If fso.FileExists("Q:\rar.exe") then 
  If fso.FileExists("c:\GESCONTROL\CLTPRYAPL.rar") then 
    Shell.Run "C:\GESCONTROL\rar.exe x -o+ -inul c:\GESCONTROL\CLTPRYAPL.rar c:\GESCONTROL", 0, True
  End If
End If

If fso.FileExists("C:\GESCONTROL\CLTPRYAPL.mdb") then
   Shell.Run "C:\GESCONTROL\CLTPRYAPL.mdb", 3, False
End If