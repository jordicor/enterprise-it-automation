' update one file with another

' if modification time is later

Function Update( source, target )

Dim f1,f2,d1,d2,c1,c2

if fs.FileExists( source ) then

   ' source file is accessible

   set f1 = fs.GetFile( source )

   d1 = f1.DateLastModified

   c1 = Year(d1) * 10000 + Month(d1) * 100 + Day(d1)

   if fs.FileExists( target ) then

      set f2 = fs.GetFile( target )

      d2 = f2.DateLastModified

      c2 = Year(d2) * 10000 + Month(d2) * 100 + Day(d2)

   else

      c2 = 0

   end if

   if c1 > c2 then

      ' overwrite local copy with new version

      f1.Copy target,True

   end if

end if

End Function

' begin script execution

Dim fs

set fs = WScript.CreateObject("Scripting.FileSystemObject")

s = "\\Server\Server_c\AVP\sigfile.dat"

t = "C:\AVP\sigfile.dat"

update s, t

