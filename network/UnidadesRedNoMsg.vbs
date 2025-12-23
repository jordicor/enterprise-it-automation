' Declaro variables
 Option Explicit
 Dim fso , objSysInfo , objNetwork , ObjUser , strGroupPath , objGroup
 Dim strLenDiv , strPathPer, strPathDep , strPathPla , strEmpLet , strUserPath , strGroup , strGroupName , fndPers
 Public wshNetwork , strSrvLet , alrConn

' Se define objeto de archivos
 Set Fso = CreateObject("Scripting.FileSystemObject")

' Inicializa objeto para configuracion de red
 Set wshNetwork = wscript.CreateObject("WScript.Network") 

' Inicializa conexion con el AD para buscar info
 Set objSysInfo = CreateObject("ADSystemInfo")
 Set objNetwork = CreateObject("Wscript.Network")

 strUserPath = "LDAP://" & objSysInfo.UserName
 Set objUser = GetObject(strUserPath)
 
'Miembro de..

 For Each strGroup in objUser.MemberOf
  strGroupPath = "LDAP://" & strGroup
  Set objGroup = GetObject(strGroupPath)
  strGroupName = objGroup.CN ' strGroupName -> nombre grupo al que pertenece
  strLenDiv = len(strGroupName)
       
  If strLenDiv = 5 THEN

   strEmpLet = left(strGroupName,3)

   servidores(strEmpLet) ' devuelve el servidor a conectar
           
   strPathPer = "\\" & strSrvLet & "\" & strEmpLet & "\" & strGroupName & "\" & wshNetwork.UserName ' Personal
   strPathDep = "\\" & strSrvLet & "\" & strEmpLet & "\" & strGroupName & "\" & strGroupName ' Departamento
   strPathPla = "\\" & strSrvLet & "\" & strEmpLet & "\" & strEmpLet ' Planta
   
  ' Conectamos unidades de red
   fndPers = 0
   
   yaConn(strPathPer)
   If alrConn = 0 then
    If fso.FolderExists(strPathPer) then 
     Conectar 70, strPathPer
     fndPers = 1 
    End if
   End if 

   yaConn(strPathDep)
   If alrConn = 0 then
    If fso.FolderExists(strPathDep) then 
     Conectar 71,strPathDep 
    End if
   End if
   
   yaConn(strPathPla)
   If alrConn = 0 then
    If fso.FolderExists(strPathPla) then 
     Conectar 72,strPathPla 
    End if
   End if
      
  End IF 
       
 Next


WScript.Quit(0)

' Devuelve servidor segun division
Sub Servidores(strEmpLet)
 Select Case strEmpLet
  Case "F02" strSrvLet = "I5835437"
  Case "F04" strSrvLet = "I5835436"
  Case "F25" strSrvLet = "I5835434"
  Case "F59" strSrvLet = "I5835433"
  Case "F80" strSrvLet = "I5835438"
 End Select
End Sub

' Mira si la unidad a conectar ya esta conectada
Sub yaConn(strPathUn)
 Dim i , Odrives
 
 Set oDrives = wshNetwork.EnumNetworkDrives
 alrConn = 0
 
 For i = 0 to oDrives.Count - 1 Step 2
  If oDrives.Item(i+1) = strPathUn then
   ' msgbox "La ruta " & strPathUn & " ya esta conectada en " & oDrives.Item(i)
   alrConn = 1
  End if
 Next
 
End Sub

' Conecta unidades 
Sub Conectar(intPathLet ,pathUnid)
 Dim ConnUnid , strPathLet
  
 ConnUnid = 0
  
 Do While ConnUnid = 0 ' Mientras no est� conectada la unidad sigue buscando la siguiente letra libre
  strPathLet = chr(intPathLet)
  If fso.DriveExists(strPathLet) then
   intPathLet = intPathLet + 1 ' Aumenta ASCII para siguiente letra
  else
   wshNetwork.MapNetworkDrive strPathLet & ":", pathUnid , 1
   ConnUnid = 1
  End if
 Loop
 
End Sub