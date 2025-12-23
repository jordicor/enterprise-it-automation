' Declaro variables
Dim IE , wshNetwork , objWMIService , colAdapters , strGroups , oDrives

' Se define pc a conectar wmi
StrComputer = "."
Set objWMIService = GetObject("winmgmts:\\"& strComputer & "\root\cimv2")

' Inicializa objeto para abrir iexplorer
Set IE = CreateObject("InternetExplorer.Application")

' Inicializa objeto para configuracion de red
Set wshNetwork = wscript.CreateObject("WScript.Network") 

' Para buscar unidades de red conectadas
Set oDrives = wshNetwork.EnumNetworkDrives


' Inicializa conexion con el AD para buscar info
Set objSysInfo = CreateObject("ADSystemInfo")
Set objNetwork = CreateObject("Wscript.Network")

strUserPath = "LDAP://" & objSysInfo.UserName
Set objUser = GetObject(strUserPath)


' Busca tarjetas de red con ip asignada
Set colAdapters = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")


' Varoles de la ventana de internet explorer
With IE
        .left=200
        .top=100
        .height=550
        .width=400
        .menubar=0
        .toolbar=0
        .statusBar=0
        .navigate "About:Blank"
        .visible=1
End With

'Espera mientras IE esta ocupado
Do while IE.busy
loop

' Abre iexplorer y muestra datos
With IE.document
        .Open
        .WriteLn "<HTML><HEAD>"
        .WriteLn "<TITLE>Datos de " & wshNetwork.username & "</TITLE></HEAD>"
        .WriteLn "<BODY>"
        .WriteLn "<b>Usuario: </b>" & wshNetwork.username & "<br>"
	.WriteLn "<b>Unidades de red conectadas actualmente</b><br>"

         For i = 0 to oDrives.Count - 1 Step 2
            .WriteLn "* " & oDrives.Item(i) & " = " & oDrives.Item(i+1) & "<br>"
         Next

	For Each objAdapter in colAdapters
	 If Not IsNull(objAdapter.IPAddress) Then
	  For i = 0 To UBound(objAdapter.IPAddress)
	  .WriteLn "<b>Direccion IP: </b>" & objAdapter.IPAddress(i) & "<br>"
	  Next
	 End If
	Next

	.WriteLN "<b>Miembro de: </b><br>"

	For Each strGroup in objUser.MemberOf
	 strGroupPath = "LDAP://" & strGroup
	 Set objGroup = GetObject(strGroupPath)
         strGroupName = objGroup.CN
         .WriteLn "* " & strGroupName 
	 .WriteLn "<br>"
	Next

        .WriteLn "</BODY>"
        .WriteLn "</HTML>"
        .Close
End With
Set IE = Nothing
WScript.Quit(0)