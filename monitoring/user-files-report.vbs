'******************* Script para buscar en todas las unidades de red ***********
'********************* ficheros en uso por el usuario activo *******************
'** Jordi Corrales *
'******************* 

Option Explicit
On error resume next

'********** Inicia programas
Main
Public fso,WshNetwork,WshShell,WshEnv,IEObj,outputWin,oDrives,i,totalkb,totalmb,curruser,TmpDir,fl,wrt,fl2

'********** Funcion main
Sub Main()

 set fso = CreateObject("Scripting.FileSystemObject")
 Set WshNetwork = WScript.CreateObject("WScript.Network")
 set WshShell = CreateObject("WScript.Shell")
 set WshEnv = WshShell.Environment("Process")
 TmpDir = WshEnv("temp")

' Define valores para crear ventana de explorer
 Set IEObj=CreateObject("InternetExplorer.Application")
 IEObj.Navigate "about:blank"
 IEObj.Height=600
 IEObj.Width=800
 IEObj.MenuBar=False
 IEObj.StatusBar=False
 IEObj.ToolBar=0
 set outputWin=IEObj.Document

 outputWin.Writeln "<HTML><HEAD><TITLE>Ficheros en uso</TITLE></HEAD><BODY><br>"


' Define archivo donde escribirá la informacion en un .txt

 curruser = wshNetwork.userDomain & "\" & wshNetwork.userName
 'fl = split(curruser,"\")
 'fl2 = fl(1)
 fl2 = wshNetwork.userName
 fl = TmpDir & "\" & fl2 & ".txt"
 Set wrt = fso.CreateTextFile(fl, True)

 Unidades()

 wrt.close

End Sub


'********** Lista unidades de red mapeadas
Sub Unidades()

 Set oDrives = WshNetwork.EnumNetworkDrives

 outputWin.Writeln "<center><h1>Listado ficheros en uso de<br>" & curruser & "</h1></center><br><hr>Se generará el archivo " & fl & " con la información mostrada<br><br><b>NO cerrar esta ventana</b> hasta mensaje con la informacion solicitada<br><hr><br><br>"

 For i = 0 to oDrives.Count - 1 Step 2 ' Busca en cada una de las unidades de red
  outputWin.Writeln "<TABLE BORDER=1 WIDTH=150><TR><TD WIDTH=100><b>" & oDrives.Item(i) & "</b></TD></TR>" ' Pone de cabecera de la tabla la unidad que mira
  wrt.write(oDrives.Item(i) & vbCRLF)
  fndOwner 0,oDrives.Item(i) ' Busca carpetas al nombre del usuario 
  fndOwner 1,oDrives.Item(i) ' Busca ficheros al nombre del usuario
  fndOwner 2,oDrives.Item(i) ' Busca recursivamente
  outputWin.Writeln "</TR></TABLE><br>"
  wrt.write(vbCRLF)
 Next

 totalmb = int(totalkb / 1024)

 outputWin.Writeln "</tr></table><br><br><hr><h2>Total Espacio<br>" & curruser & " --- " & totalmb & "MB</h2><br></BODY></HTML>"
 wrt.write("Total Espacio " & curruser & "," & totalmb & " MB" & vbCRLF)
 wscript.echo "Total espacio usado" & vbCRLF & curruser & " -- " & totalmb & " MB"

End Sub


'********** Mira el owner del fichero recursivamente
Sub fndOwner(fndfiles,strFolder)
 Dim strFileName,objShell,objFolder,folds,filowner,iszip,fullpath,isfoldr,strkb,strkbs

 Set objShell = CreateObject ("Shell.Application")
 Set objFolder = objShell.Namespace (strFolder)

' Mira cada fichero en la carpeta , muestra las carpetas o los ficheros segun la variable fndfiles
 For Each strFileName in objFolder.items
  filowner = objFolder.GetDetailsOf(strFileName, 8)
  if curruser = filowner then

' Coje la extension del archivo para ver si es zip para evitar que la tome como carpeta
    iszip = right(strFileName,3)

    if iszip <> "zip" then
     if strFileName.isFolder then
      isfoldr = 1
     else
      isfoldr = 0
     end if
    else
     isfoldr = 0
    end if

    if fndfiles = 0 then
     if isfoldr = 1 then
       outputWin.Writeln "<TD WIDTH=100>" & strFileName & "</TD></TR>" ' Pone las carpetas
       wrt.write(strFileName & "," & vbCRLF)
     end if
    else
     if isfoldr = 0 then
      if fndfiles = 1 then ' Si esta listando los archivos (no las carpetas ni recursivamente) suma el tamaño total en uso
       strkb = objFolder.GetDetailsOf(strFileName, 1)
       strkbs = split(strkb," ")
       strkb = int(strkbs(0))
       totalkb = totalkb + strkb ' Para saber el total de espacio en uso
       outputWin.Writeln "<TD WIDTH=100>" & strFileName & "</TD>" & "<TD WIDTH=50>" & objFolder.GetDetailsOf(strFileName, 1) & "</TD></TR>" ' Pone el nombre del fichero en una columna y el tamaño en otra
       wrt.write(strFileName & "," & objFolder.GetDetailsOf(strFileName, 1) & vbCRLF)
      end if 
     end if
    end if

  end if
 Next

 If fndfiles = 2 then
  IEObj.Visible=True
  For Each strFileName in objFolder.items
   filowner = objFolder.GetDetailsOf(strFileName, 8)
   if curruser = filowner then

    iszip = right(strFileName,3)

    if iszip <> "zip" then
     if strFileName.isFolder then
      outputWin.Writeln "</TR></TABLE><BR><TABLE BORDER=1 WIDTH=150><TR><TD WIDTH=100><b>" & strFileName & "</b></TD></TR>"
      wrt.write(strFileName & "," & vbCRLF)
      fndOwner 0,strFileName
      fndOwner 1,strFileName
      fndOwner 2,strFileName
     end if
    end if 

   end if
  Next
 end if

End Sub