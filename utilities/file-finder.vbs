Dim System
Dim Drive
ReDim Location(0)
Dim FileName
FileName = LCase(Trim(InputBox("Enter the name of the" &_
                  " file that you wish to search for.")))
If Len(FileName) = 0 Then Wscript.Quit
Set System = CreateObject("Scripting.FileSystemObject")
For Each Drive In System.Drives
   If Drive.IsReady And Drive.DriveType = 2 Then
      Call FindFile(Drive & "\")
   End If
Next 'Drive
Msgbox Join(Location,vbCr)


Sub FindFile(ThisFolder)
    Dim File
    Dim Folder
    For Each Folder In System.GetFolder(ThisFolder).SubFolders
       For Each File In Folder.Files
          If LCase(File.Name) = FileName Then
             Location(Ubound(Location)) = File
             ReDim Preserve Location(Ubound(Location) + 1)
          End If
       Next 'File
       Call FindFile(Folder)
    Next 'Folder
End Sub