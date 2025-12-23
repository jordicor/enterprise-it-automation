' FileName:AtchMail.vbs
Const AttachPath = "C:\Temp\Test.txt"
Dim WS, nSs, nWS, nDb, nDoc, MServer, MFile
Set WS = CreateObject("WScript.Shell")
If Not WS.AppActivate("Lotus Notes") Then
MsgBox "Please start Notes before this script.", , "No Notes"
Set WS = Nothing
WScript.Quit
End If
WS.SendKeys "% r"
Set nSs = CreateObject("Notes.NotesSession")
MServer = nSs.GetEnvironmentString("MailServer", True)
MFile = nSs.GetEnvironmentString("MailFile", True)
Set nWS = CreateObject("Notes.NotesUIWorkspace")
nWS.OpenDatabase MServer, MFile
Set nDoc = nWS.ComposeDocument("", "", "Memo" )
With nDoc
.EditMode = True
' Here set multiple recipients delimited by vbCrLf
.FieldSetText "EnterSendto", nSs.UserName & vbCrLf
.FieldSetText "Subject", "Mail by Script"
.GotoField "Body"
.FieldSetText "Body", "Mail with attachment sending test by script" & vbCrLf
.CreateObject "", "", AttachPath
.Send
End With
Set nDoc = Nothing: Set nWS = Nothing: Set nSs = Nothing