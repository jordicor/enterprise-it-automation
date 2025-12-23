strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colServiceList = objWMIService.ExecQuery _
    ("Select * from Win32_Service where Name = 'AudioSrv'")
For Each objService in colServiceList
    errReturnCode = objService.Change( , , , , "Disabled")   
Next
