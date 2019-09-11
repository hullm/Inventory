'Created by Matthew Hull 8/7/15

'On Error Resume Next

'Get the inventory database path
Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentFolder = objFSO.GetAbsolutePathName(".")
strCurrentFolder = strCurrentFolder & "\..\..\Database"
strInventoryDatabase = strCurrentFolder & "\Inventory.mdb"

'Create the connection to the inventory database
Set objConnection = CreateObject("ADODB.Connection")
strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strInventoryDatabase & ";"
objConnection.Open strConnection

strSQL = "UPDATE People INNER JOIN (Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag) ON People.ID = Assignments.AssignedTo "
strSQL = strSQL & "SET Devices.Room="
strSQLWHERE = " WHERE Assignments.Active=True And People.HomeRoom='%Teacher%'"

'*******************************************************************************************************
'Kindergarten Grade Devices

objConnection.Execute(strSQL & "'Bennett'" & Replace(strSQLWhere,"%Teacher%","Bennett, Kimberly"))
objConnection.Execute(strSQL & "'Hendry'" & Replace(strSQLWhere,"%Teacher%","Hendry / Lavigne"))

'*******************************************************************************************************
'1st Grade Devices

objConnection.Execute(strSQL & "'Abrantes'" & Replace(strSQLWhere,"%Teacher%","Abrantes / Jaeger"))
objConnection.Execute(strSQL & "'KellyK'" & Replace(strSQLWhere,"%Teacher%","Kelly, Krista"))
objConnection.Execute(strSQL & "'Zehr'" & Replace(strSQLWhere,"%Teacher%","Zehr, Anna"))

'*******************************************************************************************************
'2nd Grade Devices

objConnection.Execute(strSQL &"'Borie'"  & Replace(strSQLWhere,"%Teacher%","Borie, Nicole"))
objConnection.Execute(strSQL &"'Dudla'" & Replace(strSQLWhere,"%Teacher%","Dudla, Kellie"))
objConnection.Execute(strSQL &"'Poetzsch'" & Replace(strSQLWhere,"%Teacher%","Gearing / Poetzsch"))

'*******************************************************************************************************
'3rd Grade Devices

objConnection.Execute(strSQL & "'Allen'" & Replace(strSQLWhere,"%Teacher%","Allen, Jeffrey"))
objConnection.Execute(strSQL & "'Holderman'"& Replace(strSQLWhere,"%Teacher%","Holderman, Emily"))
objConnection.Execute(strSQL & "'Gershen'" & Replace(strSQLWhere,"%Teacher%","Byrne / Gershen" ))

'*******************************************************************************************************
'4th Grade Devices

objConnection.Execute(strSQL & "'Thomsen'" & Replace(strSQLWhere,"%Teacher%","Drapeau / Thomsen"))
objConnection.Execute(strSQL & "'Lindsay'" & Replace(strSQLWhere,"%Teacher%","Lindsay, Lisa"))

'*******************************************************************************************************
'5th Grade Devices

objConnection.Execute(strSQL & "'Crotty'" & Replace(strSQLWhere,"%Teacher%","Crotty, Jeffrey"))
objConnection.Execute(strSQL & "'Hoover'" & Replace(strSQLWhere,"%Teacher%","Hoover, Erik"))
objConnection.Execute(strSQL & "'Montesano'" & Replace(strSQLWhere,"%Teacher%","Montesano, Kelly"))

'*******************************************************************************************************
'6th Grade Devices

objConnection.Execute(strSQL & "'Butler'" & Replace(strSQLWhere,"%Teacher%","Butler, Matthew"))
objConnection.Execute(strSQL & "'Catarelli'" & Replace(strSQLWhere,"%Teacher%","Catarelli / Spring"))
objConnection.Execute(strSQL & "'Lewis'" & Replace(strSQLWhere,"%Teacher%","Lewis, Jonathan"))

'*******************************************************************************************************

MsgBox "Done"