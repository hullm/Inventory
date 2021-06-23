<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/11/18
'Last Updated 6/11/18

'This is the API for the inventory site

Option Explicit

'On Error Resume Next

Dim strLookupType, strUserName, strSQL, objLookUp, strSerialNumber, strAssetTag, objDevice, strOutput

strLookupType = Request.QueryString("Type")

Select Case strLookupType

	Case "DNSLookup"
		'?Status=DNSLookup&Serial=ABC1234 --> true/false
		Response.Write DNSLookup(Request.QueryString("Serial"))

	Case "ProxyLookup"
		'?Status=ProxyLookup&Serial=ABC1234 --> true/false
		Response.Write ProxyLookup(Request.QueryString("Serial"))

	Case "StudentID"
		'?Type=StudentID&UserName=jsmith --> 12345
		Response.Write GetStudentID(Request.QueryString("UserName"))

	Case "ComputerName"
		'?Type=ComputerName&Serial=ABC1234 --> HS-Lastname-01234
		Response.Write GetComputerName(Request.QueryString("Serial"))

	Case "AssetTag"
		'?Type=AssetTag&Serial=ABC1234 --> 1234
		Response.Write GetAssetTag(Request.QueryString("Serial"))

	Case "Tags"
		'?Type=Tags&AssetTag=1234 --> Spare,Silver
		Response.Write GetTags(Request.QueryString("AssetTag"))

	Case "Backup"
		'?Type=Backup&Serial=ABC1234 --> Backup/NoBackup
		Response.Write GetBackupInformation(Request.QueryString("Serial"))
	
	Case "GenerateComputerName"
		'?Type=GenerateComputerName&Serial=ABC1234 --> HS-Lastname-01234
		Response.Write GenerateComputerName(Request.QueryString("Serial"))

	Case "GetAssignedUserFullName"
		'?Type=GetAssignedDeviceUser&Serial=ABC1234 --> Tess Armstrong'
		Response.Write GetAssignedUserFullName(Request.QueryString("Serial"))

	Case "GetAssignedUsername"
		'?Type=GetAssignedDeviceUser&Serial=ABC1234 --> armstrongt'
		Response.Write GetAssignedUsername(Request.QueryString("Serial"))
		
	Case "AlertInfected"
		'?Type=AlertInfected&Serial=ABC1234
		SendAlertInfectedEmail(Request.QueryString("Serial"))

End Select

Function DNSLookup(strSerialNumber)

	'Checks to see if the "Disable DNS" tag is on a device.  This is used to either enable or disable 
	'the external pihole filter. 

	strSQL = "SELECT LGTag FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)
	
	If Not objLookup.EOF Then
		strSQL = "SELECT ID FROM Tags WHERE Tag='Disable DNS' AND LGTag='" & objLookUp(0) & "'"
		Set objLookUp = Application("Connection").Execute(strSQL)

		If Not objLookup.EOF Then
			DNSLookup = "true"
		Else
			DNSLookup = "false"
		End If

	End If

End Function

Function ProxyLookup(strSerialNumber)

	'Checks to see if the "Disable Proxy" tag is on the device.  This is used to either enable or disable
	'the proxy settings on the device.

	strSQL = "SELECT LGTag FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)
	
	If Not objLookup.EOF Then
		strSQL = "SELECT ID FROM Tags WHERE Tag='Disable Proxy' AND LGTag='" & objLookUp(0) & "'"
		Set objLookUp = Application("Connection").Execute(strSQL)

		If Not objLookup.EOF Then
			ProxyLookup = "true"
		Else
			ProxyLookup = "false"
		End If

	End If

End Function

Function GetStudentID(strUserName)

	'This will return the studentID of the user
	
	strSQL = "SELECT StudentID FROM PEOPLE WHERE UserName='" & strUserName & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)
	
	If Not objLookUp.EOF Then
		GetStudentID = objLookUp(0)
	End If

End Function

Function GetComputerName(strSerialNumber)

	'This will return the computer name stored in the database

	strSQL = "SELECT ComputerName FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)
	
	If Not objLookUp.EOF Then
		GetComputerName = objLookUp(0)
	End If
End Function

Function GetAssetTag(strSerialNumber)

	'This will return the asset tag of the device

	strSQL = "SELECT LGTag FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)
	
	If Not objLookup.EOF Then
		GetAssetTag = objLookUp(0)
	End If

End Function

Function GetTags(strAssetTag)

	'This will return a list of the tags associated with the device

	strSQL = "SELECT Tag FROM Tags WHERE LGTag='" & strAssetTag & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)
	
	If Not objLookup.EOF Then
		Do Until objLookUp.EOF
			strOutput = strOutput & objLookUp(0) & ", "
			objLookup.MoveNext
		Loop
	End If

	If strOutput <> "" Then
		strOutput = Left(strOutput,Len(strOutput)-2)
	End If

	GetTags = strOutput

End Function

Function GetBackupInformation(strSerialNumber)

	'Determine if the backup should be forced, or skipped on a device

	strAssetTag = GetAssetTag(strSerialNumber)

	strSQL = "Select ID FROM Tags WHERE Tag='Backup' AND LGTag='" & strAssetTag & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)

	If Not objLookUp.EOF Then
		GetBackupInformation = "Backup"
	Else

		strSQL = "Select ID FROM Tags WHERE Tag='NoBackup' AND LGTag='" & strAssetTag & "'"
		Set objLookUp = Application("Connection").Execute(strSQL)

		If Not objLookUp.EOF Then
			GetBackupInformation = "NoBackup"
		End If

	End If

End Function

Function GetAssignedUser(strSerialNumber)

	'This will return the curently assigned user

	strSQL = "Select People.LastName, People.FirstName" & vbCRLF
	strSQL = strSQL & "FROM People INNER JOIN (Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTAG) On People.ID = Assignments.AssignedTo" & vbCRLF
	strSQL = strSQL & "WHERE Assignments.Active=True AND SerialNumber='" & strSerialNumber & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)

	If Not objLookUp.EOF Then
		GetAssignedUser = objLookUp(0) & ", " & objLookUp(1)
	End If

End Function

Function GetAssignedUserFullName(strSerialNumber)

	dim strAssignedFullName, arrAssignedFullName

	strAssignedFullName = GetAssignedUser(strSerialNumber)
	arrAssignedFullName = Split(strAssignedFullName,",")
	GetAssignedUserFullName = arrAssignedFullName(1)&" "&arrAssignedFullName(0)

End Function


Function GetAssignedUsername(strSerialNumber)

	'This will return the curently assgined user

	strSQL = "SELECT People.UserName" & vbCRLF
	strSQL = strSQL & "FROM People INNER JOIN (Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTAG) On People.ID = Assignments.AssignedTo" & vbCRLF
	strSQL = strSQL & "WHERE Assignments.Active=True AND SerialNumber='" & strSerialNumber & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)

	If Not objLookUp.EOF Then
		GetAssignedUsername = objLookUp(0)
	End If

End Function

Function GetDeviceSite(strSerialNumber)

	'This will return the currently assigned site for a device

	strSQL = "SELECT Site FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)

	If Not objLookUp.EOF Then
		GetDeviceSite = objLookUp(0)
	End If

End Function

Function GetRole(strUserName)

	'Get the role of the user, student or teacher

	strSQL = "SELECT Role FROM PEOPLE WHERE UserName='" & strUserName & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)

	If Not objLookUp.EOF Then
		GetRole = objLookUp(0)
	End If

End Function

Function GetRoom(strSerialNumber)

	'Get the location for the device

	strSQL = "SELECT Room FROM Devices Where SerialNumber='" & strSerialNumber & "'"
	Set objLookUp = Application("Connection").Execute(strSQL)

	If Not objLookUp.EOF Then
		GetRoom = objLookUp(0)
	End If

End Function

Function GenerateComputerName(strSerialNumber)

	Dim strCurrentComputerName, strAssetTag, strTags, strUser,strLastName, arrUser, strSite, strSitePrefix
	Dim strFixedAssetTag, strUserName, strRole, arrTags, strTag, intGraduationYear, strRoom

	'Get the information needed to generate a name
	strCurrentComputerName = GetComputerName(strSerialNumber)
	strAssetTag = GetAssetTag(strSerialNumber)
	strTags = GetTags(strAssetTag)
	strSite = GetDeviceSite(strSerialNumber)
	strUser = GetAssignedUser(strSerialNumber)
	strUserName =GetAssignedUsername(strSerialNumber)
	strRole = GetRole(strUserName)
	strRoom = GetRoom(strSerialNumber)

	'Get the last name from the strUser variable, it's in lastname, firstname format
	If strUser <> "" Then
		arrUser = Split(strUser,",")
		strLastName = arrUser(0)
	End If

	'Set the site name to be used in the computer name
	If strSite = "High School" Then
		strSitePrefix = "HS"
	ElseIf strSite = "Elementary" Then
		strSitePrefix = "ES"
	End If

	'Pad the asset tag with 
	Select Case Len(strAssetTag)
		Case 1
			strFixedAssetTag = "0000" & strAssetTag
		Case 2
			strFixedAssetTag = "000" & strAssetTag
		Case 3
			strFixedAssetTag = "00" & strAssetTag
		Case 4
			strFixedAssetTag = "0" & strAssetTag
		Case Else
			strFixedAssetTag = strAssetTag
	End Select

	'Get the graduation year from the tag list
	arrTags = Split(strTags,",")
	intGraduationYear = ""
	For Each strTag in arrTags
		If IsNumeric(Trim(strTag)) Then
			If Trim(strTag) > 2000 Then
				If Len(Trim(strTag)) = 4 Then
					intGraduationYear = Trim(strTag)
				End If
			End If
		End If
	Next

	'If the computer is assigned to a teacher than name it to the proper way
	If strRole = "Teacher" Then
		If strLastName <> "" Then
			GenerateComputerName = strSitePrefix & "-" & strLastName & "-" & strFixedAssetTag
		Else
			GenerateComputerName = strSitePrefix & "-Spare-" & strFixedAssetTag
		End If

	'If it's a student device name it the graduation year - asset tag.
	ElseIf strRole = "Student" Then
		
		'If the graduation year is found rename it the proper way, otherwise use the tag
		If intGRaduationYear <> "" Then
			GenerateComputerName = intGraduationYear & "-" & strFixedAssetTag
		Else
			GenerateComputerName = strFixedAssetTag
			SendEmailAboutMissingTag(strSerialNumber)
		End If
	
	'If the device isn't assigned to someone try naming it properly
	Else
		If intGRaduationYear <> "" Then
			GenerateComputerName = intGraduationYear & "-" & strFixedAssetTag
		ElseIf strRoom = "Library" Then
			GenerateComputerName = strSitePrefix & "-Lib-" & strFixedAssetTag
		Else
			GenerateComputerName = strSitePrefix & "-Spare-" & strFixedAssetTag
		End If
	End If
	

End Function 

Sub SendEmailAboutMissingTag(strSerialNumber)

	'This will send out an email if a student device is missing a graduation year tag.

	Dim strSMTPPickupFolder, objMessage, objConf, strMessage, strSubject, bolHTMLMEssage
	Dim bolSendAsAdmin, strURL

	strSQL = "SELECT LGTag,ExternalIP,InternalIP,ComputerName,LastUser,OSVersion FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
	Set objDevice = Application("Connection").Execute(strSQL)

	If Not objDevice.EOF Then

		Const cdoSendUsingPickup = 1

		strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"
	
		strURL = "http://" & Request.ServerVariables("server_name")
		strURL = strURL & Left(Request.ServerVariables("path_info"),Len(Request.ServerVariables("path_info")) - 7)
		strURL = strURL & "device.asp?Tag=" & objDevice(0)
	
		'Get the message body
		strMessage =  "The device with the tag of " & objDevice(0) & " is assigned to a student but doesn't have a graduation year tag.  "
		If objDevice(4) <> "" Then
			strMessage = strMessage & "The last user was " &  objDevice(4) & ".  "
		End If

		strMessage = strMessage & vbCRLF & vbCRLF & strURL
	
		strSubject = "Student Device Missing Tag - " & objDevice(0)

		'Create the objects required to send the mail.
		Set objMessage = CreateObject("CDO.Message")
		Set objConf = objMessage.Configuration
		With objConf.Fields
			.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
			.item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
			.Update
		End With
	
		objMessage.TextBody = strMessage
		objMessage.From = Application("EMailNotifications")
		objMessage.To = Application("LostDeviceNotify")
		objMessage.Subject = strSubject
		objMessage.Send
	
		'Close objects
		Set objMessage = Nothing
		Set objConf = Nothing
	End If
   
End Sub

Sub SendAlertInfectedEmail(strSerialNumber)

	'This will send out an email if the computer is infected

	Dim strSMTPPickupFolder, objMessage, objConf, strMessage, strSubject, bolHTMLMEssage
	Dim bolSendAsAdmin, strURL

	strSQL = "SELECT LGTag,ExternalIP,InternalIP,ComputerName,LastUser,OSVersion FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
	Set objDevice = Application("Connection").Execute(strSQL)

	If Not objDevice.EOF Then

		Const cdoSendUsingPickup = 1

		strSMTPPickupFolder = "C:\Inetpub\mailroot\Pickup"
	
		strURL = "http://" & Request.ServerVariables("server_name")
		strURL = strURL & Left(Request.ServerVariables("path_info"),Len(Request.ServerVariables("path_info")) - 7)
		strURL = strURL & "device.asp?Tag=" & objDevice(0)
	
		'Get the message body
		strMessage =  "The device with the tag of " & objDevice(0) & " may be infected with ransomware.  "
		If objDevice(4) <> "" Then
			strMessage = strMessage & "The last user was " &  objDevice(4) & ".  "
		End If

		strMessage = strMessage & vbCRLF & vbCRLF & strURL
	
		strSubject = "Possibly Infected Device - " & objDevice(0)

		'Create the objects required to send the mail.
		Set objMessage = CreateObject("CDO.Message")
		Set objConf = objMessage.Configuration
		With objConf.Fields
			.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
			.item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
			.Update
		End With
	
		objMessage.TextBody = strMessage
		objMessage.From = Application("EMailNotifications")
		objMessage.To = Application("LostDeviceNotify")
		objMessage.Subject = strSubject
		objMessage.Send
	
		'Close objects
		Set objMessage = Nothing
		Set objConf = Nothing
	End If
   
End Sub
%>