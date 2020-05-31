<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/11/18
'Last Updated 6/11/18

'This is the API for the inventory site

Option Explicit

'On Error Resume Next

Dim strLookupType, strUserName, strSQL, objLookUp, strSerialNumber, strLGTag, objDevice, strAssetTag, strOutput

strLookupType = Request.QueryString("Type")

Select Case strLookupType

	Case "DNSLookup"

		strSerialNumber = Request.QueryString("Serial")

		strSQL = "SELECT LGTag FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
		Set objLookUp = Application("Connection").Execute(strSQL)
		
		If Not objLookup.EOF Then
			strSQL = "SELECT ID FROM Tags WHERE Tag='Disable DNS' AND LGTag='" & objLookUp(0) & "'"
			Set objLookUp = Application("Connection").Execute(strSQL)

			If Not objLookup.EOF Then
				Response.Write("true")
			Else
				Response.Write("false")
			End If

		End If

	
	Case "StudentID"
	
		strUserName = Request.QueryString("UserName")
	
		strSQL = "SELECT StudentID FROM PEOPLE WHERE UserName='" & strUserName & "'"
		Set objLookUp = Application("Connection").Execute(strSQL)
		
		If Not objLookUp.EOF Then
			Response.Write(objLookUp(0))
		End If

	Case "ComputerName"
		strSerialNumber = Request.QueryString("Serial")

		strSQL = "SELECT ComputerName FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
		Set objLookUp = Application("Connection").Execute(strSQL)
		
		If Not objLookUp.EOF Then
			Response.Write(objLookUp(0))
		End If

	Case "AssetTag"
		strSerialNumber = Request.QueryString("Serial")

		strSQL = "SELECT LGTag FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
		Set objLookUp = Application("Connection").Execute(strSQL)
		
		If Not objLookup.EOF Then
			Response.Write(objLookUp(0))
		End If

	Case "Tags"
		strAssetTag = Request.QueryString("AssetTag")

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

		Response.Write(strOutput)
	
	Case "Backup"
		strSerialNumber = Request.QueryString("Serial")

		strSQL = "SELECT LGTag FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
		Set objLookUp = Application("Connection").Execute(strSQL)
		
		If Not objLookUp.EOF Then

			strLGTag = objLookUp(0)

			strSQL = "Select ID FROM Tags WHERE Tag='Backup' AND LGTag='" & strLGTag & "'"
			Set objLookUp = Application("Connection").Execute(strSQL)

			If Not objLookUp.EOF Then
				Response.Write("Backup")
			Else

				strSQL = "Select ID FROM Tags WHERE Tag='NoBackup' AND LGTag='" & strLGTag & "'"
				Set objLookUp = Application("Connection").Execute(strSQL)

				If Not objLookUp.EOF Then
					Response.Write("NoBackup")
				End If

			End If

		End If

	Case "AlertInfected"
		strSerialNumber = Request.QueryString("Serial")

		strSQL = "SELECT LGTag,ExternalIP,InternalIP,ComputerName,LastUser,OSVersion FROM Devices WHERE SerialNumber='" & strSerialNumber & "'"
		Set objDevice = Application("Connection").Execute(strSQL)

		If Not objDevice.EOF Then
			SendAlertInfectedEmail
		End If

End Select

Sub SendAlertInfectedEmail

	'This will send out an email

	Dim strSMTPPickupFolder, objMessage, objConf, strMessage, strSubject, bolHTMLMEssage
	Dim bolSendAsAdmin, strURL

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
   
End Sub
%>