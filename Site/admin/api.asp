<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/11/18
'Last Updated 6/11/18

'This is the API for the inventory site

Option Explicit

'On Error Resume Next

Dim strLookupType, strUserName, strSQL, objLookUp, strSerialNumber, strLGTag

strLookupType = Request.QueryString("Type")

Select Case strLookupType

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

End Select
%>