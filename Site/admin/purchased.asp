<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 6/11/21
'Last Updated 6/11/21

'This page will show who has purchased retired MacBook Airs

'Option Explicit

On Error Resume Next

Dim strSiteVersion, bolShowLogout, strUser, strColumns, strSQL, objPurchased, strSubmitTo, strAddMessage, intAddErrID
Dim objCompleted, intPurchaseCount, intCompletedCount

'See if the user has the rights to visit this page
If AccessGranted Then
	ProcessSubmissions
Else
	DenyAccess
End If %>

<%Sub ProcessSubmissions

	intAddErrID = 0
	strAddMessage = ""

	'Check and see if anything was submitted to the site
	Select Case Request.Form("Submit")
		Case "Mark as Picked Up"
			MarkAsPickedUp
		Case "Add"
			AddAssetTag
		Case "Delete"
			DeleteAssetTag
		Case "Remove"
			RemoveEntry
	End Select
	
	'Get the URL used to submit forms
	If Request.ServerVariables("QUERY_STRING") = "" Then
		strSubmitTo = "purchased.asp"
	Else
		strSubmitTo = "purchased.asp?" & Request.ServerVariables("QUERY_STRING")
	End If

	'Grab the list of purchased 
	strSQL = "SELECT Owed.ID,RecordedDate,FirstName,LastName,Site,ClassOf,LGTag,UserName,Model,DeviceAge,Owed.Active" & vbCRLF
	strSQL = strSQL & "FROM Owed INNER JOIN People ON Owed.OwedBy = People.ID" & vbCRLF
	strSQL = strSQL & "WHERE Owed.Item LIKE '%Retired%' AND PickupDate IS NULL" & vbCRLF
	strSQL = strSQL & "ORDER BY LastName, FirstName;"
	Set objPurchased = Application("Connection").Execute(strSQL)

	'Grab the list of completed entries
	strSQL = "SELECT Owed.ID,RecordedDate,FirstName,LastName,Site,ClassOf,LGTag,PickupDate,Price,UserName,Model,DeviceAge" & vbCRLF
	strSQL = strSQL & "FROM Owed INNER JOIN People ON Owed.OwedBy = People.ID" & vbCRLF
	strSQL = strSQL & "WHERE Owed.Active=False AND Owed.Item LIKE '%Retired%' AND PickupDate IS NOT NULL" & vbCRLF
	strSQL = strSQL & "ORDER BY LastName, FirstName;"
	Set objCompleted = Application("Connection").Execute(strSQL)

	'Count the number of purchased devices
	intPurchaseCount = 0
	If Not objPurchased.EOF Then
		Do Until objPurchased.EOF
			intPurchaseCount = intPurchaseCount + 1
			objPurchased.MoveNext
		Loop
		objPurchased.MoveFirst
	End If

	'Count the number of completed devices
	intCompletedCount = 0
	If Not objCompleted.EOF Then
		Do Until objCompleted.EOF
			intCompletedCount = intCompletedCount + 1
			objCompleted.MoveNext
		Loop
		objCompleted.MoveFirst
	End If

	SetupSite
	DisplaySite

End Sub %>

<%Sub DisplaySite %>

	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	<html>
	<head>
		<title><%=Application("SiteName")%></title>
		<link rel="stylesheet" type="text/css" href="../style.css" />
		<link rel="apple-touch-icon" href="../images/inventory.png" />
		<link rel="shortcut icon" href="../images/inventory.ico" />
		<meta name="viewport" content="width=device-width,user-scalable=0" />
		<meta name="theme-color" content="#333333">
		<link rel="stylesheet" href="../assets/css/jquery-ui.css">
		<script src="../assets/js/jquery.js"></script>
		<script src="../assets/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="../assets/css/jquery.dataTables.min.css">
		<link rel="stylesheet" href="../assets/css/buttons.dataTables.min.css">
		<script src="../assets/js/jquery.dataTables.min.js"></script>
		<script src="../assets/js/dataTables.buttons.min.js"></script>
		<script src="../assets/js/buttons.colVis.min.js"></script>
		<script src="../assets/js/buttons.html5.min.js"></script>
		<script>
			$(function() {

			<%	If Not IsMobile Then %>
					$( document ).tooltip({track: true});
			<%	End If %>

			var purchasedTable = $('#Purchased').DataTable( {
					paging: false,
					"info": false,
					"autoWidth": false,
					dom: 'Bfrtip',
					"order": [],
					// stateSave: true,
					buttons: [
						{
							extend: 'colvis',
							text: 'Show/Hide Columns'
						}
				<%	If Not IsMobile Then %>		
						,
						{
							extend: 'csvHtml5',
							text: 'Download CSV',
							title: 'Purchased Devices to be Picked Up'
						}
				<%	End If %>
					]
				});
		<%	If IsMobile Then %>
				purchasedTable.columns([0,2,3,4,5]).visible(false);	
		<%	End If %>
				var completedTable = $('#Completed').DataTable( {
					paging: false,
					"info": false,
					"autoWidth": false,
					dom: 'Bfrtip',
					"order": [],
					// stateSave: true,
					buttons: [
						{
							extend: 'colvis',
							text: 'Show/Hide Columns'
						}
				<%	If Not IsMobile Then %>		
						,
						{
							extend: 'csvHtml5',
							text: 'Download CSV',
							title: 'Purchased Devices'
						}
				<%	End If %>
					]
				});
		<%	If IsMobile Then %>
				completedTable.columns([0,1,5,6]).visible(false);	
		<%	End If %>
			} );
		</script>
	</head>
	<body class="<%=strSiteVersion%>">

		<div class="Header"><%=Application("SiteName")%></div>
		<div>
			<ul class="NavBar" align="center">
				<li><a href="index.asp"><img src="../images/home.png" title="Home" height="32" width="32"/></a></li>
				<li><a href="search.asp"><img src="../images/search.png" title="Search" height="32" width="32"/></a></li>
				<li><a href="stats.asp"><img src="../images/stats.png" title="Stats" height="32" width="32"/></a></li>
				<li><a href="log.asp"><img src="../images/log.png" title="System Log" height="32" width="32"/></a></li>
				<li><a href="add.asp"><img src="../images/add.png" title="Add Person or Device" height="32" width="32"/></a></li>
				<li><a href="login.asp?action=logout"><img src="../images/logout.png" title="Log Out" height="32" width="32"/></a></li>
			</ul>
		</div>

	<%
		DisplayPurchasedTable
		DisplayCompletedTable
	%>

		<div class="Version">Version <%=Application("Version")%></div>
		<div class="CopyRight"><%=Application("Copyright")%></div>
	</body>

	</html>

<%End Sub%>

<%Sub DisplayPurchasedTable

	Dim objAssignmentInfo, strAssignmentInfo, strRowClass, strShowPickupButton%>

	<div>
		<br />
		<a href="log.asp">System Log</a> | Purchase Log (<%=intPurchaseCount%>)
		<table align="center" Class="ListView" id="Purchased">
			<thead>
				<th>Date</th>
				<th>Name</th>
				<th>Site</th>
				<th>Role</th>
				<th>Model</th>
				<th>Device Year</th>
				<th>Asset Tag</th>
				<th>Action</th>
			</thead>		
			<tbody>
		<%	Do Until objPurchased.EOF 
		
				strAssignementInfo = ""
				If IsNull(objPurchased(6)) Then
					'Get information about assignments
					strSQL = "SELECT DISTINCT Assignments.LGTag,Model,DatePurchased,Assignments.Active,Devices.Active" & vbCRLF
					strSQL = strSQL & "FROM People INNER JOIN (Devices INNER JOIN Assignments ON Devices.LGTag = Assignments.LGTag) On People.ID = Assignments.AssignedTo" & vbCRLF
					strSQL = strSQL & "WHERE People.UserName='" & objPurchased(7) & "'"
					Set objAssignmentInfo = Application("Connection").Execute(strSQL)
					strAssignementInfo = ""

					'Build the popup message
					If NOT objAssignmentInfo.EOF Then
						strAssignementInfo = "title="""
						Do Until objAssignmentInfo.EOF
							strAssignementInfo = strAssignementInfo & objAssignmentInfo(0) & " - "
							strAssignementInfo = strAssignementInfo & objAssignmentInfo(1) & " - Year "
							strAssignementInfo = strAssignementInfo & GetAge(objAssignmentInfo(2))
							If objAssignmentInfo(3) Then
								strAssignementInfo = strAssignementInfo & " - Active"
							End If
							strAssignementInfo = strAssignementInfo & " &#013 "
							objAssignmentInfo.MoveNext
						Loop
						strAssignementInfo = strAssignementInfo & """"
					End If
				End If

				'Highlight the row if the user hasn't paid for the device
				If objPurchased(10) Then
					strRowClass = " Class=""Warning"""
				Else
					strRowClass = ""
				End If

				'Disable the pickup button if a device isn't assigned to them, or of they haven't paid
				If objPurchased(6) = "" Or IsNull(objPurchased(6)) Then
					strShowPickupButton = "disabled"
				ElseIf objPurchased(10) Then
					strShowPickupButton = "disabled"
				Else 
					strShowPickupButton = ""
				End If
				%>

				<tr <%=strRowClass%>>
					<td><%=objPurchased(1)%></td>
					<td <%=strAssignementInfo%>><a href="user.asp?UserName=<%=objPurchased(7)%>&Back=Purchased&Page=purchased.asp"><%=objPurchased(3)%>,&nbsp;<%=objPurchased(2)%></a></td>
					<td id="center"><%=objPurchased(4)%></td>
					<td id="center"><%=GetRole(objPurchased(5))%></td>
					<td id="center"><%=objPurchased(8)%></td>
			<%	If IsNull(objPurchased(9)) Then %>
					<td id="center">&nbsp;</td>
			<%	Else %>
					<td id="center">Year&nbsp;<%=objPurchased(9)%></td>
			<%	End If %>
					<td id="center">
					<%	If objPurchased(6) = "" Or IsNull(objPurchased(6)) Then %>
							<form method="POST" action="<%=strSubmitTo%>">
								<input type="hidden" value="<%=objPurchased(0)%>" name="id" />
									<input class="Card InputWidthSmall" type="input" name="AssetTag" />
									<input type="submit" value="Add" name="Submit" />
							<%	If Int(intAddErrID) = Int(objPurchased(0)) Then %>
									<div class="Error"><%=strAddMessage%></div>
							<%	End If %>
							</form>
					<%	Else %>
							<form method="POST" action="<%=strSubmitTo%>">
								<a href="device.asp?Tag=<%=objPurchased(6)%>&Back=Purchased&Page=purchased.asp"><%=objPurchased(6)%></a>&nbsp;
								<input type="hidden" value="<%=objPurchased(0)%>" name="id" />
								<input type="hidden" value="<%=objPurchased(6)%>" name="AssetTag" />
								<input type="submit" value="Delete" name="Submit" />
							</form>
					<% End If %>
					</td>
					<td id="center">
					
						<form method="POST" action="<%=strSubmitTo%>">
							<input type="hidden" value="<%=objPurchased(0)%>" name="id" />
							<input type="hidden" value="<%=objPurchased(6)%>" name="AssetTag" />
							<input type="submit" value="Mark as Picked Up" name="Submit" <%=strShowPickupButton%>>
							<input type="submit" value="Remove" name="Submit" />
						</form>
					</td>
				</tr>
		<%		objPurchased.MoveNext 
			Loop %>

			</tbody>
		</table>
	</div>

<%End Sub %>

<%Sub DisplayCompletedTable

	Dim intSumOfPurchases 
	
	intSumOfPurchases = 0 %>

	<div>
		<br />
		<br />
		<br />
		History (<%=intCompletedCount%>)
		<table align="center" Class="ListView" id="Completed">
			<thead>
				<th>Date Purchased</th>
				<th>Date Picked Up</th>
				<th>Name</th>
				<th>Model</th>
				<th>Device Year</th>
				<th>Asset Tag</th>
				<th>Purchase Price</th>
			</thead>		
			<tbody>
		<%	Do Until objCompleted.EOF 
				intSumOfPurchases = intSumOfPurchases + CLng(objCompleted(8)) %>
				<tr>
					<td><%=objCompleted(1)%></td>
					<td><%=objCompleted(7)%></td>
					<td><a href="user.asp?UserName=<%=objCompleted(9)%>&Back=Purchased&Page=purchased.asp"><%=objCompleted(3)%>,&nbsp;<%=objCompleted(2)%></a></td>
					<td id="center"><%=objCompleted(10)%></td>
			<%	If IsNull(objCompleted(11)) Then %>
					<td id="center">&nbsp;</td>
			<%	Else %>
					<td id="center">Year&nbsp;<%=objCompleted(11)%></td>
			<%	End If %>		
					<td id="center"><a href="device.asp?Tag=<%=objCompleted(6)%>&Back=Purchased&Page=purchased.asp"><%=objCompleted(6)%></a></td>
					<td id="center">$<%=objCompleted(8)%></td>
				</tr>
		<%		objCompleted.MoveNext 
			Loop %>
		<%	If intSumOfPurchases <> 0 Then %>
				<tr>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td id="center">Total: $<%=intSumOfPurchases%></td>
				</tr>
		<% End If %>

			</tbody>
		</table>
	</div>

<%End Sub %>

<%Sub AddAssetTag

	Dim strSQL, intID, intTag, objUser, objDeviceCheck, strModel, intDeviceAge

	intID = Request.Form("id")
	intTag = Request.Form("AssetTag")
	bolAddTag = True

	'Check to see if the device is already purchased to someone
	strSQL = "SELECT ID FROM Owed WHERE LGTag='" & Replace(intTag,"'","''") & "'"
	Set objDeviceCheck = Application("Connection").Execute(strSQL)
	If NOT objDeviceCheck.EOF Then
		strAddMessage = "Already Purchased"
		intAddErrID = intID
		bolAddTag = False
		Exit Sub
	End If

	'Check to see if the device exists in the database and if it's assigned
	strSQL = "SELECT ID,Assigned,Model,DatePurchased FROM Devices WHERE LGTag='" & Replace(intTag,"'","''") & "'"
	Set objDeviceCheck = Application("Connection").Execute(strSQL)
	If objDeviceCheck.EOF Then
		strAddMessage = "Device Doesn't Exist"
		intAddErrID = intID
		bolAddTag = False
		Exit Sub
	ElseIf objDeviceCheck(1) Then
		strAddMessage = "Device Still Assigned <a href=""device.asp?Tag=" & intTag & "&Back=Purchased&Page=purchased.asp"">" & intTag & "</a>"
		intAddErrID = intID
		bolAddTag = False
		Exit Sub
	End If

	strModel = objDeviceCheck(2)
	intDeviceAge = GetAge(objDeviceCheck(3))

	If bolAddTag Then

		'Add the device to the purchased list
		strSQL = "UPDATE Owed SET "
		strSQL = strSQL & "LGTag='" & Replace(intTag,"'","''") & "', "
		strSQL = strSQL & "Model='" & Replace(strModel,"'","''") & "', "
		strSQL = strSQL & "DeviceAge=" & intDeviceAge & vbCRLF
		strSQL = strSQL & "WHERE ID=" & intID
		Application("Connection").Execute(strSQL)

		'Add the Purchased tag
		strSQL = "INSERT INTO Tags (LGTAG,TAG)" & vbCRLF
		strSQL = strSQL & "VALUES ('" & Replace(intTag,"'","''") & "','Purchased')"
		Application("Connection").Execute(strSQL)

		'Disable the device
		strSQL = "UPDATE Devices SET "
		strSQL = strSQL & "Active=False, "
		strSQL = strSQL & "DateDisabled=Date() "
		strSQL = strSQL & "WHERE LGTag='" & Replace(intTag,"'","''") & "'"
		Application("Connection").Execute(strSQL)

		'Get the username so we can update the log
		strSQL = "SELECT UserName" & vbCRLF
		strSQL = strSQL & "FROM Owed INNER JOIN People ON Owed.OwedBy = People.ID" & vbCRLF
		strSQL = strSQL & "WHERE Owed.ID=" & intID
		Set objUser = Application("Connection").Execute(strSQL)

		'Update the log
		UpdateLog "DeviceDisabled",intTag,"","Enabled","Disabled",""
		UpdateLog "DeviceUpdatedTagAdded",intTag,objUser(0),"","Purchased",""
		UpdateLog "PurchasedByUser",intTag,objUser(0),"","",""

	End If

End Sub%>

<%Sub DeleteAssetTag

	Dim strSQL, intID, intTag, objUser

	intID = Request.Form("id")
	intTag = Request.Form("AssetTag")

	'Remove the Purchased tag from the device
	strSQL = "DELETE FROM Tags WHERE Tag='Purchased' AND LGTag='" & Replace(intTag,"'","''") & "'"
	Application("Connection").Execute(strSQL)

	'Remove the data from the device
	strSQL = "UPDATE Owed SET LGTag=NULL, Model=Null, DeviceAge=NULL WHERE ID=" & intID
	Application("Connection").Execute(strSQL)

	'Get the username so we can update the log
	strSQL = "SELECT UserName" & vbCRLF
	strSQL = strSQL & "FROM Owed INNER JOIN People ON Owed.OwedBy = People.ID" & vbCRLF
	strSQL = strSQL & "WHERE Owed.ID=" & intID
	Set objUser = Application("Connection").Execute(strSQL)

	UpdateLog "DeviceUpdatedTagDeleted",intTag,objUser(0),"Purchased","",""
	UpdateLog "PurchasedByUserDeleted",intTag,objUser(0),"","",""

End Sub%>

<%Sub RemoveEntry

	Dim strSQL, intID, intTag, objUser

	intID = Request.Form("id")
	intTag = Request.Form("AssetTag")

	'Delete the asset tag first, this will help make sure the data is logged
	If intTag <> "" Then
		DeleteAssetTag
	End If

	'Remove the entry from the database
	strSQL = "DELETE FROM Owed WHERE ID=" & intID
	Application("Connection").Execute(strSQL)

End Sub%>

<%Sub MarkAsPickedUp

	Dim strSQL, intID, datDate, objUser

	intID = Request.Form("id")
	datDate = Date()

	'Get the username and asset tag so we can update the log
	strSQL = "SELECT UserName,LGTag" & vbCRLF
	strSQL = strSQL & "FROM Owed INNER JOIN People ON Owed.OwedBy = People.ID" & vbCRLF
	strSQL = strSQL & "WHERE Owed.ID=" & intID
	Set objUser = Application("Connection").Execute(strSQL)

	strSQL = "UPDATE Owed SET PickUpdate=#" & datDate & "# WHERE ID=" & intID
	Application("Connection").Execute(strSQL)

	UpdateLog "PurchasedDevicePickedUp",objUser(1),objUser(0),"","",""

End Sub%>

<%Sub UpdateLog(EntryType,DeviceTag,UserName,OldValue,NewValue,EventNumber)

	Dim strType, strOldNotes, strNewNotes, strOldValue, strNewValue, intTag, strUserName, datDate, datTime, strSQL, intEventNumber

	'Get the type
	If EntryType <> "" Then
		strType = EntryType
	Else 
		Exit Sub
	End If
	
	'If a notes field was updated then the data needs to be stored in the notes field of the log
	If InStr(strType,"Notes") > 0 Then
		strOldNotes = OldValue
		strNewNotes = NewValue
		strOldValue = ""
		strNewValue = ""
	Else
		strOldNotes = ""
		strNewNotes = ""
		strOldValue = OldValue
		strNewValue = NewValue
	End If
	
	If EventNumber = "" Then
		intEventNumber = 0
	Else
		intEventNumber = EventNumber
	End If
	
	'Get the other things needed for the log
	intTag = DeviceTag
	strUserName = UserName
	datDate = Date()
	datTime = Time()
	
	strSQL = "INSERT INTO Log (LGTag,UserName,Type,OldValue,NewValue,OldNotes,NewNotes,UpdatedBy,LogDate,LogTime,Active,Deleted,EventNumber)" & vbCRLF
	strSQL = strSQL & "VALUES ('"
	strSQL = strSQL & intTag & "','"
	strSQL = strSQL & Replace(strUserName,"'","''") & "','"
	strSQL = strSQL & Replace(strType,"'","''") & "','"
	strSQL = strSQL & Replace(strOldValue,"'","''") & "','"
	strSQL = strSQL & Replace(strNewValue,"'","''") & "','"
	strSQL = strSQL & Replace(strOldNotes,"'","''") & "','"
	strSQL = strSQL & Replace(strNewNotes,"'","''") & "','"
	strSQL = strSQL & Replace(strUser,"'","''") & "',#"
	strSQL = strSQL & datDate & "#,#"
	strSQL = strSQL & datTime & "#,True,False," & intEventNumber & ")"
	Application("Connection").Execute(strSQL)
	
End Sub%>

<%Function GetAge(strDate)

	Dim strMonth, strDay, strYear, intIndex, datStartofYear, datEndofYear

	strMonth = Month(Date)
	strDay = Day(Date)
	strYear = Year(Date)

	For intIndex = 0 to 100

		datStartofYear = strMonth & "/" & strDay & "/" & strYear - intIndex - 1
		datEndofYear = strMonth & "/" & strDay & "/" & strYear - intIndex
		If CDate(strDate) >= CDate(datStartofYear) And CDate(strDate) <= CDate(datEndofYear)  Then
			GetAge = intIndex + 1
			Exit For
		End If

	Next

End Function %>

<%Function GetRole(intYear)

	Dim datToday, intMonth, intCurrentYear, intGrade, strSQL, objRole
	
	'If they're an adult then get their role from the database
	If intYear <= 1000 Then
		strSQL = "SELECT Role FROM Roles WHERE RoleID=" & intYear
		Set objRole = Application("Connection").Execute(strSQL)
		
		If Not objRole.EOF Then
			GetRole = objRole(0)
		End If
	End If 
		
	'Convert the graduating year to a grade
	datToday = Date
	intMonth = DatePart("m",datToday)
	intCurrentYear = Right(DatePart("yyyy",datToday),2)
	intYear = Right(intYear,2)

	If intMonth >= 7 And intMonth <= 12 Then
		intCurrentYear = intCurrentYear + 1
	End If

	intGrade = 12 - (intYear - intCurrentYear)

	If GetRole = "" Then
	
		'Change the grade number into text
		Select Case intGrade
			Case -1
				GetRole = "PreSchool Student"
			Case 0
				GetRole = "Kindergarten Student"
			Case 1
				GetRole = "1st Grade Student"
			Case 2
				GetRole = "2nd Grade Student"
			Case 3
				GetRole = "3rd Grade Student"
			Case 4
				GetRole = "4th Grade Student"
			Case 5
				GetRole = "5th Grade Student"
			Case 6
				GetRole = "6th Grade Student"
			Case 7
				GetRole = "7th Grade Student"
			Case 8
				GetRole = "8th Grade Student"
			Case 9
				GetRole = "9th Grade Student"
			Case 10
				GetRole = "10th Grade Student"
			Case 11
				GetRole = "11th Grade Student"
			Case 12
				GetRole = "12th Grade Student"
			Case Else
				GetRole = "Graduated"
		End Select

	End If

End Function%>

<%
' Anything below here should exist on all pages
%>

<%Sub DenyAccess

	'If we're not using basic authentication then send them to the login screen
	If bolShowLogout Then
		Response.Redirect("login.asp?action=logout")
	Else

	SetupSite

	%>
	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	<html>
	<head>
		<title><%=Application("SiteName")%></title>
		<link rel="stylesheet" type="text/css" href="../style.css" />
		<link rel="apple-touch-icon" href="../images/inventory.png" />
		<link rel="shortcut icon" href="../images/inventory.ico" />
		<meta name="viewport" content="width=device-width" />
	</head>
	<body>
		<center><b>Access Denied</b></center>
	</body>
	</html>

<%End If

End Sub%>

<%Function AccessGranted

	Dim objNetwork, strUserAgent, strSQL, strRole, objNameCheckSet

	'Redirect the user the SSL version if required
	If Application("ForceSSL") Then
		If Request.ServerVariables("SERVER_PORT")=80 Then
			If Request.ServerVariables("QUERY_STRING") = "" Then
				Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
			Else
				Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
			End If
		End If
	End If

	'Get the users logon name
	Set objNetwork = CreateObject("WSCRIPT.Network")
	strUser = objNetwork.UserName
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'Check and see if anonymous access is enabled
	If LCase(Left(strUser,4)) = "iusr" Then
		strUser = GetUser
		bolShowLogout = True
	Else
		bolShowLogout = False
	End If

	'Build the SQL string, this will check the userlevel of the user.
	strSQL = "Select Role" & vbCRLF
	strSQL = strSQL & "From Sessions" & vbCRLF
	strSQL = strSQL & "WHERE UserName='" & strUser & "' And SessionID='" & Request.Cookies("SessionID") & "'"
	Set objNameCheckSet = Application("Connection").Execute(strSQL)
	strRole = objNameCheckSet(0)

	If strRole = "Admin" Then
		AccessGranted = True
	Else
		AccessGranted = False
	End If

End Function%>

<%Function GetUser

	Const USERNAME = 1

	Dim strUserAgent, strSessionID, objSessionLookup, strSQL

	'Get some needed data
	strSessionID = Request.Cookies("SessionID")
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'Send them to the logon screen if they don't have a Session ID
	If strSessionID = "" Then
		SendToLogonScreen

	'Get the username from the database
	Else

		strSQL = "SELECT ID,UserName,SessionID,IPAddress,UserAgent,ExpirationDate FROM Sessions "
		strSQL = strSQL & "WHERE UserAgent='" & Left(Replace(strUserAgent,"'","''"),250) & "' And SessionID='" & Replace(strSessionID,"'","''") & "'"
		strSQL = strSQL & " And ExpirationDate > Date()"
		Set objSessionLookup = Application("Connection").Execute(strSQL)

		'If a session isn't found for then kick them out
		If objSessionLookup.EOF Then
			SendToLogonScreen
		Else
			GetUser = objSessionLookup(USERNAME)
		End If
	End If

End Function%>

<%Function IsMobile

	Dim strUserAgent

	'Get the User Agent from the client so we know what browser they are using
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'Check the user agent for signs they are on a mobile browser
	If InStr(strUserAgent,"iPhone") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"iPad") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"Android") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"Windows Phone") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"BlackBerry") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"Nintendo") Then
		IsMobile = True
	ElseIf InStr(strUserAgent,"PlayStation Vita") Then
		IsMobile = True
	Else
		IsMobile = False
	End If

	If InStr(strUserAgent,"Nexus 9") Then
		IsMobile = False
	End If
End Function %>

<%Sub SendToLogonScreen

	Dim strReturnLink, strSourcePage

	'Build the return link before sending them away.
	strReturnLink =  "?" & Request.ServerVariables("QUERY_STRING")
	strSourcePage = Request.ServerVariables("SCRIPT_NAME")
	strSourcePage = Right(strSourcePage,Len(strSourcePage) - InStrRev(strSourcePage,"/"))
	If strReturnLink = "?" Then
		strReturnLink = "?SourcePage=" & strSourcePage
	Else
		strReturnLink = strReturnLink & "&SourcePage=" & strSourcePage
	End If

	Response.Redirect("login.asp" & strReturnLink)

End Sub %>

<%Sub SetupSite

	If IsMobile Then
		strSiteVersion = "Mobile"
	Else
		strSiteVersion = "Full"
	End If

	If Application("MultiColumn") Then
		  strColumns = "MultiColumn"
	  End If

End Sub%>