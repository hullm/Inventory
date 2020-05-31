'Created by Matthew Hull on 10/29/19
'Last Updated 10/29/19

'This script will update the list of students in the database
'Last Name,First Name,Student ID,Sex,Birthday,Date Created,Homeroom Email,Homeroom,Username,Site,ClassOf

Option Explicit

'On Error Resume Next

Dim objDBConnection

CONST FIRSTNAME = 0
CONST LASTNAME = 1
CONST CLASSOF = 2
CONST SITE = 3
CONST HOMEROOM = 4
CONST HOMEROOMEMAIL = 5
CONST STUDENTID = 6
CONST DATECREATED = 7
CONST SEX = 8
CONST BIRTHDAY = 9
Const PASSWORD = 11
CONST USERNAME = 10

'Create a connection to the database
Set objDBConnection = ConnectToDatabase

CleanExportFile
ImportNewStudents

MsgBox "Done"

Sub ImportNewStudents

    Dim objFSO, strCurrentFolder, strCSV, strSQL, txtSourceCSV,arrUserData

    'This will import new students into the database and disable the ones who are gone

    'Get the CSV path
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strCurrentFolder = objFSO.GetAbsolutePathName(".")
    strCurrentFolder = strCurrentFolder & "\CSV\"
    strCSV = strCurrentFolder & "Import.csv"

    'Disable all the student accounts in the database, only the active ones will be
    'enabled later.
    strSQL = "UPDATE People SET Active=False WHERE ClassOf > 2000"
    objDBConnection.Execute(strSQL)

    'Open the source CSV
	Set txtSourceCSV = objFSO.OpenTextFile(strCSV)
	
	'Discard the header row
	txtSourceCSV.ReadLine
	
	'Loop through each line 
	While txtSourceCSV.AtEndOfLine = False
				
		'Get the data from the row and add it to an array
		arrUserData = GetUserDataFromImportedData(txtSourceCSV.ReadLine)
		
		'If they aren't in the database then add them
		If Not ExistsInDatabase(arrUserData(STUDENTID),arrUserData(CLASSOF)) Then

            'Add the user to the inventory database
			AddUserToDatabase arrUserData
		
			'Update the log
			UpdateLog "NewStudentDetected","",arrUserData(USERNAME),"",arrUserData(LASTNAME) & ", " & arrUserData(FIRSTNAME),""


        Else

            'The user already exists, but we will update a few values in the database, 
			'this is where the account is enabled in the database
			ModifyUserInDatabase arrUserData

        End If

    Wend

    'Close Objects
    Set objFSO = Nothing
    Set objDBConnection = Nothing

End Sub

Function GetUserDataFromImportedData(strImportedData)

	'This function will return an array that contains all the information about a user from the
	'import file.

	Dim arrRow, strLastName, strFirstName, intClassOf, intStudentID, strSex, datBirthday, datDateCreated
	Dim strHomeRoom, strHomeRoomEmail, arrLastName, strSite, strUserName
	
	'On Error Resume Next

	arrRow = Split(strImportedData,",")

	'Get the variables from the row
	strLastName= Trim(arrRow(0))
	strFirstName = Trim(arrRow(1))
	intStudentID = Trim(arrRow(2))
	strSex = Trim(arrRow(3))
	datBirthday = Trim(arrRow(4))
	datDateCreated = Trim(arrRow(5))

	'Fix the homeroom variables if there are two teachers
	If InStr(arrRow(6),"/") <> 0 Then
		strHomeRoomEmail = Replace(Trim(arrRow(6))," / ",";")
		strHomeRoom = Replace(Trim(arrRow(7)),"""","")
        strUserName = Trim(arrRow(10))
        strSite = Trim(arrRow(11))
        intClassOf = Right(Trim(arrRow(12)),4)
	Else

        If Trim(arrRow(6)) = "" Then
            strHomeRoomEmail = ""
            strHomeRoom = ""
            strUserName = Trim(arrRow(10))
            strSite = Trim(arrRow(11))
            intClassOf = Right(Trim(arrRow(12)),4)
        Else
            strHomeRoomEmail = Trim(arrRow(6))
            strHomeRoom = Replace(Trim(arrRow(7)) & ", " & Trim(arrRow(8)),"""","")
            strUserName = Trim(arrRow(11))
            strSite = Trim(arrRow(12))
            intClassOf = Right(Trim(arrRow(13)),4)
        End If

	End If

	'Fix the last name
	If InStr(strLastName," ") <> 0 Then
		arrLastName = Split(strLastName," ")
	
		'Fix the suffix
		Select Case LCase(arrLastName(1))
			Case "jr", "jr."
				strLastName = arrLastName(0) & " Jr"

			Case "ii", "2", "2nd"
				strLastName = arrLastName(0) & " II"
		
			Case "iii", "3", "3rd"
				strLastName = arrLastName(0) & " III"
		
			Case "iv", "4", "4th"
				strLastName = arrLastName(0) & " IV"
			
		End Select
	End If
	
	'Add the user's data to an array
	GetUserDataFromImportedData = Array(strFirstName,strLastName,intClassOf,strSite,strHomeRoom, _
	strHomeRoomEmail,intStudentID,datDateCreated,strSex,datBirthday,strUserName,"")

End Function

Sub ModifyUserInDatabase(arrUserData)

	'This will update the settings in the inventory database for the user

	Dim strSQL
	
	'Update the settings in the database
	strSQL = "UPDATE People SET "
	strSQL = strSQL & "HomeRoom='" & Replace(arrUserData(HOMEROOM),"'","''") & "'," 
	strSQL = strSQL & "HomeRoomEmail='" & Replace(arrUserData(HOMEROOMEMAIL),"'","''") & "'," 
	strSQL = strSQL & "Site='" & Replace(arrUserData(SITE),"'","''") & "'," 
	strSQL = strSQL & "Deleted=FALSE" & "," 
	strSQL = strSQL & "Active=TRUE" &  "," 
	strSQL = strSQL & "Sex='" & arrUserData(SEX) & "',"
	strSQL = strSQL & "Birthday=#" & arrUserData(BIRTHDAY) & "#,"
	strSQL = strSQL & "DateAdded=#" & arrUserData(DATECREATED) & "# " 
	strSQL = strSQL & "WHERE Role='Student' AND StudentID=" & arrUserData(STUDENTID)
	objDBConnection.Execute(strSQL)

End Sub

Sub CleanExportFile

	'This will remove unwanted character from an input file

	Dim objFSO, strCurrentFolder, strSourceCSV, strDestinationCSV, txtSourceCSV, txtDestinationCSV
	Dim strImportedData, intIndex, strCharacter

	'Get the paths to the CSV's
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCurrentFolder = objFSO.GetAbsolutePathName(".")
	strCurrentFolder = strCurrentFolder & "\CSV\"
	strSourceCSV = strCurrentFolder & "Export.csv"
	strDestinationCSV = strCurrentFolder & "Import.csv" 

	'Open the source CSV
	Set txtSourceCSV = objFSO.OpenTextFile(strSourceCSV)
	
	'Create the destination CSV
	Set txtDestinationCSV = objFSO.CreateTextFile(strDestinationCSV)
	
	'Replace the unwanted characters 
	While txtSourceCSV.AtEndOfLine = False
		strImportedData = txtSourceCSV.ReadLine
        strImportedData = Replace(strImportedData,", Jr"," Jr")
        strImportedData = Replace(strImportedData,", III"," III")
        strImportedData = Replace(strImportedData,", IV"," IV")
		strImportedData = Replace(strImportedData,"´","")
		strImportedData = Replace(strImportedData,"  "," ")
		strImportedData = Replace(strImportedData,"!","")
        strImportedData = Replace(strImportedData,"""","")
		strImportedData = Replace(strImportedData,"#","")
		strImportedData = Replace(strImportedData,"$","")
		strImportedData = Replace(strImportedData,"%","")
		strImportedData = Replace(strImportedData,"&","")
		strImportedData = Replace(strImportedData,"(","")
		strImportedData = Replace(strImportedData,")","")
		strImportedData = Replace(strImportedData,"*","")
		strImportedData = Replace(strImportedData,"+","")
		strImportedData = Replace(strImportedData,"`","")
		strImportedData = Replace(strImportedData," 12:00:00 AM","")
		
		'Write the fixed data to the new file
		txtDestinationCSV.Write(strImportedData & vbCRLF)
	
	Wend
	
	'Close the CSV files
	txtSourceCSV.Close
	txtDestinationCSV.Close

	'Close objects
	Set objFSO = Nothing
	Set txtSourceCSV = Nothing
	Set txtDestinationCSV = Nothing

End Sub 

Function ExistsInDatabase(intStudentID,intClassOf)

	'This will check and see if a user exists in the database, it will also correct a students 
	'class off setting if their grade has changed.  The function will return a true or false.

    'On Error Resume Next

	Dim strSQL, objStudent

	'Check and see if they are in the students table
	strSQL = "SELECT ID,ClassOf" & vbCRLF
	strSQL = strSQL & "FROM People" & vbCRLF
	strSQL = strSQL & "WHERE ClassOf>2000 AND StudentID=" & intStudentID
	Set objStudent = objDBConnection.Execute(strSQL)

	'If they aren't in the people table then they aren't in the database.
	If objStudent.EOF Then
		ExistsInDatabase = False
	Else
		ExistsInDatabase = True
	End If

End Function

Sub AddUserToDatabase(arrUserData)

	'This will add the new user to the database with a password of NewAccount.  The
	'user will be flagged as deleted in the database.

    On Error Resume Next

	Dim strSQL

	strSQL = "INSERT INTO People (FirstName,LastName,Username,Role,ClassOf,Site,HomeRoom,HomeRoomEmail,Sex,Birthday,PWord,StudentID,Active,Pending,AUP,Deleted,DateAdded)" & vbCRLF
	strSQL = strSQL & "VALUES ('"
	strSQL = strSQL & Replace(arrUserData(FIRSTNAME),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(LASTNAME),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(USERNAME),"'","''") & "','"
	strSQL = strSQL & "Student','"
	strSQL = strSQL & arrUserData(CLASSOF) & "','"
	strSQL = strSQL & Replace(arrUserData(SITE),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(HOMEROOM),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(HOMEROOMEMAIL),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(SEX),"'","''") & "','"
	strSQL = strSQL & Replace(arrUserData(BIRTHDAY),"'","''") & "','"
	strSQL = strSQL & "',"
	strSQL = strSQL & arrUserData(STUDENTID) & ",True,False,"
	strSQL = strSQL & "False,False,#" & arrUserData(DATECREATED) & "#)"
    objDBConnection.Execute(strSQL)

    If Err Then
        InputBox strSQL, strSQL, strSQL
        Err.Clear
    End If
	
End Sub

Function ConnectToDatabase

	'This function returns a connection object used to run SQL commands against the database

	Dim objFSO, strCurrentFolder, strDatabase, strConnection

	'Get the database path
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCurrentFolder = objFSO.GetAbsolutePathName(".")
	strCurrentFolder = objFSO.GetParentFolderName(strCurrentFolder)
	strCurrentFolder = strCurrentFolder & "\Database\"
	strDatabase = strCurrentFolder & "Inventory.mdb"

	'Create the connection to the database
	Set ConnectToDatabase = CreateObject("ADODB.Connection")
	strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strDatabase & ";"
	ConnectToDatabase.Open strConnection
	
	Set objFSO = Nothing

End Function

Sub UpdateLog(EntryType,DeviceTag,UserName,OldValue,NewValue,EventNumber)

	'This will update the log with what change has been made

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
	
	'Zero out the event number if nothing's there
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
	
	'Write the log entry to the database
	strSQL = "INSERT INTO Log (LGTag,UserName,Type,OldValue,NewValue,OldNotes,NewNotes,UpdatedBy,LogDate,LogTime,Active,Deleted,EventNumber)" & vbCRLF
	strSQL = strSQL & "VALUES ('"
	strSQL = strSQL & intTag & "','"
	strSQL = strSQL & Replace(strUserName,"'","''") & "','"
	strSQL = strSQL & Replace(strType,"'","''") & "','"
	strSQL = strSQL & Replace(strOldValue,"'","''") & "','"
	strSQL = strSQL & Replace(strNewValue,"'","''") & "','"
	strSQL = strSQL & Replace(strOldNotes,"'","''") & "','"
	strSQL = strSQL & Replace(strNewNotes,"'","''") & "','"
	strSQL = strSQL & Replace("Automated","'","''") & "',#"
	strSQL = strSQL & datDate & "#,#"
	strSQL = strSQL & datTime & "#,True,False," & intEventNumber & ")"
	objDBConnection.Execute(strSQL)
	
End Sub