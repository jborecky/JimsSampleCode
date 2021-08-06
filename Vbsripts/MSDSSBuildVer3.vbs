'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME:MSDSSBuildVer3.vbs
'
' AUTHOR: Jim Borecky , 
' DATE  : 8/21-24/2006
'
' COMMENT: Script that takes the file created from GetUserFromServerPermissions.vbs
'			and creates a simulated MSDSS file to import into aelita
'
'
'
'			IMPORTANT!!!!!!!
'				 Need special DLL's to use this object
'				 novell-activex_core-devel-2004.10.06-1.exe
'				 novell-activex_ndap-devel-2006.06.14-1windows.zip
'				 both can be found on the Novell Web site
'
' JB Version 3 will not create any users but will create a MSDSS file For
'	NOAM and RETAIL domains
'
'
'==========================================================================

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8


if wscript.arguments.count = 1 then
	strInputFile = wscript.arguments.item(0)
Else
	ShowHelp
	wscript.quit
end If

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objGroupNDSTank = CreateObject("Scripting.Dictionary")
Set objUserNDSTank = CreateObject("Scripting.Dictionary")
Set objContainerNDSTank = CreateObject("Scripting.Dictionary")
Set objRoleNDSTank = CreateObject("Scripting.Dictionary")
Set objDic = CreateObject("Scripting.Dictionary")		'NDS Extraction Dictionary Object
Set objUserDic = CreateObject("Scripting.Dictionary")	'Dictionary for User NovellDn to LDAPDn user mapping 
Set objGroupDic = CreateObject("Scripting.Dictionary")	'Dictionary for Group NovellDn to LDAPDn group mappings
Set objFileDic =CreateObject("Scripting.Dictionary")	'Dictionsry object to load Imporrt File into
Set objUserNotFoundDic = CreateObject("Scripting.Dictionary") 'Dictionary object to get Users not found in NDS
Set objUserDuplicateDic = CreateObject("Scripting.Dictionary") 'Dictionary object to get Duplicate users list
Set objGroupDupDic = CreateObject("Scripting.Dictionary") 'Dictionary objects for Duplicate groups
Set objGroupOUDic = CreateObject("Scripting.Dictionary") 'Dictionary objects for Container groups
Set objRoleDic = CreateObject("Scripting.Dictionary") 'Dictionary ojbects for Role groups

'----------------------------------------------------------------------------

'Opening File and retrieving attributes
Set objInputFile = objfso.OpenTextFile(strInputFile,ForReading)
Do While objInputFile.AtEndOfStream <> True
	strFileText = trim(objInputFile.readline)
	ObjectSort(strFileText)	
loop
objInputFile.Close

'Add any additional Users from the groups
For Each group In objGroupNDSTank
	AddUsersFromGroupToTank(Group)
Next

'Add any Addition Users from the Containers
For Each Container In objContainerNDSTank
	AddUsersFromContainerToTank(Container)
Next

'Add any additional users from Org Roles
For Each Role In objRoleNDSTank
	AddUsersFromRoleToTank(Role)
Next

'Get Unique OU
strSession = GenerateFileName
strSessionGet = Split(strSession,vbTab)(0)
arrSession = Split(strSessionGet,"\")
strSessionGet = arrSession(ubound(arrSession))
strSessionGet = Split(strSessionGet,".")(0)
strUniqueOU = "OU=" & strSessionGet

'Destination Migration Path --Too be replaced later by variable--
dnMainOU = "OU=FUN-Tree,OU=Migrated Users,dc=migp,dc=my,dc=Bank,dc=net"

dnDestinationOU = strUniqueOU & "," & dnMainOU

'Create OU to GoTo
Call CreateOU(dnMainOU,strUniqueOU)

'temporary enumeration of users
'Set objOutputfile = objFSO.CreateTextFile("found.txt",ForWriting)
For Each user In objUserNDSTank
	'objOutputfile.WriteLine user
	WScript.Echo user
	Call CreateUserFromNDS(user,dnDestinationOU)	
Next

For Each Container In objContainerNDSTank
	'Create Group
	Call CreateContainerFromNDS(Container,dnDestinationOU)
	'Populate Group
	Call PopulateContainerFromNDSName(Container)
Next

For Each Group In objGroupNDSTank
	'Create Group
	Call CreateGroupFromNDS(Group,dnDestinationOU)
	'Populate Group
	Call PopulateGroupFromNDSName(Group)
Next

For Each Role In objRoleNDSTank
	'Create Group
	Call CreateRoleFromNDS(Role,dnDestinationOU)
	'Populate Group
	Call PopulateRoleFromNDSName(Role)
Next

'objOutputfile.Close
	'Sort Dictionary
	Call SortDictionary(objDic)

'Creating an Extra File for MSDSS while I'm here in case Microsoft fixes it.
Set objFoundFile = objfso.CreateTextFile("C:\jimWorks\Migration-" & strSessionGet & ".txt",True,True)
For Each item In objDic.Keys
	objFoundFile.WriteLine item
Next
objFoundFile.close

'Creating an extra File for Duplicate Names
Set objFoundFile = objfso.CreateTextFile("C:\jimWorks\Duplicate-" & strSessionGet & ".txt",True,True)
For Each item In objUserDuplicateDic.Keys
	objFoundFile.WriteLine item
Next
objFoundFile.close

'Creating an extra File for Missing Users
Set objFoundFile = objfso.CreateTextFile("C:\jimWorks\Missing-" & strSessionGet & ".txt",True,True)
For Each item In objUserNotFoundDic.Keys
	objFoundFile.WriteLine item
Next
objFoundFile.close

'Creating an extra File for Group Mappings
Set objFoundFile = objfso.CreateTextFile("C:\jimWorks\Groups-" & strSessionGet & ".xls",True)
For Each key In objGroupDupDic.Keys
	strItem = objGroupDupDic.Item(key)
	objFoundFile.WriteLine key & vbTab & strItem
Next
objFoundFile.close

'Building the Session file for Aelita
Call BuildSessionFile(strSession)
'-----------------------------------------------------------------------------
Function PopulateRoleFromNDSName(strNDSGroup)
	strNDSName = "NDS:\\FUN\" & strNDSGroup
	WScript.Echo "strNDSGroup:" & strNDSGroup
	Set NWDirQuery1 = CreateObject("NWDirLib.NWDirCtrl.1")
	Set results = NWDirQuery1.FindEntry(strNDSName)
	GetObjectClass = "Unknown"
	Fields = results.GetFieldValue("Role Occupant",Empty)
	
	'Get group DN
	strGroupDN = objRoleNDSTank.Item(strNDSGroup)
	
	If IsArray(Fields) then
		For k = LBound(Fields) To UBound(Fields)
			strNDSdn = objUserNDSTank.item(Fields(k))
			Call AddToGroup(strGroupDN, strNDSdn)
				WScript.Echo Fields(k)
		Next
	End If
End Function
'-----------------------------------------------------------------------------
Function PopulateGroupFromNDSName(strNDSGroup)
	strNDSName = "NDS:\\FUN\" & strNDSGroup
	WScript.Echo "strNDSGroup:" & strNDSGroup
	Set NWDirQuery1 = CreateObject("NWDirLib.NWDirCtrl.1")
	Set results = NWDirQuery1.FindEntry(strNDSName)
	GetObjectClass = "Unknown"
	Fields = results.GetFieldValue("member",Empty)
	
	'Get group DN
	strGroupDN = objGroupNDSTank.Item(strNDSGroup)
	
	If IsArray(Fields) then
		For k = LBound(Fields) To UBound(Fields)
			strNDSdn = objUserNDSTank.item(Fields(k))
			Call AddToGroup(strGroupDN, strNDSdn)
				WScript.Echo Fields(k)
		Next
	End If
End Function
'-----------------------------------------------------------------------------
Function CreateRoleFromNDS(strNDSDN,dnDestOU)
	'This was inserted because the periods in NDS was causing the object to puke we don't
	'need them so we'll just skip em..
	If instr(strNDSDN ,"Role Based Service")=0 and instr(strNDSDN ,"Tomcat-Roles")=0 then
		'Getting the Format we need. See Functions for details
		strUserDN = RemoveNDS(strNDSDN)
		strUserDN = NewFormat(strUserDN,dnDestinationOU)
		strSamAccount = Split(strUserDN,"CN=")(1) & "#"
		'Checking to see if Group Exists
		If (groupExists(strSamAccount) <> True) And (objRoleDic.Exists(strUserDN) = false) Then
			'None were to be found. Let's make it so!!!
			Call CreateGroup(dnDestOU, strSamAccount, strUserDN)
			'Getting Dn for Dictionary
			strGroupDN = GetGroupDN(strSamAccount)
			'Adding to Dictionary so we can make our Session File
			objRoleDic.Add strUserDN,strGroupDN
		'Ok here we go because we have duplicate groups in the blooming NDS tree
		ElseIf (groupExists(strSamAccount) = True) And (objRoleDic.Exists(strUserDN) = false) Then
			'Let's find us an acceptable group name
			DupCount = 1
			'Change so we can search better
			'strDuplicateGroupName = "Dup_" & DupCount & "_" & strSamAccount
			strDuplicateGroupName = strSamAccount & "_Dup_" & DupCount
			Do While groupExists(strDuplicateGroupName)
				DupCount = DupCount + 1
				'Changed so we can search better
				'strDuplicateGroupName = "Dup_" & DupCount & "_" & strSamAccount
				strDuplicateGroupName = strSamAccount & "_Dup_" & DupCount
			Loop
			'Making it so!!
			strSamAccount = strDuplicateGroupName
			'Create that group
			Call CreateGroup(dnDestOU, strSamAccount,strUserDN)
			'Get the Dn for the Dictionary
			strGroupDN = GetGroupDN(strSamAccount)
			'Add Group to Dictionary for the Session File
			objRoleDic.Add strUserDN,strGroupDN
		ElseIf (groupExists(strSamAccount) = True) And (objRoleDic.Exists(strUserDN) <> 0) Then
			'Get DN From Dict
			strGroupDN = objRoleDic.Item(strUserDN)						
		End If
		
		'Add to Container Dict
		If objRoleDic.Exists(strUserDN) = False then
			objRoleDic.Add strUserDN,strGroupDN
		End If
		
		'Change Dict object
		objRoleNDSTank.Item(strNDSDN) = strGroupDN
	End If
	
	'Populate Group
	
End Function
'-----------------------------------------------------------------------------
Function PopulateContainerFromNDSName(strNDSGroup)
	Dim results
	strNDSName = "NDS:\\FUN\" & strNDSGroup
	WScript.Echo strNDSName

	Set NWDirQuery1 = CreateObject("NWDirQueryLib.NWDirQuery")
	NWDirQuery1.FullName = strNDSName
	NWDirQuery1.Filter = "User(CN = *)"
	NWDirQuery1.Fields = "CN, Login Disabled, Given Name, Surname, Telephone Number, Group Membership"
	NWDirQuery1.SearchScope = 2
	NWDirQuery1.SearchMode = 0
	NWDirQuery1.MaximumResults = 100
	Set results = NWDirQuery1.Search
	
	' The number of accounts returned by the NDS Query
	WScript.Echo results.Count
	
	'Get group DN
	strGroupDN = objContainerNDSTank.Item(strNDSGroup)
	
	'Insert missing users
	For Each user In results
		strNDSdn = objUserNDSTank.item(user.fullname)
		Call AddToGroup(strGroupDN, strNDSdn)
				WScript.Echo user.Fullname
		
	Next
End Function
'-----------------------------------------------------------------------------
Function CreateGroupFromNDS(strNDSDN,dnDestOU)
	'This was inserted because the periods in NDS was causing the object to puke we don't
	'need them so we'll just skip em..
	If instr(strNDSDN ,"Role Based Service")=0 and instr(strNDSDN ,"Tomcat-Roles")=0 then
		'Getting the Format we need. See Functions for details
		strUserDN = RemoveNDS(strNDSDN)
		strUserDN = NewFormat(strUserDN,dnDestinationOU)
		strSamAccount = Split(strUserDN,"CN=")(1)
		'Checking to see if Group Exists
		If (groupExists(strSamAccount) <> True) And (objGroupDic.Exists(strUserDN) = false) Then
			'None were to be found. Let's make it so!!!
			Call CreateGroup(dnDestOU, strSamAccount, strUserDN)
			'Getting Dn for Dictionary
			strGroupDN = GetGroupDN(strSamAccount)
			'Adding to Dictionary so we can make our Session File
			objGroupDic.Add strUserDN,strGroupDN
		'Ok here we go because we have duplicate groups in the blooming NDS tree
		ElseIf (groupExists(strSamAccount) = True) And (objGroupDic.Exists(strUserDN) = false) Then
			'Let's find us an acceptable group name
			DupCount = 1
			'Change so we can search better
			'strDuplicateGroupName = "Dup_" & DupCount & "_" & strSamAccount
			strDuplicateGroupName = strSamAccount & "_Dup_" & DupCount
			Do While groupExists(strDuplicateGroupName)
				DupCount = DupCount + 1
				'Changed so we can search better
				'strDuplicateGroupName = "Dup_" & DupCount & "_" & strSamAccount
				strDuplicateGroupName = strSamAccount & "_Dup_" & DupCount
			Loop
			'Making it so!!
			strSamAccount = strDuplicateGroupName
			'Create that group
			Call CreateGroup(dnDestOU, strSamAccount,strUserDN)
			'Get the Dn for the Dictionary
			strGroupDN = GetGroupDN(strSamAccount)
			'Add Group to Dictionary for the Session File
			objGroupDic.Add strUserDN,strGroupDN
		ElseIf (groupExists(strSamAccount) = True) And (objGroupDic.Exists(strUserDN) <> 0) Then
			'Get DN From Dict
			strGroupDN = objGroupDic.Item(strUserDN)						
		End If
		
		'Change Dict object
		objGroupNDSTank.Item(strNDSDN) = strGroupDN
	End If
	
	'Populate Group
	
End Function
'-----------------------------------------------------------------------------
Function CreateContainerFromNDS(strNDSDN,dnDestOU)
	'This was inserted because the periods in NDS was causing the object to puke we don't
	'need them so we'll just skip em..
	If instr(strNDSDN ,"Role Based Service")=0 and instr(strNDSDN ,"Tomcat-Roles")=0 then
		'Getting the Format we need. See Functions for details
		strUserDN = RemoveNDS(strNDSDN)
		strUserDN = NewFormatOU(strUserDN,dnDestinationOU)
		arrAccount = Split(strNDSDN,"\")
		strSamAccount = arrAccount(ubound(arrAccount)) & "$"
		'Checking to see if Group Exists
		WScript.Echo "samaccount:" & strUserDN
		If (groupExists(strSamAccount) <> True) And (objGroupOUDic.Exists(strUserDN) = false) Then
			'None were to be found. Let's make it so!!!
			Call CreateGroup(dnDestOU, strSamAccount, strUserDN)
			'Getting Dn for Dictionary
			strGroupDN = GetGroupDN(strSamAccount)
			'Adding to Dictionary so we can make our Session File
			objGroupOUDic.Add strUserDN,strGroupDN
		'Ok here we go because we have duplicate groups in the blooming NDS tree
		ElseIf (groupExists(strSamAccount) = True) And (objGroupOUDic.Exists(strUserDN) = false) Then
			'Let's find us an acceptable group name
			DupCount = 1
			'Change so we can search better
			'strDuplicateGroupName = "Dup_" & DupCount & "_" & strSamAccount
			strDuplicateGroupName = strSamAccount & "_Dup_" & DupCount
			Do While groupExists(strDuplicateGroupName)
				DupCount = DupCount + 1
				'Changed so we can search better
				'strDuplicateGroupName = "Dup_" & DupCount & "_" & strSamAccount
				strDuplicateGroupName = strSamAccount & "_Dup_" & DupCount
			Loop
			'Making it so!!
			strSamAccount = strDuplicateGroupName
			'Create that group
			Call CreateGroup(dnDestOU, strSamAccount,strUserDN)
			'Get the Dn for the Dictionary
			strGroupDN = GetGroupDN(strSamAccount)
			'Add Group to Dictionary for the Session File
			objGroupOUDic.Add strUserDN,strGroupDN
		ElseIf (groupExists(strSamAccount) = True) And (objGroupOUDic.Exists(strUserDN) <> 0) Then
			'Get DN From Dict
			strGroupDN = objGroupOUDic.Item(strUserDN)						
		End If
		
		'Add to Container Dict
		If objGroupOUDic.Exists(strUserDN) = False then
			objGroupOUDic.Add strUserDN,strGroupDN
		End If
		
		'Change Dict object
		objContainerNDSTank.Item(strNDSDN) = strGroupDN
	End If
	
	'Populate Group
	
End function
'-----------------------------------------------------------------------------
Function AddUsersFromContainerToTank(strContainer)
	Dim results

	Set NWDirQuery1 = CreateObject("NWDirQueryLib.NWDirQuery")
	NWDirQuery1.FullName = strContainer
	NWDirQuery1.Filter = "User(CN = *)"
	NWDirQuery1.Fields = "CN, Login Disabled, Given Name, Surname, Telephone Number, Group Membership"
	NWDirQuery1.SearchScope = 2
	NWDirQuery1.SearchMode = 0
	NWDirQuery1.MaximumResults = 100
	Set results = NWDirQuery1.Search
	
	' The number of accounts returned by the NDS Query
	WScript.Echo results.Count
	
	'Insert missing user into Dictionary
	For Each user In results
		If objUserNDSTank.Exists(user.fullname) = False Then
				objUserNDSTank.Add user.Fullname,"user"
			End If 
	Next
End Function
'-----------------------------------------------------------------------------
Function AddUsersFromRoleToTank(strObjectNDS)
	Set NWDirQuery1 = CreateObject("NWDirLib.NWDirCtrl.1")
	Set results = NWDirQuery1.FindEntry(strObjectNDS)
	GetObjectClass = "Unknown"
	Fields = results.GetFieldValue("Role Occupant",Empty)
	If IsArray(Fields) then
		For k = LBound(Fields) To UBound(Fields)
			If objUserNDSTank.Exists(Fields(k)) = False Then
				objUserNDSTank.Add Fields(k),"user"
			End If 
		Next
	End If
	Set NWDirQuery1 = Nothing
End Function
'-----------------------------------------------------------------------------
Function AddUsersFromGroupToTank(strObjectNDS)
	Set NWDirQuery1 = CreateObject("NWDirLib.NWDirCtrl.1")
	Set results = NWDirQuery1.FindEntry(strObjectNDS)
	GetObjectClass = "Unknown"
	Fields = results.GetFieldValue("member",Empty)
	If IsArray(Fields) then
		For k = LBound(Fields) To UBound(Fields)
			If objUserNDSTank.Exists(Fields(k)) = False Then
				objUserNDSTank.Add Fields(k),"user"
			End If 
		Next
	End If
	Set NWDirQuery1 = Nothing
End Function
'-----------------------------------------------------------------------------
Function ObjectSort(strObject)
	arrObject = Split(strObject,vbTab)
	strObjectName = arrObject(0)
	strType = arrObject(1)
	Select Case strType
		Case "User"
			If objUserNDSTank.Exists(strObjectName)= false Then
				objUserNDSTank.Add strObjectname,strType
			End If
		Case "Group"
			If objGroupNDSTank.Exists(strObjectName) = False Then
				objGroupNDSTank.Add strObjectName,strType
			End If
		Case "Organizational Unit"
			If objContainerNDSTank.Exists(strObjectName) = False Then
				objContainerNDSTank.Add strObjectName, strType
			End If
		Case "Organizational Role"
			If objRoleNDSTank.Exists(strObjectName) = False Then
				objRoleNDSTank.Add strObjectName,strType
			End If
	End Select
End Function 
'-----------------------------------------------------------------------------
sub ShowHelp()
	Wscript.echo
	Wscript.echo "You need to specify an a text file to use"
	Wscript.echo "cscript ScriptName.vbs InputFileName"
	wscript.echo
	Wscript.echo "Example:"
	Wscript.echo "cscript ScriptName.vbs Input.txt"
	wscript.echo 
	
End Sub

'-----------------------------------------------------------------------------
'			Function to create a Child OU
'-----------------------------------------------------------------------------
Function CreateOU(dnParentOU,dnChildOU)
	Set objOU1 = GetObject("LDAP://" & dnParentOU)
	Set objOU2 = objOU1.Create("organizationalUnit", dnChildOU)	'Adding OU
	objOU2.SetInfo
End Function
'-----------------------------------------------------------------------------
'			Function to create a Group In
'				Active Dierctory
'-----------------------------------------------------------------------------
Function CreateGroup(strDestinationOU, strGroupName, strNovellDN)
	'Putting in pause because we are killing the DC's
	WScript.Sleep 1000
	Const ADS_GROUP_TYPE_GLOBAL_GROUP = &h2
	Const ADS_GROUP_TYPE_SECURITY_ENABLED = &h80000000
	
	WScript.Echo strDestinationOU
	WScript.Echo strGroupName
	Set objOU = GetObject("LDAP://MYDomainController001/" & strDestinationOU)
	Set objGroup = objOU.Create("Group", "CN=" & strGroupName)	'Adding Group
	
	On Error Resume Next
	Err.Clear
	objGroup.Put "sAMAccountName", strGroupName
	'Populating the Notes: field with the old NDS context
	objGroup.Put "info", strNovellDN
	objGroup.Put "groupType", ADS_GROUP_TYPE_GLOBAL_GROUP Or _
	    ADS_GROUP_TYPE_SECURITY_ENABLED							'Making all groups Global groups
	objGroup.SetInfo
	If Err.Number <> 0 Then
		WScript.echo Err.Description
	End If
	On Error GoTo 0
	
	'If InStr(strGroupname,"$") Then
	'	If not objGroupOUDic.Exists(strDictUserDn) Then
	'		objGroupOUDic.Add strNovellDN, "CN=" & strGroupName & "," & strDestinationOU
	'	End if
	'else
		'Adding Group to the Dictionary so we can print it out later
		If not objGroupDupDic.Exists(strNovellDN) then
			objGroupDupDic.Add strNovellDN, strGroupName
		End If
	'End if
End Function 
'-----------------------------------------------------------------------------
'			Function to Add a User to a Group
'-----------------------------------------------------------------------------
Function AddToGroup(strGroupNameDN, strUserToAddDN)
	Const ADS_PROPERTY_APPEND = 3 
	
	On Error Resume Next
	'Err.Clear
	'Do While Err.Number <> 0 
		Set objGroup = GetObject("LDAP://MYDomainController001/" & strGroupNameDN) 
 			objGroup.PutEx ADS_PROPERTY_APPEND, "member", _
	    Array(strUserToAddDN)	 							'Adding user Here
	
		objGroup.SetInfo
		If Err.Number <> 0 Then
			WScript.echo Err.Description
		End If
	'Loop
	On Error GoTo 0
End Function
'-----------------------------------------------------------------------------
Function CreateUserFromNDS(strNDSName,dnDestOU)

	'Getting needed info from NDS
	On Error Resume Next
	Err.Clear
	Set NWDirQuery1 = CreateObject("NWDirLib.NWDirCtrl.1")
	Set results = NWDirQuery1.FindEntry(strNDSName)
	arrName = results.GetFieldValue("CN",Empty,true)
	strName = arrName(0)
	arrFirstName = results.GetFieldValue("Given Name",Empty,true)
	strFirstName = arrFirstName(0)
	arrAcctStatus = results.GetFieldValue("Login Disabled",Empty,true)
	strAcctStatus = arrAcctStatus(0)
	arrLastName = results.GetFieldValue("Surname",Empty,true)
	strLastName = arrLastName(0)
	arrDescription = results.GetFieldValue("Description",Empty,true)
	strDescription = replace(arrDescription(0),vbTab,"")
	If Err.Number = 0 then
		WScript.Echo strName & vbTab & strAcctStatus & vbTab & strFirstName & vbTab &_
					strLastName & vbTab & strDescription
		
		'See Function for descriptions
		'Getting what we need which is NDS DN for aelita and Sam account name
		strUserDN = RemoveNDS(results.Fullname)
		strUserDN = NewFormat(strUserDN,dnDestinationOU)
		strSamAccount = Split(strUserDN,"CN=")(1)	
			
		'Checking to see if we are creating Duplicate accounts
			If UserExists(strSamAccount) <> True Then
				'We're ok to create new account
				Call CreateUser(dnDestOU,strSamAccount,strFirstName,strLastName,strUserDN,strAcctStatus,strDescription)
				'Getting Dn to add to Dictionary
				dnUserAccount = GetUserDN(strSamAccount)
				'Adding to Dictionary to use when we build Session File
				objUserDic.Add strUserDN,dnUserAccount
			Else
				'Opps have a duplicate account in NDS. Someone didn't do their job
				'That's ok we'll process it anyways. Can't slow down the migration Team
				
				'Adding duplicate user to Dictionary
				If objUserDuplicateDic.Exists(strSamAccount) = 0 then
					objUserDuplicateDic.Add strSamAccount,strSamAccount
				End If 
				
				'Looking to find a acceptable duplicate name
				DupCount = 1
				'Changed so we can search better
				'strDuplicateUser = "Dup_" & DupCount & "_" & strSamAccount
				strDuplicateUser = strSamAccount & "_Dup_" & DupCount
				Do While UserExists(strDuplicateUser)
					DupCount = DupCount + 1
					'Changed so we can search better
					'strDuplicateUser = "Dup_" & DupCount & "_" & strSamAccount
					strDuplicateUser = strSamAccount & "_Dup_" & DupCount
				Loop
				
				'Prompting user to create the additional user.
				Answer = MsgBox("Duplicate User Found: " & strSamAccount & vbNewLine & _
					"Should I create a Duplicate account?: " & strDuplicateUser ,vbYesNoCancel)
				Select Case Answer
					'Operator Selected Cancel We be done here!!!!
					Case 2
						WScript.Echo "Process terminated by User"
						WScript.Quit
					'Selected Yes we will create the duplicate with a different name
					Case 6
						'Replacing variable and continuing as before
						strSamAccount = strDuplicateUser
						'Creating user with alternate Name
						Call CreateUser(dnDestOU,strSamAccount,strFirstName,strLastName,strUserDN,strAcctStatus,strDescription)
						'Getting Dn to add to Dictionary
						dnUserAccount = GetUserDN(strSamAccount)
						'Adding to Dictionary so we can create the Session file later.
						objUserDic.Add strUserDN,dnUserAccount
					'Selected No Continue and basically merge the users here
					Case 7
						'Informing user that they will basically get merged
						MsgBox("System will use current User")
						'Getting Dn to add to Dictionary
						dnUserAccount = GetUserDN(strSamAccount)
						'Adding to Dictionary so we can create the Session file later.
						objUserDic.Add strUserDN,strSamAccount
				End Select				
			End If
			'Getting the LDAP DN of the user we just created
			strUserLDAPDN = GetUserDN(strSamAccount)
			
			'Put Accociate NDS:// with NDSDN
			objUserNDSTank.Item(strNDSName) = strUserLDAPDN
		End If
	
End Function
'-----------------------------------------------------------------------------
'			Function to Create a User in AD
'-----------------------------------------------------------------------------
Function CreateUser(strDestinationOU,strUserName,strFirst,strLast,strUserDN,strStatus,strDescription)
	'Need to create timing loop here
	Set objOU = GetObject("LDAP://MYDomainController001/" & strDestinationOU)

	Set objUser = objOU.Create("User", "CN=" & strUserName)
	objUser.Put "sAMAccountName", strUserName			'Populating Sam Account Name
	If strFirst <> "" then
		objUser.Put "givenName" , strFirst					'Populating Given Name
	End If
	If strLast <> "" then
		objUser.Put "sn", strLast							'Populating Sirname
	End If
	If strFirst <> "" And strLast <>"" Then
		objUser.Put "displayName", strFirst & " " & strLast	'Populating Display Name
	End If
	If strUserDN <> "" Then
		objUser.Put "info",strUserDN
	End If
	If strDescription <> "" Then
		objUser.Put "description",strDescription
	End If
	objUser.SetInfo
	'Only enable is account is enabled in NDS
	'WScript.Echo "Status =" & strStatus
	If strStatus = false then
		objUser.AccountDisabled = False						'Enabling Account
		objUser.SetInfo
	End if
End Function 
'-----------------------------------------------------------------------------
'			Function to Find out if the User already Exists
'-----------------------------------------------------------------------------
Function UserExists(strUserName)
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
 	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
 
	objCommand.CommandText = _
	    "<LDAP://MYDomainController001/dc=migp,dc=my,dc=bank,dc=net>;(&(objectCategory=User)" & _
	         "(samAccountName=" & strUserName & "));distinguishedName,samAccountName;subtree"
  
	'Setting up for a connection timeout
	ConnectionErr = True
	ConnectionTimeoutCount = 0
  	Do While ConnectionErr 
  		On Error Resume Next
  		Err.Clear
  		'Making the Connection here
		Set objRecordSet = objCommand.Execute
		'If the connection fails wait
		If Err.Number = 0 Then
			ConnectionErr = False
		Else
			'Didn't get a connection
			WScript.Sleep 3000
			WScript.Echo Err.Description
			ConnectionTimeoutCount = ConnectionTimeoutCount + 1
			If ConnectionTimeoutCount > 5 Then
				WScript.Echo "LDAP Query is Down"
				WScript.Quit
			End If
		End If
		On Error GoTo 0
	Loop
 
	If objRecordset.RecordCount = 0 Then	
	    UserExists = false
	Else	
		UserExists = True
	End If
	set objConnection = Nothing
	Set objCommand = nothing
End Function
'-----------------------------------------------------------------------------
'			Function to get the User DN from AD
'				Returns dishtingusihed name
'-----------------------------------------------------------------------------
Function GetUserDN(strShortUserName)
	dim strDistinguishedName
	' Getting ADO objects
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
 
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
 
 	'Setting Query String
	objCommand.CommandText = _
	    "<LDAP://MYDomainController001/dc=migp,dc=my,dc=bank,dc=net>;(&(objectCategory=User)" & _
	         "(samAccountName=" & strShortUserName & "));distinguishedName,samAccountName;subtree"
	
	'Setting up for a connection timeout
	ConnectionErr = True
	ConnectionTimeoutCount = 0
  	Do While ConnectionErr 
  		On Error Resume Next
  		Err.Clear
  		'Making the Connection here
		Set objRecordSet = objCommand.Execute
		'If the connection fails wait
		If Err.Number = 0 Then
			ConnectionErr = False
		Else
			'Didn't get a connection
			WScript.Sleep 3000
			WScript.Echo Err.Description
			ConnectionTimeoutCount = ConnectionTimeoutCount + 1
			If ConnectionTimeoutCount > 5 Then
				WScript.Echo "LDAP Query is Down"
				WScript.Quit
			End If
		End If
		On Error GoTo 0
	Loop
	    strDistinguishedName = objRecordset.Fields("distinguishedName")
	    GetUserDN = strDistinguishedName
	set objConnection = Nothing
	Set objCommand = Nothing
End Function
'-----------------------------------------------------------------------------
'			Function to get the Group DN from AD
'				Returns dishtingusihed name
'-----------------------------------------------------------------------------
Function GetGroupDN(strUserName)
	dim strDistinguishedName
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
 
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
 
	objCommand.CommandText = _
	    "<LDAP://MYDomainController001/dc=migp,dc=my,dc=bank,dc=net>;(&(objectCategory=Group)" & _
	         "(samAccountName=" & strUserName & "));distinguishedName,samAccountName;subtree"
  
	'Setting up for a connection timeout
	ConnectionErr = True
	ConnectionTimeoutCount = 0
  	Do While ConnectionErr 
  		On Error Resume Next
  		Err.Clear
  		'Making the Connection here
		Set objRecordSet = objCommand.Execute
		'If the connection fails wait
		If Err.Number = 0 Then
			ConnectionErr = False
		Else
			'Didn't get a connection
			WScript.Sleep 3000
			WScript.Echo Err.Description
			ConnectionTimeoutCount = ConnectionTimeoutCount + 1
			If ConnectionTimeoutCount > 5 Then
				WScript.Echo "LDAP Query is Down"
				WScript.Quit
			End If
		End If
		On Error GoTo 0
	Loop
	    strDistinguishedName = objRecordset.Fields("distinguishedName")
	    GetGroupDN = strDistinguishedName
	Set objConnection = Nothing
	Set objCommand = nothing
End Function
'-----------------------------------------------------------------------------
'			Function to check to see if a group exists in AD
'				Returns bool value
'-----------------------------------------------------------------------------
Function groupExists(strGroupName)
	groupExists = false
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
 	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
 
	objCommand.CommandText = _
	    "<LDAP://MYDomainController001/dc=migp,dc=my,dc=bank,dc=net>;(&(objectCategory=Group)" & _
	         "(samAccountName=" & strGroupName & "));distinguishedName,samAccountName;subtree"
  
	'Setting up for a connection timeout
	ConnectionErr = True
	ConnectionTimeoutCount = 0
  	Do While ConnectionErr 
  		On Error Resume Next
  		Err.Clear
  		'Making the Connection here
		Set objRecordSet = objCommand.Execute
		'If the connection fails wait
		If Err.Number = 0 Then
			ConnectionErr = False
		Else
			'Didn't get a connection
			WScript.Sleep 3000
			WScript.Echo Err.Description
			ConnectionTimeoutCount = ConnectionTimeoutCount + 1
			If ConnectionTimeoutCount > 5 Then
				WScript.Echo "LDAP Query is Down"
				WScript.Quit
			End If
		End If
		On Error GoTo 0
	Loop
 	If objRecordset.RecordCount = 0 Then	
	    GroupExists = false
	Else	
		GroupExists = True
	End If
	set objConnection = Nothing
	Set objCommand = Nothing
End Function
'-----------------------------------------------------------------------------
'			Function to remove the NDS prefix from the Fullname
'-----------------------------------------------------------------------------
Function RemoveNDS(strUser)
	Set NWDir = CreateObject("NWDirLib.NWDirCtrl.1")
	If InStr(strUser,"NDS:") Then
		'getting the tree name in case this is used elsewhere
		strTree = NWDir.TreeFromFullName(strUser)
		'Some magic and instant presto, you have a string ready to reformat
		strUser = Split(strUser,"\\")(1)
		strUser = Split(strUser,strTree&"\")(1)
	End If
	RemoveNDS = strUser
End Function
'-----------------------------------------------------------------------------
'			Function to Reformat to a imitation NDS DN
'-----------------------------------------------------------------------------
Function NewFormat(DN,strDestOU)
	'split the different parts
	arrDN = Split(Dn,"\")
	'put back together
	For CountDNSpace = 0 To UBound(arrDN)
		If CountDNSpace = 0 Then
			strNovellDN = "O=" & arrDN(CountDNSpace)	'The beginning name
			strNewGroupName = arrDN(CountDNSpace)
		ElseIf CountDNSpace = UBound(arrDN) Then
			strNovellDN = strNovellDN & "/" & "CN=" & arrDN(CountDNSpace) 'the End of the name
		Else
			strNovellDN = strNovellDN & "/"& "OU=" & arrDN(CountDNSpace) 'Everything in between a OU I guess
			strNewGroupName = strNewGroupName & "_" & arrDN(COuntDNSpace)
		End If
		' Add it too the Dictionary to output for migration.txt
		'if not objDic.Exists(strNovellDN)then
		'	AddToDict(strNovellDN)
		'	If InStr(strNovellDN,"CN=") < 1 Then
		'		If objGroupOUDic.Exists(strNovellDN) = 0 Then
		'			'Need to check for Duplicate Container Groups
		'			DupCount = 1
		'			strGroupCheck = strNewGroupName
		'			Do while groupExists(strGroupCheck & "_NDS$")
		'				strGroupCheck = strNewGroupName & "_DUP" & DupCount
		'				DupCount = DupCount + 1
		'			loop
		'			Call CreateGroup(strDestOU,strGroupCheck & "_NDS$",strNovellDN)
		'		End if
		'	End if
		'End if
		
	Next
	NewFormat = strNovellDN
End Function
'------------------------------------------------------------------------------------
Function NewFormatOU(DN,strDestOU)
	'split the different parts
	arrDN = Split(Dn,"\")
	'put back together
	For CountDNSpace = 0 To UBound(arrDN)
		If CountDNSpace = 0 Then
			strNovellDN = "O=" & arrDN(CountDNSpace)	'The beginning name
			strNewGroupName = arrDN(CountDNSpace)
		ElseIf CountDNSpace = UBound(arrDN) Then
			strNovellDN = strNovellDN & "/" & "OU=" & arrDN(CountDNSpace) 'the End of the name
		Else
			strNovellDN = strNovellDN & "/"& "OU=" & arrDN(CountDNSpace) 'Everything in between a OU I guess
			strNewGroupName = strNewGroupName & "_" & arrDN(COuntDNSpace)
		End If
		' Add it too the Dictionary to output for migration.txt
		'if not objDic.Exists(strNovellDN)then
		'	AddToDict(strNovellDN)
		'	If InStr(strNovellDN,"CN=") < 1 Then
		'		If objGroupOUDic.Exists(strNovellDN) = 0 Then
		'			'Need to check for Duplicate Container Groups
		'			DupCount = 1
		'			strGroupCheck = strNewGroupName
		'			Do while groupExists(strGroupCheck & "_NDS$")
		'				strGroupCheck = strNewGroupName & "_DUP" & DupCount
		'				DupCount = DupCount + 1
		'			loop
		'			Call CreateGroup(strDestOU,strGroupCheck & "_NDS$",strNovellDN)
		'		End if
		'	End if
		'End if
		
	Next
	NewFormatOU = strNovellDN
End Function

'-----------------------------------------------------------------------------
'			Function to add a DN to te Dictionary for migration.txt
'-----------------------------------------------------------------------------
Function AddToDict(DN)
	If not objDic.Exists(DN) Then
		objDic.Add DN,DN
	End if
End Function
'-----------------------------------------------------------------------------
' 			Function Returns a new GUID in a string or an empty string.
'			By: Ken Treadway
'-----------------------------------------------------------------------------
Function MakeGUID
	On Error Resume Next
	
	' Local variables.
	Dim objTL, strGUID
	
	' Be negative.
	MakeGUID = ""
	
	' Create the type library.
	Err.Clear
	Set objTL = CreateObject("Scriptlet.TypeLib")
	If (Err.Number <> 0) Then
		WScript.StdErr.WriteLine("Unable to create Scriptlet.TypeLib.")
		Exit Function
	End If
	
	' Get a new GUID.
	strGUID = objTL.Guid
	
	' Clean up the GUID.
	strGUID = Left(strGUID, (Len(strGUID)-2))
	
	' Clean up.
	Set objTL = Nothing
	
	' If we have value, return it.
	If (Len(strGUID) > 0) Then
		MakeGUID = strGUID
	End If
End Function
'-----------------------------------------------------------------------------
'			Function to build that Session file for Aelita
'-----------------------------------------------------------------------------
Function BuildSessionFile(strSessionFileName)
	Set objFSObject = CreateObject("Scripting.FileSystemObject")
	'Get File and Session Name
	arrSessionFileName = Split(strSessionFileName,vbTab)
	
	'Open File
	'This File is for Aelita 6.x
	Set objSessionFile = objfso.CreateTextFile(arrSessionFileName(0),True,True)
	'This file if for NDS Migrator
	'Set objSessionFile = objfso.CreateTextFile(arrSessionFileName(0),True)
	
	'Get random GUID
	'GUID = MakeGUID
	'insert header
	'objSessionFile.WriteLine "Session "& arrSessionFileName(1) & ": " & GUID
	'objSessionFile.WriteLine "Started: " & Now()
	objSessionFile.WriteLine "Session 22: {21AD8B68-2A42-459e-BD29-F082F47E71B2}"
	objSessionFile.WriteLine "Started: 07-26-2006 09:17"

	objSessionFile.WriteLine "NDS Tree: FUN"
	objSessionFile.WriteLine "AD Server: MYDomainController001.migp.my.bank.net"
	'Sort Dictionary
	Call SortDictionary(objGroupDic)
	'output to the Session File
	For Each Key In objGroupDic.Keys
		objSessionFile.WriteLine Key
		objSessionFile.WriteLine "Group"
		objSessionFile.WriteLine objGroupDic.Item(Key)
		objSessionFile.WriteLine "group"
	Next
	'Sort Dictionary
	Call SortDictionary(objUserDic)
	'output to Session File
	For Each Key In objUserDic.Keys
		objSessionFile.WriteLine Key
		objSessionFile.WriteLine "User"
		objSessionFile.WriteLine objUserDic.Item(Key)
		objSessionFile.WriteLine "user"
	Next
	'Sort Dictionary
	For Each Key In objGroupOUDic.Keys
		objSessionFile.WriteLine Key
		objSessionFile.WriteLine "Organizational Unit"
		objSessionFile.WriteLine objGroupOUDic.Item(Key)
		objSessionFile.WriteLine "group"
	Next
	
	For Each Key In objRoleDic.Keys
		objSessionFile.WriteLine Key
		objSessionFile.WriteLine "Organizational Role"
		objSessionFile.WriteLine objRoleDic.Item(Key)
		objSessionFile.WriteLine "group"
	Next
	
	
End Function
'-----------------------------------------------------------------------------
'			Cool new function to sort any Dictionary object and keep the items
'-----------------------------------------------------------------------------
Sub SortDictionary(objDict)
  Dim nCount, strKey
  nCount = 0
  
  '-- Redim the array to the number of keys we need 
  ReDim tmpKeyArray(objDict.Count - 1)
  ReDim tmpItemArray(objDict.Count -1)

  '-- Load the array 
  For Each strKey In objDict.Keys

    '-- Set the array element to the key 
    tmpKeyArray(nCount) = strKey
    tmpItemArray(nCount) = objDict.Item(strKey)
    WScript.Echo nCount & vbTab & tmpKeyArray(ncount)& vbTab & tmpItemArray(ncount)

    '-- Increment the count 
    nCount = nCount + 1

  Next 

	'SortArray------------------------------
  Dim iTemp, jTemp, strTempKey, strTempItem

  For iTemp = 0 To UBound(tmpKeyArray)  
    For jTemp = 0 To iTemp  

      If strComp(tmpKeyArray(jTemp), tmpKeyArray(iTemp)) > 0 Then
        'Swap the array positions
        strTempKey = tmpKeyArray(jTemp)
        strTempItem = tmpItemArray(jTemp)
        
        tmpKeyArray(jTemp) = tmpKeyArray(iTemp)
        tmpItemArray(jTemp) = tmpItemArray(iTemp)
        
        tmpKeyArray(iTemp) = strTempKey
        tmpItemArray(itemp) = strTempItem
      End If 

    Next 
  Next 

  'Reload Dictionary 
  objDict.RemoveAll
  For iTemp = 0 To UBound(tmpKeyArray)
  	WScript.Echo iTemp & vbTab & tmpKeyArray(iTemp)& vbTab & tmpItemArray(iTemp)
   	objDict.add tmpKeyArray(iTemp), tmpItemArray(iTemp)
  Next 
End Sub
'-----------------------------------------------------------------------------
'			Function to produce required File
'-----------------------------------------------------------------------------
Function GenerateFileName
	intSessionID = DatePart("d",Now)
	Count = 0
	Do While objFSO.FileExists("C:\Jimworks\Session_" & intSessionID & "-" & count & ".txt")
		count = count + 1
	Loop
	GenerateFileName = "C:\Jimworks\Session_" & intSessionID & "-" & count & ".txt" & vbTab & intSessionID
End Function
'-----------------------------------------------------------------------------
'			Function to add to a container Group
'-----------------------------------------------------------------------------
Function AddToContainerGroup(dnAccount,dnLDAP)
	'Get Proper Container Group
	arrDN = Split(dnAccount,"/")
	For DNCount = LBound(arrDN) To (UBound(arrDN)-1)
		If DNCount > 0 then
			DestDN = DestDN & "/" & arrDN(DNCount)
		Else
			DestDN = arrDN(DNCount)
		End If
	Next
	If objGroupOUDic.Exists(DestDN) Then
		OU = objGroupOUDic.Item(DestDN)
		call AddToGroup(OU,dnLDAP)
	End If
End Function
