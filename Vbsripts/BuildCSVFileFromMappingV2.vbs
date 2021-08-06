'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: BuildCSVFileFromMapping.vbs
'
' AUTHOR: Jim Borecky
' DATE  : 1/30/2007
'
' COMMENT: This script will take the group and User mappings file and create
'			a CSV file for NDSMigrator.
'
'
'			IMPORTANT!!!!!!!
'				 Need special DLL's to use this script
'				 novell-activex_core-devel-2004.10.06-1.exe
'				 novell-activex_ndap-devel-2006.06.14-1windows.zip
'				 both can be found on the Novell Web site
'
'			AND YOU MUST ALSO BE LOGGED INTO NOVELL!!!!!!!
'
'
'	Changes:
'		2-2-2007 Ver.2-Making changes so that the Novell team doesn't have
'		to supply the NDS name in a Berkley Format
'
'==========================================================================
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Set objDictionary = CreateObject("Scripting.Dictionary")
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set NWDir1 = CreateObject("NWDirLib.NWDirCtrl.1")
Set NWDirQuery1 = CreateObject("NWDirQueryLib.NWDirQuery")
Set NWUsrGrp = CreateObject("NWUsrGrpLib.NWGroup")

'Change me!!!!
strDestDomain = "LDAP://yourdomain.net"

'-------------------------- Main ------------------------------------------------

if wscript.arguments.count = 1 Then
	strInputFile = wscript.arguments.item(0)
Else
	WScript.Echo
	WScript.Echo
	WScript.Echo
	Wscript.echo
	Wscript.echo "You need to specify an a text file to use"
	Wscript.echo "cscript BuildCSVFileFromMapping.vbs InputFileName"
	wscript.echo
	Wscript.echo "Example:"
	Wscript.echo "cscript BuildCSVFileFromMapping.vbs Input.txt"
	wscript.echo 
	WScript.Echo "--------------------------------------------------------"
	WScript.Echo "Mapping Example:"
	WScript.Echo "Full NDS Path	  <tab>      adSAMAccount Name"
	WScript.Echo
	WScript.Echo "NDS://FUNNOVELL/O=MYORG/OU=US/CN=PSBACK	PSBACK"
	WScript.Echo "--------------------------------------------------------"
	WScript.Quit
End If

'Get all of the files set for action
Set objInputFile = objfso.OpenTextFile(strInputFile,ForReading)
Set objOutputFile = objFSO.CreateTextFile(split(strInputFile,".")(0) & ".csv")
Set objErrorLog = objFSO.CreateTextFile(split(strInputFile,".")(0) & "_Errors.csv")

'Place header on to import file
objOutputFile.WriteLine """Full NDS Path"",""NDS Object Type"",""Full LDAP Path"",""AD Object Type"",""GUID"",""SAM Account Name"",""User Principal Name"",""Migration Status"","
objErrorLog.WriteLine """Full NDS Path"",""NDS Object Type"",""Full LDAP Path"",""AD Object Type"",""GUID"",""SAM Account Name"",""User Principal Name"",""Migration Status"","

'Main
Do While objInputFile.AtEndOfStream <> True
	strNew = objInputFile.ReadLine
	'Split into two parts
	WScript.Echo strNew
	strFullNDS = Split(strNew,vbTab)(0)
	If InStr(strFullNDS,"\") Then
		strFullNDS = ChangeFormat(strFullNDS)
	End if
	strSAMAccount = Split(strNew,vbTab)(1)
	strNDSType = GetNDSUserType(strFullNDS)
	strADInfo = GetADInfo(strSAMAccount,strDestDomain)
	'Something is changing the slahes on the file name
	strFullNDS = replace(strFullNDS,"\","/")
	'Record to appropriate files
	If (InStr(strADInfo,"Does not exist")) Or (instr(strNDSType,"Unknown")) Then
		objErrorLog.WriteLine """" & strFullNDS & """,""" & strNDSType & """," & strADInfo & ",""" & strSAMAccount & ""","'"" & "  " & """,""" & "  " & ""","
	else
		objOutputFile.WriteLine """" & strFullNDS & """,""" & strNDSType & """," & strADInfo & ",""" & strSAMAccount & ""","'"" & "  " & """,""" & "  " & ""","
	End If
Loop

'Close up shop
objInputFile.Close
objOutputFile.Close
objErrorLog.Close

'----------------------------------------------------------------------------
'	This Function takes a distinguished Name from NDS and puts it in what
'	I believe is a LDAP berkley format.
'	INCOMING FORMAT NDS:\\FUNNOVELL\MYORG\US\GPAELITA
'	OUTGOING FORMAT NDS://NDS://FUNNOVELL/O=MYORG/OU=US/CN=GPAELITA
'----------------------------------------------------------------------------
Function ChangeFormat(strNDSName)
	strContext = Split(strNDSName,"\\")(1) 'Pulling off header
	arrName = Split(strContext,"\")
	strNewName = "NDS://"
	i=0
	For i = LBound(arrName) To UBound(arrName)
		If i = LBound(arrName) Then 
			strNewName = strNewName & arrName(i)
		ElseIf i = (LBound(arrName)+1) then
			strNewName = strNewName & "/O=" & arrName(i)
		ElseIf i= UBound(arrName) Then
			strNewName = strNewName & "/CN=" & arrName(i)
		Else
			strNewName = strNewName	& "/OU=" & arrName(i)
		End If
	Next
	ChangeFormat = strNewName
End Function
' ---------------------------------------------------------------------------
'        Gets Distinguished Name Function
'----------------------------------------------------------------------------

function GetADInfo(strAccountName, strDomain)

	dim strDistinguishedName
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
 
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
 
	objCommand.CommandText = _
	    "<" & strDomain & ">;(&(objectCategory=*)" & _
	         "(samAccountName=" & strAccountName & "));distinguishedName,samAccountName,objectClass,objectGUID;subtree"
  
	Set objRecordSet = objCommand.Execute
 
	If objRecordset.RecordCount = 0 Then	
	    WScript.Echo "sAMAccountName: " & strAccountName & " does not exist."
	    GetADInfo = """Does not exist in Domain" & """,""" & "" & """,""" & "" & """"
	Else
	    'WScript.Echo strUserName & " exists."
	    strDistinguishedName = objRecordset.Fields("distinguishedName")
	    'WScript.Echo strDistinguishedName
	    arrObjectClass = objRecordSet.Fields("objectClass")
	    objectType = getClass(arrObjectClass)
	    arrObjectGUID = objRecordSet.Fields("objectGUID")
	    strHexGuid = OctetToHexStr(arrObjectGUID)  
	    strGUID = HexGuidToGuidStr(strHexGuid)
	    'Passing back all the AD info int the proper format
	    GetADInfo = """" & strDomain & "/" & strDistinguishedName & """,""" & objectType & """,""" & strGUID & """"
	End If
 	 	
	objConnection.Close
end Function
'--------------------------------------------------------------------------------------------------
'	Function to figure out the object class and return
'--------------------------------------------------------------------------------------------------
Function GetClass(arr)
	If IsArray(arr) Then
		For i = lbound(arr) To UBound(arr)
			strCN = lcase(arr(i))
			'WScript.Echo strCN
			If strCN = "user" Then
				strType = "User"
			End If
			If strCN = "group" Then
				strType = "Group"
			End If
		Next
	End If
	'WScript.Echo ":" & strType
	GetClass = strType
End Function
'--------------------------------------------------------------------------------------------------
'	Two Functions stolen from the web that converts the Hex string into a readable GUID
'--------------------------------------------------------------------------------------------------  
Function OctetToHexStr(arrbytOctet)  
  	' Function to convert OctetString (byte array) to Hex string.  
	Dim k  
  	OctetToHexStr = ""  
  
	For k = 1 To Lenb(arrbytOctet)    
		OctetToHexStr = OctetToHexStr & Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)  
  	Next  
  	
End Function  
  
Function HexGuidToGuidStr(strGuid)  
  ' Function to convert Hex Guid to display form.  
  Dim k  
  HexGuidToGuidStr = "{"  
  
	For k = 1 To 4  
  		HexGuidToGuidStr = HexGuidToGuidStr & Mid(strGuid, 9 - 2*k, 2)  
  	Next  
  
	HexGuidToGuidStr = HexGuidToGuidStr & "-"  
  	
  	For k = 1 To 2  
  		HexGuidToGuidStr = HexGuidToGuidStr & Mid(strGuid, 13 - 2*k, 2)  
  	Next  
  
	HexGuidToGuidStr = HexGuidToGuidStr & "-"  
	For k = 1 To 2  
  		HexGuidToGuidStr = HexGuidToGuidStr & Mid(strGuid, 17 - 2*k, 2)  
  	Next  
  	
	HexGuidToGuidStr = HexGuidToGuidStr & "-" & Mid(strGuid, 17, 4)  
  	HexGuidToGuidStr = HexGuidToGuidStr & "-" & Mid(strGuid, 21)  
  	
  	HexGuidToGuidStr = HexGuidToGuidStr & "}"
  
End Function
'---------------------------------------------------------------------------------------------------
'	Script to get NDS User and Group object's or returns unknown
'---------------------------------------------------------------------------------------------------
Function GetNDSUserType(NDSFullPathName)
	On Error Resume Next
	NDSFullPathName = replace(NDSFullPathName,"/","\")
	Dim results
	strType = "Unknown"
	
	NWDirQuery1.FullName = NDSFullPathName
	NWDirQuery1.Filter = "(CN = *)" 'Change to *
	NWDirQuery1.Fields = "CN, Object Class"
	NWDirQuery1.SearchScope = 0 '0	 qrySearchEntry	 		Searches only the entry level.
								'1	 qrySearchSubordinates	Searches entries subordinate to the entry given in the search context.
								'2	 qrySearchSubtree		Searches the entire sub-tree. This is below the entry given in the search context.
	NWDirQuery1.SearchMode = 0
	NWDirQuery1.MaximumResults = 100
	Set results = NWDirQuery1.Search
	'WScript.Echo results.Count
	
	For Each obj In results
		arr = obj.FieldValue("Object Class")
		If IsArray(arr) Then
			For i = lbound(arr) To UBound(arr)
				strCN = arr(i)
				'WScript.Echo strCN
				If strCN = "User" Then
					strType = "User"
				End If
				If strCN = "Group" Then
					strType = "Group"
				End If
			Next
		End If
		'wscript.Echo obj.Fullname
		GetNDSUserType = strType
	Next
	GetNDSUserType = strType
End Function

