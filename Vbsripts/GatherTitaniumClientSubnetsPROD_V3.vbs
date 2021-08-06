'=================================================================================
' NAME: GatherTitaniumClientSubnets.vbs
'
' AUTHOR: Jim Borecky
' DATE  : 7/15/2014
'
' COMMENTS: This script will get the Titanium client subnets and 
'	create a new file and a Diff file.
'
' ------------------------------------------------------------------
' Version   Date              Initial         Comment
' -------   --------          -------         -------
'  1.0      7/15/2014        Jim Borecky       Original
'
'================================================================================
' objShell.exec ("cmd svn update C:\Projects\Project1 --username buildrobot --password IAmARobot1")
'===========================Configurable variables=====================================
strStartTime = Now
wscript.echo "GatherTaniumClientSubnets.vbs"
WScript.Echo "Start Time:" & strStartTime

'hardcoding creds here prevents prompting
account = "myId"
password = "only4Him*"

'Const xlExcel7 = 39
Const xlExcel12 = 51


bolDEBUG = False
bolSILENT = False
bolVERBOSE = False
bolSTOPONERRORS = False
bolUploadFile = True
bolLoadToSubversion = False
intQuestionNumber = 12842
strSPURL = "\\somesad.sharepoint.net\sites\risk-eis-305\tanium\Shared%20Documents\IsolatedSubnet"

strMasterFile = "SeperatedandIsolatedmaster.xlsx"
URL = "https://tanium.net/soap"

'hardcoding file here prevents arguments
FileToUpdate = "E:\TaniumSim\DifferentDirectory\SeparatedSubnets.txt"
IsolatedToUpdate = "E:\TaniumSim\DifferentDirectory\IsolatedSubnets.txt"
'================================================================================


Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
Set objDictLoadFile = CreateObject("Scripting.Dictionary")
Set objDictExcelSeparated = CreateObject("Scripting.Dictionary")
Set objDictExcelIsolated = CreateObject("Scripting.Dictionary")
Set objDictResults = CreateObject("Scripting.Dictionary")
Set objDictMaskChangeCheck = CreateObject("Scripting.Dictionary")
Set objDictSubnetVPN = CreateObject("Scripting.Dictionary")
Set objDictDiffFile = CreateObject("Scripting.Dictionary")
Set objDictDiffISOFile = CreateObject("Scripting.Dictionary")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
bolSourceFileError = False
strDupsReport = ""

'==================================================================
' Get arguements
'==================================================================
If bolDEBUG Or bolVERBOSE Then
	bolSILENT = False
End If
If bolDEBUG Then
	bolVERBOSE = True
End If
If FileToUpdate = "" Then
	if wscript.arguments.count = 1 then
		FileToUpdate = wscript.arguments.item(0)
	ElseIf wscript.arguments.count = 3 Then
		FileToUpdate = wscript.arguments.item(0)
		account = wscript.arguments.item(1)
		password = wscript.arguments.item(2)
	Else
		ShowHelp
		wscript.quit
	end If
End if


'==================================================================
'Get Creds
'==================================================================
If account = "" Then
	account = InputBox("User ID", "Tanium Credentials")
End If
If password = "" Then
	'password = InputBox("Password", "Tanium Credentials")
	WScript.Echo "WARNING:LOCATE PASSWORD WINDOW TO CONTINUE."
	password = getPassword("Tanium","Enter your password")
End If

'===================================================================
'Copy master locally
'===================================================================
strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
WScript.Echo "RETRIEVING MASTER FILE TO:" & strScriptPath & strMasterFile
Call SPCopy(strScriptPath, strMasterFile ,strSPURL)

'===================================================================
'Get Master List info
'===================================================================
' Bind to Excel object.
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If (Err.Number <> 0) Then
    On Error GoTo 0
    Wscript.Echo "Excel application not found."
    Wscript.Quit
End If
On Error GoTo 0

Set objWorkbook = objExcel.Workbooks.Open(strScriptPath & strMasterFile)

intRow = 1
objExcel.ActiveWorkbook.Sheets(1).Select()
	
Do Until objExcel.Cells(intRow,1).Value = ""
    strData = objExcel.Cells(intRow, 1).Value
    Wscript.Echo intRow & ":" & strData
    If not objDictExcelSeparated.Exists(strData) Then
    	objDictExcelSeparated.Add strData, strData
    End If
    intRow = intRow + 1
Loop

WScript.Echo "DONE READING:" & objExcel.ActiveWorkbook.ActiveSheet.Name

intRow = 1
objExcel.ActiveWorkbook.Sheets(2).Select()
	
Do Until objExcel.Cells(intRow,1).Value = ""
    strData = objExcel.Cells(intRow, 1).Value
    Wscript.Echo intRow & ":" & strData
    If not objDictExcelIsolated.Exists(strData) Then
    	objDictExcelIsolated.Add strData, strData
    End If
    intRow = intRow + 1
Loop

WScript.Echo "DONE READING:" & objExcel.ActiveWorkbook.ActiveSheet.Name
objExcel.ActiveWorkbook.Close()
objExcel.Quit

'===================================================================
' Making separate local files
'===================================================================
' FileToUpdate = "E:\TaniumSim\DifferentDirectory\SeparatedSubnets.txt"
Set objCurrentFile = objFSO.CreateTextFile(FileToUpdate,true)
For Each key In objDictExcelSeparated.Keys
	WScript.Echo "->" & key
	objCurrentFile.WriteLine key
Next
objCurrentFile.Close()
' IsolatedToUpdate = "E:\TaniumSim\DifferentDirectory\IsolatedSubnets.txt"
Set objCurrentFile = objFSO.CreateTextFile(IsolatedToUpdate,true)
For Each key In objDictExcelIsolated.Keys
	WScript.Echo "->" & key
	objCurrentFile.WriteLine key
Next
objCurrentFile.Close()

'===================================================================
'Make Tanium Query
'===================================================================
xmlResult = GetResultData(account,password,intQuestionNumber)
 
If bolDEBUG Then
	WScript.Echo Now
	WScript.Echo vbNewLine & "RESULTS:" & vbNewLine & xmlResult
End If

'Parse Tanium Query
Call parseResult_SetsField(xmlResult)
'===================================================================
'Load current Separated Subnet file
'==================================================================
Set objCurrentFile = objFSO.OpenTextFile(FileToUpdate,ForReading)
Do While objCurrentFile.AtEndOfStream <> True
	strNewLine = objCurrentFile.ReadLine
	objDictLoadFile.Add strNewLine, strNewLine
	
	'Pull the subnet mask and add to separate dictionary
	arrSubnet = split(strNewLine,"/")
	
	'Check to make sure we do not have any duplicates.
	If objDictMaskChangeCheck.Exists(arrsubnet(0)) Then
		If Not bolSILENT Then
			WScript.Echo "Please make correction in source file:"
			WScript.Echo "Subnet: " & arrsubnet(0) & " is duplicated in the source file."
			WScript.Echo "As " & strNewLine & " - " & arrsubnet(0) & "/" & objDictMaskChangeCheck.Item(arrsubnet(0)) & "."
			WScript.Echo vbNewLine & vbNewLine
		End If
		
		'Put duplicate errors in the windows log file
		subWriteEvent 1, "Please make correction in source file:" & vbNewLine & _
			             "Subnet: " & arrsubnet(0) & " is duplicated in the source file." & vbNewLine & _
			             "As " & strNewLine & " - " & arrsubnet(0) & "/" & objDictMaskChangeCheck.Item(arrsubnet(0)) & "."
		strDupsReport = strDupsReport & "Subnet: " & arrsubnet(0) & " is duplicated in the source file." & vbNewLine & _
			             "As " & strNewLine & " - " & arrsubnet(0) & "/" & objDictMaskChangeCheck.Item(arrsubnet(0)) & "." & vbNewLine
			             
		If bolSTOPONERRORS then
			bolSourceFileError = True
		End If
	Else
		'Add to check
		objDictMaskChangeCheck.Add arrSubnet(0),arrSubnet(1)
	End If
Loop
objCurrentFile.Close()

'==================================================================
'Load current Isolated Subnet file
'==================================================================
Set objCurrentFile = objFSO.OpenTextFile(IsolatedToUpdate,ForReading)
Do While objCurrentFile.AtEndOfStream <> True
	strNewLine = objCurrentFile.ReadLine
	objDictSubnetVPN.Add strNewLine, strNewLine
	
	'Pull the subnet mask and add to separate dictionary
	arrSubnet = split(strNewLine,"/")
	
	'Check to make sure we do not have any duplicates.
	If objDictMaskChangeCheck.Exists(arrsubnet(0)) Then
		If Not bolSILENT Then
			WScript.Echo "Please make correction in source file:"
			WScript.Echo "Subnet: " & arrsubnet(0) & " is duplicated in the source file."
			WScript.Echo "As " & strNewLine & " - " & arrsubnet(0) & "/" & objDictMaskChangeCheck.Item(arrsubnet(0)) & "."
			WScript.Echo vbNewLine & vbNewLine
		End If
		
		'Put duplicate errors in the windows log file
		subWriteEvent 1, "Please make correction in source file:" & vbNewLine & _
			             "Subnet: " & arrsubnet(0) & " is duplicated in the source file." & vbNewLine & _
			             "As " & strNewLine & " - " & arrsubnet(0) & "/" & objDictMaskChangeCheck.Item(arrsubnet(0)) & "."
		strDupsReport = strDupsReport & "Subnet: " & arrsubnet(0) & " is duplicated in the source file." & vbNewLine & _
			             "As " & strNewLine & " - " & arrsubnet(0) & "/" & objDictMaskChangeCheck.Item(arrsubnet(0)) & "." & vbNewLine
			             
		If bolSTOPONERRORS Then
			bolSourceFileError = True
		End If
	Else
		'Add to check
		objDictMaskChangeCheck.Add arrSubnet(0),arrSubnet(1)
	End If
Loop
objCurrentFile.Close()

'==================================================================
'If all checks out then continue
'==================================================================
If not bolSourceFileError Then
	'Add needed items <----------------------------------------------------------------------
	For Each key In objDictResults.Keys
		'Check for [no results]
		If key <> "[no results]" then
			If Not objDictLoadFile.Exists(key) Then
				arrSubnetKey = split(key,"/")
				If Not objDictMaskChangeCheck.Exists(arrSubnetKey(0)) Then
					
					If arrSubnetKey(1) = "32" Then
						If bolVERBOSE Then
							WScript.Echo key
						End If
						If CheckForSubnet(key) then
							objDictDiffISOFile.Add key, key
						End if
					Else
						If bolVERBOSE Then
							WScript.Echo "Adding subnet:" & key
						End If
						intKey = CovertToInt(Split(key,"/")(1))
						If intKey > 24 And intKey < 29 then
							objDictLoadFile.Add key, key
							objDictDiffFile.Add key, key
						End if
					End if
					
					'Add to check
					If objDictMaskChangeCheck.Exists(arrSubnetKey(0))then
						WScript.Echo "Duplicate:" & arrSubnetKey(0)
					Else
						objDictMaskChangeCheck.Add arrSubnetKey(0),arrSubnetKey(1)
						End If
			
				End If
			End If
		End If 
	Next
	
	'Create new files <----------------------------------------------------------------------
	'original file
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	
	'Get time zone information
	Set colItems = objWMIService.ExecQuery("Select * From Win32_TimeZone")
	For Each objItem In colItems
		timeZone = objItem.StandardName
	Next
	timeZone = lcase(Left(timeZone,1))

	'Get Server
	strServer = Split(URL,"/")(2)
	strServer = Split(strServer,".")(0)
	If strServer = "127" Then
		strServer = "tanium"
	End If
	
	'Build the file name
	strHour = Hour(Now)
	strMinute = Minute(now)
	strYear = Year(Now)
	strMonth = Month(Now)
	strDay = Day(Now)
	'Fix short format
	If len(strHour) = 1 Then strHour = "0" & strHour
	If len(strMinute) = 1 Then strMinute = "0" & strMinute
	If len(strMonth) = 1 Then strMonth = "0" & strMonth
	If len(strDay) = 1 Then strDay = "0" & strDay

	If InStr(FileToUpdate,"\") Then
		WScript.Echo "PASS"
		arrFileToUpdate = Split(FileToUpdate,"\")
		newFileToUpdate = arrFileToUpdate(UBound(arrFileToUpdate))
		newFileToUpdate = Split(newFileToUpdate,".")(0)
	else
		fileNameRoot = Split(FileToUpdate,".")(0)
	End If
	FileName = newFileToUpdate & "_" & strYear & strMonth & strDay & "_" & strHour & strMinute & timeZone & "t(" & strServer & ").txt"
	FileName2 = FileName
	If bolVERBOSE then
		WScript.Echo "NEW FILE NAME:" & FileName
	End If
	
	WScript.Echo "SORTING SEPARATED SUBNETS STARTED AT:" & now
	'Sort the Dictionary
	Call SortDictionary(objDictLoadFile)
	WScript.Echo "SORTING ISOLATED SUBNETS STARTED AT:" & Now
	Call SortDictionary(objDictDiffISOFile)
	

	WScript.Echo "WRITING FILE..."
	'Create new subnet file<--------------------------------------------------------------------------		
	Set objNewFile = objFSO.CreateTextFile(FileName,ForWriting)
		For Each key In objDictLoadFile.Keys
			If Trim(key) <> "" Then
				objNewFile.WriteLine key
			End If
		Next
	objNewFile.Close()'New SeparatedSubnets file
	

	'Create a new Diff file<--------------------------------------------------------------------------
	FileName = newFileToUpdate & "_DIFF_" & strYear & strMonth & strDay & "_" & strHour & strMinute & timeZone & "t(" & strServer & ").txt"
	Call SortDictionary(objDictDiffFile)
	Set objNewFile = objFSO.CreateTextFile(FileName,ForWriting)
	'Create inital reporting for dups <-----------------------------------------------------------------------------------
	objNewFile.WriteLine "======================================================="
	objNewFile.WriteLine "              DUPLICATES "
	objNewFile.WriteLine "======================================================="
	
	If strDupsReport <> "" Then
		objNewFile.WriteLine strDupsReport
	Else 
		objNewFile.WriteLine "NO DUPLICATES FOUND"
	End if
	objNewFile.WriteLine "======================================================="
	objNewFile.WriteLine "              MISSING IN SEPARATED FILE "
	objNewFile.WriteLine "======================================================="
	
	'Create inital reporting for diffs <-----------------------------------------------------------------------------------
	For Each key In objDictDiffFile.Keys
		If Trim(key) <> "" Then
			If Not InStr(key,"192.168.")> 0 Then
					objNewFile.WriteLine key
			End If
		End If
	Next
	
	'Create inital reporting for isolated <--------------------------------------------------------------------------------
	objNewFile.WriteLine "======================================================="
	objNewFile.WriteLine "              MISSING IN ISOLATED FILE "
	objNewFile.WriteLine "======================================================="
	
	For Each key In objDictDiffISOFile.Keys
		If Trim(key) <> "" Then
			objNewFile.WriteLine key
		End If
	Next
		
	objNewFile.Close() 'DIFF File
	
	
	'Upload file to sharepoint<-----------------------------------------------------------------------
	If bolUploadFile Then
		strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
		strDestinationFileName = strSharePointLocation & Replace(FileName,"_","%137")
		If Not bolSILENT then
			WScript.Echo "WARNING:COPYING DIFF FILE TO SHAREPOINT SITE:"
			WScript.Echo FileName & vbNewLine 
		End if
		Call SPCopy(strSPURL, FileName ,strScriptPath)
		If Not bolSILENT then
			WScript.Echo "WARNING:COPYING CONFIGURATION FILE TO SHAREPOINT SITE:"
			WScript.Echo FileName2 & vbNewLine 
		End if
		Call SPCopy(strSPURL, FileName2 ,strScriptPath)

	End If
	
	'Upload file to Subversion<-----------------------------------------------------------------------
	If bolLoadToSubversion then
		' WshShell.exec ("cmd svn update C:\Projects\Project1 --username buildrobot --password IAmARobot1")
	End If

End If

'Notify the user we are finished
If Not bolSILENT Then
	WScript.Echo "Start Time:" & strStartTime
	WScript.Echo "DONE:" & Now
End If
'The End
'==================================================================================
'                           FUNCTIONS AND SUBS
'                       
'==================================================================================
Function CovertToInt(strInput)
	'Cause cint sucks
	Select Case trim(strInput)
		Case "1" CovertToInt = 1
		Case "2" CovertToInt = 2
		Case "3" CovertToInt = 3
		Case "4" CovertToInt = 4
		Case "5" CovertToInt = 5
		Case "6" CovertToInt = 6
		Case "7" CovertToInt = 7
		Case "8" CovertToInt = 8
		Case "9" CovertToInt = 9
		Case "10" CovertToInt = 10
		Case "11" CovertToInt = 11
		Case "12" CovertToInt = 12
		Case "13" CovertToInt = 13
		Case "14" CovertToInt = 14
		Case "15" CovertToInt = 15
		Case "16" CovertToInt = 16
		Case "17" CovertToInt = 17
		Case "18" CovertToInt = 18
		Case "19" CovertToInt = 19
		Case "20" CovertToInt = 20
		Case "21" CovertToInt = 21
		Case "22" CovertToInt = 22
		Case "23" CovertToInt = 23
		Case "24" CovertToInt = 24
		Case "25" CovertToInt = 25
		Case "26" CovertToInt = 26
		Case "27" CovertToInt = 27
		Case "28" CovertToInt = 28
		Case "29" CovertToInt = 29
		Case "30" CovertToInt = 30
		Case "31" CovertToInt = 31
		Case "32" CovertToInt = 32
		Case Else CovertToInt = 0
	End Select	
End Function
Function CheckForSubnet(strIPAddess)
	If bolVERBOSE Then
		WScript.Echo "Checking:" & strIPAddess
	End If
	CheckForSubnet = false
	For Each Subnet In objDictSubnetVPN
		arrSubnet = Split(Subnet,"/")
		If CheckToSeeWithinSubnet(strIPAddess,arrSubnet(0),arrSubnet(1)) Then
			CheckForSubnet = True
		End If
	Next
End Function
'==================================================================================
' Sub SortDictionary(objDict)
'
' Dependancies on other Subs or Functions
'	1) None
'
'Comments:
'			Sub to sort any Dictionary object and keep the items
'			I modified this slightly to sort by integers instead of strings
'==================================================================================
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
    
    '-- Increment the count 
    nCount = nCount + 1

  Next 

	'SortArray------------------------------
  Dim iTemp, jTemp, strTempKey, strTempItem

  For iTemp = 0 To UBound(tmpKeyArray)  
    For jTemp = 0 To iTemp  
 	
 	  If strComp(tmpKeyArray(jTemp), tmpKeyArray(iTemp)) > 0 Then
      'If clng(trim(tmpKeyArray(jTemp))) > clng(trim(tmpKeyArray(iTemp))) Then
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
  	objDict.add tmpKeyArray(iTemp), tmpItemArray(iTemp)
  	
  	If bolDEBUG then
  		WScript.Echo "Sort order:" & tmpKeyArray(iTemp) & vbTab & tmpItemArray(iTemp)
  	End If
  Next 
End Sub 'SortDictionary
'==================================================================================
' Function GetResultData(tanium_username, tanium_password)
'
' Dependancies on other Subs or Functions
'	1)none
'
'Comments:
' 	This Function will Query the Tanium Service to retrieve the current saved
'	question.
'==================================================================================
Function GetResultData(tanium_username, tanium_password,intQuestion)
	SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
	Dim soapXML
		soapXML = "<?xml version=""1.0"" encoding=""utf-8""?>"
		soapXML = soapXML & "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
  		soapXML = soapXML & "<soap:Body>"
    	soapXML = soapXML &     "<tanium_soap_request xmlns=""urn:TaniumSOAP"">"
    	soapXML = soapXML & 			"<auth>"
	    soapXML = soapXML & 				"<username>" & tanium_username & "</username>"
	    soapXML = soapXML & 				"<password>" & tanium_password & "</password>"
	    soapXML = soapXML & 			"</auth>"
        soapXML = soapXML &     "<command xmlns="""">GetResultData</command>"
        soapXML = soapXML &  	   "<object_list xmlns="""">"
        soapXML = soapXML &            "<saved_question>"
        soapXML = soapXML &               "<id>" & intQuestion & "</id>"
        soapXML = soapXML &            "</saved_question>"
        soapXML = soapXML &        "</object_list>"
        soapXML = soapXML &        "</tanium_soap_request>"
        soapXML = soapXML & "</soap:Body>"
        soapXML = soapXML & "</soap:Envelope>"
	
	If bolDEBUG then
		WScript.Echo "=====QUERY====="
    	WScript.Echo soapXML
    	WScript.Echo "==============="
    End if
    
   Set resultXML = CreateObject("MSXML2.DOMDocument")
   Set returnXML = CreateObject("MSXML2.DOMDocument")	
	
	Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	
	objHTTP.open "POST", URL, False 
	objHTTP.setRequestHeader "Content-Type", "text/xml"
	objHTTP.setRequestHeader "SOAPAction", url
	objHTTP.SetOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

	objHTTP.send soapXML 
	
	strXml = objHTTP.responseText
	
	resultXML.LoadXml strXml
	
	If bolDEBUG Then
		WScript.Echo "=====RESPONSE====="
    	WScript.Echo strXml
    	WScript.Echo "=================="
    End If

	Set objNodeList = resultXML.getElementsByTagName("ResultXML")
	
	If objNodeList.length > 0 Then
		For each x in objNodeList
			GetResultData = x.Text
		Next
	Else 
		GetResultData = ""
	End If
End Function 'GetResultData
'==================================================================================
' Function parseResult_SetsField(strXML)
'
' Dependancies on other Subs or Functions
'	1)Set objDictResults = CreateObject("Scripting.Dictionary") <- Global variable
'
'Comments:
' 	This Function will parse the result_sets XML returned by the SOAP service
'==================================================================================
Function parseResult_SetsField(strXML)
	Set xmlCheck = CreateObject("MSXML2.DOMDocument")
	
	xmlCheck.loadXML strXml
	
	'============Check for timestamp to report ===========
	Set NodeCheck = xmlCheck.getElementsByTagName("now")
		If NodeCheck.length > 0 Then
			For Each x in NodeCheck
				If Not bolSILENT Then
					WScript.Echo "TIME OF LAST QUERY:" & x.text
				End If
				nowNode=x.Text
			Next
		Else 
			If Not bolSILENT Then
				WScript.Echo "WARNING:No time stamp found"
			End If
			nowNode = ""
		End If	
		
	'================Get the item count====================
	Set NodeCheck = xmlCheck.getElementsByTagName("item_count")
		If NodeCheck.length > 0 Then
			For Each x in NodeCheck
				If Not bolSILENT Then
					WScript.Echo "ITEM COUNT:" & x.text & vbNewLine & vbNewLine
				End If
				itemCount=x.Text
			Next
		Else 
			If Not bolSILENT Then
				WScript.Echo "ERROR:No Results Found"
			End If
			WScript.Quit
			itemCount = ""
		End If	
		
	'===============Isolate the result set==================
	Set NodeCheck = xmlCheck.getElementsByTagName("rs")
		If NodeCheck.length > 0 Then
			For Each x in NodeCheck
				If bolDEBUG Then
					WScript.Echo "RESULT_SETS:" & x.text
				End If
				resultSets=x.Text
			Next
		Else 
			If Not bolSILENT Then
				WScript.Echo "ERROR:No Results found"
			End If
			WScript.Quit
			resultSets = ""
		End If
		
	'===================PARSING RESULTS====================
	Dim LoadResultsArray()
	ReDim LoadResultsArray(itemCount*2)
	count = 0
	Set NodeCheck = xmlCheck.getElementsByTagName("v")
		If NodeCheck.length > 0 Then
			For Each x in NodeCheck
				LoadResultsArray(count)=x.Text
				count = count + 1
			Next
		Else 
			If Not bolSILENT Then
				WScript.Echo "ERROR:No Results found"
			End If
			WScript.Quit
			resultSets = ""
		End If
	 'Load Dictionary Object
	 For i = 0 To (UBound(LoadResultsArray)-1) Step 2
	 	If bolVERBOSE Then
			WScript.Echo LoadResultsArray(i)
		End If
		
	 	If InStr(LoadResultsArray(i),"/") > 0 then
	 		objDictResults.Add LoadResultsArray(i), LoadResultsArray(i+1)
	 	End if
	 Next
	'Sanity Check
	If bolVERBOSE Then
		For Each key In objDictResults.Keys
			If bolDEBUG Then
				WScript.Echo key & vbTab & objDictResults.Item(key)
			Else
				WScript.Echo key
			End If
		Next
	End If
	'======================================================
	
End Function 'parseResult_SetsField
'==================================================================================
' Function ReadPassword()
'
' Dependancies on other Subs or Functions
'	1)none
'
'Comments:
' 	Fundtion to get password testing 
'==================================================================================
Function ReadPassword()
	strMessage = "ENTER your password."
	Wscript.StdOut.Write strMessage
	'Do While Input <> Chr(13)
		Do While Not WScript.StdIn.AtEndOfLine
   			Input = WScript.StdIn.Read(1)
		Loop
		If Input <> Chr(13) then
			strPassword = strPassword + Input
		End if
	'Loop
	WScript.Echo strPassword
	ReadPassword = strPassword
End Function 'ReadPassword
'==================================================================================
' Sub subWriteEvent (intType, strDescription)
'
' Dependancies on Other Subs and Functions
'	1)None
'
' Comments:
'	This function allows the script to write to the system Event log.
' 0 = SUCCESS
' 1 = Error
' 2 = WARNING
' 4 = INFORMATION
' 8 = AUDIT_SUCCESS
' 16 = AUDIT_FAILURE
'==================================================================================
Sub subWriteEvent (intType, strDescription)
	  On Error Resume Next 'Turn on error handling for this sub
	  'Writes an event to the event log
	  strDescription = "Script Information Returned for: " & wScript.ScriptName & vbNewline & _
	  		strDescription
	  If bolDebug Then
	  	WScript.Echo "Attempting to write: " & strDescription
	  End If
	  WshShell.LogEvent intType, strDescription
	  strDescription = ""
	  intType = ""
	  On Error GoTo 0 'Turn error handling off
End Sub 'Write Event
'==================================================================================
' Function GetPassword(myHeader,myPrompt)
'
' Dependancies on Other Subs and Functions
'	1)None
'
' Comments:
'	This function masks the password
'==================================================================================
Function GetPassword(myHeader,myPrompt)
    Dim objIE
    ' Create an IE object
    Set objShell = WScript.CreateObject("WScript.Shell")
	Set objIE = CreateObject( "InternetExplorer.Application" )
    ' specify some of the IE window's settings
    objIE.Navigate "about:blank"
    objIE.Document.title = "Password for " & myHeader
    objIE.ToolBar        = False
    objIE.Resizable      = False
    objIE.StatusBar      = False
    objIE.Width          = 320
    objIE.Height         = 180
    ' Center the dialog window on the screen
    With objIE.Document.parentWindow.screen
        objIE.Left = (.availWidth  - objIE.Width ) \ 2
        objIE.Top  = (.availHeight - objIE.Height) \ 2
    End With
    ' Wait till IE is ready
    Do While objIE.Busy
        WScript.Sleep 200
    Loop
    ' Insert the HTML code to prompt for a password
    objIE.Document.body.innerHTML = "<div align=""center""><p>" & myPrompt _
                                  & "</p><p><input type=""password"" size=""20"" " _
                                  & "id=""Password""></p><p><input type=" _
                                  & """hidden"" id=""OK"" name=""OK"" value=""0"">" _
                                  & "<input type=""submit"" value="" OK "" " _
                                  & "onclick=""VBScript:OK.value=1""></p></div>"
    ' Hide the scrollbars
    objIE.Document.body.style.overflow = "auto"
    ' Make the window visible
    objIE.Visible = True
    'Put window on top applys to explorer
    objIE.top = True
    'Set as the active application
    objShell.AppActivate objIE
    ' Set focus on password input field
    objIE.Document.all.Password.focus

    ' Wait till the OK button has been clicked
    On Error Resume Next
    Do While objIE.Document.all.OK.value = 0 
        WScript.Sleep 200
        If Err Then    'user clicked red X (or alt-F4) to close IE window
            IELogin = Array( "", "" )
            objIE.Quit
            Set objIE = Nothing
            Exit Function
        End if
    Loop
    On Error Goto 0

    ' Read the password from the dialog window
    GetPassword = objIE.Document.all.Password.value

    ' Close and release the object
    objIE.Quit
    Set objIE = Nothing
End Function 'GetPassword
'==================================================================================
' SPCopy(SPURL, InFileName, InPath)
'
' Dependancies on Other Subs and Functions
'	1)None
'
' Comments:
'	This function uploads a file to sharepoint
'==================================================================================
Function SPCopy(SPURL, InFileName, InPath)
	On Error Resume next
    Set oFS = CreateObject("Scripting.FileSystemObject")

    SPSource = InPath & "\" & InFileName
	If InStr(SPSource,":") Then
    	SPSource = Replace(SPSource,"\\","\")
    End If
    
    If SPURL = "" Then
    	If not bolSILENT then
        	WScript.Echo "File: " & InFileName & " is not a valid file." & vbCrLf & _
            	"Not sure where to upload it to SharePoint... " & vbCrLf & VbCrLf
        End If
        Exit Function
    End If

    SPURL = SPURL & "\"
    DestFile = SPURL & "\" & InFileName
    If InStr(DestFile,":") Then
    	DestFile = Replace(DestFile,"\\","\")
    End If
    
    If InStr(SPURL,":") Then
    	SPURL = Replace(SPURL,"\\","\")
    End If
    

    ' copy the file(s) to the sharepoint path if its not already there
    Err.Clear
    If Not oFS.FileExists(DestFile) Then
        oFS.CopyFile SPSource, SPURL, True
    End If
    If Err.Number <> 0 Then
    	Call subWriteEvent(1,"Copying " & SPSource & " to " & SPURL & " | " &  Err.Description)
    End If
    On Error GoTo 0
end Function 'SPCopy
'==================================================================================
' Function CheckToSeeWithinSubnet(IPCheck,Subnet,MASK)
'		   CheckToSeeWithinSubnet(192.168.0.123,192.168.0.0,24)
' Dependancies on Other Subs and Functions
'	1)None
'
' Comments:
'	This function Checks a IP to see if within a specified subnet
'==================================================================================
Function CheckToSeeWithinSubnet(IPCheck,Subnet,MASK)
	CheckToSeeWithinSubnet = True
	'Build mask
	sMask = ""
	For i = 1 To 32
		If i <= MASK then
			sMask = sMask & "1"
		Else
			sMask =sMask & "0"
		End If
	Next
	If bolDebug then
		WScript.Echo sMask
	End if
	
	'build Mask Array
	arrMask = Array("0","0","0","0")
	
	For i = 0 To 3
		strNewOctet = Left(sMask,8)
		sMask = Right(sMask,(24-(i*8)))
		arrMask(i) = strNewOctet
	Next

	If bolDebug then	
	 	For Each octet In arrMask
	 		WScript.Echo octet
	 	Next
	End if
	
	'Separate and build octets to check
	arrIPCheck = Split(IPCheck,".")
	If ubound(arrIPCheck) <> 3 Then
		CheckToSeeWithinSubnet = False
		Exit Function
	End If
	If bolDebug Then
		For Each octet In arrIPCheck
	 		WScript.Echo octet
	 	Next
	End If 

	arrSubnet2 = Split(Subnet,".")
	If UBound(arrSubnet2)<> 3 Then
		CheckToSeeWithinSubnet = False
		Exit Function
	End If
	If bolDebug Then
	 	For Each octet In arrSubnet2
	 		WScript.Echo octet
	 	Next
	End if
	
	'Check each octet
	For check = 0 To 3
		'if Not CheckMask(ConvertBinaryToString(cint(arrIPCheck(check)),8),ConvertBinaryToString(cint(arrSubnet(check)),8),arrMask(check)) Then
		if Not CheckMask(ConvertBinaryToString(arrIPCheck(check),8),ConvertBinaryToString(arrSubnet(check),8),arrMask(check)) Then
			CheckToSeeWithinSubnet = False
			Exit Function
		End If		
	Next	
	
End Function
'==================================================================================
' Function CheckMask(bCheck,bSubnet,bMask)
'
' Dependancies on other Subs or Functions
'	1)none
'
'Comments:
' 	Checks string values
'==================================================================================
Function CheckMask(bCheck,bSubnet,bMask)
	'WScript.Echo bcheck & "," & bSubnet & "," & bmask
	result = True
	For i = 1 To 8
			If Mid(bMask,i,1)<>0 then
				If Mid(bCheck,i,1) <> Mid(bSubnet,i,1) Then
					result = False
				End If
			End If
	'		WScript.Echo Mid(bCheck,i,1) & "," & Mid(bSubnet,i,1) & "," & Mid(bMask,i,1) & "M " & result
	Next
	CheckMask = result
End Function
'==================================================================================
'Function ConvertBinaryToString(binary,length)
'
' Dependancies on other Subs or Functions
'	1)none
'
'Comments:
' 	Converts binary value to a string
'==================================================================================
Function ConvertBinaryToString(binary,length)
	strReturn = ""
	intbuffer = length - Len(binary)
	For i = 1 To intbuffer
		strReturn = strReturn & "0"
	Next
	strReturn = strReturn & binary
	ConvertBinaryToString = strReturn
End Function
'==================================================================================
' Function ConvertDecimalToBinary(intDec)
'
' Dependancies on other Subs or Functions
'	1)none
'
'Comments:
' 	Converts Decimal value to Binary
'==================================================================================
Function ConvertDecimalToBinary(intDec)
	Const maxpower = 8   ' Maximum number of binary digits supported.
    
    If intDec > 2 ^ maxpower Then
         ConvertDecimalToBinary = "error"
         Exit Function
    End If
	
    For i = maxpower To 0 Step -1
       If intDec And (2 ^ i) Then   ' Use the logical "AND" operator.
          bin = bin + "1"
       Else
          bin = bin + "0"
       End If
    Next
    ConvertDecimalToBinary = right(bin,8)  ' The bin string contains the binary number.
End Function
'==================================================================================
' sub ShowHelp()
'
' Dependancies on other Subs or Functions
'	1)none
'
'Comments:
' 	Help file for the user
'==================================================================================
sub ShowHelp()
	WScript.Echo "===================================================================="
	WScript.Echo "==               GatherTitaniumClientSubnets.vbs                  =="
	WScript.Echo "===================================================================="
	WScript.Echo "This script Queries Tanium to find subnets that may be missing"
	WScript.Echo "out of the current SeparatedSubnets file. Then creates a new file"
	Wscript.echo
	Wscript.echo "You need to specify an a text file to use"
	Wscript.echo "cscript GatherTitaniumClientSubnets.vbs InputFileName [account] [password]"
	wscript.echo
	Wscript.echo "Example:"
	Wscript.echo "cscript GatherTitaniumClientSubnets.vbs SeparatedSubnets.txt a123456 VBRocks"
	wscript.echo 	
	WScript.Echo "===================================================================="
end Sub 'ShowHelp
