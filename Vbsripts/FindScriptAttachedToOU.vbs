'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: FindScriptAttachedToOU
'
' AUTHOR: Jim Borecky
' DATE  : 9/21/2007
'
' COMMENT: This script will take a local deposit of XML files created by
'		GPOReport----.vbs and find the scripts attached to the Group Policy
'		objects given a list of Group Policies in a txt file.
'
'==========================================================================

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

if wscript.arguments.count = 2 Then
	strInputFile = wscript.arguments.item(0)
	strXMLPath = WScript.Arguments.Item(1)
Else
	ShowHelp
	wscript.quit
End If

'Lets set up a File object so we can use files
Set objFSO = CreateObject("Scripting.FileSystemObject")
	
'----------------------------------------------------------------------------
'This File will contain the GPO's we want to look at
Set objInputFile = objfso.OpenTextFile(strInputFile,ForReading)
'This is where we are dumping the results
Set objOutputFile = objFSO.CreateTextFile(split(strInputFile,".")(0) & ".out.txt")

'Let loop through the file
Do While objInputFile.AtEndOfStream <> True
	strNew = objInputFile.ReadLine
	
	'Need to remove any illegal file characters to keep the script from puking
	strFindFile = strXMLPath & "\" & CheckIllegalFileCharacters(strNew) & ".xml"
	WScript.Echo strFindFile
	
	'Just in case we are making up names
	if objFSO.FileExists(strFindFile) Then
		'loads that inital XML object
		strXML = GetXMLFile(strFindFile)
		'Off to the races
		Call GetChildXML(strNew,strXML)
	End If
Loop


objInputFile.Close
objOutputFile.Close
'-------------------------------------------------------------------
' Function to get Info in XML formt
'-------------------------------------------------------------------
Function GetChildXML(GPO,XML)
		' This is designed to use in a recursive fashion
		Set xmlOld=CreateObject("Microsoft.XMLDOM")
		Dim xmlData
		Dim xmlScriptNode
		xmlOld.async="false"
		xmlOld.loadXML(XML)
		
		'Looping through all of the Nodes
		For each x In xmlOld.documentElement.childNodes
		
		  'Trying to find the node that we are looking for.
		  if x.basename = "Script" Then
		  	'And here we go, lets echo that baby out just for some feedback
		  	strScriptSetting = x.text
		  	WScript.Echo strScriptSetting
		  	
		  	'Write it to a file so that we can look at it later
		  	objOutputFile.WriteLine GPO & vbTab & strScriptSetting
		  End If
		  
		  'Checking to see if there are any child nodes that we need to dig into.
		  If x.ChildNodes.Length <> 0 then
			  xmlData = xmlOld.documentElement.selectSingleNode(x.basename).xml
			  
			  'Have to skip some errors since this object needs help.
			  On Error Resume Next
			  Err.Clear
			  
			  'If you miss this you'll get lost. Here we are calling the 
			  'function again to get a level down.
			  if x.hasChildNodes Then
			    Call GetChildXML(GPO,xmlData)
			  End If
			  
			  
			  If Err.Number <> 0 Then
			  'Uncomment the next line if you are looking for a bug.
			'  	WScript.Echo Err.Description
			  End If
		  End If
		Next
End Function
'-------------------------------------------------------------------
' Function to extract data from XML
'-------------------------------------------------------------------
Function GetXMLAttrib(xmlEmp,strAttrib)
	'Create an XML object to transfer to
	set xmlDoc=CreateObject("Microsoft.XMLDOM")
	xmlDoc.async="false"
	'Load the new XML object with the data from the old.
	xmlDoc.loadxml(xmlEmp)
	'Newstring will contain the data from the Node that we are selecting
	NewString = xmlDoc.documentElement.selectSingleNode(strAttrib).text
	'Just in case it has no data
	If NewString <> "" Then
		'getting rid of stuff that might make my script puke.
		NewString = Replace(NewString,"'","''")
		GetXMLAttrib = NewString
	End If
end Function
'-------------------------------------------------------------------
' Function to load XML file into the system
'-------------------------------------------------------------------
Function GetXMLFile(XML)
	'Create a XML object to store our data
	Set xmlOld=CreateObject("Microsoft.XMLDOM")
	xmlOld.async="false"
	'Load that File into the Object.
	xmlOld.load(XML)		
	'Return it back
	GetXMLFile = xmlOld.xml
End Function

'-------------------------------------------------------------------
' Find illegal File Characters
'-------------------------------------------------------------------
Function CheckIllegalFileCharacters(strFileName)
	'Just a simple find and replace
	strFilename = Replace(strFileName, "*", "")
	strFilename = Replace(strFileName, "?", "")
	strFilename = Replace(strFileName, ":", "")
	strFilename = Replace(strFileName, ">", "")
	strFilename = Replace(strFileName, "<", "")
	strFilename = Replace(strFileName, "|", "")
	strFilename = Replace(strFileName, "\\", "")
	strFilename = Replace(strFileName, "\", "")
	strFilename = Replace(strFileName, "/", "")
	CheckIllegalFileCharacters = strFilename
End Function 
'------------------------------------------------------------------
' Sub to tell the user the syntax
'------------------------------------------------------------------
Sub ShowHelp()
	'Ok so you need some answers.
	Wscript.echo
	Wscript.echo "You need to specify an a text file to use"
	Wscript.echo "cscript ScriptName.vbs ListOfPoliciesFile XMLPath"
	wscript.echo
	Wscript.echo "Example:"
	Wscript.echo "cscript ScriptName.vbs Input.txt C:\PolicyXML\DVCH"
	wscript.echo 
End sub	
