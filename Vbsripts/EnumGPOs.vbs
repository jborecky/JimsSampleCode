'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: EnumGPOs.vbs
'
' AUTHOR: Jim Borecky
' DATE  : 9/11/2007
'
' COMMENT: Policy to create XML dump of the domain so that I can later 
'			parse the XML file to find the info that I'm looking for.
'
'
'	9/21/2007 I think this method has a memory leak in the method used. I 
'		have sent it off to our Microsoft representative Dewey Stevens. 
'		Hopefully he will look at it and get back to me.
'
'==========================================================================

LOBPath = "OU=WBG"
timeStart = Now()
' Create the GPM and GPMConstants objects,
'then connect to the domain.

Set GPMC = CreateObject("GPMgmt.GPM")
Set Constants = GPMC.GetConstants()

'On Error Resume next
Set GPMCDomain = GPMC.GetDomain("Rougeone.borecky.net", "DomainControllername",Constants.UseAnyDC)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutputFile = objFSO.CreateTextFile("out.txt")
Set objLOBFile = objFSO.CreateTextFile("LOB_out.txt")
Set objGPOFile = objFSO.CreateTextFile("GPO_Only.txt")


'Perform a search with a empty criteria to gather up all of the 
'GPO's
set objGPMSearchCriteria2 = GPMC.CreateSearchCriteria
Set myGPOs = GPMCDomain.SearchGPOs(objGPMSearchCriteria2)
Set objGPMSearchCriteria2 = Nothing

'Loop through the GPO's
WScript.Echo myGPOs.Count
For Each GPO In myGPOs

	'Echo out where I'm at
	wscript.Echo VbCrLf & GPO.DisplayName & vbTab & GPO.ID

	If InStr(GPO.DisplayName,"DTC") Then
		'yank out the Illegal File Characters
		Call SearchLinks(GPO)
		'fileName = CheckIllegalFileCharacters(GPO.DisplayName)
		'objOutputFile.WriteLine fileName
	End If 
	filename = ""
Next

objOutputFile.Close()
objLOBFile.Close()
objGPOFile.Close()

WScript.Echo "Start:" & timeStart
WScript.Echo "End:  " & Now()

Function SearchLinks(gpmGPO)
	bolRecord = False
	strLinks = gpmGpo.DisplayName & vbNewLine
	strLOBLinks = gpmGpo.DisplayName & vbNewLine
	
	Set gpm = CreateObject("GPMGMT.GPM")
	Set gpmSearchCriteria = gpm.CreateSearchCriteria
	Set gpmConstants = gpm.GetConstants
	gpmSearchCriteria.Add gpmConstants.SearchPropertySOMLinks, _
	gpmConstants.SearchOpContains, gpmGPO
	Set colDOUSOMs = GPMCDomain.SearchSOMs(gpmSearchCriteria)
	
	For Each gpmSOM in colDOUSOMs
		strLinks = strLinks & vbTab & gpmSOM.Path & vbNewLine
		WScript.Echo gpmSOM.Path
		If InStr(gpmSOM.Path, LOBPath) Then
				strLOBLinks = strLOBLinks & vbTab & gpmSOM.Path & vbNewLine
				bolRecord = True
		End If
	Next
	
	If bolRecord Then
		objOutputFile.WriteLine strLinks
		objLOBFile.WriteLine strLOBLinks
		objGPOFile.WriteLine gpmGPO.DisplayName
	End If 

End Function
'-------------------------------------------------------------------
' Find illegal File Characters
'-------------------------------------------------------------------
Function CheckIllegalFileCharacters(strFileName)
	'Simple Find and Replace
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
