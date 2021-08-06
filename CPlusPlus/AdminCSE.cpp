//=========================================================================================
// AdminCSE.cpp : Defines the initialization routines for the DLL.
// CREATED BY   : Jim Borecky
// CREATED DATE : 6/11/2014
//
// COMMENTS:
//		Creating a CSE Template.
//=========================================================================================

#include "stdafx.h"
#include "AdminCSE.h"
#include <activeds.h>
#include "strsafe.h"
#include "iostream"
#include "windows.h"

//.NET Namespaces
using namespace System;
using namespace System::Diagnostics;
using namespace System::Security;
using namespace System::Security::Principal;
//END .Net Namespaces

using namespace std;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CAdminCSEApp
CMyExtModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
	OBJECT_ENTRY(CLSID_MFCDll, CMyExtraClass)
END_OBJECT_MAP()

BEGIN_MESSAGE_MAP(CAdminCSEApp, CWinApp)
END_MESSAGE_MAP()


// CAdminCSEApp construction

CAdminCSEApp::CAdminCSEApp()
{
	BOOL InitInstance(void);
	int ExitInstance(void);
}


// The one and only CAdminCSEApp object

CAdminCSEApp theApp;

//=========================================================================================
// BOOL CMFCDllApp::InitInstance()
//
// Dependancies on other Subs or Functions
//		1)
//
// COMMENTS:
//		Function that gets called when the DLL gets called. This is needed to register the
//		DLL.
//=========================================================================================
BOOL CAdminCSEApp::InitInstance()
{
	_Module.Init(ObjectMap, m_hInstance);
	return CWinApp::InitInstance();
}
//=========================================================================================
// BOOL CMFCDllApp::InitInstance()
//
// Dependancies on other Subs or Functions
//		1)
//
// COMMENTS:
//		Function that gets called when the DLL exits. This is just clean up.
//=========================================================================================
int CAdminCSEApp::ExitInstance()
{
	_Module.Term();
	return CWinApp::ExitInstance();
}

//=========================================================================================
// Functions that install and uninstall the DLL into the registry.
//
// Dependancies on other Subs or Functions
//		1)
//
// COMMENTS:
//		Function that gets called when the DLL exits. This is just clean up. Most of this
//		code is aquired from the SDK.
//=========================================================================================
///////////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE
//STDAPI DllCanUnloadNow(void)
//{
//	AFX_MANAGE_STATE(AfxGetStaticModuleState());
//	return (AfxDllCanUnloadNow()==S_OK && _Module.GetLockCount()==0) ? S_OK : S_FALSE;
//}
//////////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type
STDAPI DLLGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _Module.GetClassObject(rclsid, riid, ppv);
}
/////////////////////////////////////////////////////////////////////////////////
// DLLRegisterServer - Adds entries to the system registry

BOOL x = TRUE;
STDAPI DllRegisterServer(void)
{  
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return _Module.RegisterServer(FALSE);
}

/////////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entires from the system registry

STDAPI DllUnregisterServer(void)
{
	_Module.UnregisterServer();
	return S_OK;
}




//=========================================================================================
// DWORD CALLBACK ProcessGroupPolicy( DWORD dwFlags, HANDLE hToken, HKEY hKeyRoot, PGROUP_POLICY_OBJECT pDeletedGPOList, PGROUP_POLICY_OBJECT pChangedGPOList, ASYNCCOMPLETIONHANDLE pHandle, BOOL *pbAbort, PFNSTATUSMESSAGECALLBACK pStatusCallback)
// CREATED BY   : Jim Borecky
// CREATED DATE : Unknown
// Dependancies on other Subs or Functions
//		1)
//
// COMMENTS:
//		Function that gets called by the group policy to apply the policy
//=========================================================================================
DWORD CALLBACK ProcessGroupPolicy( DWORD dwFlags,
								   HANDLE hToken,
								   HKEY hKeyRoot,
								   PGROUP_POLICY_OBJECT pDeletedGPOList,
								   PGROUP_POLICY_OBJECT pChangedGPOList,
								   ASYNCCOMPLETIONHANDLE pHandle,
								   BOOL *pbAbort,
								   PFNSTATUSMESSAGECALLBACK pStatusCallback)
{
	//http://msdn.microsoft.com/en-us/library/aa374377(VS.85).aspx
	//Add this to Bill's Service
	//BOOL RefreshPolicy( __in  BOOL bMachine);
	//http://msdn.microsoft.com/en-us/library/aa374398(VS.85).aspx


	//Setup for recording data into the event log
	String ^ssSource;
	String ^ssLog;
	String ^ssEvent;
	EventLogging EvtLog;

	ssSource = gcnew String("CSETemplate");
	ssLog = gcnew String("Application");
	
   //Create object to store Group Policy collections lists in
   PGROUP_POLICY_OBJECT pCurGPO;
   
   // Process all deleted GPOs in pDeletedGPOList.
   for( pCurGPO = pDeletedGPOList; pCurGPO; pCurGPO = pCurGPO->pNext )
   {
       if( *pbAbort )
       {
           // Abort.
           break;
       }
       //Place code here.
	   ssEvent = gcnew String("A Policy was deleted:");
	   ssEvent = ssEvent + gcnew String(pCurGPO->szGPOName);
	   EvtLog.LogEvent(ssSource,ssLog,ssEvent);
   }

   // Process all changed GPOs in pChangedGPOList.
   for( pCurGPO = pChangedGPOList; pCurGPO; pCurGPO = pCurGPO->pNext )
   {
       if( *pbAbort )
       {
           // Abort.
           break;
       }
       //Place code here
	   ssEvent = gcnew String("A Policy Changed:");
	   ssEvent = ssEvent + gcnew String(pCurGPO->szGPOName);
	   EvtLog.LogEvent(ssSource,ssLog,ssEvent);
   }

   //Code here processes everytime!
   Beep(523,500); // 523 hertz (C5) for 500 milliseconds  
	 
   //Get the reg key HKEY_LOCAL_MACHINE information that has been deposited by the ADM template.
   char Path[] = "SOFTWARE\\Policies\\CSETemplate";
   //CrossDomain		"This allows entries from domains other than the domain from which the machine resides. But will also allow invalid entries as the extension does not verify the entry."
   char CrossDomain[] = "CrossDomain";
   //Enabled			"Configures the machine to accept the entries from the Local Group Management Page located in Active Directory Users and Computers. This controls the built-in groups."
   char Enabled[] = "Enabled";
   //GrandFather		"Removes existing User Accounts that are not in the exclusion list. Default value is to be left behind."
   char GrandFather[] = "GrandFather";
   //LocalGroup		"Control local groups that are nested into the built-in groups.This will remove any local groups that are nestes into the Built-in groups."
   char LocalGroup[] = "LocalGroup";
   char Task[] = "Tasks";

	//HKEY_LOCAL_MACHINE\SOFTWARE\Policies\CSETemplate\Tasks
	//TASK_ID1 - All entries below this key will run the tasks listed in the policy."
	ssEvent = gcnew String("Entering the DLL");
	EvtLog.LogEvent(ssSource,ssLog,ssEvent);

	//theApp.funProcessTasks(hToken, hKeyRoot);

	ssEvent = gcnew String("Group Policy processed the function ProcessGroupPolicy");
	EvtLog.LogEvent(ssSource,ssLog,ssEvent);

	//Cleanup 
	delete ssSource;
	delete ssLog;
	delete ssEvent;
//	delete[] EnabledValue;

	//This ensures that the CSE will be called again when Group Policy refreshes.
    return(ERROR_OVERRIDE_NOCHANGES);
}
//=========================================================================================
// DWORD CALLBACK ProcessGroupPolicyEx( DWORD dwFlags, HANDLE hToken, HKEY hKeyRoot, PGROUP_POLICY_OBJECT pDeletedGPOList, PGROUP_POLICY_OBJECT pChangedGPOList, ASYNCCOMPLETIONHANDLE pHandle, BOOL *pbAbort, PFNSTATUSMESSAGECALLBACK pStatusCallback)
//
// Dependancies on other Subs or Functions
//		1)
//
// COMMENTS:
//		Function that gets called by the group policy to create a RSOP result.
//=========================================================================================
DWORD CALLBACK ProcessGroupPolicyEx(DWORD dwFlags, HANDLE hToken, HKEY hKeyRoot, PGROUP_POLICY_OBJECT pDeletedGPOList, PGROUP_POLICY_OBJECT pChangedGPOList, ASYNCCOMPLETIONHANDLE pHandle, BOOL *pbAbort, PFNSTATUSMESSAGECALLBACK pStatusCallback, IWbemServices *pWbemServices, HRESULT *pRsopStatus)
{
   //AfxMessageBox(_T("Group Policy processed ProcessGroupPolicyEx"),MB_OK);
   //AfxAbort();
	
   return(ERROR_SUCCESS);
}

//=========================================================================================
// EventLogging::LogEvent()
// CREATED BY  : Jim Borecky
// DATE CREATED: Unknown
// Dependancies on other Subs or Functions
//		1)
//
// COMMENTS:
//		Defines the initialization routines for the DLL.
//=========================================================================================
void EventLogging::LogEvent(String ^sSource, String ^sLog, String ^sEvent)
{

	if(!EventLog::SourceExists(sSource))
		EventLog::CreateEventSource(sSource,sLog);
			EventLog::WriteEntry(sSource,sEvent);
			return;
}

//=========================================================================================
// char* CAdminCSEApp::GetMyDWORD(char* strPath, char* strValue)
//
// Dependancies on other Subs or Functions
//		1)
//
// COMMENTS:
//		Defines the initialization routines for the DLL.
//=========================================================================================

char* CAdminCSEApp::GetMyDWORD(char* strPath, char* strValue)
{
	DWORD dwVersion;
    HKEY hMyKey;
    LONG returnStatus;
	DWORD dwType=REG_DWORD;
    DWORD dwSize=sizeof(DWORD);

	// Convert to a wchar_t*
    size_t origsize = strlen(strPath) + 1;
    const size_t newsize = 100;
    size_t convertedChars = 0;
    wchar_t wcstrPath[newsize];
    mbstowcs_s(&convertedChars, wcstrPath, origsize, strPath, _TRUNCATE);

	size_t origValuesize = strlen(strValue) + 1;
    const size_t newValuesize = 100;
    wchar_t wcstrValue[newValuesize];
    mbstowcs_s(&convertedChars, wcstrValue, origValuesize, strValue, _TRUNCATE);
   
    returnStatus = RegOpenKeyEx(HKEY_LOCAL_MACHINE, wcstrPath, 0L, KEY_READ, &hMyKey);

    if (returnStatus == ERROR_SUCCESS)
    {
        returnStatus = RegQueryValueEx(hMyKey, wcstrValue, NULL, &dwType,(LPBYTE)&dwVersion, &dwSize);
        if (returnStatus == ERROR_SUCCESS)
        {
			char buffer[100];
			char* MyChar;
			MyChar = _ultoa(dwVersion,buffer,16);
            return MyChar;
			delete[] MyChar;
        }
		else
		{
			return "FAILED_Value";
		}
    }
	else
	{
	  return "FAILED_Key";
	}

    RegCloseKey(hMyKey);
	delete[] strPath;
	delete[] strValue;
	delete[] hMyKey;

}

