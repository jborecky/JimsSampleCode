// AdminCSE.h : main header file for the AdminCSE DLL
//

#pragma once

//GUID for classess
extern const GUID CLSID_MFCDll;

// {80FBBE20-4995-46ba-9569-AB0F81F9132C}
static const GUID CLSID_MFCDll = 
{ 0x80fbbe20, 0x4995, 0x46ba, { 0x95, 0x69, 0xab, 0xf, 0x81, 0xf9, 0x13, 0x2c } };

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols

#include <windows.h>
#include <userenv.h>
#using <System.dll>
#include "stdio.h"

#include <tchar.h>
#using <mscorlib.dll>
#include<string>
#include "stdafx.h"
#include "strsafe.h"

// CAdminCSEApp
// See AdminCSE.cpp for the implementation of this class
//

using namespace System;
using namespace System::Security;
using namespace System::Security::Principal;

#define TOTALBYTES    49152
#define BYTEINCREMENT 4096

class CAdminCSEApp : public CWinApp
{
public:
	CAdminCSEApp();
	char* GetMyDWORD(char* strPath, char* strValue);
	//void funProcessTasks(HANDLE hUserToken, HKEY rootKey);
private:
	//void WtoC(char* Dest, TCHAR* Source, int SourceSize);

// Overrides
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DWORD CALLBACK ProcessGroupPolicy( DWORD dwFlags, HANDLE hToken, HKEY hKeyRoot, 
		PGROUP_POLICY_OBJECT pDeletedGPOList, PGROUP_POLICY_OBJECT pChangedGPOList, 
		ASYNCCOMPLETIONHANDLE pHandle, BOOL *pbAbort, PFNSTATUSMESSAGECALLBACK pStatusCallback);

	DWORD CALLBACK ProcessGroupPolicyEx(DWORD dwFlags, HANDLE hToken, HKEY hKeyRoot, 
		PGROUP_POLICY_OBJECT pDeletedGPOList, PGROUP_POLICY_OBJECT pChangedGPOList, 
		ASYNCCOMPLETIONHANDLE pHandle, BOOL *pbAbort, PFNSTATUSMESSAGECALLBACK pStatusCallback, 
		IWbemServices *pWbemServices, HRESULT *pRsopStatus);

	DECLARE_MESSAGE_MAP()

	// //This Function processes the tasks in the HKLM\Software\Policies\CSETemplate\Tasks key
	//int funRunTask(System::String^ strCommandLine,System::String^ strCommandLineArgs, HANDLE hUserToken);
	//int funRunTask(LPCTSTR commandLine, LPTSTR commandLineArgs, HANDLE hUserToken);
	std::wstring ParseError(LPTSTR lpszFunction);
	size_t ExecuteProcess(HANDLE hSecToken,std::wstring FullPathToExe, std::wstring Parameters, size_t SecondsToWait);
};

class CMyExtraClass : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CMyExtraClass, &CLSID_MFCDll>
{
public:
	DECLARE_REGISTRY_CLSID();

	BEGIN_COM_MAP(CMyExtraClass)
	END_COM_MAP()
};

class EventLogging
{
public:
	void LogEvent(System::String ^sSource, System::String ^sLog, System::String ^sEvent);
}; 