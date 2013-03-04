// chgoffice.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "chgoffice.h"
#include <iostream>
#include <wchar.h>
#include <Psapi.h>


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// The one and only application object

CWinApp theApp;

using namespace std;

BOOL CommonCopyFile(CString SourceFileName, CString DestFileName)
{
	CFile sourceFile ;
	CFile destFile ;
	CFileException ex;
	if (!sourceFile.Open((LPCTSTR)SourceFileName,CFile::modeRead | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		CString ErrorMsg = _T("打开文件：") ;
		ErrorMsg += SourceFileName ;
		ErrorMsg += _T("失败。\n错误信息为：\n") ;
		ErrorMsg += szError ;
		AfxMessageBox(ErrorMsg);
		return FALSE ;
	}
	else
	{
		if (!destFile.Open((LPCTSTR)DestFileName, CFile::modeWrite | CFile::shareExclusive | CFile::modeCreate, &ex))
		{
			TCHAR szError[1024];
			ex.GetErrorMessage(szError, 1024);
			CString ErrorMsg = _T("创建文件：") ;
			ErrorMsg += DestFileName ;
			ErrorMsg += _T("失败。\n错误信息为：\n") ;
			ErrorMsg += szError ;
			AfxMessageBox(ErrorMsg);
			sourceFile.Close();
			return FALSE ;
		}
		BYTE buffer[4096];
		DWORD dwRead;
		do
		{
			dwRead = sourceFile.Read(buffer, 4096);
			destFile.Write(buffer, dwRead);
		}
		while (dwRead > 0);   
		destFile.Close();
		sourceFile.Close();
	}
	return TRUE ;
}
const TCHAR WORDFILE_2003[] = _T("Normal.dot");
const TCHAR EXCELFILE_2003[] = _T("Book.xlt");
const TCHAR EXCELFILE_2003_1[] = _T("Sheet.xlt");
const TCHAR TEMPLATEDIR[] = _T("Microsoft\\Templates\\");
const TCHAR EXL_DIR[] = _T("Microsoft\\Excel\\XLSTART\\");

BOOL Copy2003Files()
{
	TCHAR MyDir[_MAX_PATH];  
	SHGetSpecialFolderPath(NULL,MyDir,CSIDL_APPDATA,0);
	CString AppDataDir = MyDir;
	AppDataDir.TrimRight('\\');
	AppDataDir += _T("\\");
	::GetModuleFileName(NULL, MyDir, _MAX_PATH);
	CString CurrentDir = MyDir;
	CurrentDir = CurrentDir.Left(CurrentDir.ReverseFind('\\'));
	CurrentDir.TrimRight('\\');
	CurrentDir += _T("\\Office2003\\");	
	BOOL bRes = CommonCopyFile(CString(CurrentDir + WORDFILE_2003),
			CString(AppDataDir + TEMPLATEDIR + WORDFILE_2003));
	if(!bRes) {
		wcout << _T("复制文件失败") << endl;
		return FALSE;
	}
	bRes = CommonCopyFile(CString(CurrentDir + EXCELFILE_2003),
		CString(AppDataDir + EXL_DIR + EXCELFILE_2003));
	if(!bRes) {
		wcout << _T("复制文件失败") << endl;
		return FALSE;
	}
	bRes = CommonCopyFile(CString(CurrentDir + EXCELFILE_2003_1),
		CString(AppDataDir + EXL_DIR + EXCELFILE_2003_1));
	if(!bRes) {
		wcout << _T("复制文件失败") << endl;
		return FALSE;
	}
	return TRUE;
}

const TCHAR WORDFILE_2007[] = _T("Normal.dotm");
const TCHAR EXCELFILE_2007[] = _T("Book.xltx");
const TCHAR EXCELFILE_2007_1[] = _T("Sheet.xltx");
BOOL Copy2007Files()
{
	TCHAR MyDir[_MAX_PATH];  
	SHGetSpecialFolderPath(NULL,MyDir,CSIDL_APPDATA,0);
	CString AppDataDir = MyDir;
	AppDataDir.TrimRight('\\');
	AppDataDir += _T("\\");
	::GetModuleFileName(NULL, MyDir, _MAX_PATH);
	CString CurrentDir = MyDir;
	CurrentDir = CurrentDir.Left(CurrentDir.ReverseFind('\\'));
	CurrentDir.TrimRight('\\');
	CurrentDir += _T("\\Office2007\\");	
	BOOL bRes = CommonCopyFile(CString(CurrentDir + WORDFILE_2007),
		CString(AppDataDir + TEMPLATEDIR + WORDFILE_2007));
	if(!bRes) {
		cout << "复制文件失败" << endl;
		return FALSE;
	}
	bRes = CommonCopyFile(CString(CurrentDir + EXCELFILE_2007),
		CString(AppDataDir + EXL_DIR + EXCELFILE_2007));
	if(!bRes) {
		cout << "复制文件失败" << endl;
		return FALSE;
	}
	bRes = CommonCopyFile(CString(CurrentDir + EXCELFILE_2007_1),
		CString(AppDataDir + EXL_DIR + EXCELFILE_2007_1));
	if(!bRes) {
		cout << "复制文件失败" << endl;
		return FALSE;
	}
	return TRUE;
}


DWORD FindProcess(TCHAR *strProcessName)
{
	DWORD aProcesses[1024], cbNeeded, cbMNeeded;
	HMODULE hMods[1024];
	HANDLE hProcess;
	TCHAR szProcessName[MAX_PATH];
	if ( !EnumProcesses( aProcesses, sizeof(aProcesses), &cbNeeded ) )  return 0;
	for(int i=0; i< (int) (cbNeeded / sizeof(DWORD)); i++)
	{
		//_tprintf(_T("%d\t"), aProcesses[i]);
		hProcess = OpenProcess(  PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, aProcesses[i]);
		EnumProcessModules(hProcess, hMods, sizeof(hMods), &cbMNeeded);
		GetModuleFileNameEx( hProcess, hMods[0], szProcessName,sizeof(szProcessName));
		CString strP = szProcessName;
		strP.MakeUpper();
		int nStart = strP.Find(strProcessName);
		if(nStart != -1)
		{
			//_tprintf(_T("%s;"), szProcessName);
			return(aProcesses[i]);
		}
		//_tprintf(_T("\n"));
	}

	return 0;
}

int KillProcess(DWORD nProcessID)
{
	HANDLE hProcessHandle;  
	hProcessHandle = ::OpenProcess( PROCESS_TERMINATE, FALSE, nProcessID );
	return ::TerminateProcess( hProcessHandle, 4 );
}

int CreateNewProcess(LPCTSTR pszExeName)
{
	PROCESS_INFORMATION piProcInfoGPS;
	STARTUPINFO siStartupInfo;
	ZeroMemory( &siStartupInfo, sizeof(siStartupInfo) );
	siStartupInfo.cb = sizeof(siStartupInfo);
	return ::CreateProcess( (LPTSTR)pszExeName, NULL, NULL, NULL, false, CREATE_DEFAULT_ERROR_MODE, NULL, NULL, &siStartupInfo,&piProcInfoGPS );
}

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int nRetCode = 0;

	// initialize MFC and print and error on failure
	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))
	{
		// TODO: change error code to suit your needs
		_tprintf(_T("Fatal Error: MFC initialization failed\n"));
		nRetCode = 1;
	}
	else
	{
		// TODO: code your application's behavior here.
		cout << "开始处理..." << endl;
		OfficeVersion::eOfficeVersion eVersion = OfficeVersion::GetOfficeVersion();
		CString strOut = OfficeVersion::GetVersionAsString(eVersion);
		wcout << (LPCTSTR)strOut << endl;
		BOOL bRes = FALSE;
		DWORD nProcessID = 0;
		nProcessID = FindProcess(_T("WINWORD.EXE"));
		if(nProcessID){
			KillProcess(nProcessID);
			Sleep(1000);
		}
		nProcessID = FindProcess(_T("EXCEL.EXE"));
		if(nProcessID) {
			KillProcess(nProcessID);
			Sleep(1000);
		}
		switch(eVersion) {
			case OfficeVersion::eOfficeVersion_2003: 
								bRes = Copy2003Files();							
								break;
			case OfficeVersion::eOfficeVersion_2007:
								bRes = Copy2007Files();
								break;
			default:
				cout << "找不到Office版本" << endl;
		}
	
		if(bRes) {
			cout << "修改成功, 启动Word检测是否正确" << endl;
			CString sApp = OfficeVersion::GetProgID(OfficeVersion::eOfficeApp_Word);
			BSTR bApp = sApp.AllocSysString();
			TCHAR szPath[255];
			OfficeVersion::GetPath(bApp, szPath, 255);
			CString strPath = szPath;
			CreateNewProcess((LPCTSTR)strPath);
		}
		cout << "按任意键退出" << endl;
		_getwch();
	}

	return nRetCode;
}
