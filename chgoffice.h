#pragma once

#include "resource.h"

class OfficeVersion
{
public:
	enum eOfficeVersion
	{
		eOfficeVersion_Unknown, // error return value
		eOfficeVersion_95,
		eOfficeVersion_97,
		eOfficeVersion_2000,
		eOfficeVersion_XP,   // XP = 2002 + marketing
		eOfficeVersion_2003,
		eOfficeVersion_2007,
	};

	enum eOfficeApp // in case you are looking for a particular app
	{
		eOfficeApp_Word,
		eOfficeApp_Excel,
		eOfficeApp_Outlook,
		eOfficeApp_Access,
		eOfficeApp_PowerPoint,
	};

	static CString GetVersionAsString(const eOfficeVersion officeVersion)
	{
		switch(officeVersion) {
case eOfficeVersion_Unknown: { return _T("Not found");       }break;
case eOfficeVersion_95:      { return _T("Office 95");       }break;
case eOfficeVersion_97:      { return _T("Office 97");       }break;
case eOfficeVersion_2000:    { return _T("Office 2000");     }break;
case eOfficeVersion_XP:      { return _T("Office XP");       }break;
case eOfficeVersion_2003:    { return _T("Office 2003");     }break;
case eOfficeVersion_2007:    { return _T("Office 2007");     }break;
default:                     { ASSERT(false); return _T(""); }break; // added another ???
		}
	}

	static CString GetApplicationAsString(const eOfficeApp officeApp)
	{
		switch(officeApp) {			
case eOfficeApp_Word:       { return _T("Word");            }break;
case eOfficeApp_Excel:      { return _T("Excel");           }break;
case eOfficeApp_Outlook:    { return _T("Outlook");         }break;
case eOfficeApp_Access:     { return _T("Access");          }break;
case eOfficeApp_PowerPoint: { return _T("Powerpoint");      }break;
default:                    { ASSERT(false); return _T(""); }break; // added another ???
		}
	}


	static CString GetProgID(const eOfficeApp officeApp)
	{
		// ProgIDs from http://support.microsoft.com/kb/240794/EN-US/
		switch(officeApp) {			
case eOfficeApp_Word:       { return _T("Word.Application");       }break;
case eOfficeApp_Excel:      { return _T("Excel.Application");      }break;
case eOfficeApp_Outlook:    { return _T("Outlook.Application");    }break;
case eOfficeApp_Access:     { return _T("Access.Application");     }break;
case eOfficeApp_PowerPoint: { return _T("Powerpoint.Application"); }break;
default:                    { ASSERT(false); return _T("");        }break; // added another ???
		}
	}

	static eOfficeVersion StringToVersion(const CString& versionString)
	{
		// mapping between the marketing version (e.g. 2003) and the behind-the-scenes version
		if(_T("7") == versionString){
			return eOfficeVersion_95;
		}else if(_T("8") == versionString){
			return eOfficeVersion_97;
		}else if(_T("9") == versionString){
			return eOfficeVersion_2000;
		}else if(_T("10") == versionString){
			return eOfficeVersion_XP;
		}else if(_T("11") == versionString){
			return eOfficeVersion_2003;
		}else if(_T("12") == versionString){
			return eOfficeVersion_2007;
		}else{
			return eOfficeVersion_Unknown; // added another ???
		}
	}

	static eOfficeVersion GetOfficeVersion()
	{
		// by default we use Word (and so on, down the list) as a proxy for "Office" 
		// (i.e. if word is there then "Office" is there)
		// if you want something more specific, then call GetApplicationVersion()

		static const eOfficeApp appsToCheck[] = {	
			eOfficeApp_Word,
			eOfficeApp_Excel,
			eOfficeApp_Outlook,
			eOfficeApp_Access,
			eOfficeApp_PowerPoint,
		};
		const int numItems( sizeof(appsToCheck) / sizeof(appsToCheck[0]) );		

		for(int i=0; i<numItems; ++i){
			const eOfficeVersion thisAppVersion( GetApplicationVersion(eOfficeApp_Word) );
			if(eOfficeVersion_Unknown != thisAppVersion){
				return thisAppVersion;
			}
		}

		return eOfficeVersion_Unknown; // probably nothing installed
	}

	static eOfficeVersion GetApplicationVersion(eOfficeApp appToCheck)
	{
		// some of this function is based on the code in the article at: http://support.microsoft.com/kb/q247985/
		const CString progID( GetProgID(appToCheck) );

		HKEY hKey( NULL);
		HKEY hKey1(NULL);

		if(ERROR_SUCCESS != ::RegOpenKeyEx(HKEY_CLASSES_ROOT, progID, 0, KEY_READ, &hKey) ){
			return eOfficeVersion_Unknown;
		}

		if(ERROR_SUCCESS != ::RegOpenKeyEx(hKey, _T("CurVer"), 0, KEY_READ, &hKey1)) {
			::RegCloseKey(hKey);
			return eOfficeVersion_Unknown;
		}

		// Get the Version information
		const int BUFFER_SIZE(255);
		ULONG cSize(BUFFER_SIZE);
		TCHAR szVersion[BUFFER_SIZE];
		const LONG lRet( ::RegQueryValueEx(hKey1, NULL, NULL, NULL, (LPBYTE)szVersion, &cSize) );

		// Close the registry keys
		::RegCloseKey(hKey1);
		::RegCloseKey(hKey);

		// Error while querying for value
		if(ERROR_SUCCESS != lRet){
			return eOfficeVersion_Unknown;
		}

		const CString progAndVersion(szVersion);
		// At this point szVersion contains the ProgID followed by a number. 
		// For example, Word 97 will return Word.Application.8 and Word 2000 will return Word.Application.9

		const int lastDot( progAndVersion.ReverseFind(_T('.')) );
		const int firstCharOfVersion( lastDot + 1); // + 1 to get rid of the dot at the front
		const CString versionString( progAndVersion.Right(progAndVersion.GetLength() - firstCharOfVersion) );

		return StringToVersion(versionString);
	}
	static BOOL GetPath(LPOLESTR szApp, LPTSTR szPath, ULONG cSize)
	{
		CLSID clsid;
		LPOLESTR pwszClsid;
		TCHAR  szKey[128];
		CHAR  szCLSID[60];
		HKEY hKey;

		// szPath must be at least 255 char in size
		if (cSize < 255)
			return FALSE;

		// Get the CLSID using ProgID
		HRESULT hr = CLSIDFromProgID(szApp, &clsid);
		if (FAILED(hr))
		{
			return FALSE;
		}

		// Convert CLSID to String
		hr = StringFromCLSID(clsid, &pwszClsid);
		if (FAILED(hr))
		{
			return FALSE;
		}

		// Convert result to ANSI
		//WideCharToMultiByte(CP_ACP, 0, pwszClsid, -1, szCLSID, 60, NULL, NULL);

		// Free memory used by StringFromCLSID

		// Format Registry Key string
		wsprintf(szKey, _T("CLSID\\%s\\LocalServer32"), pwszClsid);

		CoTaskMemFree(pwszClsid);


		// Open key to find path of application
		LONG lRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, szKey, 0, KEY_ALL_ACCESS, &hKey);
		if (lRet != ERROR_SUCCESS) 
		{
			// If LocalServer32 does not work, try with LocalServer
			wsprintf(szKey, _T("CLSID\\%s\\LocalServer"), szCLSID);
			lRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, szKey, 0, KEY_ALL_ACCESS, &hKey);
			if (lRet != ERROR_SUCCESS) 
			{
				return FALSE;
			}
		}

		// Query value of key to get Path and close the key
		lRet = RegQueryValueEx(hKey, NULL, NULL, NULL, (BYTE*)szPath, &cSize);
		RegCloseKey(hKey);
		if (lRet != ERROR_SUCCESS)
		{
			return FALSE;
		}

		// Strip off the '/Automation' switch from the path
		TCHAR *x = _tcsrchr(szPath, _T('/'));
		if(0!= x) // If no /Automation switch on the path
		{
			int result = x - szPath; 
			szPath[result]  = '\0';  // If switch there, strip it
		}   
		return TRUE;
	}

};
