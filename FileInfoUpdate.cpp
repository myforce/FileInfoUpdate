// FileInfoUpdate.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <wtypes.h>
#include <winver.h>
#include <stdio.h>
#include <map>
#include <list>
#include <Shlwapi.h>

#pragma comment(lib, "Shlwapi.lib")

#define STRING_MAX       256
#define STRING_MAX_PATH  STRING_MAX + 25 // \StringFileInfo\000004b0\<string>


#pragma comment(lib, "version.lib")

bool ParseVersion(LPTSTR & szVersion, WORD(&arrVersion)[4]);
BOOL CALLBACK EnumLanguages(HMODULE  hModule, LPCTSTR  lpszType, LPCTSTR  lpszName, WORD     wIDLanguage, LONG_PTR lParam);
int RunExternalProcessAndGetExitCode(LPTSTR szCommandLine);

int _tmain(int argc, _TCHAR* argv[])
{
	//input: <file> [args]
	//possible args:
	//		/fv <fileversion>
	//		/pv <productversion>
	//		/s  <name> <value>

	//parse input
	if (__argc <= 1)
	{
		//print usage info
		LPTSTR szUsage = 
			_T("usage: %s <file> [arguments]\r\n")
			_T("possible arguments:\r\n")
			_T("\t/fv <fileversion> : set file version\r\n")
			_T("\t/pv <productversion> : set product version\r\n")
			_T("\t/s  <name> <value> : set string value")
			_T("\vl : get versioninfo language code")
			_T("when calling without arguments, the full version info will be returned in RC format");
		_ftprintf_s(stdout, szUsage, __targv[0]);
		return 0;
	}

	if (__argc < 2)
	{
		//invalid command
		_ftprintf_s(stderr, _T("Invalid command"));
		return 1;
	}

	LPCTSTR szVersionFile = __targv[1];
	if (_taccess(szVersionFile, 06) != 0)
	{
		//no read/write access
		_ftprintf_s(stderr, _T("Access to '%s' is denied"), szVersionFile);
		return 2;
	}

	if (__argc == 3 && _tcscmp(__targv[2], _T("/vl")) == 0)
	{
		//special case: retrieve version-info language
		::SetErrorMode(SEM_NOOPENFILEERRORBOX | SEM_FAILCRITICALERRORS);

		HMODULE hModule = LoadLibrary(szVersionFile);
		if (!hModule)
			return -1;

		WORD wLanguage = 0;
		std::list<WORD> lstLanguages;
		if (::EnumResourceLanguages(hModule, RT_VERSION, MAKEINTRESOURCE(VS_VERSION_INFO), &EnumLanguages, (LONG_PTR)&lstLanguages))
		{
			if (lstLanguages.size() > 0)
				wLanguage = lstLanguages.front();
		}
		FreeLibrary(hModule);

		return (int)wLanguage;
	}

	bool bDumpRC = (__argc == 2);
	LPTSTR szNewFileVersion = NULL;
	LPTSTR szNewProductVersion = NULL;
	std::map<LPCTSTR, LPCTSTR> mapNewStringValues;

	int nParameter = 1;
	while (++nParameter < __argc)
	{
		if (_tcscmp(__targv[nParameter], _T("/fv")) == 0)
		{
			if (nParameter++ >= __argc)
				break;
			szNewFileVersion = __targv[nParameter];
		}
		if (_tcscmp(__targv[nParameter], _T("/pv")) == 0)
		{
			if (nParameter++ >= __argc)
				break;
			szNewProductVersion = __targv[nParameter];
		}
		if (_tcscmp(__targv[nParameter], _T("/s")) == 0)
		{
			if (nParameter++ >= __argc)
				break;
			LPCTSTR szString = __targv[nParameter];
			if (nParameter++ >= __argc)
				break;
			LPCTSTR szStringValue = __targv[nParameter];
			mapNewStringValues[szString] = szStringValue;
			if (_tcslen(szString) > STRING_MAX)
			{
				//error
				_ftprintf_s(stderr, _T("Impossible to set string '%s' - maximum length: %d, necessary length: %d."), szString, STRING_MAX, _tcslen(szString));
				return 3;
			}
		}
	}

	//look for the modification time
	FILETIME ftLastAccessTime = { 0 }, ftLastWriteTime = { 0 };
	HANDLE hFile = CreateFile(szVersionFile, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile != INVALID_HANDLE_VALUE)
	{
		::GetFileTime(hFile, NULL, &ftLastAccessTime, &ftLastWriteTime);
		::CloseHandle(hFile);
	}

	//determine the language for the version-info block
	// try 32 and 64 bit
	WORD wVersionResourceLanguage = LANG_NEUTRAL;
	
	TCHAR szProcessFileName[MAX_PATH] = { 0 };
	_tcscpy_s(szProcessFileName, MAX_PATH, __targv[0]);
	::PathRemoveExtension(szProcessFileName);
	LPCTSTR szProcessFileNameEnd = szProcessFileName + (_tcslen(szProcessFileName) - 2);
	if (_tcscmp(szProcessFileNameEnd, _T("32")) == 0 || _tcscmp(szProcessFileNameEnd, _T("64")) == 0)
	{
		//file ends on 32 or 64, try xxx32.exe and xxx64.exe
		TCHAR szCommandLine[MAX_PATH] = { 0 };
		_stprintf_s(szCommandLine, MAX_PATH, _T("\"%s32.exe\" \"%s\" /vl"), szProcessFileName, szVersionFile);
		int nLang = RunExternalProcessAndGetExitCode(szCommandLine);
		if (nLang <= 0)
		{
			_stprintf_s(szCommandLine, MAX_PATH, _T("\"%s64.exe\" \"%s\" /vl"), szProcessFileName, szVersionFile);
			nLang = RunExternalProcessAndGetExitCode(szCommandLine);
			if (nLang > 0)
				wVersionResourceLanguage = (WORD)nLang;
		}
	}
	else
	{
		//file doesn't end on 32 or 64, try xxx.exe and xxx64.exe
		TCHAR szCommandLine[MAX_PATH] = { 0 };
		_stprintf_s(szCommandLine, MAX_PATH, _T("\"%s.exe\" \"%s\" /vl"), szProcessFileName, szVersionFile);
		int nLang = RunExternalProcessAndGetExitCode(szCommandLine);
		if (nLang <= 0)
		{
			_stprintf_s(szCommandLine, MAX_PATH, _T("\"%s64.exe\" \"%s\" /vl"), szProcessFileName, szVersionFile);
			nLang = RunExternalProcessAndGetExitCode(szCommandLine);
			if (nLang > 0)
				wVersionResourceLanguage = (WORD)nLang;
		}
	}

	//look for file info block
	DWORD dwHandle = NULL;
	DWORD dwSize = GetFileVersionInfoSize(szVersionFile, &dwHandle);
	if (0 >= dwSize)
	{
		//error
		_ftprintf_s(stderr, _T("Failed to read file info for '%s': Error %d"), szVersionFile, ::GetLastError());

		//abort
		return 4;
	}
	LPBYTE lpBuffer = new BYTE[dwSize];
	ZeroMemory(lpBuffer, dwSize);
	if (!GetFileVersionInfo(szVersionFile, dwHandle, dwSize, lpBuffer))
	{
		//error
		_ftprintf_s(stderr, _T("Failed to read file info for '%s': Error %d"), szVersionFile, ::GetLastError());

		//abort
		delete[] lpBuffer;
		return 5;
	}

	//get fixed info
	UINT uSize = 0;
	VS_FIXEDFILEINFO * pFixedInfo = NULL;
	if (!VerQueryValue(lpBuffer, _T("\\"), (LPVOID*)&pFixedInfo, &uSize))
	{
		//error
		_ftprintf_s(stderr, _T("Failed to read fixed version info"));

		//abort
		delete[] lpBuffer;
		return 6;
	}

	//get translations
	DWORD * pdwLangages = NULL;

	LPCTSTR szSubBlockTranslation = _T("\\VarFileInfo\\Translation");
	if (!VerQueryValue(lpBuffer, (LPTSTR)szSubBlockTranslation, (LPVOID*)&pdwLangages, &uSize))
	{
		//error
		_ftprintf_s(stderr, _T("Failed to read translations: Error %d."), ::GetLastError());

		//abort
		delete[] lpBuffer;
		return 7;
	}

	int nLanguageCount = uSize / sizeof(DWORD);

	//dump file info in RC format, if requested
	if (bDumpRC)
	{
		//header
		_ftprintf_s(stdout, _T("VS_VERSION_INFO VERSIONINFO\n"));

		//fixed info
		_ftprintf_s(stdout, _T(" FILEVERSION\t%u,%u,%u,%u\n"), HIWORD(pFixedInfo->dwFileVersionMS), LOWORD(pFixedInfo->dwFileVersionMS), HIWORD(pFixedInfo->dwFileVersionLS), LOWORD(pFixedInfo->dwFileVersionLS));
		_ftprintf_s(stdout, _T(" PRODUCTVERSION\t%u,%u,%u,%u\n"), HIWORD(pFixedInfo->dwProductVersionMS), LOWORD(pFixedInfo->dwProductVersionMS), HIWORD(pFixedInfo->dwProductVersionLS), LOWORD(pFixedInfo->dwProductVersionLS));
		_ftprintf_s(stdout, _T(" FILEFLAGSMASK\t%#XL\n"), pFixedInfo->dwFileFlagsMask);
		_ftprintf_s(stdout, _T(" FILEFLAGS\t%#XL\n"), pFixedInfo->dwFileFlags);
		_ftprintf_s(stdout, _T(" FILEOS\t\t%#XL\n"), pFixedInfo->dwFileOS);
		_ftprintf_s(stdout, _T(" FILETYPE\t%#X\n"), pFixedInfo->dwFileType);
		_ftprintf_s(stdout, _T(" FILESUBTYPE\t%#X\n"), pFixedInfo->dwFileSubtype);

		//string file info
		_ftprintf_s(stdout, _T("BEGIN\n"));
		_ftprintf_s(stdout, _T("\tBLOCK \"StringFileInfo\"\n"));
		_ftprintf_s(stdout, _T("\tBEGIN\n"));
		TCHAR szLanguage[9] = { 0 };
		LPCTSTR pValueBuffer;
		TCHAR szSubBlock[STRING_MAX_PATH] = { 0 };
		std::list<LPCTSTR> lstStringFileInfo = { _T("Comments"), _T("CompanyName"), _T("FileDescription"), _T("FileVersion"), _T("InternalName"), _T("LegalCopyright"), _T("LegalTrademarks"), _T("OriginalFilename"), _T("ProductName"), _T("ProductVersion"), _T("PrivateBuild"), _T("SpecialBuild") };
		for (int nLanguage = 0; nLanguage < nLanguageCount; nLanguage++)
		{
			WORD * psLanguage = (WORD*)(pdwLangages + nLanguage);
			_stprintf_s(szLanguage, 9, _T("%04x%04x"), psLanguage[0], psLanguage[1]);

			_ftprintf_s(stdout, _T("\t\tBLOCK \"%s\"\n"), szLanguage);
			_ftprintf_s(stdout, _T("\t\tBEGIN\n"));

			for (LPCTSTR szString : lstStringFileInfo)
			{
				_stprintf_s(szSubBlock, STRING_MAX_PATH, _T("\\StringFileInfo\\%s\\%s"), szLanguage, szString);
				if (VerQueryValue(lpBuffer, (LPTSTR)szSubBlock, (LPVOID*)&pValueBuffer, &uSize))
					_ftprintf_s(stdout, _T("\t\t\tVALUE \"%s\", \"%s\"\n"), szString, pValueBuffer);
			}

			_ftprintf_s(stdout, _T("\t\tEND\n"));
		}
		_ftprintf_s(stdout, _T("\tEND\n"));

		//translation info
		_ftprintf_s(stdout, _T("\tBLOCK \"VarFileInfo\"\n"));
		_ftprintf_s(stdout, _T("\tBEGIN\n"));
		for (int nLanguage = 0; nLanguage < nLanguageCount; nLanguage++)
		{
			WORD * psLanguage = (WORD*)(pdwLangages + nLanguage);
			_ftprintf_s(stdout, _T("\t\tVALUE \"Translation\", 0x%X, %ld\n"), psLanguage[0], psLanguage[1]);
		}
		_ftprintf_s(stdout, _T("\tEND\n"));

		//footer
		_ftprintf_s(stdout, _T("END\n\n"));

		return 0;
	}

	//update fixed file info
	if (szNewFileVersion || szNewProductVersion)
	{
		WORD arrFileVersion[4] = { 0 };
		bool bNewFileVersion = ParseVersion(szNewFileVersion, arrFileVersion); //note: this modifies szNewFileVersion
		WORD arrProductVersion[4] = { 0 };
		bool bNewProductVersion = ParseVersion(szNewProductVersion, arrProductVersion); //note: this modifies szNewProductVersion

		UINT uSize = 0;
		VS_FIXEDFILEINFO * pFixedInfo = NULL;
		if (!VerQueryValue(lpBuffer, _T("\\"), (LPVOID*)&pFixedInfo, &uSize))
		{
			//error
			_ftprintf_s(stderr, _T("Failed to read fixed version info"));
			return 6;
		}

		if (bNewFileVersion)
		{
			pFixedInfo->dwFileVersionMS = MAKELONG(arrFileVersion[1], arrFileVersion[0]);
			pFixedInfo->dwFileVersionLS = MAKELONG(arrFileVersion[3], arrFileVersion[2]);
		}
		if (bNewProductVersion)
		{
			pFixedInfo->dwProductVersionMS = MAKELONG(arrProductVersion[1], arrProductVersion[0]);
			pFixedInfo->dwProductVersionLS = MAKELONG(arrProductVersion[3], arrProductVersion[2]);
		}
	}

	//update data in every language
	TCHAR szLanguage[9] = { 0 };
	LPTSTR pValueBuffer;
	TCHAR szSubBlock[STRING_MAX_PATH] = { 0 };
	for (int nLanguage = 0; nLanguage < nLanguageCount; nLanguage++)
	{
		//format the language specifier
		WORD * psLanguage = (WORD*)(pdwLangages + nLanguage);
		_stprintf_s(szLanguage, 9, _T("%04x%04x"), psLanguage[0], psLanguage[1]);

		//update string info
		for (auto kvpString : mapNewStringValues)
		{
			LPCTSTR szString = kvpString.first;
			LPCTSTR szStringValue = kvpString.second;

			_stprintf_s(szSubBlock, STRING_MAX_PATH, _T("\\StringFileInfo\\%s\\%s"), szLanguage, szString);
			if (!VerQueryValue(lpBuffer, (LPTSTR)szSubBlock, (LPVOID*)&pValueBuffer, &uSize))
			{
				//error
				_ftprintf_s(stderr, _T("Failed to set string '%s' to '%s' - string doesn't exist in file info."), szString, szStringValue);

				//abort
				delete[] lpBuffer;
				return 8;
			}

			ZeroMemory(pValueBuffer, uSize * sizeof(TCHAR));
			size_t uSizeRequired = _tcslen(szStringValue) + 1;
			if (uSize < uSizeRequired)
			{
				//error
				_ftprintf_s(stderr, _T("Failed to set string '%s' to '%s' - maximum length: %d, necessary length: %d."), szString, szStringValue, uSize - 1, _tcslen(szStringValue));

				//abort
				delete[] lpBuffer;
				return 9;
			}
			_tcscpy_s(pValueBuffer, uSizeRequired, szStringValue);
		}
	}

	//update file with new info
	HANDLE hUpdate = BeginUpdateResource(szVersionFile, FALSE);
	if (!hUpdate)
	{
		//error
		_ftprintf_s(stderr, _T("Failed to update resources for '%s': Error %d."), szVersionFile, ::GetLastError());

		//abort
		delete[] lpBuffer;
		return 10;
	}
	if (!UpdateResource(hUpdate, RT_VERSION, MAKEINTRESOURCE(VS_VERSION_INFO), wVersionResourceLanguage, lpBuffer, dwSize))
	{
		//error
		_ftprintf_s(stderr, _T("Failed to update resources for '%s': Error %d."), szVersionFile, ::GetLastError());

		//abort
		EndUpdateResource(hUpdate, TRUE);
		delete[] lpBuffer;
		return 11;
	}
	if (!EndUpdateResource(hUpdate, FALSE))
	{
		//error
		_ftprintf_s(stderr, _T("Failed to commit update resources for '%s': Error %d."), szVersionFile, ::GetLastError());

		//abort
		delete[] lpBuffer;
		return 12;
	}

	//restore last modification time
	hFile = CreateFile(szVersionFile, GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile != INVALID_HANDLE_VALUE)
	{
		::SetFileTime(hFile, NULL, &ftLastAccessTime, &ftLastWriteTime);
		::CloseHandle(hFile);
	}

	//done!
	_ftprintf_s(stdout, _T("Done\r\n"));

	delete[] lpBuffer;
	return 0;
}

bool ParseVersion(LPTSTR & szVersion, WORD(&arrVersion)[4])
{
	if (szVersion == NULL)
		return false;

	TCHAR szSeparators[] = _T(".");
	TCHAR * szToken = NULL;
	TCHAR * szNextToken = NULL;

	int nNumber = 0;
	szToken = _tcstok_s(szVersion, szSeparators, &szNextToken);
	while (szToken != NULL && nNumber < 4)
	{
		arrVersion[nNumber] = (WORD)_ttoi(szToken);
		szToken = _tcstok_s(NULL, szSeparators, &szNextToken);
		nNumber++;
	}
	szVersion = NULL;
	return true;
}

BOOL CALLBACK EnumLanguages(HMODULE hModule, LPCTSTR lpszType, LPCTSTR lpszName, WORD wIDLanguage, LONG_PTR lParam)
{
	std::list<WORD> * plstLanguages = reinterpret_cast<std::list<WORD>*>(lParam);
	if (plstLanguages)// && wIDLanguage > 0)
		plstLanguages->push_back(wIDLanguage);
	return TRUE;
}

int RunExternalProcessAndGetExitCode(LPTSTR szCommandLine)
{
	//start process
	STARTUPINFO si = { 0 };
	si.wShowWindow = SW_HIDE;
	PROCESS_INFORMATION pi = { 0 };
	if (!CreateProcess(NULL, szCommandLine, NULL, NULL, TRUE, CREATE_DEFAULT_ERROR_MODE | CREATE_UNICODE_ENVIRONMENT | CREATE_NO_WINDOW, NULL, NULL, &si, &pi))
		return -1;
	CloseHandle(pi.hThread);

	//wait until process closed
	if (WAIT_OBJECT_0 != WaitForSingleObject(pi.hProcess, 60000))
	{
		CloseHandle(pi.hProcess);
		return -2;
	}

	//get exit code
	DWORD dwExitCode = 0;
	if (!GetExitCodeProcess(pi.hProcess, &dwExitCode))
	{
		CloseHandle(pi.hProcess);
		return -3;
	}

	//done
	CloseHandle(pi.hProcess);
	return (int)dwExitCode;
}
