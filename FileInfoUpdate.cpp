// FileInfoUpdate.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <wtypes.h>
#include <winver.h>
#include <stdio.h>
#include <map>
#include <list>

#define STRING_MAX       256
#define STRING_MAX_PATH  STRING_MAX + 25 // \StringFileInfo\000004b0\<string>


#pragma comment(lib, "version.lib")

bool ParseVersion(LPTSTR & szVersion, WORD(&arrVersion)[4]);

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


	//look for file info block
	DWORD dwHandle = NULL;
	DWORD dwSize = GetFileVersionInfoSize(szVersionFile, &dwHandle);
	if (0 >= dwSize)
	{
		//error
		_ftprintf_s(stderr, _T("Failed to read file info for '%s': Error %d"), szVersionFile, ::GetLastError());
		return 4;
	}
	LPBYTE lpBuffer = new BYTE[dwSize];
	if (!GetFileVersionInfo(szVersionFile, 0, dwSize, lpBuffer))
	{
		delete[] lpBuffer;

		//error
		_ftprintf_s(stderr, _T("Failed to read file info for '%s': Error %d"), szVersionFile, ::GetLastError());
		return 5;
	}

	//get fixed info
	UINT uSize = 0;
	VS_FIXEDFILEINFO * pFixedInfo = NULL;
	if (!VerQueryValue(lpBuffer, _T("\\"), (LPVOID*)&pFixedInfo, &uSize))
	{
		//error
		_ftprintf_s(stderr, _T("Failed to read fixed version info"));
		return 6;
	}

	//get translations
	DWORD * pdwLangages = NULL;

	LPCTSTR szSubBlockTranslation = _T("\\VarFileInfo\\Translation");
	if (!VerQueryValue(lpBuffer, (LPTSTR)szSubBlockTranslation, (LPVOID*)&pdwLangages, &uSize))
	{
		delete[] lpBuffer;

		//error
		_ftprintf_s(stderr, _T("Failed to read translations: Error %d."), ::GetLastError());
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

	//update data in every language, or in the language specified
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
				delete[] lpBuffer;

				//error
				_ftprintf_s(stderr, _T("Failed to set string '%s' to '%s' - string doesn't exist in file info."), szString, szStringValue);
				return 8;
			}

			ZeroMemory(pValueBuffer, _tcslen(pValueBuffer) * sizeof(TCHAR));
			if (uSize < _tcslen(szStringValue) + 1)
			{
				delete[] lpBuffer;

				//error
				_ftprintf_s(stderr, _T("Failed to set string '%s' to '%s' - maximum length: %d, necessary length: %d."), szString, szStringValue, uSize - 1, _tcslen(szStringValue));
				return 9;
			}
			_tcscpy_s(pValueBuffer, uSize, szStringValue);
	}
	}

	//update file with new info
	HANDLE hUpdate = BeginUpdateResource(szVersionFile, FALSE);
	if (!hUpdate)
	{
		delete[] lpBuffer;

		//error
		_ftprintf_s(stderr, _T("Failed to update resources for '%s': Error %d."), szVersionFile, ::GetLastError());
		return 10;
	}
	if (!UpdateResource(hUpdate, RT_VERSION, MAKEINTRESOURCE(VS_VERSION_INFO), LANG_NEUTRAL, lpBuffer, dwSize))
	{
		delete[] lpBuffer;

		//error
		_ftprintf_s(stderr, _T("Failed to update resources for '%s': Error %d."), szVersionFile, ::GetLastError());
		return 11;
	}
	EndUpdateResource(hUpdate, FALSE);

	//done!
	delete[] lpBuffer;
	_ftprintf_s(stdout, _T("Done"));
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