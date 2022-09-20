#pragma once

#ifndef __UTILITY_H_
#define __UTILITY_H_
#include "Structs.h"
#include <filesystem>

void GetFoldersInDirectory(_StringVector &FolderList, const std::wstring &Root, bool bRecurseSubdirectories = false);

void CopyFilesByWildCard(LPCTSTR sSource, LPCTSTR sTarget, LPCTSTR sWildCard, bool bRecursive = false);
void GetFilesInDirectory(const std::wstring &WildCard, _StringVector &FileList, const std::wstring &Root, bool bRecurseSubdirectories = false);
void GetFilesInDirectory(LPCTSTR WildCard, _StringVector &FileList, LPCTSTR Root, bool bRecurseSubdirectories = false);
void GetFilesInDirectory(const CString &WildCard, _StringVector &FileList, const std::wstring &Root, bool bRecurseSubdirectories = false);
void GetFilesInDirectory(_StringVector &WildCardList, _StringVector &FileList, LPCTSTR Root, bool bRecurseSubdirectories = false);
void GetFilesInDirectory(_StringVector &WildCardList, _StringVector &FileList, const std::wstring &Root, bool bRecurseSubdirectories = false);

void CopyToClipboard(std::wstring &ws);
void CopyToClipboard(std::string &s);
void CopyToClipboard(const char *szText);
bool CopyFromClipboard(std::wstring &s);
bool CopyFromClipboard(std::string &s);
static int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData);
bool BrowseForFolder(HWND hwnd, const wchar_t *szStartFolder, const wchar_t *szTitle, std::wstring &wsFolder);
bool BrowseForFolder(HWND hwnd, const char *szStartFolder, const char *szTitle, std::string &sFolder);

template< typename T > void EraseDataVector(T &t)
{
	for (T::iterator it = t.begin(); it != t.end(); it++)
	{
		if ((*it) != NULL)
		{
			delete (*it);
			(*it) = NULL;
		}
	}

	t.erase(t.begin(), t.end());
}


#endif	// #ifndef __UTILITY_H_