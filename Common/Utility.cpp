#include "stdafx.h"
#include "Utility.h"

void GetFoldersInDirectory(_StringVector &FolderList, const std::wstring &Root, bool bRecurseSubdirectories /* = false */)
{
	HANDLE dir;
	WIN32_FIND_DATA file_data;

	if (bRecurseSubdirectories)
	{
		// First look for any folders that might exist here
		std::wstring sRootFolders(Root);
		sRootFolders += _T("\\*");
		if ((dir = FindFirstFile(sRootFolders.c_str(), &file_data)) == INVALID_HANDLE_VALUE)
			return; /* No files found */

		do {
			std::wstring file_name = file_data.cFileName;
			std::wstring full_file_name(Root);
			full_file_name += _T("\\");
			full_file_name += file_name;
			const bool is_directory = (file_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

			if (file_name.at(0) == '.')
				continue;

			if (is_directory)
			{
				GetFoldersInDirectory(FolderList, full_file_name, bRecurseSubdirectories);
			}
		} while (FindNextFile(dir, &file_data));
	}

	// Look for any matches to our wildcard
	std::wstring sFullWildCard(Root);
	sFullWildCard += _T("\\*.*");	
	if ((dir = FindFirstFile(sFullWildCard.c_str(), &file_data)) == INVALID_HANDLE_VALUE)
		return; /* No files found */

	do {
		const std::wstring file_name = file_data.cFileName;
		std::wstring full_file_name(Root);
		full_file_name += _T("\\");
		full_file_name += file_name;
		const bool is_directory = (file_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

		if (file_name.at(0) == '.')
			continue;

		if (is_directory)
		{
			FolderList.push_back(full_file_name);
		}
	} while (FindNextFile(dir, &file_data));

	FindClose(dir);
}

void CopyFilesByWildCard(LPCTSTR sSource, LPCTSTR sTarget, LPCTSTR sWildCard, bool bRecursive /* = false */)
{
	_StringVector FileList;
	_itString itFile;
	std::wstring sRelative;
	std::wstring sNewTarget;
	std::wstring sPathToTarget;
	TCHAR drive[_MAX_DRIVE];
	TCHAR dir[_MAX_DIR];
		
	GetFilesInDirectory( sWildCard, FileList, sSource, bRecursive);
	for (itFile = FileList.begin(); itFile != FileList.end(); itFile++)
	{
		// Extract the relative path from the filename
		sRelative = (*itFile).substr(_tcslen(sSource), std::string::npos);
		sNewTarget = sTarget;
		sNewTarget += sRelative;

		/*
		// Make the directory structure first, otherwise the copy command will fail
		_wsplitpath_s(sNewTarget.c_str(), drive, _MAX_DRIVE, dir, _MAX_DIR, NULL, 0, NULL, 0);
		sPathToTarget = drive;
		sPathToTarget += dir;
		*/

		_StringVector sv;
		boost::split(sv, sRelative, boost::is_any_of("\\"));
		sPathToTarget = sTarget;
		_itString it = sv.begin();
		for (int i = 0; i < sv.size() - 1 ; i++, it++)
		{
			sPathToTarget += (*it);
			sPathToTarget += _T("\\");
			_wmkdir(sPathToTarget.c_str());
		}

		
		//std::filesystem::copy((*itFile).c_str(), sNewTarget.c_str(), std::filesystem::copy_options::update_existing);
		std::filesystem::copy((*itFile).c_str(), sNewTarget.c_str(), std::filesystem::copy_options::overwrite_existing);
	}

}

void GetFilesInDirectory(_StringVector &WildCardList, _StringVector &FileList, LPCTSTR Root, bool bRecurseSubdirectories /* = false */)
{
	for (std::wstring sWildCard : WildCardList)
	{
		GetFilesInDirectory(sWildCard.c_str(), FileList, Root, bRecurseSubdirectories);
	}
}

void GetFilesInDirectory(_StringVector &WildCardList, _StringVector &FileList, const std::wstring &Root, bool bRecurseSubdirectories /* = false */)
{
	for (std::wstring sWildCard : WildCardList)
	{
		GetFilesInDirectory(sWildCard.c_str(), FileList, Root.c_str(), bRecurseSubdirectories);
	}
}

void GetFilesInDirectory(LPCTSTR WildCard, _StringVector &FileList, LPCTSTR Root, bool bRecurseSubdirectories /* = false */)
{
	std::wstring sWildCard;
	std::wstring sRoot;
	sWildCard = WildCard;
	sRoot = Root;
	GetFilesInDirectory(sWildCard, FileList, sRoot, bRecurseSubdirectories);
}

void GetFilesInDirectory(const CString &WildCard, _StringVector &FileList, const std::wstring &Root, bool bRecurseSubdirectories /* = false */)
{
	std::wstring sWildCard;
	sWildCard = (LPCTSTR)WildCard;
	GetFilesInDirectory(sWildCard, FileList, Root, bRecurseSubdirectories);
}
void GetFilesInDirectory(const std::wstring &WildCard, _StringVector &FileList, const std::wstring &Root, bool bRecurseSubdirectories /* = false */)
{
	HANDLE dir;
	WIN32_FIND_DATA file_data;

	if (bRecurseSubdirectories)
	{
		// First look for any folders that might exist here
		std::wstring sRootFolders(Root);
		sRootFolders += _T("\\*");
		if ((dir = FindFirstFile(sRootFolders.c_str(), &file_data)) == INVALID_HANDLE_VALUE)
			return; /* No files found */

		do {
			std::wstring file_name = file_data.cFileName;
			std::wstring full_file_name(Root);
			full_file_name += _T("\\");
			full_file_name += file_name;
			const bool is_directory = (file_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

			if (file_name.at(0) == '.')
				continue;

			if (is_directory)
			{
				GetFilesInDirectory(WildCard, FileList, full_file_name, bRecurseSubdirectories);
			}
		} while (FindNextFile(dir, &file_data));
	}

	// Look for any matches to our wildcard
	std::wstring sFullWildCard(Root);
	sFullWildCard += _T("\\");
	sFullWildCard += WildCard;
	if ((dir = FindFirstFile(sFullWildCard.c_str(), &file_data)) == INVALID_HANDLE_VALUE)
		return; /* No files found */

	do {
		const std::wstring file_name = file_data.cFileName;
		std::wstring full_file_name(Root);
		full_file_name += _T("\\");
		full_file_name += file_name;
		const bool is_directory = (file_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

		if (file_name.at(0) == '.')
			continue;

		if (!is_directory)
		{
			FileList.push_back(full_file_name);
		}
	} while (FindNextFile(dir, &file_data));

	FindClose(dir);
} // GetFilesInDirectory

void CopyToClipboard(std::wstring &ws)
{
	std::string s;
	s = CW2A(ws.c_str());
	CopyToClipboard(s.c_str());
}
void CopyToClipboard(std::string &s)
{
	CopyToClipboard(s.c_str());
}

void CopyToClipboard(const char *szText)
{
	HGLOBAL h;
	LPTSTR arr;

	h = GlobalAlloc(GMEM_MOVEABLE, strlen(szText) + 1);
	arr = (LPTSTR)GlobalLock(h);
	strcpy_s((char*)arr, strlen(szText) + 1, szText);	
	GlobalUnlock(h);

	::OpenClipboard(NULL);
	EmptyClipboard();
	SetClipboardData(CF_TEXT, h);
	CloseClipboard();
}

bool CopyFromClipboard(std::wstring &s)
{
	// Try opening the clipboard
	if (!::OpenClipboard(NULL))
	{
		return false;
	}

	// Get handle of clipboard object for text
	HANDLE hData = ::GetClipboardData(CF_UNICODETEXT);
	if (hData == NULL)
	{
		return false;
	}

	// Lock the handle to get the actual text pointer
	wchar_t *pszText = static_cast<wchar_t*>(::GlobalLock(hData));
	if (pszText == NULL)
	{
		return false;
	}

	// Save text in a string class instance
	s = pszText;

	// Release the lock
	::GlobalUnlock(hData);

	// Release the clipboard
	::CloseClipboard();

	return true;
}

bool CopyFromClipboard(std::string &s)
{
	// Try opening the clipboard
	if (!::OpenClipboard(NULL))
	{
		return false;
	}
	
	// Get handle of clipboard object for ANSI text
	HANDLE hData = ::GetClipboardData(CF_TEXT);
	if (hData == NULL)
	{
		return false;
	}
	
	// Lock the handle to get the actual text pointer
	char *pszText = static_cast<char*>(::GlobalLock(hData));
	if (pszText == NULL)
	{
		return false;
	}
		
	// Save text in a string class instance
	s = pszText;

	// Release the lock
	::GlobalUnlock(hData);

	// Release the clipboard
	::CloseClipboard();

	return true;
}

bool BrowseForFolder(HWND hwnd, const char *szStartFolder, const char *szTitle, std::string &sFolder)
{
	std::wstring sStart;
	std::wstring wFolder;
	std::wstring sTitle;
	sStart = CA2W(szStartFolder);
	sTitle = CA2W(szTitle);
	return BrowseForFolder(hwnd, sStart.c_str(), sTitle.c_str(), wFolder);
	sFolder = CW2A(wFolder.c_str());
}

static int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData)
{
	// If the BFFM_INITIALIZED message is received
	// set the path to the start path.
	switch (uMsg)
	{
	case BFFM_INITIALIZED:
	{
		if (NULL != lpData)
		{
			::SendMessage(hwnd, BFFM_SETSELECTION, TRUE, lpData);
		}
	}
	}

	return 0; // The function should always return 0.
}


bool BrowseForFolder(HWND hwnd, const wchar_t *szStartFolder, const wchar_t *szTitle, std::wstring &wsFolder)
{
	BROWSEINFO bi = { 0 };
	LPITEMIDLIST pidl = NULL;
	TCHAR szDisplayName[MAX_PATH];
	bool bRet = false;
	CoInitialize(NULL);
	
	szDisplayName[0] = '\0';

	bi.hwndOwner = hwnd;
	bi.pidlRoot = NULL;
	bi.pszDisplayName = szDisplayName;
	bi.lpszTitle = szTitle;
	bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
	bi.lpfn = BrowseCallbackProc;
	bi.lParam = (LPARAM)szStartFolder;
	bi.iImage = 0;

	pidl = SHBrowseForFolder(&bi);
	TCHAR  szPathName[MAX_PATH];
	if (pidl)
	{
		bRet = SHGetPathFromIDList(pidl, szPathName);
		CoTaskMemFree(pidl);
		if (bRet)
		{
			wsFolder = szPathName;			
		}
	}
	CoUninitialize();
	return bRet;
}