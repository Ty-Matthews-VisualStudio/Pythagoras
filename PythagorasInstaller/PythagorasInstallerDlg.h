
// PythagorasInstallerDlg.h : header file
//

#pragma once
#include "afxwin.h"

class PythonModule
{
public:
	std::wstring m_sName;
	std::wstring m_sX64File;
	std::wstring m_sX86File;
	bool m_bFTPFile;

	PythonModule( LPCTSTR szName, LPCTSTR szX64File, LPCTSTR szX86File)
	{
		m_sName = szName;
		if (!_tcsicmp(szX64File, _T("0")))
		{
			m_bFTPFile = false;			
		}
		else
		{			
			m_bFTPFile = true;
			m_sX64File = szX64File;
			if (_tcsicmp(szX86File, _T("0")))
			{
				m_sX86File = szX86File;
			}
			else
			{
				m_bFTPFile = false;
			}
		}
	};
	virtual ~PythonModule() {};
};

// CPythagorasInstallerDlg dialog
class CPythagorasInstallerDlg : public CDialogEx
{
// Construction
public:
	CPythagorasInstallerDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PYTHAGORASINSTALLER_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;
	std::wstringstream m_ssTempPath;
	std::wstringstream m_ssLogFile;
	std::wstring m_sModuleFile;
	std::wstring m_sPythonInstallFilex86;
	std::wstring m_sPythonInstallFilex64;
	std::vector<PythonModule> m_vPythonModules;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CRegistrySettings m_RegistrySettings;
	BOOL m_bPython;
	BOOL m_bVisualC;
	BOOL m_bPythagoras;
	_SYSTEM_INFO m_SystemInfo;
	afx_msg void OnBnClickedButtonInstall();
	bool DownloadFile(LPCTSTR szFileName, std::wstring &sLocalFile);
	bool DownloadFile(std::wstring &sFileName, std::wstring &sLocalFile);
	bool ExecuteFile(LPCTSTR szFileName, DWORD dwCreationFlags = NORMAL_PRIORITY_CLASS | CREATE_NO_WINDOW);
	bool CreateProgramFiles(LPCTSTR szLocalFile, std::wstring &sOutputPath);
	void AddToLog(LPCTSTR szText);
	void AddToLog(std::wstringstream &ssText);
	void OpenFile(LPCTSTR szFileName);
	BOOL m_bDesktopShortcut;
	BOOL m_bStartMenuShortcut;
	CButton m_btnDesktopShortcut;
	CButton m_btnStartMenuShortcut;
	afx_msg void OnBnClickedCheckPythagoras();
	BOOL m_bPythonModules;
	CListBox m_lstExistingComponents;
	void CheckForExistingComponents();
	void DownloadModuleFile();
	void DownloadPythonInstallFile();
};

/*
-------------------------------------------------------------------
Description:
Creates the actual 'lnk' file (assumes COM has been initialized).

Parameters:
pszTargetfile    - File name of the link's target.
pszTargetargs    - Command line arguments passed to link's target.
pszLinkfile      - File name of the actual link file being created.
pszDescription   - Description of the linked item.
iShowmode        - ShowWindow() constant for the link's target.
pszCurdir        - Working directory of the active link.
pszIconfile      - File name of the icon file used for the link.
iIconindex       - Index of the icon in the icon file.

Returns:
HRESULT value >= 0 for success, < 0 for failure.
--------------------------------------------------------------------
*/

HRESULT CreateShortcut(LPCWSTR pszTargetfile, LPCWSTR pszTargetargs,
	LPCWSTR pszLinkfile, LPCWSTR pszDescription,
	int iShowmode, LPCWSTR pszCurdir,
	LPCWSTR pszIconfile, int iIconindex);


class CleanUp
{
public:
	wchar_t* m_Folder;
	CleanUp() : m_Folder(NULL) {};
	virtual ~CleanUp() { if (m_Folder) CoTaskMemFree(static_cast<void*>(m_Folder)); };
};