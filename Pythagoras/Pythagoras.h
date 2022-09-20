
// Pythagoras.h : main header file for the Pythagoras application
//
#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"       // main symbols

class CPythagorasView;
class CMainFrame;

// CPythagorasApp:
// See Pythagoras.cpp for the implementation of this class
//

extern int CM_EXECUTETHREAD;
extern int CM_ENGINE_CALLBACK;

typedef enum
{
	eETWClearOutput = 0,
	eETWSelectPane,
	eETWThreadFinished
} ExecuteThreadMessageWPARAMEnum;

class CPythagorasApp : public CWinAppEx
{
public:
	HANDLE m_hMutex;
	UINT m_uiWinMsg;
	_vLPFileFolderData m_vFileFolderData;
	_vScriptFunction m_vScriptFunction;
	LPScriptFunctionStruct m_lpSelectedScript;
	DWORD_PTR m_dwSelectedFunction;
	CRegistrySettings m_RegistrySettings;
	HANDLE m_hThreadMutex;
	std::vector<HANDLE> m_ThreadHandles;
	std::wstringstream m_ssPythonEngPath;

	std::map<std::wstring, bool> m_DataMap;
	typedef std::map<std::wstring, bool>::iterator _itDataMap;
	
	std::wstring m_sPythagorasFolder;
	std::wstring m_sTempFolder;
	std::wstring m_sInspectScript;
	std::wstring m_sInjectScript;
	std::wstring m_sTemplateScript;
	std::wstring m_sTrueLocalScriptFolder;
	std::wstring m_sTrueCommonScriptFolder;
	std::wstring m_sTempLocalScriptFolder;
	std::wstring m_sTempCommonScriptFolder;
	std::wstring m_sDefaultHTMLTemplate;
	bool m_bAdminAccount;

public:
	CPythagorasApp();

	bool DownloadCommonFiles(std::wstringstream &ssPythagorasFolder);
	void UploadCommonFiles();
	void ZipCommonFiles(std::wstring &sZipFile);
	void GetScriptsFolders(std::wstring &sCommonFolder, std::wstring &sLocalFolder);
	void ClearScriptFunctions();
	void ParsePythonScripts(std::wstring& sScriptFolder, std::wstring& sTrueScriptFolder, _vScriptFunction& vScriptFunction, eScriptType ScriptType);
	void CopyLocalAndCommonScripts(std::wstring& sCommonFolder, std::wstring& sLocalFolder);
	void RefreshScripts(bool bRefreshMainFrame = true);
	void CompileScripts(bool bDeleteMemory = true);
	void ExecuteScript(bool bDeleteMemory = true);
	void ExecuteScript(_vLPFileFolderData &vFileFolderData, bool bDeleteMemory = true);
	static unsigned __stdcall ExecuteScriptThread(void *lpParams);
	void CheckScriptThreads(DWORD dwTime = 500, bool bUpdateProgress = false );
	// This function is used for debugging purposes
	void DeleteSharedMemory(bool bDisplayError = true);
	static bool StartPythonEngine(LPCTSTR szPath);
	static void TroubleshootIssues();
	static void GetPathTSLog(std::wstring &wsTSLog);
	static void OpenTSLog();
	static void AddToTSLog(LPCTSTR szText, bool bNewLine = false, bool bReset = false);
	static void AddToTSLog(std::wstringstream &ssText, bool bNewLine = false, bool bReset = false);
	void AddFileFolder(LPCTSTR szPath);
	void AddFile(LPCTSTR szFileName);
	void AddFolder(LPCTSTR szFolderName);	
	void CheckAndClearOutput();
	void AddToOutput(OutputTypeEnum eOutputType, LPCTSTR szText, bool bSelect = false);
	void SetOutput(OutputTypeEnum eOutputType, LPCTSTR szText, bool bSelect = false);
	CPythagorasView *GetView();
	CMainFrame *GetMainFrame() { return (CMainFrame *)m_pMainWnd; };
	bool AllowAdmin() { return m_bAdminAccount; };
	void CheckForAdmin();
	static void DebugOut(const wchar_t *szText, bool bNewLine = false);
	static void GetAppDirectory(std::string &sDirectory);
	static void GetAppDirectory(std::wstring &wDirectory);
	void CheckPythonVersion();
	
// Overrides
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

// Implementation
	UINT  m_nAppLook;
	BOOL  m_bHiColorIcons;

	virtual void PreLoadState();
	virtual void LoadCustomState();
	virtual void SaveCustomState();

	afx_msg void OnAppAbout();
	afx_msg void OnAppRefreshScripts();
	DECLARE_MESSAGE_MAP()
	afx_msg void OnFileSettings();
	afx_msg void OnExplorerEdit();
	afx_msg void OnOpenFolder();
	afx_msg void OnPopupExplorerNewScript();
	afx_msg void OnUpdatePopupExplorerNewScript(CCmdUI *pCmdUI);
	virtual int Run();
	afx_msg void OnHelpTroubleshootissues();
};

extern CPythagorasApp theApp;