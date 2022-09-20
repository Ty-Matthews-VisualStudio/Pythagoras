
// Pythagoras.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "afxwinappex.h"
#include "afxdialogex.h"
#include "Pythagoras.h"
#include "MainFrm.h"

#include "PythagorasDoc.h"
#include "PythagorasView.h"
#include "AppSettingsDlg.h"
#include "NewScriptDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

int CM_EXECUTETHREAD = RegisterWindowMessage(_T("CM_EXECUTETHREAD"));
int CM_ENGINE_CALLBACK = RegisterWindowMessage(_T("CM_ENGINE_CALLBACK"));

// CPythagorasApp

BEGIN_MESSAGE_MAP(CPythagorasApp, CWinAppEx)
	ON_COMMAND(ID_APP_ABOUT, &CPythagorasApp::OnAppAbout)
	// Standard file based document commands
	ON_COMMAND(ID_FILE_NEW, &CWinAppEx::OnFileNew)
	ON_COMMAND(ID_FILE_OPEN, &CWinAppEx::OnFileOpen)
	ON_COMMAND(ID_APP_REFRESH_SCRIPTS, &CPythagorasApp::OnAppRefreshScripts)
	ON_COMMAND(ID_FILE_SETTINGS, &CPythagorasApp::OnFileSettings)
	ON_COMMAND(ID_EXPLORER_EDIT, &CPythagorasApp::OnExplorerEdit)
	ON_COMMAND(ID_OPEN_FOLDER, &CPythagorasApp::OnOpenFolder)
	ON_COMMAND(ID_POPUP_EXPLORER_NEW_SCRIPT, &CPythagorasApp::OnPopupExplorerNewScript)
	ON_UPDATE_COMMAND_UI(ID_POPUP_EXPLORER_NEW_SCRIPT, &CPythagorasApp::OnUpdatePopupExplorerNewScript)
	ON_COMMAND(ID_HELP_TROUBLESHOOTISSUES, &CPythagorasApp::OnHelpTroubleshootissues)
END_MESSAGE_MAP()


// CPythagorasApp construction

CPythagorasApp::CPythagorasApp() : m_dwSelectedFunction(-1), m_bAdminAccount(false)
{
	m_bHiColorIcons = TRUE;

	// TODO: replace application ID string below with unique ID string; recommended
	// format for string is CompanyName.ProductName.SubProduct.VersionInformation
	SetAppID(_T("Pythagoras.AppID.NoVersion"));

	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

// The one and only CPythagorasApp object

CPythagorasApp theApp;


// CPythagorasApp initialization

BOOL CPythagorasApp::InitInstance()
{
	// InitCommonControlsEx() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// Set this to include all the common control classes you want to use
	// in your application.
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinAppEx::InitInstance();
		
	m_hMutex = ::CreateMutex(NULL, FALSE, _T("PythagorasMutex"));
	if (m_hMutex == NULL || ::GetLastError() == ERROR_ALREADY_EXISTS)
	{
		::MessageBox(NULL, _T("Pythagoras is already running.  Only one instance can be open at a time."), NULL, MB_OK);
		return FALSE;
	}

	// Initialize OLE libraries
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}

	AfxEnableControlContainer();
	EnableTaskbarInteraction(FALSE);

	// AfxInitRichEdit2() is required to use RichEdit control	
	// AfxInitRichEdit2();

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	// of your final executable, you should remove from the following
	// the specific initialization routines you do not need
	// Change the registry key under which our settings are stored
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization
	SetRegistryKey(_T("Local AppWizard-Generated Applications"));
	LoadStdProfileSettings(0);  // Load standard INI file options (including MRU)

	InitContextMenuManager();
	InitKeyboardManager();
	InitTooltipManager();
	CMFCToolTipInfo ttParams;
	ttParams.m_bVislManagerTheme = TRUE;
	theApp.GetTooltipManager()->SetTooltipParams(AFX_TOOLTIP_TYPE_ALL,
		RUNTIME_CLASS(CMFCToolTipCtrl), &ttParams);

	CheckForAdmin();

	// Register the application's document templates.  Document templates
	//  serve as the connection between documents, frame windows and views
	CSingleDocTemplate* pDocTemplate;
	pDocTemplate = new CSingleDocTemplate(
		IDR_MAINFRAME,
		RUNTIME_CLASS(CPythagorasDoc),
		RUNTIME_CLASS(CMainFrame),       // main SDI frame window
		RUNTIME_CLASS(CPythagorasView));
	if (!pDocTemplate)
		return FALSE;
	AddDocTemplate(pDocTemplate);
	
	// Parse command line for standard shell commands, DDE, file open
	CCommandLineInfo cmdInfo;
	ParseCommandLine(cmdInfo);

	// Enable DDE Execute open
	//EnableShellOpen();
	//RegisterShellFileTypes(TRUE);

	// Dispatch commands specified on the command line.  Will return FALSE if
	// app was launched with /RegServer, /Register, /Unregserver or /Unregister.
	if (!ProcessShellCommand(cmdInfo))
		return FALSE;

	// The one and only window has been initialized, so show and update it
	m_pMainWnd->ShowWindow(SW_SHOW);
	m_pMainWnd->UpdateWindow();
	// call DragAcceptFiles only if there's a suffix
	//  In an SDI app, this should occur after ProcessShellCommand
	// Enable drag/drop open
	m_pMainWnd->DragAcceptFiles();

	TCHAR szPath[MAX_PATH];	
	::GetModuleFileName(AfxGetApp()->m_hInstance, szPath, MAX_PATH);
	TCHAR drive[_MAX_DRIVE];
	TCHAR dir[_MAX_DIR];
	_wsplitpath_s(szPath, drive, _MAX_DRIVE, dir, _MAX_DIR, NULL, 0, NULL, 0);
	m_ssPythonEngPath.str(_T(""));
#ifdef _DEBUG
	m_ssPythonEngPath << drive << dir << _T("PythagorasEngineDEBUG.exe");
#else
	m_ssPythonEngPath << drive << dir << _T("PythagorasEngine.exe");
#endif

	m_hThreadMutex = CreateMutex(NULL, FALSE, NULL);

	CheckPythonVersion();

	// RefreshScripts() MUST come after m_ssPythonEngPath is set above, since it creates that process during execution
	RefreshScripts(true);	
			
	return TRUE;
}

int CPythagorasApp::ExitInstance()
{
	//TODO: handle additional resources you may have added
	AfxOleTerm(FALSE);

	EraseDataVector< _vLPFileFolderData >(m_vFileFolderData);
	ClearScriptFunctions();
	CheckScriptThreads(INFINITE,false);
	CloseHandle(m_hThreadMutex);

	return CWinAppEx::ExitInstance();
}

// CPythagorasApp message handlers


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonSourceCode();
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON_SOURCE_CODE, &CAboutDlg::OnBnClickedButtonSourceCode)
END_MESSAGE_MAP()

// App command to run the dialog
void CPythagorasApp::OnAppAbout()
{
	CAboutDlg aboutDlg;
	aboutDlg.DoModal();
}


// CPythagorasApp customization load/save methods

void CPythagorasApp::PreLoadState()
{
	BOOL bNameValid;
	CString strName;
	bNameValid = strName.LoadString(IDS_EDIT_MENU);
	ASSERT(bNameValid);
	GetContextMenuManager()->AddMenu(strName, IDR_POPUP_EDIT);
	bNameValid = strName.LoadString(IDS_EXPLORER);
	ASSERT(bNameValid);
	GetContextMenuManager()->AddMenu(strName, IDR_POPUP_EXPLORER);
	bNameValid = strName.LoadString(IDS_STDOUT_TAB);
	ASSERT(bNameValid);
	GetContextMenuManager()->AddMenu(strName, IDR_OUTPUT_POPUP);
}

void CPythagorasApp::LoadCustomState()
{
}

void CPythagorasApp::SaveCustomState()
{
}

afx_msg void CPythagorasApp::OnAppRefreshScripts()
{	
	RefreshScripts(true);
}

void CPythagorasApp::RefreshScripts( bool bRefreshMainFrame /* = true */)
{
	std::wstring sCommonFolder;
	std::wstring sLocalFolder;

	ClearScriptFunctions();
	GetScriptsFolders(sCommonFolder, sLocalFolder);

	if (sCommonFolder.size() > 0 && sLocalFolder.size() > 0)
	{
		ParsePythonScripts(sCommonFolder, m_sTrueCommonScriptFolder, m_vScriptFunction, ScriptTypeCommon);
		ParsePythonScripts(sLocalFolder, m_sTrueLocalScriptFolder, m_vScriptFunction, ScriptTypeLocal);

		CompileScripts(true);
	}

	if(bRefreshMainFrame)
	{
		CMainFrame *pMainFrame = (CMainFrame *)m_pMainWnd;
		pMainFrame->RefreshScripts();
	}	
}

void CPythagorasApp::ClearScriptFunctions()
{
	for (_itScriptFunction itScript = m_vScriptFunction.begin(); itScript != m_vScriptFunction.end(); itScript++)
	{
		EraseDataVector< _vFunction >((*itScript)->vFunctions);
	}
	EraseDataVector< _vScriptFunction >(m_vScriptFunction);
}

void CPythagorasApp::UploadCommonFiles()
{
	std::wstring sZipFile;
	CInternetSession InetSession(_T("Pythagoras/1.0"));	
	
	class FTPWrapper
	{
	public:
		CFtpConnection* m_pConnect = NULL;
		FTPWrapper(const FTPWrapper &pCopy) : m_pConnect(pCopy.m_pConnect) {};
		FTPWrapper(CFtpConnection* pConnect) : m_pConnect(pConnect) {};
		~FTPWrapper()
		{
			if (m_pConnect)
			{
				m_pConnect->Close();
				delete m_pConnect;
				m_pConnect = NULL;
			}
		}
	};
	
	try
	{
		// Open a connection
		FTPWrapper ftp(InetSession.GetFtpConnection(m_RegistrySettings.m_sFTPAddress.c_str(), m_RegistrySettings.m_sFTPUser.c_str(), m_RegistrySettings.m_sFTPPassword.c_str()));

		// Zip all the common files
		ZipCommonFiles(sZipFile);

		if (!ftp.m_pConnect->PutFile(sZipFile.c_str(), _T("Pythagoras.zip")))
		{
			LPVOID lpMessageBuffer;
			DWORD dwErr = ::GetLastError();
			::FormatMessage
			(
				FORMAT_MESSAGE_ALLOCATE_BUFFER |
				FORMAT_MESSAGE_FROM_SYSTEM,		// source and processing options 
				NULL,							// address of  message source 
				dwErr,						// requested message identifier 
				MAKELANGID
				(
					LANG_NEUTRAL,
					SUBLANG_DEFAULT
				),								// language identifier for requested message 
				(LPTSTR)&lpMessageBuffer,		// address of message buffer 
				0,								// maximum size of message buffer 
				NULL							// address of array of message inserts 
			);
			std::wstringstream ssOutput;
			ssOutput << _T("Error uploading ZIP file to FTP server.\nCode:") << dwErr << _T("\nMessage:") << (LPTSTR)&lpMessageBuffer;
			::LocalFree(lpMessageBuffer);			
			::MessageBox(NULL, ssOutput.str().c_str(), NULL, MB_ICONSTOP);
		}
		else
		{
			theApp.AddToOutput(eStandardOutput, _T("Common file upload complete"), true);
		}
	}
	catch (std::exception &ex)
	{
		std::wstringstream ssMessage;
		ssMessage << _T("Caught exception in UploadCommonFiles(): [Message] ") << ex.what();
		::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);
	}
	catch (CInternetException* pEx)
	{
		std::wstringstream ssMessage;
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		ssMessage << _T("Caught InternetException in UploadCommonFiles(): [Message] ") << sz << _T(" [Code] ") << pEx->m_dwError;
		::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);
		pEx->Delete();
	}
}

void CPythagorasApp::ZipCommonFiles(std::wstring &sZipFile)
{
	std::wstringstream ssZipFile;
	TCHAR lpTempPathBuffer[MAX_PATH];
	DWORD dwRetVal = GetTempPath(MAX_PATH, lpTempPathBuffer);
	if (dwRetVal > MAX_PATH || (dwRetVal == 0))
	{
		throw std::exception("GetTempPath() failed");
	}

	_StringVector vFiles;
	_StringVector vWildCards;
	vWildCards.push_back(_T("*.py"));
	vWildCards.push_back(_T("*.html"));
	vWildCards.push_back(_T("*.txt"));
	GetFilesInDirectory(vWildCards, vFiles, m_sTempCommonScriptFolder.c_str(), true);

	for (_itString it = vFiles.begin(); it != vFiles.end();)
	{
		TCHAR fname[_MAX_FNAME];
		TCHAR fext[_MAX_EXT];
		_wsplitpath_s((*it).c_str(), NULL, 0, NULL, 0, fname, _MAX_FNAME, fext, _MAX_EXT);

		std::wstring sRelativePath = (*it).substr(wcslen(m_sTempCommonScriptFolder.c_str()));
		std::wstring sFName = fname;
		sFName += fext;
		bool bSkip = false;
#if 0
		// Only copy scripts that begin with "Pythagoras"... all else should be removed
		if (!_tcsnicmp(sRelativePath.c_str(), _T("\\Scripts\\Local"), wcslen(_T("\\Scripts\\Local"))))
		{
			if (_tcsnicmp(fname, _T("Pythagoras"), wcslen(_T("Pythagoras"))))
			{
				bSkip = true;
			}
		}
		

		// Only copy Default.html and Tutorial.html from the Templates path
		if (!_tcsnicmp(sRelativePath.c_str(), _T("\\Templates"), wcslen(_T("\\Templates"))))
		{
			if (_tcsnicmp(fname, _T("Default"), wcslen(_T("Default"))) && _tcsnicmp(fname, _T("Tutorial"), wcslen(_T("Tutorial"))))
			{
				bSkip = true;
			}
		}

		// Don't copy anything from the Temp folder
		if (!_tcsnicmp(sRelativePath.c_str(), _T("\\Scripts\\Temp"), wcslen(_T("\\Scripts\\Temp"))))
		{
			bSkip = true;
		}
#else
#define SKIP_FILE(s) { if (!_tcsnicmp(sFName.c_str(), _T(s), wcslen(_T(s)))) { bSkip = true; } }
		SKIP_FILE("PythagorasInject.py");
		SKIP_FILE("PythagorasInspect.py");
		SKIP_FILE("DefaultTemplate.html");
		
#endif

		if (bSkip)
		{
			it = vFiles.erase(it);
		}
		else
		{
			it++;
		}
	}

	ssZipFile << lpTempPathBuffer << _T("PythagorasCommonFiles.zip");
	ZipFile(ssZipFile.str().c_str(), m_sTempCommonScriptFolder.c_str(), vFiles);

	sZipFile = ssZipFile.str().c_str();
}

bool CPythagorasApp::DownloadCommonFiles(std::wstringstream &ssPythagorasFolder)
{
	static unsigned int iCounter = 0;
	bool bReturn = true;
	CInternetSession InetSession(_T("Pythagoras/1.0"));
	TCHAR lpTempPathBuffer[MAX_PATH];
	DWORD dwRetVal = 0;

	class FTPWrapper
	{
	public:
		CFtpConnection* m_pConnect;
		FTPWrapper() : m_pConnect(NULL) {};
		FTPWrapper(const FTPWrapper &pCopy) : m_pConnect(pCopy.m_pConnect) {};
		FTPWrapper(CFtpConnection* pConnect) : m_pConnect(pConnect) {};
		~FTPWrapper()
		{
			if (m_pConnect)
			{
				m_pConnect->Close();
				delete m_pConnect;
				m_pConnect = NULL;
			}
		}
	};

	if (AllowAdmin())
	{
		// Only ask if it's the second (or more) time we've refreshed.  First happens automatically when the application first starts.
		if (iCounter++ > 0)
		{
			int nRet = AfxMessageBox(_T("Do you want to download and unzip common files?"), MB_YESNO);
			if (nRet == IDNO)
			{
				return true;
			}
		}		
	}
	
	
	
	try
	{
		InetSession.SetOption(INTERNET_OPTION_CONNECT_TIMEOUT, 20 * 1000); // Timeout after 10 seconds
		// Open a connection
		FTPWrapper ftp(InetSession.GetFtpConnection(m_RegistrySettings.m_sFTPAddress.c_str(), m_RegistrySettings.m_sFTPUser.c_str(), m_RegistrySettings.m_sFTPPassword.c_str()));

		// use a file find object to enumerate files
		CFtpFileFind finder(ftp.m_pConnect);

		//  Get the temp path env string (no guarantee it's a valid path).
		dwRetVal = GetTempPath(MAX_PATH, lpTempPathBuffer);
		if (dwRetVal > MAX_PATH || (dwRetVal == 0))
		{
			throw std::exception("GetTempPath() failed");
		}		
		std::wstringstream ssZipFile;

		// Download file
		ssZipFile << lpTempPathBuffer << _T("PythagorasCommonFiles.zip");
		ftp.m_pConnect->GetFile(_T("PythagorasCommonFiles.zip"), ssZipFile.str().c_str(), TRUE, FILE_ATTRIBUTE_NORMAL, FTP_TRANSFER_TYPE_BINARY, 1);

		// Now unzip it
		UnzipFile(ssZipFile.str().c_str(), ssPythagorasFolder.str().c_str());

		// Delete it
		DeleteFile(ssZipFile.str().c_str());

		AddToOutput(eStandardOutput, _T("Common file download complete"), true);
	}
	catch (std::exception &ex)
	{
		std::wstringstream ssMessage;
		ssMessage << _T("Caught exception in DownloadCommonFiles(): [Message] ") << ex.what();
		::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);
		bReturn = false;
	}
	catch (CInternetException* pEx)
	{
		std::wstringstream ssMessage;
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);		
		ssMessage << _T("Caught InternetException in DownloadCommonFiles(): [Message] ") << sz << _T(" [Code] ") << pEx->m_dwError;
		::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);		
		pEx->Delete();
		bReturn = false;
	}

	return bReturn;
}

void CPythagorasApp::GetScriptsFolders(std::wstring &sCommonFolder, std::wstring &sLocalFolder)
{
	wchar_t* localMyDocuments = 0;
	std::wstringstream ssTemplates;
	std::wstringstream ssTemplatesCommon;
	std::wstringstream ssTemplatesLocal;
	std::wstringstream ssCommon;
	std::wstringstream ssLocal;
	std::wstringstream ssPythagoras;
	std::wstringstream ssScripts;
	std::wstringstream ssTemp;
	std::wstringstream ssDelete;
	std::wstringstream ssFileName;
	std::wstringstream PyCache;
	std::wstring sAppDirectory;
	const SECURITY_ATTRIBUTES *psa = NULL;
	CMemBuffer mbTutorialScript, mbInjectScript, mbInspectScript, mbHTMLTemplate;
	TCHAR lpTempPathBuffer[MAX_PATH];	
		
	mbTutorialScript.InitFromResource(IDR_TUTORIAL_SCRIPT);
	mbInjectScript.InitFromResource(IDR_INJECT_SCRIPT);
	mbInspectScript.InitFromResource(IDR_INSPECT_SCRIPT);
	mbHTMLTemplate.InitFromResource(IDR_DEFAULT_HTML_TEMPLATE);
		
	{
		DWORD dwRetVal = GetTempPath(MAX_PATH, lpTempPathBuffer);
		if (dwRetVal > MAX_PATH || (dwRetVal == 0))
		{
			GetAppDirectory(sAppDirectory);
			ssTemp.str(_T(""));
			ssTemp << sAppDirectory.c_str() << _T("Temp");
		}
		else
		{
			ssTemp.str(_T(""));
			ssTemp << lpTempPathBuffer << _T("PythagorasTemp");
		}
		
		// First thing, delete any existing temp folder underneath our application folder
		ssDelete.str(_T(""));
		ssDelete.write(ssTemp.str().c_str(), ssTemp.str().size());
		std::error_code ec;
		try
		{			
			std::filesystem::remove_all(ssDelete.str().c_str(), ec);
		}
		catch (...)
		{
		}
		
		// Now remake it
		_wmkdir(ssTemp.str().c_str());

		PyCache.str(_T(""));
		PyCache.write(ssTemp.str().c_str(), ssTemp.str().size());
		PyCache << _T("\\PyCache");
		_wmkdir(PyCache.str().c_str());
		m_sTempFolder = PyCache.str().c_str();

		ssScripts.str(_T(""));
		ssScripts.write(ssTemp.str().c_str(), ssTemp.str().size());
		ssScripts << _T("\\Scripts");
		_wmkdir(ssScripts.str().c_str());

		ssLocal.str(_T(""));
		ssLocal.write(ssScripts.str().c_str(), ssScripts.str().size());
		ssLocal << _T("\\Local");
		_wmkdir(ssLocal.str().c_str());
		sLocalFolder = ssLocal.str();
		m_sTempLocalScriptFolder = sLocalFolder;
		
		ssFileName.str(_T(""));
		ssFileName.write(ssLocal.str().c_str(), ssLocal.str().size());
		ssFileName << _T("\\PythagorasInject.py");
		mbInjectScript.WriteToFile(ssFileName);
		ssFileName.str(_T(""));
		ssFileName.write(ssLocal.str().c_str(), ssLocal.str().size());
		ssFileName << _T("\\PythagorasInspect.py");		
		mbInspectScript.WriteToFile(ssFileName);
		
		ssCommon.str(_T(""));
		ssCommon.write(ssScripts.str().c_str(), ssScripts.str().size());
		ssCommon << _T("\\Common");
		_wmkdir(ssCommon.str().c_str());
		sCommonFolder = ssCommon.str();
		m_sTempCommonScriptFolder = sCommonFolder;

		ssFileName.str(_T(""));
		ssFileName.write(ssCommon.str().c_str(), ssCommon.str().size());
		ssFileName << _T("\\PythagorasInject.py");
		m_sInjectScript = ssFileName.str();
		mbInjectScript.WriteToFile(m_sInjectScript);		
		
		ssFileName.str(_T(""));
		ssFileName.write(ssCommon.str().c_str(), ssCommon.str().size());
		ssFileName << _T("\\PythagorasInspect.py");
		m_sInspectScript = ssFileName.str();
		mbInspectScript.WriteToFile(m_sInspectScript);		
		
		ssFileName.str(_T(""));
		ssFileName.write(ssCommon.str().c_str(), ssCommon.str().size());
		ssFileName << _T("\\DefaultTemplate.html");
		m_sTemplateScript = ssFileName.str();
		m_sDefaultHTMLTemplate = ssFileName.str();
		mbHTMLTemplate.WriteToFile(m_sDefaultHTMLTemplate);			
	}

	if (!SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Documents, 0, NULL, &localMyDocuments)))
	{
		AfxMessageBox(_T("Unable to find MyDocuments folder"));
		return;
	}

	// Make the Local documents directory structure
	switch (m_RegistrySettings.m_iLocalScriptsOption)	
	{
	case 0: // MyDocuments
	{
		ssPythagoras.str(_T(""));
		ssPythagoras << localMyDocuments << _T("\\Pythagoras");
		_wmkdir(ssPythagoras.str().c_str());

		ssScripts.str(_T(""));
		ssScripts.write(ssPythagoras.str().c_str(), ssPythagoras.str().size());
		ssScripts << _T("\\Scripts");
		_wmkdir(ssScripts.str().c_str());

		ssLocal.str(_T(""));
		ssLocal.write(ssScripts.str().c_str(), ssScripts.str().size());
		ssLocal << _T("\\Local");
		_wmkdir(ssLocal.str().c_str());
		m_sTrueLocalScriptFolder = ssLocal.str().c_str();
	}
	break;

	case 1: // Custom folder
	{
		m_sTrueLocalScriptFolder = m_RegistrySettings.m_sLocalScriptsFolder.c_str();		
	}
	break;
	}
	
	// Create the Tutorial.py script in the local script folder
	ssFileName.str(_T(""));
	ssFileName.write(m_sTrueLocalScriptFolder.c_str(), m_sTrueLocalScriptFolder.size());
	ssFileName << _T("\\Tutorial.py");
	if (AllowAdmin())
	{
		// Overwrite it every time, so we can just edit the one in our solution folder
		mbTutorialScript.WriteToFile(ssFileName);
	}
	else
	{
		mbTutorialScript.WriteToFile(ssFileName, MEMBUFFER_FLAG_WRITEFILE_SKIP_IF_EXISTS);
	}
	
	switch (m_RegistrySettings.m_iCommonScriptsOption)
	{
	case 0: // MyDocuments
	{
		ssPythagoras.str(_T(""));
		ssPythagoras << localMyDocuments << _T("\\Pythagoras");
		_wmkdir(ssPythagoras.str().c_str());

		ssScripts.str(_T(""));
		ssScripts.write(ssPythagoras.str().c_str(), ssPythagoras.str().size());
		ssScripts << _T("\\Scripts");
		_wmkdir(ssScripts.str().c_str());

		ssCommon.str(_T(""));
		ssCommon.write(ssScripts.str().c_str(), ssScripts.str().size());
		ssCommon << _T("\\Common");
		_wmkdir(ssCommon.str().c_str());
		m_sTrueCommonScriptFolder = ssCommon.str().c_str();
	}
	break;

	case 1: // Custom folder
	{
		m_sTrueCommonScriptFolder = m_RegistrySettings.m_sCommonScriptsFolder.c_str();
	}
	break;
	}

	// Create the DefaultTemplate.html file in our common folder, but only if none currently exists.
	ssFileName.str(_T(""));
	ssFileName.write(m_sTrueCommonScriptFolder.c_str(), m_sTrueCommonScriptFolder.size());
	ssFileName << _T("\\DefaultTemplate.html");
	mbHTMLTemplate.WriteToFile(ssFileName, MEMBUFFER_FLAG_WRITEFILE_SKIP_IF_EXISTS);

	if (m_RegistrySettings.m_iDownloadCommon != 0)
	{
		// Download files from FTP server into a temporary path
		TCHAR lpTempPathBuffer[MAX_PATH];
		DWORD dwRetVal = GetTempPath(MAX_PATH, lpTempPathBuffer);
		if (dwRetVal > MAX_PATH || (dwRetVal == 0))
		{
			::AfxMessageBox(_T("GetTempPath() failed"));
			return;
		}
		ssPythagoras.str(_T(""));
		ssPythagoras << lpTempPathBuffer << _T("Pythagoras");
		// Delete anything presently in that folder
		std::error_code ec;
		try
		{
			std::filesystem::remove_all(ssPythagoras.str().c_str(), ec);
		}
		catch (...)
		{
		}
		_wmkdir(ssPythagoras.str().c_str());
		
		DownloadCommonFiles(ssPythagoras);
		std::filesystem::copy(ssPythagoras.str().c_str(), m_sTempCommonScriptFolder.c_str(), std::filesystem::copy_options::recursive | std::filesystem::copy_options::update_existing);		
	}
	
	CoTaskMemFree(static_cast<void*>(localMyDocuments));
	CopyLocalAndCommonScripts(sCommonFolder, sLocalFolder);
}

void CPythagorasApp::CopyLocalAndCommonScripts(std::wstring& sCommonFolder, std::wstring& sLocalFolder)
{
	// Copy over all Python files from the true local script folder to our temp script folder
	std::wstring sWildCard;
	std::filesystem::path source;
	std::filesystem::path target;
	_StringVector FileList;

	try {
		
		//GetFilesInDirectory(_T("*.py"), FileList, m_sTrueLocalScriptFolder.c_str(), true);
		CopyFilesByWildCard(m_sTrueLocalScriptFolder.c_str(), sLocalFolder.c_str(), _T("*.py"), true);
		CopyFilesByWildCard(m_sTrueCommonScriptFolder.c_str(), sCommonFolder.c_str(), _T("*.py"), true);
		CopyFilesByWildCard(m_sTrueLocalScriptFolder.c_str(), sLocalFolder.c_str(), _T("*.html"), true);
		CopyFilesByWildCard(m_sTrueCommonScriptFolder.c_str(), sCommonFolder.c_str(), _T("*.html"), true);
		/*
		source = m_sTrueLocalScriptFolder.c_str();
		target = sLocalFolder.c_str();
		std::filesystem::copy(source, target, std::filesystem::copy_options::recursive | std::filesystem::copy_options::update_existing);

		source = m_sTrueCommonScriptFolder.c_str();
		target = sCommonFolder.c_str();
		std::filesystem::copy(source, target, std::filesystem::copy_options::recursive | std::filesystem::copy_options::update_existing);
		*/
	}
	catch (std::exception& ex) { 
		std::wstringstream ssMessage;
		ssMessage << _T("Caught exception in CopyLocalAndCommonScripts(): [Message] ") << ex.what();
		::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);		
	}	
}


void CPythagorasApp::ParsePythonScripts(std::wstring &sScriptFolder, std::wstring& sTrueScriptFolder, _vScriptFunction &vScriptFunction, eScriptType ScriptType)
{
	// Go through all of the Python scripts in the "Scripts" folder, and find the functions defined therein
	std::wstring sWildCard(L"*.py");
	_StringVector FileList;
	_itString itString;
	TCHAR Scriptdir[_MAX_DIR];
	TCHAR dir[_MAX_DIR];
	TCHAR fname[_MAX_FNAME];
	TCHAR *SubPath = NULL;
	bool bHidden = false;
	std::wstringstream ssTruePathToScript;

	GetFilesInDirectory(sWildCard, FileList, sScriptFolder, true);

	if (FileList.size() == 0)
	{
		return;
	}
		
	_wsplitpath_s(sScriptFolder.c_str(), NULL, 0, Scriptdir, _MAX_DIR, NULL, 0, NULL, 0);
	// Remove the script with our inspect function
	for (itString = FileList.begin(); itString != FileList.end(); )
	{
		bHidden = false;
		_wsplitpath_s((*itString).c_str(), NULL, 0, dir, _MAX_DIR, fname, _MAX_FNAME, NULL, 0);
				
		// Remove any files that begin with "Pythagoras"
		if (!_tcsnicmp(fname, _T("Pythagoras"), lstrlenW(_T("Pythagoras"))))
		{
			bHidden = true;
			
#if 0
			// Look specifically for our inspect and inject scripts
			if (!_tcsicmp(fname, _T("PythagorasInspect")))
			{
				if (ScriptType == ScriptTypeCommon)
				{
					m_sInspectScript = (*itString);
				}
			}

			if (!_tcsicmp(fname, _T("PythagorasInject")))
			{
				if (ScriptType == ScriptTypeCommon)
				{
					m_sInjectScript = (*itString);
				}
			}

			if (!_tcsicmp(fname, _T("PythagorasTemplate")))
			{
				if (ScriptType == ScriptTypeCommon)
				{
					m_sTemplateScript = (*itString);
				}
			}
#endif
			
			/*if (!theApp.AllowAdmin())*/
			{
				itString = FileList.erase(itString);
				continue;
			}			
		}

		// Remove any files that begin with an underscore
		if (!_tcsnicmp(fname, _T("_"), 1))
		{
			bHidden = true;
			if (!theApp.AllowAdmin())
			{
				itString = FileList.erase(itString);
				continue;
			}
		}

		LPScriptFunctionStruct lpNew = new ScriptFunctionStruct;

		// Remove the common folder from the path to the file
		SubPath = dir + _tcslen(Scriptdir);
		std::wstring sRelative(SubPath);

		ssTruePathToScript.str(_T(""));
		ssTruePathToScript << sTrueScriptFolder.c_str() << (*itString).substr(sScriptFolder.length(), std::string::npos);		
		lpNew->sTruePathToScript = ssTruePathToScript.str().c_str();

		// Now replace all slashes with .
		for( std::wstring::iterator it = sRelative.begin(); it != sRelative.end(); ++it)
		{
			if ((*it) == '\\' || (*it) == '/')
			{
				(*it) = '.';
			}
		}
		lpNew->sRelativePath = sRelative;
		lpNew->sRelativePath += fname;
		lpNew->sPathToScript = (*itString);
		lpNew->ScriptType = ScriptType;
		lpNew->bErrors = false;
		lpNew->bHidden = bHidden;
		vScriptFunction.push_back(lpNew);
		itString++;
	}
}

void CPythagorasApp::CompileScripts(bool bDeleteMemory /* = true */)
{
	using namespace boost::interprocess;
	using namespace PythonEngine;

	class shm_remove
	{
	public:
		bool m_bRemove;
		shm_remove(bool bRemove) : m_bRemove(bRemove) { if (bRemove) shared_memory_object::remove(cSharedMemoryString); }
		~shm_remove() { if (m_bRemove) shared_memory_object::remove(cSharedMemoryString); }
	};

	shm_remove remove(bDeleteMemory);

	try
	{
		const int MEMORY_GROW_SIZE = 1048576;

		{
			wmanaged_shared_memory segment(create_only, cSharedMemoryString, MEMORY_GROW_SIZE);

			// An allocator convertible to any allocator<T, segment_manager_t> type
			void_allocator alloc_inst(segment.get_segment_manager());

			// Construct the shared memory map and fill them
			_SharedExecuteOptions *ExecuteOptions = segment.construct<_SharedExecuteOptions>(cExecuteOptionsString)(alloc_inst);
			ExecuteOptions->m_ProcessMode = ePMCompileScript;
			ExecuteOptions->m_InspectFile = m_sInspectScript.c_str();
			ExecuteOptions->m_InjectFile = m_sInjectScript.c_str();
			ExecuteOptions->m_DefaultHTMLTemplate = m_sDefaultHTMLTemplate.c_str();
			ExecuteOptions->m_LocalScriptsFolder = m_sTrueLocalScriptFolder.c_str();
			ExecuteOptions->m_CommonScriptsFolder = m_sTrueCommonScriptFolder.c_str();
			ExecuteOptions->m_TempFolder = m_sTempFolder.c_str();
			ExecuteOptions->m_lHWND = reinterpret_cast<unsigned long>(GetMainFrame()->m_hWnd);
			ExecuteOptions->m_lMessageID = CM_ENGINE_CALLBACK;

			_SharedCompileScriptVector *CompileScriptVector = segment.construct<_SharedCompileScriptVector>(cCompileScriptString)(alloc_inst);
		}
		// At this point the segment will be unmapped since the segment object has gone out of scope.
		// Grow the segment if the amount of free memory has dipped below a certain amount
		__int64 iFreeMemory = -1;
		__int64 iAdd = 0;
		int index = 0;
		for (_itScriptFunction itScriptFunction = m_vScriptFunction.begin(); itScriptFunction != m_vScriptFunction.end(); itScriptFunction++, index++)
		{
			iAdd = (*itScriptFunction)->sPathToScript.size() * 4;
			if (iFreeMemory != -1 && iFreeMemory < iAdd)
			{
				managed_shared_memory::grow(cSharedMemoryString, MEMORY_GROW_SIZE);
			}

			// Map into the shared memory
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);

			// An allocator convertible to any allocator<T, segment_manager_t> type
			void_allocator alloc_inst(segment.get_segment_manager());

			_SharedCompileScriptVector *CompileScriptVector = segment.find<_SharedCompileScriptVector>(cCompileScriptString).first;

			_SharedCompileScript Record(alloc_inst);
			Record.m_PathToFile = (*itScriptFunction)->sPathToScript.c_str();
			Record.m_Index = index;

			CompileScriptVector->push_back(Record);
			iFreeMemory = segment.get_free_memory();
		}

		if (bDeleteMemory)
		{
			// Launch child process			
			StartPythonEngine(m_ssPythonEngPath.str().c_str());
			
			// Map into the shared memory			
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);
			
			_SharedCompileScriptVector *CompileScriptVector = segment.find<_SharedCompileScriptVector>(cCompileScriptString).first;
			_SharedCompileScriptIterator itScript = CompileScriptVector->begin();

			while (itScript != CompileScriptVector->end())
			{
				std::wstring sError;
				ConvertString((*itScript).m_ErrorMessage, sError);

				if (sError.size() == 0)
				{
					m_vScriptFunction.at((*itScript).m_Index)->bErrors = false;

					_SharedStringIterator itFunction;
					_SharedStringIterator itDescription;
					_SharedStringIterator itDisplayName;
															
					for (
						itFunction = (*itScript).m_Functions.begin(), itDescription = (*itScript).m_Descriptions.begin(), itDisplayName = (*itScript).m_DisplayNames.begin();
						itFunction != (*itScript).m_Functions.end() && itDescription != (*itScript).m_Descriptions.end() && itDisplayName != (*itScript).m_DisplayNames.end();
						itFunction++, itDescription++, itDisplayName++
						)
					{
						LPFunctionStruct lpNew = new FunctionStruct;
						ConvertString((*itFunction).m_String, lpNew->sFunctionName);
						ConvertString((*itDescription).m_String, lpNew->sDescription);
						ConvertString((*itDisplayName).m_String, lpNew->sDisplayName);

						m_vScriptFunction.at((*itScript).m_Index)->vFunctions.push_back(lpNew);						
					}
				}
				else
				{
					m_vScriptFunction.at((*itScript).m_Index)->bErrors = true;
					m_vScriptFunction.at((*itScript).m_Index)->ssErrorMessage << sError.c_str();
				}
				itScript++;
			}

			// Check if child has destroyed the vector, if not, destroy it
			if (segment.find<_SharedCompileScriptVector>(cCompileScriptString).first)
			{
				segment.destroy<_SharedCompileScriptVector>(cCompileScriptString);
			}

			if (segment.find<_SharedExecuteOptions>(cExecuteOptionsString).first)
			{
				segment.destroy<_SharedExecuteOptions>(cExecuteOptionsString);
			}
		}
	}
	catch (interprocess_exception &ex)
	{
		std::wstringstream sOutput;
		std::wstring sWhat = CA2W(ex.what());
		sOutput << _T("CPythagorasApp::CompileScripts()\nCaught interprocess_exception.\nCode=") << ex.get_error_code() << _T("\nMessage=") << sWhat.c_str();
		AfxMessageBox(sOutput.str().c_str());
	}
}

void CPythagorasApp::ExecuteScript(bool bDeleteMemory /* = true */)
{
	ExecuteScript(m_vFileFolderData, bDeleteMemory);
}

void CPythagorasApp::ExecuteScript(_vLPFileFolderData &vFileFolderData, bool bDeleteMemory /* = true */)
{	
	bool bClearOutput = m_RegistrySettings.m_iClearOutput == 0 ? false : true;
	std::wstring sInjectScript(_T(""));
	if (m_dwSelectedFunction == -1 || !m_lpSelectedScript)
	{
		AfxMessageBox(_T("Please select a function to execute."));
		return;
	}

	DWORD dwWait = WaitForSingleObject(m_hThreadMutex, 500);
	if (dwWait != WAIT_OBJECT_0)
	{
		int iReturn = AfxMessageBox(_T("A script is currently running, do you want to add this new call to the queue?"), MB_YESNO);
		if (iReturn == IDNO)
		{
			return;
		}
		// Reset this; we're going to create a locking thread that will wait for the prior ones to finish.  Don't clear
		// the output in this case 
		bClearOutput = false;
	}
	else
	{
		ReleaseMutex(m_hThreadMutex);
	}

	//CheckAndClearOutput();
	GetView()->m_pcScriptExecution.ShowWindow(SW_SHOW);

	// Copy over the latest scripts, as they might be editing on the fly
	CopyLocalAndCommonScripts(m_sTempCommonScriptFolder, m_sTempLocalScriptFolder);
	if (AllowAdmin())
	{
		std::wstringstream ssPath;
		if (m_lpSelectedScript->ScriptType == ScriptTypeLocal)
		{
			ssPath << m_sTempLocalScriptFolder.c_str() << _T("\\PythagorasInject.py");
		}
		else
		{
			ssPath << m_sTempCommonScriptFolder.c_str() << _T("\\PythagorasInject.py");
		}
		sInjectScript = ssPath.str().c_str();
	}
	else
	{
		sInjectScript = m_sInjectScript;
	}
	
	LPFunctionStruct lpFunction = m_lpSelectedScript->vFunctions.at(m_dwSelectedFunction);	
	
	unsigned int iThreadID = 0;
	LPExecuteScriptThreadParams lpThreadParams = new ExecuteScriptThreadParams;
	lpThreadParams->bDeleteMemory = bDeleteMemory;	
	lpThreadParams->bClearOutput = bClearOutput;
	lpThreadParams->sFunctionName = lpFunction->sFunctionName;
	lpThreadParams->sPathToScript = m_lpSelectedScript->sPathToScript;
	lpThreadParams->sRelativePath = m_lpSelectedScript->sRelativePath;
	lpThreadParams->sInjectScript = sInjectScript;
	lpThreadParams->sDefaultHTMLTemplate = m_sDefaultHTMLTemplate;
	lpThreadParams->sLocalScriptsFolder = m_sTrueLocalScriptFolder;
	lpThreadParams->sCommonScriptsFolder = m_sTrueCommonScriptFolder;
	lpThreadParams->sTempFolder = m_sTempFolder;
	lpThreadParams->hMutex = m_hThreadMutex;
	lpThreadParams->pMainFrame = (CMainFrame *)m_pMainWnd;
	lpThreadParams->sPythonEnginePath = m_ssPythonEngPath.str();

	for (LPFileFolderDataStruct lpSource : vFileFolderData)
	{
		if (lpSource->Type == FileType)
		{
			lpThreadParams->vFileData.push_back(lpSource->sFullName);
		}
		else
		{
			lpThreadParams->vFolderData.push_back(lpSource->sFullName);
		}		
	}
	
	HANDLE hThread = (HANDLE)_beginthreadex(NULL, 0, &CPythagorasApp::ExecuteScriptThread, (void *)lpThreadParams, 0, &iThreadID);
	if (hThread)
	{
		m_ThreadHandles.push_back(hThread);
	}
	CheckScriptThreads(500,true);
}

void CPythagorasApp::CheckScriptThreads(DWORD dwTime /* = 500 */, bool bUpdateProgress /* = false */ )
{
	for (std::vector<HANDLE>::iterator itThread = m_ThreadHandles.begin(); itThread != m_ThreadHandles.end(); )
	{
		// Wait to see if the thread has finished... if so, remove it from our queue
		DWORD dwWait = WaitForSingleObject(*itThread, dwTime);
		if (dwWait == WAIT_OBJECT_0)
		{
			CloseHandle(*itThread);
			itThread = m_ThreadHandles.erase(itThread);
		}
		else
		{
			itThread++;
		}
	}

	if(m_ThreadHandles.size() == 0 && bUpdateProgress)
	{
		GetView()->m_pcScriptExecution.ShowWindow(SW_HIDE);				
	}	
}

/* static */unsigned __stdcall CPythagorasApp::ExecuteScriptThread(void *lpParams)
{	
	using namespace boost::interprocess;
	using namespace PythonEngine;
	LPExecuteScriptThreadParams lpThreadParams = reinterpret_cast<LPExecuteScriptThreadParams>(lpParams);

	if (!lpThreadParams)
	{
		return 0;
	}
	
	// Try to obtain ownership of the mutex... wait indefinitely until we can own it.  Necessary to prevent
	// multiple threads from attempting to write to the same interprocess memory space
	DWORD dwWait = WaitForSingleObject( lpThreadParams->hMutex, INFINITE);
	if (dwWait != WAIT_OBJECT_0)
	{
		if (dwWait == WAIT_FAILED)
		{
			LPVOID lpMsgBuf;
			DWORD dw = ::GetLastError();
			FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL, dw, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPTSTR)&lpMsgBuf, 0, NULL);

			// Display the error
			std::wstringstream ssError;
			ssError << _T("WaitForSingleObject() failed\nMessage=") << (LPTSTR)lpMsgBuf;
			::MessageBox(NULL, ssError.str().c_str(), NULL, MB_ICONSTOP);
		}
		return 0;
	}

	if (lpThreadParams->bClearOutput)
	{
		::PostMessage(lpThreadParams->pMainFrame->m_hWnd, CM_EXECUTETHREAD, eETWClearOutput, 0);
	}

	class shm_remove
	{
	public:
		LPExecuteScriptThreadParams m_lpThreadParams;
		shm_remove(LPExecuteScriptThreadParams lpThreadParams) : m_lpThreadParams(lpThreadParams)
		{
			if (lpThreadParams->bDeleteMemory) shared_memory_object::remove(cSharedMemoryString);
		}
		~shm_remove()
		{
			if (m_lpThreadParams->bDeleteMemory) shared_memory_object::remove(cSharedMemoryString);
			ReleaseMutex(m_lpThreadParams->hMutex);
			delete m_lpThreadParams;
		}
	};

	shm_remove remove(lpThreadParams);

	try
	{
		const int MEMORY_GROW_SIZE = 33554432;

		{
			// Create shared memory.  33554432 (32 MB) should be enough for ~144k files, far more than
			// would ever be expected
			wmanaged_shared_memory segment(create_only, cSharedMemoryString, MEMORY_GROW_SIZE);

			// An allocator convertible to any allocator<T, segment_manager_t> type
			void_allocator alloc_inst(segment.get_segment_manager());

			// Construct the shared memory map and fill them
			_SharedExecuteOptions *ExecuteOptions = segment.construct<_SharedExecuteOptions>(cExecuteOptionsString)(alloc_inst);
			ExecuteOptions->m_ProcessMode = ePMCallFunction;
			ExecuteOptions->m_Function = lpThreadParams->sFunctionName.c_str();
			ExecuteOptions->m_Script = lpThreadParams->sPathToScript.c_str();
			ExecuteOptions->m_ScriptRegistry = lpThreadParams->sRelativePath.c_str();
			ExecuteOptions->m_InjectFile = lpThreadParams->sInjectScript.c_str();
			ExecuteOptions->m_DefaultHTMLTemplate = lpThreadParams->sDefaultHTMLTemplate.c_str();
			ExecuteOptions->m_LocalScriptsFolder = lpThreadParams->sLocalScriptsFolder.c_str();
			ExecuteOptions->m_CommonScriptsFolder = lpThreadParams->sCommonScriptsFolder.c_str();
			ExecuteOptions->m_TempFolder = lpThreadParams->sTempFolder.c_str();			
			ExecuteOptions->m_lHWND = reinterpret_cast<unsigned long>(lpThreadParams->pMainFrame->m_hWnd);
			ExecuteOptions->m_lMessageID = CM_ENGINE_CALLBACK;
			HWND hTy = NULL;
			
			_SharedFileArray *FileNameArray = segment.construct<_SharedFileArray>(cFilenameArrayString)(alloc_inst);
			_SharedFolderArray *FolderNameArray = segment.construct<_SharedFolderArray>(cFoldernameArrayString)(alloc_inst);
			_SharedScriptCallback *Callback = segment.construct<_SharedScriptCallback>(cScriptCallbackString)(alloc_inst);
			_SharedStringArray *StringArray = segment.construct<_SharedStringArray>(cStringArrayString)(alloc_inst);
		}
		// At this point the segment will be unmapped since the segment object has gone out of scope.
		// Grow the segment if the amount of free memory has dipped below a certain amount
		__int64 iFreeMemory = -1;
		__int64 iAdd = 0;
		for (std::wstring sFullName : lpThreadParams->vFileData )
		{
			iAdd = sFullName.size() * 4;
			if (iFreeMemory != -1 && iFreeMemory < iAdd)
			{
				managed_shared_memory::grow(cSharedMemoryString, MEMORY_GROW_SIZE);
			}

			// Map into the shared memory
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);

			// An allocator convertible to any allocator<T, segment_manager_t> type
			void_allocator alloc_inst(segment.get_segment_manager());
			_SharedFileArray *FileNameArray = segment.find<_SharedFileArray>(cFilenameArrayString).first;
			
			_SharedString Record(alloc_inst);			
			Record.m_String = sFullName.c_str();
			FileNameArray->m_Filenames.push_back(Record);
			iFreeMemory = segment.get_free_memory();
		}

		iFreeMemory = -1;
		iAdd = 0;
		for (std::wstring sFullName : lpThreadParams->vFolderData)
		{
			iAdd = sFullName.size() * 4;
			if (iFreeMemory != -1 && iFreeMemory < iAdd)
			{
				managed_shared_memory::grow(cSharedMemoryString, MEMORY_GROW_SIZE);
			}

			// Map into the shared memory
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);

			// An allocator convertible to any allocator<T, segment_manager_t> type
			void_allocator alloc_inst(segment.get_segment_manager());
			_SharedFolderArray *FolderNameArray = segment.find<_SharedFolderArray>(cFoldernameArrayString).first;

			_SharedString Record(alloc_inst);
			Record.m_String = sFullName.c_str();
			FolderNameArray->m_Foldernames.push_back(Record);
			iFreeMemory = segment.get_free_memory();
		}

		if (lpThreadParams->bDeleteMemory)
		{
			std::wstringstream ssOutput;
			
			ssOutput << _T("Calling function [") << lpThreadParams->sFunctionName.c_str() << _T("] in script [") << lpThreadParams->sPathToScript.c_str() << _T("]");
			lpThreadParams->pMainFrame->m_wndOutput.AddToOutput(eStandardOutput, ssOutput.str().c_str());

			// Launch child process
			StartPythonEngine(lpThreadParams->sPythonEnginePath.c_str());
			
			// Map into the shared memory
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);
			// An allocator convertible to any allocator<T, segment_manager_t> type
			void_allocator alloc_inst(segment.get_segment_manager());

			// Construct the shared memory map and fill them
			_SharedExecuteOptions *ExecuteOptions = segment.find<_SharedExecuteOptions>(cExecuteOptionsString).first;
			_SharedStringIterator itString;

			for (itString = ExecuteOptions->m_OutputVector.begin(); itString != ExecuteOptions->m_OutputVector.end(); itString++)
			{
				std::wstring sValue;
				ConvertString((*itString).m_String, sValue);
				// Can't select the active tab from this thread, triggers ASSERT() errors
				lpThreadParams->pMainFrame->m_wndOutput.AddToOutput(eStandardOutput, sValue, false);				
			}

			for (itString = ExecuteOptions->m_ErrorVector.begin(); itString != ExecuteOptions->m_ErrorVector.end(); itString++)
			{
				std::wstring sValue;
				ConvertString((*itString).m_String, sValue);
				// Can't select the active tab from this thread, triggers ASSERT() errors
				lpThreadParams->pMainFrame->m_wndOutput.AddToOutput(eStandardError, sValue, false);
			}

			ssOutput.str(_T(""));
			ssOutput << _T("Script finished in ") << ExecuteOptions->m_ExecutionTime << _T(" seconds");
			lpThreadParams->pMainFrame->m_wndOutput.AddToOutput(eStandardOutput, ssOutput.str().c_str(),false);

			if (ExecuteOptions->m_ErrorVector.size() > 0)
			{
				// Send a message to the caller to set the focus onto the error output pane
				::PostMessage(lpThreadParams->pMainFrame->m_hWnd, CM_EXECUTETHREAD, eETWSelectPane, eStandardError);
			}
			else
			{
				// Send a message to the caller to set the focus onto the standard output pane
				::PostMessage(lpThreadParams->pMainFrame->m_hWnd, CM_EXECUTETHREAD, eETWSelectPane, eStandardOutput);
			}


			// Check if child has destroyed the vector, if not, destroy it
			if (segment.find<_SharedFileArray>(cFilenameArrayString).first)
			{
				segment.destroy<_SharedFileArray>(cFilenameArrayString);
			}
			if (segment.find<_SharedFolderArray>(cFoldernameArrayString).first)
			{
				segment.destroy<_SharedFolderArray>(cFoldernameArrayString);
			}
			if (segment.find<_SharedExecuteOptions>(cExecuteOptionsString).first)
			{
				segment.destroy<_SharedExecuteOptions>(cExecuteOptionsString);
			}
			if (segment.find<_SharedScriptCallback>(cScriptCallbackString).first)
			{
				segment.destroy<_SharedScriptCallback>(cScriptCallbackString);
			}
			if (segment.find<_SharedStringArray>(cStringArrayString).first)
			{
				segment.destroy<_SharedStringArray>(cStringArrayString);
			}			
		}
	}
	catch (interprocess_exception &ex)
	{
		std::wstringstream sOutput;
		sOutput << _T("Caught interprocess_exception.\nCode=") << ex.get_error_code() << _T("\nMessage=") << ex.what();
		AfxMessageBox(sOutput.str().c_str());
	}

	::PostMessage(lpThreadParams->pMainFrame->m_hWnd, CM_EXECUTETHREAD, eETWThreadFinished, 0);
	return 0;
}

void CPythagorasApp::CheckAndClearOutput()
{
	if (m_RegistrySettings.m_iClearOutput)
	{
		CMainFrame *pMainFrame = (CMainFrame *)m_pMainWnd;
		pMainFrame->m_wndOutput.SetForegroundWindow();
		pMainFrame->m_wndOutput.ClearOutput();
	}
}

void CPythagorasApp::AddToOutput(OutputTypeEnum eOutputType, LPCTSTR szText, bool bSelect /* = false */)
{
	CMainFrame *pMainFrame = (CMainFrame *)m_pMainWnd;
	pMainFrame->m_wndOutput.SetForegroundWindow();
	pMainFrame->m_wndOutput.AddToOutput(eOutputType, szText, bSelect);
}

void CPythagorasApp::SetOutput(OutputTypeEnum eOutputType, LPCTSTR szText, bool bSelect /* = false */)
{
	CMainFrame *pMainFrame = (CMainFrame *)m_pMainWnd;
	pMainFrame->m_wndOutput.SetForegroundWindow();
	pMainFrame->m_wndOutput.SetOutput(eOutputType, szText, bSelect);
}

/*static */bool CPythagorasApp::StartPythonEngine(LPCTSTR szPath)
{
	PROCESS_INFORMATION processInformation = { 0 };
	STARTUPINFO startupInfo = { 0 };
	startupInfo.cb = sizeof(startupInfo);	
	bool bReturn = false;
	std::wstringstream ssError;
		
	// Create the process
	CString csPath = szPath;
	BOOL result = CreateProcess(NULL, csPath.GetBuffer(),
		NULL, NULL, FALSE,
		NORMAL_PRIORITY_CLASS | CREATE_NO_WINDOW,
		NULL, NULL, &startupInfo, &processInformation);

	if (!result)
	{
		// CreateProcess() failed
		// Get the error from the system
		LPVOID lpMsgBuf;
		DWORD dw = ::GetLastError();
		FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
			NULL, dw, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPTSTR)&lpMsgBuf, 0, NULL);

		// Display the error
		ssError.str(_T(""));
		ssError << _T("CPythagorasApp::StartPythonEngine()\nFailed to start PythonEngine.\nCommand=") << szPath << _T("\nMessage=") << (LPTSTR)lpMsgBuf;
		::MessageBox(NULL, ssError.str().c_str(), NULL, MB_ICONSTOP);

		// Free resources created by the system
		LocalFree(lpMsgBuf);		
	}
	else
	{
		// Successfully created the process.  Wait for it to finish.
		WaitForSingleObject(processInformation.hProcess, INFINITE);
		DWORD exitCode;

		// Get the exit code.
		result = GetExitCodeProcess(processInformation.hProcess, &exitCode);
		ssError.str(_T(""));
		ssError << _T("CPythagorasApp::StartPythonEngine()\nFailed to start PythonEngine.\nExit code = ") << std::hex << exitCode;
		
		switch (exitCode)
		{
		case 0:
		{
			bReturn = true;
		}
		break;
		case 0xC000007B:
		{
			// Something failed with starting the engine, this might be a Python error, or an issue with mixing/matching 32-bit and 64-bit
			ssError << _T("\nSTATUS_INVALID_IMAGE_FORMAT\nPotential 32-bit and 64-bit conflict.");			
		}
		break;
		default:
		{			
		}
		break;
		}

		if (!bReturn)
		{
			int iReturn = ::MessageBox(NULL, ssError.str().c_str(), NULL, MB_ICONSTOP | MB_YESNO);
			if (iReturn == IDYES)
			{

			}
		}
		
		// Close the handles.
		CloseHandle(processInformation.hProcess);
		CloseHandle(processInformation.hThread);

		bReturn = !result ? false : true;
	}
	return bReturn;
}

CPythagorasView *CPythagorasApp::GetView()
{
	CMainFrame *pMainFrame = (CMainFrame *)m_pMainWnd;
	return (CPythagorasView *)pMainFrame->GetActiveView();
}

void CPythagorasApp::AddFileFolder(LPCTSTR szPath)
{
	if (boost::filesystem::is_directory(szPath))
	{
		AddFolder(szPath);		
	}
	else
	{
		AddFile(szPath);
	}
}

void CPythagorasApp::AddFile(LPCTSTR szFileName)
{
#define MAXLENGTH 35

	// Check to see if this file is already in our list (use the map)
	_itDataMap itFile = m_DataMap.find(szFileName);
	if (itFile == m_DataMap.end())
	{
		TCHAR drive[_MAX_DRIVE];
		TCHAR dir[_MAX_DIR];
		TCHAR fname[_MAX_FNAME];
		_wsplitpath_s(szFileName, drive, _MAX_DRIVE, dir, _MAX_DIR, fname, _MAX_FNAME, NULL, 0);

		m_RegistrySettings.m_sExplorerPath = drive;
		m_RegistrySettings.m_sExplorerPath += dir;
		CString sShort;
		CString sAdd;

		sShort = szFileName;
		int l = sShort.GetLength();
		if (l > MAXLENGTH)
		{
			sAdd = L"...";
			sAdd.Append(sShort.Mid(l - MAXLENGTH, MAXLENGTH));
		}
		else
		{
			sAdd = sShort;
		}

		LPFileFolderDataStruct pNewFD = new FileFolderDataStruct;
		pNewFD->Type = FileType;
		pNewFD->sFullName = szFileName;
		pNewFD->sShortName = fname;
		pNewFD->sDisplayName = sAdd;
		pNewFD->bSelected = false;
		m_vFileFolderData.push_back(pNewFD);

		m_DataMap.insert(std::make_pair(szFileName, true));

		GetView()->RefreshFileList();		
	}	
}

void CPythagorasApp::AddFolder(LPCTSTR szFolderName)
{
#define MAXLENGTH 35

	// Check to see if this folder is already in our list (use the map)
	_itDataMap itData = m_DataMap.find(szFolderName);
	if (itData == m_DataMap.end())
	{
		_StringVector sv;
		std::wstring sExplorerPath;
		boost::split(sv, szFolderName, boost::is_any_of("\\/"));
		int i = 0;
		for (_itString it = sv.begin(); it != sv.end(); it++, i++)
		{
			if (i != sv.size() - 1)
			{
				sExplorerPath += (*it).c_str();
				sExplorerPath += _T("\\");
			}			
		}
		m_RegistrySettings.m_sExplorerPath = sExplorerPath.c_str();
		
		CString sShort;
		CString sAdd;

		sShort = szFolderName;
		int l = sShort.GetLength();
		if (l > MAXLENGTH)
		{
			sAdd = L"...";
			sAdd.Append(sShort.Mid(l - MAXLENGTH, MAXLENGTH));
		}
		else
		{
			sAdd = sShort;
		}

		LPFileFolderDataStruct pNewFD = new FileFolderDataStruct;
		pNewFD->Type = FolderType;
		pNewFD->sFullName = szFolderName;		
		pNewFD->sDisplayName = sAdd;
		pNewFD->bSelected = false;
		
		sv.erase(sv.begin(), sv.end());
		boost::split(sv, szFolderName, boost::is_any_of("\\/"));
		_ritString it = sv.rbegin();
		if (it != sv.rend())
		{
			pNewFD->sShortName = (*it).c_str();			
		}
		
		m_vFileFolderData.push_back(pNewFD);
		m_DataMap.insert(std::make_pair(szFolderName, true));

		GetView()->RefreshFileList();
	}
}

void CPythagorasApp::DeleteSharedMemory(bool bDisplayError /*= true*/)
{
	using namespace boost::interprocess;
	using namespace PythonEngine;

	class shm_remove
	{
	public:
		bool m_bRemove;
		shm_remove(bool bRemove) : m_bRemove(bRemove) { }
		~shm_remove() { if (m_bRemove) shared_memory_object::remove(cSharedMemoryString); }
	};

	shm_remove remove(true);

	try
	{				
		// Map into the shared memory
		wmanaged_shared_memory segment(open_only, cSharedMemoryString);
			
		// Check if child has destroyed the vector, if not, destroy it
		if (segment.find<_SharedFileArray>(cFilenameArrayString).first)
		{
			segment.destroy<_SharedFileArray>(cFilenameArrayString);
		}
		if (segment.find<_SharedExecuteOptions>(cExecuteOptionsString).first)
		{
			segment.destroy<_SharedExecuteOptions>(cExecuteOptionsString);
		}
	}
	catch (interprocess_exception &ex)
	{
		if (bDisplayError)
		{
			std::wstringstream sOutput;
			sOutput << _T("Caught interprocess_exception.\nCode=") << ex.get_error_code() << _T("\nMessage=") << ex.what();
			AfxMessageBox(sOutput.str().c_str());
		}		
	}
}

void CPythagorasApp::OnFileSettings()
{
	AppSettingsDlg dlg;
	int iReturn = dlg.DoModal();
	if (iReturn & APP_SETTINGS_RETURN_OK && iReturn & APP_SETTINGS_RETURN_MODIFIED_PATHS)
	{
		RefreshScripts();
	}
}


void CPythagorasApp::OnExplorerEdit()
{
	if (m_lpSelectedScript)
	{
		if (m_lpSelectedScript->ScriptType == ScriptTypeCommon && !AllowAdmin())
		{
			SetOutput(eStandardError, _T("Cannot edit common scripts; any changes would be overwritten."), true);
			return;
		}
		PROCESS_INFORMATION processInformation = { 0 };
		STARTUPINFO startupInfo = { 0 };
		startupInfo.cb = sizeof(startupInfo);
		bool bReturn = true;

		// Create the process
		CString csPath = _T("\"");
		csPath += theApp.m_RegistrySettings.m_sScriptEditor.c_str();		
		csPath += _T("\" \"");
		if (m_lpSelectedScript->ScriptType == ScriptTypeCommon)
		{
			std::wstring wsMessage;
			wsMessage = _T("NOTE: Editing file ");
			wsMessage += m_lpSelectedScript->sPathToScript;
			SetOutput(eStandardOutput, wsMessage.c_str(), true);
			csPath += m_lpSelectedScript->sPathToScript.c_str();
		}
		else
		{
			csPath += m_lpSelectedScript->sTruePathToScript.c_str();
		}
		csPath += _T("\"");
		BOOL result = CreateProcess(NULL, csPath.GetBuffer(),
			NULL, NULL, FALSE,
			NORMAL_PRIORITY_CLASS | CREATE_NO_WINDOW,
			NULL, NULL, &startupInfo, &processInformation);
		if (!result)
		{
			AfxMessageBox(_T("Could not start script editor.  Check settings for valid executable file."));
		}		
	}
}


void CPythagorasApp::OnOpenFolder()
{
	TCHAR drive[_MAX_DRIVE];
	TCHAR dir[_MAX_DIR];
	if (m_lpSelectedScript)
	{
		_wsplitpath_s(m_lpSelectedScript->sTruePathToScript.c_str(), drive, _MAX_DRIVE, dir, _MAX_DIR, NULL, 0, NULL, 0);
		PROCESS_INFORMATION processInformation = { 0 };
		STARTUPINFO startupInfo = { 0 };
		startupInfo.cb = sizeof(startupInfo);
		bool bReturn = true;

		// Create the process
		CString csPath = _T("explorer.exe \"");
		csPath += drive;
		csPath += dir;
		csPath += _T("\"");
		BOOL result = CreateProcess(NULL, csPath.GetBuffer(),
			NULL, NULL, FALSE,
			NORMAL_PRIORITY_CLASS | CREATE_NO_WINDOW,
			NULL, NULL, &startupInfo, &processInformation);
		
	}
}

void CPythagorasApp::CheckForAdmin()
{
#define INFO_BUFFER_SIZE 32767
	TCHAR  infoBuf[INFO_BUFFER_SIZE];
	DWORD  bufCharCount = INFO_BUFFER_SIZE;
#define CHECK_NAME(str) m_bAdminAccount = m_bAdminAccount || !_tcsicmp(infoBuf, str)

	m_bAdminAccount = false;
	// Get user name. 
	if (GetUserName(infoBuf, &bufCharCount))
	{
		CHECK_NAME(_T("A4LG8ZZ"));
	}
}

void CPythagorasApp::OnPopupExplorerNewScript()
{
	CNewScriptDlg dlg;
	CMemBuffer mbTemplateScript;
	bool bValid = false;
	CMainFrame* pMainFrame = (CMainFrame*)m_pMainWnd;
	TCHAR drive[_MAX_DRIVE];
	TCHAR dir[_MAX_DIR];
	if (!pMainFrame)
	{
		return;
	}

	mbTemplateScript.InitFromResource(IDR_TEMPLATE_SCRIPT);
	while (!bValid)
	{
		if (dlg.DoModal() != IDOK)
		{
			return;
		}
		
		// Build the path to the new file
		std::wstringstream ssNewScript;
		LPScriptViewNodeStruct lpSelNode = pMainFrame->GetScriptViewNode();
		if (lpSelNode)
		{
			if (lpSelNode->lpScriptFunction)
			{
				// We already know the path since they've selected an existing script				
				_wsplitpath_s(lpSelNode->lpScriptFunction->sTruePathToScript.c_str(), drive, _MAX_DRIVE, dir, _MAX_DIR, NULL, 0, NULL, 0);
				ssNewScript.str(_T(""));
				ssNewScript << drive << dir << dlg.m_sName.GetBuffer() << _T(".py");				
			}
			else
			{
				if (lpSelNode->uiLocalScriptIndex != -1)
				{					
					if (lpSelNode->lpChildScriptFunction)
					{						
						_wsplitpath_s(lpSelNode->lpChildScriptFunction->sTruePathToScript.c_str(), drive, _MAX_DRIVE, dir, _MAX_DIR, NULL, 0, NULL, 0);
						ssNewScript.str(_T(""));
						ssNewScript << drive << dir << dlg.m_sName.GetBuffer() << _T(".py");						
					}
					else
					{
						// In theory this should never occur
						return;
					}
				}
				else
				{
					if (lpSelNode->hParentNode == NULL)
					{
						// We're the root, just get the path to the local script folder
						ssNewScript.str(_T(""));
						ssNewScript << m_sTrueLocalScriptFolder.c_str() << _T("\\") << dlg.m_sName.GetBuffer() << _T(".py");						
					}

				}
			}

			// Check to see if that file already exists			
			std::filesystem::path source = ssNewScript.str().c_str();
			std::error_code ec;
			if (std::filesystem::exists(source, ec))
			{
				AfxMessageBox(_T("A script with that name already exists."));
			}
			else
			{
				mbTemplateScript.WriteToFile(ssNewScript.str().c_str());
				RefreshScripts();
				bValid = true;
			}			
		}		
	}
}


void CPythagorasApp::OnUpdatePopupExplorerNewScript(CCmdUI *pCmdUI)
{
	CMainFrame *pMainFrame = (CMainFrame *)m_pMainWnd;
	if (pMainFrame)
	{
		LPScriptViewNodeStruct lpSelNode = pMainFrame->GetScriptViewNode();

		if (lpSelNode)
		{
			if (lpSelNode->lpScriptFunction || lpSelNode->uiLocalScriptIndex != -1)
			{
				pCmdUI->Enable();
				return;
			}
		}
	}	
	pCmdUI->Enable(false);
}

/*static */void CPythagorasApp::DebugOut(const wchar_t *szText, bool bNewLine)
{
	FILE *wp = NULL;
#ifdef _DEBUG	
	TCHAR lpTempPathBuffer[MAX_PATH];
	std::wstringstream wsFormat;

	DWORD dwRetVal = GetTempPath(MAX_PATH,          // length of the buffer
		lpTempPathBuffer); // buffer for path 
	if (dwRetVal <= MAX_PATH && (dwRetVal != 0))
	{
		wsFormat << lpTempPathBuffer << _T("Pythagoras.txt");
		errno_t err = _wfopen_s(&wp, wsFormat.str().c_str(), _T("a"));
		if (err == 0)
		{
			fwprintf(wp, _T("%s%s"), szText, bNewLine ? _T("\n") : _T(""));
			fclose(wp);
		}
	}
#endif
}

/* static */ void CPythagorasApp::GetAppDirectory(std::string &sDirectory)
{
	std::wstring s;
	GetAppDirectory(s);
	sDirectory.append(s.begin(), s.end());
}

/* static */ void CPythagorasApp::GetAppDirectory(std::wstring &wDirectory)
{
	TCHAR szPath[MAX_PATH];
	TCHAR Drive[MAX_PATH], Dir[MAX_PATH];
	GetModuleFileName(AfxGetApp()->m_hInstance, szPath, MAX_PATH);

	_wsplitpath_s(szPath, Drive, MAX_PATH, Dir, MAX_PATH, NULL, 0, NULL, 0);

	wDirectory.append(Drive);
	wDirectory.append(Dir);
}

int CPythagorasApp::Run()
{
	int iReturn = 0;

	try
	{
		iReturn = CWinAppEx::Run();
	}
	catch (std::exception &ex)
	{
		std::wstringstream ssMessage;
		std::wstring ws;
		ws = L"Caught exception: ";
		ssMessage << ws << ex.what();
		::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);		
	}
	catch (...)
	{
		std::wstringstream ssMessage;
		std::wstring ws;
		ws = L"Caught unhandled exception:\n";

		LPVOID lpMessageBuffer;
		::FormatMessage
		(
			FORMAT_MESSAGE_ALLOCATE_BUFFER |
			FORMAT_MESSAGE_FROM_SYSTEM,		// source and processing options 
			NULL,							// address of  message source 
			::GetLastError(),						// requested message identifier 
			MAKELANGID
			(
				LANG_NEUTRAL,
				SUBLANG_DEFAULT
			),								// language identifier for requested message 
			(LPTSTR)&lpMessageBuffer,		// address of message buffer 
			0,								// maximum size of message buffer 
			NULL							// address of array of message inserts 
		);
		ssMessage << ws << (LPTSTR)&lpMessageBuffer;
		::LocalFree(lpMessageBuffer);

		::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);		
	}
	return iReturn;
}


void CPythagorasApp::OnHelpTroubleshootissues()
{
	// TODO: Add your command handler code here
	TroubleshootIssues();
}

// Troubleshooting log
/* static */void CPythagorasApp::GetPathTSLog(std::wstring &wsTSLog)
{
	std::wstringstream wsTempFile;
	std::wstring wsAppDirectory;
	TCHAR szPath[MAX_PATH];
	TCHAR Drive[MAX_PATH], Dir[MAX_PATH];

	GetModuleFileName(AfxGetApp()->m_hInstance, szPath, MAX_PATH);
	_wsplitpath_s(szPath, Drive, MAX_PATH, Dir, MAX_PATH, NULL, 0, NULL, 0);

	wsAppDirectory.append(Drive);
	wsAppDirectory.append(Dir);
	wsTempFile << wsAppDirectory.c_str() << _T("Troubleshooting.txt");
	wsTSLog = wsTempFile.str().c_str();
}

/* static */void CPythagorasApp::AddToTSLog(LPCTSTR szText, bool bNewLine /* = false */, bool bReset /* = false */)
{
	FILE *wp = NULL;
	errno_t err = 0;	
	std::wstring wsTSLog;	
	
	CString sTime = CTime::GetCurrentTime().Format("%m-%d-%Y %H:%M:%S");	
	GetPathTSLog(wsTSLog);
	
	err = _wfopen_s(&wp, wsTSLog.c_str(), bReset ? _T("w") : _T("a"));
	if (err == 0)
	{
		fwprintf(wp, _T("[%ws]%ws%ws"), sTime, szText, bNewLine ? _T("\n") : _T(""));
		fclose(wp);
	}	
}

void CPythagorasApp::AddToTSLog(std::wstringstream &ssText, bool bNewLine /* = false */, bool bReset /* = false */)
{
	AddToTSLog(ssText.str().c_str(), bNewLine, bReset);
}

/*static */void CPythagorasApp::TroubleshootIssues()
{
#define LARGE_PATH 32768
#define PYTHON_MODULES _T("pyparsing xlsxwriter cycler pytz setuptools matplotlib numpy scipy")
	FILE* fp = NULL;
	errno_t err = 0;
	DWORD dwRet = 0;
	wchar_t szPath[LARGE_PATH];
	std::wstringstream ssCmd;
	std::wstringstream ssFile;
	std::wstringstream ssTempPath;
	std::wstringstream ssScript;
	std::wstring sVersion;
	std::wstring sModuleName;
	std::string sCmd;
	std::wstring sVCRuntimeVersion;
	int iVisualCInstalled = 0;
	wchar_t version[1024];
	wchar_t module_name[1024];
	TCHAR lpTempPathBuffer[LARGE_PATH];
	_StringVector sv;
	_itString it;
	CRegistryHelper rhHelper;
	_SYSTEM_INFO SystemInfo;
	SHELLEXECUTEINFO shInfo;
	bool b32Bit = false;

	try
	{
		// Get the TEMP directory
		dwRet = GetTempPath(LARGE_PATH, lpTempPathBuffer);
		if (dwRet <= LARGE_PATH && (dwRet != 0))
		{
			ssTempPath << lpTempPathBuffer;
		}

		AddToTSLog(_T("Environment variables..."), true, true);
		ssCmd.str(_T(""));
		dwRet = ::GetEnvironmentVariable(_T("PATH"), szPath, LARGE_PATH);
		ssCmd << _T("PATH=") << szPath;
		AddToTSLog(ssCmd, true);

		rhHelper.SetMainKey(HKEY_LOCAL_MACHINE);
		rhHelper.SetBaseSubKey(_T("Software\\Microsoft\\VisualStudio\\14.0\\VC\\Runtimes\\x64"));
		rhHelper.AddItem(&iVisualCInstalled, -1, 0, 1, _T("Installed"), _T(""), 0);
		rhHelper.AddItem(&sVCRuntimeVersion, _T(""), _T("Version"), _T(""));
		rhHelper.ReadRegistry();

		AddToTSLog(_T("Looking for system components..."), true);
		if (iVisualCInstalled != -1)
		{
			ssCmd.str(_T(""));
			ssCmd << _T("Found Visual C++ 14.0 Runtime version ") << sVCRuntimeVersion.c_str();
			AddToTSLog(ssCmd, true);
		}
		else
		{
			AddToTSLog(_T("Did not find Visual C++ 14.0 Runtime"), true);
		}

		GetNativeSystemInfo(&SystemInfo);
		switch (SystemInfo.wProcessorArchitecture)
		{
		case PROCESSOR_ARCHITECTURE_AMD64:
		{
			b32Bit = false;
			AddToTSLog(_T("Detected 64-bit OS"), true);
		}
		break;

		case PROCESSOR_ARCHITECTURE_INTEL:
		{
			b32Bit = true;
			AddToTSLog(_T("Detected 32-bit OS"), true);
		}
		break;
		}


		// Get Python version
		AddToTSLog(_T("Checking Python installation..."), true);
		Py_Initialize();
		std::string sVersionString = Py_GetVersion();
		std::wstring swVersionString = std::wstring(sVersionString.begin(), sVersionString.end());
		Py_Finalize();

		ssCmd.str(_T(""));
		ssCmd << _T("Using Python version  ") << swVersionString.c_str();
		AddToTSLog(ssCmd, true);

		ssFile.str(_T(""));
		ssFile << ssTempPath.str().c_str() << _T("PythagorasModuleDependencies.out");
		ssScript.str(_T(""));
		ssScript << ssTempPath.str().c_str() << _T("PythagorasModuleDependencies.py");

		err = _wfopen_s(&fp, ssScript.str().c_str(), _T("w"));
		if (err == 0)
		{
			// There should be a way to do this directly with Py_ API functions, but I haven't had the time to look at it in depth...
			fprintf(fp, "%s%s%s%s%s%s",
				"from pkgutil import iter_modules\n",
				"import pkg_resources\n",
				"import sys\n",
				"for module_name in sys.argv[1:]:\n",
				"\tif module_name in (name for loader,name,ispkg in iter_modules()):\n",
				"\t\tprint(module_name + ' version ' + pkg_resources.get_distribution(module_name).version)");
			fclose(fp);
		}

		ssCmd.str(_T(""));
		ssCmd << (_T("python ")) << ssScript.str().c_str() << _T(" ");
		ssCmd << PYTHON_MODULES;
		ssCmd << _T("> ") << ssFile.str().c_str();
		sCmd = CW2A(ssCmd.str().c_str());

		ssCmd.str(_T(""));
		ssCmd << _T("Looking for existing Python modules ") << PYTHON_MODULES;
		AddToTSLog(ssCmd, true);
		system(sCmd.c_str());
		err = _wfopen_s(&fp, ssFile.str().c_str(), _T("r"));
		if (err == 0)
		{
			while (fgetws(module_name, 1024, fp))
			{
				sModuleName = module_name;
				CMemBuffer::TrimRight(sModuleName, '\n');
				ssCmd.str(_T(""));
				ssCmd << _T("Found Python module ") << sModuleName.c_str();
				AddToTSLog(ssCmd, true);
			}
			fclose(fp);
		}
	}
	catch (...)
	{
	}
	OpenTSLog();
}

/*static */ void CPythagorasApp::OpenTSLog()
{
	PROCESS_INFORMATION processInformation = { 0 };
	STARTUPINFO startupInfo = { 0 };
	startupInfo.cb = sizeof(startupInfo);
	bool bReturn = true;
	std::wstring wsTSLog;
	GetPathTSLog(wsTSLog);

	// Create the process
	CString csPath = theApp.m_RegistrySettings.m_sScriptEditor.c_str();
	csPath += _T(" ");
	csPath += wsTSLog.c_str();
	CreateProcess(NULL, csPath.GetBuffer(),
		NULL, NULL, FALSE,
		NORMAL_PRIORITY_CLASS | CREATE_NO_WINDOW,
		NULL, NULL, &startupInfo, &processInformation);
}

void CAboutDlg::OnBnClickedButtonSourceCode()
{
	CMemBuffer mbSourceCode;	
	std::wstring sFilter(_T("Zip files (*.zip)|*.zip|All files (*.*)|*.*||"));
	CFileDialog FileDlg(FALSE, _T(".zip"), _T("PythagorasSourceCode.zip"), OFN_OVERWRITEPROMPT, sFilter.c_str(), NULL, NULL);
	CString str;
	CString sTitle(_T("Please select an output file"));
	FileDlg.GetOFN().lpstrTitle = sTitle;
	
	if (FileDlg.DoModal() == IDOK)
	{
		CString FileName = FileDlg.GetPathName();
		DWORD dwAttrib = GetFileAttributes(FileName);

		if (dwAttrib != INVALID_FILE_ATTRIBUTES && !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY))
		{
			// File exists already... check to see if we can get access to the file; if not, it's probably open
			HANDLE fh;
			fh = CreateFile(FileName, GENERIC_READ, 0 /* no sharing! exclusive */, NULL, OPEN_EXISTING, 0, NULL);
			if ((fh != NULL) && (fh != INVALID_HANDLE_VALUE))
			{
				// We can access the file, good to go...
				CloseHandle(fh);
			}
			else
			{
				AfxMessageBox(_T("That file appears to be open, please close it first."));
				return;
			}
		}

		mbSourceCode.InitFromResource(IDR_SOURCE_CODE);
		mbSourceCode.WriteToFile(FileName);
	}
}

void CPythagorasApp::CheckPythonVersion()
{
	// The purpose of this function is two-fold: first, it's to retrieve the user's Python version; second, it's
	// to invoke the Python library code so that the compiler links to PythonXX.dll.  Once linked, the application
	// will not start unless PythonXX.dll is in the path.  This will provide a more direct error message to the user
	// at the very beginning, rather than waiting until PythonEngine is launched.  If PythonEngine fails it's not
	// always easy to determine why it failed since the error message is hidden from the user.
	Py_Initialize();

	// Py_GetVersion() will return something like the following:
	// 3.7.6 (tags/v3.7.6:43364a7ae0, Dec 19 2019, 00:44:22) [MSC v.1916 64 bit (AMD64)] 
	std::string sVersionString = Py_GetVersion();	
	std::wstring swVersionString = std::wstring(sVersionString.begin(), sVersionString.end());
	std::wstring stVersion, stMajor, stMinor, stMicro;

	_StringVector sv;
	boost::split(sv, swVersionString, boost::is_any_of(" "));
	if(sv.size() > 1)	
	{
		stVersion = sv.at(0).c_str();
		sv.clear();
		boost::split(sv, stVersion, boost::is_any_of("."));
		if (sv.size() == 3)
		{
			stMajor = sv.at(0).c_str();
			stMinor = sv.at(1).c_str();
			stMicro = sv.at(2).c_str();
		}
	}

	std::wstringstream ssOutput;
	ssOutput << _T("Using Python ") << swVersionString.c_str();
	GetMainFrame()->m_wndOutput.AddToOutput(eDescription, ssOutput.str().c_str());

	Py_Finalize();
}