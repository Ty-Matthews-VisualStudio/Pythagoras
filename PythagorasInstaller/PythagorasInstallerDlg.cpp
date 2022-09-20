
// PythagorasInstallerDlg.cpp : implementation file
//

#include "stdafx.h"
#include "PythagorasInstaller.h"
#include "PythagorasInstallerDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CPythagorasInstallerDlg dialog



CPythagorasInstallerDlg::CPythagorasInstallerDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_PYTHAGORASINSTALLER_DIALOG, pParent)
	, m_bPython(TRUE)
	, m_bVisualC(TRUE)
	, m_bPythagoras(TRUE)
	, m_bDesktopShortcut(TRUE)
	, m_bStartMenuShortcut(TRUE)
	, m_bPythonModules(FALSE)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CPythagorasInstallerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Check(pDX, IDC_CHECK_PYTHON, m_bPython);
	DDX_Check(pDX, IDC_CHECK_VISUALC, m_bVisualC);
	DDX_Check(pDX, IDC_CHECK_PYTHAGORAS, m_bPythagoras);
	DDX_Check(pDX, IDC_CHECK_DESKTOP_ICON, m_bDesktopShortcut);
	DDX_Check(pDX, IDC_CHECK_START_MENU_SHORTCUT, m_bStartMenuShortcut);
	DDX_Control(pDX, IDC_CHECK_DESKTOP_ICON, m_btnDesktopShortcut);
	DDX_Control(pDX, IDC_CHECK_START_MENU_SHORTCUT, m_btnStartMenuShortcut);
	DDX_Check(pDX, IDC_CHECK_PYTHON_MODULES, m_bPythonModules);
	DDX_Control(pDX, IDC_LIST_EXISTING, m_lstExistingComponents);
}

BEGIN_MESSAGE_MAP(CPythagorasInstallerDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_INSTALL, &CPythagorasInstallerDlg::OnBnClickedButtonInstall)
	ON_BN_CLICKED(IDC_CHECK_PYTHAGORAS, &CPythagorasInstallerDlg::OnBnClickedCheckPythagoras)
END_MESSAGE_MAP()


// CPythagorasInstallerDlg message handlers

BOOL CPythagorasInstallerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// Get temp directory
	TCHAR lpTempPathBuffer[MAX_PATH];
	DWORD dwRetVal = GetTempPath(MAX_PATH, lpTempPathBuffer);
	if (dwRetVal <= MAX_PATH && (dwRetVal != 0))
	{
		m_ssTempPath << lpTempPathBuffer;
		m_ssLogFile << lpTempPathBuffer << _T("Pythagoras_Installer_Log.txt");
	}
	else
	{
		AfxMessageBox(_T("Failed to get temp folder."));
		OnCancel();
	}

	m_bPython = m_RegistrySettings.m_iInstallPython == 0 ? false : true;
	m_bPythonModules = m_RegistrySettings.m_iInstallPythonModules == 1 ? true : false;
	m_bVisualC = m_RegistrySettings.m_iInstallVisualC == 0 ? false : true;
	m_bPythagoras = m_RegistrySettings.m_iInstallPythagoras == 0 ? false : true;
	m_bDesktopShortcut = m_RegistrySettings.m_iCreateDesktopShortcut == 0 ? false : true;
	m_bStartMenuShortcut = m_RegistrySettings.m_iCreateStartMenuShortcut == 0 ? false : true;

	DownloadModuleFile();
	DownloadPythonInstallFile();
	CheckForExistingComponents();

	UpdateData(FALSE);
	OnBnClickedCheckPythagoras();
	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CPythagorasInstallerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CPythagorasInstallerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CPythagorasInstallerDlg::OnBnClickedButtonInstall()
{
	
	std::wstring sPythagorasEXE;
	std::wstring sDummy;

	UpdateData(TRUE);
	if (!m_bPython && !m_bPythonModules && !m_bVisualC && !m_bPythagoras && !m_bDesktopShortcut && !m_bStartMenuShortcut)
	{
		AfxMessageBox(_T("Nothing to install!"));
		return;
	}

	m_RegistrySettings.m_iInstallPython = m_bPython ? true : false;
	m_RegistrySettings.m_iInstallPythonModules = m_bPythonModules ? true : false;
	m_RegistrySettings.m_iInstallVisualC = m_bVisualC ? true : false;
	m_RegistrySettings.m_iInstallPythagoras = m_bPythagoras ? true : false;
	m_RegistrySettings.m_iCreateDesktopShortcut = m_bDesktopShortcut ? true : false;
	m_RegistrySettings.m_iCreateStartMenuShortcut = m_bStartMenuShortcut ? true : false;

	std::wstring sFileName;
	std::wstringstream ssCommandLine;
	std::wstringstream ssFormat;

	std::wstring sPythonFile;
	std::wstring sVisualCFile;
	std::wstring sPythagorasFile;
	std::wstring sPythagorasEngineFile;
	
	switch (m_SystemInfo.wProcessorArchitecture)
	{	
	case PROCESSOR_ARCHITECTURE_AMD64:
	{
		//sPythonFile = _T("x64/python-3.6.0-amd64.exe");
		sPythonFile = m_sPythonInstallFilex64;
		sVisualCFile = _T("x64/mu_visual_cpp_redistributable_for_visual_studio_2015_x64_6837978.exe");		
		sPythagorasFile = _T("x64/Pythagoras.exe");
		sPythagorasEngineFile = _T("x64/PythagorasEngine.exe");
	}
	break;

	case PROCESSOR_ARCHITECTURE_INTEL:
	{
		//sPythonFile = _T("x86/python-3.6.0.exe");		
		sPythonFile = m_sPythonInstallFilex86;
		sVisualCFile = _T("x86/mu_visual_cpp_redistributable_for_visual_studio_2015_x86_6837977.exe");
		sPythagorasFile = _T("x86/Pythagoras.exe");
		sPythagorasEngineFile = _T("x86/PythagorasEngine.exe");
	}
	break;
	}	
	
	if (m_bPython)
	{
		AddToLog(_T("Downloading Python installer"));
		if (DownloadFile(sPythonFile, sFileName))
		{
			AddToLog(_T("Launching Python installer"));
			if (!ExecuteFile(sFileName.c_str()))
			{
				AddToLog(_T("ERROR Failed to launch Python installer"));
				AfxMessageBox(_T("Failed to install Python."));
				return;
			}
			AddToLog(_T("Success"));
		}
		else
		{
			AddToLog(_T("ERROR Failed to download Python installer"));
			AfxMessageBox(_T("Failed to download Python installer."));
			return;
		}

		// Try to refresh our variables after installing Python, to update the path.  Otherwise, we won't be able to 
		// install pip in the next section...		
		SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, (LPARAM) "Environment", SMTO_ABORTIFHUNG, 5000, NULL);
	}

	if (m_bPythonModules)
	{
#if 0
		AddToLog(_T("Upgrading pip"));
		ssCommandLine.str(_T(""));
		ssCommandLine << _T("python -m pip install --upgrade pip --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org");
		if (!ExecuteFile(ssCommandLine.str().c_str(), NORMAL_PRIORITY_CLASS | CREATE_NEW_CONSOLE))
		{
			AddToLog(_T("ERROR Failed to install/upgrade pip"));
			AfxMessageBox(_T("Failed to install/upgrade pip."));
			return;
		}
		AddToLog(_T("Success"));
#endif
		m_vPythonModules.push_back(PythonModule(_T("pyparsing"), _T("0"), _T("0")));
		m_vPythonModules.push_back(PythonModule(_T("xlsxwriter"), _T("0"), _T("0")));
		m_vPythonModules.push_back(PythonModule(_T("cycler"), _T("0"), _T("0")));
		m_vPythonModules.push_back(PythonModule(_T("pytz"), _T("0"), _T("0")));
		m_vPythonModules.push_back(PythonModule(_T("setuptools"), _T("0"), _T("0")));
		m_vPythonModules.push_back(PythonModule(_T("matplotlib"), _T("0"), _T("0")));
		m_vPythonModules.push_back(PythonModule(_T("numpy"), _T("0"), _T("0")));
		m_vPythonModules.push_back(PythonModule(_T("scipy"), _T("0"), _T("0")));
		
		std::wstringstream ssParams;
		std::wstring wParams;
		for (PythonModule& Module : m_vPythonModules)
		{
			bool bReturn = true;
			SHELLEXECUTEINFO shExInfo = { 0 };
			shExInfo.cbSize = sizeof(shExInfo);
			shExInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
			shExInfo.hwnd = this->m_hWnd;
			shExInfo.lpVerb = _T("runas");                // Operation to perform
			shExInfo.lpFile = _T("c:\\Users\\a4lg8zz\\Python\\Python.exe");       // Application to start    
			ssParams.str(_T(""));
			ssParams << _T("-m pip install ") << Module.m_sName << _T(" --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org");
			wParams = ssParams.str();
			shExInfo.lpParameters = wParams.c_str();                  // Additional parameters
			shExInfo.lpDirectory = 0;
			shExInfo.nShow = SW_SHOW;
			shExInfo.hInstApp = 0;

			if (ShellExecuteEx(&shExInfo))
			{
				WaitForSingleObject(shExInfo.hProcess, INFINITE);
				CloseHandle(shExInfo.hProcess);
			}
		}
#if 0
		for (PythonModule &Module : m_vPythonModules)
		{
			ssFormat.str(_T(""));
			ssFormat << _T("Installing ") << Module.m_sName.c_str() << _T(" module");
			AddToLog(ssFormat);

			if (Module.m_bFTPFile)
			{
				std::wstring sFile;
				switch (m_SystemInfo.wProcessorArchitecture)
				{
				case PROCESSOR_ARCHITECTURE_AMD64:
				{
					sFile = Module.m_sX64File;
				}
				break;

				case PROCESSOR_ARCHITECTURE_INTEL:
				{
					sFile = Module.m_sX86File;					
				}
				break;
				}

				if (DownloadFile(sFile, sFileName))
				{
					ssCommandLine.str(_T(""));
					ssCommandLine << _T("pip install ") << sFileName.c_str();
				}
				else
				{
					ssFormat.str(_T(""));
					ssFormat << _T("ERROR Failed to download ") << Module.m_sName.c_str() << _T(" wheel file");
					AddToLog(ssFormat);
					ssFormat.str(_T(""));
					ssFormat << _T("Failed to download ") << Module.m_sName.c_str() << _T(" wheel file");
					AfxMessageBox(ssFormat.str().c_str());
					return;
				}
			}
			else
			{
				ssCommandLine.str(_T(""));
				ssCommandLine << _T("pip install ") << Module.m_sName.c_str();
			}

			if (!ExecuteFile(ssCommandLine.str().c_str(), NORMAL_PRIORITY_CLASS | CREATE_NEW_CONSOLE))
			{
				ssFormat.str(_T(""));
				ssFormat << _T("ERROR Failed to install ") << Module.m_sName.c_str() << _T(" module");
				AddToLog(ssFormat.str().c_str());
				AfxMessageBox(ssFormat.str().c_str());
				return;
			}
			AddToLog(_T("Success"));			
		}
#endif  // #if 0
	}

	if (m_bVisualC)
	{
		AddToLog(_T("Downloading Visual C++ redistributable installer"));
		if (DownloadFile(sVisualCFile, sFileName))
		{
			AddToLog(_T("Launching Visual C++ redistributable installer"));
			if (!ExecuteFile(sFileName.c_str()))
			{
				AddToLog(_T("ERROR Failed to launch Visual C++ redistributable installer"));
				AfxMessageBox(_T("Failed to install Visual C++ Redistributable."));
				return;
			}
			AddToLog(_T("Success"));
		}
		else
		{
			AddToLog(_T("ERROR Failed to download Visual C++ redistributable installer"));
			AfxMessageBox(_T("Failed to download Visual C++ redistributable installer."));
			return;
		}
	}

	if (m_bPythagoras)
	{
		AddToLog(_T("Downloading Pythagoras executable"));
		if (DownloadFile(sPythagorasFile, sFileName))
		{
			AddToLog(_T("Creating ProgramFiles directory structure"));
			if (!CreateProgramFiles(sFileName.c_str(), sPythagorasEXE))
			{
				AddToLog(_T("ERROR Failed to create ProgramFiles directory structure, check to make sure running installer with elevated privileges"));
				AfxMessageBox(_T("Failed to install Pythagoras."));
				return;
			}
			AddToLog(_T("Success"));
		}
		else
		{
			AddToLog(_T("ERROR Failed to download Pythagoras executable"));
			AfxMessageBox(_T("Failed to download Pythagoras executable."));
			return;
		}
		AddToLog(_T("Downloading Pythagoras Engine executable"));
		if (DownloadFile(sPythagorasEngineFile, sFileName))
		{
			AddToLog(_T("Copying Pythagoras Engine executable to ProgramFiles directory structure"));
			if (!CreateProgramFiles(sFileName.c_str(), sDummy))
			{
				AfxMessageBox(_T("Failed to install PythagorasEngine."));
				return;
			}
			AddToLog(_T("Success"));
		}
		else
		{
			AddToLog(_T("ERROR Failed to download Pythagoras Engine executable"));
			AfxMessageBox(_T("Failed to download Pythagoras Engine executable."));
			return;
		}

		if (m_bStartMenuShortcut)
		{
			std::wstringstream ssFolder;
			CleanUp myCleanUp;
			AddToLog(_T("Obtaining path to FOLDERID_CommonPrograms"));
			if (!SUCCEEDED(SHGetKnownFolderPath(FOLDERID_CommonPrograms, 0, NULL, &myCleanUp.m_Folder)))
			{
				AddToLog(_T("During install, SHGetKnownFolderPath(FOLDERID_Programs) FAILED"));
				AfxMessageBox(_T("Failed to get start menu folder."));
				return;
			}
			DWORD dwErr = 0;
			ssFolder << myCleanUp.m_Folder << _T("\\3M\\");
			AddToLog(_T("Creating start menu path for 3M\\Pythagoras"));
			if (!CreateDirectory(ssFolder.str().c_str(), NULL))
			{
				dwErr = ::GetLastError();
				if (dwErr != ERROR_ALREADY_EXISTS)
				{
					std::wstringstream ssMsg;
					ssMsg << _T("During install CreateDirectory() FAILED, return code [") << dwErr << _T("]  Directory attempted: [") << ssFolder.str().c_str() << _T("]");
					AddToLog(ssMsg.str().c_str());
					AfxMessageBox(_T("Failed to create start menu folder."));
					return;
				}
			}
			AddToLog(_T("Creating Pythagoras.lnk under start menu"));
			ssFolder << _T("Pythagoras.lnk");
			std::wstring sArgs(_T(""));
			std::wstring sDescription(_T("Pythagoras application"));
			if (CreateShortcut(sPythagorasEXE.c_str(), sArgs.c_str(), ssFolder.str().c_str(), sDescription.c_str(), SW_SHOWNORMAL, sArgs.c_str(), sPythagorasEXE.c_str(), 0) < 0)
			{
				AddToLog(_T("Failed to create shortcut under start menu"));
				AfxMessageBox(_T("Failed to create start menu shortcut."));
				return;
			}
		}

		if (m_bDesktopShortcut)
		{
			std::wstringstream ssFolder;
			CleanUp myCleanUp;
			AddToLog(_T("Obtaining path to FOLDERID_Desktop"));
			if (!SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Desktop, 0, NULL, &myCleanUp.m_Folder)))
			{
				AddToLog(_T("During install, SHGetKnownFolderPath(FOLDERID_Desktop) FAILED"));
				AfxMessageBox(_T("Failed to get path to desktop."));
				return;
			}
			ssFolder << myCleanUp.m_Folder << _T("\\Pythagoras.lnk");
			AddToLog(_T("Creating Pythagoras.lnk on desktop"));
			std::wstring sArgs(_T(""));
			std::wstring sDescription(_T("Pythagoras application"));
			if (CreateShortcut(sPythagorasEXE.c_str(), sArgs.c_str(), ssFolder.str().c_str(), sDescription.c_str(), SW_SHOWNORMAL, sArgs.c_str(), sPythagorasEXE.c_str(), 0) < 0)
			{
				AddToLog(_T("Failed to create shortcut on desktop"));
				AfxMessageBox(_T("Failed to create desktop shortcut."));
				return;
			}
		}
	}

	std::wstringstream ssMessage;
	ssMessage << _T("Installation complete, no errors detected.  Check error log ") << m_ssLogFile.str().c_str() << _T(" for more details");
	OpenFile(m_ssLogFile.str().c_str());	
	AfxMessageBox(ssMessage.str().c_str());	
	OnOK();
}

bool CPythagorasInstallerDlg::CreateProgramFiles(LPCTSTR szSourceFile, std::wstring &sOutputPath)
{
	wchar_t* localProgramFiles = 0;
	std::wstring sArgv;
	std::wstringstream ssProgramFiles;
	std::wstringstream ssCWD;
	const SECURITY_ATTRIBUTES *psa = NULL;
	switch (m_SystemInfo.wProcessorArchitecture)
	{
	case PROCESSOR_ARCHITECTURE_AMD64:
	{
#if 0
		// This code does not work when the application is 32-bit...
		HRESULT hr = SHGetKnownFolderPath(FOLDERID_ProgramFilesX64, 0, NULL, &localProgramFiles);
		if (!SUCCEEDED(hr))
		{
			switch (hr)
			{
			case E_FAIL:
			{
				AddToLog(_T("In CreateProgramFiles(), SHGetKnownFolderPath(FOLDERID_ProgramFilesX64) exited with code E_FAIL"));
			}
			break;

			case E_INVALIDARG:
			{
				AddToLog(_T("In CreateProgramFiles(), SHGetKnownFolderPath(FOLDERID_ProgramFilesX64) exited with code E_INVALIDARG"));
			}
			break;
			}
			
			return false;
		}
#else
		CRegistryHelper RegHelper;
		RegHelper.SetMainKey(HKEY_LOCAL_MACHINE);
		RegHelper.SetBaseSubKey(_T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion"));
		std::wstring sProgramFiles;
		RegHelper.AddItem(&sProgramFiles, _T(""), _T("ProgramFilesDir"), _T(""));
		RegHelper.ReadRegistry(KEY_WOW64_64KEY);
		ssProgramFiles << sProgramFiles.c_str();
#endif
	}
	break;

	case PROCESSOR_ARCHITECTURE_INTEL:
	{
		if (!SUCCEEDED(SHGetKnownFolderPath(FOLDERID_ProgramFiles, 0, NULL, &localProgramFiles)))
		{
			AddToLog(_T("In CreateProgramFiles(), SHGetKnownFolderPath(FOLDERID_ProgramFiles) FAILED"));
			return false;
		}
		ssProgramFiles << localProgramFiles;
		CoTaskMemFree(static_cast<void*>(localProgramFiles));
	}
	break;
	}	

	ssProgramFiles << _T("\\3M\\");
	DWORD dwErr = 0;
	if (!CreateDirectory(ssProgramFiles.str().c_str(), NULL))
	{
		dwErr = ::GetLastError();
		if (dwErr != ERROR_ALREADY_EXISTS)
		{
			std::wstringstream ssMsg;
			ssMsg << _T("In CreateProgramFiles(), CreateDirectory() FAILED, return code [") << dwErr << _T("]  Directory attempted: [") << ssProgramFiles.str().c_str() << _T("]");
			AddToLog(ssMsg);
			return false;
		}
	}
	ssProgramFiles << _T("Pythagoras\\");
	if (!CreateDirectory(ssProgramFiles.str().c_str(), NULL))
	{
		dwErr = ::GetLastError();
		if (dwErr != ERROR_ALREADY_EXISTS)
		{
			std::wstringstream ssMsg;
			ssMsg << _T("In CreateProgramFiles(), CreateDirectory() FAILED, return code [") << dwErr << _T("]  Directory attempted: [") << ssProgramFiles.str().c_str() << _T("]");
			AddToLog(ssMsg);
			return false;
		}
	}

#if 0
	int iRet = _wmkdir(ssProgramFiles.str().c_str());
	if (iRet != 0 && iRet != EEXIST )
	{
		std::wstringstream ssMsg;
		ssMsg << _T("In CreateProgramFiles(), _wmkdir FAILED, return code [") << iRet << _T("]  Directory attempted: [") << ssProgramFiles.str().c_str() << _T("]");
		AddToLog(ssMsg.str().c_str());
		return false;
	}
#endif
		
	TCHAR fname[_MAX_FNAME];
	TCHAR ext[_MAX_EXT];
	_wsplitpath_s(szSourceFile, NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);

	std::wstringstream ssEXEFile;
	ssEXEFile << ssProgramFiles.str().c_str();
	ssEXEFile << fname << ext;

	sOutputPath = ssEXEFile.str().c_str();
	
	return CopyFile(szSourceFile, ssEXEFile.str().c_str(), FALSE);
}

bool CPythagorasInstallerDlg::ExecuteFile(LPCTSTR szFileName, DWORD dwCreationFlags /* NORMAL_PRIORITY_CLASS | CREATE_NO_WINDOW */)
{
	PROCESS_INFORMATION processInformation = { 0 };
	STARTUPINFO startupInfo = { 0 };
	startupInfo.cb = sizeof(startupInfo);
	bool bReturn = true;

	// Create the process
	CString csPath = szFileName;
	BOOL result = CreateProcess(NULL, csPath.GetBuffer(),
		NULL, NULL, FALSE,
		dwCreationFlags,
		NULL, NULL, &startupInfo, &processInformation);

	if (!result)
	{
		// CreateProcess() failed
		// Get the error from the system
		LPWSTR lpMsgBuf = NULL;
		DWORD dw = ::GetLastError();
		FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
			NULL, dw, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&lpMsgBuf, 0, NULL);

		// Log the error
		std::wstringstream ssError;
		ssError << _T("In ExecuteFile(), CreateProcess() failed\nCommand=") << szFileName << _T("\nMessage=") << (LPWSTR)lpMsgBuf;
		AddToLog(ssError);
		//::MessageBox(NULL, ssError.str().c_str(), NULL, MB_ICONSTOP);

		// Free resources created by the system
		LocalFree(lpMsgBuf);
		bReturn = false;
	}
	else
	{
		// Successfully created the process.  Wait for it to finish.
		WaitForSingleObject(processInformation.hProcess, INFINITE);
		DWORD exitCode;

		// Get the exit code.
		result = GetExitCodeProcess(processInformation.hProcess, &exitCode);

		// Close the handles.
		CloseHandle(processInformation.hProcess);
		CloseHandle(processInformation.hThread);

		bReturn = !result ? false : true;
	}
	return bReturn;
}

bool CPythagorasInstallerDlg::DownloadFile(std::wstring &sFileName, std::wstring &sLocalFile)
{
	return DownloadFile(sFileName.c_str(), sLocalFile);
}

bool CPythagorasInstallerDlg::DownloadFile(LPCTSTR szFileName, std::wstring &sLocalFile)
{
	CInternetSession InetSession(_T("PythagorasInstaller/1.0"));
	bool bReturn = false;

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
		FTPWrapper ftp(InetSession.GetFtpConnection(_T("ecpweb.mmm.com"), _T("Pythagoras"), _T("PythonFTP2017!")));
		
		// Build local filename
		TCHAR fname[_MAX_FNAME];
		TCHAR ext[_MAX_EXT];
		_wsplitpath_s(szFileName, NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);
		sLocalFile = m_ssTempPath.str().c_str();
		sLocalFile += fname;
		sLocalFile += ext;

		if (!ftp.m_pConnect->GetFile(szFileName, sLocalFile.c_str(),FALSE))
		{
			LPWSTR lpMessageBuffer;
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
				(LPWSTR)&lpMessageBuffer,		// address of message buffer 
				0,								// maximum size of message buffer 
				NULL							// address of array of message inserts 
			);
			std::wstringstream ssOutput;
			ssOutput << _T("Error downloading ") << szFileName << _T(" from FTP server.\nCode:") << dwErr << _T("\nMessage:") << (LPWSTR)&lpMessageBuffer;
			::LocalFree(lpMessageBuffer);
			//::MessageBox(NULL, ssOutput.str().c_str(), NULL, MB_ICONSTOP);
			AddToLog(ssOutput);
		}		
		else
		{
			bReturn = true;
		}
	}
	catch (std::exception &ex)
	{
		std::wstringstream ssMessage;
		ssMessage << _T("Caught exception in DownloadFile(): [Message] ") << ex.what();
		//::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);
		AddToLog(ssMessage);
	}
	catch (CInternetException* pEx)
	{
		std::wstringstream ssMessage;
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		ssMessage << _T("Caught InternetException in DownloadFile(): [Message] ") << sz << _T(" [Code] ") << pEx->m_dwError;
		//::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);
		AddToLog(ssMessage);
		pEx->Delete();
	}

	return bReturn;
}

void CPythagorasInstallerDlg::AddToLog(LPCTSTR szText)
{
	FILE *wp = NULL;
	errno_t err;	
	CString sTime = CTime::GetCurrentTime().Format("%m-%d-%Y %H:%M:%S");	

	err = _wfopen_s(&wp, m_ssLogFile.str().c_str(), _T("a"));
	if (err == 0)
	{		
		fprintf(wp, "[%ws]%ws\n", sTime, szText);
		fclose(wp);
	}
}

void CPythagorasInstallerDlg::AddToLog(std::wstringstream &ssText)
{
	AddToLog(ssText.str().c_str());
}

void CPythagorasInstallerDlg::OpenFile(LPCTSTR szFileName)
{
	SHELLEXECUTEINFO shInfo;
	memset(&shInfo, 0, sizeof(SHELLEXECUTEINFO));
	shInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	shInfo.fMask = SEE_MASK_NOASYNC;
	shInfo.hwnd = NULL;
	shInfo.lpVerb = _T("open");	
	shInfo.lpFile = szFileName;
	ShellExecuteEx(&shInfo);
}


HRESULT CreateShortcut(LPCWSTR pszTargetfile, LPCWSTR pszTargetargs,
	LPCWSTR pszLinkfile, LPCWSTR pszDescription,
	int iShowmode, LPCWSTR pszCurdir,
	LPCWSTR pszIconfile, int iIconindex)
{
	HRESULT       hRes;                  /* Returned COM result code */
	IShellLink*   pShellLink;            /* IShellLink object pointer */
	IPersistFile* pPersistFile;          /* IPersistFile object pointer */
	WCHAR wszLinkfile[MAX_PATH]; /* pszLinkfile as Unicode
								 string */
	int           iWideCharsWritten;     /* Number of wide characters
										 written */
	CoInitialize(NULL);
	hRes = E_INVALIDARG;
	if (
		(pszTargetfile != NULL) && (wcslen(pszTargetfile) > 0) &&
		(pszTargetargs != NULL) &&
		(pszLinkfile != NULL) && (wcslen(pszLinkfile) > 0) &&
		(pszDescription != NULL) &&
		(iShowmode >= 0) &&
		(pszCurdir != NULL) &&
		(pszIconfile != NULL) &&
		(iIconindex >= 0)
		)
	{
		hRes = CoCreateInstance(
			CLSID_ShellLink,     /* pre-defined CLSID of the IShellLink
								 object */
			NULL,                 /* pointer to parent interface if part of
								  aggregate */
			CLSCTX_INPROC_SERVER, /* caller and called code are in same
								  process */
			IID_IShellLink,      /* pre-defined interface of the
								 IShellLink object */
			(LPVOID*)&pShellLink);         /* Returns a pointer to the IShellLink
										   object */
		if (SUCCEEDED(hRes))
		{
			/* Set the fields in the IShellLink object */
			hRes = pShellLink->SetPath(pszTargetfile);
			hRes = pShellLink->SetArguments(pszTargetargs);
			if (wcslen(pszDescription) > 0)
			{
				hRes = pShellLink->SetDescription(pszDescription);
			}
			if (iShowmode > 0)
			{
				hRes = pShellLink->SetShowCmd(iShowmode);
			}
			if (wcslen(pszCurdir) > 0)
			{
				hRes = pShellLink->SetWorkingDirectory(pszCurdir);
			}
			if (wcslen(pszIconfile) > 0 && iIconindex >= 0)
			{
				hRes = pShellLink->SetIconLocation(pszIconfile, iIconindex);
			}

			/* Use the IPersistFile object to save the shell link */
			hRes = pShellLink->QueryInterface(
				IID_IPersistFile,         /* pre-defined interface of the
										  IPersistFile object */
				(LPVOID*)&pPersistFile);            /* returns a pointer to the
													IPersistFile object */
			if (SUCCEEDED(hRes))
			{
				/*
				iWideCharsWritten = MultiByteToWideChar(CP_ACP, 0,
					pszLinkfile, -1,
					wszLinkfile, MAX_PATH);
				*/
				hRes = pPersistFile->Save(pszLinkfile, TRUE);
				pPersistFile->Release();
			}
			pShellLink->Release();
		}

	}
	CoUninitialize();
	return (hRes);
}

void CPythagorasInstallerDlg::OnBnClickedCheckPythagoras()
{
	UpdateData(TRUE);
	m_btnDesktopShortcut.EnableWindow(m_bPythagoras);
	m_btnStartMenuShortcut.EnableWindow(m_bPythagoras);
}

void CPythagorasInstallerDlg::CheckForExistingComponents()
{
	std::wstringstream ssCmd;
	std::wstringstream ssFile;
	bool bHasPython = false;
	wchar_t version[1024];
	
	AddToLog(_T("Looking for existing components..."));

	ssFile << m_ssTempPath.str().c_str() << _T("PythagorasPythonVersion.txt");
	ssCmd << (_T("python --version > ")) << ssFile.str().c_str();
	std::string sCmd;
	sCmd = CW2A(ssCmd.str().c_str());
	system(sCmd.c_str());
	FILE *fp = NULL;
	errno_t err = _wfopen_s(&fp, ssFile.str().c_str(), _T("r"));
	if (err == 0)
	{		
		if (fgetws(version, 1024, fp))
		{
			m_lstExistingComponents.AddString(version);
			bHasPython = true;
		}
		fclose(fp);
	}
	
	if (!bHasPython)
	{
		AddToLog(_T("No python installation detected."));
	}
	else
	{
		std::wstring sVersion;
		sVersion = version;
		CMemBuffer::TrimRight(sVersion, '\n');
		ssCmd.str(_T(""));
		ssCmd << _T("Python version ") << sVersion.c_str() << _T(" found.");
		AddToLog(ssCmd);
		std::wstringstream ssScript;
		ssFile.str(_T(""));
		ssFile << m_ssTempPath.str().c_str() << _T("PythagorasModuleDependencies.out");
		ssScript << m_ssTempPath.str().c_str() << _T("PythagorasModuleDependencies.py");
		
		err = _wfopen_s(&fp, ssScript.str().c_str(), _T("w"));
		if (err != 0)
		{
			return;
		}
		fprintf(fp, "from pkgutil import iter_modules\nimport sys\nfor module_name in sys.argv[1:]:\n\tif module_name in (name for loader,name,ispkg in iter_modules()):\n\t\tprint(module_name)");
		fclose(fp);
		
		ssCmd.str(_T(""));
		ssCmd << (_T("python ")) << ssScript.str().c_str() << _T(" ");
		for (PythonModule &Module : m_vPythonModules)
		{
			ssCmd << Module.m_sName.c_str() << _T(" ");
		}		
		ssCmd << _T("> ") << ssFile.str().c_str();
		std::string sCmd;
		sCmd = CW2A(ssCmd.str().c_str());
		
		AddToLog(_T("Looking for existing Python modules..."));
		system(sCmd.c_str());
		err = _wfopen_s(&fp, ssFile.str().c_str(), _T("r"));
		if (err == 0)
		{
			wchar_t module_name[1024];
			while(fgetws(module_name, 1024, fp))
			{
				std::wstring sModuleName;
				sModuleName = module_name;
				CMemBuffer::TrimRight(sModuleName, '\n');
				ssCmd.str(_T(""));
				ssCmd << _T("Python module ") << sModuleName.c_str();
				m_lstExistingComponents.AddString(ssCmd.str().c_str());

				ssCmd.str(_T(""));
				ssCmd << _T("Found Python module ") << sModuleName.c_str();
				AddToLog(ssCmd);
			}			
			fclose(fp);		
		}			
	}	

	if (m_RegistrySettings.m_iVisualCInstalled != -1)
	{
		m_lstExistingComponents.AddString(_T("Visual C++ 14.0 Runtime"));
		AddToLog(_T("Found Visual C++ 14.0 Runtime"));
	}
		
	GetNativeSystemInfo(&m_SystemInfo);
	switch (m_SystemInfo.wProcessorArchitecture)
	{
	case PROCESSOR_ARCHITECTURE_AMD64:
	{
		m_lstExistingComponents.AddString(_T("64-bit OS"));
		AddToLog(_T("Detected 64-bit OS"));
	}
	break;

	case PROCESSOR_ARCHITECTURE_INTEL:
	{
		m_lstExistingComponents.AddString(_T("32-bit OS"));
		AddToLog(_T("Detected 32-bit OS"));
	}
	break;
	}

	
}

void CPythagorasInstallerDlg::DownloadModuleFile()
{
	// Get the modules file
	if (!DownloadFile(_T("PythonModules.txt"), m_sModuleFile))
	{
		AddToLog(_T("Failed to download PythonModules.txt"));
		return;
	}

	FILE *fp = NULL;
	errno_t err;
	
	err = _wfopen_s(&fp, m_sModuleFile.c_str(), _T("r"));
	if (err == 0)
	{
		wchar_t line[1024];
		while (_fgetts(line, 1024, fp))
		{
			std::wstring sModuleName = line;
			boost::trim_right_if(sModuleName, boost::is_any_of(" \n\r"));
			if (sModuleName.size() > 0)
			{
				std::vector<std::wstring> sv;
				boost::split(sv, sModuleName, boost::is_any_of(","));
				m_vPythonModules.push_back(PythonModule(sv.at(0).c_str(), sv.at(1).c_str(), sv.at(2).c_str()));
			}
		}
		fclose(fp);
	}
}

void CPythagorasInstallerDlg::DownloadPythonInstallFile()
{
	// Get the modules file
	std::wstring sFileName;
	if (!DownloadFile(_T("PythonInstall.txt"), sFileName))
	{
		AddToLog(_T("Failed to download PythonInstall.txt"));
		return;
	}

	FILE *fp = NULL;
	errno_t err;
	m_sPythonInstallFilex86 = _T("");
	m_sPythonInstallFilex64 = _T("");

	err = _wfopen_s(&fp, sFileName.c_str(), _T("r"));
	if (err == 0)
	{
		wchar_t line[1024];
		while (_fgetts(line, 1024, fp))
		{
			std::wstring sOSLine = line;
			boost::trim_right_if(sOSLine, boost::is_any_of(" \n\r"));
			if (sOSLine.size() > 0)
			{
				std::vector<std::wstring> sv;
				boost::split(sv, sOSLine, boost::is_any_of(","));
				if (sv.size() == 2)
				{
					if (!_tcsncicmp(sv.at(0).c_str(), _T("x86"), 3))
					{
						m_sPythonInstallFilex86 = sv.at(1);
					}
					if (!_tcsncicmp(sv.at(0).c_str(), _T("x64"), 3))
					{
						m_sPythonInstallFilex64 = sv.at(1);
					}
				}
			}
		}
		fclose(fp);
	}

	if (m_sPythonInstallFilex86.size() == 0 || m_sPythonInstallFilex64.size() == 0)
	{
		AddToLog(_T("Failed to parse PythonInstall.txt for x86 and x64 files"));
		return;
	}
}