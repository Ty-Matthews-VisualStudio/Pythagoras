
// PythagorasInstaller.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "PythagorasInstaller.h"
#include "PythagorasInstallerDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CPythagorasInstallerApp

BEGIN_MESSAGE_MAP(CPythagorasInstallerApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CPythagorasInstallerApp construction

CPythagorasInstallerApp::CPythagorasInstallerApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only CPythagorasInstallerApp object

CPythagorasInstallerApp theApp;


// CPythagorasInstallerApp initialization

BOOL CPythagorasInstallerApp::InitInstance()
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

	CoInitialize(NULL);

	CWinApp::InitInstance();


	AfxEnableControlContainer();

	// Create the shell manager, in case the dialog contains
	// any shell tree view or shell list view controls.
	CShellManager *pShellManager = new CShellManager;

	// Activate "Windows Native" visual manager for enabling themes in MFC controls
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	CPythagorasInstallerDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with OK
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}
	else if (nResponse == -1)
	{
		TRACE(traceAppMsg, 0, "Warning: dialog creation failed, so application is terminating unexpectedly.\n");
		TRACE(traceAppMsg, 0, "Warning: if you are using MFC controls on the dialog, you cannot #define _AFX_NO_MFC_CONTROLS_IN_DIALOGS.\n");
	}

	// Delete the shell manager created above.
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

#ifndef _AFXDLL
	ControlBarCleanUp();
#endif

	CoUninitialize();

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}

