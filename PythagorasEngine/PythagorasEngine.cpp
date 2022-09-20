
// PythagorasEngine.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "PythagorasEngine.h"
#include "PythagorasEngineDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CPythagorasEngineApp

BEGIN_MESSAGE_MAP(CPythagorasEngineApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CPythagorasEngineApp construction

CPythagorasEngineApp::CPythagorasEngineApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

/*static */void CPythagorasEngineApp::DebugOut(const wchar_t* szText, bool bNewLine)
{
	FILE* wp = NULL;
#ifdef _DEBUG	
	TCHAR lpTempPathBuffer[MAX_PATH];
	std::wstringstream wsFormat;

	DWORD dwRetVal = GetTempPath(MAX_PATH,          // length of the buffer
		lpTempPathBuffer); // buffer for path 
	if (dwRetVal <= MAX_PATH && (dwRetVal != 0))
	{
		wsFormat << lpTempPathBuffer << _T("PythagorasEngine.txt");
		errno_t err = _wfopen_s(&wp, wsFormat.str().c_str(), _T("a"));
		if (err == 0)
		{
			fwprintf(wp, _T("%s%s"), szText, bNewLine ? _T("\n") : _T(""));
			fclose(wp);
		}
	}
#endif
}


// The one and only CPythagorasEngineApp object

CPythagorasEngineApp theApp;


// CPythagorasEngineApp initialization

BOOL CPythagorasEngineApp::InitInstance()
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

	DebugOut(_T("1"),true);

	CWinApp::InitInstance();
	AfxEnableControlContainer();

	// Create the shell manager, in case the dialog contains
	// any shell tree view or shell list view controls.
	CShellManager *pShellManager = new CShellManager;

	// Activate "Windows Native" visual manager for enabling themes in MFC controls
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	DebugOut(_T("2"), true);

	PythonEngine::ExecuteParams EP;
	PythonEngine::PythonSession PySession;
	PythonEngine::PythonSession::m_pExecuteParams = &EP;
	EP.FV.SetExecuteParams(reinterpret_cast<DWORD_PTR>(&EP));

	EP.alloc_inst = NULL;
	EP.pSharedStrings = NULL;

	DebugOut(_T("3"), true);
	
	std::string s;
	if (!PythonEngine::ExtractSharedMemory(&EP,s))
	{
		DebugOut(_T("Failed to extract shared memory"), true);
		boost::iterator_range<std::string::const_iterator> rng;
		if (boost::ifind_first(s, std::string("Unsupported version")))
		{
			DebugOut(_T("Verion wrong?"), true);
			::MessageBox(NULL, L"Please update client-side application.  This script will not function with current version.", NULL, MB_ICONSTOP);
		}
		else
		{
			CString cs(s.c_str());
			DebugOut(cs, true);			
			::MessageBox(NULL, cs, NULL, MB_ICONSTOP);
		}
	}
	else
	{
		DebugOut(_T("EP.ProcessMode == PythonEngine::ePMCallFunction"), true);
		if (EP.ProcessMode == PythonEngine::ePMCallFunction)
		{		
			// dlg.m_pExecuteParams = &EP;
			// INT_PTR nResponse = dlg.DoModal();
			INT_PTR nResponse = IDOK;
			if (nResponse == IDOK)
			{
				ULONGLONG ulStart, ulEnd;
#define TIMER_START ulStart = ::GetTickCount64();
#define TIMER_END ulEnd = ::GetTickCount64();
#define TIME_ELAPSED ((ulEnd - ulStart) / 1000.0)

				TIMER_START;
				bool bReturn = PythonEngine::ExecuteScript(&EP, PythonEngine::eEMExecute);
				TIMER_END;
				EP.dExecutionTime = TIME_ELAPSED;

				if (!bReturn)
				{
					CString cs(EP.ssMessage.str().c_str());
					TRACE(cs);
					::MessageBox(NULL, cs, NULL, MB_ICONSTOP);
				}
				else
				{
					if (!PythonEngine::SetSharedMemory(&EP, s))
					{
						CString cs(s.c_str());
						::MessageBox(NULL, cs, NULL, MB_ICONSTOP);
					}
#if 0
					if (EP.m_bOpenFiles)
					{
						EP.FV.OpenSaveFiles();
					}
#endif
				}
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
		}
	}	

	// Delete the shell manager created above.
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

#ifndef _AFXDLL
	ControlBarCleanUp();
#endif

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}

