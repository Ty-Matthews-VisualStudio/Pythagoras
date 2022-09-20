
// PythagorasEngine.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CPythagorasEngineApp:
// See PythagorasEngine.cpp for the implementation of this class
//

class CPythagorasEngineApp : public CWinApp
{
public:
	CPythagorasEngineApp();

// Overrides
public:
	virtual BOOL InitInstance();
	static void DebugOut(const wchar_t* szText, bool bNewLine);

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern CPythagorasEngineApp theApp;