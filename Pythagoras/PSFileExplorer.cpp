// PSFileExplorer.cpp : implementation file
//

#include "stdafx.h"
#include "Pythagoras.h"
#include "PSFileExplorer.h"


// PSFileExplorer

IMPLEMENT_DYNAMIC(PSFileExplorer, CPropertySheet)

PSFileExplorer::PSFileExplorer(UINT nIDCaption, CWnd* pParentWnd, UINT iSelectPage)
	:CPropertySheet(nIDCaption, pParentWnd, iSelectPage)
{
	AddPage(&m_ppFileTabExplorer);
	AddPage(&m_ppFileTabRegex);
}

PSFileExplorer::PSFileExplorer(LPCTSTR pszCaption, CWnd* pParentWnd, UINT iSelectPage)
	:CPropertySheet(pszCaption, pParentWnd, iSelectPage)
{
	AddPage(&m_ppFileTabExplorer);
	AddPage(&m_ppFileTabRegex);
	AddPage(&m_ppFileTabDatabase);
}

PSFileExplorer::~PSFileExplorer()
{
}


BEGIN_MESSAGE_MAP(PSFileExplorer, CPropertySheet)
	ON_WM_CREATE()
END_MESSAGE_MAP()


// PSFileExplorer message handlers


int PSFileExplorer::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CPropertySheet::OnCreate(lpCreateStruct) == -1)
		return -1;

	return 0;
}

void PSFileExplorer::AddSelectedFiles()
{
	int iSelPage = GetPageIndex(GetActivePage());
	switch (iSelPage)
	{
	default:
	case 0:
	{
		m_ppFileTabExplorer.AddSelectedFiles();
	}
	break;

	case 1:
	{
		m_ppFileTabRegex.OnBnClickedButtonRegexAddSelected();
	}
	break;

	case 2:
	{
		//m_ppFileTabDatabase
	}
	break;
	}
}