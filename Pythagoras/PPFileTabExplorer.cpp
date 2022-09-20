// PPFileTabExplorer.cpp : implementation file
//

#include "stdafx.h"
#include "Pythagoras.h"
#include "PPFileTabExplorer.h"
#include "afxdialogex.h"


// PPFileTabExplorer dialog

IMPLEMENT_DYNAMIC(PPFileTabExplorer, CPropertyPage)

PPFileTabExplorer::PPFileTabExplorer()
	: CPropertyPage(IDD_FILE_TAB_EXPLORER)
{

}

PPFileTabExplorer::~PPFileTabExplorer()
{
}

void PPFileTabExplorer::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	//DDX_Control(pDX, IDC_MFCSHELLLIST_EXPLORER, m_wndExplorerTree);
	//DDX_Control(pDX, IDC_MFCSHELLLIST_EXPLORER, m_wndExplorerList);	
}


BEGIN_MESSAGE_MAP(PPFileTabExplorer, CPropertyPage)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_ACTIVATE()
END_MESSAGE_MAP()


// PPFileTabExplorer message handlers


int PPFileTabExplorer::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CPropertyPage::OnCreate(lpCreateStruct) == -1)
		return -1;	

	return 0;
}


void PPFileTabExplorer::OnSize(UINT nType, int cx, int cy)
{
	CPropertyPage::OnSize(nType, cx, cy);
	
}


void PPFileTabExplorer::OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized)
{	
	CPropertyPage::OnActivate(nState, pWndOther, bMinimized);

	int i = 0;
	// TODO: Add your message handler code here
}

void PPFileTabExplorer::Activate()
{
	static bool bOnce = false;

	if( !bOnce)
	{
		CWnd* pPlaceHolder = GetDlgItem(IDC_STATIC_EXPLORER_TREE_PLACEHOLDER);
		CRect rcPlaceHolder;
		pPlaceHolder->GetWindowRect(rcPlaceHolder);
		ScreenToClient(&rcPlaceHolder);
		
		m_wndExplorerTree.Create(/*WS_CHILD | */WS_VISIBLE | LVS_REPORT | TVS_HASLINES | TVS_LINESATROOT | TVS_HASBUTTONS,
			rcPlaceHolder,
			this,
			IDC_EXPLORER_TREE_CONTROL);

		m_wndExplorerTree.SetWindowPos(&CWnd::wndTop, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
		rcPlaceHolder.bottom -= 5;
		m_wndExplorerTree.MoveWindow(&rcPlaceHolder);
		
		pPlaceHolder = GetDlgItem(IDC_STATIC_EXPLORER_LIST_PLACEHOLDER);
		pPlaceHolder->GetWindowRect(rcPlaceHolder);
		ScreenToClient(&rcPlaceHolder);

		m_wndExplorerList.Create(/*WS_CHILD | */WS_VISIBLE | LVS_REPORT,
			rcPlaceHolder,
			this,
			IDC_EXPLORER_LIST_CONTROL);

		m_wndExplorerList.SetWindowPos(&CWnd::wndTop, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
		rcPlaceHolder.bottom -= 5;
		m_wndExplorerList.MoveWindow(&rcPlaceHolder);
		
		m_wndExplorerTree.SetRelatedList(&m_wndExplorerList);		
		m_wndExplorerList.SetRelatedTree(&m_wndExplorerTree);
		bOnce = true;

		m_wndExplorerTree.SelectPath(theApp.m_RegistrySettings.m_sExplorerPath.c_str());
		m_wndExplorerList.DisplayFolder(theApp.m_RegistrySettings.m_sExplorerPath.c_str());
	}	
}

BOOL PPFileTabExplorer::OnInitDialog()
{
	CPropertyPage::OnInitDialog();
	GetParentFrame()->RecalcLayout();

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

void PPFileTabExplorer::AddSelectedFiles()
{
	m_wndExplorerList.AddSelectedFiles();
}