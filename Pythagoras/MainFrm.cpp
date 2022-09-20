
// MainFrm.cpp : implementation of the CMainFrame class
//

#include "stdafx.h"
#include "Pythagoras.h"
#include "PythagorasView.h"

#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CMainFrame

IMPLEMENT_DYNCREATE(CMainFrame, CFrameWndEx)

const int  iMaxUserToolbars = 10;
const UINT uiFirstUserToolBarId = AFX_IDW_CONTROLBAR_FIRST + 40;
const UINT uiLastUserToolBarId = uiFirstUserToolBarId + iMaxUserToolbars - 1;

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWndEx)
	ON_WM_CREATE()
	ON_COMMAND(ID_VIEW_CUSTOMIZE, &CMainFrame::OnViewCustomize)
	ON_REGISTERED_MESSAGE(AFX_WM_CREATETOOLBAR, &CMainFrame::OnToolbarCreateNew)
	ON_COMMAND_RANGE(ID_VIEW_APPLOOK_WIN_2000, ID_VIEW_APPLOOK_WINDOWS_7, &CMainFrame::OnApplicationLook)
	ON_UPDATE_COMMAND_UI_RANGE(ID_VIEW_APPLOOK_WIN_2000, ID_VIEW_APPLOOK_WINDOWS_7, &CMainFrame::OnUpdateApplicationLook)
	ON_WM_SETTINGCHANGE()
	ON_NOTIFY(TVN_SELCHANGED, IDC_TREE_SCRIPTS_VIEW, &CMainFrame::OnTvnSelchangedScriptsTree)
	ON_NOTIFY(NM_DBLCLK, IDC_TREE_SCRIPTS_VIEW, &CMainFrame::OnNMDblclkScriptsTree)
	ON_NOTIFY(TVN_SELCHANGED, IDC_TREE_FUNCTIONS_VIEW, &CMainFrame::OnTvnSelchangedFunctionsTree)	
	ON_REGISTERED_MESSAGE(CM_EXECUTETHREAD, &CMainFrame::OnExecuteThreadCustomMessage)
	ON_REGISTERED_MESSAGE(CM_ENGINE_CALLBACK, &CMainFrame::OnEngineCallbackCustomMessage)
	ON_COMMAND(ID_VIEW_TOOLBAR, &CMainFrame::OnViewToolbar)
	ON_COMMAND(ID_VIEW_SCRIPTS, &CMainFrame::OnViewScripts)
	ON_COMMAND(ID_VIEW_FUNCTIONS, &CMainFrame::OnViewFunctions)	
END_MESSAGE_MAP()

static UINT indicators[] =
{
	ID_SEPARATOR,           // status line indicator
	ID_INDICATOR_CAPS,
	ID_INDICATOR_NUM,
	ID_INDICATOR_SCRL,
};

// CMainFrame construction/destruction

CMainFrame::CMainFrame() : m_bDockingInitialized(false)
{
	// TODO: add member initialization code here
	theApp.m_nAppLook = theApp.GetInt(_T("ApplicationLook"), ID_VIEW_APPLOOK_WINDOWS_7);
}

CMainFrame::~CMainFrame()
{
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CFrameWndEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	if (!m_wndMenuBar.Create(this))
	{
		TRACE0("Failed to create menubar\n");
		return -1;      // fail to create
	}

	m_wndMenuBar.SetPaneStyle(m_wndMenuBar.GetPaneStyle() | CBRS_SIZE_DYNAMIC | CBRS_TOOLTIPS | CBRS_FLYBY);

	// prevent the menu bar from taking the focus on activation
	CMFCPopupMenu::SetForceMenuFocus(FALSE);

#if 0
	if (!m_wndToolBar.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC) ||
		!m_wndToolBar.LoadToolBar(theApp.m_bHiColorIcons ? IDR_MAINFRAME_256 : IDR_MAINFRAME))
	{
		TRACE0("Failed to create toolbar\n");
		return -1;      // fail to create
	}

	CString strToolBarName;
	bNameValid = strToolBarName.LoadString(IDS_TOOLBAR_STANDARD);
	ASSERT(bNameValid);
	m_wndToolBar.SetWindowText(strToolBarName);
	m_wndToolBar.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndToolBar);
	
	CString strCustomize;
	bNameValid = strCustomize.LoadString(IDS_TOOLBAR_CUSTOMIZE);
	ASSERT(bNameValid);
	m_wndToolBar.EnableCustomizeButton(TRUE, ID_VIEW_CUSTOMIZE, strCustomize);

	// Allow user-defined toolbars operations:
	InitUserToolbars(NULL, uiFirstUserToolBarId, uiLastUserToolBarId);
#endif

	if (!m_wndStatusBar.Create(this))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}
	m_wndStatusBar.SetIndicators(indicators, sizeof(indicators)/sizeof(UINT));

	// TODO: Delete these five lines if you don't want the toolbar and menubar to be dockable
	m_wndMenuBar.EnableDocking(CBRS_ALIGN_ANY);
	
	EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndMenuBar);
	

	// enable Visual Studio 2005 style docking window behavior
	CDockingManager::SetDockingMode(DT_SMART);
	// enable Visual Studio 2005 style docking window auto-hide behavior
	EnableAutoHidePanes(CBRS_ALIGN_ANY);

	// Load menu item image (not placed on any standard toolbars):
	CMFCToolBar::AddToolBarForImageCollection(IDR_MENU_IMAGES, theApp.m_bHiColorIcons ? IDB_MENU_IMAGES_24 : 0);

	// create docking windows
	if (!CreateDockingWindows())
	{
		TRACE0("Failed to create docking windows\n");
		return -1;
	}

	// set the visual manager and style based on persisted value
#if 1
	OnApplicationLook(theApp.m_nAppLook);
#endif

	m_wndScriptsView.SetMainFrame(this);

	m_wndScriptsView.EnableDocking(CBRS_ALIGN_ANY);
	m_wndFunctionsView.EnableDocking(CBRS_ALIGN_ANY);
	m_wndOutput.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndScriptsView);
	DockPane(&m_wndOutput);
	
	CDockablePane* pTabbedBar = NULL;
	if (m_wndFunctionsView.AttachToTabWnd(&m_wndScriptsView, DM_SHOW, TRUE, &pTabbedBar))
	{		
	}
#if 0
	m_wndProperties.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndProperties);
#endif

	
		
#if 0
	// Enable toolbar and docking window menu replacement
	EnablePaneMenu(TRUE, ID_VIEW_CUSTOMIZE, strCustomize, ID_VIEW_TOOLBAR);

	// enable quick (Alt+drag) toolbar customization
	CMFCToolBar::EnableQuickCustomization();

	if (CMFCToolBar::GetUserImages() == NULL)
	{
		// load user-defined toolbar images
		if (m_UserImages.Load(_T(".\\UserImages.bmp")))
		{
			CMFCToolBar::SetUserImages(&m_UserImages);
		}
	}

	// enable menu personalization (most-recently used commands)
	// TODO: define your own basic commands, ensuring that each pulldown menu has at least one basic command.
	CList<UINT, UINT> lstBasicCommands;

	lstBasicCommands.AddTail(ID_FILE_NEW);
	lstBasicCommands.AddTail(ID_FILE_OPEN);
	lstBasicCommands.AddTail(ID_FILE_SAVE);
	lstBasicCommands.AddTail(ID_FILE_PRINT);
	lstBasicCommands.AddTail(ID_APP_EXIT);
	lstBasicCommands.AddTail(ID_EDIT_CUT);
	lstBasicCommands.AddTail(ID_EDIT_PASTE);
	lstBasicCommands.AddTail(ID_EDIT_UNDO);
	lstBasicCommands.AddTail(ID_APP_ABOUT);
	lstBasicCommands.AddTail(ID_VIEW_STATUS_BAR);
	lstBasicCommands.AddTail(ID_VIEW_TOOLBAR);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2003);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_VS_2005);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2007_BLUE);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2007_SILVER);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2007_BLACK);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2007_AQUA);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_WINDOWS_7);
	lstBasicCommands.AddTail(ID_SORTING_SORTALPHABETIC);
	lstBasicCommands.AddTail(ID_SORTING_SORTBYTYPE);
	lstBasicCommands.AddTail(ID_SORTING_SORTBYACCESS);
	lstBasicCommands.AddTail(ID_SORTING_GROUPBYTYPE);

	CMFCToolBar::SetBasicCommands(lstBasicCommands);
#endif

	RefreshScripts();
	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	cs.style &= ~(LONG)FWS_ADDTOTITLE;
	if( !CFrameWndEx::PreCreateWindow(cs) )
		return FALSE;
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return TRUE;
}

BOOL CMainFrame::CreateDockingWindows()
{
	BOOL bNameValid;

	// Create class view
	CString strScriptsView;
	bNameValid = strScriptsView.LoadString(IDS_SCRIPTS_VIEW);
	ASSERT(bNameValid);
	if (!m_wndScriptsView.Create(strScriptsView, this, CRect(0, 0, 200, 200), TRUE, ID_VIEW_FILEVIEW, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_LEFT | CBRS_FLOAT_MULTI))
	{
		TRACE0("Failed to create Scripts View window\n");
		return FALSE; // failed to create
	}

	// Create file view
	CString strFunctionsView;
	bNameValid = strFunctionsView.LoadString(IDS_FUNCTIONS_VIEW);
	ASSERT(bNameValid);
	if (!m_wndFunctionsView.Create(strFunctionsView, this, CRect(0, 0, 200, 200), TRUE, ID_VIEW_CLASSVIEW, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_LEFT| CBRS_FLOAT_MULTI))
	{
		TRACE0("Failed to create Functions View window\n");
		return FALSE; // failed to create
	}

	// Create output window
	CString strOutputWnd;
	bNameValid = strOutputWnd.LoadString(IDS_OUTPUT_WND);
	ASSERT(bNameValid);
	if (!m_wndOutput.Create(strOutputWnd, this, CRect(0, 0, 100, 100), TRUE, ID_VIEW_OUTPUTWND, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_BOTTOM | CBRS_FLOAT_MULTI))
	{
		TRACE0("Failed to create Output window\n");
		return FALSE; // failed to create
	}

#if 0
	// Create properties window
	CString strPropertiesWnd;
	bNameValid = strPropertiesWnd.LoadString(IDS_PROPERTIES_WND);
	ASSERT(bNameValid);
	if (!m_wndProperties.Create(strPropertiesWnd, this, CRect(0, 0, 200, 200), TRUE, ID_VIEW_PROPERTIESWND, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_RIGHT | CBRS_FLOAT_MULTI))
	{
		TRACE0("Failed to create Properties window\n");
		return FALSE; // failed to create
	}
#endif

	SetDockingWindowIcons(theApp.m_bHiColorIcons);
	m_bDockingInitialized = true;
	return TRUE;
}

void CMainFrame::SetDockingWindowIcons(BOOL bHiColorIcons)
{
	HICON hFileViewIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_FILE_VIEW_HC : IDI_FILE_VIEW), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndScriptsView.SetIcon(hFileViewIcon, FALSE);

	HICON hClassViewIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_CLASS_VIEW_HC : IDI_CLASS_VIEW), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndFunctionsView.SetIcon(hClassViewIcon, FALSE);

	HICON hOutputBarIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_OUTPUT_WND_HC : IDI_OUTPUT_WND), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndOutput.SetIcon(hOutputBarIcon, FALSE);

#if 0
	HICON hPropertiesBarIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_PROPERTIES_WND_HC : IDI_PROPERTIES_WND), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndProperties.SetIcon(hPropertiesBarIcon, FALSE);
#endif

}

// CMainFrame diagnostics

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWndEx::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWndEx::Dump(dc);
}
#endif //_DEBUG


// CMainFrame message handlers

void CMainFrame::OnViewCustomize()
{
	CMFCToolBarsCustomizeDialog* pDlgCust = new CMFCToolBarsCustomizeDialog(this, TRUE /* scan menus */);
	pDlgCust->EnableUserDefinedToolbars();
	pDlgCust->Create();
}

LRESULT CMainFrame::OnToolbarCreateNew(WPARAM wp,LPARAM lp)
{
	LRESULT lres = CFrameWndEx::OnToolbarCreateNew(wp,lp);
	if (lres == 0)
	{
		return 0;
	}

	CMFCToolBar* pUserToolbar = (CMFCToolBar*)lres;
	ASSERT_VALID(pUserToolbar);

	BOOL bNameValid;
	CString strCustomize;
	bNameValid = strCustomize.LoadString(IDS_TOOLBAR_CUSTOMIZE);
	ASSERT(bNameValid);

	pUserToolbar->EnableCustomizeButton(TRUE, ID_VIEW_CUSTOMIZE, strCustomize);
	return lres;
}

void CMainFrame::OnApplicationLook(UINT id)
{
	CWaitCursor wait;

	theApp.m_nAppLook = id;

	switch (theApp.m_nAppLook)
	{
	case ID_VIEW_APPLOOK_WIN_2000:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManager));
		break;

	case ID_VIEW_APPLOOK_OFF_XP:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOfficeXP));
		break;

	case ID_VIEW_APPLOOK_WIN_XP:
		CMFCVisualManagerWindows::m_b3DTabsXPTheme = TRUE;
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));
		break;

	case ID_VIEW_APPLOOK_OFF_2003:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2003));
		CDockingManager::SetDockingMode(DT_SMART);
		break;

	case ID_VIEW_APPLOOK_VS_2005:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerVS2005));
		CDockingManager::SetDockingMode(DT_SMART);
		break;

	case ID_VIEW_APPLOOK_VS_2008:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerVS2008));
		CDockingManager::SetDockingMode(DT_SMART);
		break;

	case ID_VIEW_APPLOOK_WINDOWS_7:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows7));
		CDockingManager::SetDockingMode(DT_SMART);
		break;

	default:
		switch (theApp.m_nAppLook)
		{
		case ID_VIEW_APPLOOK_OFF_2007_BLUE:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_LunaBlue);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_BLACK:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_ObsidianBlack);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_SILVER:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_Silver);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_AQUA:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_Aqua);
			break;
		}

		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2007));
		CDockingManager::SetDockingMode(DT_SMART);
	}

	m_wndOutput.UpdateFonts();
	RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME | RDW_ERASE);

	theApp.WriteInt(_T("ApplicationLook"), theApp.m_nAppLook);
}

void CMainFrame::OnUpdateApplicationLook(CCmdUI* pCmdUI)
{
	pCmdUI->SetRadio(theApp.m_nAppLook == pCmdUI->m_nID);
}


BOOL CMainFrame::LoadFrame(UINT nIDResource, DWORD dwDefaultStyle, CWnd* pParentWnd, CCreateContext* pContext) 
{
	// base class does the real work

	if (!CFrameWndEx::LoadFrame(nIDResource, dwDefaultStyle, pParentWnd, pContext))
	{
		return FALSE;
	}


	// enable customization button for all user toolbars
	BOOL bNameValid;
	CString strCustomize;
	bNameValid = strCustomize.LoadString(IDS_TOOLBAR_CUSTOMIZE);
	ASSERT(bNameValid);

	for (int i = 0; i < iMaxUserToolbars; i ++)
	{
		CMFCToolBar* pUserToolbar = GetUserToolBarByIndex(i);
		if (pUserToolbar != NULL)
		{
			pUserToolbar->EnableCustomizeButton(TRUE, ID_VIEW_CUSTOMIZE, strCustomize);
		}
	}

	return TRUE;
}


void CMainFrame::OnSettingChange(UINT uFlags, LPCTSTR lpszSection)
{
	CFrameWndEx::OnSettingChange(uFlags, lpszSection);
	m_wndOutput.UpdateFonts();
}

void CMainFrame::RefreshScripts()
{
	m_wndScriptsView.BuildScriptTree();
	OnViewScripts();
}

void CMainFrame::OnNMDblclkScriptsTree(NMHDR *pNMHDR, LRESULT *pResult)
{
	m_wndFunctionsView.ShowPane(TRUE, FALSE, TRUE);
	*pResult = 0;
}

LPScriptViewNodeStruct CMainFrame::GetScriptViewNode()
{
	if (m_bDockingInitialized)
	{
		CTreeCtrl *pScriptTree = m_wndScriptsView.GetTreeCtrl();
		if (::IsWindow(m_wndScriptsView.m_hWnd))
		{
			HTREEITEM hSel = pScriptTree->GetSelectedItem();
			if (hSel)
			{
				DWORD_PTR dwp = pScriptTree->GetItemData(hSel);
				return reinterpret_cast<LPScriptViewNodeStruct>(dwp);
			}
		}
	}	
	
	return NULL;	
}

void CMainFrame::OnTvnSelchangedScriptsTree(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
	HTREEITEM hNew = pNMTreeView->itemNew.hItem;
	HTREEITEM hOld = pNMTreeView->itemOld.hItem;

	theApp.m_lpSelectedScript = NULL;

	if (hNew)
	{
		CTreeCtrl *pScriptTree = m_wndScriptsView.GetTreeCtrl();		
		DWORD_PTR dwp = pScriptTree->GetItemData(hNew);
		if (dwp != 0)
		{
			LPScriptViewNodeStruct lpNode = reinterpret_cast<LPScriptViewNodeStruct>(dwp);
			theApp.m_lpSelectedScript = lpNode->lpScriptFunction;
			m_wndFunctionsView.RefreshContent();

			if (lpNode->lpScriptFunction && lpNode->lpScriptFunction->bErrors)
			{
				m_wndOutput.SetOutput(eDescription, lpNode->lpScriptFunction->ssErrorMessage.str(), true);
			}
		}
	}

	// TODO: Add your control notification handler code here
	*pResult = 0;
}

void CMainFrame::OnTvnSelchangedFunctionsTree(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
	HTREEITEM hNew = pNMTreeView->itemNew.hItem;

	if (hNew)
	{
		CTreeCtrl *pFunctionsTree = m_wndFunctionsView.GetTreeCtrl();
		DWORD_PTR dwp = pFunctionsTree->GetItemData(hNew);
		if(theApp.m_lpSelectedScript && dwp < theApp.m_lpSelectedScript->vFunctions.size())
		{
			LPFunctionStruct lpFunction = theApp.m_lpSelectedScript->vFunctions.at(dwp);
			m_wndOutput.SetOutput(eDescription, lpFunction->sDescription, true);
			theApp.m_dwSelectedFunction = dwp;
		}
		else
		{
			m_wndOutput.SetOutput(eDescription, _T(""));			
			theApp.m_dwSelectedFunction = -1;
		}		
	}
}

LRESULT CMainFrame::OnEngineCallbackCustomMessage(WPARAM wp, LPARAM lp)
{
	using namespace PythonEngine;
	LRESULT lReturn = 0;
	CPythagorasView *pView = (CPythagorasView *)GetActiveView();

	try
	{
		wmanaged_shared_memory segment(open_only, cSharedMemoryString);
		_SharedScriptCallback *Callback = segment.find<_SharedScriptCallback>(cScriptCallbackString).first;
		void_allocator alloc_inst(segment.get_segment_manager());
		
		switch (wp)
		{
		case eMCMessageBox:
		{
			if (!Callback)
			{
				return 0;
			}
			std::wstring sTitle;
			ConvertString(Callback->m_String, sTitle);
			lReturn = AfxMessageBox(sTitle.c_str(), (unsigned int)lp);
		}
		break;

		case eMCStepIt:
		{
			pView->IncrementProgressBar((unsigned int)lp);
		}
		break;

		case eMCSetStatus:
		{
			if (!Callback)
			{
				return 0;
			}
			std::wstring sStatus;
			ConvertString(Callback->m_String, sStatus);
			SetStatusBarText(sStatus.c_str());
		}
		break;

		case eMCPrint:
		{
			if (!Callback)
			{
				return 0;
			}
			std::wstring sPrint;
			ConvertString(Callback->m_String, sPrint);
			m_wndOutput.SetForegroundWindow();
			m_wndOutput.AddToOutput(eStandardOutput, sPrint, true);
		}
		break;
		
		case eMCStepRange:
		{
			pView->m_pcScriptExecution.SetRange(0, (unsigned int)lp);
		}
		break;

		case eMCAddToRegex:
		{
			_SharedStringArray *FileList = segment.find<_SharedStringArray>(cStringArrayString).first;
			_SharedStringIterator itFileName;
			std::wstring sFileName;
			eModuleCallbackRegexAction eAction = static_cast<eModuleCallbackRegexAction>(lp);
			if (eAction == eMCRegexReplace)
			{
				pView->RegexResetContent();
			}
			pView->RegexClear();
			for (itFileName = FileList->m_Strings.begin(); itFileName != FileList->m_Strings.end(); itFileName++)
			{				
				ConvertString((*itFileName).m_String, sFileName);
				pView->RegexAddFile(sFileName.c_str(), NULL);
			}
			pView->RegexDisplayResultList();
		}
		break;
		}
	}
	catch (interprocess_exception &ex)
	{
		ex.what();
	}
	catch (...)
	{
	}
	
	return lReturn;
}

LRESULT CMainFrame::OnExecuteThreadCustomMessage(WPARAM wp, LPARAM lp)
{
	switch (static_cast<ExecuteThreadMessageWPARAMEnum>(wp))
	{
	case eETWClearOutput:
	{
		SetFocus();
		ShowPane(&m_wndOutput, TRUE, FALSE, TRUE);
		m_wndOutput.ClearOutput();
		RecalcLayout();
	}
	break;

	case eETWSelectPane:
	{
		SetFocus();
		ShowPane(&m_wndOutput, TRUE, FALSE, TRUE);
		m_wndOutput.m_wndTabs.SetActiveTab(static_cast<OutputTypeEnum>(lp));
		RecalcLayout();		
	}
	break;

	case eETWThreadFinished:
	{
		SetStatusBarText(_T("Script finished"));
		theApp.CheckScriptThreads(1000, true);
	}
	break;

	default:
	{
		ASSERT(FALSE);
	}
	break;
	}
	return 0;
}

void CMainFrame::OnViewToolbar()
{
	SetFocus();
	ShowPane(&m_wndOutput, TRUE, FALSE, TRUE);	
	RecalcLayout();	
}


void CMainFrame::OnViewScripts()
{
	SetFocus();
	m_wndScriptsView.ShowPane(TRUE, FALSE, TRUE);
	RecalcLayout();
}


void CMainFrame::OnViewFunctions()
{
	SetFocus();
	m_wndFunctionsView.ShowPane(TRUE, FALSE, TRUE);
	RecalcLayout();	
}

void CMainFrame::SetStatusBarText(LPCTSTR szText, int nPaneId /*= 0*/)
{
	m_wndStatusBar.SetPaneText(nPaneId, szText);
}