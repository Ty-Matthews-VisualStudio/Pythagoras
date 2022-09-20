
// PythagorasView.cpp : implementation of the CPythagorasView class
//

#include "stdafx.h"
// SHARED_HANDLERS can be defined in an ATL project implementing preview, thumbnail
// and search filter handlers and allows sharing of document code with that project.
#ifndef SHARED_HANDLERS
#include "Pythagoras.h"
#endif

#include "PythagorasDoc.h"
#include "PythagorasView.h"
#include "ButtonMine.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CPythagorasView

IMPLEMENT_DYNCREATE(CPythagorasView, CFormView)

BEGIN_MESSAGE_MAP(CPythagorasView, CFormView)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
	ON_WM_CREATE()
	ON_BN_CLICKED(IDC_BUTTON_FILES_UP, &CPythagorasView::OnBnClickedButtonFilesUp)
	ON_BN_CLICKED(IDC_BUTTON_FILES_DOWN, &CPythagorasView::OnBnClickedButtonFilesDown)
	ON_BN_CLICKED(IDC_BUTTON_EXECUTE, &CPythagorasView::OnBnClickedButtonExecute)
	ON_BN_CLICKED(IDC_BUTTON_DEBUG, &CPythagorasView::OnBnClickedButtonDebug)
	ON_BN_CLICKED(IDC_BUTTON_DELETE, &CPythagorasView::OnBnClickedButtonDelete)
	ON_WM_SIZING()
	ON_WM_SIZE()
	ON_BN_CLICKED(IDC_BUTTON_FILES_ADD_SELECTED, &CPythagorasView::OnBnClickedButtonFilesAddSelected)
	ON_BN_CLICKED(IDC_BUTTON_FILES_DELETE, &CPythagorasView::OnBnClickedButtonFilesDelete)
	ON_BN_CLICKED(IDC_BUTTON_FILES_SELECT_ALL, &CPythagorasView::OnBnClickedButtonFilesSelectAll)
	ON_BN_CLICKED(IDC_BUTTON_FILES_DELETE_ALL, &CPythagorasView::OnBnClickedButtonFilesDeleteAll)
	ON_COMMAND(ID_RUN_SELECTED_ENTRIES, &CPythagorasView::OnRunSelectedEntries)
	ON_UPDATE_COMMAND_UI(ID_RUN_SELECTED, &CPythagorasView::OnUpdateRunSelected)
	ON_COMMAND(ID_RUN_SELECTED, &CPythagorasView::OnRunSelected)
END_MESSAGE_MAP()

// CPythagorasView construction/destruction

CPythagorasView::CPythagorasView()
	: CFormView(IDD_PYTHAGORAS_FORM)
	, m_sStaticFiles(_T(""))
{
	// TODO: add construction code here

}

CPythagorasView::~CPythagorasView()
{
}

void CPythagorasView::DoDataExchange(CDataExchange* pDX)
{
	CFormView::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_FILES, m_lstFiles);
	DDX_Control(pDX, IDC_FILE_PROPSHEET_PLACEHOLDER, m_wndPSPlaceHolder);
	DDX_Control(pDX, IDC_BUTTON_FILES_UP, m_btnFilesMoveUp);
	DDX_Control(pDX, IDC_BUTTON_FILES_DOWN, m_btnFilesMoveDown);
	DDX_Control(pDX, IDC_BUTTON_EXECUTE, m_btnFilesExecute);
	DDX_Control(pDX, IDC_BUTTON_DELETE, m_btnDeleteSharedMemory);
	DDX_Control(pDX, IDC_BUTTON_DEBUG, m_btnSetSharedMemory);
	DDX_Control(pDX, IDC_BUTTON_FILES_DELETE, m_btnFilesDelete);
	DDX_Control(pDX, IDC_BUTTON_FILES_ADD_SELECTED, m_btnFilesAddSelected);
	DDX_Control(pDX, IDC_BUTTON_FILES_SELECT_ALL, m_btnFilesSelectAll);
	DDX_Control(pDX, IDC_BUTTON_FILES_DELETE_ALL, m_btnFilesDeleteAll);
	DDX_Text(pDX, IDC_STATIC_LIST, m_sStaticFiles);
	DDX_Control(pDX, IDC_PROGRESS_EXECUTION, m_pcScriptExecution);
}

BOOL CPythagorasView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return CFormView::PreCreateWindow(cs);
}

void CPythagorasView::OnInitialUpdate()
{
	CFormView::OnInitialUpdate();
	GetParentFrame()->RecalcLayout();
	ResizeParentToFit();

	CWnd* pwndPropSheetHolder = GetDlgItem(IDC_FILE_PROPSHEET_PLACEHOLDER);
	m_FileTabWrapper.m_pFileExplorer = new PSFileExplorer(_T(""),this);
	m_FileTabWrapper.m_pFileExplorer->Create(this,
		WS_CHILD | WS_VISIBLE, 0);
	
	// fit the property sheet into the place holder window, and show it
	CRect rectPropSheet;
	pwndPropSheetHolder->GetWindowRect(rectPropSheet);
	ScreenToClient(&rectPropSheet);	
	rectPropSheet.bottom -= 4;
	m_FileTabWrapper.m_pFileExplorer->SetWindowPos(&CWnd::wndTop, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
	m_FileTabWrapper.m_pFileExplorer->MoveWindow(&rectPropSheet);

	if( AfxGetMainWnd() != NULL)
	{
		m_FileTabWrapper.m_pFileExplorer->Activate();
	}

	m_sStaticFiles = _T("Files to process:");
	m_pcScriptExecution.ShowWindow(SW_HIDE);
	m_pcScriptExecution.SetRange(0, 100);
	m_pcScriptExecution.SetStep(1);

	UpdateData(FALSE);

	m_btnFilesMoveUp.SetImage(IDB_UPARROW, RGB(0xFF, 0xFF, 0xFF));
	m_btnFilesMoveUp.SetToolTip("Move selected entries up");
	m_btnFilesMoveDown.SetImage(IDB_DOWNARROW, RGB(0xFF, 0xFF, 0xFF));
	m_btnFilesMoveDown.SetToolTip("Move selected entries down");
	m_btnFilesExecute.SetImage(IDB_EXECUTE, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesExecute.SetToolTip("Execute function");
	m_btnFilesDelete.SetImage(IDB_DELETEALL, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesDelete.SetToolTip("Remove selected entries");
	m_btnFilesAddSelected.SetImage(IDB_ADDALL, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesAddSelected.SetToolTip("Add selected entries");
	m_btnFilesSelectAll.SetImage(IDB_SELECTALL, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesSelectAll.SetToolTip("Select all entries");
	m_btnFilesDeleteAll.SetImage(IDB_DELETE, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesDeleteAll.SetToolTip("Delete all entries");
	
#ifdef _DEBUG
	m_btnDeleteSharedMemory.ShowWindow(SW_NORMAL);
	m_btnSetSharedMemory.ShowWindow(SW_NORMAL);
#endif
}

void CPythagorasView::OnRButtonUp(UINT /* nFlags */, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

void CPythagorasView::OnContextMenu(CWnd* /* pWnd */, CPoint point)
{
#ifndef SHARED_HANDLERS
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EDIT, point.x, point.y, this, TRUE);
#endif
}


// CPythagorasView diagnostics

#ifdef _DEBUG
void CPythagorasView::AssertValid() const
{
	CFormView::AssertValid();
}

void CPythagorasView::Dump(CDumpContext& dc) const
{
	CFormView::Dump(dc);
}

CPythagorasDoc* CPythagorasView::GetDocument() const // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CPythagorasDoc)));
	return (CPythagorasDoc*)m_pDocument;
}
#endif //_DEBUG


// CPythagorasView message handlers


int CPythagorasView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CFormView::OnCreate(lpCreateStruct) == -1)
		return -1;

	return 0;
}


void CPythagorasView::OnActivateView(BOOL bActivate, CView* pActivateView, CView* pDeactiveView)
{
	CFormView::OnActivateView(bActivate, pActivateView, pDeactiveView);
	if (bActivate&& AfxGetMainWnd() != NULL)
	{
		//m_FileTabWrapper.m_pFileExplorer->Activate();		
	}
}


void CPythagorasView::OnBnClickedButtonFilesUp()
{
	int i = 0;
	std::vector< int > vSelectedFiles;
	std::vector< int >::iterator itFile;

	int iCount = GetSelectedEntries(vSelectedFiles);

	for (itFile = vSelectedFiles.begin(); itFile != vSelectedFiles.end(); itFile++)
	{
		int iIndex = (*itFile);
		LPFileFolderDataStruct lpMover = theApp.m_vFileFolderData.at(iIndex);

		// Is there an item above us that isn't one to be moved up?  If so, swap spots		
		if (iIndex > 0)
		{
			if (!theApp.m_vFileFolderData.at(iIndex - 1)->bSelected)
			{
				// Swap them
				LPFileFolderDataStruct lpHold = NULL;
				_itLPFileFolderData itSwitch1, itSwitch2;
				itSwitch1 = theApp.m_vFileFolderData.begin() + iIndex - 1;
				itSwitch2 = theApp.m_vFileFolderData.begin() + iIndex;
				lpHold = (*itSwitch1);
				(*itSwitch1) = (*itSwitch2);
				(*itSwitch2) = lpHold;
			}
		}
	}
	RefreshFileList();
}

void CPythagorasView::RefreshFileList()
{
	// Take care of the list control
	m_lstFiles.ResetContent();
	int i = 0, iItem = 0;
	_itLPFileFolderData itData;
	bool bFiles = false;
	bool bFolders = false;

	for (itData = theApp.m_vFileFolderData.begin(); itData != theApp.m_vFileFolderData.end(); itData++, i++)
	{
		if ((*itData))
		{
			if ((*itData)->Type == FileType)
			{
				bFiles = true;
			}
			if ((*itData)->Type == FolderType)
			{
				bFolders = true;
			}
			iItem = m_lstFiles.AddString((*itData)->sDisplayName.c_str());
			m_lstFiles.SetItemData(iItem, reinterpret_cast< DWORD_PTR >((*itData)));
			if ((*itData)->bSelected)
			{
				m_lstFiles.SetSel(iItem);
			}
		}
	}

	if (bFiles || bFolders)
	{
		if (bFiles && bFolders)
		{
			m_sStaticFiles = _T("Files and folders");
		}
		if (bFiles && !bFolders)
		{
			m_sStaticFiles = _T("Files");
		}
		if (!bFiles && bFolders)
		{
			m_sStaticFiles = _T("Folders");
		}

		m_sStaticFiles += _T(" to process:");
	}
	else
	{
		m_sStaticFiles = _T("Files to process:");
	}
	
	UpdateData(FALSE);
}

int CPythagorasView::GetSelectedEntries(std::vector< int > &vSelection)
{
	CArray<int, int> aListBoxSel;
	int nCount = m_lstFiles.GetSelCount();

	aListBoxSel.SetSize(nCount);
	_itLPFileFolderData itData;
	nCount = m_lstFiles.GetSelItems(nCount, aListBoxSel.GetData());

	for (itData = theApp.m_vFileFolderData.begin(); itData != theApp.m_vFileFolderData.end(); itData++)
	{
		(*itData)->bSelected = false;
	}

	for (int i = 0; i < nCount; i++)
	{
		theApp.m_vFileFolderData.at(aListBoxSel.GetAt(i))->bSelected = true;
		vSelection.push_back(aListBoxSel.GetAt(i));
	}

	// Dump the selection array.
	AFXDUMP(aListBoxSel);
	//aListBoxSel.Dump();

	return nCount;
}



void CPythagorasView::OnBnClickedButtonFilesDown()
{
	int i = 0;
	std::vector< int > vSelectedEntries;
	std::vector< int >::reverse_iterator ritData;

	int iCount = GetSelectedEntries(vSelectedEntries);

	for (ritData = vSelectedEntries.rbegin(); ritData != vSelectedEntries.rend(); ritData++)
	{
		int iIndex = (*ritData);
		LPFileFolderDataStruct lpMover = theApp.m_vFileFolderData.at(iIndex);

		// Is there an item below us that isn't one to be moved down?  If so, swap spots		
		if (iIndex < theApp.m_vFileFolderData.size() - 1)
		{
			if (!theApp.m_vFileFolderData.at(iIndex + 1)->bSelected)
			{
				// Swap them
				LPFileFolderDataStruct lpHold = NULL;
				_itLPFileFolderData itSwitch1, itSwitch2;
				itSwitch1 = theApp.m_vFileFolderData.begin() + iIndex + 1;
				itSwitch2 = theApp.m_vFileFolderData.begin() + iIndex;
				lpHold = (*itSwitch1);
				(*itSwitch1) = (*itSwitch2);
				(*itSwitch2) = lpHold;
			}
		}
	}
	RefreshFileList();
}


void CPythagorasView::OnBnClickedButtonExecute()
{
	theApp.ExecuteScript();
}


void CPythagorasView::OnBnClickedButtonDebug()
{	
	theApp.DeleteSharedMemory(false);
	theApp.ExecuteScript(false);
	theApp.AddToOutput(eStandardOutput, _T("Shared memory set, launch debugger."), true);
}


void CPythagorasView::OnBnClickedButtonDelete()
{
	theApp.DeleteSharedMemory(false);
	theApp.AddToOutput(eStandardOutput, _T("Shared memory deleted."), true);
}


void CPythagorasView::OnSizing(UINT fwSide, LPRECT pRect)
{
	CFormView::OnSizing(fwSide, pRect);
}


void CPythagorasView::OnSize(UINT nType, int cx, int cy)
{
	CFormView::OnSize(nType, cx, cy);
		
	if (IsWindow(this->m_hWnd) && IsWindow(m_lstFiles.m_hWnd))
	{
		CRect rParent;
		//GetClientRect(&rParent);
		CRect rChild;

		GetWindowRect(rParent);
		ScreenToClient(&rParent);
				
		m_lstFiles.GetWindowRect(&rChild);
		ScreenToClient(&rChild);
		rChild.bottom = rParent.bottom - 5;
		m_lstFiles.MoveWindow(rChild, TRUE);


		/*
		m_wndPSPlaceHolder.GetWindowRect(&rChild);
		ScreenToClient(&rChild);
		rChild.bottom = rParent.bottom - 5;
		rChild.right = rParent.right - 5;
		m_wndPSPlaceHolder.MoveWindow(rChild, TRUE);
		*/

		/*
		if(m_FileTabWrapper.m_pFileExplorer )
		{
			m_FileTabWrapper.m_pFileExplorer->GetWindowRect(&rChild);
			ScreenToClient(&rChild);
			rChild.bottom = rParent.bottom - 5;
			rChild.right = rParent.right - 5;
			m_FileTabWrapper.m_pFileExplorer->MoveWindow(rChild, TRUE);

			m_FileTabWrapper.m_pFileExplorer->m_ppFileTabExplorer.GetWindowRect(&rChild);
			ScreenToClient(&rChild);
			rChild.bottom = rParent.bottom - 5;
			rChild.right = rParent.right - 5;
			m_FileTabWrapper.m_pFileExplorer->m_ppFileTabExplorer.MoveWindow(rChild, TRUE);
		}
		*/
	}	
}


void CPythagorasView::OnBnClickedButtonFilesAddSelected()
{
	m_FileTabWrapper.m_pFileExplorer->AddSelectedFiles();
}


void CPythagorasView::OnBnClickedButtonFilesDelete()
{	
	std::vector< int > vSelectedEntries;	
	GetSelectedEntries(vSelectedEntries);
	_itLPFileFolderData itData;

	for (itData = theApp.m_vFileFolderData.begin(); itData != theApp.m_vFileFolderData.end(); )
	{
		if ((*itData)->bSelected)
		{			
			theApp.m_DataMap.erase((*itData)->sFullName.c_str());
			delete (*itData);
			itData = theApp.m_vFileFolderData.erase(itData);
		}
		else
		{
			itData++;
		}
	}
	RefreshFileList();
}


void CPythagorasView::OnBnClickedButtonFilesSelectAll()
{
	_itLPFileFolderData itData;
	bool bSelect = (m_lstFiles.GetSelCount() == m_lstFiles.GetCount() ? false : true);
	for (itData = theApp.m_vFileFolderData.begin(); itData != theApp.m_vFileFolderData.end(); itData++)
	{
		(*itData)->bSelected = bSelect;
	}
	RefreshFileList();
}


void CPythagorasView::OnBnClickedButtonFilesDeleteAll()
{
	_itLPFileFolderData itData;
	for (itData = theApp.m_vFileFolderData.begin(); itData != theApp.m_vFileFolderData.end(); )
	{
		theApp.m_DataMap.erase((*itData)->sFullName.c_str());
		delete (*itData);
		itData = theApp.m_vFileFolderData.erase(itData);
	}
	RefreshFileList();
}


void CPythagorasView::OnRunSelectedEntries()
{
	int iIndex = m_FileTabWrapper.m_pFileExplorer->GetActiveIndex();
	switch (iIndex)
	{
	default:
	case 0:
		break;

	case 1:
	{
		m_FileTabWrapper.m_pFileExplorer->m_ppFileTabRegex.RunSelected();
	}
	break;

	case 2:
		break;
	}
}

void CPythagorasView::OnUpdateRunSelected(CCmdUI *pCmdUI)
{
	int nCount = m_FileTabWrapper.m_pFileExplorer->m_ppFileTabRegex.m_lstFiles.GetSelCount();
	pCmdUI->Enable(nCount != 0);
}

void CPythagorasView::OnRunSelected()
{
	m_FileTabWrapper.m_pFileExplorer->m_ppFileTabRegex.RunSelected();
}

void CPythagorasView::IncrementProgressBar(unsigned int nSteps /* = 1 */)
{
	for (int i = 0; i < nSteps; i++)
	{
		m_pcScriptExecution.StepIt();
	}	
}

void CPythagorasView::RegexAddFile(LPCTSTR szFileName, LPCTSTR szRoot)
{
	m_FileTabWrapper.m_pFileExplorer->m_ppFileTabRegex.AddFile(szFileName, szRoot);	
}

void CPythagorasView::RegexClear()
{
	m_FileTabWrapper.m_pFileExplorer->m_ppFileTabRegex.ClearResultList();
}

void CPythagorasView::RegexResetContent()
{	
	m_FileTabWrapper.m_pFileExplorer->m_ppFileTabRegex.ResetContent();
}

void CPythagorasView::RegexDisplayResultList()
{
	m_FileTabWrapper.m_pFileExplorer->m_ppFileTabRegex.DisplaySearchList();
}

