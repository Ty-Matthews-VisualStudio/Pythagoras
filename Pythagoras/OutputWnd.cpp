
#include "stdafx.h"

#include "OutputWnd.h"
#include "Resource.h"
#include "MainFrm.h"
#include "Pythagoras.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// COutputBar

COutputWnd::COutputWnd()
{
}

COutputWnd::~COutputWnd()
{
}

BEGIN_MESSAGE_MAP(COutputWnd, CDockablePane)
	ON_WM_CREATE()
	ON_WM_SIZE()	
END_MESSAGE_MAP()

int COutputWnd::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDockablePane::OnCreate(lpCreateStruct) == -1)
		return -1;

	CRect rectDummy;
	rectDummy.SetRectEmpty();

	// Create tabs window:
	if (!m_wndTabs.Create(CMFCTabCtrl::STYLE_FLAT, rectDummy, this, 1))
	{
		TRACE0("Failed to create output tab window\n");
		return -1;      // fail to create
	}

	// Create output panes:
	const DWORD dwStyle = LBS_NOINTEGRALHEIGHT | WS_CHILD | WS_VISIBLE | WS_HSCROLL | WS_VSCROLL;

	if (!m_wndDescription.Create(dwStyle, rectDummy, &m_wndTabs, 2) ||
		!m_wndStandardOutput.Create(dwStyle, rectDummy, &m_wndTabs, 3) ||
		!m_wndStandardError.Create(dwStyle, rectDummy, &m_wndTabs, 4))
	{
		TRACE0("Failed to create output windows\n");
		return -1;      // fail to create
	}

	UpdateFonts();

	CString strTabName;
	BOOL bNameValid;

	// Attach list windows to tab:
	bNameValid = strTabName.LoadString(IDS_DESCRIPTION_TAB);
	ASSERT(bNameValid);
	m_wndTabs.AddTab(&m_wndDescription, strTabName, (UINT)0);
	bNameValid = strTabName.LoadString(IDS_STDOUT_TAB);
	ASSERT(bNameValid);
	m_wndTabs.AddTab(&m_wndStandardOutput, strTabName, (UINT)1);
	bNameValid = strTabName.LoadString(IDS_STDERR_TAB);
	ASSERT(bNameValid);
	m_wndTabs.AddTab(&m_wndStandardError, strTabName, (UINT)2);

	return 0;
}

void COutputWnd::SetOutput(OutputTypeEnum eType, const std::wstring &sOutput, bool bSelect /*= false */)
{
	SetOutput(eType, sOutput.c_str(), bSelect);
}

COutputList *COutputWnd::GetOutputList(OutputTypeEnum eType)
{
	COutputList *pwndOutput = NULL;
	switch (eType)
	{
	default:
	case eDescription:
	{
		pwndOutput = &m_wndDescription;
	}
	break;

	case eStandardOutput:
	{
		pwndOutput = &m_wndStandardOutput;
	}
	break;

	case eStandardError:
	{
		pwndOutput = &m_wndStandardError;
	}
	break;
	}

	return pwndOutput;
}

void COutputWnd::SetOutput(OutputTypeEnum eType, const wchar_t *pOutput, bool bSelect /*= false */)
{	
	COutputList *pwndOutput = GetOutputList(eType);
	pwndOutput->ResetContent();
	AddToOutput(eType, pOutput, bSelect);
}

void COutputWnd::AddToOutput(OutputTypeEnum eType, const std::wstring &sOutput, bool bSelect /*= false */)
{
	AddToOutput(eType, sOutput.c_str(), bSelect);	
}

void COutputWnd::AddToOutput(OutputTypeEnum eType, const wchar_t *pOutput, bool bSelect /*= false */)
{
	COutputList *pwndOutput = GetOutputList(eType);	
	AddToOutput(pwndOutput, pOutput);
	if (bSelect)
	{
		m_wndTabs.SetActiveTab(eType);
	}
}

void COutputWnd::AddToOutput(COutputList *pOutputList, const std::wstring &sOutput)
{
	AddToOutput(pOutputList, sOutput.c_str());
}

void COutputWnd::AddToOutput(COutputList *pOutputList, const wchar_t *pOutput)
{
	_StringVector sv;
	boost::split(sv, pOutput, boost::is_any_of("\n\r"));
	for (_itString it = sv.begin(); it != sv.end(); it++)
	{		
		boost::trim_right_if((*it), boost::is_any_of(" \n\r"));
		if ((*it).size() > 0)
		{
			pOutputList->AddString((*it).c_str());
		}
	}	
}

void COutputWnd::OnSize(UINT nType, int cx, int cy)
{
	CDockablePane::OnSize(nType, cx, cy);

	// Tab control should cover the whole client area:
	m_wndTabs.SetWindowPos (NULL, -1, -1, cx, cy, SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOZORDER);
}

void COutputWnd::AdjustHorzScroll(CListBox& wndListBox)
{
	CClientDC dc(this);
	CFont* pOldFont = dc.SelectObject(&afxGlobalData.fontRegular);

	int cxExtentMax = 0;

	for (int i = 0; i < wndListBox.GetCount(); i ++)
	{
		CString strItem;
		wndListBox.GetText(i, strItem);

		cxExtentMax = max(cxExtentMax, (int)dc.GetTextExtent(strItem).cx);
	}

	wndListBox.SetHorizontalExtent(cxExtentMax);
	dc.SelectObject(pOldFont);
}

void COutputWnd::UpdateFonts()
{
	m_wndDescription.SetFont(&afxGlobalData.fontRegular);
	m_wndStandardOutput.SetFont(&afxGlobalData.fontRegular);
	m_wndStandardError.SetFont(&afxGlobalData.fontRegular);
}

void COutputWnd::SelectOutput(OutputTypeEnum eType)
{
	CMainFrame *pMainFrame = theApp.GetMainFrame();
	if (pMainFrame)
	{
		pMainFrame->SetFocus();
		pMainFrame->ShowPane(this, TRUE, FALSE, TRUE);
		pMainFrame->m_wndOutput.m_wndTabs.SetActiveTab(eType);
		pMainFrame->RecalcLayout();
	}
}

void COutputWnd::ClearOutput()
{	
	m_wndStandardOutput.ResetContent();
	m_wndStandardError.ResetContent();
}

/////////////////////////////////////////////////////////////////////////////
// COutputList1

COutputList::COutputList()
{
}

COutputList::~COutputList()
{
}

BEGIN_MESSAGE_MAP(COutputList, CListBox)
	ON_WM_CONTEXTMENU()
	ON_COMMAND(ID_EDIT_COPY, OnEditCopy)
	ON_COMMAND(ID_EDIT_CLEAR, OnEditClear)
	ON_COMMAND(ID_VIEW_OUTPUTWND, OnViewOutput)
	ON_COMMAND(ID_EDIT_SAVE, OnEditSave)
	ON_WM_WINDOWPOSCHANGING()
END_MESSAGE_MAP()
/////////////////////////////////////////////////////////////////////////////
// COutputList message handlers

void COutputList::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_OUTPUT_POPUP, point.x, point.y, this, TRUE);
#if 0
	CMenu menu;
	menu.LoadMenu(IDR_OUTPUT_POPUP);

	CMenu* pSumMenu = menu.GetSubMenu(0);

	if (AfxGetMainWnd()->IsKindOf(RUNTIME_CLASS(CMDIFrameWndEx)))
	{
		CMFCPopupMenu* pPopupMenu = new CMFCPopupMenu;

		if (!pPopupMenu->Create(this, point.x, point.y, (HMENU)pSumMenu->m_hMenu, FALSE, TRUE))
			return;

		((CMDIFrameWndEx*)AfxGetMainWnd())->OnShowPopupMenu(pPopupMenu);
		UpdateDialogControls(this, FALSE);
	}
#endif

	SetFocus();
}

void COutputList::OnEditCopy()
{
	std::wstring sCopy;
	int n = GetCount();
	for (int i = 0; i < n; i++)
	{
		CString sLine;
		GetText(i, sLine);
		sCopy += sLine.GetBuffer();
		sCopy += _T("\r\n");
	}
	CopyToClipboard(sCopy);
}

void COutputList::OnEditSave()
{
	std::wstring sFilter(_T("Text files (*.txt)|*.txt|All files (*.*)|*.*||"));
	CFileDialog FileDlg(FALSE, _T(".txt"), NULL, OFN_OVERWRITEPROMPT, sFilter.c_str(), NULL, NULL);
	CString str;
	int nMaxFiles = 256;
	int nBufferSz = nMaxFiles * _MAX_PATH + 1;
	FileDlg.GetOFN().lpstrFile = str.GetBuffer(nBufferSz);
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

		FILE *wp = NULL;
		errno_t err = _wfopen_s(&wp, FileName, _T("w"));
		if (err == 0)
		{
			int n = GetCount();
			for (int i = 0; i < n; i++)
			{
				CString sLine;
				GetText(i, sLine);
				fwprintf_s(wp, _T("%s\n"), sLine.GetBuffer());
			}
			fclose(wp);
		}
	}
}


void COutputList::OnEditClear()
{
	ResetContent();
}

void COutputList::OnViewOutput()
{
	ViewOutput();
}

void COutputList::ViewOutput()
{
	CDockablePane* pParentBar = DYNAMIC_DOWNCAST(CDockablePane, GetOwner());
	CMDIFrameWndEx* pMainFrame = DYNAMIC_DOWNCAST(CMDIFrameWndEx, GetTopLevelFrame());

	if (pMainFrame != NULL && pParentBar != NULL)
	{
		pMainFrame->SetFocus();
		pMainFrame->ShowPane(pParentBar, TRUE, FALSE, FALSE);
		pMainFrame->RecalcLayout();
	}
}


