// RegExHistoryDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Pythagoras.h"
#include "RegExHistoryDlg.h"
#include "afxdialogex.h"


// RegExHistoryDlg dialog

IMPLEMENT_DYNAMIC(RegExHistoryDlg, CDialogEx)

RegExHistoryDlg::RegExHistoryDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_DIALOG_REGEX_HISTORY, pParent), m_pRegexHistory(NULL)
{

}

RegExHistoryDlg::~RegExHistoryDlg()
{
}

void RegExHistoryDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_ENTRIES, m_lstEntries);
}


BEGIN_MESSAGE_MAP(RegExHistoryDlg, CDialogEx)
	ON_NOTIFY(NM_DBLCLK, IDC_LIST_ENTRIES, &RegExHistoryDlg::OnNMDblclkListEntries)
	ON_BN_CLICKED(IDC_BUTTON_CLEAR_HISTORY, &RegExHistoryDlg::OnBnClickedButtonClearHistory)
END_MESSAGE_MAP()


// RegExHistoryDlg message handlers


BOOL RegExHistoryDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	CRect rect;
	m_lstEntries.GetClientRect(&rect);
	
	m_lstEntries.InsertColumn(0, _T("Date"), LVCFMT_LEFT, rect.Width() * 0.17);
	m_lstEntries.InsertColumn(1, _T("Root Folder"), LVCFMT_LEFT, rect.Width() * 0.3);
	m_lstEntries.InsertColumn(2, _T("Wildcard"), LVCFMT_LEFT, rect.Width() * 0.1);
	m_lstEntries.InsertColumn(3, _T("RegEx"), LVCFMT_LEFT, rect.Width() * 0.20);
	m_lstEntries.InsertColumn(4, _T("File Matches"), LVCFMT_LEFT, rect.Width() * 0.11);
	m_lstEntries.InsertColumn(5, _T("Folder Matches"), LVCFMT_LEFT, rect.Width() * 0.13);

	RefreshEntries();

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

void RegExHistoryDlg::RefreshEntries()
{
	m_lstEntries.DeleteAllItems();
	if (m_pRegexHistory)
	{
		std::wstringstream ssFormat;
		int i = 0;
		for (LPRegexHistory lpEntry : *m_pRegexHistory)
		{
			LVITEM lvi;
			memset(&lvi, 0, sizeof(LVITEM));

			lvi.mask = LVIF_TEXT | LVIF_PARAM;
			lvi.iItem = i++;
			lvi.iSubItem = 0;
			lvi.pszText = (LPTSTR)(LPCTSTR)(lpEntry->sDate.c_str());
			lvi.lParam = (LPARAM)lpEntry;
			int iItem = m_lstEntries.InsertItem(&lvi);
			m_lstEntries.SetItemText(iItem, 1, (LPTSTR)(LPCTSTR)(lpEntry->sRootFolder.c_str()));
			m_lstEntries.SetItemText(iItem, 2, (LPTSTR)(LPCTSTR)(lpEntry->sWildcard.c_str()));
			m_lstEntries.SetItemText(iItem, 3, (LPTSTR)(LPCTSTR)(lpEntry->sRegEx.c_str()));
			ssFormat.str(_T(""));
			ssFormat << lpEntry->uiFileMatchCount;
			m_lstEntries.SetItemText(iItem, 4, (LPTSTR)(LPCTSTR)(ssFormat.str().c_str()));
			ssFormat.str(_T(""));
			ssFormat << lpEntry->uiFolderMatchCount;
			m_lstEntries.SetItemText(iItem, 5, (LPTSTR)(LPCTSTR)(ssFormat.str().c_str()));
		}
	}
	m_pSelectedEntry = NULL;
}

void RegExHistoryDlg::OnOK()
{
	if (m_lstEntries.GetSelectedCount() > 0)
	{
		int iSel = m_lstEntries.GetNextItem(-1, LVNI_SELECTED);		
		m_pSelectedEntry = reinterpret_cast<LPRegexHistory>(m_lstEntries.GetItemData(iSel));
	}

	CDialogEx::OnOK();
}


void RegExHistoryDlg::OnNMDblclkListEntries(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	OnOK();
	*pResult = 0;
}


void RegExHistoryDlg::OnBnClickedButtonClearHistory()
{
	int iRet = AfxMessageBox(_T("Are you sure you want to clear all history entries?"), MB_YESNO);
	if (iRet == IDYES)
	{
		EraseDataVector<std::vector<LPRegexHistory>>(*m_pRegexHistory);
		RefreshEntries();
	}
}
