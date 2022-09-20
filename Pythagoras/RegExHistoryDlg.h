#pragma once
#include "afxcmn.h"


// RegExHistoryDlg dialog

class RegExHistoryDlg : public CDialogEx
{
	DECLARE_DYNAMIC(RegExHistoryDlg)

public:
	RegExHistoryDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~RegExHistoryDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_REGEX_HISTORY };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CListCtrl m_lstEntries;
	_vRegexHistory *m_pRegexHistory;
	LPRegexHistory m_pSelectedEntry;
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	void RefreshEntries();
	afx_msg void OnNMDblclkListEntries(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButtonClearHistory();
};
