#pragma once
#include "afxwin.h"
#include "ButtonMine.h"

// PPFileTabRegex dialog
class PPFileTabRegex : public CPropertyPage
{
	DECLARE_DYNAMIC(PPFileTabRegex)

public:
	PPFileTabRegex();
	virtual ~PPFileTabRegex();

	_StringVector m_svFileList;
	_StringVector m_svShortFileList;
	_vRegexHistory m_RegexHistory;

	ButtonMine m_btnFilesSelectAll;
	ButtonMine m_btnFilesAddSelected;
	ButtonMine m_btnFilesRemoveSelected;
	ButtonMine m_btnFilesRemoveAll;
	ButtonMine m_btnFilesDeleteUnselected;
	ButtonMine m_btnFilesAddAll;
	ButtonMine m_btnRunSelected;

	int m_iSearchOption;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_FILE_TAB_REGEX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support	
	void ChangeSelected(bool bRemoveSelected);
	void RunRegex();

	DECLARE_MESSAGE_MAP()
public:
	CString m_sRootFolder;
	CString m_sRegex;
	BOOL m_bRecurseSubfolders;
	BOOL m_bClearSearchList;
	CListBox m_lstFiles;
	afx_msg void OnBnClickedButtonRunRegex();
	virtual BOOL OnInitDialog();
	CString m_sWildcard;
	afx_msg void OnBnClickedButtonRegexAddAll();
	afx_msg void OnBnClickedButtonRegexClearList();
	afx_msg void OnBnClickedButtonRegexAddSelected();
	afx_msg void OnBnClickedButtonRegexRemove();
	afx_msg void OnBnClickedButtonRegexClearUnselected();
	afx_msg void OnBnClickedButtonRegexSelectAll();
	
	CComboBox m_cbSearchOption;		
	BOOL m_bRetrieveParentFolder;
	afx_msg void OnBnClickedButtonRegexRunSelected();
	void RunSelected();
	void AddFile(LPCTSTR szFileName, LPCTSTR szRootFolder);
	void DisplaySearchList();
	void ResetContent();
	void ClearResultList();
	void ReadRegexHistory();
	void WriteRegexHistory();
	void AddToRegexHistory( unsigned int uiFileCount, unsigned int uiFolderCount);
	afx_msg void OnBnClickedButtonHistory();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	void SaveResultList();
	CEdit m_edRegexString;
	CEdit m_edWildcard;
	afx_msg void OnBnClickedButtonRegexSave();
};
