#pragma once


// CustomShellListCtrl

class CustomShellListCtrl : public CMFCShellListCtrl
{
	DECLARE_DYNAMIC(CustomShellListCtrl)

protected:
	CMFCShellTreeCtrl *m_pRelatedTree;
public:
	CustomShellListCtrl();
	virtual ~CustomShellListCtrl();

	void SetRelatedTree(CMFCShellTreeCtrl *pRelatedTree);
	void AddSelectedFiles();

protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg BOOL OnNMDblclk(NMHDR *pNMHDR, LRESULT *pResult);
};


