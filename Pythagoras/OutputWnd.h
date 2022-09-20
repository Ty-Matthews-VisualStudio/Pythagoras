
#pragma once

/////////////////////////////////////////////////////////////////////////////
// COutputList window

class COutputList : public CListBox
{
// Construction
public:
	COutputList();

// Implementation
public:
	virtual ~COutputList();
	void ViewOutput();

protected:
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	afx_msg void OnEditCopy();
	afx_msg void OnEditClear();
	afx_msg void OnViewOutput();
	afx_msg void OnEditSave();

	DECLARE_MESSAGE_MAP()
};

class COutputWnd : public CDockablePane
{
// Construction
public:
	COutputWnd();

	void UpdateFonts();

// Attributes
public:
	CMFCTabCtrl	m_wndTabs;

	COutputList m_wndDescription;
	COutputList m_wndStandardOutput;
	COutputList m_wndStandardError;

protected:	
	void AdjustHorzScroll(CListBox& wndListBox);

// Implementation
public:
	virtual ~COutputWnd();
	void SelectOutput(OutputTypeEnum eType);
	void SetOutput(OutputTypeEnum eType, const std::wstring &sOutput, bool bSelect = false );
	void SetOutput(OutputTypeEnum eType, const wchar_t *pOutput, bool bSelect = false);
	void AddToOutput(OutputTypeEnum eType, const std::wstring &sOutput, bool bSelect = false);
	void AddToOutput(OutputTypeEnum eType, const wchar_t *pOutput, bool bSelect = false);
	void AddToOutput(COutputList *pOutputList, const std::wstring &sOutput);
	void AddToOutput(COutputList *pOutputList, const wchar_t *pOutput);
	COutputList *GetOutputList(OutputTypeEnum eType);
	void ClearOutput();

protected:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);

	DECLARE_MESSAGE_MAP()
public:
	
};

