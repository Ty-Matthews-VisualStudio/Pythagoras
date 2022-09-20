
// PythagorasView.h : interface of the CPythagorasView class
//

#pragma once
#include "afxcmn.h"
#include "afxvslistbox.h"
#include "afxwin.h"
#include "ButtonMine.h"

class CPythagorasDoc;

class FileTabWrapper
{
public:
	PSFileExplorer *m_pFileExplorer;

	FileTabWrapper() : m_pFileExplorer(NULL) {};
	~FileTabWrapper() {
		if (m_pFileExplorer) delete m_pFileExplorer;
		m_pFileExplorer = NULL;
	};
};

class BitmapButtonWrapper
{
public:
	CBitmapButton m_Button;

	BitmapButtonWrapper() {};
	~BitmapButtonWrapper() {};	
	void LoadBitmaps(UINT nID, UINT nIDSel = 0, UINT nIDFocus = 0, UINT nIDDisabled = 0)
	{
		m_Button.LoadBitmaps(nID, nIDSel, nIDFocus, nIDDisabled);		
	}
	void SetSubClass(CButton &Sub, UINT nID, CWnd *pParentWnd)
	{
		LONG lStyle = GetWindowLong(Sub.m_hWnd, GWL_STYLE);
		SetWindowLong(Sub.m_hWnd, GWL_STYLE, lStyle | BS_OWNERDRAW);		
		Sub.SubclassDlgItem(nID, pParentWnd);
	}
	void SizeToContent()
	{
		m_Button.SizeToContent();
	}
};

class CPythagorasView : public CFormView
{
protected: // create from serialization only
	CPythagorasView();
	DECLARE_DYNCREATE(CPythagorasView)

public:
#ifdef AFX_DESIGN_TIME
	enum{ IDD = IDD_PYTHAGORAS_FORM };
#endif

// Attributes
public:
	CPythagorasDoc* GetDocument() const;

// Operations
public:
	FileTabWrapper m_FileTabWrapper;	
	BitmapButtonWrapper m_btnMoveUp;
	BitmapButtonWrapper m_btnMoveDown;
	BitmapButtonWrapper m_btnExecute;
	ButtonMine m_btnFilesMoveUp;
	ButtonMine m_btnFilesMoveDown;
	ButtonMine m_btnFilesExecute;
	ButtonMine m_btnFilesDelete;
	ButtonMine m_btnFilesAddSelected;
	ButtonMine m_btnFilesSelectAll;
	ButtonMine m_btnFilesDeleteAll;

// Overrides
public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual void OnInitialUpdate(); // called first time after construct

// Implementation
public:
	virtual ~CPythagorasView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:
	

// Generated message map functions
protected:
	afx_msg void OnFilePrintPreview();
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	DECLARE_MESSAGE_MAP()
public:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	virtual void OnActivateView(BOOL bActivate, CView* pActivateView, CView* pDeactiveView);

	void AddFile(LPCTSTR szFileName);
	void RegexAddFile(LPCTSTR szFileName, LPCTSTR szRoot);
	void RegexClear();
	void RegexResetContent();
	void RegexDisplayResultList();
	int GetSelectedEntries(std::vector< int > &vSelection);
	void RefreshFileList();
	CListBox m_lstFiles;
	afx_msg void OnBnClickedButtonFilesUp();
	afx_msg void OnBnClickedButtonFilesDown();
	afx_msg void OnBnClickedButtonExecute();
	afx_msg void OnBnClickedButtonDebug();
	afx_msg void OnBnClickedButtonDelete();
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	CStatic m_wndPSPlaceHolder;			
	CButton m_btnDeleteSharedMemory;
	CButton m_btnSetSharedMemory;
	afx_msg void OnBnClickedButtonFilesAddSelected();		
	afx_msg void OnBnClickedButtonFilesDelete();	
	afx_msg void OnBnClickedButtonFilesSelectAll();	
	afx_msg void OnBnClickedButtonFilesDeleteAll();
	CString m_sStaticFiles;
	afx_msg void OnRunSelectedEntries();
	afx_msg void OnUpdateRunSelected(CCmdUI *pCmdUI);
	afx_msg void OnRunSelected();	
	CProgressCtrl m_pcScriptExecution;
	void IncrementProgressBar(unsigned int nSteps = 1);
};

#ifndef _DEBUG  // debug version in PythagorasView.cpp
inline CPythagorasDoc* CPythagorasView::GetDocument() const
   { return reinterpret_cast<CPythagorasDoc*>(m_pDocument); }
#endif

