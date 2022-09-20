
// PythagorasEngineDlg.h : header file
//

#pragma once

#include "PythonEngine.h"

// CPythagorasEngineDlg dialog
class CPythagorasEngineDlg : public CDHtmlDialog
{
// Construction
public:
	CPythagorasEngineDlg(CWnd* pParent = NULL);	// standard constructor
	PythonEngine::LPExecuteParams m_pExecuteParams;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PYTHAGORASENGINE_DIALOG, IDH = IDR_HTML_PYTHAGORASENGINE_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

	HRESULT OnButtonOK(IHTMLElement *pElement);
	HRESULT OnButtonCancel(IHTMLElement *pElement);
	
// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	DECLARE_DHTML_EVENT_MAP()
public:
	virtual void OnDocumentComplete(LPDISPATCH pDisp, LPCTSTR szUrl);
	virtual void OnBeforeNavigate(LPDISPATCH pDisp, LPCTSTR szUrl);
};
