
// MainFrm.h : interface of the CMainFrame class
//

#pragma once
#include "ScriptsView.h"
#include "FunctionsView.h"
#include "OutputWnd.h"
#include "PropertiesWnd.h"

class CMainFrame : public CFrameWndEx
{	
protected: // create from serialization only
	CMainFrame();
	DECLARE_DYNCREATE(CMainFrame)

// Attributes
public:
	COutputWnd      m_wndOutput;

// Operations
public:
	void RefreshScripts();

// Overrides
public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual BOOL LoadFrame(UINT nIDResource, DWORD dwDefaultStyle = WS_OVERLAPPEDWINDOW | FWS_ADDTOTITLE, CWnd* pParentWnd = NULL, CCreateContext* pContext = NULL);

// Implementation
public:
	virtual ~CMainFrame();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:  // control bar embedded members
	CMFCMenuBar       m_wndMenuBar;
	CMFCToolBar       m_wndToolBar;
	CMFCStatusBar     m_wndStatusBar;
	CMFCToolBarImages m_UserImages;
public:
	CScriptsView      m_wndScriptsView;
	CFunctionsView    m_wndFunctionsView;	
	CPropertiesWnd    m_wndProperties;
	bool m_bDockingInitialized;

// Generated message map functions
protected:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnViewCustomize();
	afx_msg LRESULT OnToolbarCreateNew(WPARAM wp, LPARAM lp);
	afx_msg void OnApplicationLook(UINT id);
	afx_msg void OnUpdateApplicationLook(CCmdUI* pCmdUI);
	afx_msg void OnSettingChange(UINT uFlags, LPCTSTR lpszSection);
	afx_msg void OnTvnSelchangedScriptsTree(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkScriptsTree(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnTvnSelchangedFunctionsTree(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg LRESULT OnExecuteThreadCustomMessage(WPARAM wp, LPARAM lp);
	afx_msg LRESULT OnEngineCallbackCustomMessage(WPARAM wp, LPARAM lp);
	DECLARE_MESSAGE_MAP()

	BOOL CreateDockingWindows();
	void SetDockingWindowIcons(BOOL bHiColorIcons);
public:
	afx_msg void OnViewToolbar();
	afx_msg void OnViewScripts();
	afx_msg void OnViewFunctions();

	void SetStatusBarText(LPCTSTR szText, int nPaneId = 0);
	LPScriptViewNodeStruct GetScriptViewNode();
	CScriptsView* GetScriptsView()
	{
		return &m_wndScriptsView;
	}
};


