
#pragma once

#include "ViewTree.h"

class CClassToolBar : public CMFCToolBar
{
	virtual void OnUpdateCmdUI(CFrameWnd* /*pTarget*/, BOOL bDisableIfNoHndler)
	{
		CMFCToolBar::OnUpdateCmdUI((CFrameWnd*) GetOwner(), bDisableIfNoHndler);
	}

	virtual BOOL AllowShowOnList() const { return FALSE; }
};

class CFunctionsView : public CDockablePane
{
public:
	CFunctionsView();
	virtual ~CFunctionsView();

	void AdjustLayout();
	void OnChangeVisualStyle();
	
protected:
	CClassToolBar m_wndToolBar;
	CViewTree m_wndFunctionsView;
	CImageList m_ClassViewImages;
	UINT m_nCurrSort;
	
// Overrides
public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	void RefreshContent();
	CTreeCtrl *GetTreeCtrl()
	{
		return static_cast<CTreeCtrl *>(&m_wndFunctionsView);
	}

protected:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	afx_msg void OnPaint();
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg LRESULT OnChangeActiveTab(WPARAM, LPARAM);
	afx_msg void OnSort(UINT id);
	afx_msg void OnUpdateSort(CCmdUI* pCmdUI);

	DECLARE_MESSAGE_MAP()
};

