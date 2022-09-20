#pragma once
#include "afxshelltreectrl.h"
#include "afxshelllistctrl.h"
#include "CustomShellListCtrl.h"

#ifndef __PPFileTabExplorer_H__
#define __PPFileTabExplorer_H__

// PPFileTabExplorer dialog

class PPFileTabExplorer : public CPropertyPage
{
	DECLARE_DYNAMIC(PPFileTabExplorer)

public:
	PPFileTabExplorer();
	virtual ~PPFileTabExplorer();
	void Activate();


// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_FILE_TAB_EXPLORER };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CMFCShellTreeCtrl m_wndExplorerTree;	
	CustomShellListCtrl m_wndExplorerList;
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized);
	virtual BOOL OnInitDialog();
	void AddSelectedFiles();
};

#endif		// #ifndef __PPFileTabExplorer_H__