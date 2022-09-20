#pragma once
#include "afxeditbrowsectrl.h"
#include "afxwin.h"


// AppSettingsDlg dialog
#define APP_SETTINGS_RETURN_CANCEL			0x1
#define APP_SETTINGS_RETURN_OK				0x2
#define APP_SETTINGS_RETURN_MODIFIED_PATHS	0x4

class AppSettingsDlg : public CDialogEx
{
	DECLARE_DYNAMIC(AppSettingsDlg)

public:
	AppSettingsDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~AppSettingsDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_APP_SETTINGS };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:	
	virtual BOOL OnInitDialog();
	virtual INT_PTR DoModal();
	CString m_sScriptEditor;
	CMFCEditBrowseCtrl m_ScriptEditorBrowseCtrl;
	BOOL m_bClearOutput;
	CString m_sFTPAddress;
	CString m_sFTPUsername;
	CString m_sFTPPassword;
	afx_msg void OnBnClickedButtonAppSettingsZipCommon();
	CButton m_btnZipCommon;
	BOOL m_bDownloadCommon;
	CEdit m_edFTPAddress;
	CEdit m_edFTPUsername;
	CEdit m_edFTPPassword;
	afx_msg void OnBnClickedCheckAppSettingsDownloadCommon();
	void EnableDisableControls( bool bUpdate = FALSE );
	int m_iLocalScriptsFolder;
	CMFCEditBrowseCtrl m_LocalScriptsFolderCtrl;
	afx_msg void OnBnClickedRadioLocalScriptsMydocuments();
	afx_msg void OnBnClickedRadioLocalScriptsCustomFolder();
	CString m_sLocalScriptsFolder;
	CMFCEditBrowseCtrl m_CommonScriptsFolderCtrl;
	CString m_sCommonScriptsFolder;
	int m_iCommonScriptsFolder;
	afx_msg void OnBnClickedRadioCommonScriptsMydocuments();
	afx_msg void OnBnClickedRadioCommonScriptsCustomFolder();	
};
