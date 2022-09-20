#pragma once


// CNewScriptDlg dialog

class CNewScriptDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CNewScriptDlg)

public:
	CNewScriptDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CNewScriptDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_NEW_SCRIPT };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CString m_sName;
};
