// NewScriptDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Pythagoras.h"
#include "NewScriptDlg.h"
#include "afxdialogex.h"


// CNewScriptDlg dialog

IMPLEMENT_DYNAMIC(CNewScriptDlg, CDialogEx)

CNewScriptDlg::CNewScriptDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_DIALOG_NEW_SCRIPT, pParent)
	, m_sName(_T(""))
{

}

CNewScriptDlg::~CNewScriptDlg()
{
}

void CNewScriptDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_NAME, m_sName);
}


BEGIN_MESSAGE_MAP(CNewScriptDlg, CDialogEx)
END_MESSAGE_MAP()


// CNewScriptDlg message handlers
