// AppSettingsDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Pythagoras.h"
#include "AppSettingsDlg.h"
#include "afxdialogex.h"


// AppSettingsDlg dialog

IMPLEMENT_DYNAMIC(AppSettingsDlg, CDialogEx)

AppSettingsDlg::AppSettingsDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_DIALOG_APP_SETTINGS, pParent)	
	, m_sScriptEditor(_T(""))
	, m_bClearOutput(FALSE)
	, m_sFTPAddress(_T(""))
	, m_sFTPUsername(_T(""))
	, m_sFTPPassword(_T(""))
	, m_bDownloadCommon(TRUE)
	, m_iLocalScriptsFolder(0)	
	, m_sLocalScriptsFolder(_T(""))
	, m_sCommonScriptsFolder(_T(""))
	, m_iCommonScriptsFolder(0)
{

}

AppSettingsDlg::~AppSettingsDlg()
{
}

void AppSettingsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_MFCEDITBROWSE_APP_SETTINGS_SCRIPT_EDITOR, m_sScriptEditor);
	DDX_Control(pDX, IDC_MFCEDITBROWSE_APP_SETTINGS_SCRIPT_EDITOR, m_ScriptEditorBrowseCtrl);
	DDX_Control(pDX, IDC_MFCEDITBROWSE_APP_SETTINGS_LOCAL_SCRIPT_FOLDER, m_LocalScriptsFolderCtrl);
	DDX_Check(pDX, IDC_CHECK_APP_SETTINGS_CLEAR_OUTPUT, m_bClearOutput);
	DDX_Text(pDX, IDC_EDIT_FTP_ADDRESS, m_sFTPAddress);
	DDX_Text(pDX, IDC_EDIT_FTP_USERNAME, m_sFTPUsername);
	DDX_Text(pDX, IDC_EDIT_FTP_PASSWORD, m_sFTPPassword);
	DDX_Control(pDX, IDC_BUTTON_APP_SETTINGS_ZIP_COMMON, m_btnZipCommon);
	DDX_Check(pDX, IDC_CHECK_APP_SETTINGS_DOWNLOAD_COMMON, m_bDownloadCommon);
	DDX_Control(pDX, IDC_EDIT_FTP_ADDRESS, m_edFTPAddress);
	DDX_Control(pDX, IDC_EDIT_FTP_USERNAME, m_edFTPUsername);
	DDX_Control(pDX, IDC_EDIT_FTP_PASSWORD, m_edFTPPassword);
	DDX_Radio(pDX, IDC_RADIO_LOCAL_SCRIPTS_MYDOCUMENTS, m_iLocalScriptsFolder);
	DDX_Text(pDX, IDC_MFCEDITBROWSE_APP_SETTINGS_LOCAL_SCRIPT_FOLDER, m_sLocalScriptsFolder);
	DDX_Control(pDX, IDC_MFCEDITBROWSE_APP_SETTINGS_COMMON_SCRIPT_FOLDER, m_CommonScriptsFolderCtrl);
	DDX_Text(pDX, IDC_MFCEDITBROWSE_APP_SETTINGS_COMMON_SCRIPT_FOLDER, m_sCommonScriptsFolder);
	DDX_Radio(pDX, IDC_RADIO_COMMON_SCRIPTS_MYDOCUMENTS, m_iCommonScriptsFolder);
}


BEGIN_MESSAGE_MAP(AppSettingsDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON_APP_SETTINGS_ZIP_COMMON, &AppSettingsDlg::OnBnClickedButtonAppSettingsZipCommon)
	ON_BN_CLICKED(IDC_CHECK_APP_SETTINGS_DOWNLOAD_COMMON, &AppSettingsDlg::OnBnClickedCheckAppSettingsDownloadCommon)
	ON_BN_CLICKED(IDC_RADIO_LOCAL_SCRIPTS_MYDOCUMENTS, &AppSettingsDlg::OnBnClickedRadioLocalScriptsMydocuments)
	ON_BN_CLICKED(IDC_RADIO_LOCAL_SCRIPTS_CUSTOM_FOLDER, &AppSettingsDlg::OnBnClickedRadioLocalScriptsCustomFolder)
	ON_BN_CLICKED(IDC_RADIO_COMMON_SCRIPTS_MYDOCUMENTS, &AppSettingsDlg::OnBnClickedRadioCommonScriptsMydocuments)
	ON_BN_CLICKED(IDC_RADIO_COMMON_SCRIPTS_CUSTOM_FOLDER, &AppSettingsDlg::OnBnClickedRadioCommonScriptsCustomFolder)
END_MESSAGE_MAP()


// AppSettingsDlg message handlers


BOOL AppSettingsDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
		
	m_sScriptEditor = theApp.m_RegistrySettings.m_sScriptEditor.c_str();	
	m_ScriptEditorBrowseCtrl.EnableFileBrowseButton(_T("exe"), _T("Executable files|*.exe||"));
	m_LocalScriptsFolderCtrl.EnableFolderBrowseButton(0, BIF_NEWDIALOGSTYLE);
	m_CommonScriptsFolderCtrl.EnableFolderBrowseButton(0, BIF_NEWDIALOGSTYLE);

	m_bClearOutput = theApp.m_RegistrySettings.m_iClearOutput == 1 ? true : false;
	m_bDownloadCommon = theApp.m_RegistrySettings.m_iDownloadCommon == 1 ? true : false;
	m_sFTPAddress = theApp.m_RegistrySettings.m_sFTPAddress.c_str();
	m_sFTPUsername = theApp.m_RegistrySettings.m_sFTPUser.c_str();
	m_sFTPPassword = theApp.m_RegistrySettings.m_sFTPPassword.c_str();

	m_sLocalScriptsFolder = theApp.m_RegistrySettings.m_sLocalScriptsFolder.c_str();
	m_iLocalScriptsFolder = theApp.m_RegistrySettings.m_iLocalScriptsOption;
	m_sCommonScriptsFolder = theApp.m_RegistrySettings.m_sCommonScriptsFolder.c_str();
	m_iCommonScriptsFolder = theApp.m_RegistrySettings.m_iCommonScriptsOption;
	
	if (!theApp.AllowAdmin())
	{
		m_btnZipCommon.ShowWindow(SW_HIDE);
	}
	//m_btnZipCommon.EnableWindow(theApp.AllowAdmin());
		
	EnableDisableControls(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


INT_PTR AppSettingsDlg::DoModal()
{
	INT_PTR iReturn = CDialogEx::DoModal();
#define CHECKMODIFIED_S(s1,s2,Flag)	{ std::wstring sCompare; sCompare = s2; if( _tcsicmp(s1.c_str(), sCompare.c_str())) { Flag |= APP_SETTINGS_RETURN_MODIFIED_PATHS; }}
#define CHECKMODIFIED_I(i1,i2,Flag)	{ if( i1 != i2 ) { Flag |= APP_SETTINGS_RETURN_MODIFIED_PATHS; }}
	
	if (iReturn == IDOK)
	{
		iReturn = APP_SETTINGS_RETURN_OK;
		theApp.m_RegistrySettings.m_sScriptEditor = m_sScriptEditor;
		theApp.m_RegistrySettings.m_iClearOutput = m_bClearOutput ? 1 : 0;

		if (m_bDownloadCommon && theApp.m_RegistrySettings.m_iDownloadCommon == 0)
		{
			iReturn |= APP_SETTINGS_RETURN_MODIFIED_PATHS;
		}
		theApp.m_RegistrySettings.m_iDownloadCommon = m_bDownloadCommon ? 1 : 0;
				
		//m_sFTPAddress.GetBuffer()
		CHECKMODIFIED_S(theApp.m_RegistrySettings.m_sFTPAddress, m_sFTPAddress, iReturn);
		CHECKMODIFIED_S(theApp.m_RegistrySettings.m_sFTPUser, m_sFTPUsername, iReturn);
		CHECKMODIFIED_S(theApp.m_RegistrySettings.m_sFTPPassword, m_sFTPPassword, iReturn);
		CHECKMODIFIED_S(theApp.m_RegistrySettings.m_sLocalScriptsFolder, m_sLocalScriptsFolder, iReturn);
		CHECKMODIFIED_S(theApp.m_RegistrySettings.m_sCommonScriptsFolder, m_sCommonScriptsFolder, iReturn);
		CHECKMODIFIED_I(theApp.m_RegistrySettings.m_iLocalScriptsOption, m_iLocalScriptsFolder, iReturn);
		CHECKMODIFIED_I(theApp.m_RegistrySettings.m_iCommonScriptsOption, m_iCommonScriptsFolder, iReturn);
		
		theApp.m_RegistrySettings.m_sFTPAddress = m_sFTPAddress;
		theApp.m_RegistrySettings.m_sFTPUser = m_sFTPUsername;
		theApp.m_RegistrySettings.m_sFTPPassword = m_sFTPPassword;
		theApp.m_RegistrySettings.m_iLocalScriptsOption = m_iLocalScriptsFolder;
		theApp.m_RegistrySettings.m_sLocalScriptsFolder = m_sLocalScriptsFolder;
		theApp.m_RegistrySettings.m_iCommonScriptsOption = m_iCommonScriptsFolder;
		theApp.m_RegistrySettings.m_sCommonScriptsFolder = m_sCommonScriptsFolder;
	}
	else
	{
		if (iReturn == IDCANCEL)
		{
			iReturn = APP_SETTINGS_RETURN_CANCEL;
		}
	}
	return iReturn;
}


void AppSettingsDlg::OnBnClickedButtonAppSettingsZipCommon()
{
	theApp.UploadCommonFiles();
}


void AppSettingsDlg::OnBnClickedCheckAppSettingsDownloadCommon()
{
	EnableDisableControls(TRUE);
}

void AppSettingsDlg::EnableDisableControls(bool bUpdate /* = FALSE */)
{
	UpdateData(bUpdate);
	m_edFTPAddress.EnableWindow(m_bDownloadCommon);
	m_edFTPUsername.EnableWindow(m_bDownloadCommon);
	m_edFTPPassword.EnableWindow(m_bDownloadCommon);

	m_LocalScriptsFolderCtrl.EnableWindow(m_iLocalScriptsFolder == 1);
	m_CommonScriptsFolderCtrl.EnableWindow(m_iCommonScriptsFolder == 1);
}


void AppSettingsDlg::OnBnClickedRadioLocalScriptsMydocuments()
{
	EnableDisableControls(TRUE);
}


void AppSettingsDlg::OnBnClickedRadioLocalScriptsCustomFolder()
{
	EnableDisableControls(TRUE);
}


void AppSettingsDlg::OnBnClickedRadioCommonScriptsMydocuments()
{
	EnableDisableControls(TRUE);
}


void AppSettingsDlg::OnBnClickedRadioCommonScriptsCustomFolder()
{
	EnableDisableControls(TRUE);
}
