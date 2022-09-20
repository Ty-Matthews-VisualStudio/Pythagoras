#include "StdAfx.h"
#include "RegistrySettings.h"
#include "../Common/RegistryHelper.h"

CRegistrySettings::CRegistrySettings(void)
{
	m_rhHelper.SetMainKey( HKEY_CURRENT_USER );
	m_rhHelper.SetBaseSubKey( _T("Software\\3M\\Pythagoras"));
	m_rhHelper.AddItem( &m_sExplorerPath, _T(""), _T("Explorer path"), _T("") );
	m_rhHelper.AddItem( &m_sRegexRootFolder, _T(""), _T("Regex root folder"), _T(""));
	m_rhHelper.AddItem(&m_sRegexDefault, _T(""), _T("Default regex"), _T(""));
	m_rhHelper.AddItem(&m_sRegexWildcard, _T("*.csv"), _T("Regex wildcard"), _T(""));
	m_rhHelper.AddItem(&m_sScriptEditor, _T("notepad.exe"), _T("Script editor"), _T(""));

	m_rhHelper.AddItem(&m_sFTPAddress, _T("ecpweb.mmm.com"), _T("FTP address"), _T(""));
	m_rhHelper.AddItem(&m_sFTPUser, _T("Pythagoras"), _T("FTP user"), _T(""));
	m_rhHelper.AddItem(&m_sFTPPassword, _T("PythonFTP2017!"), _T("FTP password"), _T(""));
	m_rhHelper.AddItem(&m_sLocalScriptsFolder, _T(""), _T("Local scripts folder"), _T(""));
	m_rhHelper.AddItem(&m_sCommonScriptsFolder, _T(""), _T("Common scripts folder"), _T(""));
	
	m_rhHelper.AddItem( &m_iFunction, 0, 0, 0, _T("Function"), _T(""), REGISTRY_MIN );
	m_rhHelper.AddItem( &m_iOpenFile, 1, 0, 1, _T("Open file"), _T(""), REGISTRY_MIN | REGISTRY_MAX );
	m_rhHelper.AddItem( &m_iAddFilterIndex, 0, 0, 0, _T("Add file filter index"), _T(""), 0 );
	m_rhHelper.AddItem( &m_iSaveFilterIndex, 0, 0, 0, _T("Save file filter index"), _T(""), 0 );
	m_rhHelper.AddItem(&m_iRegexRecurseSubfolders, 1, 0, 1, _T("Regex recurse subfolders"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iRegexClearSearchList, 1, 0, 1, _T("Regex clear search list"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iClearOutput, 0, 0, 1, _T("Clear output before executing function"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iDownloadCommon, 1, 0, 1, _T("Download common scripts"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iRegexSearchOption, 0, eFilenames, eFilesAndFolders, _T("Regex search option"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iRegexRetrieveParentFolder, 0, 0, 1, _T("Regex retrieve parent folder"), _T(""), REGISTRY_MIN | REGISTRY_MAX);

	m_rhHelper.AddItem(&m_iLocalScriptsOption, 0, 0, 1, _T("Local scripts location"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iCommonScriptsOption, 0, 0, 1, _T("Common scripts location"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	
	LoadSettings();	
}

CRegistrySettings::~CRegistrySettings(void)
{
	SaveSettings();
}

void CRegistrySettings::LoadSettings()
{
	m_rhHelper.ReadRegistry();
	m_rhHelper.WriteRegistry();
}

void CRegistrySettings::SaveSettings()
{
	m_rhHelper.WriteRegistry();
}

