#pragma once

class CRegistrySettings
{
public:			
	std::wstring m_sScriptEditor;
	std::wstring m_sExplorerPath;
	std::wstring m_sRegexRootFolder;
	std::wstring m_sRegexWildcard;
	std::wstring m_sRegexDefault;

	std::wstring m_sFTPAddress;
	std::wstring m_sFTPUser;
	std::wstring m_sFTPPassword;
	std::wstring m_sLocalScriptsFolder;
	std::wstring m_sCommonScriptsFolder;
	
	unsigned int m_iAddFilterIndex;
	unsigned int m_iSaveFilterIndex;
	unsigned int m_iFunction;
	unsigned int m_iOpenFile;
	unsigned int m_iRegexRecurseSubfolders;
	unsigned int m_iRegexClearSearchList;
	unsigned int m_iClearOutput;
	unsigned int m_iDownloadCommon;
	unsigned int m_iRegexSearchOption;
	unsigned int m_iRegexRetrieveParentFolder;
	unsigned int m_iLocalScriptsOption;
	unsigned int m_iCommonScriptsOption;
	
	CRegistryHelper m_rhHelper;

protected:
	virtual void LoadSettings();

public:
	CRegistrySettings(void);
	virtual ~CRegistrySettings(void);		
	virtual void SaveSettings();
};
