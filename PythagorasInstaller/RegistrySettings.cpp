#include "StdAfx.h"
#include "RegistrySettings.h"
#include "../Common/RegistryHelper.h"

CRegistrySettings::CRegistrySettings(void)
{
	m_rhHelper.SetMainKey( HKEY_CURRENT_USER );
	m_rhHelper.SetBaseSubKey( _T("Software\\3M\\PythagorasInstaller"));
	m_rhHelper.AddItem( &m_iInstallPython, 1, 0, 1, _T("Install Python"), _T(""), REGISTRY_MIN | REGISTRY_MAX );
	m_rhHelper.AddItem(&m_iInstallPythonModules, 1, 0, 1, _T("Install Python Modules"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iInstallVisualC, 1, 0, 1, _T("Install Visual C++"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iInstallPythagoras, 1, 0, 1, _T("Install Pythagoras"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iCreateStartMenuShortcut, 1, 0, 1, _T("Create Start Menu Shortcut"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	m_rhHelper.AddItem(&m_iCreateDesktopShortcut, 1, 0, 1, _T("Create Desktop Shortcut"), _T(""), REGISTRY_MIN | REGISTRY_MAX);
	
	LoadSettings();	

	CRegistryHelper rhHelper;
	rhHelper.SetMainKey(HKEY_LOCAL_MACHINE);
	rhHelper.SetBaseSubKey(_T("Software\\Microsoft\\VisualStudio\\14.0\\VC\\Runtimes\\x64"));
	rhHelper.AddItem(&m_iVisualCInstalled, -1, 0, 1, _T("Installed"), _T(""), 0);
	rhHelper.ReadRegistry();
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

