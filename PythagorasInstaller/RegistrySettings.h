#pragma once

class CRegistrySettings
{
public:			
	unsigned int m_iInstallPython;
	unsigned int m_iInstallPythonModules;
	unsigned int m_iInstallVisualC;
	unsigned int m_iInstallPythagoras;
	unsigned int m_iCreateStartMenuShortcut;
	unsigned int m_iCreateDesktopShortcut;

	int m_iVisualCInstalled;
	
	CRegistryHelper m_rhHelper;

protected:
	virtual void LoadSettings();

public:
	CRegistrySettings(void);
	virtual ~CRegistrySettings(void);		
	virtual void SaveSettings();
};
