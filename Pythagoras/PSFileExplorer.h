#pragma once

// PSFileExplorer

class PSFileExplorer : public CPropertySheet
{
	DECLARE_DYNAMIC(PSFileExplorer)

public:
	PPFileTabExplorer m_ppFileTabExplorer;
	PPFileTabRegex m_ppFileTabRegex;
	PPFileTabDatabase m_ppFileTabDatabase;

public:	
	PSFileExplorer(UINT nIDCaption, CWnd* pParentWnd = NULL, UINT iSelectPage = 0);
	PSFileExplorer(LPCTSTR pszCaption, CWnd* pParentWnd = NULL, UINT iSelectPage = 0);
	virtual ~PSFileExplorer();

	void Activate()
	{
		m_ppFileTabExplorer.Activate();
		//m_ppFileTabRegex.Activate();
		//m_ppFileTabDatabase.Activate();
	}
	void AddSelectedFiles();

protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
};


