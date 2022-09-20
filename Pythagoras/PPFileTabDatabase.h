#pragma once


// PPFileTabDatabase dialog

class PPFileTabDatabase : public CPropertyPage
{
	DECLARE_DYNAMIC(PPFileTabDatabase)

public:
	PPFileTabDatabase();
	virtual ~PPFileTabDatabase();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_FILE_TAB_DATABASE};
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
};
