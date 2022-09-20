// PPFileTabDatabase.cpp : implementation file
//

#include "stdafx.h"
#include "Pythagoras.h"
#include "PPFileTabDatabase.h"
#include "afxdialogex.h"


// PPFileTabDatabase dialog

IMPLEMENT_DYNAMIC(PPFileTabDatabase, CPropertyPage)

PPFileTabDatabase::PPFileTabDatabase()
	: CPropertyPage(IDD_FILE_TAB_DATABASE)
{

}

PPFileTabDatabase::~PPFileTabDatabase()
{
}

void PPFileTabDatabase::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(PPFileTabDatabase, CPropertyPage)
END_MESSAGE_MAP()


// PPFileTabDatabase message handlers
