// CustomShellListCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "Pythagoras.h"
#include "CustomShellListCtrl.h"
#include "PythagorasView.h"

// CustomShellListCtrl

IMPLEMENT_DYNAMIC(CustomShellListCtrl, CMFCShellListCtrl)

CustomShellListCtrl::CustomShellListCtrl() : m_pRelatedTree(NULL)
{

}

CustomShellListCtrl::~CustomShellListCtrl()
{
}


BEGIN_MESSAGE_MAP(CustomShellListCtrl, CMFCShellListCtrl)
	ON_NOTIFY_REFLECT_EX(NM_DBLCLK, &CustomShellListCtrl::OnNMDblclk)
END_MESSAGE_MAP()



// CustomShellListCtrl message handlers




BOOL CustomShellListCtrl::OnNMDblclk(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	SHORT sKeyState = 0;
	bool bAltPressed = false;

	if (pNMItemActivate->iItem == -1)
	{
		return FALSE;
	}

	sKeyState = ::GetKeyState(VK_MENU);
	bAltPressed = (sKeyState & 0x8000) == 0 ? false : true;

	sKeyState = ::GetKeyState(VK_CONTROL);
	bAltPressed = bAltPressed || sKeyState & 0x8000;
	
	// ToDo: check to see if it's a file or a folder... if it's a file, we process it
	// here and add it to our list.  If not, just call our parent so that the explorer
	// view will function as if the user wanted to enter that folder.
	CString sPath;
	GetItemPath(sPath, pNMItemActivate->iItem);
	if (sPath.GetLength() > 0)
	{
		if (boost::filesystem::is_regular_file(sPath.GetBuffer()))
		{
			theApp.AddFile(sPath);
			*pResult = 0;
			return TRUE;
		}
		else
		{
			if (bAltPressed && boost::filesystem::is_directory(sPath.GetBuffer()))
			{
				theApp.AddFolder(sPath);
				*pResult = 0;
				return TRUE;
			}
			else
			{
				if (m_pRelatedTree)
				{
					m_pRelatedTree->SelectPath(sPath);
					DisplayFolder(sPath);
				}
			}			
		}
	}	
	return FALSE;
}

void CustomShellListCtrl::SetRelatedTree(CMFCShellTreeCtrl *pRelatedTree)
{	
	m_pRelatedTree = pRelatedTree;
}

void CustomShellListCtrl::AddSelectedFiles()
{
	POSITION pos = GetFirstSelectedItemPosition();	
	while (pos)
	{
		int nItem = GetNextSelectedItem(pos);
		CString sPath;
		GetItemPath(sPath, nItem);
		if (sPath.GetLength() > 0)
		{
			if (boost::filesystem::is_regular_file(sPath.GetBuffer()))
			{
				theApp.AddFile(sPath);
			}
			else
			{
				if (boost::filesystem::is_directory(sPath.GetBuffer()))
				{
					theApp.AddFolder(sPath);					
				}
			}
		}
	}
}