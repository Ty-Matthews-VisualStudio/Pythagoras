// PPFileTabRegex.cpp : implementation file
//

#include "stdafx.h"
#include "Pythagoras.h"
#include "PPFileTabRegex.h"
#include "afxdialogex.h"
#include "RegExHistoryDlg.h"

// PPFileTabRegex dialog

IMPLEMENT_DYNAMIC(PPFileTabRegex, CPropertyPage)

PPFileTabRegex::PPFileTabRegex()
	: CPropertyPage(IDD_FILE_TAB_REGEX)
	, m_sRootFolder(_T(""))
	, m_sRegex(_T(""))
	, m_bRecurseSubfolders(FALSE)
	, m_bClearSearchList(FALSE)
	, m_sWildcard(_T(""))
	, m_iSearchOption(0)
	, m_bRetrieveParentFolder(FALSE)
{

}

PPFileTabRegex::~PPFileTabRegex()
{
	WriteRegexHistory();
	EraseDataVector<std::vector<LPRegexHistory>>(m_RegexHistory);
}

void PPFileTabRegex::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_MFCEDITBROWSE_ROOT_FOLDER, m_sRootFolder);
	DDX_Text(pDX, IDC_EDIT_REGEX, m_sRegex);
	DDX_Check(pDX, IDC_CHECK_RECURSE_SUBFOLDERS, m_bRecurseSubfolders);
	DDX_Check(pDX, IDC_CHECK_CLEAR_SEARCH_LIST, m_bClearSearchList);
	DDX_Control(pDX, IDC_LIST_SEARCH_RESULTS, m_lstFiles);
	DDX_Text(pDX, IDC_EDIT_REGEX_WILDCARD, m_sWildcard);
	DDX_Control(pDX, IDC_BUTTON_REGEX_SELECT_ALL, m_btnFilesSelectAll);
	DDX_Control(pDX, IDC_BUTTON_REGEX_ADD_SELECTED, m_btnFilesAddSelected);
	DDX_Control(pDX, IDC_BUTTON_REGEX_REMOVE, m_btnFilesRemoveSelected);
	DDX_Control(pDX, IDC_BUTTON_REGEX_CLEAR_LIST, m_btnFilesRemoveAll);
	DDX_Control(pDX, IDC_BUTTON_REGEX_CLEAR_UNSELECTED, m_btnFilesDeleteUnselected);
	DDX_Control(pDX, IDC_BUTTON_REGEX_ADD_ALL, m_btnFilesAddAll);
	DDX_Control(pDX, IDC_COMBO_REGEX_SEARCH_OPTION, m_cbSearchOption);
	DDX_CBIndex(pDX, IDC_COMBO_REGEX_SEARCH_OPTION, m_iSearchOption);
	DDX_Check(pDX, IDC_CHECK_REGEX_RETRIEVE_PARENT_FOLDER, m_bRetrieveParentFolder);
	DDX_Control(pDX, IDC_BUTTON_REGEX_RUN_SELECTED, m_btnRunSelected);
	DDX_Control(pDX, IDC_EDIT_REGEX, m_edRegexString);
	DDX_Control(pDX, IDC_EDIT_REGEX_WILDCARD, m_edWildcard);
}


BEGIN_MESSAGE_MAP(PPFileTabRegex, CPropertyPage)
	ON_BN_CLICKED(IDC_BUTTON_RUN_REGEX, &PPFileTabRegex::OnBnClickedButtonRunRegex)
	ON_BN_CLICKED(IDC_BUTTON_REGEX_ADD_ALL, &PPFileTabRegex::OnBnClickedButtonRegexAddAll)
	ON_BN_CLICKED(IDC_BUTTON_REGEX_CLEAR_LIST, &PPFileTabRegex::OnBnClickedButtonRegexClearList)
	ON_BN_CLICKED(IDC_BUTTON_REGEX_ADD_SELECTED, &PPFileTabRegex::OnBnClickedButtonRegexAddSelected)
	ON_BN_CLICKED(IDC_BUTTON_REGEX_REMOVE, &PPFileTabRegex::OnBnClickedButtonRegexRemove)
	ON_BN_CLICKED(IDC_BUTTON_REGEX_CLEAR_UNSELECTED, &PPFileTabRegex::OnBnClickedButtonRegexClearUnselected)
	ON_BN_CLICKED(IDC_BUTTON_REGEX_SELECT_ALL, &PPFileTabRegex::OnBnClickedButtonRegexSelectAll)
	ON_BN_CLICKED(IDC_BUTTON_REGEX_RUN_SELECTED, &PPFileTabRegex::OnBnClickedButtonRegexRunSelected)
	ON_BN_CLICKED(IDC_BUTTON_HISTORY, &PPFileTabRegex::OnBnClickedButtonHistory)
	ON_BN_CLICKED(IDC_BUTTON_REGEX_SAVE, &PPFileTabRegex::OnBnClickedButtonRegexSave)
END_MESSAGE_MAP()


// PPFileTabRegex message handlers
BOOL PPFileTabRegex::OnInitDialog()
{
	CPropertyPage::OnInitDialog();

	m_sRootFolder = theApp.m_RegistrySettings.m_sRegexRootFolder.c_str();
	m_sRegex = theApp.m_RegistrySettings.m_sRegexDefault.c_str();
	m_bRecurseSubfolders = theApp.m_RegistrySettings.m_iRegexRecurseSubfolders == 0 ? false : true;
	m_bClearSearchList = theApp.m_RegistrySettings.m_iRegexClearSearchList == 0 ? false : true;
	m_sWildcard = theApp.m_RegistrySettings.m_sRegexWildcard.c_str();

	m_btnFilesSelectAll.SetImage(IDB_SELECTALL, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesSelectAll.SetToolTip("Select all files");
	m_btnFilesAddSelected.SetImage(IDB_ADDALL, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesAddSelected.SetToolTip("Add selected files");
	m_btnFilesRemoveSelected.SetImage(IDB_DELETEALL, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesRemoveSelected.SetToolTip("Remove selected files");
	m_btnFilesRemoveAll.SetImage(IDB_DELETE, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesRemoveAll.SetToolTip("Remove all files");
	m_btnFilesDeleteUnselected.SetImage(IDB_DELETEUNSELECTED, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesDeleteUnselected.SetToolTip("Remove unselected files");
	m_btnFilesAddAll.SetImage(IDB_ADDSELECTED, RGB(0xcc, 0xcc, 0xcc));
	m_btnFilesAddAll.SetToolTip("Add all files");
	m_btnRunSelected.SetImage(IDB_RUNSELECTED, RGB(0xcc, 0xcc, 0xcc));
	m_btnRunSelected.SetToolTip("Run script on selected files");

	int iItem = m_cbSearchOption.AddString(_T("File names only"));
	m_cbSearchOption.SetItemData(iItem, eFilenames);
	iItem = m_cbSearchOption.AddString(_T("Folder names only"));
	m_cbSearchOption.SetItemData(iItem, eFoldernames);
	iItem = m_cbSearchOption.AddString(_T("File and folder names"));
	m_cbSearchOption.SetItemData(iItem, eFilesAndFolders);
	iItem = m_cbSearchOption.AddString(_T("Existing result set"));
	m_cbSearchOption.SetItemData(iItem, eExistingResults);
	

	m_iSearchOption = theApp.m_RegistrySettings.m_iRegexSearchOption;
	m_bRetrieveParentFolder = theApp.m_RegistrySettings.m_iRegexRetrieveParentFolder == 0 ? false : true;

	ReadRegexHistory();

	UpdateData(FALSE);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


void PPFileTabRegex::OnBnClickedButtonRunRegex()
{
	RunRegex();
}

void PPFileTabRegex::RunRegex()
{
	UpdateData(TRUE);
	std::string sRegex;
	sRegex = CW2A(m_sRegex.GetBuffer());
	boost::regex ex;	
	_StringVector vFiles;
	_StringVector vFolders;
	std::wstring sRoot;
	sRoot = m_sRootFolder;
	std::wstringstream ssOutput;
	ULONGLONG ulStart, ulEnd;

	// Boost regex will throw an exception if the regex syntax is invalid
	bool bSyntaxError = false;
	std::wstringstream ssMsg;
	try
	{
		ex = sRegex.c_str();
	}
	catch (const boost::regex_error& e)
	{
		bSyntaxError = true;
		ssMsg << _T("Caught boost::regex_error exception.  Code [") << e.code() << _T("] Message ");
		switch (e.code())
		{
		case boost::regex_constants::error_backref: ssMsg << _T("[error_backref: invalid back reference]"); break;
		case boost::regex_constants::error_badbrace: ssMsg << _T("[error_badbrace: invalid count in a { } expression]"); break;
		case boost::regex_constants::error_badrepeat: ssMsg << _T("[error_badrepeat: a repetition character (*, ?, +, or {) was not preceded by a valid regular expression]"); break;
		case boost::regex_constants::error_brace: ssMsg << _T("[error_brace: mismatched brace '{' or '}']"); break;
		case boost::regex_constants::error_brack: ssMsg << _T("[error_brack: mismatched bracket '[' or ']']"); break;
		case boost::regex_constants::error_collate: ssMsg << _T("[error_collate: invalid collating element name]"); break;
		case boost::regex_constants::error_complexity: ssMsg << _T("[error_complexity: the requested match is too complex]"); break;
		case boost::regex_constants::error_ctype: ssMsg << _T("[error_ctype: invalid character class name]"); break;
		case boost::regex_constants::error_escape: ssMsg << _T("[error_escape: invalid escape character or trailing escape]"); break;
		case boost::regex_constants::error_paren: ssMsg << _T("[error_paren: mismatched parentheses '(' or ')']"); break;
		case boost::regex_constants::error_range: ssMsg << _T("[error_range: invalid character range (e.g., [z-a])]"); break;
		case boost::regex_constants::error_space: ssMsg << _T("[error_space: insufficient memory to handle this regular expression]"); break;
		case boost::regex_constants::error_stack: ssMsg << _T("[error_stack: insufficient memory to evaluate a match]"); break;
		default: ssMsg << _T("[unknown error code]"); break;
		}
	}
	catch (...)
	{
		bSyntaxError = true;
		ssMsg << _T("[invalid regular expression]");
	}
	if (bSyntaxError)
	{
		ssMsg << std::endl << std::endl << _T("Please correct regular expression syntax and try again.");
		AfxMessageBox(ssMsg.str().c_str());
		return;
	}
	
#define TIMER_START ulStart = ::GetTickCount64();
#define TIMER_END ulEnd = ::GetTickCount64();
#define TIME_ELAPSED ((ulEnd - ulStart) / 1000.0)
		
	theApp.m_RegistrySettings.m_sRegexRootFolder = m_sRootFolder;
	theApp.m_RegistrySettings.m_sRegexDefault = m_sRegex;
	theApp.m_RegistrySettings.m_iRegexRecurseSubfolders = m_bRecurseSubfolders ? 1 : 0;
	theApp.m_RegistrySettings.m_iRegexClearSearchList = m_bClearSearchList ? 1 : 0;
	theApp.m_RegistrySettings.m_sRegexWildcard = m_sWildcard;
	theApp.m_RegistrySettings.m_iRegexSearchOption = m_iSearchOption;
	theApp.m_RegistrySettings.m_iRegexRetrieveParentFolder = m_bRetrieveParentFolder ? 1 : 0;

	m_lstFiles.ResetContent();
	UpdateData(FALSE);

	RegexSearchOptionEnum eSearchOption = (RegexSearchOptionEnum)m_cbSearchOption.GetItemData(m_iSearchOption);
	if (eSearchOption == eExistingResults)
	{
		ssOutput << _T("Regex searching existing results");
		vFiles = m_svFileList;
	}
	else
	{
		ssOutput << _T("Regex searching folder [") << (LPCTSTR)m_sRootFolder << _T("]");
	}
	theApp.AddToOutput(eStandardOutput, ssOutput.str().c_str());

	if (m_bClearSearchList || eSearchOption == eExistingResults)
	{
		OnBnClickedButtonRegexClearList();
	}
	
	if (eSearchOption == eFilenames || eSearchOption == eFilesAndFolders)
	{
		TIMER_START;
		GetFilesInDirectory(m_sWildcard, vFiles, sRoot, m_bRecurseSubfolders);
		TIMER_END;
		ssOutput.str(_T(""));
		ssOutput << _T("Found ") << vFiles.size() << _T(" file(s) using wildcard [") << (LPCTSTR)m_sWildcard << _T("] in ") << TIME_ELAPSED << _T(" seconds.");
		theApp.AddToOutput(eStandardOutput, ssOutput.str().c_str());
	}

	if (eSearchOption == eFoldernames || eSearchOption == eFilesAndFolders)
	{
		TIMER_START;
		GetFoldersInDirectory(vFolders, sRoot, m_bRecurseSubfolders);		
		TIMER_END;
		ssOutput.str(_T(""));
		ssOutput << _T("Found ") << vFolders.size() << _T(" folder(s) in ") << TIME_ELAPSED << _T(" seconds.");
		theApp.AddToOutput(eStandardOutput, ssOutput.str().c_str());
	}
	
	int iAddCount = 0;
	unsigned int uiFileCount = 0;
	unsigned int uiFolderCount = 0;
	TCHAR fname[_MAX_FNAME];
	TCHAR ext[_MAX_EXT];
	
	TIMER_START;
	for (const std::wstring &itFile : vFiles)
	{		
		std::wstring st;
		std::string sFileName;
		_wsplitpath_s(itFile.c_str(), NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);
		st = fname;
		st += ext;

		sFileName = CW2A(st.c_str());

		std::string::const_iterator start, end;
		start = sFileName.begin();
		end = sFileName.end();
		boost::match_results<std::string::const_iterator> what;
		boost::match_flag_type flags = boost::match_default;		
		if (boost::regex_search(start, end, what, ex, flags))		
		{
			uiFileCount++;			
			if (!m_bRetrieveParentFolder)
			{
				st = itFile.c_str();
			}
			else
			{
				st = _T("");
				_StringVector sv;
				boost::split(sv, itFile, boost::is_any_of("\\/"));
				for (int i = 0; i < sv.size() - 1; i++)
				{
					st += sv.at(i).c_str();
					st += _T("\\");
				}
			}
			
			// Check to see if this file is already in our list
			bool bAdd = true;
			for (_itString itExistingFile = m_svFileList.begin(); bAdd && itExistingFile != m_svFileList.end(); itExistingFile++)
			{
				if (!_tcsicmp(st.c_str(), (*itExistingFile).c_str()))
				{
					bAdd = false;					
				}
			}

			if (bAdd)
			{	
				m_svFileList.push_back(st);
				
				// Get the substring without our root				
				std::wstring wShort;
				wShort = st.substr(sRoot.size(), st.size() - sRoot.size());
				m_svShortFileList.push_back(wShort);
				iAddCount++;
			}			
		}
	}
	for (const std::wstring &itFolder : vFolders)
	{
		std::string sFolderName;

		_StringVector sv;
		boost::split(sv, itFolder, boost::is_any_of("\\/"));
		_ritString rit = sv.rbegin();
		if (rit == sv.rend())
		{
			continue;
		}
		sFolderName = CW2A(rit->c_str());

		std::string::const_iterator start, end;
		start = sFolderName.begin();
		end = sFolderName.end();
		boost::match_results<std::string::const_iterator> what;
		boost::match_flag_type flags = boost::match_default;
		if (boost::regex_search(start, end, what, ex, flags))
		{
			uiFolderCount++;			
			// Check to see if this folder is already in our list
			bool bAdd = true;
			for (_itString itExistingFile = m_svFileList.begin(); bAdd && itExistingFile != m_svFileList.end(); itExistingFile++)
			{
				if (!_tcsicmp(itFolder.c_str(), (*itExistingFile).c_str()))
				{
					bAdd = false;
				}
			}

			if (bAdd)
			{
				m_svFileList.push_back(itFolder);
				// Get the substring without our root
				std::wstring wShort;
				wShort = itFolder.substr(sRoot.size(), itFolder.size() - sRoot.size());
				m_svShortFileList.push_back(wShort);
				iAddCount++;
			}
		}
	}
	TIMER_END;

	AddToRegexHistory(uiFileCount, uiFolderCount);

	ssOutput.str(_T(""));
	ssOutput << _T("Regex matched ") << (uiFileCount + uiFolderCount) << _T(" entries in ") << TIME_ELAPSED << _T(" seconds.  Added ") << iAddCount << _T(" new entries to search results.");
	theApp.AddToOutput(eStandardOutput, ssOutput.str().c_str(), true);

	DisplaySearchList();
}


void PPFileTabRegex::DisplaySearchList()
{
	// Now populate our list control
	int iIndex = 0;
	for (const std::wstring &itFile : m_svShortFileList)
	{
		int iItem = m_lstFiles.AddString(itFile.c_str());
		m_lstFiles.SetItemData(iItem, iIndex);
		iIndex++;
	}
	UpdateData(FALSE);
}

void PPFileTabRegex::OnBnClickedButtonRegexAddAll()
{
	for( const std::wstring &s : m_svFileList)
	{
		theApp.AddFileFolder(s.c_str());
	}	
}


void PPFileTabRegex::OnBnClickedButtonRegexClearList()
{
	ResetContent();
}

void PPFileTabRegex::ResetContent()
{
	m_svFileList.erase(m_svFileList.begin(), m_svFileList.end());
	m_svShortFileList.erase(m_svShortFileList.begin(), m_svShortFileList.end());
	m_lstFiles.ResetContent();
	UpdateData(FALSE);
}


void PPFileTabRegex::OnBnClickedButtonRegexAddSelected()
{
	CArray<int, int> aListBoxSel;
	int nCount = m_lstFiles.GetSelCount();

	aListBoxSel.SetSize(nCount);	
	nCount = m_lstFiles.GetSelItems(nCount, aListBoxSel.GetData());
		
	for (int i = 0; i < nCount; i++)
	{
		int iIndex = m_lstFiles.GetItemData(aListBoxSel.GetAt(i));
		theApp.AddFileFolder(m_svFileList.at(iIndex).c_str());
	}

	// Dump the selection array.
	AFXDUMP(aListBoxSel);	
}


void PPFileTabRegex::OnBnClickedButtonRegexRemove()
{
	ChangeSelected(true);
}


void PPFileTabRegex::OnBnClickedButtonRegexClearUnselected()
{
	ChangeSelected(false);
}

void PPFileTabRegex::ChangeSelected(bool bRemoveSelected)
{
	CArray<int, int> aListBoxSel;
	int nCount = m_lstFiles.GetSelCount();

	aListBoxSel.SetSize(nCount);
	nCount = m_lstFiles.GetSelItems(nCount, aListBoxSel.GetData());

	_StringVector svFileList;
	_StringVector svShortFileList;

	int i = 0;
	for (const std::wstring &s : m_svFileList)
	{
		bool bDelete = bRemoveSelected ? false : true;
		for (int n = 0; n < nCount; n++)
		{
			int iIndex = m_lstFiles.GetItemData(aListBoxSel.GetAt(n));
			if (i == iIndex)
			{
				bDelete = bRemoveSelected ? true : false;
				break;
			}
		}
		if (!bDelete)
		{
			svFileList.push_back(s);
			svShortFileList.push_back(m_svShortFileList.at(i));
		}
		i++;
	}

	OnBnClickedButtonRegexClearList();
	m_svFileList = svFileList;
	m_svShortFileList = svShortFileList;

	DisplaySearchList();

	// Dump the selection array.
	AFXDUMP(aListBoxSel);
}

void PPFileTabRegex::OnBnClickedButtonRegexSelectAll()
{
	bool bSelect = (m_lstFiles.GetSelCount() == m_lstFiles.GetCount() ? false : true);
	int nCount = m_lstFiles.GetCount();
	SetRedraw(FALSE);
	int topIndex = m_lstFiles.GetTopIndex();
	for (int i = 0; i < nCount; i++)
	{
		m_lstFiles.SetSel(i, bSelect);
	}
	m_lstFiles.SetTopIndex(topIndex);
	SetRedraw(TRUE);
	m_lstFiles.RedrawWindow();
}

void PPFileTabRegex::OnBnClickedButtonRegexRunSelected()
{
	RunSelected();
}

void PPFileTabRegex::RunSelected()
{
	CArray<int, int> aListBoxSel;
	int nCount = m_lstFiles.GetSelCount();
	if (nCount <= 0)
	{
		theApp.AddToOutput(eStandardOutput, _T("No entries selected."), true);
		return;
	}

	aListBoxSel.SetSize(nCount);
	nCount = m_lstFiles.GetSelItems(nCount, aListBoxSel.GetData());
	_vLPFileFolderData vFileFolderData;

	for (int i = 0; i < nCount; i++)
	{
		std::wstring sPath;
		int iIndex = m_lstFiles.GetItemData(aListBoxSel.GetAt(i));
		sPath = m_svFileList.at(iIndex).c_str();
		LPFileFolderDataStruct pNewFD = new FileFolderDataStruct;
		
		pNewFD->sFullName = sPath;
		if (boost::filesystem::is_directory(sPath))
		{
			pNewFD->Type = FolderType;
		}
		else
		{
			pNewFD->Type = FileType;
		}
		vFileFolderData.push_back(pNewFD);
	}

	theApp.ExecuteScript(vFileFolderData);
	EraseDataVector<_vLPFileFolderData>(vFileFolderData);

	// Dump the selection array.
	AFXDUMP(aListBoxSel);
}

void PPFileTabRegex::ReadRegexHistory()
{
	// Open the file
	std::wstringstream ssHistoryFile;
	ssHistoryFile << theApp.m_sPythagorasFolder.c_str() << _T("\\") << _T("RegexHistory.dat");
	CMemBuffer mbData;
	LPBYTE pData = mbData.InitFromFile(ssHistoryFile.str().c_str());
	LPRegexHistoryHeader lpHeader = (LPRegexHistoryHeader)pData;
	if(pData)
	{
		unsigned int *pEntryOffset = (unsigned int *)(pData + sizeof(RegexHistoryHeader));
		for (unsigned int i = 0; i < lpHeader->uiNumEntries; i++)
		{
			LPRegexHistoryEntry lpEntry = (LPRegexHistoryEntry)(pData + *pEntryOffset);			
			wchar_t *pStringData = (wchar_t *)(&lpEntry->pData);

			LPRegexHistory lpNewEntry = new RegexHistory;
			lpNewEntry->sDate.append(pStringData,lpEntry->uiDateSize);
			pStringData += lpEntry->uiDateSize;
			lpNewEntry->sRootFolder.append(pStringData, lpEntry->uiRootFolderSize);
			pStringData += lpEntry->uiRootFolderSize;
			lpNewEntry->sWildcard.append(pStringData, lpEntry->uiWildCardSize);
			pStringData += lpEntry->uiWildCardSize;
			lpNewEntry->sRegEx.append(pStringData, lpEntry->uiRegExSize);

			lpNewEntry->uiSearchOption = lpEntry->uiSearchOption;
			lpNewEntry->uiFileMatchCount = lpEntry->uiFileMatchCount;
			lpNewEntry->uiFolderMatchCount = lpEntry->uiFolderMatchCount;
			
			m_RegexHistory.push_back(lpNewEntry);
			
			pEntryOffset++;
		}
		
	}
}

void PPFileTabRegex::AddToRegexHistory(unsigned int uiFileCount, unsigned int uiFolderCount)
{
	LPRegexHistory lpNewEntry = new RegexHistory;
	auto t = std::time(nullptr);
	auto tm = *std::localtime(&t);
	std::wstringstream ssDate;
	ssDate << std::put_time(&tm, _T("%m-%d-%Y %H:%M:%S"));

	lpNewEntry->sDate = ssDate.str().c_str();
	lpNewEntry->sRootFolder = m_sRootFolder;
	lpNewEntry->sWildcard = m_sWildcard;
	lpNewEntry->sRegEx = m_sRegex;
	lpNewEntry->uiSearchOption = m_iSearchOption;
	lpNewEntry->uiFileMatchCount = uiFileCount;
	lpNewEntry->uiFolderMatchCount = uiFolderCount;
	m_RegexHistory.push_back(lpNewEntry);
}

void PPFileTabRegex::WriteRegexHistory()
{
	if (m_RegexHistory.size() == 0)
	{
		return;
	}
	
	// Create the file
	std::wstringstream ssHistoryFile;
	ssHistoryFile << theApp.m_sPythagorasFolder.c_str() << _T("\\") << _T("RegexHistory.dat");
	CMemBuffer mbHeader((DWORD)sizeof(RegexHistoryHeader), MEMBUFFER_FLAG_ZEROMEMORY | MEMBUFFER_FLAG_MIN_SIZE_8192);
	CMemBuffer mbData((DWORD)0, MEMBUFFER_FLAG_ZEROMEMORY | MEMBUFFER_FLAG_MIN_SIZE_8192);
	LPRegexHistoryHeader lpHeader = (LPRegexHistoryHeader)mbHeader.GetBuffer();
	
	BYTE byEntryBlock[] = { 0x0, 0x0, 0x0, 0x0 };
	unsigned int *pEntryOffset = (unsigned int *)byEntryBlock;
	
	lpHeader->uiVersion = 0x1;
	lpHeader->uiNumEntries = m_RegexHistory.size();
	unsigned int iEntrySize = 0;

	(*pEntryOffset) = m_RegexHistory.size() * sizeof(unsigned int) + sizeof(RegexHistoryHeader);
	
	for (LPRegexHistory lpEntry : m_RegexHistory)
	{
		CMemBuffer mbEntry;
		iEntrySize = sizeof(RegexHistoryEntry) - sizeof(BYTE) + (lpEntry->sDate.size() + lpEntry->sRegEx.size() + lpEntry->sRootFolder.size() + lpEntry->sWildcard.size()) * sizeof(wchar_t);
		LPRegexHistoryEntry lpEntryHeader = (LPRegexHistoryEntry)mbEntry.GetBuffer(iEntrySize);
		LPBYTE pbyStringData = &lpEntryHeader->pData;
				
		const wchar_t *pStringData = lpEntry->sDate.data();		
		unsigned int iSize = lpEntry->sDate.size() * sizeof(wchar_t);
		lpEntryHeader->uiDateSize = lpEntry->sDate.size();
		memcpy(pbyStringData, (void *)lpEntry->sDate.data(), iSize);
		pbyStringData += iSize;

		iSize = lpEntry->sRootFolder.size() * sizeof(wchar_t);
		lpEntryHeader->uiRootFolderSize = lpEntry->sRootFolder.size();
		memcpy(pbyStringData, (void *)lpEntry->sRootFolder.data(), iSize);
		pbyStringData += iSize;

		iSize = lpEntry->sWildcard.size() * sizeof(wchar_t);
		lpEntryHeader->uiWildCardSize = lpEntry->sWildcard.size();
		memcpy(pbyStringData, (void *)lpEntry->sWildcard.data(), iSize);
		pbyStringData += iSize;

		iSize = lpEntry->sRegEx.size() * sizeof(wchar_t);
		lpEntryHeader->uiRegExSize = lpEntry->sRegEx.size();
		memcpy(pbyStringData, (void *)lpEntry->sRegEx.data(), iSize);

		lpEntryHeader->uiSearchOption = lpEntry->uiSearchOption;
		lpEntryHeader->uiFileMatchCount = lpEntry->uiFileMatchCount;
		lpEntryHeader->uiFolderMatchCount = lpEntry->uiFolderMatchCount;
				
		mbHeader.AppendBuffer(byEntryBlock, 4);
		(*pEntryOffset) += iEntrySize;
		mbData.AppendBuffer(&mbEntry);
	}

	mbHeader.AppendBuffer(&mbData);
	mbHeader.WriteToFile(ssHistoryFile.str().c_str());
}

void PPFileTabRegex::OnBnClickedButtonHistory()
{
	RegExHistoryDlg dlg;
	dlg.m_pRegexHistory = &m_RegexHistory;
	if (dlg.DoModal() == IDOK)
	{
		if (dlg.m_pSelectedEntry)
		{
			m_sRootFolder = dlg.m_pSelectedEntry->sRootFolder.c_str();
			m_sWildcard = dlg.m_pSelectedEntry->sWildcard.c_str();
			m_sRegex = dlg.m_pSelectedEntry->sRegEx.c_str();
			m_iSearchOption = dlg.m_pSelectedEntry->uiSearchOption;
			UpdateData(FALSE);
		}		
	}
}


BOOL PPFileTabRegex::PreTranslateMessage(MSG* pMsg)
{
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN)
	{
		CWnd *pFocus = GetFocus();
		if (pFocus == (CWnd*)&m_edRegexString)
		{
			RunRegex();
			return TRUE;
		}
		if (pFocus == (CWnd*)&m_edWildcard)
		{
			m_edRegexString.SetFocus();
			return TRUE;
		}
	}

	return CPropertyPage::PreTranslateMessage(pMsg);
}


void PPFileTabRegex::SaveResultList()
{

}

void PPFileTabRegex::OnBnClickedButtonRegexSave()
{
	// TODO: Add your control notification handler code here
}

void PPFileTabRegex::AddFile(LPCTSTR szFileName, LPCTSTR szRootFolder)
{
	// Check to see if this file is already in our list
	bool bAdd = true;
	for (_itString itExistingFile = m_svFileList.begin(); bAdd && itExistingFile != m_svFileList.end(); itExistingFile++)
	{
		if (!_tcsicmp(szFileName, (*itExistingFile).c_str()))
		{
			bAdd = false;
		}
	}

	if (bAdd)
	{
		m_svFileList.push_back(szFileName);

		if(szRootFolder)
		{
			// Get the substring without our root				
			std::wstring wShort;
			std::wstring sRoot(szRootFolder);
			std::wstring st(szFileName);
			wShort = st.substr(sRoot.size(), st.size() - sRoot.size());
			m_svShortFileList.push_back(wShort);
		}
		else
		{
			m_svShortFileList.push_back(szFileName);
		}		
	}	
}

void PPFileTabRegex::ClearResultList()
{
	m_lstFiles.ResetContent();
}