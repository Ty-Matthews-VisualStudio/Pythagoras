#pragma once

#ifndef __STRUCTS_H_
#define __STRUCTS_H_

#include <string>
#include <sstream>
#include <vector>

typedef enum
{
	eFilenames = 0,
	eFoldernames,
	eFilesAndFolders,
	eExistingResults
} RegexSearchOptionEnum;

typedef enum
{
	eDescription = 0,
	eStandardOutput,
	eStandardError
} OutputTypeEnum;

typedef enum
{
	FileType = 0,
	FolderType
} eFileFolderType;

typedef struct
{
	eFileFolderType Type;
	std::wstring sFullName;
	std::wstring sShortName;
	std::wstring sDisplayName;
	std::wstring sDescription;	
	BOOL bSelected;
	BYTE color[3];
} FileFolderDataStruct, *LPFileFolderDataStruct;

typedef enum
{
	ScriptTypeLocal = 0,
	ScriptTypeCommon
} eScriptType;

typedef struct
{
	std::wstring sFunctionName;
	std::wstring sDescription;
	std::wstring sDisplayName;
} FunctionStruct, *LPFunctionStruct;

typedef std::vector< LPFunctionStruct > _vFunction;
typedef _vFunction::iterator _itFunction;

typedef struct
{
	std::wstring sTruePathToScript;
	std::wstring sPathToScript;
	std::wstring sRelativePath;
	_vFunction vFunctions;
	eScriptType ScriptType;
	bool bHidden;
	bool bErrors;
	std::wstringstream ssErrorMessage;
} ScriptFunctionStruct, *LPScriptFunctionStruct;

typedef std::vector< LPScriptFunctionStruct > _vScriptFunction;
typedef _vScriptFunction::iterator _itScriptFunction;

typedef struct
{
	HTREEITEM hParentNode;
	bool bFolder;
	unsigned int uiLocalScriptIndex;
	LPScriptFunctionStruct lpScriptFunction;
	LPScriptFunctionStruct lpChildScriptFunction;
} ScriptViewNodeStruct, *LPScriptViewNodeStruct;

typedef std::vector< LPScriptViewNodeStruct > _vLPScriptViewNode;

typedef std::vector< LPFileFolderDataStruct > _vLPFileFolderData;
typedef _vLPFileFolderData::iterator _itLPFileFolderData;

typedef std::vector< FileFolderDataStruct > _vFileFolderData;
typedef _vFileFolderData::iterator _itFileFolderData;

typedef std::vector< std::wstring > _StringVector;
typedef std::vector< std::wstring >::iterator _itString;
typedef std::vector< std::wstring >::reverse_iterator _ritString;

class CMainFrame;
typedef struct
{
	HANDLE hMutex;
	_StringVector vFileData;
	_StringVector vFolderData;	
	bool bDeleteMemory;
	bool bClearOutput;
	CMainFrame *pMainFrame;
	std::wstring sFunctionName;
	std::wstring sPathToScript;
	std::wstring sRelativePath;
	std::wstring sInjectScript;
	std::wstring sDefaultHTMLTemplate;
	std::wstring sLocalScriptsFolder;
	std::wstring sCommonScriptsFolder;
	std::wstring sTempFolder;
	std::wstring sPythonEnginePath;
} ExecuteScriptThreadParams, *LPExecuteScriptThreadParams;

typedef struct
{
	std::wstring sDate;
	std::wstring sRootFolder;
	std::wstring sWildcard;
	std::wstring sRegEx;
	unsigned int uiSearchOption;
	unsigned int uiFileMatchCount;
	unsigned int uiFolderMatchCount;
} RegexHistory, *LPRegexHistory;

typedef std::vector<LPRegexHistory> _vRegexHistory;

typedef struct
{
	unsigned int uiVersion;
	unsigned int uiNumEntries;
} RegexHistoryHeader, *LPRegexHistoryHeader;

typedef struct
{
	unsigned int uiDateSize;
	unsigned int uiRootFolderSize;	
	unsigned int uiWildCardSize;
	unsigned int uiRegExSize;
	unsigned int uiSearchOption;
	unsigned int uiFileMatchCount;
	unsigned int uiFolderMatchCount;
	BYTE pData;
} RegexHistoryEntry, *LPRegexHistoryEntry;


#endif // #ifndef __STRUCTS_H_