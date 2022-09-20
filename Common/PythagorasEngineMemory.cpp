#include "stdafx.h"
#include <sstream>
#include <string>
#include "PythagorasEngineMemory.h"

namespace PythonEngine
{
	const char *cSharedMemoryString = "Pythagoras Shared Memory";
	const wchar_t *cFilenameArrayString = _T("Filename Array");
	const wchar_t *cFoldernameArrayString = _T("Foldername Array");
	const wchar_t *cExecuteOptionsString = _T("Execute Options");
	const wchar_t *cCompileScriptString = _T("Compile Script");
	const wchar_t *cScriptCallbackString = _T("Script Callback");
	const wchar_t *cStringArrayString = _T("String Array");

	void ConvertString(char_string &cs, std::string &s)
	{
		std::string ls(cs.begin(), cs.end());
		s = ls;
	}

	void ConvertString(wchar_string &cs, std::wstring &s)
	{		
		std::wstring ws(cs.begin(), cs.end());
		s = ws;
	}
}