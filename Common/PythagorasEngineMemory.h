#pragma once
#ifndef __PYTHON_ENGINE_MEMORY_H_
#define __PYTHON_ENGINE_MEMORY_H_

#include <boost/interprocess/managed_shared_memory.hpp>
#include <boost/container/scoped_allocator.hpp>
#include <boost/interprocess/allocators/allocator.hpp>
#include <boost/interprocess/containers/vector.hpp>
#include <boost/interprocess/containers/string.hpp>
#include <string>
#include <sstream>

using namespace boost::interprocess;

namespace PythonEngine
{
	typedef enum
	{
		ePMCompileScript = 0,
		ePMCallFunction		
	} eProcessMode;

	typedef enum
	{
		eMCMessageBox = 0,
		eMCStepIt,
		eMCSetStatus,
		eMCPrint,
		eMCStepRange,
		eMCAddToRegex
	} eModuleCallback;
	typedef enum
	{
		eMCRegexUnknown = 0,
		eMCRegexAdd,
		eMCRegexReplace		
	} eModuleCallbackRegexAction;

	extern const char *cSharedMemoryString;
	extern const wchar_t *cFilenameArrayString;
	extern const wchar_t *cFoldernameArrayString;
	extern const wchar_t *cExecuteOptionsString;
	extern const wchar_t *cCompileScriptString;
	extern const wchar_t *cScriptCallbackString;
	extern const wchar_t *cStringArrayString;
	
	//Typedefs of allocators and containers
#ifdef _UNICODE
	typedef boost::interprocess::wmanaged_shared_memory::segment_manager                       segment_manager_t;
#else
	typedef boost::interprocess::managed_shared_memory::segment_manager                       segment_manager_t;
#endif
	typedef boost::container::scoped_allocator_adaptor<allocator<void, segment_manager_t> >
		void_allocator;
	
	typedef void_allocator::rebind<int>::other                           int_allocator;
	typedef boost::interprocess::vector<int, int_allocator>              int_vector;

	typedef void_allocator::rebind<char>::other                          char_allocator;	
	typedef boost::interprocess::basic_string<char, std::char_traits<char>, char_allocator>   char_string;
	typedef boost::interprocess::basic_string<char, std::char_traits<char>, char_allocator>::iterator _char_string_iterator;

	typedef void_allocator::rebind<wchar_t>::other								wchar_allocator;
	typedef boost::interprocess::basic_string<wchar_t, std::char_traits<wchar_t>, wchar_allocator>   wchar_string;
	typedef boost::interprocess::basic_string<wchar_t, std::char_traits<wchar_t>, wchar_allocator>::iterator _wchar_string_iterator;
		
	class _SharedString
	{
	public:
#ifdef _UNICODE
		wchar_string m_String;
#else
		char_string	m_String;
#endif

		//Since void_allocator is convertible to any other allocator<T>, we can simplify
		//the initialization taking just one allocator for all inner containers.
		typedef void_allocator allocator_type;

		_SharedString(_SharedString const& other, const allocator_type &void_alloc)
			: m_String(other.m_String, void_alloc)
		{}
		_SharedString(const allocator_type &void_alloc)
			: m_String(void_alloc)
		{}
	};

	typedef void_allocator::rebind<_SharedString>::other    _SharedStringAllocator;
	typedef boost::interprocess::vector<_SharedString, _SharedStringAllocator>   _SharedStringVector;
	typedef boost::interprocess::vector<_SharedString, _SharedStringAllocator>::iterator _SharedStringIterator;

	class _SharedCompileScript
	{
	public:
		int m_Index;
#ifdef _UNICODE
		wchar_string m_PathToFile;
		wchar_string m_ErrorMessage;
#else
		char_string	m_PathToFile;
		char_string m_ErrorMessage;
#endif
		_SharedStringVector m_Functions;
		_SharedStringVector m_Descriptions;
		_SharedStringVector m_DisplayNames;

		//Since void_allocator is convertible to any other allocator<T>, we can simplify
		//the initialization taking just one allocator for all inner containers.
		typedef void_allocator allocator_type;

		_SharedCompileScript(_SharedCompileScript const& other, const allocator_type &void_alloc)
			: m_Index(other.m_Index), m_PathToFile(other.m_PathToFile, void_alloc), m_ErrorMessage(other.m_ErrorMessage, void_alloc), m_Functions(other.m_Functions,void_alloc), m_Descriptions(other.m_Descriptions, void_alloc), m_DisplayNames(other.m_DisplayNames, void_alloc)
		{}
		_SharedCompileScript(const allocator_type &void_alloc)
			: m_Index(0), m_PathToFile(void_alloc), m_ErrorMessage(void_alloc), m_Functions(void_alloc), m_Descriptions(void_alloc), m_DisplayNames(void_alloc)
		{}
	};

	typedef void_allocator::rebind<_SharedCompileScript>::other _SharedCompileScriptAllocator;
	typedef boost::interprocess::vector<_SharedCompileScript, _SharedCompileScriptAllocator>   _SharedCompileScriptVector;
	typedef boost::interprocess::vector<_SharedCompileScript, _SharedCompileScriptAllocator>::iterator _SharedCompileScriptIterator;

	class _SharedStringArray
	{
	public:
		_SharedStringVector m_Strings;

		//Since void_allocator is convertible to any other allocator<T>, we can simplify
		//the initialization taking just one allocator for all inner containers.
		typedef void_allocator allocator_type;

		_SharedStringArray(_SharedStringArray const& other, const allocator_type &void_alloc)
			: m_Strings(other.m_Strings, void_alloc)
		{}
		_SharedStringArray(const allocator_type &void_alloc)
			: m_Strings(void_alloc)
		{}
	};

	class _SharedFileArray
	{
	public:
		_SharedStringVector m_Filenames;
		
		//Since void_allocator is convertible to any other allocator<T>, we can simplify
		//the initialization taking just one allocator for all inner containers.
		typedef void_allocator allocator_type;

		_SharedFileArray(_SharedFileArray const& other, const allocator_type &void_alloc)
			: m_Filenames(other.m_Filenames, void_alloc)
		{}
		_SharedFileArray(const allocator_type &void_alloc)
			: m_Filenames(void_alloc)
		{}
	};

	class _SharedFolderArray
	{
	public:
		_SharedStringVector m_Foldernames;

		//Since void_allocator is convertible to any other allocator<T>, we can simplify
		//the initialization taking just one allocator for all inner containers.
		typedef void_allocator allocator_type;

		_SharedFolderArray(_SharedFolderArray const& other, const allocator_type &void_alloc)
			: m_Foldernames(other.m_Foldernames, void_alloc)
		{}
		_SharedFolderArray(const allocator_type &void_alloc)
			: m_Foldernames(void_alloc)
		{}
	};

	class _SharedExecuteOptions
	{
	public:
		eProcessMode m_ProcessMode;
		int m_OpenFiles;
		double m_ExecutionTime;
		unsigned long m_lHWND;
		unsigned long m_lMessageID;

#ifdef _UNICODE
		wchar_string m_InspectFile;
		wchar_string m_InjectFile;
		wchar_string m_TempFolder;
		wchar_string m_DefaultHTMLTemplate;
		wchar_string m_LocalScriptsFolder;
		wchar_string m_CommonScriptsFolder;
		wchar_string m_Script;
		wchar_string m_ScriptRegistry;
		wchar_string m_Function;
#else
		char_string m_InspectFile;
		char_string m_InjectFile;
		char_string m_TempFolder;
		char_string m_DefaultHTMLTemplate;
		char_string m_LocalScriptsFolder;
		char_string m_CommonScriptsFolder;
		char_string	m_Script;
		char_string m_ScriptRegistry;
		char_string	m_Function;
#endif
		
		_SharedStringVector m_OutputVector;
		_SharedStringVector m_ErrorVector;

		//Since void_allocator is convertible to any other allocator<T>, we can simplify
		//the initialization taking just one allocator for all inner containers.
		typedef void_allocator allocator_type;

		_SharedExecuteOptions(_SharedExecuteOptions const& other, const allocator_type &void_alloc)
			:
			m_lHWND(other.m_lHWND), 
			m_lMessageID(other.m_lMessageID), 
			m_ProcessMode(other.m_ProcessMode),
			m_InspectFile(other.m_InspectFile),
			m_InjectFile(other.m_InjectFile),
			m_TempFolder(other.m_TempFolder),
			m_DefaultHTMLTemplate(other.m_DefaultHTMLTemplate),
			m_LocalScriptsFolder(other.m_LocalScriptsFolder),
			m_CommonScriptsFolder(other.m_CommonScriptsFolder),
			m_OpenFiles(other.m_OpenFiles),
			m_Script(other.m_Script, void_alloc),
			m_Function(other.m_Script, void_alloc),
			m_ScriptRegistry(other.m_ScriptRegistry, void_alloc),
			m_OutputVector(other.m_OutputVector, void_alloc),
			m_ErrorVector(other.m_ErrorVector,void_alloc)
		{}
		_SharedExecuteOptions(const allocator_type &void_alloc)
			:
			m_lHWND(0),
			m_lMessageID(0),
			m_ProcessMode(ePMCompileScript),
			m_OpenFiles(0),
			m_InspectFile(void_alloc),
			m_InjectFile(void_alloc),
			m_TempFolder(void_alloc),
			m_DefaultHTMLTemplate(void_alloc),
			m_LocalScriptsFolder(void_alloc),
			m_CommonScriptsFolder(void_alloc),
			m_Script(void_alloc),
			m_ScriptRegistry(void_alloc),
			m_Function(void_alloc),
			m_OutputVector(void_alloc),
			m_ErrorVector(void_alloc)
		{}
	};
	typedef void_allocator::rebind<_SharedExecuteOptions>::other _SharedExecuteOptionsAllocator;

	class _SharedScriptCallback
	{
	public:
#ifdef _UNICODE
		wchar_string m_String;
#else
		char_string m_String;
#endif
		//Since void_allocator is convertible to any other allocator<T>, we can simplify
		//the initialization taking just one allocator for all inner containers.
		typedef void_allocator allocator_type;

		_SharedScriptCallback(_SharedScriptCallback const& other, const allocator_type &void_alloc)
			: m_String(other.m_String)
		{}
		_SharedScriptCallback(const allocator_type &void_alloc)
			: m_String(void_alloc)
		{}
	};
	typedef void_allocator::rebind<_SharedScriptCallback>::other _SharedScriptCallbackAllocator;
		
	void ConvertString(char_string &cs, std::string &s);
	void ConvertString(wchar_string &cs, std::wstring &s);
}

#endif		// #ifndef __PYTHON_ENGINE_MEMORY_H_