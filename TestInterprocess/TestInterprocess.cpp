// TestInterprocess.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

#ifdef _DEBUG
#undef _DEBUG
#include "C:\\Users\\a4lg8zz\\Python\\include\\Python.h"
#define _DEBUG
#else
#include "C:\\Users\\a4lg8zz\\Python\\include\\Python.h"
#endif

#include <boost/interprocess/managed_shared_memory.hpp>
#include <boost/container/scoped_allocator.hpp>
#include <boost/interprocess/allocators/allocator.hpp>
#include <boost/interprocess/containers/vector.hpp>
#include <boost/interprocess/containers/string.hpp>
#include <string>
#include <stdlib.h>
#include <iostream>

typedef boost::interprocess::wmanaged_shared_memory::segment_manager segment_manager_t;
typedef boost::container::scoped_allocator_adaptor<boost::interprocess::allocator<void, segment_manager_t> > void_allocator;
typedef void_allocator::rebind<wchar_t>::other wchar_allocator;
typedef boost::interprocess::basic_string<wchar_t, std::char_traits<wchar_t>, wchar_allocator> wchar_string;
typedef boost::interprocess::basic_string<wchar_t, std::char_traits<wchar_t>, wchar_allocator>::iterator _wchar_string_iterator;

class _SharedString
{
public:
	wchar_string m_String;

	//Since void_allocator is convertible to any other allocator<T>, we can simplify
	//the initialization taking just one allocator for all inner containers.
	typedef void_allocator allocator_type;

	_SharedString(_SharedString const& other, const allocator_type &void_alloc) : m_String(other.m_String, void_alloc) {}
	_SharedString(const allocator_type &void_alloc) : m_String(void_alloc) {}
};

typedef void_allocator::rebind<_SharedString>::other _SharedStringAllocator;
typedef boost::interprocess::vector<_SharedString, _SharedStringAllocator>   _SharedStringVector;
typedef boost::interprocess::vector<_SharedString, _SharedStringAllocator>::iterator _SharedStringIterator;

const char *cSharedMemory = "SharedMem";
const wchar_t *cSharedStrings = _T("SharedStrings");

void Owner(int argc, char **argv)
{
	using namespace boost::interprocess;
	class shm_remove
	{
	public:
		shm_remove() { shared_memory_object::remove(cSharedMemory); }
		~shm_remove() { shared_memory_object::remove(cSharedMemory); }
	};
	shm_remove remove;
	wmanaged_shared_memory segment(create_only, cSharedMemory, 65536);

	// An allocator convertible to any allocator<T, segment_manager_t> type
	void_allocator alloc_inst(segment.get_segment_manager());

	_SharedStringVector *StringVector = segment.construct<_SharedStringVector>(cSharedStrings)(alloc_inst);
	
	_SharedString Record(alloc_inst);	
	Record.m_String = _T("This is a test of the emergency broadcast system");
	StringVector->push_back(Record);

	Record.m_String = _T("C:\\Users\\A4LG8ZZ\\Documents\\Code\\VS2015\\Pythagoras\\x64");
	StringVector->push_back(Record);
	
	// Execute this program again, except with a parameter.  This will trigger the Child() code
	std::string sCmd(argv[0]);
	sCmd += " child";
	system(sCmd.c_str());

	if (segment.find<_SharedStringVector>(cSharedStrings).first)
	{
		segment.destroy<_SharedStringVector>(cSharedStrings);
	}
}

void Child(int argc, char **argv)
{
	using namespace boost::interprocess;
	char drive[_MAX_DRIVE];
	char dir[_MAX_DIR];
	std::string sFolder;
	_SharedStringIterator it;
	
	PyObject* objPath = NULL;
	PyObject* objFolder = NULL;	
	PyObject* pName = NULL;
	PyObject* pModule = NULL;
	PyObject* pFunc = NULL;
	PyObject* pList = NULL;
	PyObject* pValue = NULL;
	PyObject* pArgs = NULL;
	unsigned int iArg = 0;
	
	try
	{
		wmanaged_shared_memory segment(open_only, cSharedMemory);

		_SharedStringVector *StringVector = segment.find<_SharedStringVector>(cSharedStrings).first;

		_splitpath_s(argv[0], drive, _MAX_DRIVE, dir, _MAX_DIR, NULL, 0, NULL, 0);
		sFolder = drive;
		sFolder += dir;

		Py_Initialize();
		objPath = PySys_GetObject((char*)"path");
		objFolder = PyUnicode_FromString(sFolder.c_str());
		PyList_Append(objPath, objFolder);

		pName = PyUnicode_FromString("Test");
		pModule = PyImport_Import(pName);
		pFunc = PyObject_GetAttrString(pModule, "MyFunc");

		pList = PyList_New(StringVector->size());
		for (iArg = 0, it = StringVector->begin(); it != StringVector->end(); it++, iArg++)
		{
			std::wstring ws((*it).m_String.begin(), (*it).m_String.end());
			std::wcout << ws.c_str() << "\n";
			pValue = PyUnicode_FromWideChar(ws.c_str(), ws.size());
			PyList_SetItem(pList, iArg, pValue);
		}

		pArgs = PyTuple_New(1);
		PyTuple_SetItem(pArgs, 0, pList);

		pValue = PyObject_CallObject(pFunc, pArgs);
	}
	catch (interprocess_exception &ex)
	{
		std::cout << "Caught interprocess_exception.  Code=" << ex.get_error_code() << ": " << ex.what() << "\n";
	}
	catch (...)
	{
	}
	
}

int main(int argc, char **argv)
{
	if (argc == 1)
	{
		Owner(argc, argv);
	}
	else
	{
		Child(argc, argv);
	}	
	
    return 0;
}

