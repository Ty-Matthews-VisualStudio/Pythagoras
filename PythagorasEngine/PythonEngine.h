#pragma once
#ifndef __PYTHON_ENGINE_H_
#define __PYTHON_ENGINE_H_

#include <sstream>
#include <ostream>
#include <vector>
#include <string>
#include <MsHTML.h>
#include <boost/algorithm/string/split.hpp>
#include <boost/interprocess/managed_shared_memory.hpp>
#include <boost/interprocess/containers/vector.hpp>
#include <boost/interprocess/allocators/allocator.hpp>
#include "..\Common\PythagorasEngineMemory.h"

// For PyStrContainer:
#define RETURN_ENUM(str,e) if (!_stricmp(m_String, str)) {return e;}

namespace PythonEngine
{
	
	typedef std::vector< std::wstring > _vStrings;
	typedef _vStrings::iterator _itString;
	typedef std::vector<std::pair<std::wstring, std::wstring>> _vStringPairs;
	typedef _vStringPairs::iterator _itStringPair;
	
	typedef enum
	{
		eEMQueryVariables = 0,
		eEMExecute
	} eExecuteMode;
	
	typedef enum
	{
		eVIDUnknown = 0,
		eVIDComboBox,
		eVIDCheckBox,
		eVIDText,
		eVIDHidden
	} eVariableID;

	typedef enum
	{
		eHKIUnknown = 0,
		eHKITemplate,
		eHKIWidth,
		eHKIHeight,
		eHKIBGColor,
		eHKIReplaceStrings		
	} eHTMLKeyID;

	typedef enum
	{
		eVersionUnknown = 0,
		eVersion_1_0
	} eVersionID;

	typedef enum
	{
		eRIDUnknown = 0,
		eRIDOutput,
		eRIDError,
		eRIDReturn
	} eReturnID;

	typedef enum
	{
		eVTNull = 0,
		eVTString,
		eVTNumber
	} eValueType;

	typedef enum
	{
		eCAUnknown = 0,
		eCAGet,
		eCAPut
	} eClipboardAction;

	class PyStrContainer
	{
	public:
		char *m_String;
		std::wstring m_wString;
		PyStrContainer();
		PyStrContainer(const PyStrContainer &copy);
		PyStrContainer(PyObject* PyObj, const std::stringstream &ssMessage);
		PyStrContainer(PyObject* PyObj, const std::string &sMessage);
		PyStrContainer(PyObject* PyObj, const char *szMessage);
		virtual ~PyStrContainer();
		virtual void Clear();		
		void SetString(const char *s)
		{
			if (!s)
			{
				return;
			}
			Clear();
			m_String = _strdup(s);
			if (m_String)
			{
				m_wString = CA2W(m_String);
			}			
		}
		void SetString(const wchar_t *s)
		{
			if (!s)
			{
				return;
			}
			Clear();			
			m_wString = s;
			m_String = _strdup(CW2A(s));			
		}
		void SetString(std::string &s) { SetString(s.c_str()); };
		void SetString(std::wstring &s) { SetString(s.c_str()); };
		void SetString(PyObject* PyObj, const char *szMessage);
		void SetString(PyObject* PyObj, std::stringstream &ssMessage);
		void SetStringFromBytes(PyObject* PyObj, const char *szMessage);
		friend std::ostream& operator<<(std::ostream& os, const PyStrContainer& PyStr)
		{
			os << (PyStr.m_String ? PyStr.m_String : "");
			return os;
		}
		LPCSTR c_str()
		{
			return m_String;
		}
		const wchar_t *w_str()
		{
			return m_wString.c_str();
		}
		operator const char *()
		{
			return m_String;
		}
		operator const wchar_t *()
		{
			return m_wString.c_str();
		}
		operator eVariableID()
		{			
			RETURN_ENUM("ComboBox", eVIDComboBox);
			RETURN_ENUM("CheckBox", eVIDCheckBox);
			RETURN_ENUM("Text", eVIDText);
			return eVIDUnknown;
		}
		operator eHTMLKeyID()
		{
			RETURN_ENUM("Template", eHKITemplate);
			RETURN_ENUM("Width", eHKIWidth);
			RETURN_ENUM("Height", eHKIHeight);
			RETURN_ENUM("BGColor", eHKIBGColor);
			RETURN_ENUM("ReplaceStrings", eHKIReplaceStrings);
			return eHKIUnknown;
		}
		operator eVersionID()
		{
			RETURN_ENUM("1.0", eVersion_1_0);
			return eVersionUnknown;
		}
		operator eReturnID()
		{
			RETURN_ENUM("Output", eRIDOutput);
			RETURN_ENUM("Error", eRIDError);
			RETURN_ENUM("Return", eRIDReturn);
			return eRIDUnknown;
		}
		operator eClipboardAction()
		{			
			RETURN_ENUM("Get", eCAGet);
			RETURN_ENUM("Put", eCAPut);
			return eCAUnknown;
		}
		operator eModuleCallbackRegexAction()
		{
			RETURN_ENUM("Add", eMCRegexAdd);
			RETURN_ENUM("Replace", eMCRegexReplace);
			return eMCRegexUnknown;
		}

	private:
		void Init(PyObject* PyObj, const char *szMessage);
	};

	class BaseItem
	{
	public:
		std::wstring m_sName;
		std::wstring m_sDescription;
		bool m_bRegistry;
		std::wstring m_sRegistrySubKey;

		BaseItem(PyStrContainer &StrName, PyStrContainer &StrDescription, bool bRegistry) : m_bRegistry(bRegistry)
		{
			SetName(StrName);
			SetDescription(StrDescription);
		}
		void SetName(PyStrContainer &StrName)
		{
			m_sName = CA2W(StrName.m_String);
		}
		void SetDescription(PyStrContainer &StrDescription)
		{
			m_sDescription = CA2W(StrDescription.m_String);
		}
		void ReadRegistryString(std::wstring &sValue);
		void ReadRegistryInteger(int &iValue, int min, int max);
		void WriteRegistryString(std::wstring &sValue);
		void WriteRegistryInteger(int iValue);
		virtual void ReadRegistry(LPCTSTR szSubKey) = 0;
		virtual void WriteRegistry() = 0;
		virtual void SetFromDefault() = 0;
	};

	class SaveFileItem : public BaseItem
	{
	public:
		std::wstring m_sFilter;
		std::wstring m_sDefExt;
		std::wstring m_Value;
		std::wstring m_Default;

		SaveFileItem(PyStrContainer &StrName, PyStrContainer &StrDescription, bool bRegistry, PyObject *PyDict);
		virtual ~SaveFileItem();
		void ReadRegistry(LPCTSTR szSubKey);
		void WriteRegistry();
		void SetFromDefault()
		{
			m_Value = m_Default;
		};
	};
	typedef std::vector< SaveFileItem > _vSaveFiles;
	typedef _vSaveFiles::iterator _itSaveFile;
	
	class TextItem : public BaseItem
	{
	public:
		std::wstring m_Value;
		std::wstring m_Default;
		unsigned long m_lMaxLength;
		bool m_bPassword;
		bool m_bHidden;
		
		TextItem(PyStrContainer &StrName, PyStrContainer &StrDescription, bool bRegistry, PyObject *PyDict);
		virtual ~TextItem();
		void ReadRegistry(LPCTSTR szSubKey);
		void WriteRegistry();
		void SetFromDefault()
		{
			m_Value = m_Default;
		};
	};
	typedef std::pair<unsigned int, TextItem> _TextItemPair;
	typedef std::vector< _TextItemPair > _vTextItems;
	typedef _vTextItems::iterator _itTextItem;

	class ComboBoxEntry
	{
	public:
		std::wstring m_sKey;
		std::wstring m_sValue;
		eValueType m_ValueType;

		ComboBoxEntry() : m_sKey(L""), m_sValue(L""), m_ValueType(eVTNull) {};
		ComboBoxEntry(ComboBoxEntry const& copy) : m_sKey(copy.m_sKey), m_sValue(copy.m_sValue), m_ValueType(copy.m_ValueType) {};
		ComboBoxEntry(const wchar_t *sKey, const wchar_t *sValue, const eValueType eVT) : m_sKey(sKey), m_sValue(sValue), m_ValueType(eVT) {};
		ComboBoxEntry(const std::wstring &sKey, const std::wstring &sValue, const eValueType eVT) : m_sKey(sKey), m_sValue(sValue), m_ValueType(eVT) {};
		~ComboBoxEntry() {};
		void ReadRegistry(LPCTSTR szSubKey);
		void WriteRegistry();
	};

	//typedef std::pair<std::string, std::string> ComboBoxEntry;
	typedef std::vector<ComboBoxEntry> _vComboBoxEntries;
	typedef _vComboBoxEntries::iterator _itComboBoxEntry;
	
	class ComboBoxItem : public BaseItem
	{
	public:
		_vComboBoxEntries m_Entries;
		eValueType m_ValueType;
		std::wstring m_Value;
		std::wstring m_Default;		

		ComboBoxItem(PyStrContainer &StrName, PyStrContainer &StrDescription, bool bRegistry, PyObject *PyDict);
		virtual ~ComboBoxItem();
		void AddEntry(Py_ssize_t pos, PyObject *objKey, PyObject *objValue, const char *szText);
		void ReadRegistry(LPCTSTR szSubKey);
		void WriteRegistry();
		void SetFromDefault()
		{
			m_Value = m_Default;
		};
	};
	typedef std::pair<unsigned int, ComboBoxItem> _ComboBoxPair;
	typedef std::vector< _ComboBoxPair > _vComboBoxes;
	typedef _vComboBoxes::iterator _itComboBox;

	class CheckBoxItem : public BaseItem
	{
	public:
		bool m_Value;
		bool m_Default;

		CheckBoxItem(PyStrContainer &StrName, PyStrContainer &StrDescription, bool bRegistry, PyObject *PyDict);
		virtual ~CheckBoxItem();
		void ReadRegistry(LPCTSTR szSubKey);
		void WriteRegistry();
		void SetFromDefault() {
			m_Value = m_Default;
		}
	};
	typedef std::pair<unsigned int, CheckBoxItem > _CheckBoxPair;
	typedef std::vector< _CheckBoxPair > _vCheckBoxes;
	typedef _vCheckBoxes::iterator _itCheckBox;
		
	class FormVariables
	{
	public:
		std::wstring m_sHTMLTemplate;
		unsigned int m_uiWidth;
		unsigned int m_uiHeight;
		std::wstring m_sBGColor;
		_vStringPairs m_vReplaceStrings;
		_vComboBoxes m_vComboBoxes;
		_vCheckBoxes m_vCheckBoxes;
		_vTextItems m_vTextItems;
		// This is really a ExecuteParams pointer, but it's a huge circular reference mess trying to get forward declarations between structs and classes to work cleanly
		DWORD_PTR m_lpExecuteParams;

		FormVariables() : m_lpExecuteParams(NULL), m_sHTMLTemplate(L""), m_uiWidth(320), m_uiHeight(200), m_sBGColor(L"#ffffff") {};
		virtual ~FormVariables() {};
		void DoDataExchange(IHTMLElementCollection *pCollection);
		void SetValuesFromDefaults();
		CString BuildHTMLForm();		
		void AddToDictionary(PyObject *pDict);
		void ParseHTMLDict(PyObject *pDict);
		void ClearForm();
		void SetExecuteParams(DWORD_PTR lpExecuteParams) { m_lpExecuteParams = lpExecuteParams; };
	};

	typedef std::vector< _SharedCompileScript *> _vScriptFunctions;
	typedef _vScriptFunctions::iterator _itScriptFunctions;	
		
	class ExecuteParams
	{
	public:
		void_allocator* alloc_inst;
		eProcessMode ProcessMode;
		std::wstring sInspectScript;
		std::wstring sInjectScript;
		std::wstring sTempFolder;
		std::wstring sDefaultHTMLTemplate;
		std::wstring sLocalScriptsFolder;
		std::wstring sCommonScriptsFolder;
		FormVariables FV;
		_vStrings vFilenames;
		_vStrings vFoldernames;
		_vStrings vOpenFiles;
		_vScriptFunctions vScriptFunctions;
		_SharedStringArray* pSharedStrings;
		_vStrings vPythonStdOut;
		_vStrings vPythonStdErr;
		std::wstring sPathToScript;
		std::wstring sScriptRegistry;
		std::wstring sFunction;
		std::wstringstream ssMessage;
		bool m_bOpenFiles;
		bool bSuccess;
		double dExecutionTime;
		ExecuteParams() {};
		virtual ~ExecuteParams() {};
	};	
	typedef ExecuteParams *LPExecuteParams;

		
	class PyRefHandler
	{
	public:
		std::vector< PyObject * > m_vPyObjects;

		PyRefHandler() {}
		~PyRefHandler()
		{
			std::vector< PyObject * >::iterator itObject;
			for (itObject = m_vPyObjects.begin(); itObject != m_vPyObjects.end(); itObject++)
			{
				if ((*itObject))
				{
					Py_XDECREF((*itObject));
					(*itObject) = NULL;
				}
			}
			m_vPyObjects.erase(m_vPyObjects.begin(), m_vPyObjects.end());
		}
		PyObject *AddRef(PyObject *pRef)
		{
#if 0
			// Calling Py_XDECREF on these objects has the nasty side effect of crashing the program.
			// Attempted to sort out which Python objects need to the reference decremented, but
			// ultimately it doesn't really matter since the entire Python interpreter goes away
			// between each execution of the program.
			if (pRef)
			{
				m_vPyObjects.push_back(pRef);
			}
#endif
			return pRef;
		}
	};
	
	void ParseFormVariables_v1_0(PyObject *objDict, LPExecuteParams lpParams, bool bRegistryOnly);
	void ParseFormVariables(PyObject *objDict, LPExecuteParams lpParams, bool bRegistryOnly);
	void ParseReturnVariables_v1_0(PyObject *objDict, LPExecuteParams lpParams);
	void ParseReturnVariables(PyObject *objDict, LPExecuteParams lpParams);
	void ParseKeyValue(PyStrContainer &strKey, PyObject *objValue);	
	void FormatPythonError(std::stringstream &sError);
	void FormatPythonError(std::wstringstream &sError);
	bool ExecuteScript(LPExecuteParams lpParams, eExecuteMode eEM);
	bool ExtractSharedMemory(LPExecuteParams lpParams, std::wstring &sResult);
	bool ExtractSharedMemory(LPExecuteParams lpParams, std::string &sResult);
	bool SetSharedMemory(LPExecuteParams lpParams, std::wstring &sResult);
	bool SetSharedMemory(LPExecuteParams lpParams, std::string &sResult);
	void GetDictString(PyObject *pObj, const char *szLabel, PyStrContainer &str, const char *szMessage);
	bool CompileScripts(LPExecuteParams lpParams, void_allocator *alloc_inst);
	void GetAppDirectory(std::string &sDirectory);
	void GetAppDirectory(std::wstring &wDirectory);
	void GetDocumentsDirectory(std::string &sDirectory);
	void GetDocumentsDirectory(std::wstring &wDirectory);
	void ExtractPythonListAsString(std::vector<std::string> &sVector, PyObject *pList);
	void ExtractPythonListAsString(std::vector<std::wstring> &sVector, PyObject *pList);
	void ExtractPythonListAsInteger(std::vector<int> &iVector, PyObject *pList);
	void ExtractPythonListAsFloat(std::vector<double> &dVector, PyObject *pList);
	void ExtractPythonListAsBool(std::vector<bool> &bVector, PyObject *pList);
	void ReplaceLocaleSpecificStrings(std::string &sString, FormVariables *pFormVariables);
	void ReplaceLocaleSpecificStrings(std::wstring &wString, FormVariables* pFormVariables);
	void OpenFiles(LPExecuteParams lpParams);
	bool SetSharedStringArray(_StringVector &vStrings, std::wstring &sResult);
	bool SetSharedStringArray(_StringVector *pStrings, std::wstring &sResult);

	class PyDictParser
	{
	public:
		std::map< std::wstring, std::vector<bool> > m_BoolVector;
		typedef std::map< std::wstring, std::vector<bool> >::iterator _itBool;
		std::map< std::wstring, std::vector<int> > m_IntVector;
		typedef std::map< std::wstring, std::vector<int> >::iterator _itInt;
		std::map< std::wstring, std::vector<double> > m_FloatVector;
		typedef std::map< std::wstring, std::vector<double> >::iterator _itFloat;
		std::map< std::wstring, std::vector<std::wstring> > m_StringVector;
		typedef std::map< std::wstring, std::vector<std::wstring> >::iterator _itString;
		
		PyDictParser() {};
		virtual ~PyDictParser() {};

		void AddBool(const wchar_t *szKey);
		void AddInt(const wchar_t *szKey);
		void AddFloat(const wchar_t *szKey);
		void AddString(const wchar_t *szKey);
		bool ParseDictionary(PyObject *pDict, const char *szFunction);
		int GetSize(const wchar_t *szKey);
		bool GetBoolAt(const wchar_t *szKey, int Pos, bool &b);
		bool GetStringAt(const wchar_t *szKey, int Pos, std::wstring &s);
		bool GetIntAt(const wchar_t *szKey, int Pos, int &i);
		bool GetFloatAt(const wchar_t *szKey, int Pos, double &d);

		std::vector<std::wstring> *GetStringVector(const wchar_t *szKey);
	};

	PyMODINIT_FUNC PyInit_Module(void);

	class PythonSession
	{
	public:
		typedef struct
		{
			HWND hwndCallback;
			UINT MessageID;
			LPExecuteParams lpExecuteParams;
		} EngineStateStruct, *LPEngineStateStruct;

		EngineStateStruct m_EngineState;
		static LPExecuteParams m_pExecuteParams;

		PythonSession()
		{
			wchar_t *program = _T("PythagorasEngine");

			/* Add a built-in module, before Py_Initialize */
			PyImport_AppendInittab("pythagoras", PyInit_Module);

			/* Pass program name to the Python interpreter */
			Py_SetProgramName(program);

			Py_Initialize();
		}
		~PythonSession()
		{
			try
			{
				Py_Finalize();
			}
			catch (...)
			{
			}
		}

		static void SetCallbackString(LPCTSTR szText);
		static void SetCallbackString(const char *szText);
		static void GetBaseSubKey(LPExecuteParams lpExecuteParams, std::wstring &sRegistrySubKey);
		static void ReadRegistryString(LPExecuteParams lpExecuteParams, std::wstring &sName, std::wstring &sValue);
		static void WriteRegistryString(LPExecuteParams lpExecuteParams, std::wstring &sName, std::wstring &sValue);
		
		// C functions accessible within Python:
		static PyObject *Module_MessageBox(PyObject *self, PyObject *args);
		static PyObject *Module_Status(PyObject *self, PyObject *args);
		static PyObject *Module_SetStepRange(PyObject *self, PyObject *args);
		static PyObject *Module_Step(PyObject *self, PyObject *args);
		static PyObject *Module_Print(PyObject *self, PyObject *args);		
		static PyObject *Module_DisplayForm(PyObject *self, PyObject *args);
		static PyObject *Module_GetRegistryValues(PyObject *self, PyObject *args);
		static PyObject *Module_OpenFile(PyObject *self, PyObject *args);
		static PyObject *Module_SaveFile(PyObject *self, PyObject *args);
		static PyObject *Module_Clipboard(PyObject *self, PyObject *args);
		static PyObject *Module_BrowseFolder(PyObject *self, PyObject *args);
		static PyObject *Module_DatabaseCreateEntry(PyObject *self, PyObject *args);
		static PyObject *Module_EntryAddTag(PyObject *self, PyObject *args);
		static PyObject *Module_AddToRegexList(PyObject *self, PyObject *args);
		static PyObject *Module_AddToProcessList(PyObject *self, PyObject *args);
	};
	
	// Some useful functions that should exist in the Python API....
	int PyDict_SetItemWideChar(PyObject *p, const wchar_t *key, PyObject *val);
	PyObject* PyObject_GetAttrWideChar(PyObject *o, const wchar_t *attr_name);
	void PyErr_SetString(PyObject *type, const wchar_t *message);

	void DebugOut(const wchar_t *szText, bool bNewLine = false);	
}

#endif		// #ifndef __PYTHON_ENGINE_H_