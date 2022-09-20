
#include "stdafx.h"
#include "PythonEngine.h"
#include "..\Common\PythagorasEngineMemory.h"
#include "PythagorasEngineDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace PythonEngine
{		
	PyStrContainer::PyStrContainer() : m_String(NULL)
	{
	}
	PyStrContainer::PyStrContainer(const PyStrContainer &copy)
	{		
		m_String = _strdup(copy.m_String);
	}
	PyStrContainer::PyStrContainer(PyObject* PyObj, const std::stringstream &ssMessage) : m_String(NULL)
	{
		Init(PyObj, ssMessage.str().c_str());
	}
	PyStrContainer::PyStrContainer(PyObject* PyObj, const std::string &sMessage) : m_String(NULL)
	{
		Init(PyObj, sMessage.c_str());
	}
	PyStrContainer::PyStrContainer(PyObject* PyObj, const char *szMessage) : m_String(NULL)
	{
		Init(PyObj, szMessage);
	}
	void PyStrContainer::SetString(PyObject* PyObj, const char *szMessage)
	{		
		Init(PyObj, szMessage);
	}
	void PyStrContainer::SetString(PyObject* PyObj, std::stringstream &ssMessage)
	{
		Init(PyObj, ssMessage.str().c_str());
	}
	void PyStrContainer::SetStringFromBytes(PyObject* PyObj, const char *szMessage)
	{
		Clear();
		if (PyObj)
		{			
			m_String = _strdup(PyBytes_AS_STRING(PyObj));
			if (!m_String)
			{
				throw std::exception(szMessage);
			}
			else
			{
				m_wString = CA2W(m_String);
			}
		}		
	}

	void PyStrContainer::Init(PyObject* PyObj, const char *szMessage)
	{
		Clear();
		if (PyObj)
		{
			if (PyUnicode_Check(PyObj))
			{
				//PyObject* PyStr = PyUnicode_AsEncodedString(PyObj, "utf-8", "Error ~");
				PyObject* PyStr = PyUnicode_AsEncodedString(PyObj, "utf-8", "ignore");
				if (PyStr)
				{
					m_String = _strdup(PyBytes_AS_STRING(PyStr));
				}
			}
			if (!m_String)
			{
				throw std::exception(szMessage);
			}
			else
			{
				m_wString = CA2W(m_String);
			}
		}		
	}
	PyStrContainer::~PyStrContainer()
	{		
		Clear();
	}
	
	void PyStrContainer::Clear()
	{
		if (m_String)
		{
			free(m_String);
			m_String = NULL;
		}
		m_wString.clear();
	}

	void GetDictString(PyObject *pObj, const char *szLabel, PyStrContainer &str, const char *szMessage)
	{		
		PyObject *objStr = PyDict_GetItemString(pObj, szLabel);
		str.SetString(objStr, szMessage);
	}

	void BaseItem::ReadRegistryString(std::wstring &sDefault)
	{
		if (m_bRegistry && m_sRegistrySubKey.size() > 0)
		{
			CRegistryHelper rhHelper;
			std::wstring sValue;
			rhHelper.SetMainKey(HKEY_CURRENT_USER);
			rhHelper.SetBaseSubKey(m_sRegistrySubKey.c_str());
			rhHelper.AddItem(&sValue, sDefault.c_str(), m_sName.c_str(), _T(""));
			rhHelper.ReadRegistry();
			rhHelper.WriteRegistry();
			sDefault = sValue;
		}
	}

	void BaseItem::ReadRegistryInteger(int &iDefault, int min, int max)
	{
		// Ints (and all other BaseItem values) are written as strings, to protect against the user changing the datatype 
		// on us in the future
		std::wstringstream ssValue;
		std::wstring sDefault;
		ssValue << iDefault;
		sDefault = ssValue.str().c_str();
		ReadRegistryString(sDefault);
		iDefault = _ttoi(sDefault.c_str());
		iDefault = min(iDefault, max);
		iDefault = max(iDefault, min);
	}

	void BaseItem::WriteRegistryString(std::wstring &sValue)
	{
		if (m_bRegistry && m_sRegistrySubKey.size() > 0)
		{
			CRegistryHelper rhHelper;
			rhHelper.SetMainKey(HKEY_CURRENT_USER);
			rhHelper.SetBaseSubKey(m_sRegistrySubKey.c_str());
			rhHelper.AddItem(&sValue, _T(""), m_sName.c_str(), _T(""));
			rhHelper.WriteRegistry();
		}
	}

	void BaseItem::WriteRegistryInteger(int iValue)
	{
		// Ints (and all other BaseItem values) are written as strings, to protect against the user changing the datatype 
		// on us in the future
		std::wstringstream ssValue;
		std::wstring sValue;
		ssValue << iValue;
		sValue = ssValue.str().c_str();
		WriteRegistryString(sValue);
	}

	SaveFileItem::SaveFileItem(PyStrContainer &StrName, PyStrContainer &StrDescription, bool bRegistry, PyObject *PyDict) : BaseItem(StrName, StrDescription, bRegistry)
	{
		PyStrContainer StrFilter, StrExtension;
		PyObject *objFilter = PyDict_GetItemString(PyDict, "Filter");
		std::stringstream ss;
		ss << "Error with SaveFile Variable '" << StrName << "': ";
		
		if (objFilter)
		{
			ss << "Filter must be string";
			StrFilter.SetString(objFilter, ss.str().c_str());
		}
		else
		{
			StrFilter.SetString("All Files (*.*)|*.*|");
		}
		m_sFilter = CA2W(StrFilter);

		PyObject *objExtension = PyDict_GetItemString(PyDict, "DefaultExtension");
		ss.str("");
		ss << "Error with SaveFile Variable '" << StrName << "': ";
		if (objExtension)
		{
			ss << "DefaultExtension must be string";
			StrExtension.SetString(objExtension, ss.str().c_str());
		}
		else
		{
			StrExtension.SetString("xlsx");
		}
		m_sDefExt = CA2W(StrExtension);
	}
	
	SaveFileItem::~SaveFileItem()
	{

	}

	void SaveFileItem::ReadRegistry(LPCTSTR szSubKey)
	{
		m_sRegistrySubKey = szSubKey;
		ReadRegistryString(m_Default);		
	}

	void SaveFileItem::WriteRegistry()
	{	
		WriteRegistryString(m_Value);
	}

	TextItem::TextItem(PyStrContainer &StrName, PyStrContainer &StrDescription, bool bRegistry, PyObject *PyDict) : BaseItem(StrName, StrDescription, bRegistry)
	{	
		m_lMaxLength = -1;
		m_bPassword = false;
		m_bHidden = false;
		std::stringstream ss;
		ss << "Error with Text Variable '" << StrName << "': ";
		PyRefHandler RefHandler;
		
		PyObject *objDefault = PyDict_GetItemString(PyDict, "Default");
		PyObject *objExc = NULL;
		std::wstringstream ssDefault;
		if (objDefault)
		{
			if (!PyNumber_Check(objDefault) && !PyUnicode_Check(objDefault))
			{
				ss << "Default must be either a string or a number";
				throw std::exception(ss.str().c_str());
			}
			if (PyNumber_Check(objDefault))
			{
				Py_ssize_t stDefault = PyNumber_AsSsize_t(objDefault, objExc);
				if (stDefault == PY_SSIZE_T_MIN || stDefault == PY_SSIZE_T_MAX)
				{
					ss << "Default value could not be converted to integer";
					throw std::exception(ss.str().c_str());
				}
				if (stDefault == -1 && PyErr_Occurred())
				{
					if (PyErr_Occurred())
					{
						PyErr_Clear();
					}
					// Not an integer, try as a float
					PyObject *pFloat = RefHandler.AddRef(PyNumber_Float(objDefault));
					double d = PyFloat_AS_DOUBLE(pFloat);
					ssDefault << d;
				}
				else
				{
					ssDefault << stDefault;
				}
				m_Default = ssDefault.str();
			}
			if (PyUnicode_Check(objDefault))
			{
				ss << "Default value could not be converted to string";
				PyStrContainer StrKey(objDefault, ss);
				m_Default = CA2W(StrKey);
			}			
		}

		ss.str("");
		ss << "Error with Text Variable '" << StrName << "': ";
		PyObject *objMaxLength = PyDict_GetItemString(PyDict, "MaxLength");
		if (objMaxLength)
		{
			if (!PyNumber_Check(objMaxLength))
			{
				ss << "MaxLength must be an integer";
				throw std::exception(ss.str().c_str());
			}
			Py_ssize_t stMaxLength = PyNumber_AsSsize_t(objMaxLength, objExc);
			if (stMaxLength == PY_SSIZE_T_MIN || stMaxLength == PY_SSIZE_T_MAX)
			{
				ss << "MaxLength value could not be converted to integer";
				throw std::exception(ss.str().c_str());
			}
			if (stMaxLength == -1 && PyErr_Occurred())
			{
				if (PyErr_Occurred())
				{
					PyErr_Clear();
				}
				// Not an integer
				ss << "MaxLength must be an integer";
				throw std::exception(ss.str().c_str());
			}
			else
			{
				m_lMaxLength = stMaxLength;
			}
		}

		PyObject *objPassword = PyDict_GetItemString(PyDict, "Password");
		if (objPassword)
		{			
			if (!PyBool_Check(objPassword))
			{
				throw std::exception("Password value must be boolean");
			}
			m_bPassword = (objPassword == Py_False) ? false : true;
		}
		PyObject *objHidden = PyDict_GetItemString(PyDict, "Hidden");
		if (objHidden)
		{
			if (!PyBool_Check(objHidden))
			{
				throw std::exception("Hidden value must be boolean");
			}
			m_bHidden = (objHidden == Py_False) ? false : true;
		}
	}
	TextItem::~TextItem()
	{

	}
	
	void TextItem::ReadRegistry(LPCTSTR szSubKey)
	{
		m_sRegistrySubKey = szSubKey;
		ReadRegistryString(m_Default);
	}
	
	void TextItem::WriteRegistry()
	{
		WriteRegistryString(m_Value);
	}
		
	ComboBoxItem::ComboBoxItem(PyStrContainer &StrName, PyStrContainer &StrDescription, bool bRegistry, PyObject *PyDict) : BaseItem(StrName, StrDescription, bRegistry)
	{		
		std::stringstream ss;
		ss << "Error with ComboBox Variable '" << StrName << "': ";

		// _vComboBoxEntries m_Entries;
		PyObject *objEntries = PyDict_GetItemString(PyDict, "Entries");
		if (!objEntries)
		{
			ss << "No entries specified";
			throw std::exception(ss.str().c_str());
		}
		if (!PyDict_Check(objEntries))
		{
			ss << "Entries value must be a dictionary object";
			throw std::exception(ss.str().c_str());
		}

		PyObject *objKey = NULL;
		PyObject *objValue = NULL;		
		Py_ssize_t pos = 0;
		
		while (PyDict_Next(objEntries, &pos, &objKey, &objValue))
		{
			// TODO: Check to make sure the key is either a number or a string (no tuples)
			// PyNumber_Check and PyUnicode_Check
			if (!PyNumber_Check(objKey) && !PyUnicode_Check(objKey))
			{
				ss << "Key for entry " << pos << " must be either a string or a number";
				throw std::exception(ss.str().c_str());
			}
			AddEntry(pos, objKey, objValue, ss.str().c_str());
		}
		PyObject *objDefault = PyDict_GetItemString(PyDict, "Default");
		PyObject *objExc = NULL;
		std::wstringstream ssDefault;
		if (objDefault)
		{
			if (!PyNumber_Check(objDefault) && !PyUnicode_Check(objDefault))
			{
				ss << "Default must be either a string or a number";
				throw std::exception(ss.str().c_str());
			}
			if (PyNumber_Check(objDefault))
			{
				Py_ssize_t stDefault = PyNumber_AsSsize_t(objDefault, objExc);
				if (stDefault == PY_SSIZE_T_MIN || stDefault == PY_SSIZE_T_MAX)
				{
					ss << "Default value could not be converted to integer";
					throw std::exception(ss.str().c_str());
				}
				ssDefault << stDefault;
			}
			if (PyUnicode_Check(objDefault))
			{
				ss << "Default value could not be converted to string";
				PyStrContainer StrKey(objDefault, ss);
				ssDefault << StrKey;
			}
			m_Default = ssDefault.str();
		}
	}

	
	ComboBoxItem::~ComboBoxItem()
	{
	}

	void ComboBoxItem::ReadRegistry(LPCTSTR szSubKey)
	{
		m_sRegistrySubKey = szSubKey;
		ReadRegistryString(m_Default);		
	}
	void ComboBoxItem::WriteRegistry()
	{
		WriteRegistryString(m_Value);		
	}

	void ComboBoxItem::AddEntry(Py_ssize_t pos, PyObject *objKey, PyObject *objValue, const char *szText)
	{
		std::wstringstream ssKey;
		std::stringstream ssMsg;
		ssMsg << szText;
		PyStrContainer StrValue;
		PyObject *objExc = NULL;
		eValueType lValueType;
		if (PyNumber_Check(objKey))
		{
			lValueType = eVTNumber;
			Py_ssize_t stKey = PyNumber_AsSsize_t(objKey, objExc);
			if (stKey == PY_SSIZE_T_MIN || stKey == PY_SSIZE_T_MAX)
			{
				ssMsg << "Key for entry " << pos << " could not be converted to integer";
				throw std::exception(ssMsg.str().c_str());
			}
			ssKey << stKey;
		}
		if (PyUnicode_Check(objKey))
		{
			lValueType = eVTString;
			std::wstring ws;
			ssMsg << "Could not convert key for entry " << pos << " to string";
			PyStrContainer StrKey(objKey, ssMsg);
			ws = StrKey;
			ssKey << ws.c_str();
		}
		ssMsg << "Value for entry " << pos << " must be string";
		StrValue.SetString(objValue, ssMsg.str().c_str());
		m_Entries.push_back(ComboBoxEntry(ssKey.str().c_str(), StrValue, lValueType ));
	}

	CheckBoxItem::CheckBoxItem(PyStrContainer &StrName, PyStrContainer &StrDescription, bool bRegistry, PyObject *PyDict) : BaseItem(StrName, StrDescription, bRegistry), m_Default(false)
	{		
		std::stringstream ss;
		ss << "Error with CheckBox Variable '" << StrName << "': ";
		PyObject *objDefault = PyDict_GetItemString(PyDict, "Default");
		if (objDefault)
		{
			if (!PyBool_Check(objDefault))
			{
				throw std::exception("Default value must be boolean");
			}
			m_Default = (objDefault == Py_False) ? false : true;
		}
	}
	CheckBoxItem::~CheckBoxItem()
	{

	}
	void CheckBoxItem::ReadRegistry(LPCTSTR szSubKey)
	{
		m_sRegistrySubKey = szSubKey;
		int iValue = m_Default ? 1 : 0;
		ReadRegistryInteger(iValue, 0, 1);
		m_Default = iValue == 0 ? false : true;
	}
	void CheckBoxItem::WriteRegistry()
	{
		int iValue = m_Value ? 1 : 0;
		WriteRegistryInteger(iValue);
	}

	void FormVariables::SetValuesFromDefaults()
	{
		for (_CheckBoxPair &CheckBox : m_vCheckBoxes)
		{
			CheckBox.second.SetFromDefault();
		}
		for (_ComboBoxPair &ComboBox : m_vComboBoxes)
		{
			ComboBox.second.SetFromDefault();
			if (ComboBox.second.m_Entries.size() == 0)
			{
				// We haven't been given any entries, so we cannot determine what the output data type should be since
				// all values are stored in the registry as strings regardless.  Set the value type to string and 
				// let the calling function make the decision
				ComboBox.second.m_ValueType = eVTString;
			}
			else
			{
				for (ComboBoxEntry Entry : ComboBox.second.m_Entries)
				{
					if (!_tcsicmp(Entry.m_sKey.c_str(), ComboBox.second.m_Value.c_str()))
					{
						ComboBox.second.m_ValueType = Entry.m_ValueType;
						break;
					}
				}
			}
			
		}
		for (_TextItemPair &TextItem : m_vTextItems)
		{
			TextItem.second.SetFromDefault();
		}
	}
	
	void FormVariables::DoDataExchange(IHTMLElementCollection *pCollection)
	{
		variant_t vtID;
		::VariantInit(&vtID);
		vtID.vt = VT_I4;
		vtID.lVal = 0;
		HRESULT hr;
		
		for (_itCheckBox itCheckBox = m_vCheckBoxes.begin(); itCheckBox != m_vCheckBoxes.end(); itCheckBox++)
		{
			IHTMLInputElement *pInput;
			IDispatch *pItem = NULL;
			variant_t vtName;
			::VariantInit(&vtName);
			vtName.SetString(CW2A((*itCheckBox).second.m_sName.c_str()));
			pCollection->item(vtName, vtID, &pItem);
			if (pItem)
			{
				hr = pItem->QueryInterface(IID_IHTMLInputElement, (void **)&pInput);
				pItem->Release();
				if (!FAILED(hr))
				{	
					VARIANT_BOOL vtBool;
					pInput->get_checked(&vtBool);
					(*itCheckBox).second.m_Value = (vtBool == VARIANT_TRUE) ? true : false;
					pInput->Release();
					(*itCheckBox).second.WriteRegistry();
				}
			}
		}
		for (_itComboBox itComboBox = m_vComboBoxes.begin(); itComboBox != m_vComboBoxes.end(); itComboBox++)
		{
			IHTMLSelectElement *pSelect;
			IDispatch *pItem = NULL;
			variant_t vtName;
			::VariantInit(&vtName);
			vtName.SetString(CW2A((*itComboBox).second.m_sName.c_str()));
			pCollection->item(vtName, vtID, &pItem);
			if (pItem)
			{
				hr = pItem->QueryInterface(IID_IHTMLSelectElement, (void **)&pSelect);
				pItem->Release();
				if (!FAILED(hr))
				{
					BSTR bstr;
					pSelect->get_value(&bstr);
					if (bstr)
					{
						std::unique_ptr<char> b2s(_com_util::ConvertBSTRToString(bstr));
						(*itComboBox).second.m_Value = CA2W(b2s.get());
					}
					else
					{
						(*itComboBox).second.m_Value = _T("");
					}
					pSelect->Release();
					(*itComboBox).second.WriteRegistry();
				}
			}

			// Go through all of the entries for this combobox and find the matching key.  Then set our class-level value type member to
			// the corresponding value type
			_itComboBoxEntry itEntry;
			for (itEntry = (*itComboBox).second.m_Entries.begin(); itEntry != (*itComboBox).second.m_Entries.end(); itEntry++)
			{
				if (!_tcsicmp((*itEntry).m_sKey.c_str(), (*itComboBox).second.m_Value.c_str()))
				{
					(*itComboBox).second.m_ValueType = (*itEntry).m_ValueType;
					break;
				}
			}
		}		
		for (_itTextItem itTextItem = m_vTextItems.begin(); itTextItem != m_vTextItems.end(); itTextItem++)
		{
			IHTMLInputElement *pInput;
			IDispatch *pItem = NULL;
			variant_t vtName;
			::VariantInit(&vtName);
			vtName.SetString(CW2A((*itTextItem).second.m_sName.c_str()));
			pCollection->item(vtName, vtID, &pItem);
			if (pItem)
			{
				hr = pItem->QueryInterface(IID_IHTMLInputElement, (void **)&pInput);
				pItem->Release();
				if (!FAILED(hr))
				{
					BSTR bstr;
					pInput->get_value(&bstr);
					if (bstr)
					{
						std::unique_ptr<char> b2s(_com_util::ConvertBSTRToString(bstr));
						(*itTextItem).second.m_Value = CA2W(b2s.get());
						(*itTextItem).second.WriteRegistry();
					}
					pInput->Release();
				}
			}
		}		
	}
	
	CString FormVariables::BuildHTMLForm()
	{
		// First load the desired template.  If none specified, load the default
		CMemBuffer mbTemplate;
		CMemBuffer mbCheckBox;
		CMemBuffer mbComboBox;
		CMemBuffer mbComboBoxOption;
		CMemBuffer mbInput;		
		std::stringstream ssMsg;
		//std::wstring wsDocDir;
		std::wstringstream wsDefaultTemplate;
		LPExecuteParams lpExecuteParams = reinterpret_cast<LPExecuteParams>(m_lpExecuteParams);
		
		//GetDocumentsDirectory(wsDocDir);
		if (m_sHTMLTemplate.size() > 0)
		{
			if (!mbTemplate.InitFromFile(m_sHTMLTemplate.c_str()))
			{
				std::string s(m_sHTMLTemplate.begin(), m_sHTMLTemplate.end());
				ssMsg << "Could not open HTML template '" << s << "'";
				throw std::exception(ssMsg.str().c_str());
			}
		}
		else
		{
			//wsDefaultTemplate << wsDocDir;
			//wsDefaultTemplate << _T("\\Pythagoras\\Templates\\Common\\Default.html");			
			DebugOut(_T("HTML template:"), true);
			DebugOut(lpExecuteParams->sDefaultHTMLTemplate.c_str(), true);
			if (!mbTemplate.InitFromFile(lpExecuteParams->sDefaultHTMLTemplate.c_str()))
			{				
				ssMsg << "Could not open HTML template '" << CW2A(lpExecuteParams->sDefaultHTMLTemplate.c_str()) << "'";
				throw std::exception(ssMsg.str().c_str());
			}
		}

		// First things first, go through all of the replace strings and replace them in the block.  We want to do this before
		// parsing out the individual blocks so that we don't have to execute this multiple times.
		for (_itStringPair itStringPair = m_vReplaceStrings.begin(); itStringPair != m_vReplaceStrings.end(); itStringPair++)
		{			
			mbTemplate.ReplaceString((*itStringPair).first.c_str(), (*itStringPair).second.c_str());
		}

		mbTemplate.MoveStringBlock("<!--BEGIN_CHECKBOX-->", "<!--END_CHECKBOX-->", &mbCheckBox);
		mbTemplate.MoveStringBlock("<!--BEGIN_COMBOBOX-->", "<!--END_COMBOBOX-->", &mbComboBox);
		mbTemplate.MoveStringBlock("<!--BEGIN_INPUT-->", "<!--END_INPUT-->", &mbInput);
		
		mbCheckBox.ReplaceString("<!--BEGIN_CHECKBOX-->", "");
		mbCheckBox.ReplaceString("<!--END_CHECKBOX-->", "");

		mbComboBox.ReplaceString("<!--BEGIN_COMBOBOX-->", "");
		mbComboBox.ReplaceString("<!--END_COMBOBOX-->", ""); 

		mbInput.ReplaceString("<!--BEGIN_INPUT-->", "");
		mbInput.ReplaceString("<!--END_INPUT-->", "");
				
		mbComboBox.MoveStringBlock("<!--BEGIN_OPTION-->", "<!--END_OPTION-->", &mbComboBoxOption);

		mbComboBoxOption.ReplaceString("<!--BEGIN_OPTION-->", "");
		mbComboBoxOption.ReplaceString("<!--END_OPTION-->", "");

		CMemBuffer mbCheckCopy;
		CMemBuffer mbComboCopy;
		CMemBuffer mbOptionCopy;
		CMemBuffer mbInputCopy;
		
		_itCheckBox itCheckBox = m_vCheckBoxes.begin();
		_itComboBox itComboBox = m_vComboBoxes.begin();
		_itTextItem itTextItem = m_vTextItems.begin();
		unsigned int iOrder = 0;	
		// This 2048 number is arbitrary, intended to be larger than remotely possible (we would never expect a 
		// GUI to have over 2000 fields on it).  This is just to avoid 'while(true)' and infinite looping in case
		// we mess up the break sequence down below...
		while(iOrder < 2048)		
		{	
			if (itCheckBox != m_vCheckBoxes.end())
			{
				if ((*itCheckBox).first == iOrder)
				{
					mbCheckBox.CopyBuffer(mbCheckCopy);
					mbCheckCopy.ReplaceString(_T("<!--ID-->"), (*itCheckBox).second.m_sName.c_str());
					mbCheckCopy.ReplaceString(_T("<!--DESCRIPTION-->"), (*itCheckBox).second.m_sDescription.c_str());
					mbCheckCopy.ReplaceString("<!--CHECKED-->", (*itCheckBox).second.m_Default ? "checked" : "");

					mbTemplate.InsertBuffer("<!--INSERT_INPUT_ELEMENTS-->", mbCheckCopy, MEMBUFFER_FLAG_INSERT_BEFORE);
					itCheckBox++;
				}
			}
			if (itComboBox != m_vComboBoxes.end())
			{
				if ((*itComboBox).first == iOrder)
				{
					mbComboBox.CopyBuffer(mbComboCopy);
					mbComboCopy.ReplaceString(_T("<!--ID-->"), (*itComboBox).second.m_sName.c_str());
					mbComboCopy.ReplaceString(_T("<!--DESCRIPTION-->"), (*itComboBox).second.m_sDescription.c_str());

					_itComboBoxEntry itCBEntry = (*itComboBox).second.m_Entries.begin();
					while (itCBEntry != (*itComboBox).second.m_Entries.end())
					{
						mbComboBoxOption.CopyBuffer(mbOptionCopy);
						mbOptionCopy.ReplaceString(_T("<!--KEY-->"), (*itCBEntry).m_sKey.c_str());
						mbOptionCopy.ReplaceString(_T("<!--VALUE-->"), (*itCBEntry).m_sValue.c_str());
						if (!_tcsicmp((*itComboBox).second.m_Default.c_str(), (*itCBEntry).m_sKey.c_str()))
						{
							mbOptionCopy.ReplaceString("<!--SELECTED-->", " selected");
						}
						else
						{
							mbOptionCopy.ReplaceString("<!--SELECTED-->", "");
						}
						mbComboCopy.InsertBuffer("<!--INSERT_OPTIONS-->", mbOptionCopy, MEMBUFFER_FLAG_INSERT_BEFORE);
						itCBEntry++;
					}

					mbTemplate.InsertBuffer("<!--INSERT_INPUT_ELEMENTS-->", mbComboCopy, MEMBUFFER_FLAG_INSERT_BEFORE);

					itComboBox++;
				}
			}
			if (itTextItem != m_vTextItems.end())
			{
				if ((*itTextItem).first == iOrder)
				{
					mbInput.CopyBuffer(mbInputCopy);
					mbInputCopy.ReplaceString(_T("<!--ID-->"), (*itTextItem).second.m_sName.c_str());
					if (!(*itTextItem).second.m_bHidden)
					{
						mbInputCopy.ReplaceString(_T("<!--DESCRIPTION-->"), (*itTextItem).second.m_sDescription.c_str());
					}
					else
					{
						mbInputCopy.ReplaceString(_T("<!--DESCRIPTION-->"), _T(""));
					}
					mbInputCopy.ReplaceString(_T("<!--DEFAULT-->"), (*itTextItem).second.m_Default.c_str());
					if ((*itTextItem).second.m_bPassword)
					{
						mbInputCopy.ReplaceString("<!--TYPE-->", "password");
					}
					else
					{
						if ((*itTextItem).second.m_bHidden)
						{
							mbInputCopy.ReplaceString("<!--TYPE-->", "hidden");
						}
						else
						{
							mbInputCopy.ReplaceString("<!--TYPE-->", "text");
						}
					}
					if ((*itTextItem).second.m_lMaxLength == -1)
					{
						mbInputCopy.ReplaceString("<!--MAXLENGTH-->", "");
					}
					else
					{
						std::stringstream ssFormat;
						ssFormat << (*itTextItem).second.m_lMaxLength;
						mbInputCopy.ReplaceString("<!--MAXLENGTH-->", ssFormat.str().c_str());
					}

					mbTemplate.InsertBuffer("<!--INSERT_INPUT_ELEMENTS-->", mbInputCopy, MEMBUFFER_FLAG_INSERT_BEFORE);
					itTextItem++;
				}
			}			
			if (itCheckBox == m_vCheckBoxes.end() && itComboBox == m_vComboBoxes.end() && itTextItem == m_vTextItems.end())
			{
				break;
			}
			iOrder++;

			if (iOrder > m_vCheckBoxes.size() + m_vComboBoxes.size() + m_vTextItems.size())
			{
				break;
			}
		}
		return CString(mbTemplate.GetBufferAsString());
	}

	void FormVariables::AddToDictionary(PyObject *pDict)
	{
		PyObject *pValue = NULL;
		PyRefHandler RefHandler;

		// Go through all of the formvariables and append them to the dictionary.
		for (_CheckBoxPair CheckItem : m_vCheckBoxes)
		{
			PyDict_SetItemWideChar(pDict, CheckItem.second.m_sName.c_str(), CheckItem.second.m_Value ? Py_True : Py_False);
		}

		for (_ComboBoxPair ComboItem : m_vComboBoxes)
		{
			switch (ComboItem.second.m_ValueType)
			{
			case eVTString:
			{
				pValue = RefHandler.AddRef(PyUnicode_FromWideChar(ComboItem.second.m_Value.c_str(),-1));
				PyDict_SetItemWideChar(pDict, ComboItem.second.m_sName.c_str(), pValue);
			}
			break;

			case eVTNumber:
			{
				pValue = Py_BuildValue("l", _tstoi(ComboItem.second.m_Value.c_str()));
				PyDict_SetItemWideChar(pDict, ComboItem.second.m_sName.c_str(), pValue);
			}
			break;
			}
		}

		for (_TextItemPair TextItem : m_vTextItems )
		{			
			pValue = RefHandler.AddRef(PyUnicode_FromWideChar(TextItem.second.m_Value.c_str(),-1));
			PyDict_SetItemWideChar(pDict, TextItem.second.m_sName.c_str(), pValue);
		}
	}

	void FormVariables::ParseHTMLDict(PyObject *pDict)
	{
		PyRefHandler RefHandler;
		PyObject *objKey = NULL;
		PyObject *objValue = NULL;		
		Py_ssize_t pos = 0;
		std::stringstream ss;
		ss << "Error with Variable 'HTML': ";
		while (PyDict_Next(pDict, &pos, &objKey, &objValue))
		{
			PyStrContainer StrKey(objKey, "Dictionary keys must be strings");
			switch (static_cast<eHTMLKeyID>(StrKey))
			{
			case eHKITemplate:
			{
				PyStrContainer StrTemplate(objValue, "Template must be a string");
				m_sHTMLTemplate = StrTemplate.w_str();
				ReplaceLocaleSpecificStrings(m_sHTMLTemplate, this);
			}
			break;
			case eHKIWidth:
			{
				if (PyNumber_Check(objValue))
				{
					PyObject *objExc = NULL;
					Py_ssize_t stValue = PyNumber_AsSsize_t(objValue, objExc);
					if (stValue == PY_SSIZE_T_MIN || stValue == PY_SSIZE_T_MAX)
					{
						ss << "Width could not be converted to integer";
						throw std::exception(ss.str().c_str());
					}
					m_uiWidth = stValue;
				}
				else
				{
					ss << "Width must be integer";
					throw std::exception(ss.str().c_str());
				}
			}
			break;
			case eHKIHeight:
			{
				if (PyNumber_Check(objValue))
				{
					PyObject *objExc = NULL;
					Py_ssize_t stValue = PyNumber_AsSsize_t(objValue, objExc);
					if (stValue == PY_SSIZE_T_MIN || stValue == PY_SSIZE_T_MAX)
					{
						ss << "Height could not be converted to integer";
						throw std::exception(ss.str().c_str());
					}
					m_uiHeight = stValue;
				}
				else
				{
					ss << "Height must be integer";
					throw std::exception(ss.str().c_str());
				}
			}
			break;
			case eHKIBGColor:
			{
				PyStrContainer StrBGColor(objValue, "BGColor must be a string");
				m_sBGColor = StrBGColor;
			}
			break;
			case eHKIReplaceStrings:
			{
				if (!PyDict_Check(objValue))
				{
					ss << "ReplaceString must be a dictionary object";
					throw std::exception(ss.str().c_str());
				}
				PyObject *objRSKey = NULL;
				PyObject *objRSValue = NULL;
				Py_ssize_t RSpos = 0;
				std::stringstream ssMsg;
				ssMsg << ss.str();
				while (PyDict_Next(objValue, &RSpos, &objRSKey, &objRSValue))
				{
					ssMsg << "ReplaceString key at position " << RSpos << " must be string";
					PyStrContainer StrRSKey(objRSKey, ssMsg);
					ssMsg.str("");
					ssMsg << "Error with Variable 'HTML': ReplaceString value at position " << RSpos << " must be string";
					PyStrContainer StrRSValue(objRSValue, ssMsg);
					m_vReplaceStrings.push_back(std::make_pair(StrRSKey.w_str(), StrRSValue.w_str()));
				}
			}
			break;
			case eHKIUnknown:
			default:
			{
				ss << "Unrecognized key '" << StrKey << "'";
				throw std::exception(ss.str().c_str());
			}
			}
		}
	}

	void FormVariables::ClearForm()
	{
		m_vReplaceStrings.erase(m_vReplaceStrings.begin(), m_vReplaceStrings.end());
		m_vComboBoxes.erase(m_vComboBoxes.begin(), m_vComboBoxes.end());
		m_vCheckBoxes.erase(m_vCheckBoxes.begin(), m_vCheckBoxes.end());
		m_vTextItems.erase(m_vTextItems.begin(), m_vTextItems.end());
	}

	void FormatPythonError(std::wstringstream &sError)
	{
		std::stringstream ss;
		std::wstring s;
		FormatPythonError(ss);
		s = CA2W(ss.str().c_str());
		sError.clear();
		sError << s.c_str();
	}

	void FormatPythonError(std::stringstream &sError)
	{
		PyRefHandler RefHandler;
		PyStrContainer strExcType;
		PyStrContainer strExcValue;
		PyStrContainer strExcTraceback;
		if (PyErr_Occurred())
		{
			PyObject *exc_type = NULL, *exc_value = NULL, *exc_tb = NULL;

			PyErr_Fetch(&exc_type, &exc_value, &exc_tb);
			RefHandler.AddRef(exc_type);
			RefHandler.AddRef(exc_value);
			RefHandler.AddRef(exc_tb);

			PyObject* str_exc_type = RefHandler.AddRef(PyObject_Repr(exc_type)); //Now a unicode object
			strExcType.SetString(str_exc_type, "");			

			PyObject* str_exc_value = RefHandler.AddRef(PyObject_Repr(exc_value));
			strExcValue.SetString(str_exc_value, "");			

			/* See if we can get a full traceback */
			PyObject *module_name = NULL, *pyth_module = NULL, *pyth_func = NULL;
			module_name = RefHandler.AddRef(PyUnicode_FromString("traceback"));
			pyth_module = RefHandler.AddRef(PyImport_Import(module_name));

			if (pyth_module)
			{
				pyth_func = RefHandler.AddRef(PyObject_GetAttrString(pyth_module, "format_exception"));
				if (pyth_func && PyCallable_Check(pyth_func)) {
					PyObject *pyth_val = NULL;

					pyth_val = RefHandler.AddRef(PyObject_CallFunctionObjArgs(pyth_func, exc_type, exc_value, exc_tb, NULL));

					if (pyth_val && PyList_Check(pyth_val))
					{
						PyObject *msg = PyBytes_FromString("");
						Py_ssize_t i = 0, lines_len = PyList_GET_SIZE(pyth_val);
						for (; i < lines_len; i++)
						{
							PyBytes_ConcatAndDel(&msg, PyList_GET_ITEM(pyth_val, i));
							RefHandler.AddRef(msg);
							if (msg == NULL)
							{
								break;
							}
						}

						try
						{
							strExcTraceback.SetStringFromBytes(msg, "");
						}
						catch (...)
						{
						}
						
					}
				}
			}
		}

		sError << "Error executing Python script: " << strExcType << strExcValue << (strExcTraceback.m_String ? strExcTraceback : "");
	}

	void ParseFormVariables_v1_0(PyObject *objDict, LPExecuteParams lpParams, bool bRegistryOnly)
	{
		PyRefHandler RefHandler;
		PyObject *objKey = NULL;
		PyObject *objValue = NULL;
		PyObject *sType = NULL;
		PyObject *sDesc = NULL;
		PyObject *sRegistry = NULL;
		PyObject *objType = NULL;
		PyObject *objDesc = NULL;
		PyObject *objRegistry = NULL;
		Py_ssize_t pos = 0;
		unsigned int iOrder = 0;

		while (PyDict_Next(objDict, &pos, &objKey, &objValue))
		{
			PyStrContainer StrKey(objKey, "FormVariables dictionary keys must be strings");
			std::stringstream ss;
			ss << "Error with Variable '" << StrKey << "': ";
						
			// Handle reserved keys...
			if (!_stricmp("Version", StrKey))
			{
				// Already handled, just skip it...
				continue;
			}
			if (!PyDict_Check(objValue))
			{
				ss << "not a dictionary object";
				throw std::exception(ss.str().c_str());
			}
			if (!_stricmp("HTML", StrKey))
			{
				if (!PyDict_Check(objValue))
				{
					ss << "not a dictionary object";
					throw std::exception(ss.str().c_str());
				}
				lpParams->FV.ParseHTMLDict(objValue);
				continue;
			}

			// Handle user-defined keys...
			sType = RefHandler.AddRef(PyUnicode_FromString("Type"));
			objType = PyDict_GetItem(objValue, sType);
			if (!objType)
			{
				ss << "does not specify Type";
				throw std::exception(ss.str().c_str());
			}
			std::stringstream ssType;
			ssType << ss.str() << "Type value must be string";
			PyStrContainer StrType(objType, ssType);

			sDesc = RefHandler.AddRef(PyUnicode_FromString("Description"));
			objDesc = PyDict_GetItem(objValue, sDesc);
			PyStrContainer StrDesc;
			if (!objDesc )
			{
				if (!bRegistryOnly)
				{
					ss << "does not specify Description";
					throw std::exception(ss.str().c_str());
				}
			}
			else
			{
				std::stringstream ssDesc;
				ssDesc << ss.str() << "Description value must be string";
				StrDesc.SetString(objDesc, ssDesc);
			}
			
			bool bRegistry = true;
			// If we're only looking to get default values from the registry, then no need to look for this key...
			if (!bRegistryOnly)
			{
				sRegistry = RefHandler.AddRef(PyUnicode_FromString("Registry"));
				objRegistry = PyDict_GetItem(objValue, sRegistry);
				if (objRegistry)
				{
					if (!PyBool_Check(objRegistry))
					{
						ss << "Registry value must be boolean";
						throw std::exception(ss.str().c_str());
					}
					bRegistry = (objRegistry == Py_False) ? false : true;
				}
			}

			switch (static_cast<eVariableID>(StrType))
			{
			case eVIDComboBox:
			{
				lpParams->FV.m_vComboBoxes.push_back(std::make_pair(iOrder++, ComboBoxItem(StrKey, StrDesc, bRegistry, objValue)));
			}
			break;

			case eVIDCheckBox:
			{
				lpParams->FV.m_vCheckBoxes.push_back(std::make_pair(iOrder++, CheckBoxItem(StrKey, StrDesc, bRegistry, objValue)));
			}
			break;

			case eVIDText:
			{
				lpParams->FV.m_vTextItems.push_back(std::make_pair(iOrder++, TextItem(StrKey, StrDesc, bRegistry, objValue)));
			}
			break;

			case eVIDUnknown:
			default:
			{
				ss << "Unrecognized type '" << StrType << "' for Variable '" << StrKey << "'";
				throw std::exception(ss.str().c_str());
			}
			break;
			}
		}

		// Now go through and set the defaults according to what's in the registry
		
		std::wstringstream ssSubKey;
		std::vector<std::wstring> sv;
		boost::split(sv, lpParams->sScriptRegistry, boost::is_any_of("."));
		ssSubKey << _T("Software\\3M\\Pythagoras\\ScriptSettings\\");
		for (std::wstring ws : sv)
		{
			ssSubKey << ws.c_str() << _T("\\");
		}
		ssSubKey << lpParams->sFunction.c_str() << _T("\\");

		for (std::pair<unsigned int, ComboBoxItem> &Item : lpParams->FV.m_vComboBoxes)
		{
			Item.second.ReadRegistry(ssSubKey.str().c_str());
		}
		for (std::pair<unsigned int, CheckBoxItem> &Item : lpParams->FV.m_vCheckBoxes)
		{
			Item.second.ReadRegistry(ssSubKey.str().c_str());
		}
		for (std::pair<unsigned int, TextItem> &Item : lpParams->FV.m_vTextItems)
		{
			Item.second.ReadRegistry(ssSubKey.str().c_str());
		}
		
	}

	void ParseFormVariables(PyObject *objDict, LPExecuteParams lpParams, bool bRegistryOnly)
	{
		PyRefHandler RefHandler;
		PyObject *objKey = NULL;
		PyObject *objValue = NULL;
		PyObject *sType = NULL;
		PyObject *sDesc = NULL;
		PyObject *objType = NULL;
		Py_ssize_t pos = 0;
		if (!PyDict_Check(objDict))
		{
			throw std::exception("FormVariables must be a dictionary object");
		}

		// First find the version... this will dictate how everything else will be handled
		eVersionID lVersionID = eVersionUnknown;
		while (PyDict_Next(objDict, &pos, &objKey, &objValue))
		{
			PyStrContainer StrKey(objKey, "FormVariables dictionary keys must be strings");
#if 0
			if (!_stricmp("Mode", StrKey))
			{
				throw std::exception("'Mode' is a reserved keyword and cannot be used as a variable name");
			}
			if (!_stricmp("SourceFiles", StrKey))
			{
				throw std::exception("'SourceFiles' is a reserved keyword and cannot be used as a variable name");
			}
#endif
			if (!_stricmp("Version", StrKey))
			{
				std::stringstream ssMsg;
				ssMsg << "Error with Version key: Value must be a string";
				PyStrContainer StrVersion(objValue, ssMsg);
				lVersionID = static_cast<eVersionID>(StrVersion);
			}
		}

		// Erase any existing form variables
		lpParams->FV.ClearForm();

		// Call the right function depending on version.  
		switch (lVersionID)
		{
		case eVersion_1_0:
		{
			ParseFormVariables_v1_0(objDict, lpParams, bRegistryOnly);
		}
		break;

		case eVersionUnknown:
		default:
		{
			throw std::exception("Unrecognized version during parsing of FormVariables");
		}
		break;
		}
	}

	void ParseReturnVariables_v1_0(PyObject *objDict, LPExecuteParams lpParams)
	{
		PyRefHandler RefHandler;
		PyObject *objKey = NULL;
		PyObject *objValue = NULL;
		PyObject *sType = NULL;
		PyObject *sDesc = NULL;
		PyObject *objType = NULL;
		PyObject *objDesc = NULL;
		Py_ssize_t pos = 0;
		unsigned int iOrder = 0;

		while (PyDict_Next(objDict, &pos, &objKey, &objValue))
		{
			PyStrContainer StrKey(objKey, "Return variable dictionary keys must be strings");
			std::stringstream ss;
			ss << "Error with Variable '" << StrKey << "': ";

			// Handle reserved keys...
			if (!_stricmp("Version", StrKey))
			{
				// Already handled, just skip it...
				continue;
			}			
			if (!_stricmp("OpenFiles", StrKey))
			{
				if (PyList_Check(objValue))
				{
					ExtractPythonListAsString(lpParams->vOpenFiles, objValue);
				}
				else
				{
					PyStrContainer StrFileName(objValue, "OpenFiles variable must be a list of strings or a single string");
					lpParams->vOpenFiles.push_back(StrFileName.w_str());
				}
				continue;
			}
		}
	}
	void ParseReturnVariables(PyObject *objDict, LPExecuteParams lpParams)
	{
		PyRefHandler RefHandler;
		PyObject *objKey = NULL;
		PyObject *objValue = NULL;
		PyObject *sType = NULL;
		PyObject *sDesc = NULL;
		PyObject *objType = NULL;
		Py_ssize_t pos = 0;
		if (!PyDict_Check(objDict))
		{
			// This code should not be reachable since we check for a dictionary object in the calling function...
			throw std::exception("Function must return a dictionary object");
		}

		// First find the version... this will dictate how everything else will be handled
		eVersionID lVersionID = eVersionUnknown;
		while (PyDict_Next(objDict, &pos, &objKey, &objValue))
		{
			PyStrContainer StrKey(objKey, "Dictionary keys for return variable must be strings");
#if 0
			if (!_stricmp("Mode", StrKey))
			{
				throw std::exception("'Mode' is a reserved keyword and cannot be used as a variable name");
			}
			if (!_stricmp("SourceFiles", StrKey))
			{
				throw std::exception("'SourceFiles' is a reserved keyword and cannot be used as a variable name");
			}
			if (!_stricmp("SourceFolders", StrKey))
			{
				throw std::exception("'SourceFolders' is a reserved keyword and cannot be used as a variable name");
			}
#endif
			if (!_stricmp("Version", StrKey))
			{
				std::stringstream ssMsg;
				ssMsg << "Error with Version key: Value must be a string";
				PyStrContainer StrVersion(objValue, ssMsg);
				lVersionID = static_cast<eVersionID>(StrVersion);
			}
		}

		// Call the right function depending on version.  
		switch (lVersionID)
		{
		case eVersion_1_0:
		{
			ParseReturnVariables_v1_0(objDict, lpParams);
		}
		break;

		case eVersionUnknown:
		default:
		{
			throw std::exception("Unrecognized version during parsing of return variable");
		}
		break;
		}
	}

	void ParseKeyValue(PyStrContainer &strKey, PyObject *objValue)
	{		
	}

	bool ExecuteScript(LPExecuteParams lpParams, eExecuteMode eEM)
	{	
		PyRefHandler RefHandler;
		std::wstringstream ssBaseFolder;
		PyObject *pName = NULL, *pModule = NULL, *pDict = NULL, *pFunc = NULL;
		PyObject *pArgs = NULL, *pValue = NULL, *pList = NULL;
		int iArg = 0;
		_itString itString;

		TCHAR drive[_MAX_DRIVE];
		TCHAR dir[_MAX_DIR];
		TCHAR fname[_MAX_FNAME];
		
		_wsplitpath_s(lpParams->sPathToScript.c_str(), drive, _MAX_DRIVE, dir, _MAX_DIR, fname, _MAX_FNAME, NULL, 0);
		ssBaseFolder << drive << dir;
		lpParams->bSuccess = true;
		
		try
		{			
			switch (eEM)
			{
#if 0
			case eEMQueryVariables:
			{
				// First call the script with "QueryVariables" to find out what we need to prompt the user for
				PyObject* objPath = PySys_GetObject((char*)"path");
				PyObject* objFolder = RefHandler.AddRef(PyUnicode_FromWideChar(ssBaseFolder.str().c_str(), ssBaseFolder.str().size()));
				PyList_Append(objPath, objFolder);

				pName = RefHandler.AddRef(PyUnicode_FromWideChar(fname,-1));
				pModule = RefHandler.AddRef(PyImport_Import(pName));
				if (!pModule)
				{
					throw std::exception("Failed to import module");
				}
				pFunc = RefHandler.AddRef(PyObject_GetAttrWideChar(pModule, lpParams->sFunction.c_str()));
				if (!pFunc || !PyCallable_Check(pFunc))
				{
					throw std::exception("Function is not valid or not callable");
				}
				pDict = RefHandler.AddRef(PyDict_New());
				pValue = RefHandler.AddRef(PyUnicode_FromString("QueryVariables"));
				PyDict_SetItemString(pDict, "Mode", pValue);
				
				pValue = RefHandler.AddRef(PyUnicode_FromString("1.0"));
				PyDict_SetItemString(pDict, "Version", pValue);

				// Add the filenames
				pList = RefHandler.AddRef(PyList_New(lpParams->vFilenames.size()));
				for (iArg = 0, itString = lpParams->vFilenames.begin(); itString != lpParams->vFilenames.end(); itString++, iArg++)
				{
					pValue = RefHandler.AddRef(PyUnicode_FromWideChar((*itString).c_str(), (*itString).size()));
					PyList_SetItem(pList, iArg, pValue);
				}
				PyDict_SetItemString(pDict, "SourceFiles", pList);

				// Add the folder names
				pList = RefHandler.AddRef(PyList_New(lpParams->vFoldernames.size()));
				for (iArg = 0, itString = lpParams->vFoldernames.begin(); itString != lpParams->vFoldernames.end(); itString++, iArg++)
				{
					pValue = RefHandler.AddRef(PyUnicode_FromWideChar((*itString).c_str(), (*itString).size()));
					PyList_SetItem(pList, iArg, pValue);
				}
				PyDict_SetItemString(pDict, "SourceFolders", pList);
				
				pArgs = RefHandler.AddRef(PyTuple_New(1));
				PyTuple_SetItem(pArgs, 0, pDict);

				// QueryVariables should return a dictionary object
				pValue = RefHandler.AddRef(PyObject_CallObject(pFunc, pArgs));
				if (!pValue)
				{
					throw std::exception("Script in QueryVariables mode did not return a dictionary object");
				}
				else
				{
					if (pValue != Py_None)
					{
						ParseFormVariables(pValue, lpParams);
					}
					else
					{
						throw std::exception("Script in QueryVariables mode did not return a dictionary object");
					}
				}
			}
			break;
#endif
			case eEMExecute:
			default:
			{
				// Add the base folder to the Python path
				PyObject* objPath = PySys_GetObject((char*)"path");
				PyObject* objFolder = RefHandler.AddRef(PyUnicode_FromWideChar(ssBaseFolder.str().c_str(), ssBaseFolder.str().size()));
				PyList_Append(objPath, objFolder);
				
				// Load the scripts into MemBuffer, append the inject script, then write it out to a temp python script
				CMemBuffer mbScript;
				std::wstring wScript(lpParams->sPathToScript.begin(), lpParams->sPathToScript.end());
				std::wstring wInject(lpParams->sInjectScript.begin(), lpParams->sInjectScript.end());
				std::wstring wTempFolder(lpParams->sTempFolder.begin(), lpParams->sTempFolder.end());
				if (!mbScript.InitFromFile(wScript.c_str()))
				{
					throw std::exception("Failed to load python script");
				}
				mbScript.AppendString("\n");
				if (!mbScript.AddFile(wInject.c_str()))
				{
					throw std::exception("Failed to inject python script");
				}
				// Create temp file path
				std::wstring wTempScript;
				wTempScript.append(wTempFolder.begin(), wTempFolder.end());
				wTempScript.append(_T("\\PythagorasTemp.py"));
				if (!mbScript.WriteToFile(wTempScript.c_str()))
				{
					throw std::exception("Failed to write to temporary script");
				}				
				
				// Add the 'Temp' path to the Python path				
				objFolder = RefHandler.AddRef(PyUnicode_FromWideChar(lpParams->sTempFolder.c_str(),lpParams->sTempFolder.size()));
				PyList_Append(objPath, objFolder);

				pName = RefHandler.AddRef(PyUnicode_FromString("PythagorasTemp"));
				pModule = RefHandler.AddRef(PyImport_Import(pName));
				if (!pModule)
				{
					throw std::exception("Failed to import module");
				}				
				pFunc = RefHandler.AddRef(PyObject_GetAttrString(pModule, "_PythagorasCall"));
				if (!pFunc || !PyCallable_Check(pFunc))
				{
					throw std::exception("Function is not valid or not callable");
				}
				// Now execute the script... first argument is the name of the function to call, second is a dictionary containing
				// various configuration parameters (version, etc).  Note the actual function that will be called here (_PythagorasCall)
				// was appended to the script file
				pArgs = RefHandler.AddRef(PyTuple_New(2));
				pValue = RefHandler.AddRef(PyUnicode_FromWideChar(lpParams->sFunction.c_str(), lpParams->sFunction.size()));
				PyTuple_SetItem(pArgs, 0, pValue);
				
				pDict = RefHandler.AddRef(PyDict_New());
				//pValue = RefHandler.AddRef(PyUnicode_FromString("Execute"));
				//PyDict_SetItemString(pDict, "Mode", pValue);
				pValue = RefHandler.AddRef(PyUnicode_FromString("1.1"));
				PyDict_SetItemString(pDict, "Version", pValue);

#if 0
				// Add all the form variables
				lpParams->FV.AddToDictionary(pDict);
#endif

				// Add the filenames
				pList = RefHandler.AddRef(PyList_New(lpParams->vFilenames.size()));
				for (iArg = 0, itString = lpParams->vFilenames.begin(); itString != lpParams->vFilenames.end(); itString++, iArg++)
				{
					pValue = RefHandler.AddRef(PyUnicode_FromWideChar((*itString).c_str(), (*itString).size()));
					PyList_SetItem(pList, iArg, pValue);
				}
				PyDict_SetItemString(pDict, "SourceFiles", pList);				
				
				// Add the folder names
				pList = RefHandler.AddRef(PyList_New(lpParams->vFoldernames.size()));
				for (iArg = 0, itString = lpParams->vFoldernames.begin(); itString != lpParams->vFoldernames.end(); itString++, iArg++)
				{
					pValue = RefHandler.AddRef(PyUnicode_FromWideChar((*itString).c_str(), (*itString).size()));
					PyList_SetItem(pList, iArg, pValue);
				}
				PyDict_SetItemString(pDict, "SourceFolders", pList);
				
				PyTuple_SetItem(pArgs, 1, pDict);

				pValue = RefHandler.AddRef(PyObject_CallObject(pFunc, pArgs));
				if (!pValue)
				{
					throw std::exception("Function did not return a dictionary object");
				}
				else
				{
					if(!PyDict_Check(pValue))
					{						
						throw std::exception("Function call did not return a dictionary object");
					}
					PyObject *objRSKey = NULL;
					PyObject *objRSValue = NULL;
					Py_ssize_t RSpos = 0;
					std::stringstream ssMsg;					
					while (PyDict_Next(pValue, &RSpos, &objRSKey, &objRSValue))
					{
						ssMsg.str("");
						ssMsg << "Dictionary key from function return at position " << RSpos << " must be string";
						PyStrContainer StrRSKey(objRSKey, ssMsg);

						switch (static_cast<eReturnID>(StrRSKey))
						{
						case eRIDOutput:
						{
							if (!PyList_Check(objRSValue))
							{
								throw std::exception("Output field in return object must be list");
							}							
							ExtractPythonListAsString(lpParams->vPythonStdOut, objRSValue);
						}
						break;

						case eRIDError:
						{
							if (!PyList_Check(objRSValue))
							{
								throw std::exception("Error field in return object must be list");
							}
							ExtractPythonListAsString(lpParams->vPythonStdErr, objRSValue);
						}
						break;

						case eRIDReturn:
						{							
							if(PyDict_Check(objRSValue))
							{
								ParseReturnVariables(objRSValue, lpParams);
								//OpenFiles(lpParams);
							}
							else
							{
								// Accept no return value, but all else must be a dictionary
								if (objRSValue != Py_None)
								{
									throw std::exception("Function must return a dictionary object");
								}
							}
						}
						break;

						case eRIDUnknown:
						default:
						{
						}
						break;
						}
					}
				}
			}
			}			
		}
		catch (std::exception &ex)
		{
			std::wstring sWhat;
			sWhat = CA2W(ex.what());
			lpParams->ssMessage << _T("Python library error: ") << sWhat.c_str();
			if (PyErr_Occurred())
			{
				std::wstringstream ssPythonErr;
				FormatPythonError(ssPythonErr);
				lpParams->ssMessage << _T("\nPyErr: ") << ssPythonErr.str();
			}			
			lpParams->bSuccess = false;
		}
		catch (...)
		{
			std::wstringstream ssMessage;
			std::wstring ws;
			ws = L"Caught unhandled exception:\n";

			LPVOID lpMessageBuffer;
			::FormatMessage
			(
				FORMAT_MESSAGE_ALLOCATE_BUFFER |
				FORMAT_MESSAGE_FROM_SYSTEM,		// source and processing options 
				NULL,							// address of  message source 
				::GetLastError(),						// requested message identifier 
				MAKELANGID
				(
					LANG_NEUTRAL,
					SUBLANG_DEFAULT
				),								// language identifier for requested message 
				(LPTSTR)&lpMessageBuffer,		// address of message buffer 
				0,								// maximum size of message buffer 
				NULL							// address of array of message inserts 
			);
			ssMessage << ws << (LPTSTR)&lpMessageBuffer;
			::LocalFree(lpMessageBuffer);

			lpParams->ssMessage << ssMessage.str();
			lpParams->bSuccess = false;			
		}
		
		return lpParams->bSuccess;
	}

	bool CompileScripts(LPExecuteParams lpParams, void_allocator *alloc_inst)
	{		
		PyObject *pName = NULL;
		PyObject *pValue = NULL;
		PyObject *pModule = NULL;
		PyObject *pFunc = NULL;

		TCHAR drive[_MAX_DRIVE];
		TCHAR dir[_MAX_DIR];
		std::wstringstream ssFolder;
				
		_wsplitpath_s( lpParams->sInspectScript.c_str(), drive, _MAX_DRIVE, dir, _MAX_DIR, NULL, 0, NULL, 0);
		DebugOut(_T("CompileScripts 1\n"));

		PyRefHandler RefHandler;
		
		// Add the root scripts folder to our search path
		ssFolder << drive << dir;
		PyObject* sysPath = PySys_GetObject((char*)"path");
		PyObject* programName = RefHandler.AddRef(PyUnicode_FromWideChar(ssFolder.str().c_str(),-1));
		PyList_Append(sysPath, programName);

		DebugOut(_T("CompileScripts 2\n"));
		
		// When loading an external python script, don't add the ".py" extension
		pName = RefHandler.AddRef(PyUnicode_FromString("PythagorasInspect"));
		pModule = RefHandler.AddRef(PyImport_Import(pName));

		std::wstringstream ssMsg;
		bool bSuccess = false;

		DebugOut(_T("CompileScripts 3\n"));

		if (!pModule)
		{			
			// Failed to load the python file
			if (PyErr_Occurred())
			{
				FormatPythonError(ssMsg);
			}
			else
			{
				ssMsg << _T("Python error: Could not load inspect script.");
			}
			lpParams->ssMessage << ssMsg.str();
			return false;
		}		
		DebugOut(_T("CompileScripts 4\n"));
		pFunc = RefHandler.AddRef(PyObject_GetAttrString(pModule, "_PythagorasInspect"));
		if (!pFunc || !PyCallable_Check(pFunc))
		{
			if (PyErr_Occurred())
			{
				FormatPythonError(ssMsg);
			}
			else
			{
				ssMsg << _T("Python error: Could not load inspect function.");
			}
			lpParams->ssMessage << ssMsg.str();
			return false;
		}
		
		DebugOut(_T("CompileScripts 5\n"));
		bSuccess = true;
		int i = 0;
		std::stringstream sFormat;
		for (_itScriptFunctions itScript = lpParams->vScriptFunctions.begin(); itScript != lpParams->vScriptFunctions.end(); itScript++, i++)
		{
			// We'll pass this index along to the Python script so that we can refer to the 
			// actual item in the RelativePaths vector
			sFormat.str("");
			sFormat << i;

			DebugOut(_T("CompileScripts 6\n"));
			PyObject *pDict = NULL, *pArgs = NULL;
			pDict = RefHandler.AddRef(PyDict_New());

			std::wstring sPathToFile;
			ConvertString((*itScript)->m_PathToFile, sPathToFile);
			DebugOut(_T("CompileScripts 7\n"));

			// Set the version
			pValue = RefHandler.AddRef(PyUnicode_FromString("1.0"));
			PyDict_SetItemString(pDict, "Version", pValue);

			DebugOut(_T("CompileScripts 8\n"));
			// Load up the script
			//pValue = RefHandler.AddRef(PyUnicode_DecodeFSDefault(sPathToFile.c_str()));
			pValue = RefHandler.AddRef(PyUnicode_FromWideChar(sPathToFile.c_str(), -1));
			PyDict_SetItemString(pDict, sFormat.str().c_str(), pValue);

			DebugOut(_T("CompileScripts 9\n"));
			// Can't pass a Dictionary object directly, must create a Tuple of exactly one item and pass that...
			pArgs = RefHandler.AddRef(PyTuple_New(1));
			PyTuple_SetItem(pArgs, 0, pDict);

			DebugOut(_T("CompileScripts 10\n"));
			//Here we call the Python function and get its return value
			pValue = RefHandler.AddRef(PyObject_CallObject(pFunc, pArgs));
			if (pValue)
			{
				DebugOut(_T("CompileScripts 11\n"));
				// inspect function returns a list
				Py_ssize_t NumFunctions = PyList_Size(pValue);
				std::wstringstream ss;
				for (int iArg = 0; iArg < NumFunctions; iArg++)
				{
					DebugOut(_T("CompileScripts 12\n"));
					PyObject *pKeyValue = PyList_GetItem(pValue, iArg);
					DebugOut(_T("CompileScripts 12.1\n"));
					//PyObject* pyStr = RefHandler.AddRef(PyUnicode_AsEncodedString(pKeyValue, "utf-8", "Error ~"));
					PyObject* pyStr = RefHandler.AddRef(PyUnicode_AsEncodedString(pKeyValue, "utf-8", "ignore"));
					DebugOut(_T("CompileScripts 12.2\n"));
					if (!pyStr)
					{
						DebugOut(_T("CompileScripts 12.2 ... pyStr == NULL\n"));
					}
					else
					{
						ss << pyStr << std::endl;
						DebugOut(ss.str().c_str());
					}					
					char *sFunction = _strdup(PyBytes_AS_STRING(pyStr));
					DebugOut(_T("CompileScripts 12.3\n"));
					std::vector<std::string> sv;
					boost::split(sv, sFunction, boost::is_any_of("|"));
					DebugOut(_T("CompileScripts 12.4\n"));
					free(sFunction);

					DebugOut(_T("CompileScripts 13\n"));
					if (sv.size() >= 4)
					{
						DebugOut(_T("CompileScripts 14\n"));
						if (!_stricmp(sv.at(1).c_str(), "raise"))
						{
							// Must have been an error with compiling the file...
							// Have to convert back to unicode							
							(*itScript)->m_ErrorMessage = CA2W(sv.at(2).c_str());
						}
						else
						{
							_SharedString ssFunction(*alloc_inst);
							ssFunction.m_String = CA2W(sv.at(1).c_str());
							(*itScript)->m_Functions.push_back(ssFunction);

							_SharedString ssDescription(*alloc_inst);
							ssDescription.m_String = CA2W(sv.at(2).c_str());
							(*itScript)->m_Descriptions.push_back(ssDescription);

							_SharedString ssDisplayName(*alloc_inst);
							ssDisplayName.m_String = CA2W(sv.at(3).c_str());
							(*itScript)->m_DisplayNames.push_back(ssDisplayName);
						}
					}
					else
					{						
						(*itScript)->m_ErrorMessage = _T("From _PythagorasInspect(): Function descriptor string is not formatted properly.");
					}
				}
			}
			else
			{
				DebugOut(_T("CompileScripts 15\n"));
				if (PyErr_Occurred())
				{
					FormatPythonError(ssMsg);
				}
				else
				{
					ssMsg << _T("Python error: Could not load inspect function.");
				}
				(*itScript)->m_ErrorMessage = ssMsg.str().c_str();
			}
		}

		return bSuccess;
	}

	bool ExtractSharedMemory(LPExecuteParams lpParams, std::string &sResult)
	{
		std::wstring wResult;
		bool bReturn = ExtractSharedMemory(lpParams, wResult);
		sResult = CW2A(wResult.c_str());
		return bReturn;
	}

	bool ExtractSharedMemory(LPExecuteParams lpParams, std::wstring &sResult)
	{
		std::wstringstream sOutput;
		bool bReturn = true;
				
		try
		{
			DebugOut(_T("ExtractSharedMemory 1\n"));
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);

			_SharedExecuteOptions *ExecuteOptions = segment.find<_SharedExecuteOptions>(cExecuteOptionsString).first;
			void_allocator alloc_inst(segment.get_segment_manager());

			lpParams->alloc_inst = &alloc_inst;
			lpParams->ProcessMode = ExecuteOptions->m_ProcessMode;
			ConvertString(ExecuteOptions->m_InspectFile, lpParams->sInspectScript);
			ConvertString(ExecuteOptions->m_InjectFile, lpParams->sInjectScript);
			ConvertString(ExecuteOptions->m_TempFolder, lpParams->sTempFolder);
			ConvertString(ExecuteOptions->m_DefaultHTMLTemplate, lpParams->sDefaultHTMLTemplate);
			ConvertString(ExecuteOptions->m_LocalScriptsFolder, lpParams->sLocalScriptsFolder);
			ConvertString(ExecuteOptions->m_CommonScriptsFolder, lpParams->sCommonScriptsFolder);

			DebugOut(_T("ExtractSharedMemory 2\n"));

			switch (lpParams->ProcessMode)
			{
			case ePMCompileScript:
			{
				DebugOut(_T("ExtractSharedMemory 2.1\n"));
				_SharedCompileScriptVector *CompileScriptVector = segment.find<_SharedCompileScriptVector>(cCompileScriptString).first;				
				_SharedCompileScriptIterator it = CompileScriptVector->begin();
				std::string sFile;
				while (it != CompileScriptVector->end())
				{					
					lpParams->vScriptFunctions.push_back(&(*it));
					it++;
				}
				sResult = lpParams->ssMessage.str();
				DebugOut(_T("ExtractSharedMemory 2.14\n"));
				bReturn = CompileScripts(lpParams,&alloc_inst);
			}
			break;

			case ePMCallFunction:
			{
				DebugOut(_T("ExtractSharedMemory 2.21\n"));
				lpParams->m_bOpenFiles = ExecuteOptions->m_OpenFiles;
				ConvertString(ExecuteOptions->m_Function, lpParams->sFunction);
				ConvertString(ExecuteOptions->m_Script, lpParams->sPathToScript);
				ConvertString(ExecuteOptions->m_ScriptRegistry, lpParams->sScriptRegistry);

				DebugOut(_T("ExtractSharedMemory 2.22\n"));
				
				//Find the vector using the c-string name
				_SharedFileArray *FileNameArray = segment.find<_SharedFileArray>(cFilenameArrayString).first;
				_SharedFolderArray *FolderNameArray = segment.find<_SharedFolderArray>(cFoldernameArrayString).first;
				_SharedStringArray *StringArray = segment.find<_SharedStringArray>(cStringArrayString).first;
				lpParams->pSharedStrings = StringArray;		

				DebugOut(_T("ExtractSharedMemory 2.23\n"));
				
				_SharedStringIterator it = FileNameArray->m_Filenames.begin();
				std::wstring sFile;
				while (it != FileNameArray->m_Filenames.end())
				{
					ConvertString((*it).m_String, sFile);
					lpParams->vFilenames.push_back(sFile);
					it++;
				}

				DebugOut(_T("ExtractSharedMemory 2.24\n"));
				it = FolderNameArray->m_Foldernames.begin();
				std::wstring sFolder;
				while (it != FolderNameArray->m_Foldernames.end())
				{
					ConvertString((*it).m_String, sFolder);
					lpParams->vFoldernames.push_back(sFolder);
					it++;
				}

#if 0
				// Execute in QueryVariables mode first... the actual calling of the function will only occur after 
				// the user has chosen the properties (and maybe not even then)
				bReturn = ExecuteScript(lpParams, eEMQueryVariables);
				sResult = lpParams->ssMessage.str();
#endif
				DebugOut(_T("ExtractSharedMemory 3"));
			}
			break;
			}
		}
		catch (interprocess_exception &ex)
		{
			sOutput.str(_T(""));
			std::wstring sWhat = CA2W(ex.what());
			sOutput << _T("PythonEngine::ExtractSharedMemory()\nCaught interprocess_exception.\nCode=") << ex.get_error_code() << _T("\nMessage=") << sWhat.c_str();
			sResult = sOutput.str();
			bReturn = false;
		}
		catch (...)
		{
			LPVOID lpMessageBuffer;
			DWORD dwErr = ::GetLastError();
			::FormatMessage
			(
				FORMAT_MESSAGE_ALLOCATE_BUFFER |
				FORMAT_MESSAGE_FROM_SYSTEM,		// source and processing options 
				NULL,							// address of  message source 
				dwErr,						// requested message identifier 
				MAKELANGID
				(
					LANG_NEUTRAL,
					SUBLANG_DEFAULT
				),								// language identifier for requested message 
				(LPTSTR)&lpMessageBuffer,		// address of message buffer 
				0,								// maximum size of message buffer 
				NULL							// address of array of message inserts 
			);			
			sOutput.str(_T(""));
			sOutput << _T("PythonEngine::ExtractSharedMemory()\nCaught unknown exception.\nCode:") << dwErr << _T("\nMessage:") << (LPTSTR)&lpMessageBuffer;
			sResult = sOutput.str();
			::LocalFree(lpMessageBuffer);
			bReturn = false;
		}		
		return bReturn;
	}

	bool SetSharedMemory(LPExecuteParams lpParams, std::string &sResult)
	{
		std::wstring wResult;
		bool bReturn = SetSharedMemory(lpParams, wResult);
		sResult = CW2A(wResult.c_str());
		return bReturn;
	}

	bool SetSharedMemory(LPExecuteParams lpParams, std::wstring &sResult)
	{
		std::wstringstream sOutput;
		bool bReturn = true;

		try
		{
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);
			
			_SharedExecuteOptions *ExecuteOptions = segment.find<_SharedExecuteOptions>(cExecuteOptionsString).first;
			void_allocator alloc_inst(segment.get_segment_manager());
		
			switch (lpParams->ProcessMode)
			{
			case ePMCompileScript:
			{
				// Nothing to do
			}
			break;

			case ePMCallFunction:
			{
				ExecuteOptions->m_ExecutionTime = lpParams->dExecutionTime;
				_itString itString;				
				for (itString = lpParams->vPythonStdOut.begin(); itString != lpParams->vPythonStdOut.end(); itString++)
				{
					_SharedString Record(alloc_inst);
					Record.m_String = (*itString).c_str();
					ExecuteOptions->m_OutputVector.push_back(Record);
				}

				for (itString = lpParams->vPythonStdErr.begin(); itString != lpParams->vPythonStdErr.end(); itString++)
				{
					_SharedString Record(alloc_inst);
					Record.m_String = (*itString).c_str();
					ExecuteOptions->m_ErrorVector.push_back(Record);
				}				
			}
			break;
			}
		}
		catch (interprocess_exception &ex)
		{
			sOutput.str(_T(""));
			std::wstring sWhat = CA2W(ex.what());
			sOutput << _T("PythonEngine::SetSharedMemory()\nCaught interprocess_exception.\nCode=") << ex.get_error_code() << _T("\nMessage=") << sWhat.c_str();
			sResult = sOutput.str();
			bReturn = false;
		}
		catch (...)
		{
			LPVOID lpMessageBuffer;
			DWORD dwErr = ::GetLastError();
			::FormatMessage
			(
				FORMAT_MESSAGE_ALLOCATE_BUFFER |
				FORMAT_MESSAGE_FROM_SYSTEM,		// source and processing options 
				NULL,							// address of  message source 
				dwErr,						// requested message identifier 
				MAKELANGID
				(
					LANG_NEUTRAL,
					SUBLANG_DEFAULT
				),								// language identifier for requested message 
				(LPTSTR)&lpMessageBuffer,		// address of message buffer 
				0,								// maximum size of message buffer 
				NULL							// address of array of message inserts 
			);
			sOutput.str(_T(""));
			sOutput << _T("PythonEngine::SetSharedMemory()\nCaught unknown exception.\nCode:") << dwErr << _T("\nMessage:") << (LPTSTR)&lpMessageBuffer;
			sResult = sOutput.str();
			::LocalFree(lpMessageBuffer);
			bReturn = false;
		}
		return bReturn;
	}

	void GetAppDirectory(std::string &sDirectory)
	{	
		std::wstring s;
		GetAppDirectory(s);
		sDirectory.append(s.begin(), s.end());
	}
	
	void GetAppDirectory(std::wstring &wDirectory)
	{
		TCHAR szPath[MAX_PATH];
		TCHAR Drive[MAX_PATH], Dir[MAX_PATH];		
		GetModuleFileName(AfxGetApp()->m_hInstance, szPath, MAX_PATH);
		
		_wsplitpath_s(szPath, Drive, MAX_PATH, Dir, MAX_PATH, NULL, 0, NULL, 0);
		
		wDirectory.append(Drive);
		wDirectory.append(Dir);
	}

	void GetDocumentsDirectory(std::string &sDirectory)
	{
		std::wstring ws;
		GetDocumentsDirectory(ws);
		sDirectory.clear();
		sDirectory.append(ws.begin(), ws.end());
	}

	void GetDocumentsDirectory(std::wstring &wDirectory)
	{
		wchar_t* localMyDocuments = 0;
		// Get the "My Documents" folder
		if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Documents, 0, NULL, &localMyDocuments)))
		{
			wDirectory = localMyDocuments;
			CoTaskMemFree(static_cast<void*>(localMyDocuments));
		}
	}

	void ReplaceLocaleSpecificStrings(std::wstring &wString, FormVariables* pFormVariables)
	{
		std::string sString;		
		sString.append(wString.begin(), wString.end());
		ReplaceLocaleSpecificStrings(sString, pFormVariables);
		wString.clear();
		wString.append(sString.begin(), sString.end());
	}	
	
	void ReplaceLocaleSpecificStrings(std::string &sString, FormVariables* pFormVariables)
	{
		typedef std::pair<std::string, std::string> valReplace;
		std::vector<valReplace> svReplace;		
		std::vector<valReplace>::iterator itString;
		std::string sAppDir;
		//std::string sDocDir;
		std::stringstream ssScripts;
		std::stringstream ssTemplates;
		TCHAR WorkingDir[_MAX_PATH];
		std::string sWorkingDir;		

		//GetDocumentsDirectory(sDocDir);
		LPExecuteParams pExecuteParams = reinterpret_cast<LPExecuteParams>(pFormVariables->m_lpExecuteParams);
		//pExecuteParams->sCommonScriptsFolder
		ssScripts.str("");
		ssScripts << CW2A(pExecuteParams->sCommonScriptsFolder.c_str());
		//ssTemplates << sDocDir << "\\Pythagoras\\Templates\\";
		svReplace.push_back(std::make_pair("<common_scripts_dir>", ssScripts.str()));
		ssScripts.str("");
		ssScripts << CW2A(pExecuteParams->sLocalScriptsFolder.c_str());
		svReplace.push_back(std::make_pair("<local_scripts_dir>", ssScripts.str()));
		//svReplace.push_back(std::make_pair("<templates_dir>", ssTemplates.str()));
		
		GetAppDirectory(sAppDir);		
		svReplace.push_back(std::make_pair("<app_dir>", sAppDir));

		_tgetcwd(WorkingDir, _MAX_PATH);
		sWorkingDir = CW2A(WorkingDir);
		sWorkingDir += "\\";
		svReplace.push_back(std::make_pair("<working_dir>", sWorkingDir));

		for (itString = svReplace.begin(); itString != svReplace.end(); itString++)
		{
			boost::ireplace_all(sString, (*itString).first, (*itString).second);
		}		
	}

	void ExtractPythonListAsString(std::vector<std::wstring> &sVector, PyObject *pList)
	{
		// Assumes caller has determined whether or not the object is a list... will not throw exception here
		if (!pList)
		{
			return;
		}
		if (!PyList_Check(pList))
		{
			return;
		}
		Py_ssize_t iCount = PyList_GET_SIZE(pList);
		for (Py_ssize_t i = 0; i < iCount; i++)
		{
			PyObject *pValue = PyList_GetItem(pList, i);
			if (!pValue)
			{
				continue;
			}
			try
			{
				PyStrContainer strOutput(pValue, "");
				sVector.push_back(strOutput.w_str());
			}
			catch (std::exception &e)
			{
				// Just ignore this case, if there's nothing there so be it
			}
		}
	}

	void ExtractPythonListAsString(std::vector<std::string> &sVector, PyObject *pList)
	{
		// Assumes caller has determined whether or not the object is a list... will not throw exception here
		if (!pList)
		{
			return;
		}
		if (!PyList_Check(pList))
		{
			return;
		}
		Py_ssize_t iCount = PyList_GET_SIZE(pList);
		for (Py_ssize_t i = 0; i < iCount; i++)
		{
			PyObject *pValue = PyList_GetItem(pList, i);
			if (!pValue)
			{			
				continue;
			}
			try
			{
				PyStrContainer strOutput(pValue, "");
				sVector.push_back(strOutput.c_str());
			}
			catch (std::exception &e)
			{
				// Just ignore this case, if there's nothing there so be it
			}			
		}
	}

	void ExtractPythonListAsInteger(std::vector<int> &iVector, PyObject *pList)
	{
		// Assumes caller has determined whether or not the object is a list... will not throw exception here
		if (!pList)
		{
			return;
		}
		if (!PyList_Check(pList))
		{
			return;
		}
		Py_ssize_t iCount = PyList_GET_SIZE(pList);
		for (Py_ssize_t i = 0; i < iCount; i++)
		{
			PyObject *pValue = PyList_GetItem(pList, i);
			PyObject *objExc = NULL;
			if (!pValue)
			{
				continue;
			}
			
			Py_ssize_t stDefault = PyNumber_AsSsize_t(pValue, objExc);
			if (stDefault == PY_SSIZE_T_MIN || stDefault == PY_SSIZE_T_MAX)
			{
				continue;
			}
			iVector.push_back(stDefault);			
		}
	}

	void ExtractPythonListAsFloat(std::vector<double> &dVector, PyObject *pList)
	{
		// Assumes caller has determined whether or not the object is a list... will not throw exception here
		if (!pList)
		{
			return;
		}
		if (!PyList_Check(pList))
		{
			return;
		}
		Py_ssize_t iCount = PyList_GET_SIZE(pList);
		for (Py_ssize_t i = 0; i < iCount; i++)
		{
			PyObject *pValue = PyList_GetItem(pList, i);			
			if (!pValue)
			{
				continue;
			}
			
			PyObject *pFloat = PyNumber_Float(pValue);
			dVector.push_back(PyFloat_AS_DOUBLE(pFloat));			
		}
	}

	void ExtractPythonListAsBool(std::vector<bool> &bVector, PyObject *pList)
	{
		// Assumes caller has determined whether or not the object is a list... will not throw exception here
		if (!pList)
		{
			return;
		}
		if (!PyList_Check(pList))
		{
			return;
		}
		Py_ssize_t iCount = PyList_GET_SIZE(pList);
		for (Py_ssize_t i = 0; i < iCount; i++)
		{
			PyObject *pValue = PyList_GetItem(pList, i);
			if (!pValue)
			{
				continue;
			}
			if (PyBool_Check(pValue))
			{
				bVector.push_back(pValue == Py_False ? false : true);
			}
		}
	}

	void PyDictParser::AddBool(const wchar_t *szKey)
	{
		std::vector<bool> BoolVector;
		m_BoolVector.insert(std::make_pair(szKey, BoolVector));
	}
	void PyDictParser::AddInt(const wchar_t *szKey)
	{
		std::vector<int> IntVector;
		m_IntVector.insert(std::make_pair(szKey, IntVector));
	}
	void PyDictParser::AddFloat(const wchar_t *szKey)
	{
		std::vector<double> FloatVector;
		m_FloatVector.insert(std::make_pair(szKey, FloatVector));
	}
	void PyDictParser::AddString(const wchar_t *szKey)
	{
		std::vector<std::wstring> StringVector;
		m_StringVector.insert(std::make_pair(szKey, StringVector));
	}

	bool PyDictParser::ParseDictionary(PyObject *pDict, const char *szFunction)
	{
		std::wstring sKey;
		std::stringstream ssMessage;
		PyObject *objKey = NULL;
		PyObject *objValue = NULL;
		PyObject *objExc = NULL;
		Py_ssize_t pos = 0;

		if (!PyDict_Check(pDict))
		{
			ssMessage << "Error with argument to [" << szFunction << "]: must be dictionary object.";
			PyErr_SetString(PyExc_TypeError, ssMessage.str().c_str());
			return false;
		}

		while (PyDict_Next(pDict, &pos, &objKey, &objValue))
		{
			if (!PyUnicode_Check(objKey))
			{
				ssMessage << "Error with argument to [" << szFunction << "]: dictionary keys must be string values.";
				PyErr_SetString(PyExc_TypeError, ssMessage.str().c_str());
				return false;
			}

			ssMessage << "Error with argument to [" << szFunction << "]: dictionary key could not be converted to string.";
			try
			{
				PyStrContainer StrKey(objKey, ssMessage);
				sKey = CA2W(StrKey);
			}
			catch (...)
			{
				return false;
			}


			// Search for this key
			_itBool itBool = m_BoolVector.find(sKey.c_str());
			if (itBool != m_BoolVector.end())
			{
				ssMessage << "Error with dictionary argument to [" << szFunction << "] key [" << CW2A(sKey.c_str()) << "]: boolean value expected.";
				if (PyBool_Check(objValue))
				{
					(*itBool).second.push_back(objValue == Py_False ? false : true);
				}
				else
				{
					if (PyList_Check(objValue))
					{
						ExtractPythonListAsBool((*itBool).second, objValue);
					}
					else
					{
						PyErr_SetString(PyExc_TypeError, ssMessage.str().c_str());
						return false;
					}
				}
			}

			_itInt itInt = m_IntVector.find(sKey.c_str());
			if (itInt != m_IntVector.end())
			{
				ssMessage << "Error with dictionary argument to [" << szFunction << "] key [" << CW2A(sKey.c_str()) << "]: integer value expected.";
				if (PyNumber_Check(objValue))
				{
					Py_ssize_t stDefault = PyNumber_AsSsize_t(objValue, objExc);
					if (stDefault == PY_SSIZE_T_MIN || stDefault == PY_SSIZE_T_MAX)
					{
						PyErr_SetString(PyExc_TypeError, ssMessage.str().c_str());
						return false;
					}
					(*itInt).second.push_back(stDefault);
				}
				else
				{
					// See if it's a list
					if (PyList_Check(objValue))
					{
						ExtractPythonListAsInteger((*itInt).second, objValue);
					}
					else
					{
						PyErr_SetString(PyExc_TypeError, ssMessage.str().c_str());
						return false;
					}
				}
			}

			_itFloat itFloat = m_FloatVector.find(sKey.c_str());
			if (itFloat != m_FloatVector.end())
			{
				ssMessage << "Error with dictionary argument to [" << szFunction << "] key [" << CW2A(sKey.c_str()) << "]: float value expected.";
				if (PyNumber_Check(objValue))
				{
					PyObject *pFloat = PyNumber_Float(objValue);
					(*itFloat).second.push_back(PyFloat_AS_DOUBLE(pFloat));
				}
				else
				{
					// See if it's a list
					if (PyList_Check(objValue))
					{
						ExtractPythonListAsFloat((*itFloat).second, objValue);
					}
					else
					{
						PyErr_SetString(PyExc_TypeError, ssMessage.str().c_str());
						return false;
					}
				}
			}

			_itString itString = m_StringVector.find(sKey.c_str());
			if (itString != m_StringVector.end())
			{
				ssMessage << "Error with dictionary argument to [" << szFunction << "] key [" << CW2A(sKey.c_str()) << "]: string value expected.";
				if (PyUnicode_Check(objValue))
				{
					try
					{
						PyStrContainer StrValue(objValue, ssMessage);
						std::wstring sValue = CA2W(StrValue);
						(*itString).second.push_back(sValue);
					}
					catch (...)
					{
						return false;
					}
				}
				else
				{
					// See if it's a list
					if (PyList_Check(objValue))
					{
						ExtractPythonListAsString((*itString).second, objValue);
					}
					else
					{
						PyErr_SetString(PyExc_TypeError, ssMessage.str().c_str());
						return false;
					}
				}
			}
		}

		return true;
	}

	int PyDictParser::GetSize(const wchar_t *szKey)
	{
		// Search for this key
		_itInt itInt = m_IntVector.find(szKey);
		if (itInt != m_IntVector.end())
		{
			return (*itInt).second.size();
		}

		_itFloat itFloat = m_FloatVector.find(szKey);
		if (itFloat != m_FloatVector.end())
		{
			return (*itFloat).second.size();
		}

		_itString itString = m_StringVector.find(szKey);
		if (itString != m_StringVector.end())
		{
			return (*itString).second.size();
		}
		return 0;
	}

	bool PyDictParser::GetBoolAt(const wchar_t *szKey, int Pos, bool &b)
	{
		_itBool itBool = m_BoolVector.find(szKey);
		if (itBool != m_BoolVector.end())
		{
			if (Pos < (*itBool).second.size())
			{
				b = (*itBool).second.at(Pos);
				return true;
			}
		}
		return false;
	}

	bool PyDictParser::GetStringAt(const wchar_t *szKey, int Pos, std::wstring &s)
	{
		_itString itString = m_StringVector.find(szKey);
		if (itString != m_StringVector.end())
		{
			if (Pos < (*itString).second.size())
			{
				s = (*itString).second.at(Pos);
				return true;
			}
		}
		return false;
	}

	std::vector<std::wstring> *PyDictParser::GetStringVector(const wchar_t *szKey)
	{
		_itString itString = m_StringVector.find(szKey);
		if (itString != m_StringVector.end())
		{
			return &(*itString).second;			
		}
		return NULL;
	}

	bool PyDictParser::GetIntAt(const wchar_t *szKey, int Pos, int &i)
	{
		_itInt itInt = m_IntVector.find(szKey);
		if (itInt != m_IntVector.end())
		{
			if (Pos < (*itInt).second.size())
			{
				i = (*itInt).second.at(Pos);
				return true;
			}
		}
		return false;
	}

	bool PyDictParser::GetFloatAt(const wchar_t *szKey, int Pos, double &d)
	{
		_itFloat itFloat = m_FloatVector.find(szKey);
		if (itFloat != m_FloatVector.end())
		{
			if (Pos < (*itFloat).second.size())
			{
				d = (*itFloat).second.at(Pos);
				return true;
			}
		}
		return false;
	}

	int PyDict_SetItemWideChar(PyObject *p, const wchar_t *key, PyObject *val)
	{
		std::string sKey;
		sKey = CW2A(key);
		return PyDict_SetItemString(p, sKey.c_str(), val);
	}

	PyObject* PyObject_GetAttrWideChar(PyObject *o, const wchar_t *attr_name)
	{
		std::string sName;
		sName = CW2A(attr_name);
		return PyObject_GetAttrString(o, sName.c_str());
	}

	void OpenFiles(LPExecuteParams lpParams)
	{
		for( std::wstring sFile : lpParams->vOpenFiles )
		{
			// Check to make sure the file exists
			FILE *rp = NULL;
			errno_t err = _wfopen_s(&rp, sFile.c_str(), _T("r"));
			if (err == 0)
			{
				fclose(rp);
				SHELLEXECUTEINFO shInfo;
				memset(&shInfo, 0, sizeof(SHELLEXECUTEINFO));
				shInfo.cbSize = sizeof(SHELLEXECUTEINFO);
				shInfo.fMask |= SEE_MASK_FLAG_DDEWAIT;
				shInfo.hwnd = NULL;
				shInfo.lpVerb = _T("open");
				CString csFile(sFile.c_str());
				shInfo.lpFile = csFile;

				BOOL bEx = ShellExecuteEx(&shInfo);
			}
		}
	}

	// *****************************************************************
	// Functions available within Python

	PyObject *PythonSession::Module_Step(PyObject *self, PyObject *args)
	{
		unsigned long numsteps = 1;
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));		

		if (args)
		{
			Py_ssize_t iSize = PyTuple_Size(args);
			if (iSize > 0)
			{
				if (!PyArg_ParseTuple(args, "k", &numsteps))
				{
					return NULL;
				}
			}
		}
		
		if (lpState->hwndCallback)
		{
			::PostMessage(lpState->hwndCallback, lpState->MessageID, eMCStepIt, numsteps);
		}
		Py_INCREF(Py_None);
		return Py_None;
	}

	PyObject *PythonSession::Module_SetStepRange(PyObject *self, PyObject *args)
	{
		unsigned long range = 100;
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));

		if (args)
		{
			Py_ssize_t iSize = PyTuple_Size(args);
			if (iSize > 0)
			{
				if (!PyArg_ParseTuple(args, "k", &range))
				{
					return NULL;
				}
			}
		}

		if (lpState->hwndCallback)
		{
			::PostMessage(lpState->hwndCallback, lpState->MessageID, eMCStepRange, range);
		}
		Py_INCREF(Py_None);
		return Py_None;
	}

	PyObject *PythonSession::Module_MessageBox(PyObject *self, PyObject *args)
	{
		std::wstring wsMessage = _T("Message");
		std::wstring wsTitle = _T("Pythagoras");
		int iButtons = 0;
		int iIcon = 0;
		long lReturn = 0;
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));
		PyObject *pArg = NULL;
		PyDictParser DictParser;

		if (args)
		{
			if (PyTuple_GET_SIZE(args) > 0)
			{
				pArg = PyTuple_GetItem(args, 0);
				if (pArg)
				{
					DictParser.AddString(_T("Title"));
					DictParser.AddString(_T("Message"));
					DictParser.AddInt(_T("Buttons"));
					DictParser.AddInt(_T("Icon"));
					if (!DictParser.ParseDictionary(pArg, "messagebox"))
					{
						return NULL;
					}
					DictParser.GetStringAt(_T("Title"), 0, wsTitle);
					DictParser.GetStringAt(_T("Message"), 0, wsMessage);
					DictParser.GetIntAt(_T("Buttons"), 0, iButtons);
					DictParser.GetIntAt(_T("Icon"), 0, iIcon);
				}
			}
		}
		
		lReturn = ::MessageBox(lpState->hwndCallback, wsMessage.c_str(), wsTitle.c_str(), iButtons + iIcon);
		return PyLong_FromLong(lReturn);
	}

	PyObject *PythonSession::Module_Status(PyObject *self, PyObject *args)
	{
		const char *status;
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));

		if (!PyArg_ParseTuple(args, "s", &status))
			return NULL;
		SetCallbackString(status);
		if (lpState->hwndCallback)
		{
			::SendMessage(lpState->hwndCallback, lpState->MessageID, eMCSetStatus, 0);
		}		
		Py_INCREF(Py_None);
		return Py_None;
	}

	PyObject *PythonSession::Module_Print(PyObject *self, PyObject *args)
	{
		const char *text;
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));

		if (!PyArg_ParseTuple(args, "s", &text))
			return NULL;
		SetCallbackString(text);
		if (lpState->hwndCallback)
		{
			::SendMessage(lpState->hwndCallback, lpState->MessageID, eMCPrint, 0);
		}
		Py_INCREF(Py_None);
		return Py_None;
	}

	PyObject *PythonSession::Module_GetRegistryValues(PyObject *self, PyObject *args)
	{
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));		
		PyObject *pReturn = NULL;
		PyObject *pArg = NULL;

		if (!args)
		{
			return NULL;
		}
		pArg = PyTuple_GetItem(args, 0);
		if (!pArg)
		{
			return NULL;
		}
		if (!PyDict_Check(pArg))
		{
			return NULL;
		}

		try
		{
			ParseFormVariables(pArg, PythonSession::m_pExecuteParams, true);
			PythonSession::m_pExecuteParams->FV.SetValuesFromDefaults();			
			pReturn = PyDict_New();
			Py_INCREF(pReturn);

			// Build the return dictionary
			PythonSession::m_pExecuteParams->FV.AddToDictionary(pReturn);			
		}
		catch (std::exception &ex)
		{
			std::wstring sWhat;
			std::wstringstream ssMessage;
			sWhat = CA2W(ex.what());
			ssMessage << _T("Caught exception: ") << sWhat.c_str();
			if (PyErr_Occurred())
			{
				std::wstringstream ssPythonErr;
				FormatPythonError(ssPythonErr);
				ssMessage << _T("\nPyErr: ") << ssPythonErr.str();
			}
			::MessageBox(NULL, ssMessage.str().c_str(), _T("Python library error"), MB_OK | MB_ICONSTOP);
		}
		catch (...)
		{
			std::wstringstream ssMessage;
			std::wstring ws;
			ws = L"Caught unhandled exception:\n";

			LPVOID lpMessageBuffer;
			::FormatMessage
			(
				FORMAT_MESSAGE_ALLOCATE_BUFFER |
				FORMAT_MESSAGE_FROM_SYSTEM,		// source and processing options 
				NULL,							// address of  message source 
				::GetLastError(),						// requested message identifier 
				MAKELANGID
				(
					LANG_NEUTRAL,
					SUBLANG_DEFAULT
				),								// language identifier for requested message 
				(LPTSTR)&lpMessageBuffer,		// address of message buffer 
				0,								// maximum size of message buffer 
				NULL							// address of array of message inserts 
			);
			ssMessage << ws << (LPTSTR)&lpMessageBuffer;
			::LocalFree(lpMessageBuffer);

			::MessageBox(NULL, ssMessage.str().c_str(), _T("Python library error"), MB_OK | MB_ICONSTOP);
		}

		if (pReturn)
		{
			return pReturn;
		}
		Py_INCREF(Py_None);
		return Py_None;
	}

	PyObject *PythonSession::Module_DisplayForm(PyObject *self, PyObject *args)
	{
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));
		CPythagorasEngineDlg dlg;
		PyObject *pReturn = NULL;
		PyObject *pArg = NULL;
		
		if (!args)
		{
			return NULL;
		}
		pArg = PyTuple_GetItem(args, 0);
		if (!pArg)
		{
			return NULL;
		}
		if (!PyDict_Check(pArg))
		{
			return NULL;
		}

		try
		{
			ParseFormVariables(pArg, PythonSession::m_pExecuteParams, false);

			dlg.m_pExecuteParams = PythonSession::m_pExecuteParams;
			dlg.m_hWndOwner = lpState->hwndCallback;
			INT_PTR nResponse = dlg.DoModal();
			if (nResponse == IDOK)
			{
				pReturn = PyDict_New();
				Py_INCREF(pReturn);
				
				// Build the return dictionary
				PythonSession::m_pExecuteParams->FV.AddToDictionary(pReturn);
			}			
		}
		catch (std::exception &ex)
		{
			std::wstring sWhat;
			std::wstringstream ssMessage;
			sWhat = CA2W(ex.what());
			ssMessage << _T("Caught exception: ") << sWhat.c_str();
			if (PyErr_Occurred())
			{
				std::wstringstream ssPythonErr;
				FormatPythonError(ssPythonErr);
				ssMessage << _T("\nPyErr: ") << ssPythonErr.str();
			}
			::MessageBox(NULL, ssMessage.str().c_str(), _T("Python library error"), MB_OK | MB_ICONSTOP);
		}
		catch (...)
		{
			std::wstringstream ssMessage;
			std::wstring ws;
			ws = L"Caught unhandled exception:\n";

			LPVOID lpMessageBuffer;
			::FormatMessage
			(
				FORMAT_MESSAGE_ALLOCATE_BUFFER |
				FORMAT_MESSAGE_FROM_SYSTEM,		// source and processing options 
				NULL,							// address of  message source 
				::GetLastError(),						// requested message identifier 
				MAKELANGID
				(
					LANG_NEUTRAL,
					SUBLANG_DEFAULT
				),								// language identifier for requested message 
				(LPTSTR)&lpMessageBuffer,		// address of message buffer 
				0,								// maximum size of message buffer 
				NULL							// address of array of message inserts 
			);
			ssMessage << ws << (LPTSTR)&lpMessageBuffer;
			::LocalFree(lpMessageBuffer);

			::MessageBox(NULL, ssMessage.str().c_str(), _T("Python library error"), MB_OK | MB_ICONSTOP);
		}

		if (pReturn)
		{
			return pReturn;
		}
		Py_INCREF(Py_None);
		return Py_None;
	}

	PyObject *PythonSession::Module_OpenFile(PyObject *self, PyObject *args)
	{
		const char *filename = NULL;
		std::wstring wsFileName;
		FILE *rp = NULL;
		errno_t err = 0;
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));

		if (!PyArg_ParseTuple(args, "s", &filename))
			return NULL;

		// Check to make sure the file exists		
		wsFileName = CA2W(filename);
		err = _wfopen_s(&rp, wsFileName.c_str(), _T("r"));
		if (err == 0)
		{
			fclose(rp);
			SHELLEXECUTEINFO shInfo;
			memset(&shInfo, 0, sizeof(SHELLEXECUTEINFO));
			shInfo.cbSize = sizeof(SHELLEXECUTEINFO);
			shInfo.fMask = SEE_MASK_NOASYNC;
			shInfo.fMask = 0;
			shInfo.hwnd = NULL;
			shInfo.lpVerb = _T("open");
			CString csFile(wsFileName.c_str());
			shInfo.lpFile = csFile;

			BOOL bEx = ShellExecuteEx(&shInfo);
		}
		
		Py_INCREF(Py_None);
		return Py_None;
	}

	PyObject *PythonSession::Module_SaveFile(PyObject *self, PyObject *args)
	{
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));		
		PyObject *pReturn = NULL;
		PyObject *pArg = NULL;
		std::stringstream ss;		
		bool bReturnIfOpen = false;
		std::wstring sFilter(_T("All Files (*.*)|*.*|"));
		std::wstring sDefExt(_T(""));
		std::wstring sTitle(_T("Pythagoras"));
		std::string sReturn;
		PyDictParser DictParser;
		
		if (args)
		{			
			if(PyTuple_GET_SIZE(args) > 0)
			{
				pArg = PyTuple_GetItem(args, 0);
				if (pArg)
				{
					DictParser.AddString(_T("Title"));
					DictParser.AddString(_T("Filter"));
					DictParser.AddString(_T("DefaultExtension"));
					DictParser.AddBool(_T("ReturnIfOpen"));
					if (!DictParser.ParseDictionary(pArg, "savefile"))
					{
						return NULL;
					}
					DictParser.GetStringAt(_T("Title"), 0, sTitle);
					DictParser.GetStringAt(_T("Filter"), 0, sFilter);
					DictParser.GetStringAt(_T("DefaultExtension"), 0, sDefExt);
					DictParser.GetBoolAt(_T("ReturnIfOpen"), 0, bReturnIfOpen);
				}
			}
		}
		
		bool bSuccess = false;
		CString FileName;
		do
		{		
			CFileDialog FileDlg(FALSE, sDefExt.c_str(), NULL, OFN_OVERWRITEPROMPT, sFilter.c_str(), NULL, NULL);
			FileDlg.m_ofn.nFilterIndex = 0;

			CString str;
			int nMaxFiles = 256;
			int nBufferSz = nMaxFiles * _MAX_PATH + 1;
			FileDlg.GetOFN().lpstrFile = str.GetBuffer(nBufferSz);
			FileDlg.GetOFN().lpstrTitle = sTitle.c_str();
			if (FileDlg.DoModal() == IDOK)
			{
				FileName = FileDlg.GetPathName();
				DWORD dwAttrib = GetFileAttributes(FileName);

				if (dwAttrib != INVALID_FILE_ATTRIBUTES && !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY))
				{
					// File exists already... check to see if we can get access to the file; if not, it's probably open
					HANDLE fh;
					fh = CreateFile(FileName, GENERIC_READ, 0 /* no sharing! exclusive */, NULL, OPEN_EXISTING, 0, NULL);
					if ((fh != NULL) && (fh != INVALID_HANDLE_VALUE))
					{
						// We can access the file, good to go...
						CloseHandle(fh);
						bSuccess = true;
					}
					else
					{	
						if (bReturnIfOpen)
						{
							// Set to true so we break out of the loop... user has requested return if file is open rather than continuously loop.
							// We don't set pReturn in this case so that the function returns Py_None.
							bSuccess = true;
						}
						else
						{
							AfxMessageBox(_T("That file appears to be open, please close it first."));
						}
					}
				}
				else
				{
					// The file doesn't exist
					bSuccess = true;
				}
				if (bSuccess)
				{
					std::string sReturn;
					sReturn = CW2A(FileName);
					pReturn = PyUnicode_FromString(sReturn.c_str());
					Py_INCREF(pReturn);
				}
			}
			else
			{
				// Just break out of the loop, user has cancelled
				bSuccess = true;
			}
		} while (!bSuccess);
		
		if (pReturn)
		{
			return pReturn;
		}
		Py_INCREF(Py_None);
		return Py_None;
	}

	PyObject *PythonSession::Module_Clipboard(PyObject *self, PyObject *args)
	{
		PyObject *pArg = NULL;
		PyObject *pReturn = NULL;
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));
		PyDictParser DictParser;
		std::wstring sAction(_T("Get"));
		std::wstring sData(_T(""));
		std::stringstream ssMessage;
		PyStrContainer psAction;

		if (args)
		{
			if (PyTuple_GET_SIZE(args) > 0)
			{
				pArg = PyTuple_GetItem(args, 0);
				if (pArg)
				{
					DictParser.AddString(_T("Action"));
					DictParser.AddString(_T("Data"));					
					if (!DictParser.ParseDictionary(pArg, "clipboard"))
					{
						return NULL;
					}
					DictParser.GetStringAt(_T("Action"), 0, sAction);
					DictParser.GetStringAt(_T("Data"), 0, sData);					
				}
			}
		}
		psAction.SetString(sAction);
		switch (static_cast<eClipboardAction>(psAction))
		{
		case eCAGet:
		{			
			if( CopyFromClipboard(sData) )			
			{
				pReturn = PyUnicode_FromWideChar(sData.c_str(), -1);
				Py_INCREF(pReturn);
			}	
			else
			{
				// The default behavior will be to return 'None' to the script if the copy function failed rather
				// than generate an exception.  It may have failed because the data type on the clipboard is not text
				// (not really an exception-level event), or it may have failed because a lock could not be issued (probably
				// should be an exception).  The former case is by far the more probable one.
			}
		}
		break;

		case eCAPut:
		{
			CopyToClipboard(sData);
		}
		break;

		default:
		{
			ssMessage << "Error with argument to [clipboard]: unrecognized action.";
			PyErr_SetString(PyExc_TypeError, ssMessage.str().c_str());
			return NULL;
		}
		break;
		}
		if (pReturn)
		{
			return pReturn;
		}
		Py_INCREF(Py_None);
		return Py_None;
	}

	PyObject *PythonSession::Module_BrowseFolder(PyObject *self, PyObject *args)
	{
		PyObject *pArg = NULL;
		PyObject *pReturn = NULL;
		std::wstring wsTitle;
		std::wstring wsStartFolder;
		std::wstring sName(_T("pythagoras.browsefolder"));
		std::wstring sValue;
		const char *startfolder;
		const char *title;
		std::wstring wsFolder;
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));
		bool bRegistry = false;
		
		if (!PyArg_ParseTuple(args, "ss", &startfolder, &title))
			return NULL;
		
		wsStartFolder = CA2W(startfolder);
		wsTitle = CA2W(title);

		if (!_tcsicmp(wsStartFolder.c_str(), _T("Registry")))
		{
			bRegistry = true;
			ReadRegistryString(lpState->lpExecuteParams, sName, sValue);
			wsStartFolder = sValue;
		}
				
		if (BrowseForFolder(lpState->hwndCallback, wsStartFolder.c_str(), wsTitle.c_str(), wsFolder))
		{
			if (bRegistry)
			{
				WriteRegistryString(lpState->lpExecuteParams, sName, wsFolder);
			}
			pReturn = PyUnicode_FromWideChar(wsFolder.c_str(), wsFolder.size());
			Py_INCREF(pReturn);
		}
		if (pReturn)
		{
			return pReturn;
		}
		Py_INCREF(Py_None);
		return Py_None;
	}

	/*static */PyObject *PythonSession::Module_AddToRegexList(PyObject *self, PyObject *args)
	{
		PyObject *pArg = NULL;
		PyObject *pReturn = NULL;
		LPEngineStateStruct lpState = reinterpret_cast<LPEngineStateStruct>(PyModule_GetState(self));
		PyDictParser DictParser;
		std::wstring sAction(_T("Add"));
		std::wstring sData(_T(""));
		std::wstring sResult(_T(""));
		std::stringstream ssMessage;
		PyStrContainer psAction;

		if (args)
		{
			if (PyTuple_GET_SIZE(args) > 0)
			{
				pArg = PyTuple_GetItem(args, 0);
				if (pArg)
				{
					DictParser.AddString(_T("Action"));
					DictParser.AddString(_T("Data"));
					if (!DictParser.ParseDictionary(pArg, "addregexlist"))
					{
						return NULL;
					}
					DictParser.GetStringAt(_T("Action"), 0, sAction);

					_StringVector *pDataVector = DictParser.GetStringVector(_T("Data"));
					if (!SetSharedStringArray(pDataVector, sResult))
					{
						PyErr_SetString(PyExc_TypeError, sResult.c_str());
						return NULL;
					}
					else
					{
						psAction.SetString(sAction);
						eModuleCallbackRegexAction eAction = static_cast<eModuleCallbackRegexAction>(psAction);
						if (lpState->hwndCallback)
						{
							::SendMessage(lpState->hwndCallback, lpState->MessageID, eMCAddToRegex, eAction);
						}
					}
				}
			}
		}
				
		
		if (pReturn)
		{
			return pReturn;
		}
		Py_INCREF(Py_None);
		return Py_None;
	}
	
	/*static */PyObject *PythonSession::Module_AddToProcessList(PyObject *self, PyObject *args)
	{
		Py_INCREF(Py_None);
		return Py_None;
	}

	/* static */void PythonSession::GetBaseSubKey(LPExecuteParams lpExecuteParams, std::wstring &sRegistrySubKey)
	{
		std::wstringstream ssSubKey;
		std::vector<std::wstring> sv;
		boost::split(sv, lpExecuteParams->sScriptRegistry, boost::is_any_of("."));
		ssSubKey << _T("Software\\3M\\Pythagoras\\ScriptSettings\\");
		for (std::wstring ws : sv)
		{
			ssSubKey << ws.c_str() << _T("\\");
		}
		ssSubKey << lpExecuteParams->sFunction.c_str() << _T("\\");
		sRegistrySubKey = ssSubKey.str().c_str();
	}

	/* static */void PythonSession::ReadRegistryString(LPExecuteParams lpExecuteParams, std::wstring &sName, std::wstring &sValue )
	{
		CRegistryHelper rhHelper;
		std::wstring sRegistrySubKey;
		GetBaseSubKey(lpExecuteParams, sRegistrySubKey);
		rhHelper.SetMainKey(HKEY_CURRENT_USER);
		rhHelper.SetBaseSubKey(sRegistrySubKey.c_str());
		rhHelper.AddItem(&sValue, sValue.c_str(), sName.c_str(), _T(""));
		rhHelper.ReadRegistry();
		rhHelper.WriteRegistry();
	}

	/* static */void PythonSession::WriteRegistryString(LPExecuteParams lpExecuteParams, std::wstring &sName, std::wstring &sValue)
	{
		CRegistryHelper rhHelper;
		std::wstring sRegistrySubKey;
		GetBaseSubKey(lpExecuteParams, sRegistrySubKey);
		rhHelper.SetMainKey(HKEY_CURRENT_USER);
		rhHelper.SetBaseSubKey(sRegistrySubKey.c_str());
		rhHelper.AddItem(&sValue, sValue.c_str(), sName.c_str(), _T(""));		
		rhHelper.WriteRegistry();
	}

	void PythonSession::SetCallbackString(const char *szText)
	{
		std::wstring ws;
		ws = CA2W(szText);
		SetCallbackString(ws.c_str());
	}
	void PythonSession::SetCallbackString(LPCTSTR szText)
	{
		try
		{
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);
			_SharedScriptCallback *Callback = segment.find<_SharedScriptCallback>(cScriptCallbackString).first;
			void_allocator alloc_inst(segment.get_segment_manager());
			if (Callback)
			{
				Callback->m_String = szText;
			}
		}
		catch (interprocess_exception &ex)
		{
		}
		catch (...)
		{		
		}		
	}

	static PyMethodDef PythagorasMethods[] = {
		{ "step",  PythonSession::Module_Step, METH_VARARGS, "Increment the GUI progress bar." },
		{ "steprange",  PythonSession::Module_SetStepRange, METH_VARARGS, "Set the GUI progress bar range." },
		{ "messagebox",  PythonSession::Module_MessageBox, METH_VARARGS, "Display a messagebox." },
		{ "status",  PythonSession::Module_Status, METH_VARARGS, "Set the text of the status bar." },
		{ "print",  PythonSession::Module_Print, METH_VARARGS, "Print text to the standard output." },
		{ "displayform",  PythonSession::Module_DisplayForm, METH_VARARGS, "Acquire input from user." },
		{ "getregistryvalues",  PythonSession::Module_GetRegistryValues, METH_VARARGS, "Retrieve variable values from registry without displaying the form." },
		{ "openfile",  PythonSession::Module_OpenFile, METH_VARARGS, "Open file." },
		{ "savefile",  PythonSession::Module_SaveFile, METH_VARARGS, "Save file." },
		{ "clipboard",  PythonSession::Module_Clipboard, METH_VARARGS, "Interact with the clipboard." },
		{ "browsefolder",  PythonSession::Module_BrowseFolder, METH_VARARGS, "Browse for a folder." },
		{ "addregexlist",  PythonSession::Module_AddToRegexList, METH_VARARGS, "Adds list of strings to regex results list." },
		{ "addprocesslist",  PythonSession::Module_AddToProcessList, METH_VARARGS, "Adds list of strings to process list." },
		{ NULL, NULL, 0, NULL }        /* Sentinel */
	};

	static struct PyModuleDef PythagorasModule = {
		PyModuleDef_HEAD_INIT,
		"Pythagoras",   /* name of module */
		"Provides a set of functions that enable interaction with the Pythagoras client application.", /* module documentation, may be NULL */
		sizeof(PythonSession::EngineStateStruct),       /* size of per-interpreter state of the module,
										 or -1 if the module keeps state in global variables. */
		PythagorasMethods
	};
	
	LPExecuteParams PythonSession::m_pExecuteParams = NULL;
	
	PyMODINIT_FUNC PyInit_Module(void)
	{
		PyObject *pModule = PyModule_Create(&PythagorasModule);
		PythonSession::LPEngineStateStruct lpState = reinterpret_cast<PythonSession::LPEngineStateStruct>(PyModule_GetState(pModule));
		try
		{
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);
			_SharedExecuteOptions *ExecuteOptions = segment.find<_SharedExecuteOptions>(cExecuteOptionsString).first;
			lpState->hwndCallback = reinterpret_cast<HWND>(ExecuteOptions->m_lHWND);
			lpState->MessageID = (WPARAM)ExecuteOptions->m_lMessageID;
			lpState->lpExecuteParams = PythonSession::m_pExecuteParams;
		}
		catch (interprocess_exception &ex)
		{
		}
		catch (...)
		{		
		}
		return pModule;
	}

	void DebugOut(const wchar_t *szText, bool bNewLine)
	{
		FILE *wp = NULL;
		TCHAR lpTempPathBuffer[MAX_PATH];
		std::wstringstream wsFormat;

#ifdef _DEBUG
		DWORD dwRetVal = GetTempPath(MAX_PATH,          // length of the buffer
			lpTempPathBuffer); // buffer for path 
		if (dwRetVal <= MAX_PATH && (dwRetVal != 0))
		{
			wsFormat << lpTempPathBuffer << _T("PythagorasEngine.txt");
			errno_t err = _wfopen_s(&wp, wsFormat.str().c_str(), _T("a"));
			if (err == 0)
			{
				fwprintf(wp, _T("%s%s"), szText, bNewLine ? _T("\n") : _T(""));
				fclose(wp);
			}
		}		
#endif
#if 0
		std::wstringstream wsTempFile;
		std::wstring wsApp;
		GetAppDirectory(wsApp);
		wsTempFile << wsApp.c_str() << _T("PythagorasEngine.txt");
		errno_t err = _wfopen_s(&wp, wsTempFile.str().c_str(), _T("a"));
		if (err == 0)
		{
			fwprintf(wp, _T("%s%s"), szText, bNewLine ? _T("\n") : _T(""));
			fclose(wp);
		}
#endif
	}

	bool SetSharedStringArray(_StringVector &vStrings, std::wstring &sResult)
	{
		return SetSharedStringArray(&vStrings, sResult);
	}

	bool SetSharedStringArray(_StringVector *pStrings, std::wstring &sResult)
	{
		bool bReturn = true;
		std::wstringstream sOutput;
		try
		{
			wmanaged_shared_memory segment(open_only, cSharedMemoryString);
			_SharedExecuteOptions *ExecuteOptions = segment.find<_SharedExecuteOptions>(cExecuteOptionsString).first;
			void_allocator alloc_inst(segment.get_segment_manager());

			//Find the vector using the c-string name
			_SharedStringArray *StringArray = segment.find<_SharedStringArray>(cStringArrayString).first;

			if (StringArray->m_Strings.size() > 0)
			{
				StringArray->m_Strings.erase(StringArray->m_Strings.begin(), StringArray->m_Strings.end());
			}
			for (std::wstring &ws : *pStrings)
			{				
				_SharedString ssData(alloc_inst);
				ssData.m_String = ws.c_str();
				StringArray->m_Strings.push_back(ssData);
			}		
		}
		catch (interprocess_exception &ex)
		{
			sOutput.str(_T(""));
			sOutput << _T("Caught interprocess_exception.\nCode=") << ex.get_error_code() << _T("\nMessage=") << CA2W(ex.what());
			sResult = sOutput.str();
			bReturn = false;
		}
		catch (...)
		{
			LPVOID lpMessageBuffer;
			DWORD dwErr = ::GetLastError();
			::FormatMessage
			(
				FORMAT_MESSAGE_ALLOCATE_BUFFER |
				FORMAT_MESSAGE_FROM_SYSTEM,		// source and processing options 
				NULL,							// address of  message source 
				dwErr,						// requested message identifier 
				MAKELANGID
				(
					LANG_NEUTRAL,
					SUBLANG_DEFAULT
				),								// language identifier for requested message 
				(LPTSTR)&lpMessageBuffer,		// address of message buffer 
				0,								// maximum size of message buffer 
				NULL							// address of array of message inserts 
			);
			sOutput.str(_T(""));
			sOutput << _T("Caught unknown exception.\nCode:") << dwErr << _T("\nMessage:") << (LPTSTR)&lpMessageBuffer;
			sResult = sOutput.str();
			::LocalFree(lpMessageBuffer);
			bReturn = false;
		}
		return bReturn;
	}

	void PyErr_SetString(PyObject *type, const wchar_t *message)
	{
		std::string sMessage;
		sMessage = CW2A(message);
		PyErr_SetString(type, sMessage.c_str());
	}
}

