
// PythagorasEngineDlg.cpp : implementation file
//

#include "stdafx.h"
#include "PythagorasEngine.h"
#include "PythagorasEngineDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CPythagorasEngineDlg dialog

BEGIN_DHTML_EVENT_MAP(CPythagorasEngineDlg)
	DHTML_EVENT_ONCLICK(_T("CPythagorasEngineDlg_ButtonOK"), OnButtonOK)
	DHTML_EVENT_ONCLICK(_T("CPythagorasEngineDlg_ButtonCancel"), OnButtonCancel)
END_DHTML_EVENT_MAP()


CPythagorasEngineDlg::CPythagorasEngineDlg(CWnd* pParent /*=NULL*/)
	: CDHtmlDialog(IDD_PYTHAGORASENGINE_DIALOG, IDR_HTML_PYTHAGORASENGINE_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CPythagorasEngineDlg::DoDataExchange(CDataExchange* pDX)
{
	if (pDX->m_bSaveAndValidate)
	{
		IHTMLDocument2 *pDoc = NULL;
		if (CDHtmlDialog::GetDHtmlDocument(&pDoc) == S_OK)
		{
			IHTMLElementCollection *pCollection;			
			if (pDoc->get_all(&pCollection) == S_OK)
			{				
				m_pExecuteParams->FV.DoDataExchange(pCollection);
				pCollection->Release();
			}			
			pDoc->close();
			pDoc->Release();			
		}
	}
	CDHtmlDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CPythagorasEngineDlg, CDHtmlDialog)
END_MESSAGE_MAP()


// CPythagorasEngineDlg message handlers

BOOL CPythagorasEngineDlg::OnInitDialog()
{
	CDHtmlDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	SetWindowPos(NULL, 0, 0, m_pExecuteParams->FV.m_uiWidth, m_pExecuteParams->FV.m_uiHeight, SWP_NOMOVE | SWP_NOZORDER);
	SetForegroundWindow();

	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CPythagorasEngineDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDHtmlDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CPythagorasEngineDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

HRESULT CPythagorasEngineDlg::OnButtonOK(IHTMLElement* /*pElement*/)
{	
	OnOK();
	return S_OK;
}

HRESULT CPythagorasEngineDlg::OnButtonCancel(IHTMLElement* /*pElement*/)
{
	OnCancel();
	return S_OK;
}


void CPythagorasEngineDlg::OnDocumentComplete(LPDISPATCH pDisp, LPCTSTR szUrl)
{	
	IHTMLDocument2 *pDoc = NULL;
	IHTMLElement *pBody = NULL;
	try
	{	
		if (CDHtmlDialog::GetDHtmlDocument(&pDoc) == S_OK)
		{
#if 1			
			pDoc->get_body(&pBody);
			CString str = m_pExecuteParams->FV.BuildHTMLForm();
			BSTR bstr = str.AllocSysString();
			pBody->put_innerHTML(bstr);
			::SysFreeString(bstr);
			pBody->Release();
#endif

			IHTMLElementCollection *pCollection = NULL;
			HRESULT hr = pDoc->get_all(&pCollection);
			if (!FAILED(hr))
			{
				IDispatch *pDisp = NULL;
				IHTMLBodyElement *pBody = NULL;
				IHTMLElement *pBodyElem = NULL;
				variant_t vtID;

				::VariantInit(&vtID);
				vtID.vt = VT_I4;
				vtID.lVal = 3;
				pCollection->item(vtID, vtID, &pDisp);
				if (pDisp)
				{
					pDisp->QueryInterface(IID_IHTMLBodyElement, (void **)&pBody);
					if (pBody)
					{
						variant_t vtBGColor;
						::VariantInit(&vtBGColor);
						if (m_pExecuteParams->FV.m_sBGColor.size() != 0)
						{
							std::string s;
							s = CW2A(m_pExecuteParams->FV.m_sBGColor.c_str());
							vtBGColor.SetString(s.c_str());
						}
						else
						{
							vtBGColor.SetString("lightgrey");
						}
						pBody->put_bgColor(vtBGColor);
						pBody->Release();
					}
#if 0
					pDisp->QueryInterface(IID_IHTMLElement, (void **)&pBodyElem);
					if (pBodyElem)
					{
						CString str = m_pExecuteParams->FV.BuildHTMLForm(ssTemplate.str().c_str());
						BSTR bstr = str.AllocSysString();
						pBodyElem->put_innerHTML(bstr);
						pBodyElem->Release();
					}
#endif
					pDisp->Release();
				}
				pCollection->Release();
			}

		pDoc->close();
		pDoc->Release();
		}
	}
	catch (std::exception &ex)
	{
		std::wstringstream ssMessage;
		std::wstring ws;
		ws = L"Caught exception: ";
		ssMessage << ws << ex.what();
		::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);
		OnCancel();
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
		
		::MessageBox(NULL, ssMessage.str().c_str(), NULL, MB_ICONSTOP);
		OnCancel();
	}

	CDHtmlDialog::OnDocumentComplete(pDisp, szUrl);
}


void CPythagorasEngineDlg::OnBeforeNavigate(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	// TODO: Add your specialized code here and/or call the base class	
	CDHtmlDialog::OnBeforeNavigate(pDisp, szUrl);
}
