// KCOMData.cpp : Implementation of CKCOMDataApp and DLL registration.

#include "stdafx.h"
#include "KCOMData.h"
#include <initguid.h>
#include "KCOMData_i.c"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

CKCOMDataApp NEAR theApp;
HWND m_hWndMain, m_hWndDropDown;

const GUID CDECL BASED_CODE _tlid =
		{ 0x29288c2, 0x2fff, 0x11d3, { 0xb4, 0x46, 0, 0x80, 0xc8, 0xf1, 0x85, 0x22 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;

int CKCOMDataApp::m_nBringMsg = RegisterWindowMessage(_T("Bring KCOM RichDropDown"));

////////////////////////////////////////////////////////////////////////////
// CKCOMDataApp::InitInstance - DLL initialization

BOOL CKCOMDataApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();
	if (!InitATL())
		return FALSE;


	if (bInit)
	{
		GXInit();
	}

	return bInit;
}


////////////////////////////////////////////////////////////////////////////
// CKCOMDataApp::ExitInstance - DLL termination

int CKCOMDataApp::ExitInstance()
{
	// TODO: Add your own module termination code here.

	_Module.Term();


	return COleControlModule::ExitInstance();
}


/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return _Module.RegisterServer(TRUE);

	return NOERROR;
}


/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	_Module.UnregisterServer(TRUE); //TRUE indicates that typelib is unreg'd

	return NOERROR;
}

	
CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

STDAPI DllCanUnloadNow(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return (AfxDllCanUnloadNow()==S_OK && _Module.GetLockCount()==0) ? S_OK : S_FALSE;
}
/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	if(AfxDllGetClassObject(rclsid, riid, ppv) == S_OK)
		return S_OK;
	return _Module.GetClassObject(rclsid, riid, ppv);
}

BOOL CKCOMDataApp::InitATL()
{
	_Module.Init(ObjectMap, AfxGetInstanceHandle());
	return TRUE;

}

void ReportError()
{
	IErrorInfo *pErrorInfo = NULL;
	CString strError;
	BSTR bstrDescription = NULL;
	HINSTANCE hInstResources = AfxGetResourceHandle();

	// go and get the error information
	//
	HRESULT hr = GetErrorInfo(0, &pErrorInfo);
	if (FAILED(hr) || !pErrorInfo)
		strError.LoadString(IDS_ERROR_GENERAL);
	else
	{
		// get the source and the description
		//
		hr = pErrorInfo->GetDescription(&bstrDescription);
		pErrorInfo->Release();

		USES_CONVERSION;
		if (SUCCEEDED(hr))
		{
			strError = W2T(bstrDescription);
			SysFreeString(bstrDescription);
		}
	}

	// show the error!
	//
	AfxMessageBox(strError, MB_ICONEXCLAMATION|MB_OK|MB_TASKMODAL);
}

#ifdef _AFXDLL

short AFXAPI _AfxShiftState()
{
	BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
	BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
	BOOL bAlt   = (GetKeyState(VK_MENU) < 0);

	return (short)(bShift + (bCtrl << 1) + (bAlt << 2));
}

void AFXAPI _AfxToggleBorderStyle(CWnd* pWnd)
{
	if (pWnd->m_hWnd != NULL)
	{
		// toggle border style and force redraw of border
		::SetWindowLong(pWnd->m_hWnd, GWL_STYLE, pWnd->GetStyle() ^ WS_BORDER);
		::SetWindowPos(pWnd->m_hWnd, NULL, 0, 0, 0, 0,
			SWP_DRAWFRAME | SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
	}
}

void AFXAPI _AfxToggleAppearance(CWnd* pWnd)
{
	if (pWnd->m_hWnd != NULL)
	{
		// toggle border style and force redraw of border
		::SetWindowLong(pWnd->m_hWnd, GWL_EXSTYLE, pWnd->GetExStyle() ^
			WS_EX_CLIENTEDGE);
		::SetWindowPos(pWnd->m_hWnd, NULL, 0, 0, 0, 0,
			SWP_DRAWFRAME | SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
	}
}

#endif

__declspec(dllexport) LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (nCode >=0 && nCode != HC_NOREMOVE)
	{	
		if (wParam != WM_MOUSEMOVE)
		{
			TRACE("hwnd is %x", ((MOUSEHOOKSTRUCT*)lParam)->hwnd);
			TRACE("message is %x \n", wParam);
		}
		// if the button clicked outside the dropdown window, close it
		if ((wParam == WM_LBUTTONDOWN || wParam == WM_NCLBUTTONDOWN) && ((MOUSEHOOKSTRUCT*)lParam)->hwnd != m_hWndDropDown)
		{
			::SendMessage(m_hWndMain, WM_KEYDOWN, VK_ESCAPE, 1);
		}
	}

	return 0;
}

__declspec(dllexport) void SetHost(HWND hWndMain, HWND hWndDropDown)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	
	m_hWndMain = hWndMain;
	m_hWndDropDown = hWndDropDown;
}
