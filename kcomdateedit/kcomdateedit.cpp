// editbox.cpp : Implementation of CKCOMDateEditApp and DLL registration.

#include "stdafx.h"
#include "KCOMDateEdit.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


CKCOMDateEditApp NEAR theApp;

const GUID CDECL BASED_CODE _tlid =
		{ 0xf6d0e404, 0x2d34, 0x11d2, { 0x9a, 0x46, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;


////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditApp::InitInstance - DLL initialization

BOOL CKCOMDateEditApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();
	
	if (bInit)
	{
		// TODO: Add your own module initialization code here.
	}

	return bInit;
}


////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditApp::ExitInstance - DLL termination

int CKCOMDateEditApp::ExitInstance()
{
	// TODO: Add your own module termination code here.

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

	return NOERROR;
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
