// KCOMDBGrid.cpp : Implementation of CKCOMDBGridApp and DLL registration.

#include "stdafx.h"
#include "KCOMDBGrid.h"
#include <initguid.h>
#include "KCOMDBGrid_i.c"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


CKCOMDBGridApp NEAR theApp;

const GUID CDECL BASED_CODE _tlid =
		{ 0xac212643, 0xfbe2, 0x11d2, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;


////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridApp::InitInstance - DLL initialization

BOOL CKCOMDBGridApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();
	if (!InitATL())
		return FALSE;


	if (bInit)
	{
		// TODO: Add your own module initialization code here.
	}

	return bInit;
}


////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridApp::ExitInstance - DLL termination

int CKCOMDBGridApp::ExitInstance()
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

BOOL CKCOMDBGridApp::InitATL()
{
	_Module.Init(ObjectMap, AfxGetInstanceHandle());
	return TRUE;

}
