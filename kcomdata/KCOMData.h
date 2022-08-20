#if !defined(AFX_KCOMDATA_H__029288D3_2FFF_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_KCOMDATA_H__029288D3_2FFF_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMData.h : main header file for KCOMDATA.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols
#include "KCOMData_i.h"

/////////////////////////////////////////////////////////////////////////////
// CKCOMDataApp : See KCOMData.cpp for implementation.

class CKCOMDataApp : public COleControlModule
{
public:
	
	// the common registered message
	static int m_nBringMsg;

	BOOL InitInstance();
	int ExitInstance();
private:
	BOOL InitATL();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMDATA_H__029288D3_2FFF_11D3_B446_0080C8F18522__INCLUDED)
