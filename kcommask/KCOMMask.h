#if !defined(AFX_KCOMMASK_H__D1728E2B_4985_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMMASK_H__D1728E2B_4985_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMMask.h : main header file for KCOMMASK.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskApp : See KCOMMask.cpp for implementation.

class CKCOMMaskApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMMASK_H__D1728E2B_4985_11D3_A7FE_0080C8763FA4__INCLUDED)
