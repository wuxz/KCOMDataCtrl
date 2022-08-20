#if !defined(AFX_KCOMDBGRID_H__AC21264C_FBE2_11D2_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMDBGRID_H__AC21264C_FBE2_11D2_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMDBGrid.h : main header file for KCOMDBGRID.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols
#include "KCOMDBGrid_i.h"

/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridApp : See KCOMDBGrid.cpp for implementation.

class CKCOMDBGridApp : public COleControlModule
{
public:
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

#endif // !defined(AFX_KCOMDBGRID_H__AC21264C_FBE2_11D2_A7FE_0080C8763FA4__INCLUDED)
