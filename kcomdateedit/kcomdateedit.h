#if !defined(AFX_EDITBOX_H__F6D0E40D_2D34_11D2_9A46_0080C8763FA4__INCLUDED_)
#define AFX_EDITBOX_H__F6D0E40D_2D34_11D2_9A46_0080C8763FA4__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

// editbox.h : main header file for EDITBOX.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols

// predefined color
#define CWHITE RGB(255, 255, 255)
#define CGRAY RGB(200, 200, 200)
#define CRED RGB(255, 0, 0)
#define CBLUE RGB(0, 0, 255)
#define CMAGENTA RGB(0, 255, 255)
#define CBLACK RGB(0, 0, 0)

/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditApp : See editbox.cpp for implementation.

class CKCOMDateEditApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EDITBOX_H__F6D0E40D_2D34_11D2_9A46_0080C8763FA4__INCLUDED)
