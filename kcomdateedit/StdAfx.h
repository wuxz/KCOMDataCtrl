#if !defined(AFX_STDAFX_H__F6D0E40B_2D34_11D2_9A46_0080C8763FA4__INCLUDED_)
#define AFX_STDAFX_H__F6D0E40B_2D34_11D2_9A46_0080C8763FA4__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afxctl.h>         // MFC support for ActiveX Controls

#include <afxcmn.h>

#define STOCKPROP_BACKCOLOR     0x00000001
#define STOCKPROP_CAPTION       0x00000002
#define STOCKPROP_FONT          0x00000004
#define STOCKPROP_FORECOLOR     0x00000008
#define STOCKPROP_TEXT          0x00000010
#define STOCKPROP_BORDERSTYLE   0x00000020
#define STOCKPROP_ENABLED       0x00000040
#define STOCKPROP_APPEARANCE    0x00000080

short AFXAPI _AfxShiftState();
void AFXAPI _AfxToggleBorderStyle(CWnd* pWnd);
void AFXAPI _AfxToggleAppearance(CWnd* pWnd);

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__F6D0E40B_2D34_11D2_9A46_0080C8763FA4__INCLUDED_)
