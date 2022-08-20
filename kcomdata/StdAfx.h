#if !defined(AFX_STDAFX_H__029288D1_2FFF_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_STDAFX_H__029288D1_2FFF_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afxctl.h>         // MFC support for ActiveX Controls
#include <afxext.h>         // MFC extensions
#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Comon Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

// Delete the two includes below if you do not wish to use the MFC
//  database classes
#include <afxdb.h>			// MFC database classes
#include <afxdao.h>			// MFC DAO database classes

#include <afxtempl.h>
#include <msdatsrc.h>

#include <config\ogado.h>
#include <gxall.h>

#define _ATL_APARTMENT_THREADED
#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>
#include <atldbcli.h>

// the mouse hook callback proc
__declspec(dllexport) LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam);

// assign value to the global variables
__declspec(dllexport) void SetHost(HWND hWndMain, HWND hWndDropDown);

extern HWND m_hWndMain, m_hWndDropDown;

#define STOCKPROP_BACKCOLOR     0x00000001
#define STOCKPROP_CAPTION       0x00000002
#define STOCKPROP_FONT          0x00000004
#define STOCKPROP_FORECOLOR     0x00000008
#define STOCKPROP_TEXT          0x00000010
#define STOCKPROP_BORDERSTYLE   0x00000020
#define STOCKPROP_ENABLED       0x00000040
#define STOCKPROP_APPEARANCE    0x00000080

struct BindData
{
	// bookmark
	BYTE bookmark[4];

	// string data
	TCHAR strData[255][256];

	// status data
	DWORD dwStatus[255];
};

struct RowData
{
	TCHAR strData[255][256];
	DWORD dwStatus[255];
};

struct ColumnData
{
	// column's name
	CString strName;

	// the ordinal in raw columns info
	int nColumn;

	// the data type in raw columns info
	VARTYPE vt;
};

#define DEFAULTCOLOR RGB(0, 0, 1)

struct ColumnProp
{
	ColumnProp()
	{
		clrBack = clrFore = clrHeadBack = clrHeadFore = DEFAULTCOLOR;
		bForceLock = FALSE;
		bForceNoNullable = FALSE;
		bNullable = TRUE;
		nDataType = 0;
		bPromptInclude = TRUE;
		nDataField = -1;
		nComboBoxStyle = 0;
	};

	long nAlignment;
	BOOL bAllowSizing;
	CString strCaption;
	long nCaptionAlignment;
	long nCase;
	CString strDataField;
	int nDataField;
	COLORREF clrBack, clrFore, clrHeadBack, clrHeadFore;
	long nFieldLen;
	BOOL bLocked;
	long nDataType;
	CString strMask;
	CString strName;
	BOOL bNullable;
	CString strPromptChar;
	BOOL bPromptInclude;
	long nStyle;
	BOOL bVisible;
	ROWCOL nWidth;
	long nComboBoxStyle;

	// the col index in the grid
	ROWCOL nColIndex;
	
	// whether force this column to be locked because it is readonly in datasource
	BOOL bForceLock;

	// whether force this column to be not nullable because it is needed by datasource
	BOOL bForceNoNullable;
	
	// the choice list data of the combobox style cell
	CString strChoiceList;
};

static VARTYPE arrVt[8] = {VT_BSTR, VT_BOOL, VT_I1, VT_I2, VT_I4, VT_R4, VT_CY, VT_DATE};

// the max number of the rows fetch once
#define FETCHONCE 300

// the function used in mfc source file, stolen by me :)
short AFXAPI _AfxShiftState();
void AFXAPI _AfxToggleBorderStyle(CWnd* pWnd);
void AFXAPI _AfxToggleAppearance(CWnd* pWnd);

// the function to format the ole db error code into string
void ReportError();

// the default funt data used to initialize the head font property, stolen by me from mfc source codes :)
AFX_STATIC_DATA const FONTDESC _KCOMFontDescDefault =
	{ sizeof(FONTDESC), OLESTR("MS Sans Serif"), FONTSIZE(8), FW_NORMAL,
	  DEFAULT_CHARSET, FALSE, FALSE, FALSE };

// the class used friend class defination, so it's protected "m_bAutoDelete" member can be access from these classes.
class __CMemFile : public CMemFile
{
	friend class CKCOMRichGridCtrl;
	friend class CKCOMRichComboCtrl;
	friend class CKCOMRichDropDownCtrl;
};

// the struction used to exchange information between host control and dropdown control
struct BringInfo
{
	long wParam;
	long lParam;
	HWND hWnd;
	TCHAR strText[256];
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__029288D1_2FFF_11D3_B446_0080C8F18522__INCLUDED_)
