#if !defined(AFX_GRIDINPUT_H__CBD24583_1764_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_GRIDINPUT_H__CBD24583_1764_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GridInput.h : header file
//

// the grid for inputing the list data for combobox style column

/////////////////////////////////////////////////////////////////////////////
// CGridInput window

class CGridInput : public CGXBrowserWnd
{
// Construction
public:
	CGridInput();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGridInput)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CGridInput();

	// Generated message map functions
protected:
	//{{AFX_MSG(CGridInput)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GRIDINPUT_H__CBD24583_1764_11D3_A7FE_0080C8763FA4__INCLUDED_)
