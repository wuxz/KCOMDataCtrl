#if !defined(AFX_KCOMEDITCONTROL_H__E2454882_1A7A_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMEDITCONTROL_H__E2454882_1A7A_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// KCOMEditControl.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CKCOMEditControl window

class CKCOMEditControl : public CGXEditControl
{
// Construction
public:
	CKCOMEditControl(CGXGridCore* pGrid, UINT nID);

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMEditControl)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual BOOL GetValue(CString& strResult);
	virtual ~CKCOMEditControl();

	// Generated message map functions
protected:
//	virtual void DrawBackground(CDC* pDC, const CRect& rect, const CGXStyle& style, BOOL bErase);
//	virtual CRect GetCellRect(ROWCOL nRow, ROWCOL nCol, LPRECT rectItem, const CGXStyle* pStyle);
	//{{AFX_MSG(CKCOMEditControl)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMEDITCONTROL_H__E2454882_1A7A_11D3_A7FE_0080C8763FA4__INCLUDED_)
