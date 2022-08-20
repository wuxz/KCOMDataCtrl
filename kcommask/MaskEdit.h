#if !defined(AFX_MASKEDIT_H__D1728E37_4985_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_MASKEDIT_H__D1728E37_4985_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// MaskEdit.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CMaskEdit window

class CKCOMMaskCtrl;

class CMaskEdit : public CEdit
{
// Construction
public:
	CMaskEdit(CKCOMMaskCtrl * pCtrl);

// Attributes
public:
	CKCOMMaskCtrl * m_pCtrl;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMaskEdit)
	//}}AFX_VIRTUAL

// Implementation
public:
	void MoveRight();
	void MoveLeft();
	virtual ~CMaskEdit();

	// Generated message map functions
protected:
	//{{AFX_MSG(CMaskEdit)
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg	void OnPaste();
	afx_msg	void OnCut();
	afx_msg void OnClear();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MASKEDIT_H__D1728E37_4985_11D3_A7FE_0080C8763FA4__INCLUDED_)
