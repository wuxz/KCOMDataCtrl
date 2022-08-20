#if !defined(AFX_KCOMCOMBOBOX_H__659BF441_1CC9_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMCOMBOBOX_H__659BF441_1CC9_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// KCOMComboBox.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CKCOMComboBox window

// override to implement the dropdown list style

class CKCOMComboBox : public CGXComboBox
{
// Construction
public:
	CKCOMComboBox(CGXGridCore* pGrid, UINT nEditID, UINT nListBoxID, UINT nFlags, BOOL bDropDownList = FALSE);

// Attributes
public:
	BOOL m_bDropDownList;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMComboBox)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual BOOL GetValue(CString &strResult);
	virtual ~CKCOMComboBox();

	// Generated message map functions
protected:
	virtual BOOL OnGridKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	virtual BOOL OnGridChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	//{{AFX_MSG(CKCOMComboBox)
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMCOMBOBOX_H__659BF441_1CC9_11D3_A7FE_0080C8763FA4__INCLUDED_)
