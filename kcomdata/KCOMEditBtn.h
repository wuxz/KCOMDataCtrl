#if !defined(AFX_KCOMEDITBTN_H__676AA100_1E4E_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMEDITBTN_H__676AA100_1E4E_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// KCOMEditBtn.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CKCOMEditBtn window

class CKCOMEditBtn : public CGXEditControl
{
	DECLARE_CONTROL(CKCOMEditBtn)
	DECLARE_DYNAMIC(CKCOMEditBtn)

// Construction
public:
	CKCOMEditBtn(CGXGridCore* pGrid, UINT nEditID);

// Attributes
public:

protected:
	CGXChild * m_pButton;
	int m_nDefaultHeight;
	int m_nExtraFrame;
	BOOL m_bSizeToContent;       // set FALSE if button shall have the same height as cell
	BOOL m_bInactiveDrawAllCell; // draw text over pushbutton when inactive
	BOOL m_bInactiveDrawButton;  // draw pushbutton when inactive

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMEditBtn)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual BOOL GetValue(CString &strResult);
	virtual void OnInitChildren(ROWCOL nRow, ROWCOL nCol, const CRect& rect);
	virtual CSize CalcSize(CDC* pDC, ROWCOL nRow, ROWCOL nCol, const CGXStyle& style, const CGXStyle* pStandardStyle, BOOL bVert);
	virtual void Init(ROWCOL nRow, ROWCOL nCol);
	virtual CSize AddBorders(CSize size, const CGXStyle& style);
	virtual CRect GetCellRect(ROWCOL nRow, ROWCOL nCol, LPRECT rectItem = NULL, const CGXStyle* pStyle = NULL);
	virtual ~CKCOMEditBtn();

	// Generated message map functions
protected:
	//{{AFX_MSG(CKCOMEditBtn)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMEDITBTN_H__676AA100_1E4E_11D3_A7FE_0080C8763FA4__INCLUDED_)
