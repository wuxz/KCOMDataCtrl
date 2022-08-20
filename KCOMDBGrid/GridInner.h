#if !defined(AFX_GRIDINNER_H__D46ED320_FCA2_11D2_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_GRIDINNER_H__D46ED320_FCA2_11D2_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GridInner.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGridInner window
#include "gridctrl.h"

class CKCOMDBGridCtrl;

class CGridInner : public CGridCtrl
{
// Construction
public:
	CGridInner(CKCOMDBGridCtrl * pCtrl);

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGridInner)
	//}}AFX_VIRTUAL

// Implementation
public:
	void HideInsertImage();
	void ShowInsertImage();
	virtual void OnFixedRowClick(CCellID& cell);
	virtual void OnFixedColClick(CCellID& cell);
	BOOL m_bInEdit;
	virtual void OnEditCell(int nRow, int nCol, UINT nChar);
	virtual void OnEndEditCell(int nRow, int nCol, CString str);
	virtual CCellID SetFocusCell(CCellID cell);
	virtual CString GetItemText(int nRow, int nCol);
	virtual ~CGridInner();

	// Generated message map functions
protected:
	BOOL m_bColDirty[255];
	BOOL m_bInsertMode;
	virtual void CreateInPlaceEditControl(CRect& rect, DWORD dwStyle, UINT nID, int nRow, int nCol, LPCTSTR szText, int nChar);
	CImageList m_ImageList;
	CKCOMDBGridCtrl* m_pCtrl;
	BOOL m_bDirty;
	int m_nRowDirty;

	//{{AFX_MSG(CGridInner)
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnMButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	friend class CKCOMDBGridCtrl;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GRIDINNER_H__D46ED320_FCA2_11D2_A7FE_0080C8763FA4__INCLUDED_)
