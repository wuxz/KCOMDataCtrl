#if !defined(AFX_GRIDINNER_H__0CFC2336_12A9_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_GRIDINNER_H__0CFC2336_12A9_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GridInner.h : header file
//

// the grid window shown in grid control as a child window

/////////////////////////////////////////////////////////////////////////////
// CGridInner window
class CKCOMRichGridCtrl;

class CGridInner : public CGXBrowserWnd
{
// Construction
public:
	CGridInner(CKCOMRichGridCtrl * pCtrl);

// Attributes
public:
	BOOL m_bColDirty[255];

protected:
	CKCOMRichGridCtrl * m_pCtrl;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGridInner)
	//}}AFX_VIRTUAL

// Implementation
public:
	BOOL Moveto(ROWCOL nRow);
	// the content of the modified fields of current row
	CStringArray m_arContent;

	ROWCOL GetRecordCount();
	BOOL IsAddNewMode();
	
	// cancel the modification of current row, so that the grid can be
	// updated according to the change of the outer datasource
	virtual void CancelRecord();

	virtual BOOL SetCurrentCell(ROWCOL nRow, ROWCOL nCol, UINT flags = GX_SCROLLINVIEW | GX_UPDATENOW);
	BOOL IsColDirty(ROWCOL nCol);
	ROWCOL GetColFromIndex(ROWCOL nColIndex);
	virtual ~CGridInner();

	// Generated message map functions
protected:
	virtual BOOL OnGridSysKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	virtual BOOL OnGridKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	virtual BOOL OnLoadCellStyle(ROWCOL nRow, ROWCOL nCol, CGXStyle &style, LPCTSTR pszExistingValue);
	virtual BOOL StoreMoveCols(ROWCOL nFromCol, ROWCOL nToCol, ROWCOL nDestCol, BOOL bProcessed = FALSE);
	virtual void OnModifyCell (ROWCOL nRow, ROWCOL nCol);
	virtual BOOL DoScroll(int direction, ROWCOL nCell);
	virtual BOOL OnLButtonHitRowCol(ROWCOL nHitRow, ROWCOL nHitCol, ROWCOL nDragRow, ROWCOL nDragCol, CPoint point, UINT flags, WORD nHitState);
	virtual BOOL MoveCurrentCell(int direction, UINT nCell = 1 , BOOL bCanSelectRange = TRUE);
	virtual BOOL OnStartTracking(ROWCOL nRow, ROWCOL nCol, int nTrackingMode);
	virtual ROWCOL AddNew();
	virtual BOOL StoreStyleRowCol(ROWCOL nRow, ROWCOL nCol, const CGXStyle* pStyle, GXModifyType mt = gxOverride, int nType = 0);
	virtual void OnClickedButtonRowCol(ROWCOL nRow, ROWCOL nCol);
	virtual int HitTest(CPoint &pt, ROWCOL* pnRow, ROWCOL* pnCol, CRect* pRectHit);
	virtual void DeleteRecords();
	virtual BOOL OnFlushRecord(ROWCOL nRow, CString* ps = NULL);
	virtual void OnGridInitialUpdate();
	//{{AFX_MSG(CGridInner)
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnMButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnMButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	friend class CGridColumn;
	friend class CGridColumns;
	friend class CKCOMRichGridCtrl;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GRIDINNER_H__0CFC2336_12A9_11D3_A7FE_0080C8763FA4__INCLUDED_)
