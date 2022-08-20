#if !defined(AFX_GRIDINPPG_H__CBD24582_1764_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_GRIDINPPG_H__CBD24582_1764_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GridInPpg.h : header file
//

// the grid in the column property page to provide WYSWYG effect

/////////////////////////////////////////////////////////////////////////////
// CGridInPpg window

class CKCOMRichGridPropColumnPage;

class CGridInPpg : public CGXBrowserWnd
{
// Construction
public:
	CGridInPpg();

	SetPagePtr(CKCOMRichGridPropColumnPage * pPage);

// Attributes
public:
	BOOL m_bCanDirty;

protected:
	CKCOMRichGridPropColumnPage * m_pPage;
// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGridInPpg)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual BOOL StoreRemoveRows(ROWCOL nFromRow, ROWCOL nToRow, BOOL bProcessed = FALSE);
	virtual BOOL StoreColWidth(ROWCOL nCol, int nWidth = 0);
	virtual void OnModifyCell (ROWCOL nRow, ROWCOL nCol);
	ROWCOL GetFieldFromIndex(ROWCOL nColIndex);
	virtual BOOL OnLButtonClickedRowCol(ROWCOL nRow, ROWCOL nCol, UINT nFlags, CPoint pt);
	ROWCOL GetColFromIndex(ROWCOL nColIndex);
	void SetCaption(ROWCOL nCol, LPCTSTR lpszNewValue);
	void SetCaptionAlignment(ROWCOL nCol, long nNewValue);
	void SetStyle(ROWCOL nCol, long nNewValue);
	void SetWidth(ROWCOL nCol, long nNewValue);
	void SetVisible(ROWCOL nCol, BOOL bNewValue);
	void SetLocked(ROWCOL nCol, BOOL bNewValue);
	void SetHeadBackColor(ROWCOL nCol, OLE_COLOR nNewValue);
	void SetHeadForeColor(ROWCOL nCol, OLE_COLOR nNewValue);
	void SetBackColor(ROWCOL nCol, OLE_COLOR nNewValue);
	void SetForeColor(ROWCOL nCol, OLE_COLOR nNewValue);
	void SetAlignment(ROWCOL nCol, long nNewValue);
	virtual ~CGridInPpg();

	// Generated message map functions
protected:
	virtual void OnGridInitialUpdate();
	BOOL m_bLButtonDown;
	//{{AFX_MSG(CGridInPpg)
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GRIDINPPG_H__CBD24582_1764_11D3_A7FE_0080C8763FA4__INCLUDED_)
