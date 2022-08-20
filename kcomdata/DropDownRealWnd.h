#if !defined(AFX_DROPDOWNREALWND_H__50C62AA2_32DA_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_DROPDOWNREALWND_H__50C62AA2_32DA_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DropDownRealWnd.h : header file
//

// the window shown as the dropdown, used in dropdown and combobox control

/////////////////////////////////////////////////////////////////////////////
// CDropDownRealWnd window
class CKCOMRichDropDownCtrl;
class CKCOMRichComboCtrl;

class CDropDownRealWnd : public CGXBrowserWnd
{
// Construction
public:
	CDropDownRealWnd(CKCOMRichDropDownCtrl * pCtrl);
	CDropDownRealWnd(CKCOMRichComboCtrl * pCtrl);

// Attributes
public:
	CKCOMRichDropDownCtrl * m_pDropDownCtrl;
	CKCOMRichComboCtrl * m_pComboCtrl;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDropDownRealWnd)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual BOOL SetCurrentCell(ROWCOL nRow, ROWCOL nCol, UINT flags = GX_SCROLLINVIEW | GX_UPDATENOW);
	
	// move the current row
	virtual BOOL MoveTo(ROWCOL nRow);

	// the helper function to show/hide the window and modify the status variable
	void Showwindow(BOOL bShow);

	ROWCOL GetColFromIndex(ROWCOL nColIndex);
	virtual ~CDropDownRealWnd();

	// Generated message map functions
protected:
	virtual BOOL DoScroll(int direction, ROWCOL nCell);
	void ReturnText(ROWCOL nRow);
	virtual BOOL ProcessKeys(CWnd *pSender, UINT nMessage, UINT nChar, UINT nRepCnt, UINT flags);
	virtual BOOL OnLButtonClickedRowCol(ROWCOL nRow, ROWCOL nCol, UINT nFlags, CPoint pt);
	virtual BOOL OnLoadCellStyle(ROWCOL nRow, ROWCOL nCol, CGXStyle &style, LPCTSTR pszExistingValue);
	//{{AFX_MSG(CDropDownRealWnd)
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	friend class CKCOMRichDropDownCtrl;
	friend class CKCOMRichComboCtrl;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DROPDOWNREALWND_H__50C62AA2_32DA_11D3_A7FE_0080C8763FA4__INCLUDED_)
