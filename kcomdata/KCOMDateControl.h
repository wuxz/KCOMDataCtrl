#if !defined(AFX_KCOMDATECONTROL_H__659BF443_1CC9_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMDATECONTROL_H__659BF443_1CC9_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// KCOMDateControl.h : header file
//

// implements the date editbox style cell

/////////////////////////////////////////////////////////////////////////////
// CKCOMDateControl window

class CKCOMDateControl : public CGXEditControl
{
	DECLARE_CONTROL(CKCOMDateControl)
	DECLARE_DYNAMIC(CKCOMDateControl)

// Construction
public:
	CKCOMDateControl(CGXGridCore* pGrid, UINT nID = 0);

// Attributes
public:
	virtual BOOL OnGridKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	virtual BOOL Store();
	virtual void SetValue(LPCTSTR pszRawValue);
	virtual BOOL GetValue(CString& strResult);
	virtual void Init(ROWCOL nRow, ROWCOL nCol);

protected:
	void CalcData();
	BOOL SetTextIn(CString strNewValue);
	BOOL m_bTextNull;
	void CalcMingGuo();
	BOOL IsValidFormat(CString strNewFormat); 
	// check if the input is finished
	BOOL IsFinished();
	// move the caret right, leave the delimiter character
	void MoveRight(int nSteps = 0);
	// move the caret left, leave the delimiter character
	void MoveLeft(int nSteps = 0);
	// calculate the data owned by myself
	void CalcText();
	void SetDateFormat(LPCTSTR lpszNewValue);
	void SetText(LPCTSTR lpszNewValue);
	void SetDate(short nYear, short nMonth, short nDay);

protected:
	BOOL m_bMingGuo, m_bFormatMingGuo;
	// format used to input data
	CString strFormat;
	// the data used when active
	CString strData;
	// the data used to draw text
	CString strCurrentData;
	// the delimiter characters
	CString strDelimiter[4];

	CString m_strText;

	// the actually data and their length
	int year, month, day, nYearLength, nMonthLength, nDayLength;
	// the start and end position of data
	int nYearPosition[2], nMonthPosition[2], nDayPosition[2];
	// the amount of delimiers
	int nDelimiters;
	// the start and posiiton of delimiters
	int nDelimiterPosition[8];
	// the position of the focus
	int nCurrentPosition;
	// the position of the first unfinished character's position
	int nYearFinish, nMonthFinish, nDayFinish;
	// the previous data
	SYSTEMTIME m_PrevData;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMDateControl)
	//}}AFX_VIRTUAL

// Implementation
public:
	void ButtonDown(USHORT iButton, UINT, CPoint point);
	BOOL UpdateEditStyle();
	void Hide();
	virtual void Draw(CDC* pDC, CRect rect, ROWCOL nRow, ROWCOL nCol, const CGXStyle& style, const CGXStyle* pStandardStyle);
	virtual BOOL OnGridChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	virtual BOOL GetControlText(CString& strDisplay, ROWCOL nRow, ROWCOL nCol, LPCTSTR pszRawValue, const CGXStyle& style);
	void SetWindowText(CString strNewText);
	virtual void SetCurrentText(const CString& str);
	void InitMask(const CGXStyle& style);
	virtual ~CKCOMDateControl();

	// Generated message map functions
protected:
	//{{AFX_MSG(CKCOMDateControl)
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnSysChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnSysKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnSysKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMDATECONTROL_H__659BF443_1CC9_11D3_A7FE_0080C8763FA4__INCLUDED_)
