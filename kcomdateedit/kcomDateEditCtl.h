#if !defined(AFX_EDITBOXCTL_H__F6D0E414_2D34_11D2_9A46_0080C8763FA4__INCLUDED_)
#define AFX_EDITBOXCTL_H__F6D0E414_2D34_11D2_9A46_0080C8763FA4__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

// EditboxCtl.h : Declaration of the CKCOMDateEditCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditCtrl : See EditboxCtl.cpp for implementation.
#include "DlgDatePicker.h" // The dialog to choose a date
enum
{
	Left = 0, 
	Center = 1, 
	Right = 2
};

class CKCOMDateEditCtrl : public COleControl
{
	DECLARE_DYNCREATE(CKCOMDateEditCtrl)

// Constructor
public:
	CKCOMDateEditCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMDateEditCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void OnFontChanged();
	virtual void OnForeColorChanged();
	virtual void OnBackColorChanged();
	virtual void OnKeyDownEvent(USHORT nChar, USHORT nShiftState);
	virtual void OnKeyPressEvent(USHORT nChar);
	virtual BOOL OnGetDisplayString(DISPID dispid, CString& strValue);
	virtual BOOL OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray);
	virtual BOOL OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut);
	//}}AFX_VIRTUAL

// Implementation
protected:
	void ExchangeStockProps(CPropExchange *pPX);
	void CalcData();
	BOOL SetTextIn(CString strNewValue);
	CString m_strText;
	BOOL m_bTextNull;
	void CalcMingGuo();
	void DisplayCaret(); // display the caret
	OLE_COLOR m_ForeColor;
	OLE_COLOR m_BackColor;
	COLORREF m_clrText;
	CFont * m_pFont;
	LOGFONT m_LogFont;
	BOOL m_bMingGuo, m_bDisplayMingGuo, m_bFormatMingGuo, m_bDisplayFormatMingGuo;

	// check the validity of the input format
	BOOL IsValidFormat(CString strNewFormat); 
	// check if the input is finished
	BOOL IsFinished();
	// move the caret right, leave the delimiter character
	void MoveRight(int nSteps = 0);
	// move the caret left, leave the delimiter character
	void MoveLeft(int nSteps = 0);
	// calculate the data owned by myself
	void CalcText();

	CDlgDatePicker dlgDatePicker;
	// format used to input data
	CString strFormat;
	// format used to display when inactive
	CString strDisplayFormat;
	// the data used when active
	CString strData;
	// the data used to draw text
	CString strCurrentData;
	// the data used when inactive
	CString strDisplayData;
	// the delimiter characters
	CString strDelimiter[4];
	// TextAlignment
	short m_ta;

	// the actually data and their length
	int year, month, day, nYearLength, nMonthLength, nDayLength;
	// the start and end position of data
	int nYearPosition[2], nMonthPosition[2], nDayPosition[2];
	// the length of the data
	int nYearDisplayLength, nMonthDisplayLength, nDayDisplayLength;
	// the amount of delimiers
	int nDelimiters;
	// the start and posiiton of delimiters
	int nDelimiterPosition[8];
	// the position to show caret
	CPoint ptCaretPosition;
	// the position of the focus
	int nCurrentPosition;
	// the position of the first unfinished character's position
	int nYearFinish, nMonthFinish, nDayFinish;
	// the previous data
	SYSTEMTIME m_PrevData;

	~CKCOMDateEditCtrl();
	DECLARE_OLECREATE_EX(CKCOMDateEditCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CKCOMDateEditCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CKCOMDateEditCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CKCOMDateEditCtrl)		// Type name and misc status

	// Subclassed control support
	BOOL PreCreateWindow(CREATESTRUCT& cs);

// Message maps
	//{{AFX_MSG(CKCOMDateEditCtrl)
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	afx_msg void AboutBox();

// Dispatch maps
	//{{AFX_DISPATCH(CKCOMDateEditCtrl)
	BOOL m_bReadOnly;
	afx_msg void OnReadOnlyChanged();
	BOOL m_bCanYearNegative;
	afx_msg void OnCanYearNegativeChanged();
	afx_msg BSTR GetDateFormat();
	afx_msg void SetDateFormat(LPCTSTR lpszNewValue);
	afx_msg BSTR GetDisplayDateFormat();
	afx_msg void SetDisplayDateFormat(LPCTSTR lpszNewValue);
	afx_msg BSTR GetText();
	afx_msg void SetText(LPCTSTR lpszNewValue);
	afx_msg short GetTextAlign();
	afx_msg void SetTextAlign(short nNewValue);
	afx_msg void SetDate(short nYear, short nMonth, short nDay);
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()

// Event maps
	//{{AFX_EVENT(CKCOMDateEditCtrl)
	void FireChange()
		{FireEvent(eventidChange,EVENT_PARAM(VTS_NONE));}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	void MoveToNextControl(BOOL bForward = TRUE);
	enum {
	//{{AFX_DISP_ID(CKCOMDateEditCtrl)
	dispidDateFormat = 3L,
	dispidDisplayDateFormat = 4L,
	dispidText = 5L,
	dispidReadOnly = 1L,
	dispidTextAlign = 6L,
	dispidCanYearNegative = 2L,
	dispidSetDate = 7L,
	eventidChange = 1L,
	//}}AFX_DISP_ID
	};

	friend class CDlgDatePicker;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EDITBOXCTL_H__F6D0E414_2D34_11D2_9A46_0080C8763FA4__INCLUDED)
