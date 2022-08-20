#if !defined(AFX_DLGDATEPICKER_H__B1FAA422_3288_11D2_BA76_0080C8763FA4__INCLUDED_)
#define AFX_DLGDATEPICKER_H__B1FAA422_3288_11D2_BA76_0080C8763FA4__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000
// DlgDatePicker.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDlgDatePicker dialog
class CKCOMDateEditCtrl;

class CDlgDatePicker : public CDialog
{
// Construction
public:
	// we use m_Date to save the user's choice
	COleDateTime m_Date;
	CDlgDatePicker(CWnd* pParent = NULL);   // standard constructor
	CKCOMDateEditCtrl * m_pCtrl;

// Dialog Data
	//{{AFX_DATA(CDlgDatePicker)
	enum { IDD = IDD_DATEPICKER };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgDatePicker)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
protected:
	void OnRightMonth(); // move to the next month
	void OnRightYear(); // move to the next year
	void OnLeftMonth(); // move to the previous month
	void OnLeftYear(); // move to the previous year
	void ShiftYear(int nStep); // move the year
	void ShiftMonth(int nStep); // move the month

	int Cells[6][7]; // the array saves the day number of each cell
	void CalcCells(); // calculate the content of cells array
	// get how many days in the given month
	int GetDaysInMonth(int year, int month); 
	// draw the grid and cells
	void DrawGrid(CPaintDC& dc);
	// display the buttons
	void DrawButton();
	// display the current chosen date at the bottom of the dialog box
	void DrawTitle(CPaintDC& dc);

	// the buttons to flip the calendar
	CButton * m_pButtonLeftYear;
	CButton * m_pButtonLeftMonth;
	CButton * m_pButtonRightYear;
	CButton * m_pButtonRightMonth;
	CComboBox m_ComboMonth;
	CEdit m_EditYear;
	CSpinButtonCtrl m_SpinYear;

	// the rect used to position the buttons
	CRect rectLeftYear, rectLeftMonth, rectRightYear, rectRightMonth, 
		rectComboMonth, rectEditYear, rectSpinYear;

	// the temporary rect variable
	CRect rect;

	// the name of the 7 days in one week
	CString m_strDayName[7];
	// the current chosen date
	COleDateTime m_CurrentDate;
	// the current chosen date's data and today
	int year, month, day, today;

	CFont m_Font;
	CFont * m_pFontPrev;
	CBrush brhGray, brhWhite, brhBlue;
	// Generated message map functions
	//{{AFX_MSG(CDlgDatePicker)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnSelchangeComboMonth();
	afx_msg void OnDeltaposSpinYear(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnChangeEditYear();
	afx_msg void OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGDATEPICKER_H__B1FAA422_3288_11D2_BA76_0080C8763FA4__INCLUDED_)
