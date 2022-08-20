#if !defined(AFX_EDITBOXPPG_H__F6D0E416_2D34_11D2_9A46_0080C8763FA4__INCLUDED_)
#define AFX_EDITBOXPPG_H__F6D0E416_2D34_11D2_9A46_0080C8763FA4__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

// EditboxPpg.h : Declaration of the CKCOMDateEditPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CKCOMDateEditPropPage : See EditboxPpg.cpp.cpp for implementation.

class CKCOMDateEditPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CKCOMDateEditPropPage)
	DECLARE_OLECREATE_EX(CKCOMDateEditPropPage)

// Constructor
public:
	CString strFormat[11];
	CKCOMDateEditPropPage();

// Dialog Data
	//{{AFX_DATA(CKCOMDateEditPropPage)
	enum { IDD = IDD_PROPPAGE_KCOMDateEdit };
	CString	m_DateFormat;
	int		m_Style;
	CString	m_DisplayDateFormat;
	BOOL	m_bReadOnly;
	int		m_TextAlign;
	int		m_sAppearance;
	int		m_sBorderStyle;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CKCOMDateEditPropPage)
	afx_msg void OnRadio1();
	afx_msg void OnRadio2();
	afx_msg void OnRadio3();
	afx_msg void OnRadio4();
	afx_msg void OnRadio5();
	afx_msg void OnRadio6();
	afx_msg void OnRadio7();
	afx_msg void OnRadio8();
	afx_msg void OnRadio9();
	afx_msg void OnRadio10();
	afx_msg void OnRadio11();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EDITBOXPPG_H__F6D0E416_2D34_11D2_9A46_0080C8763FA4__INCLUDED)
