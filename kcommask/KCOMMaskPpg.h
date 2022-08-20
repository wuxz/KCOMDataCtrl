#if !defined(AFX_KCOMMASKPPG_H__D1728E35_4985_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMMASKPPG_H__D1728E35_4985_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMMaskPpg.h : Declaration of the CKCOMMaskPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CKCOMMaskPropPage : See KCOMMaskPpg.cpp.cpp for implementation.

class CKCOMMaskPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CKCOMMaskPropPage)
	DECLARE_OLECREATE_EX(CKCOMMaskPropPage)

// Constructor
public:
	CKCOMMaskPropPage();

// Dialog Data
	//{{AFX_DATA(CKCOMMaskPropPage)
	enum { IDD = IDD_PROPPAGE_KCOMMASK };
	BOOL	m_bAutoTab;
	BOOL	m_bPromptInclude;
	BOOL	m_bHideSelection;
	CString	m_strMask;
	int		m_sAppearance;
	int		m_sBorderStyle;
	short	m_nMaxLength;
	CString	m_strPromptChar;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CKCOMMaskPropPage)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMMASKPPG_H__D1728E35_4985_11D3_A7FE_0080C8763FA4__INCLUDED)
