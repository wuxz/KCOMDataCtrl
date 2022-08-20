#if !defined(AFX_KCOMDBGRIDPPG_H__AC212656_FBE2_11D2_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMDBGRIDPPG_H__AC212656_FBE2_11D2_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMDBGridPpg.h : Declaration of the CKCOMDBGridPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridPropPage : See KCOMDBGridPpg.cpp.cpp for implementation.

class CKCOMDBGridPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CKCOMDBGridPropPage)
	DECLARE_OLECREATE_EX(CKCOMDBGridPropPage)

// Constructor
public:
	CKCOMDBGridPropPage();

// Dialog Data
	//{{AFX_DATA(CKCOMDBGridPropPage)
	enum { IDD = IDD_PROPPAGE_KCOMDBGRID };
	int		m_nDataMode;
	BOOL	m_bShowRecordSelector;
	BOOL	m_bAllowAddNew;
	BOOL	m_bAllowDelete;
	BOOL	m_bReadOnly;
	BOOL	m_bAllowArrowKeys;
	BOOL	m_bAllowRowSizing;
	long	m_nDefColWidth;
	long	m_nHeaderHeight;
	long	m_nRowHeight;
	int		m_nGridLineStyle;
	long	m_nRows;
	long	m_nCols;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	void UpdateControls();
	//{{AFX_MSG(CKCOMDBGridPropPage)
	afx_msg void OnSelendokListDatamode();
	afx_msg void OnCheckReadonly();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMDBGRIDPPG_H__AC212656_FBE2_11D2_A7FE_0080C8763FA4__INCLUDED)
