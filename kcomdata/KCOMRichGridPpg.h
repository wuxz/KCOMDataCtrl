#if !defined(AFX_KCOMRICHGRIDPPG_H__029288DC_2FFF_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_KCOMRICHGRIDPPG_H__029288DC_2FFF_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMRichGridPpg.h : Declaration of the CKCOMRichGridPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropPage : See KCOMRichGridPpg.cpp.cpp for implementation.

class CKCOMRichGridPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CKCOMRichGridPropPage)
	DECLARE_OLECREATE_EX(CKCOMRichGridPropPage)

// Constructors
public:
	void UpdateControls();
	CKCOMRichGridPropPage();

// Dialog Data
	//{{AFX_DATA(CKCOMRichGridPropPage)
	enum { IDD = IDD_PROPPAGE_KCOMRICHGRID };
	int		m_nDataMode;
	BOOL	m_bAllowAddNew;
	BOOL	m_bAllowDelete;
	BOOL	m_bAllowUpdate;
	BOOL	m_bAllowRowSizing;
	BOOL	m_bAllowColMoving;
	BOOL	m_bAllowColSizing;
	BOOL	m_bColumnHeaders;
	BOOL	m_bRecordSelectors;
	BOOL	m_bSelectByCell;
	CString	m_strCaption;
	CString	m_strFieldSeparator;
	int		m_nBorderStyle;
	int		m_nCaptionAlignment;
	int		m_nDividerStyle;
	int		m_nDividerType;
	long	m_nCols;
	long	m_nDefColWidth;
	long	m_nHeaderHeight;
	long	m_nRowHeight;
	long	m_nRows;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);        // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CKCOMRichGridPropPage)
	afx_msg void OnSelendokComboDatamode();
	afx_msg void OnCheckAllowupdate();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMRICHGRIDPPG_H__029288DC_2FFF_11D3_B446_0080C8F18522__INCLUDED)
