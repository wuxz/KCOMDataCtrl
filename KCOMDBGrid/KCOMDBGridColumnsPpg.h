#if !defined(AFX_KCOMDBGRIDCOLUMNSPPG_H__191FE9C2_052F_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_KCOMDBGRIDCOLUMNSPPG_H__191FE9C2_052F_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// KCOMDBGridColumnsPpg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridColumnsPpg : Property page dialog
#include "KCOMDBGridCtl.h"

class CKCOMDBGridColumnsPpg : public COlePropertyPage
{
	DECLARE_DYNCREATE(CKCOMDBGridColumnsPpg)
	DECLARE_OLECREATE_EX(CKCOMDBGridColumnsPpg)

// Constructors
public:
	CKCOMDBGridColumnsPpg();

// Dialog Data
	//{{AFX_DATA(CKCOMDBGridColumnsPpg)
	enum { IDD = IDD_PROPPAGE_KCOMDBGRID_COLUMNS };
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);        // DDX/DDV support
	int m_nColumnBindNum[255], m_nColumnsBind;
	TCHAR m_strColumnHeader[255][40];
	CKCOMDBGridCtrl * m_pCtrl;

	CComboBox * m_pComboColumn, * m_pComboField;
	CEdit * m_pEditHeader;

// Message maps
protected:
	//{{AFX_MSG(CKCOMDBGridColumnsPpg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSelendokComboColumn();
	afx_msg void OnSelendokComboDatafield();
	afx_msg void OnChangeEditHeader();
	afx_msg void OnButtonReset();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMDBGRIDCOLUMNSPPG_H__191FE9C2_052F_11D3_B446_0080C8F18522__INCLUDED_)
