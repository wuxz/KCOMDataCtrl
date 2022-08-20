#if !defined(AFX_KCOMRICHDROPDOWNPPG_H__029288E6_2FFF_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_KCOMRICHDROPDOWNPPG_H__029288E6_2FFF_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMRichDropDownPpg.h : Declaration of the CKCOMRichDropDownPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CKCOMRichDropDownPropPage : See KCOMRichDropDownPpg.cpp.cpp for implementation.

class CKCOMRichDropDownPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CKCOMRichDropDownPropPage)
	DECLARE_OLECREATE_EX(CKCOMRichDropDownPropPage)

// Constructor
public:
	CKCOMRichDropDownPropPage();

// Dialog Data
	//{{AFX_DATA(CKCOMRichDropDownPropPage)
	enum { IDD = IDD_PROPPAGE_KCOMRICHDROPDOWN };
	int		m_nDataMode;
	BOOL	m_bColumnHeaders;
	BOOL	m_bListAutoPosition;
	BOOL	m_bListWidthAutoSize;
	int		m_nDividerStyle;
	int		m_nDividerType;
	int		m_nBorderStyle;
	CString	m_strFieldSeparator;
	long	m_nRows;
	long	m_nCols;
	long	m_nDefColWidth;
	long	m_nHeaderHeight;
	long	m_nRowHeight;
	long	m_nMaxDropDownItems;
	long	m_nMinDropDownItems;
	long	m_nListWidth;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CKCOMRichDropDownPropPage)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMRICHDROPDOWNPPG_H__029288E6_2FFF_11D3_B446_0080C8F18522__INCLUDED)
