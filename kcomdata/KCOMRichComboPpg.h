#if !defined(AFX_KCOMRICHCOMBOPPG_H__029288E1_2FFF_11D3_B446_0080C8F18522__INCLUDED_)
#define AFX_KCOMRICHCOMBOPPG_H__029288E1_2FFF_11D3_B446_0080C8F18522__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMRichComboPpg.h : Declaration of the CKCOMRichComboPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboPropPage : See KCOMRichComboPpg.cpp.cpp for implementation.

class CKCOMRichComboPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CKCOMRichComboPropPage)
	DECLARE_OLECREATE_EX(CKCOMRichComboPropPage)

// Constructor
public:
	CKCOMRichComboPropPage();

// Dialog Data
	//{{AFX_DATA(CKCOMRichComboPropPage)
	enum { IDD = IDD_PROPPAGE_KCOMRICHCOMBO };
	BOOL	m_bReadOnly;
	BOOL	m_bColumnHeaders;
	BOOL	m_bListAutoPosition;
	BOOL	m_bListWidthAutoSize;
	int		m_nDividerStyle;
	int		m_nDividerType;
	int		m_nBorderStyle;
	int		m_nDataMode;
	CString	m_strFieldSeparator;
	long	m_nRows;
	long	m_nCols;
	long	m_n;
	long	m_nMinDropDownItems;
	long	m_nDefColWidth;
	long	m_nHeaderHeight;
	long	m_nRowHeight;
	long	m_nListWidth;
	CString	m_strFormat;
	long	m_nMaxLength;
	int		m_nStyle;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CKCOMRichComboPropPage)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMRICHCOMBOPPG_H__029288E1_2FFF_11D3_B446_0080C8F18522__INCLUDED)
