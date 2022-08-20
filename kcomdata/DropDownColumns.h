#if !defined(AFX_DROPDOWNCOLUMNS_H__491325C1_3068_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_DROPDOWNCOLUMNS_H__491325C1_3068_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DropDownColumns.h : header file
//

// implements the IDropdownColumns interface

/////////////////////////////////////////////////////////////////////////////
// CDropDownColumns command target
class CKCOMRichDropDownCtrl;
class CKCOMRichComboCtrl;

class CDropDownColumns : public CCmdTarget
{
	DECLARE_DYNCREATE(CDropDownColumns)

	CDropDownColumns();           // protected constructor used by dynamic creation

// Attributes
public:

// Operations
public:
	virtual BOOL GetDispatchIID(IID* pIID);
	ROWCOL GetRowColFromVariant(VARIANT va);
	SetCtrlPtr(CKCOMRichDropDownCtrl * pCtrl);
	SetCtrlPtr(CKCOMRichComboCtrl * pCtrl);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDropDownColumns)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CDropDownColumns();
	CKCOMRichDropDownCtrl * m_pDropDownCtrl;
	CKCOMRichComboCtrl * m_pComboCtrl;

	// Generated message map functions
	//{{AFX_MSG(CDropDownColumns)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_OLETYPELIB(CDropDownColumns)      // GetTypeInfo

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CDropDownColumns)
	afx_msg short GetCount();
	afx_msg void Remove(const VARIANT FAR& ColIndex);
	afx_msg void Add(const VARIANT FAR& ColIndex);
	afx_msg void RemoveAll();
	afx_msg LPDISPATCH GetItem(const VARIANT FAR& ColIndex);
	//}}AFX_DISPATCH
	afx_msg LPUNKNOWN _NewEnum();
	DECLARE_DISPATCH_MAP()

	BEGIN_INTERFACE_PART(EnumVARIANT, IEnumVARIANT)
		STDMETHOD(Next)(THIS_ unsigned long celt, VARIANT FAR* rgvar, 
			unsigned long FAR* pceltFetched);
	    STDMETHOD(Skip)(THIS_ unsigned long celt) ;
		STDMETHOD(Reset)(THIS) ;
	    STDMETHOD(Clone)(THIS_ IEnumVARIANT FAR* FAR* ppenum) ;
		XEnumVARIANT() ;        // constructor to set m_posCurrent
	    int m_nCurrent ; // Next() requires we keep track of our current item
	END_INTERFACE_PART(EnumVARIANT)    

	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DROPDOWNCOLUMNS_H__491325C1_3068_11D3_A7FE_0080C8763FA4__INCLUDED_)
