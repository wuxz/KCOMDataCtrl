#if !defined(AFX_COLUMNS_H__61724101_12B3_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_COLUMNS_H__61724101_12B3_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GridColumns.h : header file
//

// implements the IGridColumns interface

/////////////////////////////////////////////////////////////////////////////
// CGridColumns command target
class CKCOMRichGridCtrl;

class CGridColumns : public CCmdTarget
{
	DECLARE_DYNCREATE(CGridColumns)

	CGridColumns();           // protected constructor used by dynamic creation

// Attributes
public:

// Operations
public:
	virtual BOOL GetDispatchIID(IID* pIID);
	ROWCOL GetRowColFromVariant(VARIANT va);
	SetCtrlPtr(CKCOMRichGridCtrl * pCtrl);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGridColumns)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CGridColumns();
	CKCOMRichGridCtrl * m_pCtrl;

	// Generated message map functions
	//{{AFX_MSG(CGridColumns)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_OLETYPELIB(CGridColumns)      // GetTypeInfo

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CGridColumns)
	afx_msg short GetCount();
	afx_msg void Remove(const VARIANT FAR& ColIndex);
	afx_msg void Add(const VARIANT FAR& ColIndex);
	afx_msg void RemoveAll();
	afx_msg LPDISPATCH Item(const VARIANT FAR& ColIndex);
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

#endif // !defined(AFX_COLUMNS_H__61724101_12B3_11D3_A7FE_0080C8763FA4__INCLUDED_)
