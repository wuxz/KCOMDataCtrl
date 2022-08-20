#if !defined(AFX_SelBookmarks_H__8119E221_3B61_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_SelBookmarks_H__8119E221_3B61_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// SelBookmarks.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CSelBookmarks command target
class CKCOMRichGridCtrl;

class CSelBookmarks : public CCmdTarget
{
	DECLARE_DYNCREATE(CSelBookmarks)

	CSelBookmarks();           // protected constructor used by dynamic creation

// Attributes
public:

protected:
	virtual BOOL GetDispatchIID(IID* pIID);
	SetCtrlPtr(CKCOMRichGridCtrl * pCtrl);

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSelBookmarks)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CSelBookmarks();
	CKCOMRichGridCtrl * m_pCtrl;

	// Generated message map functions
	//{{AFX_MSG(CSelBookmarks)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_OLETYPELIB(CSelBookmarks)      // GetTypeInfo

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CSelBookmarks)
	afx_msg long GetCount();
	afx_msg void Remove(long Index);
	afx_msg void Add(const VARIANT FAR& Bookmark);
	afx_msg void RemoveAll();
	afx_msg VARIANT Item(long Index);
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

	friend class CKCOMRichGridCtrl;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SelBookmarks_H__8119E221_3B61_11D3_A7FE_0080C8763FA4__INCLUDED_)
