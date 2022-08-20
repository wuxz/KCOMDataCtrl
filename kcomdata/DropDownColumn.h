#if !defined(AFX_DROPDOWNCOLUMN_H__491325C4_3068_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_DROPDOWNCOLUMN_H__491325C4_3068_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DropDownColumn.h : header file
//

// implement the dropdowncolumn interface for dropdown control and combobox control

/////////////////////////////////////////////////////////////////////////////
// CDropDownColumn command target
class CKCOMRichDropDownCtrl;
class CKCOMRichComboCtrl;

class CDropDownColumn : public CCmdTarget
{
	DECLARE_DYNCREATE(CDropDownColumn)

	CDropDownColumn();           // protected constructor used by dynamic creation

// Attributes
public:

// Operations
public:
	virtual BOOL GetDispatchIID(IID* pIID);
	SetCtrlPtr(CKCOMRichDropDownCtrl * pCtrl);
	SetCtrlPtr(CKCOMRichComboCtrl * pCtrl);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDropDownColumn)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CDropDownColumn();
	BSTR GetCellText(long RowIndex);

protected:
	CKCOMRichDropDownCtrl * m_pDropDownCtrl;
	CKCOMRichComboCtrl * m_pComboCtrl;
	ColumnProp prop;

	// Generated message map functions
	//{{AFX_MSG(CDropDownColumn)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_OLETYPELIB(CDropDownColumns)      // GetTypeInfo

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CDropDownColumn)
	afx_msg long GetAlignment();
	afx_msg void SetAlignment(long nNewValue);
	afx_msg OLE_COLOR GetBackColor();
	afx_msg void SetBackColor(OLE_COLOR nNewValue);
	afx_msg BSTR GetCaption();
	afx_msg void SetCaption(LPCTSTR lpszNewValue);
	afx_msg long GetCaptionAlignment();
	afx_msg void SetCaptionAlignment(long nNewValue);
	afx_msg long GetCase();
	afx_msg void SetCase(long nNewValue);
	afx_msg BSTR GetDataField();
	afx_msg void SetDataField(LPCTSTR lpszNewValue);
	afx_msg OLE_COLOR GetForeColor();
	afx_msg void SetForeColor(OLE_COLOR nNewValue);
	afx_msg OLE_COLOR GetHeadBackColor();
	afx_msg void SetHeadBackColor(OLE_COLOR nNewValue);
	afx_msg OLE_COLOR GetHeadForeColor();
	afx_msg void SetHeadForeColor(OLE_COLOR nNewValue);
	afx_msg BSTR GetName();
	afx_msg void SetName(LPCTSTR lpszNewValue);
	afx_msg BOOL GetVisible();
	afx_msg void SetVisible(BOOL bNewValue);
	afx_msg long GetWidth();
	afx_msg void SetWidth(long nNewValue);
	afx_msg BSTR GetText();
	afx_msg BSTR CellText(const VARIANT FAR& Bookmark);
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

	friend class CKCOMRichDropDownCtrl;
	friend class CKCOMRichComboCtrl;
	friend class CDropDownColumns;
	friend class CKCOMRichGridPropColumnPage;
	friend class CDropDownRealWnd;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DROPDOWNCOLUMN_H__491325C4_3068_11D3_A7FE_0080C8763FA4__INCLUDED_)
