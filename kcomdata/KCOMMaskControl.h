#if !defined(AFX_KCOMMASKCONTROL_H__4DDD2120_1DB5_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMMASKCONTROL_H__4DDD2120_1DB5_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// KCOMMaskControl.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskControl window

class CKCOMMaskControl : public CGXMaskControl
{
// Construction
public:
	CKCOMMaskControl(CGXGridCore* pGrid, UINT nID, CRuntimeClass* pDataClass = RUNTIME_CLASS(CGXMaskData));

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMMaskControl)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual BOOL ConvertControlTextToValue(CString& str, ROWCOL nRow, ROWCOL nCol, const CGXStyle* pOldStyle);
	BOOL IsPromptInclude();
	virtual BOOL GetControlText(CString& strDisplay, ROWCOL nRow, ROWCOL nCol, LPCTSTR pszRawValue, const CGXStyle& style);
	virtual void SetValue(LPCTSTR pszRawValue);
	virtual BOOL GetValue(CString& strResult);
	virtual void GetCurrentText(CString& sMaskedResult);
	virtual ~CKCOMMaskControl();

	// Generated message map functions
protected:
	//{{AFX_MSG(CKCOMMaskControl)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMMASKCONTROL_H__4DDD2120_1DB5_11D3_A7FE_0080C8763FA4__INCLUDED_)
