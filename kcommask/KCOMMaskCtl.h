#if !defined(AFX_KCOMMASKCTL_H__D1728E33_4985_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_KCOMMASKCTL_H__D1728E33_4985_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// KCOMMaskCtl.h : Declaration of the CKCOMMaskCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskCtrl : See KCOMMaskCtl.cpp for implementation.
class CMaskEdit;

class CKCOMMaskCtrl : public COleControl
{
	DECLARE_DYNCREATE(CKCOMMaskCtrl)

// Constructor
public:
	CKCOMMaskCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKCOMMaskCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual void OnFontChanged();
	virtual BOOL OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray);
	virtual BOOL OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut);
	virtual BOOL OnGetDisplayString(DISPID dispid, CString& strValue);
	//}}AFX_VIRTUAL

// Implementation
protected:
	BOOL m_bAutoTab;
	CString m_strMask;
	short m_nMaxLength;
	CString m_strPromptChar;
	BOOL m_bPromptInclude;
	
	// the brush to draw the background
	CBrush * m_pBrhBack;

	// the handle of the using font for editbox
	HFONT m_hFont;

	// the pointer of the editbox window
	CMaskEdit * m_pEdit;

	// the signal to mark whether current char is delimiter or data
	BOOL m_bIsDelimiter[255];

	CString m_strInputMask;

	BOOL m_bIsInitializing;
protected:
	void ExchangeStockProps(CPropExchange *pPX);
	void CalcMask();
	BOOL IsValidChar(TCHAR nChar, int nPosition);
	BOOL IsValidMask(CString strMask);
	void MoveRight();
	void CalcText();
	void SetDisplayText(CString strText);
	~CKCOMMaskCtrl();

	DECLARE_OLECREATE_EX(CKCOMMaskCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CKCOMMaskCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CKCOMMaskCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CKCOMMaskCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CKCOMMaskCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CKCOMMaskCtrl)
	afx_msg BOOL GetPromptInclude();
	afx_msg void SetPromptInclude(BOOL bNewValue);
	afx_msg BOOL GetAutoTab();
	afx_msg void SetAutoTab(BOOL bNewValue);
	afx_msg short GetMaxLength();
	afx_msg void SetMaxLength(short nNewValue);
	afx_msg BSTR GetMask();
	afx_msg void SetMask(LPCTSTR lpszNewValue);
	afx_msg long GetSelLength();
	afx_msg void SetSelLength(long nNewValue);
	afx_msg long GetSelStart();
	afx_msg void SetSelStart(long nNewValue);
	afx_msg BSTR GetSelText();
	afx_msg void SetSelText(LPCTSTR lpszNewValue);
	afx_msg BSTR GetClipText();
	afx_msg void SetClipText(LPCTSTR lpszNewValue);
	afx_msg BSTR GetPromptChar();
	afx_msg void SetPromptChar(LPCTSTR lpszNewValue);
	afx_msg BSTR GetText();
	afx_msg void SetText(LPCTSTR lpszNewValue);
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// Event maps
	//{{AFX_EVENT(CKCOMMaskCtrl)
	void FireChange()
		{FireEvent(eventidChange,EVENT_PARAM(VTS_NONE));}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	enum {
	//{{AFX_DISP_ID(CKCOMMaskCtrl)
	dispidPromptInclude = 1L,
	dispidAutoTab = 2L,
	dispidMaxLength = 3L,
	dispidMask = 4L,
	dispidSelLength = 5L,
	dispidSelStart = 6L,
	dispidSelText = 7L,
	dispidClipText = 8L,
	dispidPromptChar = 9L,
	dispidText = 10L,
	eventidChange = 1L,
	//}}AFX_DISP_ID
	};

	friend class CMaskEdit;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KCOMMASKCTL_H__D1728E33_4985_11D3_A7FE_0080C8763FA4__INCLUDED)
