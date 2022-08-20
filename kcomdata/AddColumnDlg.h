#if !defined(AFX_ADDCOLUMNDLG_H__FE9646A0_18CD_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_ADDCOLUMNDLG_H__FE9646A0_18CD_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AddColumnDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAddColumnDlg dialog
class CKCOMRichGridPropColumnPage;

class CAddColumnDlg : public CDialog
{
// Construction
public:
	void SetPagePtr(CKCOMRichGridPropColumnPage * pPage);
	CAddColumnDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAddColumnDlg)
	enum { IDD = IDD_DIALOG_ADDCOLUMN };
	CString	m_strName;
	//}}AFX_DATA

protected:
	CKCOMRichGridPropColumnPage * m_pPage;
// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddColumnDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAddColumnDlg)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDCOLUMNDLG_H__FE9646A0_18CD_11D3_A7FE_0080C8763FA4__INCLUDED_)
