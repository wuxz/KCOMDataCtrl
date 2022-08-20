// AddColumnDlg.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "AddColumnDlg.h"
#include "KCOMRichGridColumnPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAddColumnDlg dialog


CAddColumnDlg::CAddColumnDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAddColumnDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAddColumnDlg)
	m_strName = _T("");
	//}}AFX_DATA_INIT
}


void CAddColumnDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAddColumnDlg)
	DDX_Text(pDX, IDC_EDIT_COLUMNNAME, m_strName);
	//}}AFX_DATA_MAP
	if (pDX->m_bSaveAndValidate)
	{
		if (m_strName.IsEmpty())
		{
			pDX->Fail();
			return;
		}

		for (int i = 0; i < m_pPage->m_apColumnProp.GetSize(); i ++)
		{
			// name should be unique
			if (!m_strName.CompareNoCase(m_pPage->m_apColumnProp[i]
				->strName))
			{
				AfxMessageBox(IDS_ERROR_NAMENOTUNIQUE);
				pDX->Fail();
				return;
			}
		}
	}
}


BEGIN_MESSAGE_MAP(CAddColumnDlg, CDialog)
	//{{AFX_MSG_MAP(CAddColumnDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddColumnDlg message handlers

void CAddColumnDlg::SetPagePtr(CKCOMRichGridPropColumnPage *pPage)
{
	m_pPage = pPage;
}
