// KCOMDBGridColumnsPpg.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdbgrid.h"
#include "KCOMDBGridColumnsPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridColumnsPpg dialog

IMPLEMENT_DYNCREATE(CKCOMDBGridColumnsPpg, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CKCOMDBGridColumnsPpg, COlePropertyPage)
	//{{AFX_MSG_MAP(CKCOMDBGridColumnsPpg)
	ON_CBN_SELENDOK(IDC_COMBO_COLUMN, OnSelendokComboColumn)
	ON_CBN_SELENDOK(IDC_COMBO_DATAFIELD, OnSelendokComboDatafield)
	ON_EN_CHANGE(IDC_EDIT_HEADER, OnChangeEditHeader)
	ON_BN_CLICKED(IDC_BUTTON_RESET, OnButtonReset)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

// {191FE9C1-052F-11D3-B446-0080C8F18522}
IMPLEMENT_OLECREATE_EX(CKCOMDBGridColumnsPpg, "KCOMDBGrid.CKCOMDBGridColumnsPpg",
	0x191fe9c1, 0x52f, 0x11d3, 0xb4, 0x46, 0x0, 0x80, 0xc8, 0xf1, 0x85, 0x22)


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridColumnsPpg::CKCOMDBGridColumnsPpgFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMDBGridColumnsPpg

BOOL CKCOMDBGridColumnsPpg::CKCOMDBGridColumnsPpgFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Define string resource for page type; replace '0' below with ID.

	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_KCOMDBGRID_PPG_COLUMNS);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridColumnsPpg::CKCOMDBGridColumnsPpg - Constructor

// TODO: Define string resource for page caption; replace '0' below with ID.

CKCOMDBGridColumnsPpg::CKCOMDBGridColumnsPpg() :
	COlePropertyPage(IDD, IDS_KCOMDBGRID_PPG_COLUMNS_CAPTION)
{
	//{{AFX_DATA_INIT(CKCOMDBGridColumnsPpg)
	//}}AFX_DATA_INIT
	
	m_nColumnsBind = 0;
	m_pCtrl = NULL;
	for (int i = 0; i < 255; i ++)
	{
		lstrcpy(m_strColumnHeader[i], _T(""));
		m_nColumnBindNum[i] = -1;
	}
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridColumnsPpg::DoDataExchange - Moves data between page and properties

void CKCOMDBGridColumnsPpg::DoDataExchange(CDataExchange* pDX)
{
	// NOTE: ClassWizard will add DDP, DDX, and DDV calls here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//{{AFX_DATA_MAP(CKCOMDBGridColumnsPpg)
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);

	if (pDX->m_bSaveAndValidate)
	{
		m_pCtrl->m_nColumnsBind = 0;
		int nColumn = 0;
		
		for (int i = 0; i < m_pCtrl->m_apColumnData.GetSize(); i ++)
		{
			if (m_nColumnBindNum[i] >= 0)
			{
				m_pCtrl->m_nColumnBindNum[nColumn] = m_nColumnBindNum[i];
				lstrcpy(m_pCtrl->m_strColumnHeader[nColumn], m_strColumnHeader[i]);
				nColumn ++;
				m_pCtrl->m_nColumnsBind ++;
			}
		}

		m_pCtrl->SetModifiedFlag();
	}
	else
	{
		if (!m_pCtrl || m_pCtrl->m_nDataMode != 0 || m_nColumnsBind == 0)
		{
			GetDlgItem(IDC_COMBO_COLUMN)->EnableWindow(FALSE);
			GetDlgItem(IDC_EDIT_HEADER)->EnableWindow(FALSE);
			GetDlgItem(IDC_COMBO_DATAFIELD)->EnableWindow(FALSE);
			GetDlgItem(IDC_BUTTON_RESET)->EnableWindow(FALSE);
		}
		else
		{
			GetDlgItem(IDC_COMBO_COLUMN)->EnableWindow(TRUE);
			GetDlgItem(IDC_EDIT_HEADER)->EnableWindow(TRUE);
			GetDlgItem(IDC_COMBO_DATAFIELD)->EnableWindow(TRUE);
			GetDlgItem(IDC_BUTTON_RESET)->EnableWindow(TRUE);
		}
	}
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridColumnsPpg message handlers

BOOL CKCOMDBGridColumnsPpg::OnInitDialog() 
{
	COlePropertyPage::OnInitDialog();
	
	m_pComboColumn = (CComboBox *)GetDlgItem(IDC_COMBO_COLUMN);
	m_pEditHeader = (CEdit *)GetDlgItem(IDC_EDIT_HEADER);
	m_pComboField = (CComboBox *)GetDlgItem(IDC_COMBO_DATAFIELD);

	int i, nColumns;
	ULONG Ulong;
	LPDISPATCH FAR *m_lpDispatch = GetObjectArray(&Ulong);

	// Get the CCmdTarget object associated to any one of the above
	// array elements
	m_pCtrl = (CKCOMDBGridCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
	if (!m_pCtrl)
	{
		GetDlgItem(IDC_COMBO_COLUMN)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_HEADER)->EnableWindow(FALSE);
		GetDlgItem(IDC_COMBO_DATAFIELD)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_RESET)->EnableWindow(FALSE);
		return FALSE;
	}

	m_pCtrl->GetColumnInfo();
	nColumns = m_pCtrl->m_apColumnData.GetSize();
	if (m_pCtrl->m_nColumnsBind == 0)
	{
		m_nColumnsBind = nColumns;
		for (i = 0; i < m_nColumnsBind; i ++)
		{
			m_nColumnBindNum[i] = i;
			lstrcpy(m_strColumnHeader[i], m_pCtrl->m_apColumnData[i]->strName);
		}
	}
	else
	{
		m_nColumnsBind = m_pCtrl->m_nColumnsBind;
		for (i = 0; i < m_nColumnsBind; i ++)
		{
			m_nColumnBindNum[i] = m_pCtrl->m_nColumnBindNum[i];
			lstrcpy(m_strColumnHeader[i], m_pCtrl->m_strColumnHeader[i]);
		}
	}
	
	m_pComboColumn->ResetContent();
	CString strColumn;

	m_pComboField->AddString(_T(""));
	for (i = 0; i < nColumns; i ++)
	{
		strColumn.Format("Column%3d", i);
		m_pComboColumn->AddString(strColumn);
		m_pComboField->AddString(m_pCtrl->m_apColumnData[i]->strName);
	}

	if (m_nColumnsBind)
	{
		m_pComboColumn->SetCurSel(0);
		OnSelendokComboColumn();
	}

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CKCOMDBGridColumnsPpg::OnSelendokComboColumn() 
{
	int i = m_pComboColumn->GetCurSel();
	m_pComboField->SetCurSel(m_nColumnBindNum[i] + 1);
	m_pEditHeader->SetWindowText(m_strColumnHeader[i]);
}

void CKCOMDBGridColumnsPpg::OnSelendokComboDatafield() 
{
	int nColumn = m_pComboColumn->GetCurSel();
	int nField = m_pComboField->GetCurSel();
	
	m_nColumnBindNum[nColumn] = nField - 1;
	SetModifiedFlag();
}

void CKCOMDBGridColumnsPpg::OnChangeEditHeader() 
{
	int nColumn = m_pComboColumn->GetCurSel();

	m_pEditHeader->GetWindowText(m_strColumnHeader[nColumn], 39);
	SetModifiedFlag();
}

void CKCOMDBGridColumnsPpg::OnButtonReset() 
{
	for (int i = 0; i < m_pCtrl->m_apColumnData.GetSize(); i ++)
	{
		m_pComboColumn->SetCurSel(i);
		m_pComboField->SetCurSel(i + 1);
		m_pEditHeader->SetWindowText(m_pCtrl->m_apColumnData[i]->strName);
		OnSelendokComboDatafield();
	}
}
