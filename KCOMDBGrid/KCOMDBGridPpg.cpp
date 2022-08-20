// KCOMDBGridPpg.cpp : Implementation of the CKCOMDBGridPropPage property page class.

#include "stdafx.h"
#include "KCOMDBGrid.h"
#include "KCOMDBGridPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CKCOMDBGridPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CKCOMDBGridPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CKCOMDBGridPropPage)
	ON_CBN_SELENDOK(IDC_COMBO_DATAMODE, OnSelendokListDatamode)
	ON_BN_CLICKED(IDC_CHECK_READONLY, OnCheckReadonly)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMDBGridPropPage, "KCOMDBGRID.KCOMDBGridPropPage.1",
	0xac212647, 0xfbe2, 0x11d2, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridPropPage::CKCOMDBGridPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMDBGridPropPage

BOOL CKCOMDBGridPropPage::CKCOMDBGridPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_KCOMDBGRID_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridPropPage::CKCOMDBGridPropPage - Constructor

CKCOMDBGridPropPage::CKCOMDBGridPropPage() :
	COlePropertyPage(IDD, IDS_KCOMDBGRID_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CKCOMDBGridPropPage)
	m_nDataMode = -1;
	m_bShowRecordSelector = FALSE;
	m_bAllowAddNew = FALSE;
	m_bAllowDelete = FALSE;
	m_bReadOnly = FALSE;
	m_bAllowArrowKeys = FALSE;
	m_bAllowRowSizing = FALSE;
	m_nDefColWidth = 0;
	m_nHeaderHeight = 0;
	m_nRowHeight = 0;
	m_nGridLineStyle = -1;
	m_nRows = 0;
	m_nCols = 0;
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridPropPage::DoDataExchange - Moves data between page and properties

void CKCOMDBGridPropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CKCOMDBGridPropPage)
	DDP_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode, _T("DataMode") );
	DDX_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode);
	DDP_Check(pDX, IDC_CHECK_SHOWRECORDSELECTOR, m_bShowRecordSelector, _T("ShowRecordSelector") );
	DDX_Check(pDX, IDC_CHECK_SHOWRECORDSELECTOR, m_bShowRecordSelector);
	DDP_Check(pDX, IDC_CHECK_ALLOWADDNEW, m_bAllowAddNew, _T("AllowAddNew") );
	DDX_Check(pDX, IDC_CHECK_ALLOWADDNEW, m_bAllowAddNew);
	DDP_Check(pDX, IDC_CHECK_ALLOWDELETE, m_bAllowDelete, _T("AllowDelete") );
	DDX_Check(pDX, IDC_CHECK_ALLOWDELETE, m_bAllowDelete);
	DDP_Check(pDX, IDC_CHECK_READONLY, m_bReadOnly, _T("ReadOnly") );
	DDX_Check(pDX, IDC_CHECK_READONLY, m_bReadOnly);
	DDP_Check(pDX, IDC_CHECK_ALLOWARROWKEYS, m_bAllowArrowKeys, _T("AllowArrowKeys") );
	DDX_Check(pDX, IDC_CHECK_ALLOWARROWKEYS, m_bAllowArrowKeys);
	DDP_Check(pDX, IDC_CHECK_ALLOWAROWRESIZING, m_bAllowRowSizing, _T("AllowRowSizing") );
	DDX_Check(pDX, IDC_CHECK_ALLOWAROWRESIZING, m_bAllowRowSizing);
	DDP_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth, _T("DefColWidth") );
	DDX_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth);
	DDP_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight, _T("HeaderHeight") );
	DDX_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight);
	DDP_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight, _T("RowHeight") );
	DDX_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight);
	DDP_CBIndex(pDX, IDC_COMBO_GRIDLINESTYLE, m_nGridLineStyle, _T("GridLineStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_GRIDLINESTYLE, m_nGridLineStyle);
	DDP_Text(pDX, IDC_EDIT_ROWS, m_nRows, _T("Rows") );
	DDX_Text(pDX, IDC_EDIT_ROWS, m_nRows);
	DDV_MinMaxLong(pDX, m_nRows, 0, 65536);
	DDP_Text(pDX, IDC_EDIT_COLS, m_nCols, _T("Cols") );
	DDX_Text(pDX, IDC_EDIT_COLS, m_nCols);
	DDV_MinMaxLong(pDX, m_nCols, 0, 256);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
	
	if (!pDX->m_bSaveAndValidate)
		UpdateControls();
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridPropPage message handlers

void CKCOMDBGridPropPage::OnSelendokListDatamode() 
{
	m_nDataMode = ((CComboBox *)GetDlgItem(IDC_COMBO_DATAMODE))->GetCurSel();
	UpdateControls();
}

void CKCOMDBGridPropPage::OnCheckReadonly() 
{
	m_bReadOnly = ((CButton *)GetDlgItem(IDC_CHECK_READONLY))->	GetCheck();
	UpdateControls();
}

void CKCOMDBGridPropPage::UpdateControls()
{
	if (m_nDataMode == 0)
	{
		GetDlgItem(IDC_SPIN_ROWS)->EnableWindow(FALSE);
		GetDlgItem(IDC_SPIN_COLS)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_ROWS)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_COLS)->EnableWindow(FALSE);
	}
	else
	{
		GetDlgItem(IDC_SPIN_ROWS)->EnableWindow(TRUE);
		GetDlgItem(IDC_SPIN_COLS)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_ROWS)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_COLS)->EnableWindow(TRUE);
	}

	if (m_bReadOnly)
	{
		((CButton *)GetDlgItem(IDC_CHECK_ALLOWADDNEW))->SetCheck(0);
		((CButton *)GetDlgItem(IDC_CHECK_ALLOWDELETE))->SetCheck(0);
		GetDlgItem(IDC_CHECK_ALLOWADDNEW)->EnableWindow(FALSE);
		GetDlgItem(IDC_CHECK_ALLOWDELETE)->EnableWindow(FALSE);
	}
	else
	{
		GetDlgItem(IDC_CHECK_ALLOWADDNEW)->EnableWindow(TRUE);
		GetDlgItem(IDC_CHECK_ALLOWDELETE)->EnableWindow(TRUE);
	}
}
