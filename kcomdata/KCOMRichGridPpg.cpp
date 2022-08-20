// KCOMRichGridPpg.cpp : Implementation of the CKCOMRichGridPropPage property page class.

#include "stdafx.h"
#include "KCOMData.h"
#include "KCOMRichGridPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CKCOMRichGridPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CKCOMRichGridPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CKCOMRichGridPropPage)
	ON_CBN_SELENDOK(IDC_COMBO_DATAMODE, OnSelendokComboDatamode)
	ON_BN_CLICKED(IDC_CHECK_ALLOWUPDATE, OnCheckAllowupdate)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

// {8BACE102-16B9-11D3-B446-0080C8F18522}
IMPLEMENT_OLECREATE_EX(CKCOMRichGridPropPage, "KCOMRichGrid.KCOMRichGridPpg.1",
	0x8bace102, 0x16b9, 0x11d3, 0xb4, 0x46, 0x0, 0x80, 0xc8, 0xf1, 0x85, 0x22)


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropPage::CKCOMRichGridPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMRichGridPropPage

BOOL CKCOMRichGridPropPage::CKCOMRichGridPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Define string resource for page type; replace '0' below with ID.

	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_KCOMRICHGRID_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropPage::CKCOMRichGridPropPage - Constructor

// TODO: Define string resource for page caption; replace '0' below with ID.

CKCOMRichGridPropPage::CKCOMRichGridPropPage() :
	COlePropertyPage(IDD, IDS_KCOMRICHGRID_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CKCOMRichGridPropPage)
	m_nDataMode = -1;
	m_bAllowAddNew = FALSE;
	m_bAllowDelete = FALSE;
	m_bAllowUpdate = FALSE;
	m_bAllowRowSizing = FALSE;
	m_bAllowColMoving = FALSE;
	m_bAllowColSizing = FALSE;
	m_bColumnHeaders = FALSE;
	m_bRecordSelectors = FALSE;
	m_bSelectByCell = FALSE;
	m_strCaption = _T("");
	m_strFieldSeparator = _T("");
	m_nBorderStyle = -1;
	m_nCaptionAlignment = -1;
	m_nDividerStyle = -1;
	m_nDividerType = -1;
	m_nCols = 0;
	m_nDefColWidth = 0;
	m_nHeaderHeight = 0;
	m_nRowHeight = 0;
	m_nRows = 0;
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropPage::DoDataExchange - Moves data between page and properties

void CKCOMRichGridPropPage::DoDataExchange(CDataExchange* pDX)
{
	// NOTE: ClassWizard will add DDP, DDX, and DDV calls here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//{{AFX_DATA_MAP(CKCOMRichGridPropPage)
	DDP_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode, _T("DataMode") );
	DDX_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode);
	DDP_Check(pDX, IDC_CHECK_ALLOWUPDATE, m_bAllowUpdate, _T("AllowUpdate") );
	DDX_Check(pDX, IDC_CHECK_ALLOWUPDATE, m_bAllowUpdate);
	DDP_Check(pDX, IDC_CHECK_ALLOWADDNEW, m_bAllowAddNew, _T("AllowAddNew") );
	DDX_Check(pDX, IDC_CHECK_ALLOWADDNEW, m_bAllowAddNew);
	DDP_Check(pDX, IDC_CHECK_ALLOWDELETE, m_bAllowDelete, _T("AllowDelete") );
	DDX_Check(pDX, IDC_CHECK_ALLOWDELETE, m_bAllowDelete);
	DDP_Check(pDX, IDC_CHECK_ALLOWROWRESIZING, m_bAllowRowSizing, _T("AllowRowSizing") );
	DDX_Check(pDX, IDC_CHECK_ALLOWROWRESIZING, m_bAllowRowSizing);
	DDP_Check(pDX, IDC_CHECK_ALLOWCOLMOVING, m_bAllowColMoving, _T("AllowColumnMoving") );
	DDX_Check(pDX, IDC_CHECK_ALLOWCOLMOVING, m_bAllowColMoving);
	DDP_Check(pDX, IDC_CHECK_ALLOWCOLSIZING, m_bAllowColSizing, _T("AllowColumnSizing") );
	DDX_Check(pDX, IDC_CHECK_ALLOWCOLSIZING, m_bAllowColSizing);
	DDP_Check(pDX, IDC_CHECK_COLUMNHEADERS, m_bColumnHeaders, _T("ColumnHeaders") );
	DDX_Check(pDX, IDC_CHECK_COLUMNHEADERS, m_bColumnHeaders);
	DDP_Check(pDX, IDC_CHECK_SHOWRECORDSELECTORS, m_bRecordSelectors, _T("RecordSelectors") );
	DDX_Check(pDX, IDC_CHECK_SHOWRECORDSELECTORS, m_bRecordSelectors);
	DDP_Check(pDX, IDC_CHECK_SELECTBYCELL, m_bSelectByCell, _T("SelectByCell") );
	DDX_Check(pDX, IDC_CHECK_SELECTBYCELL, m_bSelectByCell);
	DDP_Text(pDX, IDC_EDIT_CAPTION, m_strCaption, _T("Caption") );
	DDX_Text(pDX, IDC_EDIT_CAPTION, m_strCaption);
	DDP_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator, _T("FieldSeparator") );
	DDX_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator);
	DDP_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_nBorderStyle, _T("BorderStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_nBorderStyle);
	DDP_CBIndex(pDX, IDC_COMBO_CAPTIONALIGNMENT, m_nCaptionAlignment, _T("CaptionAlignment") );
	DDX_CBIndex(pDX, IDC_COMBO_CAPTIONALIGNMENT, m_nCaptionAlignment);
	DDP_CBIndex(pDX, IDC_COMBO_DIVIDERSTYLE, m_nDividerStyle, _T("DividerStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_DIVIDERSTYLE, m_nDividerStyle);
	DDP_CBIndex(pDX, IDC_COMBO_DIVIDERTYPE, m_nDividerType, _T("DividerType") );
	DDX_CBIndex(pDX, IDC_COMBO_DIVIDERTYPE, m_nDividerType);
	DDP_Text(pDX, IDC_EDIT_COLS, m_nCols, _T("Cols") );
	DDX_Text(pDX, IDC_EDIT_COLS, m_nCols);
	DDV_MinMaxLong(pDX, m_nCols, 0, 255);
	DDP_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth, _T("DefColWidth") );
	DDX_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth);
	DDP_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight, _T("HeaderHeight") );
	DDX_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight);
	DDP_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight, _T("RowHeight") );
	DDX_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight);
	DDP_Text(pDX, IDC_EDIT_ROWS, m_nRows, _T("Rows") );
	DDX_Text(pDX, IDC_EDIT_ROWS, m_nRows);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);

	if (!pDX->m_bSaveAndValidate)
		UpdateControls();
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropPage message handlers

void CKCOMRichGridPropPage::OnSelendokComboDatamode() 
{
	m_nDataMode = ((CComboBox *)GetDlgItem(IDC_COMBO_DATAMODE))->GetCurSel();
	UpdateControls();
}

void CKCOMRichGridPropPage::OnCheckAllowupdate() 
{
	m_bAllowUpdate = ((CButton *)GetDlgItem(IDC_CHECK_ALLOWUPDATE))->GetCheck();
	UpdateControls();
}

void CKCOMRichGridPropPage::UpdateControls()
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

	if (!m_bAllowUpdate)
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
