// KCOMRichComboPpg.cpp : Implementation of the CKCOMRichComboPropPage property page class.

#include "stdafx.h"
#include "KCOMData.h"
#include "KCOMRichComboPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CKCOMRichComboPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CKCOMRichComboPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CKCOMRichComboPropPage)
	// NOTE - ClassWizard will add and remove message map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMRichComboPropPage, "KCOMDATA.KCOMRichComboPropPage.1",
	0x29288ca, 0x2fff, 0x11d3, 0xb4, 0x46, 0, 0x80, 0xc8, 0xf1, 0x85, 0x22)


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboPropPage::CKCOMRichComboPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMRichComboPropPage

BOOL CKCOMRichComboPropPage::CKCOMRichComboPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_KCOMRICHCOMBO_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboPropPage::CKCOMRichComboPropPage - Constructor

CKCOMRichComboPropPage::CKCOMRichComboPropPage() :
	COlePropertyPage(IDD, IDS_KCOMRICHCOMBO_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CKCOMRichComboPropPage)
	m_bReadOnly = FALSE;
	m_bColumnHeaders = FALSE;
	m_bListAutoPosition = FALSE;
	m_bListWidthAutoSize = FALSE;
	m_nDividerStyle = -1;
	m_nDividerType = -1;
	m_nBorderStyle = -1;
	m_nDataMode = -1;
	m_strFieldSeparator = _T("");
	m_nRows = 0;
	m_nCols = 0;
	m_n = 0;
	m_nMinDropDownItems = 0;
	m_nDefColWidth = 0;
	m_nHeaderHeight = 0;
	m_nRowHeight = 0;
	m_nListWidth = 0;
	m_strFormat = _T("");
	m_nMaxLength = 0;
	m_nStyle = -1;
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboPropPage::DoDataExchange - Moves data between page and properties

void CKCOMRichComboPropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CKCOMRichComboPropPage)
	DDP_Check(pDX, IDC_CHECK_READONLY, m_bReadOnly, _T("ReadOnly") );
	DDX_Check(pDX, IDC_CHECK_READONLY, m_bReadOnly);
	DDP_Check(pDX, IDC_CHECK_COLUMNHEADERS, m_bColumnHeaders, _T("ColumnHeaders") );
	DDX_Check(pDX, IDC_CHECK_COLUMNHEADERS, m_bColumnHeaders);
	DDP_Check(pDX, IDC_CHECK_LISTAUTOPOSITION, m_bListAutoPosition, _T("ListAutoPosition") );
	DDX_Check(pDX, IDC_CHECK_LISTAUTOPOSITION, m_bListAutoPosition);
	DDP_Check(pDX, IDC_CHECK_LISTWIDTHAUTOSIZE, m_bListWidthAutoSize, _T("ListWidthAutoSize") );
	DDX_Check(pDX, IDC_CHECK_LISTWIDTHAUTOSIZE, m_bListWidthAutoSize);
	DDP_CBIndex(pDX, IDC_COMBO_DIVIDERSTYLE, m_nDividerStyle, _T("DividerStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_DIVIDERSTYLE, m_nDividerStyle);
	DDP_CBIndex(pDX, IDC_COMBO_DIVIDERTYPE, m_nDividerType, _T("DividerType") );
	DDX_CBIndex(pDX, IDC_COMBO_DIVIDERTYPE, m_nDividerType);
	DDP_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_nBorderStyle, _T("BorderStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_nBorderStyle);
	DDP_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode, _T("DataMode") );
	DDX_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode);
	DDP_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator, _T("FieldSeparator") );
	DDX_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator);
	DDP_Text(pDX, IDC_EDIT_ROWS, m_nRows, _T("Rows") );
	DDX_Text(pDX, IDC_EDIT_ROWS, m_nRows);
	DDP_Text(pDX, IDC_EDIT_COLS, m_nCols, _T("Cols") );
	DDX_Text(pDX, IDC_EDIT_COLS, m_nCols);
	DDP_Text(pDX, IDC_EDIT_MAXDROPDOWNITEMS, m_n, _T("MaxDropDownItems") );
	DDX_Text(pDX, IDC_EDIT_MAXDROPDOWNITEMS, m_n);
	DDP_Text(pDX, IDC_EDIT_MINDROPDOWNITEMS, m_nMinDropDownItems, _T("MinDropDownItems") );
	DDX_Text(pDX, IDC_EDIT_MINDROPDOWNITEMS, m_nMinDropDownItems);
	DDP_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth, _T("DefColWidth") );
	DDX_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth);
	DDP_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight, _T("HeaderHeight") );
	DDX_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight);
	DDP_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight, _T("RowHeight") );
	DDX_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight);
	DDP_Text(pDX, IDC_EDIT_LISTWIDTH, m_nListWidth, _T("ListWidth") );
	DDX_Text(pDX, IDC_EDIT_LISTWIDTH, m_nListWidth);
	DDP_Text(pDX, IDC_EDIT_FORMAT, m_strFormat, _T("Format") );
	DDX_Text(pDX, IDC_EDIT_FORMAT, m_strFormat);
	DDP_Text(pDX, IDC_EDIT_MAXLENGTH, m_nMaxLength, _T("MaxLength") );
	DDX_Text(pDX, IDC_EDIT_MAXLENGTH, m_nMaxLength);
	DDP_CBIndex(pDX, IDC_COMBO_STYLE, m_nStyle, _T("Style") );
	DDX_CBIndex(pDX, IDC_COMBO_STYLE, m_nStyle);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboPropPage message handlers
