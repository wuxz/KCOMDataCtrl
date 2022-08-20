// KCOMMaskPpg.cpp : Implementation of the CKCOMMaskPropPage property page class.

#include "stdafx.h"
#include "KCOMMask.h"
#include "KCOMMaskPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CKCOMMaskPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CKCOMMaskPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CKCOMMaskPropPage)
	// NOTE - ClassWizard will add and remove message map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMMaskPropPage, "KCOMMASK.KCOMMaskPropPage.1",
	0xd1728e26, 0x4985, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskPropPage::CKCOMMaskPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMMaskPropPage

BOOL CKCOMMaskPropPage::CKCOMMaskPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_KCOMMASK_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskPropPage::CKCOMMaskPropPage - Constructor

CKCOMMaskPropPage::CKCOMMaskPropPage() :
	COlePropertyPage(IDD, IDS_KCOMMASK_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CKCOMMaskPropPage)
	m_bAutoTab = FALSE;
	m_bPromptInclude = FALSE;
	m_bHideSelection = FALSE;
	m_strMask = _T("");
	m_sAppearance = -1;
	m_sBorderStyle = -1;
	m_nMaxLength = 0;
	m_strPromptChar = _T("");
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskPropPage::DoDataExchange - Moves data between page and properties

void CKCOMMaskPropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CKCOMMaskPropPage)
	DDP_Check(pDX, IDC_CHECK_AUTOTATB, m_bAutoTab, _T("AutoTab") );
	DDX_Check(pDX, IDC_CHECK_AUTOTATB, m_bAutoTab);
	DDP_Check(pDX, IDC_CHECK_PROMPTINCLUDE, m_bPromptInclude, _T("PromptInclude") );
	DDX_Check(pDX, IDC_CHECK_PROMPTINCLUDE, m_bPromptInclude);
	DDP_Text(pDX, IDC_EDIT_MASK, m_strMask, _T("Mask") );
	DDX_Text(pDX, IDC_EDIT_MASK, m_strMask);
	DDP_CBIndex(pDX, IDC_COMBO_APPEARANCE, m_sAppearance, _T("Appearance") );
	DDX_CBIndex(pDX, IDC_COMBO_APPEARANCE, m_sAppearance);
	DDP_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_sBorderStyle, _T("BorderStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_sBorderStyle);
	DDP_Text(pDX, IDC_EDIT_MAXLENGTH, m_nMaxLength, _T("MaxLength") );
	DDX_Text(pDX, IDC_EDIT_MAXLENGTH, m_nMaxLength);
	DDV_MinMaxInt(pDX, m_nMaxLength, 1, 255);
	DDP_Text(pDX, IDC_EDIT_PROMPTCHAR, m_strPromptChar, _T("PromptChar") );
	DDX_Text(pDX, IDC_EDIT_PROMPTCHAR, m_strPromptChar);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMMaskPropPage message handlers
