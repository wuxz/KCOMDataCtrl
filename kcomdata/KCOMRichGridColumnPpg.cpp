// KCOMRichGridPropPage.cpp : Implementation of the CKCOMRichGridPropColumnPage property page class.

#include "stdafx.h"
#include "kcomdata.h"
#include "KCOMRichGridColumnPpg.h"
#include "KCOMRichGridCtl.h"
#include "GridInputDlg.h"
#include "AddColumnDlg.h"
#include "KCOMEditControl.h"
#include "KCOMComboBox.h"
#include "KCOMDateControl.h"
#include "KCOMMaskControl.h"
#include "KCOMEditBtn.h"
#include "GridInner.h"
#include "FieldsDlg.h"
#include "KCOMRichDropDownCtl.h"
#include "DropDownRealWnd.h"
#include "KCOMRichComboCtl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define FRAMELEFT 325
#define FRAMETOP 35

IMPLEMENT_DYNCREATE(CKCOMRichGridPropColumnPage, COlePropertyPage)

#define pComboColumn ((CComboBox *)GetDlgItem(IDC_COMBO_COLUMN))
#define pStaticFrame GetDlgItem(IDC_STATIC_FRAME)
#define pListProperty ((CListBox *)GetDlgItem(IDC_LIST_PROPERTY))
#define pFields GetDlgItem(IDC_BUTTON_FIELDS)
#define pDeleteColumn GetDlgItem(IDC_BUTTON_DELETECOLUMN)

#define	pAlignLeft ((CButton *)GetDlgItem(IDC_RADIO_ALIGNMENT_LEFT))
#define	pAlignRight ((CButton *)GetDlgItem(IDC_RADIO_ALIGNMENT_RIGHT))
#define	pAlignCenter ((CButton *)GetDlgItem(IDC_RADIO_ALIGNMENT_CENTER))
#define pAllowSizing ((CButton *)GetDlgItem(IDC_CHECK_ALLOWSIZING))

#define	pCaption GetDlgItem(IDC_EDIT_CAPTION)

#define pCapAlignLeft ((CButton *)GetDlgItem(IDC_RADIO_CAPTIONALIGNMENT_LEFT))
#define pCapAlignRight ((CButton *)GetDlgItem(IDC_RADIO_CAPTIONALIGNMENT_RIGHT))
#define pCapAlignCenter ((CButton *)GetDlgItem(IDC_RADIO_CAPTIONALIGNMENT_CENTER))
#define pCapAlignAsAlignment ((CButton *)GetDlgItem(IDC_RADIO_CAPTIONALIGNMENT_ASALIGNMENT))

#define pCaseNone ((CButton *)GetDlgItem(IDC_RADIO_CASE_NONE))
#define pCaseUpper ((CButton *)GetDlgItem(IDC_RADIO_CASE_UPPER))
#define pCaseLower ((CButton *)GetDlgItem(IDC_RADIO_CASE_LOWER))

#define pDataField ((CComboBox *)GetDlgItem(IDC_COMBO_DATAFIELD))

#define pComboBoxStyle ((CComboBox *)GetDlgItem(IDC_COMBO_COMBOBOXSTYLE))

#define	pBackColor GetDlgItem(IDC_BUTTON_BACKCOLOR)
#define	pForeColor GetDlgItem(IDC_BUTTON_FORECOLOR)
#define pHeadBackColor GetDlgItem(IDC_BUTTON_HEADBACKCOLOR)
#define pHeadForeColor GetDlgItem(IDC_BUTTON_HEADFORECOLOR)

#define	pFieldLen GetDlgItem(IDC_EDIT_FIELDLEN)

#define pLocked ((CButton *)GetDlgItem(IDC_CHECK_LOCKED))

#define pText ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_TEXT))
#define pBoolean ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_BOOLEAN))
#define pByte ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_BYTE))
#define pInteger ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_INTEGER))
#define pLong ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_LONG))
#define pSingle ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_SINGLE))
#define pCurrency ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_CURRENCY))
#define pDate ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_DATE))
#define pDateMask GetDlgItem(IDC_COMBO_DATEMASK)
#define pStaticDateMask GetDlgItem(IDC_STATIC_DATEMASK)

#define	pMask GetDlgItem(IDC_EDIT_MASK)

#define pName GetDlgItem(IDC_EDIT_NAME)

#define pNullable ((CButton *)GetDlgItem(IDC_CHECK_NULLABLE))

#define pPromptChar GetDlgItem(IDC_EDIT_PROMPTCHAR)

#define pStyEdit ((CButton *)GetDlgItem(IDC_RADIO_STYLE_EDIT))
#define pStyEditButton ((CButton *)GetDlgItem(IDC_RADIO_STYLE_EDITBUTTON))
#define pStyCheckBox ((CButton *)GetDlgItem(IDC_RADIO_STYLE_CHECKBOX))
#define pStyComboBox ((CButton *)GetDlgItem(IDC_RADIO_STYLE_COMBOBOX))
#define pSetupComboBox (GetDlgItem(IDC_BUTTON_SETUPCOMBOBOX))
#define pStyButton ((CButton *)GetDlgItem(IDC_RADIO_STYLE_BUTTON))

#define pVisible ((CButton *)GetDlgItem(IDC_CHECK_VISIBLE))

#define pWidth GetDlgItem(IDC_EDIT_WIDTH)

#define pPromptInclude ((CButton *)GetDlgItem(IDC_CHECK_PROMPTINCLUDE))
/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CKCOMRichGridPropColumnPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CKCOMRichGridPropColumnPage)
	ON_LBN_SELCHANGE(IDC_LIST_PROPERTY, OnSelchangeListProperty)
	ON_CBN_SELENDOK(IDC_COMBO_COLUMN, OnSelendokComboColumn)
	ON_BN_CLICKED(IDC_BUTTON_BACKCOLOR, OnButtonBackcolor)
	ON_BN_CLICKED(IDC_BUTTON_FORECOLOR, OnButtonForecolor)
	ON_BN_CLICKED(IDC_BUTTON_HEADBACKCOLOR, OnButtonHeadbackcolor)
	ON_BN_CLICKED(IDC_BUTTON_HEADFORECOLOR, OnButtonHeadforecolor)
	ON_BN_CLICKED(IDC_CHECK_ALLOWSIZING, OnCheckAllowsizing)
	ON_BN_CLICKED(IDC_CHECK_LOCKED, OnCheckLocked)
	ON_BN_CLICKED(IDC_CHECK_NULLABLE, OnCheckNullable)
	ON_BN_CLICKED(IDC_CHECK_VISIBLE, OnCheckVisible)
	ON_EN_KILLFOCUS(IDC_EDIT_NAME, OnKillfocusEditName)
	ON_EN_KILLFOCUS(IDC_EDIT_CAPTION, OnKillfocusEditCaption)
	ON_EN_KILLFOCUS(IDC_EDIT_FIELDLEN, OnKillfocusEditFieldlen)
	ON_EN_KILLFOCUS(IDC_EDIT_MASK, OnKillfocusEditMask)
	ON_EN_KILLFOCUS(IDC_EDIT_PROMPTCHAR, OnKillfocusEditPromptchar)
	ON_EN_KILLFOCUS(IDC_EDIT_WIDTH, OnKillfocusEditWidth)
	ON_BN_CLICKED(IDC_RADIO_ALIGNMENT_CENTER, OnRadioAlignmentCenter)
	ON_BN_CLICKED(IDC_RADIO_ALIGNMENT_LEFT, OnRadioAlignmentLeft)
	ON_BN_CLICKED(IDC_RADIO_ALIGNMENT_RIGHT, OnRadioAlignmentRight)
	ON_BN_CLICKED(IDC_RADIO_CAPTIONALIGNMENT_ASALIGNMENT, OnRadioCaptionalignmentAsalignment)
	ON_BN_CLICKED(IDC_RADIO_CAPTIONALIGNMENT_CENTER, OnRadioCaptionalignmentCenter)
	ON_BN_CLICKED(IDC_RADIO_CAPTIONALIGNMENT_LEFT, OnRadioCaptionalignmentLeft)
	ON_BN_CLICKED(IDC_RADIO_CAPTIONALIGNMENT_RIGHT, OnRadioCaptionalignmentRight)
	ON_BN_CLICKED(IDC_RADIO_CASE_LOWER, OnRadioCaseLower)
	ON_BN_CLICKED(IDC_RADIO_CASE_NONE, OnRadioCaseNone)
	ON_BN_CLICKED(IDC_RADIO_CASE_UPPER, OnRadioCaseUpper)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_BOOLEAN, OnRadioDatatypeBoolean)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_BYTE, OnRadioDatatypeByte)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_CURRENCY, OnRadioDatatypeCurrency)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_DATE, OnRadioDatatypeDate)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_INTEGER, OnRadioDatatypeInteger)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_LONG, OnRadioDatatypeLong)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_SINGLE, OnRadioDatatypeSingle)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_TEXT, OnRadioDatatypeText)
	ON_BN_CLICKED(IDC_RADIO_STYLE_BUTTON, OnRadioStyleButton)
	ON_BN_CLICKED(IDC_RADIO_STYLE_CHECKBOX, OnRadioStyleCheckbox)
	ON_BN_CLICKED(IDC_RADIO_STYLE_COMBOBOX, OnRadioStyleCombobox)
	ON_BN_CLICKED(IDC_RADIO_STYLE_EDIT, OnRadioStyleEdit)
	ON_BN_CLICKED(IDC_RADIO_STYLE_EDITBUTTON, OnRadioStyleEditbutton)
	ON_BN_CLICKED(IDC_BUTTON_SETUPCOMBOBOX, OnButtonSetupcombobox)
	ON_BN_CLICKED(IDC_BUTTON_ADDCOLUMN, OnButtonAddcolumn)
	ON_BN_CLICKED(IDC_BUTTON_DELETECOLUMN, OnButtonDeletecolumn)
	ON_BN_CLICKED(IDC_BUTTON_RESET, OnButtonReset)
	ON_BN_CLICKED(IDC_CHECK_PROMPTINCLUDE, OnCheckPromptinclude)
	ON_CBN_SELENDOK(IDC_COMBO_DATAFIELD, OnSelendokComboDatafield)
	ON_BN_CLICKED(IDC_BUTTON_FIELDS, OnButtonFields)
	ON_CBN_SELENDOK(IDC_COMBO_DATEMASK, OnSelendokComboDatemask)
	ON_CBN_SELENDOK(IDC_COMBO_COMBOBOXSTYLE, OnSelendokComboComboboxstyle)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMRichGridPropColumnPage, "KCOMRICHGRID.KCOMRichGridColumnPropPage.1",
	0xcfc2326, 0x12a9, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropColumnPage::CKCOMRichGridPropColumnPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMRichGridPropColumnPage

BOOL CKCOMRichGridPropColumnPage::CKCOMRichGridPropColumnPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_KCOMRICHGRID_COLUMNS_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropColumnPage::CKCOMRichGridPropColumnPage - Constructor

CKCOMRichGridPropColumnPage::CKCOMRichGridPropColumnPage() :
	COlePropertyPage(IDD, IDS_KCOMRICHGRID_COLUMNS_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CKCOMRichGridPropColumnPage)
	//}}AFX_DATA_INIT
	m_pGridCtrl = NULL;
	m_pDropDownCtrl = NULL;
	m_pComboCtrl = NULL;
	m_wndGrid.SetPagePtr(this);
	m_nCurrentCol = -1;
	m_nColumns = 0;
}

CKCOMRichGridPropColumnPage::~CKCOMRichGridPropColumnPage()
{
	for (int i = 0; i < m_apColumnProp.GetSize(); i ++)
		delete m_apColumnProp[i];

	m_apColumnProp.RemoveAll();	
}
/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropColumnPage::DoDataExchange - Moves data between page and properties

void CKCOMRichGridPropColumnPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CKCOMRichGridPropColumnPage)
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);

	if (!pDX->m_bSaveAndValidate)
	{
		if (m_pGridCtrl)
		{
			m_wndGrid.ChangeStandardStyle(CGXStyle().SetInterior(m_pGridCtrl->TranslateColor(m_pGridCtrl->GetBackColor())));
			m_wndGrid.ChangeStandardStyle(CGXStyle().SetTextColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->GetForeColor())));
			m_wndGrid.RowHeaderStyle().SetTextColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->m_clrHeadFore));
			m_wndGrid.RowHeaderStyle().SetInterior(m_pGridCtrl->TranslateColor(m_pGridCtrl->m_clrHeadBack));
			m_wndGrid.ColHeaderStyle().SetTextColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->m_clrHeadFore));
			m_wndGrid.ColHeaderStyle().SetInterior(m_pGridCtrl->TranslateColor(m_pGridCtrl->m_clrHeadBack));
			OnSelchangeListProperty();
		}
		else if (m_pDropDownCtrl)
		{
			m_wndGrid.ChangeStandardStyle(CGXStyle().SetInterior(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->GetBackColor())));
			m_wndGrid.ChangeStandardStyle(CGXStyle().SetTextColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->GetForeColor())));
			m_wndGrid.RowHeaderStyle().SetTextColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->m_clrHeadFore));
			m_wndGrid.RowHeaderStyle().SetInterior(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->m_clrHeadBack));
			m_wndGrid.ColHeaderStyle().SetTextColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->m_clrHeadFore));
			m_wndGrid.ColHeaderStyle().SetInterior(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->m_clrHeadBack));
			OnSelchangeListProperty();
		}
		else if (m_pComboCtrl)
		{
			m_wndGrid.ChangeStandardStyle(CGXStyle().SetInterior(m_pComboCtrl->TranslateColor(m_pComboCtrl->GetBackColor())));
			m_wndGrid.ChangeStandardStyle(CGXStyle().SetTextColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->GetForeColor())));
			m_wndGrid.RowHeaderStyle().SetTextColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->m_clrHeadFore));
			m_wndGrid.RowHeaderStyle().SetInterior(m_pComboCtrl->TranslateColor(m_pComboCtrl->m_clrHeadBack));
			m_wndGrid.ColHeaderStyle().SetTextColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->m_clrHeadFore));
			m_wndGrid.ColHeaderStyle().SetInterior(m_pComboCtrl->TranslateColor(m_pComboCtrl->m_clrHeadBack));
			OnSelchangeListProperty();
		}
	}
	else
	{
		int nProperty = CalcPropertyOrdinal(pListProperty->GetCurSel());
		
		if (nProperty == 3)
			OnKillfocusEditCaption();
		else if(nProperty == 8)
			OnKillfocusEditFieldlen();
		else if(nProperty == 13)
			OnKillfocusEditMask();
		else if(nProperty == 14)
			OnKillfocusEditName();
		else if(nProperty == 16)
			OnKillfocusEditPromptchar();
		else if(nProperty == 19)
			OnKillfocusEditWidth();

		int i, j;
		int nCols = m_apColumnProp.GetSize();
		int nRows = m_wndGrid.GetRowCount();

		if (m_pGridCtrl)
		{
			for (i = 0; i < m_pGridCtrl->m_apColumnsProp.GetSize(); i ++)
				delete m_pGridCtrl->m_apColumnsProp[i];
			
			m_pGridCtrl->m_apColumnsProp.RemoveAll();
			m_pGridCtrl->m_arContent.RemoveAll();
			
			ColumnProp * pCol, * pColNew;
			m_pGridCtrl->m_pGridInner->SetRowCount(nRows - 1);
			
			for (i = 0; i < nCols; i ++)
			{
				pColNew = new ColumnProp;
				pCol = m_apColumnProp[m_wndGrid.GetFieldFromIndex(m_wndGrid.
					GetColIndex(i + 1))];
				
				pColNew->bAllowSizing = pCol->bAllowSizing;
				pColNew->bForceLock = pCol->bForceLock;
				pColNew->bForceNoNullable = pCol->bForceNoNullable;
				pColNew->bLocked = pCol->bLocked;
				pColNew->bNullable = pCol->bNullable;
				pColNew->bVisible = pCol->bVisible;
				pColNew->clrBack = pCol->clrBack;
				pColNew->clrFore = pCol->clrFore;
				pColNew->clrHeadBack = pCol->clrHeadBack;
				pColNew->clrHeadFore = pCol->clrHeadFore;
				pColNew->nAlignment = pCol->nAlignment;
				pColNew->nCaptionAlignment = pCol->nCaptionAlignment;
				pColNew->nCase = pCol->nCase;
				pColNew->nDataType = pCol->nDataType;
				pColNew->nFieldLen = pCol->nFieldLen;
				pColNew->nStyle = pCol->nStyle;
				pColNew->nWidth = pCol->nWidth;
				pColNew->strCaption = pCol->strCaption;
				pColNew->strChoiceList = pCol->strChoiceList;
				pColNew->strDataField = pCol->strDataField;
				pColNew->strMask = pCol->strMask;
				pColNew->strName = pCol->strName;
				pColNew->strPromptChar = pCol->strPromptChar;
				pColNew->bPromptInclude = pCol->bPromptInclude;
				pColNew->nComboBoxStyle = pCol->nComboBoxStyle;
				
				m_pGridCtrl->m_apColumnsProp.Add(pColNew);
				
				for (j = 1; j < nRows; j ++)
					m_pGridCtrl->m_arContent.Add(m_wndGrid.GetValueRowCol(
					j, i + 1));
			}
			
			m_pGridCtrl->InitGridFromProp();
		}
		else if (m_pDropDownCtrl)
		{
			for (i = 0; i < m_pDropDownCtrl->m_apColumnsProp.GetSize(); i ++)
				delete m_pDropDownCtrl->m_apColumnsProp[i];
			
			m_pDropDownCtrl->m_apColumnsProp.RemoveAll();
			m_pDropDownCtrl->m_arContent.RemoveAll();
			
			ColumnProp * pCol, * pColNew;

			if (m_pDropDownCtrl->m_pDrawGrid)
				m_pDropDownCtrl->m_pDrawGrid->SetRowCount(nRows - 1);
			else if (m_pDropDownCtrl->m_pDropDownRealWnd)
				m_pDropDownCtrl->m_pDropDownRealWnd->SetRowCount(nRows - 1);
			
			for (i = 0; i < nCols; i ++)
			{
				pColNew = new ColumnProp;
				pCol = m_apColumnProp[m_wndGrid.GetFieldFromIndex(m_wndGrid.
					GetColIndex(i + 1))];
				
				pColNew->bVisible = pCol->bVisible;
				pColNew->clrBack = pCol->clrBack;
				pColNew->clrFore = pCol->clrFore;
				pColNew->clrHeadBack = pCol->clrHeadBack;
				pColNew->clrHeadFore = pCol->clrHeadFore;
				pColNew->nAlignment = pCol->nAlignment;
				pColNew->nCaptionAlignment = pCol->nCaptionAlignment;
				pColNew->nCase = pCol->nCase;
				pColNew->nWidth = pCol->nWidth;
				pColNew->strCaption = pCol->strCaption;
				pColNew->strDataField = pCol->strDataField;
				pColNew->strName = pCol->strName;
				
				m_pDropDownCtrl->m_apColumnsProp.Add(pColNew);
				
				for (j = 1; j < nRows; j ++)
					m_pDropDownCtrl->m_arContent.Add(m_wndGrid.GetValueRowCol(
					j, i + 1));
			}
			
			m_pDropDownCtrl->InitGridFromProp();
		}
		else if (m_pComboCtrl)
		{
			for (i = 0; i < m_pComboCtrl->m_apColumnsProp.GetSize(); i ++)
				delete m_pComboCtrl->m_apColumnsProp[i];
			
			m_pComboCtrl->m_apColumnsProp.RemoveAll();
			m_pComboCtrl->m_arContent.RemoveAll();
			
			ColumnProp * pCol, * pColNew;

			if (m_pComboCtrl->m_pDropDownRealWnd)
				m_pComboCtrl->m_pDropDownRealWnd->SetRowCount(nRows - 1);
			
			for (i = 0; i < nCols; i ++)
			{
				pColNew = new ColumnProp;
				pCol = m_apColumnProp[m_wndGrid.GetFieldFromIndex(m_wndGrid.
					GetColIndex(i + 1))];
				
				pColNew->bVisible = pCol->bVisible;
				pColNew->clrBack = pCol->clrBack;
				pColNew->clrFore = pCol->clrFore;
				pColNew->clrHeadBack = pCol->clrHeadBack;
				pColNew->clrHeadFore = pCol->clrHeadFore;
				pColNew->nAlignment = pCol->nAlignment;
				pColNew->nCaptionAlignment = pCol->nCaptionAlignment;
				pColNew->nCase = pCol->nCase;
				pColNew->nWidth = pCol->nWidth;
				pColNew->strCaption = pCol->strCaption;
				pColNew->strDataField = pCol->strDataField;
				pColNew->strName = pCol->strName;
				
				m_pComboCtrl->m_apColumnsProp.Add(pColNew);
				
				for (j = 1; j < nRows; j ++)
					m_pComboCtrl->m_arContent.Add(m_wndGrid.GetValueRowCol(
					j, i + 1));
			}
			
			m_pComboCtrl->InitGridFromProp();
		}
	}
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridPropColumnPage message handlers

BOOL CKCOMRichGridPropColumnPage::OnInitDialog() 
{
	COlePropertyPage::OnInitDialog();
	
	InitControlsVars();
	HideControls();
	if (!m_pGridCtrl && !m_pDropDownCtrl && !m_pComboCtrl)
		return FALSE;

	int i, j, k;
	ColumnProp * pColCtrl;

	m_wndGrid.LockUpdate();
	m_wndGrid.SetRowCount(3);
	m_wndGrid.SetColCount(0);
	m_wndGrid.m_bCanDirty = FALSE;

	int nColumns;
	int nRecords;
	int nRows;

	if (m_pGridCtrl)
	{
		m_pGridCtrl->GetColumnInfo();
		nColumns = m_pGridCtrl->m_apColumnData.GetSize();
		
		pDataField->ResetContent();
		pDataField->AddString(_T(" "));
		
		for (i = 0; i < nColumns; i ++)
			pDataField->AddString(m_pGridCtrl->m_apColumnData[i]->strName);
		
		if (m_pGridCtrl->m_nDataMode != 0)
		{
			nColumns = m_pGridCtrl->m_apColumnsProp.GetSize();
			nRecords = m_pGridCtrl->m_arContent.GetSize();
			nRows = m_pGridCtrl->GetRows();
			k = 0;
			
			m_wndGrid.SetRowCount(nRows);
			
			for (i = 0; i < nColumns; i ++)
			{
				pColCtrl = m_pGridCtrl->m_apColumnsProp[i];
				InsertCol(pColCtrl);
				for (j = 1; j <= nRows && k < nRecords; j ++, k ++)
				{
					m_wndGrid.Edit(j);
					m_wndGrid.SetValueRange(CGXRange(j, i + 1), m_pGridCtrl->
						m_arContent[k]);
					m_wndGrid.Update();
				}
			}
		}
		else
		{
			nColumns = m_pGridCtrl->m_apColumnsProp.GetSize();
			
			for (i = 0; i < nColumns; i ++)
			{
				pColCtrl = m_pGridCtrl->m_apColumnsProp[i];
				InsertCol(pColCtrl);
			}
		}
	}
	else if (m_pDropDownCtrl)
	{
		m_pDropDownCtrl->GetColumnInfo();
		nColumns = m_pDropDownCtrl->m_apColumnData.GetSize();
		
		pDataField->ResetContent();
		pDataField->AddString(_T(" "));
		
		for (i = 0; i < nColumns; i ++)
			pDataField->AddString(m_pDropDownCtrl->m_apColumnData[i]->strName);
		
		if (m_pDropDownCtrl->m_nDataMode != 0)
		{
			nColumns = m_pDropDownCtrl->m_apColumnsProp.GetSize();
			nRecords = m_pDropDownCtrl->m_arContent.GetSize();
			nRows = m_pDropDownCtrl->GetRows();
			k = 0;
			
			m_wndGrid.SetRowCount(nRows);
			
			for (i = 0; i < nColumns; i ++)
			{
				pColCtrl = m_pDropDownCtrl->m_apColumnsProp[i];
				InsertCol(pColCtrl);
				for (j = 1; j <= nRows && k < nRecords; j ++, k ++)
				{
					m_wndGrid.Edit(j);
					m_wndGrid.SetValueRange(CGXRange(j, i + 1), m_pDropDownCtrl->
						m_arContent[k]);
					m_wndGrid.Update();
				}
			}
		}
		else
		{
			nColumns = m_pDropDownCtrl->m_apColumnsProp.GetSize();
			
			for (i = 0; i < nColumns; i ++)
			{
				pColCtrl = m_pDropDownCtrl->m_apColumnsProp[i];
				InsertCol(pColCtrl);
			}
		}
	}
	else if (m_pComboCtrl)
	{
		m_pComboCtrl->GetColumnInfo();
		nColumns = m_pComboCtrl->m_apColumnData.GetSize();
		
		pDataField->ResetContent();
		pDataField->AddString(_T(" "));
		
		for (i = 0; i < nColumns; i ++)
			pDataField->AddString(m_pComboCtrl->m_apColumnData[i]->strName);
		
		if (m_pComboCtrl->m_nDataMode != 0)
		{
			nColumns = m_pComboCtrl->m_apColumnsProp.GetSize();
			nRecords = m_pComboCtrl->m_arContent.GetSize();
			nRows = m_pComboCtrl->GetRows();
			k = 0;
			
			m_wndGrid.SetRowCount(nRows);
			
			for (i = 0; i < nColumns; i ++)
			{
				pColCtrl = m_pComboCtrl->m_apColumnsProp[i];
				InsertCol(pColCtrl);
				for (j = 1; j <= nRows && k < nRecords; j ++, k ++)
				{
					m_wndGrid.Edit(j);
					m_wndGrid.SetValueRange(CGXRange(j, i + 1), m_pComboCtrl->
						m_arContent[k]);
					m_wndGrid.Update();
				}
			}
		}
		else
		{
			nColumns = m_pComboCtrl->m_apColumnsProp.GetSize();
			
			for (i = 0; i < nColumns; i ++)
			{
				pColCtrl = m_pComboCtrl->m_apColumnsProp[i];
				InsertCol(pColCtrl);
			}
		}
	}

	if (m_nColumns == 0)
		pDeleteColumn->EnableWindow(FALSE);

	m_wndGrid.LockUpdate(FALSE);
	m_wndGrid.Invalidate();
	m_wndGrid.m_bCanDirty = TRUE;

	return FALSE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CKCOMRichGridPropColumnPage::InitControlsVars()
{
	m_wndGrid.SubclassDlgItem(IDC_GRID, this);
	m_wndGrid.Initialize();
	m_wndGrid.GetParam()->EnableUndo(FALSE);
	m_wndGrid.RowHeaderStyle().SetBorders(gxBorderAll, CGXPen().SetStyle(PS_NULL));
	m_wndGrid.GetParam()->EnableTrackRowHeight(GX_TRACK_ALL);
	m_wndGrid.GetParam()->EnableSelection(GX_SELROW | GX_SELCOL);
	m_wndGrid.GetParam()->EnableMoveRows(FALSE);
	m_wndGrid.RegisterControl(GX_IDS_CTRL_EDIT, new CKCOMEditControl(
		&m_wndGrid, GX_IDS_CTRL_EDIT));
	m_wndGrid.RegisterControl(GX_IDS_CTRL_COMBOBOX, new CKCOMComboBox(
		&m_wndGrid, GX_IDS_CTRL_COMBOBOX, GX_IDS_CTRL_COMBOBOX, GXCOMBO_NOTEXTFIT));
	m_wndGrid.RegisterControl(GX_IDS_CTRL_COMBOLIST, new CKCOMComboBox(
		&m_wndGrid, GX_IDS_CTRL_COMBOBOX, GX_IDS_CTRL_COMBOBOX, GXCOMBO_NOTEXTFIT, TRUE));
	m_wndGrid.RegisterControl(KCOM_IDS_CTRL_DATEEDIT, new CKCOMDateControl(
		&m_wndGrid, KCOM_IDS_CTRL_DATEEDIT));
	m_wndGrid.RegisterControl(GX_IDS_CTRL_MASKEDIT, new CKCOMMaskControl(
		&m_wndGrid, KCOM_IDS_CTRL_DATEEDIT));
	m_wndGrid.RegisterControl(KCOM_IDS_CTRL_EDITBUTTON, new CKCOMEditBtn(
		&m_wndGrid, KCOM_IDS_CTRL_DATEEDIT));
	
	ULONG Ulong;
	LPDISPATCH FAR *m_lpDispatch = GetObjectArray(&Ulong);

	// Get the CCmdTarget object associated to any one of the above
	// array elements
	if (Ulong)
	{
		DISPPARAMS tParam;
		tParam.rgdispidNamedArgs = NULL;
		tParam.cArgs = 0;
		tParam.cNamedArgs = 0;
		tParam.rgvarg = 0;
		VARIANT va;
		VariantInit(&va);
		HRESULT hr = m_lpDispatch[0]->Invoke(255, IID_NULL, LOCALE_USER_DEFAULT,
			DISPATCH_METHOD | DISPATCH_PROPERTYGET, &tParam, &va, NULL, NULL);
		if (SUCCEEDED(hr))
		{
			if (va.iVal == 0)
				m_pGridCtrl = (CKCOMRichGridCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
			else if (va.iVal == 1)
				m_pDropDownCtrl = (CKCOMRichDropDownCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
			else if (va.iVal == 2)
				m_pComboCtrl = (CKCOMRichComboCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
		}
		VariantClear(&va);
	}

	pListProperty->AddString("Alignment");
	if (m_pGridCtrl)
		pListProperty->AddString("AllowSizing");
	pListProperty->AddString("BackColor");
	pListProperty->AddString("Caption");
	pListProperty->AddString("CaptionAlignment");
	pListProperty->AddString("Case");
	if (m_pGridCtrl || m_pComboCtrl)
		pListProperty->AddString("ComboBoxStyle");
	pListProperty->AddString("DataField");
	if (m_pGridCtrl)
		pListProperty->AddString("DataType");
	if (m_pGridCtrl)
		pListProperty->AddString("FieldLen");
	pListProperty->AddString("ForeColor");
	pListProperty->AddString("HeadBackColor");
	pListProperty->AddString("HeadForeColor");
	if (m_pGridCtrl)
		pListProperty->AddString("Locked");
	if (m_pGridCtrl)
		pListProperty->AddString("Mask");
	pListProperty->AddString("Name");
	if (m_pGridCtrl)
		pListProperty->AddString("Nullable");
	if (m_pGridCtrl)
		pListProperty->AddString("PromptChar");
	if (m_pGridCtrl)
		pListProperty->AddString("Style");
	pListProperty->AddString("Visible");
	pListProperty->AddString("Width");
	if (m_pGridCtrl)
		pListProperty->AddString("PromptInclude");
/*
	pAlignLeft->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 60, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pAlignRight->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 90, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pAlignCenter->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 120, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pAllowSizing->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pCaption->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pCapAlignLeft->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 50, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pCapAlignRight->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 80, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pCapAlignCenter->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 110, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pCapAlignAsAlignment->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 140, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pCaseNone->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 70, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pCaseUpper->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pCaseLower->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 130, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pDataField->SetWindowPos(NULL, FRAMELEFT + 90, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pBackColor->SetWindowPos(NULL, FRAMELEFT + 90, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pForeColor->SetWindowPos(NULL, FRAMELEFT + 90, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pHeadBackColor->SetWindowPos(NULL, FRAMELEFT + 90, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pHeadForeColor->SetWindowPos(NULL, FRAMELEFT + 90, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pFieldLen->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pLocked->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pText->SetWindowPos(NULL, FRAMELEFT + 50, FRAMETOP + 50, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pBoolean->SetWindowPos(NULL, FRAMELEFT + 50, FRAMETOP + 80, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pByte->SetWindowPos(NULL, FRAMELEFT + 50, FRAMETOP + 110, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pInteger->SetWindowPos(NULL, FRAMELEFT + 50, FRAMETOP + 140, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pLong->SetWindowPos(NULL, FRAMELEFT + 150, FRAMETOP + 50, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pSingle->SetWindowPos(NULL, FRAMELEFT + 150, FRAMETOP + 80, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pCurrency->SetWindowPos(NULL, FRAMELEFT + 150, FRAMETOP + 110, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pDate->SetWindowPos(NULL, FRAMELEFT + 150, FRAMETOP + 140, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pMask->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pName->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pNullable->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pPromptChar->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pStyEdit->SetWindowPos(NULL, FRAMELEFT + 50, FRAMETOP + 35, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pStyEditButton->SetWindowPos(NULL, FRAMELEFT + 50, FRAMETOP + 65, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pStyCheckBox->SetWindowPos(NULL, FRAMELEFT + 50, FRAMETOP + 95, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pStyComboBox->SetWindowPos(NULL, FRAMELEFT + 50, FRAMETOP + 125, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pSetupComboBox->SetWindowPos(NULL, FRAMELEFT + 150, FRAMETOP + 125, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
	pStyButton->SetWindowPos(NULL, FRAMELEFT + 50, FRAMETOP + 155, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pVisible->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);


	pWidth->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);

	pPromptInclude->SetWindowPos(NULL, FRAMELEFT + 100, FRAMETOP + 100, 0, 0, SWP_NOZORDER | SWP_NOSIZE |SWP_NOREDRAW);
*/
	if (!m_pGridCtrl && !m_pDropDownCtrl && !m_pComboCtrl)
	{
		GetDlgItem(IDC_BUTTON_ADDCOLUMN)->EnableWindow(FALSE);
		pDeleteColumn->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_RESET)->EnableWindow(FALSE);
		pFields->EnableWindow(FALSE);
		pComboColumn->EnableWindow(FALSE);
		pListProperty->EnableWindow(FALSE);
		m_wndGrid.EnableWindow(FALSE);
		return;
	}
}

void CKCOMRichGridPropColumnPage::HideControls()
{
	pAlignLeft->ShowWindow(SW_HIDE);
	pAlignRight->ShowWindow(SW_HIDE);
	pAlignCenter->ShowWindow(SW_HIDE);

	pAllowSizing->ShowWindow(SW_HIDE);

	pCaption->ShowWindow(SW_HIDE);

	pCapAlignLeft->ShowWindow(SW_HIDE);
	pCapAlignRight->ShowWindow(SW_HIDE);
	pCapAlignCenter->ShowWindow(SW_HIDE);
	pCapAlignAsAlignment->ShowWindow(SW_HIDE);

	pCaseNone->ShowWindow(SW_HIDE);
	pCaseUpper->ShowWindow(SW_HIDE);
	pCaseLower->ShowWindow(SW_HIDE);

	pComboBoxStyle->ShowWindow(SW_HIDE);

	pBackColor->ShowWindow(SW_HIDE);
	pForeColor->ShowWindow(SW_HIDE);
	pHeadBackColor->ShowWindow(SW_HIDE);
	pHeadForeColor->ShowWindow(SW_HIDE);

	pDataField->ShowWindow(SW_HIDE);
	
	pFieldLen->ShowWindow(SW_HIDE);

	pLocked->ShowWindow(SW_HIDE);

	pText->ShowWindow(SW_HIDE);
	pBoolean->ShowWindow(SW_HIDE);
	pByte->ShowWindow(SW_HIDE);
	pInteger->ShowWindow(SW_HIDE);
	pLong->ShowWindow(SW_HIDE);
	pSingle->ShowWindow(SW_HIDE);
	pCurrency->ShowWindow(SW_HIDE);
	pDate->ShowWindow(SW_HIDE);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);

	pMask->ShowWindow(SW_HIDE);

	pName->ShowWindow(SW_HIDE);

	pNullable->ShowWindow(SW_HIDE);

	pPromptChar->ShowWindow(SW_HIDE);

	pStyEdit->ShowWindow(SW_HIDE);
	pStyEditButton->ShowWindow(SW_HIDE);
	pStyCheckBox->ShowWindow(SW_HIDE);
	pStyComboBox->ShowWindow(SW_HIDE);
	pStyButton->ShowWindow(SW_HIDE);
	pSetupComboBox->ShowWindow(SW_HIDE);

	pVisible->ShowWindow(SW_HIDE);

	pWidth->ShowWindow(SW_HIDE);

	pPromptInclude->ShowWindow(SW_HIDE);

	pStaticFrame->SetWindowText("");
}

void CKCOMRichGridPropColumnPage::OnSelchangeListProperty() 
{
	int nProperty = CalcPropertyOrdinal(pListProperty->GetCurSel());
	
	HideControls();

	if (m_pGridCtrl)
		pFields->EnableWindow(m_pGridCtrl->m_nDataMode == 0);
	else if (m_pDropDownCtrl)
		pFields->EnableWindow(m_pDropDownCtrl->m_nDataMode == 0);
	else if (m_pComboCtrl)
		pFields->EnableWindow(m_pComboCtrl->m_nDataMode == 0);

	m_nCurrentCol = pComboColumn->GetCurSel();
	if (m_nCurrentCol >= m_nColumns || m_nCurrentCol < 0)
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];
	CString strText;

	switch (nProperty)
	{
	case 0:
		pAlignLeft->ShowWindow(SW_SHOW);
		pAlignRight->ShowWindow(SW_SHOW);
		pAlignCenter->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Alignment");

		CheckRadio(pAlignLeft, pCol->nAlignment);
		break;

	case 1:
		pAllowSizing->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("AllowSizing");

		pAllowSizing->SetCheck(pCol->bAllowSizing ? 1 : 0);
		break;

	case 2:
		pBackColor->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("BackColor");
		break;

	case 3:
		pCaption->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Caption");

		pCaption->SetWindowText(pCol->strCaption);
		break;

	case 4:
		pCapAlignLeft->ShowWindow(SW_SHOW);
		pCapAlignRight->ShowWindow(SW_SHOW);
		pCapAlignCenter->ShowWindow(SW_SHOW);
		pCapAlignAsAlignment->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("CaptionAlignment");

		CheckRadio(pCapAlignLeft, pCol->nCaptionAlignment);
		break;
	
	case 5:
		pCaseNone->ShowWindow(SW_SHOW);
		pCaseUpper->ShowWindow(SW_SHOW);
		pCaseLower->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Case");

		CheckRadio(pCaseNone, pCol->nCase);
		break;

	case 6:
		pComboBoxStyle->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("ComboBoxStyle");

		pComboBoxStyle->SetCurSel(pCol->nComboBoxStyle);
		break;

	case 7:
		pDataField->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("DataField");
		
		if (m_pGridCtrl)
			pDataField->EnableWindow(m_pGridCtrl->m_nDataMode == 0);
		else if (m_pDropDownCtrl)
			pDataField->EnableWindow(m_pDropDownCtrl->m_nDataMode == 0);
		else if (m_pComboCtrl)
			pDataField->EnableWindow(m_pComboCtrl->m_nDataMode == 0);

		pDataField->SelectString(-1, pCol->strDataField.IsEmpty() ? 
			_T(" ") : pCol->strDataField);
		break;

	case 8:
		pText->ShowWindow(SW_SHOW);
		pBoolean->ShowWindow(SW_SHOW);
		pByte->ShowWindow(SW_SHOW);
		pInteger->ShowWindow(SW_SHOW);
		pLong->ShowWindow(SW_SHOW);
		pSingle->ShowWindow(SW_SHOW);
		pCurrency->ShowWindow(SW_SHOW);
		pDate->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("DataType");

		CheckRadio(pText, pCol->nDataType);
		pDateMask->ShowWindow(pCol->nDataType == 7 ? SW_SHOW : SW_HIDE);
		pStaticDateMask->ShowWindow(pCol->nDataType == 7 ? SW_SHOW : SW_HIDE);
		break;

	case 9:
		pFieldLen->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("FieldLen");

		strText.Format("%d", pCol->nFieldLen);
		pFieldLen->SetWindowText(strText);

		if (m_pGridCtrl)
			pFieldLen->EnableWindow(m_pGridCtrl->m_nDataMode != 0 || pCol->strDataField.GetLength() == 0);
		break;

	case 10:
		pForeColor->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("ForeColor");
		break;

	case 11:
		pHeadBackColor->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("HeadBackColor");
		break;

	case 12:
		pHeadForeColor->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("HeadForeColor");
		break;

	case 13:
		pLocked->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Locked");

		pLocked->SetCheck(pCol->bLocked ? 1 : 0);
		if (m_pGridCtrl)
			pLocked->EnableWindow(m_pGridCtrl->m_nDataMode != 0 || !pCol->bForceLock);
		break;

	case 14:
		pMask->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Mask");

		pMask->SetWindowText(pCol->strMask);
		break;

	case 15:
		pName->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Name");

		pName->SetWindowText(pCol->strName);
		break;

	case 16:
		pNullable->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Nullable");

		pNullable->SetCheck(pCol->bNullable ? 1 : 0);
		if (m_pGridCtrl)
			pNullable->EnableWindow(m_pGridCtrl->m_nDataMode != 0 || !pCol->bForceNoNullable);
		break;

	case 17:
		pPromptChar->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("PromptChar");

		pPromptChar->SetWindowText(pCol->strPromptChar);
		break;

	case 18:
		pStyEdit->ShowWindow(SW_SHOW);
		pStyEditButton->ShowWindow(SW_SHOW);
		pStyComboBox->ShowWindow(SW_SHOW);
		pStyCheckBox->ShowWindow(SW_SHOW);
		pStyButton->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Style");

		CheckRadio(pStyEdit, pCol->nStyle);
		pSetupComboBox->ShowWindow(pStyComboBox->GetCheck() == 1 ?
			SW_SHOW : SW_HIDE);
		break;

	case 19:
		pVisible->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Visible");

		pVisible->SetCheck(pCol->bVisible ? 1 : 0);
		break;

	case 20:
		pWidth->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Width");

		strText.Format("%d", pCol->nWidth);
		pWidth->SetWindowText(strText);
		break;

	case 21:
		pPromptInclude->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("PromptInclude");

		pPromptInclude->SetCheck(pCol->bPromptInclude ? 1 : 0);
		break;
	}
}

void CKCOMRichGridPropColumnPage::OnSelendokComboColumn() 
{
	OnSelchangeListProperty();

	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];
	ROWCOL nCol = m_wndGrid.GetColFromIndex(pCol->nColIndex);
	ROWCOL nRow, nColCurr;

	m_wndGrid.GetCurrentCell(nRow, nColCurr);
	m_wndGrid.SetCurrentCell(nRow, nCol);
}

void CKCOMRichGridPropColumnPage::CheckRadio(CButton *pButton, int nValue)
{
	CWnd * pCtrl = pButton;

	ASSERT(::GetWindowLong(pCtrl->m_hWnd, GWL_STYLE) & WS_GROUP);
	ASSERT(::SendMessage(pCtrl->m_hWnd, WM_GETDLGCODE, 0, 0L) & DLGC_RADIOBUTTON);

	// walk all children in group
	int iButton = 0;
	do
	{
		if (::SendMessage(pCtrl->m_hWnd, WM_GETDLGCODE, 0, 0L) & DLGC_RADIOBUTTON)
		{
			// control in group is a radio button
			// select button
			::SendMessage(pCtrl->m_hWnd, BM_SETCHECK, (iButton == nValue), 0L);
			iButton++;
		}
		else
		{
			TRACE0("Warning: skipping non-radio button in group.\n");
		}
		pCtrl = pCtrl->GetWindow(GW_HWNDNEXT);

	} while (pCtrl != NULL && !(GetWindowLong(pCtrl->m_hWnd, GWL_STYLE) 
		& WS_GROUP));
}

void CKCOMRichGridPropColumnPage::OnButtonBackcolor() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	COLORREF clrBack;
	
	if (m_pGridCtrl)
		OleTranslateColor(pCol->clrBack == DEFAULTCOLOR ? m_pGridCtrl->GetBackColor() : pCol->clrBack, NULL, &clrBack);
	else if (m_pDropDownCtrl)
		OleTranslateColor(pCol->clrBack == DEFAULTCOLOR ? m_pDropDownCtrl->GetBackColor() : pCol->clrBack, NULL, &clrBack);
	else if (m_pComboCtrl)
		OleTranslateColor(pCol->clrBack == DEFAULTCOLOR ? m_pComboCtrl->GetBackColor() : pCol->clrBack, NULL, &clrBack);

	CColorDialog clrDlg(clrBack, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
	{
		if (m_pGridCtrl)
		{
			pCol->clrBack = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetBackColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrBack);
		}
		else if (m_pDropDownCtrl)
		{
			pCol->clrBack = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetBackColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrBack);
		}
		else if (m_pComboCtrl)
		{
			pCol->clrBack = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetBackColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrBack);
		}
	}
}

void CKCOMRichGridPropColumnPage::OnButtonForecolor() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	COLORREF clrFore;
	
	if (m_pGridCtrl)
		OleTranslateColor(pCol->clrFore == DEFAULTCOLOR ? m_pGridCtrl->GetForeColor() : pCol->clrFore, NULL, &clrFore);
	else if (m_pDropDownCtrl)
		OleTranslateColor(pCol->clrFore == DEFAULTCOLOR ? m_pDropDownCtrl->GetForeColor() : pCol->clrFore, NULL, &clrFore);
	else if (m_pComboCtrl)
		OleTranslateColor(pCol->clrFore == DEFAULTCOLOR ? m_pComboCtrl->GetForeColor() : pCol->clrFore, NULL, &clrFore);

	CColorDialog clrDlg(clrFore, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
	{
		if (m_pGridCtrl)
		{
			pCol->clrFore = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetForeColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrFore);
		}
		else if (m_pDropDownCtrl)
		{
			pCol->clrFore = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetForeColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrFore);
		}
		else if (m_pComboCtrl)
		{
			pCol->clrFore = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetForeColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrFore);
		}
	}
}

void CKCOMRichGridPropColumnPage::OnButtonHeadbackcolor() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	COLORREF clrHeadBack;
	
	if (m_pGridCtrl)
		OleTranslateColor(pCol->clrHeadBack == DEFAULTCOLOR ? m_pGridCtrl->m_clrHeadBack : pCol->clrHeadBack, NULL, &clrHeadBack);
	else if (m_pDropDownCtrl)
		OleTranslateColor(pCol->clrHeadBack == DEFAULTCOLOR ? m_pDropDownCtrl->m_clrHeadBack : pCol->clrHeadBack, NULL, &clrHeadBack);
	else if (m_pComboCtrl)
		OleTranslateColor(pCol->clrHeadBack == DEFAULTCOLOR ? m_pComboCtrl->m_clrHeadBack : pCol->clrHeadBack, NULL, &clrHeadBack);

	CColorDialog clrDlg(clrHeadBack, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
	{
		if (m_pGridCtrl)
		{
			pCol->clrHeadBack = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetHeadBackColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrHeadBack);
		}
		else if (m_pDropDownCtrl)
		{
			pCol->clrHeadBack = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetHeadBackColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrHeadBack);
		}
		else if (m_pComboCtrl)
		{
			pCol->clrHeadBack = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetHeadBackColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrHeadBack);
		}
	}
}

void CKCOMRichGridPropColumnPage::OnButtonHeadforecolor() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	COLORREF clrHeadFore;
	
	if (m_pGridCtrl)
		OleTranslateColor(pCol->clrHeadFore == DEFAULTCOLOR ? m_pGridCtrl->m_clrHeadFore : pCol->clrHeadFore, NULL, &clrHeadFore);
	else if (m_pDropDownCtrl)
		OleTranslateColor(pCol->clrHeadFore == DEFAULTCOLOR ? m_pDropDownCtrl->m_clrHeadFore : pCol->clrHeadFore, NULL, &clrHeadFore);
	else if (m_pComboCtrl)
		OleTranslateColor(pCol->clrHeadFore == DEFAULTCOLOR ? m_pComboCtrl->m_clrHeadFore : pCol->clrHeadFore, NULL, &clrHeadFore);

	CColorDialog clrDlg(clrHeadFore, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
	{
		if (m_pGridCtrl)
		{
			pCol->clrHeadFore = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetHeadForeColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrHeadFore);
		}
		else if (m_pDropDownCtrl)
		{
			pCol->clrHeadFore = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetHeadForeColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrHeadFore);
		}
		else if (m_pComboCtrl)
		{
			pCol->clrHeadFore = (OLE_COLOR) clrDlg.GetColor();
			m_wndGrid.SetHeadForeColor(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->clrHeadFore);
		}
	}
}

void CKCOMRichGridPropColumnPage::OnCheckAllowsizing() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->bAllowSizing = pAllowSizing->GetCheck();
}

void CKCOMRichGridPropColumnPage::OnCheckLocked() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->bLocked = pLocked->GetCheck();
}

void CKCOMRichGridPropColumnPage::OnCheckNullable() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->bNullable = pNullable->GetCheck();
}

void CKCOMRichGridPropColumnPage::OnCheckVisible() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->bVisible = pVisible->GetCheck();
	m_wndGrid.SetVisible(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->bVisible);
}

void CKCOMRichGridPropColumnPage::OnKillfocusEditName() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	CString strName;
	pName->GetWindowText(strName);
	if (!strName.CompareNoCase(pCol->strName))
		return;

	for (int i = 0; i < m_apColumnProp.GetSize(); i ++)
		if (!m_apColumnProp[i]->strName.CompareNoCase(strName))
			return;

	pCol->strName = strName;

	pComboColumn->InsertString(m_nCurrentCol, strName);
	pComboColumn->DeleteString(m_nCurrentCol + 1);
}

void CKCOMRichGridPropColumnPage::OnKillfocusEditCaption() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCaption->GetWindowText(pCol->strCaption);
	m_wndGrid.SetCaption(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->strCaption);
}

void CKCOMRichGridPropColumnPage::OnKillfocusEditFieldlen() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	CString strFieldLen;
	pFieldLen->GetWindowText(strFieldLen);

	int nFieldLen = atoi(strFieldLen);
	if (nFieldLen <= 0 || nFieldLen > 255)
		return;

	pCol->nFieldLen = nFieldLen;
}

void CKCOMRichGridPropColumnPage::OnKillfocusEditMask() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	if (!((CEdit * )pMask)->GetModify())
		return;

	pMask->GetWindowText(pCol->strMask);
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_INPUTMASK, pCol->strMask));
	UpdateStyle(pCol, m_nCurrentCol + 1);
}

void CKCOMRichGridPropColumnPage::OnKillfocusEditPromptchar() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	CString strPromptChar;
	pPromptChar->GetWindowText(strPromptChar);
	if (strPromptChar.GetLength() != 1)
		return;

	pCol->strPromptChar = strPromptChar;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_INPUTPROMPT, strPromptChar));
}

void CKCOMRichGridPropColumnPage::OnKillfocusEditWidth() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	CString strWidth;
	pWidth->GetWindowText(strWidth);

	int nWidth = atoi(strWidth);
	if (nWidth < 0 || nWidth > 15000)
		return;

	pCol->nWidth = nWidth;
	m_wndGrid.SetWidth(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->nWidth);
}

void CKCOMRichGridPropColumnPage::OnRadioAlignmentCenter() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nAlignment = 2;
	m_wndGrid.SetAlignment(m_wndGrid.GetColFromIndex(pCol->nColIndex), 2);
}

void CKCOMRichGridPropColumnPage::OnRadioAlignmentLeft() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nAlignment = 0;
	m_wndGrid.SetAlignment(m_wndGrid.GetColFromIndex(pCol->nColIndex), 0);
}

void CKCOMRichGridPropColumnPage::OnRadioAlignmentRight() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nAlignment = 1;
	m_wndGrid.SetAlignment(m_wndGrid.GetColFromIndex(pCol->nColIndex), 1);
}

void CKCOMRichGridPropColumnPage::OnRadioCaptionalignmentAsalignment() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nCaptionAlignment = 3;
	m_wndGrid.SetCaptionAlignment(m_wndGrid.GetColFromIndex(pCol->nColIndex), pCol->nAlignment);
}

void CKCOMRichGridPropColumnPage::OnRadioCaptionalignmentCenter() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nCaptionAlignment = 2;
	m_wndGrid.SetCaptionAlignment(m_wndGrid.GetColFromIndex(pCol->nColIndex), 2);
}

void CKCOMRichGridPropColumnPage::OnRadioCaptionalignmentLeft() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nCaptionAlignment = 0;
	m_wndGrid.SetCaptionAlignment(m_wndGrid.GetColFromIndex(pCol->nColIndex), 0);
}

void CKCOMRichGridPropColumnPage::OnRadioCaptionalignmentRight() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nCaptionAlignment = 1;
	m_wndGrid.SetCaptionAlignment(m_wndGrid.GetColFromIndex(pCol->nColIndex), 1);
}

void CKCOMRichGridPropColumnPage::OnRadioCaseLower() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nCase = 2;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_CASE, pCol->nCase));
}

void CKCOMRichGridPropColumnPage::OnRadioCaseNone() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nCase = 0;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_CASE, pCol->nCase));
}

void CKCOMRichGridPropColumnPage::OnRadioCaseUpper() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nCase = 1;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_CASE, pCol->nCase));
}

void CKCOMRichGridPropColumnPage::OnRadioDatatypeBoolean() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nDataType = 1;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_DATATYPE, pCol->nDataType));
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioDatatypeByte() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nDataType = 2;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_DATATYPE, pCol->nDataType));
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioDatatypeCurrency() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nDataType = 6;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_DATATYPE, pCol->nDataType));
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioDatatypeDate() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nDataType = 7;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_DATATYPE, pCol->nDataType));
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pDateMask->ShowWindow(SW_SHOW);
	pStaticDateMask->ShowWindow(SW_SHOW);
}

void CKCOMRichGridPropColumnPage::OnRadioDatatypeInteger() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nDataType = 3;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_DATATYPE, pCol->nDataType));
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioDatatypeLong() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nDataType = 4;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_DATATYPE, pCol->nDataType));
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioDatatypeSingle() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nDataType = 5;
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_DATATYPE, pCol->nDataType));
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioDatatypeText() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nDataType = 0;
	UpdateStyle(pCol, m_nCurrentCol + 1);
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_DATATYPE, pCol->nDataType));
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioStyleButton() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nStyle = 4;
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pSetupComboBox->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioStyleCheckbox() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nStyle = 2;
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pSetupComboBox->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioStyleCombobox() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nStyle = 3;
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pSetupComboBox->ShowWindow(SW_SHOW);
}

void CKCOMRichGridPropColumnPage::OnRadioStyleEdit() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nStyle = 0;
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pSetupComboBox->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnRadioStyleEditbutton() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->nStyle = 1;
	UpdateStyle(pCol, m_nCurrentCol + 1);
	pSetupComboBox->ShowWindow(SW_HIDE);
}

void CKCOMRichGridPropColumnPage::OnButtonSetupcombobox() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	CGridInputDlg dlgInput;
	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];
	
	dlgInput.m_strChoiceList = pCol->strChoiceList;
	if (dlgInput.DoModal() == IDOK)
	{
		pCol->strChoiceList = dlgInput.m_strChoiceList;
		m_wndGrid.SetStyleRange(CGXRange().SetCols(m_wndGrid.GetColFromIndex(
			pCol->nColIndex)), CGXStyle().SetChoiceList(pCol->strChoiceList));
	}
}

void CKCOMRichGridPropColumnPage::OnButtonAddcolumn() 
{
	CAddColumnDlg dlg;

	dlg.SetPagePtr(this);
	if (dlg.DoModal() == IDOK)
	{
		m_wndGrid.LockUpdate();
		InsertCol(NULL);
		m_wndGrid.LockUpdate(FALSE);
		m_wndGrid.Redraw();

		m_apColumnProp[m_nColumns - 1]->strCaption = m_apColumnProp[
			m_nColumns - 1]->strName = dlg.m_strName;
		m_wndGrid.SetCaption(m_nColumns, dlg.m_strName);
		pComboColumn->AddString(dlg.m_strName);

		pComboColumn->SetCurSel(m_nColumns - 1);
	}
}

void CKCOMRichGridPropColumnPage::OnButtonDeletecolumn() 
{
	m_nCurrentCol = pComboColumn->GetCurSel();
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	ROWCOL nCol = m_wndGrid.GetColFromIndex(pCol->nColIndex);
	delete pCol;
	m_apColumnProp.RemoveAt(m_nCurrentCol);
	m_nColumns --;

	m_wndGrid.RemoveCols(nCol, nCol);
	pComboColumn->DeleteString(m_nCurrentCol);
	
	OnSelendokComboColumn();

	if (m_nColumns == 0)
		pDeleteColumn->EnableWindow(FALSE);
}

void CKCOMRichGridPropColumnPage::InsertCol(ColumnProp * pColSource)
{
	ColumnProp * pCol = new ColumnProp;
	if (pColSource)
	{
		pCol->bAllowSizing = pColSource->bAllowSizing;
		pCol->bLocked = pColSource->bLocked;
		pCol->bNullable = pColSource->bNullable;
		pCol->bVisible = pColSource->bVisible;
		pCol->clrBack = pColSource->clrBack;
		pCol->clrFore = pColSource->clrFore;
		pCol->clrHeadBack = pColSource->clrHeadBack;
		pCol->clrHeadFore = pColSource->clrHeadFore;
		pCol->bForceLock = pColSource->bForceLock;
		pCol->bForceNoNullable = pColSource->bForceNoNullable;
		pCol->nAlignment = pColSource->nAlignment;
		pCol->nCaptionAlignment = pColSource->nCaptionAlignment;
		pCol->nCase = pColSource->nCase;
		pCol->nDataType = pColSource->nDataType;
		pCol->nFieldLen = pColSource->nFieldLen;
		pCol->nStyle = pColSource->nStyle;
		pCol->nWidth = pColSource->nWidth;
		pCol->strCaption = pColSource->strCaption;
		pCol->strDataField = pColSource->strDataField;
		pCol->strMask = pColSource->strMask;
		pCol->strName = pColSource->strName;
		pCol->strPromptChar = pColSource->strPromptChar;
		pCol->strChoiceList = pColSource->strChoiceList;
		pCol->bPromptInclude = pColSource->bPromptInclude;
		pCol->nComboBoxStyle = pColSource->nComboBoxStyle;
	}
	else
	{
		pCol->bAllowSizing = TRUE;
		pCol->bLocked = FALSE;
		pCol->bNullable = TRUE;
		pCol->bVisible = TRUE;
		pCol->bForceLock = FALSE;
		pCol->bForceNoNullable = FALSE;
		pCol->nAlignment = 2;
		pCol->nCaptionAlignment = 2;
		pCol->nCase = 0;
		pCol->nDataType = 0;
		pCol->nDataField = -1;
		pCol->nFieldLen = 255;
		pCol->nStyle = 0;
		if (m_pGridCtrl)
			pCol->nWidth = m_pGridCtrl->m_nDefColWidth;
		else if (m_pDropDownCtrl)
			pCol->nWidth = m_pDropDownCtrl->m_nDefColWidth;
		else if (m_pComboCtrl)
			pCol->nWidth = m_pComboCtrl->m_nDefColWidth;
		pCol->strPromptChar = _T("_");
		pCol->bPromptInclude = TRUE;
	}

	m_apColumnProp.Add(pCol);
	m_nColumns ++;
	m_wndGrid.InsertCols(m_nColumns, 1);
	UpdateStyle(pCol, m_nColumns);
	pCol->nColIndex = m_wndGrid.GetColIndex(m_nColumns);

	m_wndGrid.SetCaption(m_nColumns, pCol->strCaption);
	m_wndGrid.SetAlignment(m_nColumns, pCol->nAlignment);
	m_wndGrid.SetBackColor(m_nColumns, pCol->clrBack);
	m_wndGrid.SetCaptionAlignment(m_nColumns, pCol->nCaptionAlignment);
	m_wndGrid.SetForeColor(m_nColumns, pCol->clrFore);
	m_wndGrid.SetHeadBackColor(m_nColumns, pCol->clrHeadBack);
	m_wndGrid.SetHeadForeColor(m_nColumns, pCol->clrHeadFore);
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nColumns), CGXStyle().SetUserAttribute(
		GX_IDS_UA_INPUTMASK, pCol->strMask));
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nColumns), CGXStyle().SetUserAttribute(
		GX_IDS_UA_PROMPTINCLUDE, pCol->bPromptInclude ? _T("T") : _T("F")));
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nColumns), CGXStyle().SetUserAttribute(
		GX_IDS_UA_INPUTPROMPT, pCol->strPromptChar));
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nColumns), CGXStyle().SetUserAttribute(
		GX_IDS_UA_CASE, pCol->nCase));
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nColumns), CGXStyle().SetChoiceList(
		pCol->strChoiceList));

	if (pColSource)
	{
		m_wndGrid.SetVisible(m_nColumns, pCol->bVisible);
		m_wndGrid.SetWidth(m_nColumns, pCol->nWidth);
		pComboColumn->AddString(pCol->strName);
	}
}

void CKCOMRichGridPropColumnPage::OnButtonReset() 
{
	for (int i = 0; i < m_nColumns; i ++)
		delete m_apColumnProp[i];

	m_apColumnProp.RemoveAll();

	pComboColumn->ResetContent();
	m_wndGrid.SetColCount(0);
	pDeleteColumn->EnableWindow(FALSE);
	
	m_nColumns = 0;
	OnSelendokComboColumn();
}

void CKCOMRichGridPropColumnPage::OnCheckPromptinclude() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pCol->bPromptInclude = pPromptInclude->GetCheck();
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_PROMPTINCLUDE, pCol->bPromptInclude ? _T("T") : _T("F")));
}

void CKCOMRichGridPropColumnPage::UpdateStyle(ColumnProp *pCol, ROWCOL nCol)
{
	WORD nCtrl = 0;
	
	if (pCol->nDataType == 7)
		nCtrl = KCOM_IDS_CTRL_DATEEDIT;
	else if (pCol->strMask.GetLength())
		nCtrl = GX_IDS_CTRL_MASKEDIT;
	else if (pCol->nStyle == 3)
	{
		if (pCol->nComboBoxStyle == 0)
			nCtrl = GX_IDS_CTRL_COMBOBOX;
		else
			nCtrl = GX_IDS_CTRL_COMBOLIST;
	}

	if (nCtrl > 0)
	{
		m_wndGrid.SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(nCtrl));
		return;
	}
	
	m_wndGrid.SetStyle(nCol, pCol->nStyle);
}

void CKCOMRichGridPropColumnPage::OnSelendokComboDatafield() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];

	pDataField->GetWindowText(pCol->strDataField);
	if (pCol->strDataField == _T(" "))
		pCol->strDataField.Empty();
	else
		pCol->nFieldLen = 255;
}

void CKCOMRichGridPropColumnPage::OnButtonFields() 
{
	CFieldsDlg dlg;
	int i;

	dlg.m_arField.RemoveAll();
	
	if (m_pGridCtrl)
	{
		for (i = 0; i < m_pGridCtrl->m_apColumnData.GetSize(); i ++)
			dlg.m_arField.Add(m_pGridCtrl->m_apColumnData[i]->strName);
	}
	else if (m_pDropDownCtrl)
	{
		for (i = 0; i < m_pDropDownCtrl->m_apColumnData.GetSize(); i ++)
			dlg.m_arField.Add(m_pDropDownCtrl->m_apColumnData[i]->strName);
	}
	else if (m_pComboCtrl)
	{
		for (i = 0; i < m_pComboCtrl->m_apColumnData.GetSize(); i ++)
			dlg.m_arField.Add(m_pComboCtrl->m_apColumnData[i]->strName);
	}

	if (dlg.DoModal() == IDOK)
	{
		ColumnProp col;

		m_wndGrid.LockUpdate();

		for (i = 0; i < dlg.m_arField.GetSize(); i ++)
		{
			col.strDataField = dlg.m_arField[i];

			if (m_pGridCtrl)
			{
				col.nDataField = m_pGridCtrl->GetFieldFromStr(col.strDataField);
				col.bForceNoNullable = !(m_pGridCtrl->m_pColumnInfo[col.nDataField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE);
				col.bForceLock = !(m_pGridCtrl->m_pColumnInfo[col.nDataField].dwFlags & DBCOLUMNFLAGS_WRITE);
				col.bLocked = col.bForceLock;
				col.bNullable = !col.bForceNoNullable;
				col.bAllowSizing = TRUE;
				col.nDataType = m_pGridCtrl->GetDataTypeFromStr(col.strDataField);
				col.nFieldLen = 255;
				col.nStyle = 0;
				col.strPromptChar = _T("_");
				col.bPromptInclude = TRUE;
				col.nWidth = m_pGridCtrl->m_nDefColWidth;
			}
			else if (m_pDropDownCtrl)
			{
				col.nDataField = m_pDropDownCtrl->GetFieldFromStr(col.strDataField);
				col.nDataField = m_pDropDownCtrl->GetFieldFromStr(col.strDataField);
				col.nWidth = m_pDropDownCtrl->m_nDefColWidth;
			}
			else if (m_pComboCtrl)
			{
				col.nDataField = m_pComboCtrl->GetFieldFromStr(col.strDataField);
				col.nDataField = m_pComboCtrl->GetFieldFromStr(col.strDataField);
				col.nWidth = m_pComboCtrl->m_nDefColWidth;
			}

			col.strName = col.strCaption = FindUniqueName(col.strDataField);

			col.bVisible = TRUE;
			col.nAlignment = 2;
			col.nCaptionAlignment = 2;
			col.nCase = 0;
			
			InsertCol(&col);
		}

		m_wndGrid.LockUpdate(FALSE);
		m_wndGrid.Redraw();
	}
}

CString CKCOMRichGridPropColumnPage::FindUniqueName(CString& strRecm)
{
	static int nOrdinal = 0;

	for (int i = 0; i < m_apColumnProp.GetSize(); i ++)
		if (!strRecm.CompareNoCase(m_apColumnProp[i]->strName))
		{
			strRecm.Format("%s%3d", strRecm, ++ nOrdinal);
			return FindUniqueName(strRecm);
		}

		return strRecm;
}

void CKCOMRichGridPropColumnPage::OnSelendokComboDatemask() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];
	pDateMask->GetWindowText(pCol->strMask);
	m_wndGrid.SetStyleRange(CGXRange().SetCols(m_nCurrentCol + 1), CGXStyle().SetUserAttribute(
		GX_IDS_UA_INPUTMASK, pCol->strMask));
}

int CKCOMRichGridPropColumnPage::CalcPropertyOrdinal(int nProperty)
{
	int nOrdinal = nProperty;

	if (m_pDropDownCtrl)
	{
		if (nOrdinal == 1)
			nOrdinal = 2;
		else if (nOrdinal == 2)
			nOrdinal = 3;
		else if (nOrdinal == 3)
			nOrdinal = 4;
		else if (nOrdinal == 4)
			nOrdinal = 5;
		else if (nOrdinal == 5)
			nOrdinal = 7;
		else if (nOrdinal == 6)
			nOrdinal = 10;
		else if (nOrdinal == 7)
			nOrdinal = 11;
		else if (nOrdinal == 8)
			nOrdinal = 12;
		else if (nOrdinal == 9)
			nOrdinal = 15;
		else if (nOrdinal == 10)
			nOrdinal = 19;
		else if (nOrdinal == 11)
			nOrdinal = 20;
	}
	else if (m_pComboCtrl)
	{
		if (nOrdinal == 1)
			nOrdinal = 2;
		else if (nOrdinal == 2)
			nOrdinal = 3;
		else if (nOrdinal == 3)
			nOrdinal = 4;
		else if (nOrdinal == 4)
			nOrdinal = 5;
		else if (nOrdinal == 5)
			nOrdinal = 6;
		else if (nOrdinal == 6)
			nOrdinal = 7;
		else if (nOrdinal == 7)
			nOrdinal = 10;
		else if (nOrdinal == 8)
			nOrdinal = 11;
		else if (nOrdinal == 9)
			nOrdinal = 12;
		else if (nOrdinal == 10)
			nOrdinal = 15;
		else if (nOrdinal == 11)
			nOrdinal = 19;
		else if (nOrdinal == 12)
			nOrdinal = 20;
	}

	return nOrdinal;
}

void CKCOMRichGridPropColumnPage::OnSelendokComboComboboxstyle() 
{
	if (m_nCurrentCol < 0 || m_nCurrentCol >= m_apColumnProp.GetSize())
		return;

	ColumnProp * pCol = m_apColumnProp[m_nCurrentCol];
	pCol->nComboBoxStyle = pComboBoxStyle->GetCurSel();
	UpdateStyle(pCol, m_nCurrentCol + 1);
}
