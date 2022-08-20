// KCOMRichComboCtl.cpp : Implementation of the CKCOMRichComboCtrl ActiveX Control class.

#include "stdafx.h"
#include "KCOMData.h"
#include "KCOMRichComboCtl.h"
#include "KCOMRichComboPpg.h"
#include <oledberr.h>
#include <msstkppg.h>
#include "DropDownColumn.h"
#include "DropDownColumns.h"
#include "DropDownRealWnd.h"
#include "KCOMRichGridColumnPpg.h"
#include "ComboEdit.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CKCOMRichComboCtrl, COleControl)

BEGIN_INTERFACE_MAP(CKCOMRichComboCtrl, COleControl)
	// Map the interfaces to their implementations
	INTERFACE_PART(CKCOMRichComboCtrl, IID_IRowsetNotify, RowsetNotify)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) CKCOMRichComboCtrl::XRowsetNotify::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, RowsetNotify)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMRichComboCtrl::XRowsetNotify::Release(void)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, RowsetNotify)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMRichComboCtrl::XRowsetNotify::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, RowsetNotify)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMRichComboCtrl::XRowsetNotify::OnFieldChange(IRowset * prowset,
	HROW hRow, ULONG cColumns, ULONG rgColumns[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_COLUMN_SET:
		case DBREASON_COLUMN_RECALCULATED:
		{
			ROWCOL nRow;
			
			// get row ordinal from hrow
			nRow = pThis->GetRowFromHRow(hRow);
			if (nRow == GX_INVALID || nRow > (ROWCOL)pThis->m_pDropDownRealWnd->OnGetRecordCount())
			{
				// this is a new record, fetch it :)
				pThis->m_bEndReached = FALSE;
				pThis->FetchNewRows();
				return S_OK;
			}

			for (ULONG i = 0; i < cColumns; i ++)
				pThis->m_pDropDownRealWnd->RedrawRowCol(nRow, rgColumns[i]);
		}

		return S_OK;
		break;

		default:
			return DB_S_UNWANTEDREASON;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP CKCOMRichComboCtrl::XRowsetNotify::OnRowChange(IRowset * prowset,
	ULONG cRows, const HROW rghRows[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, RowsetNotify)

	
	if (DBEVENTPHASE_ABOUTTODO == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{	
		switch (eReason)
		{
		case DBREASON_ROW_DELETE:
		case DBREASON_ROW_UNDOINSERT:
		{
			ROWCOL nRow;

			// we should remember the record ordinal for in some cases
			// we can not get the ordinal of the deleted records using
			// oledb interface
			// This method may be buggy but I do not know why the datagrid
			// of MS can work fine. Maybe he cached this information like me?
			for (ULONG i = 0; i < cRows; i ++)
			{
				nRow = pThis->GetRowFromHRow(rghRows[i]);
				if (nRow != GX_INVALID)
					pThis->DeleteRowFromSrc(nRow);
			}
		}

		default:
			return S_OK;
		}
	}

	if (DBEVENTPHASE_DIDEVENT == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_ROW_ASYNCHINSERT:
		case DBREASON_ROW_INSERT:
		{
		pThis->m_bEndReached = FALSE;
			pThis->FetchNewRows();
		}

		return S_OK;
		break;

		// in some cases I can not get the currect record ordinal, so I
		// have to cache the information for undelete
		case DBREASON_ROW_UNDODELETE:
		{
			for (ULONG i = 0; i < cRows; i ++)
				pThis->UndeleteRecordFromHRow(rghRows[i]);
		}

		return S_OK;
		break;

		case DBREASON_ROW_UNDOCHANGE:
		case DBREASON_ROW_UPDATE:
		{
			ROWCOL nRow;
			BOOL bUpdated = FALSE;
			for (ULONG i = 0; i < cRows; i ++)
			{
				nRow = pThis->GetRowFromHRow(rghRows[i]);
				if (nRow != GX_INVALID)
				{
					pThis->m_pDropDownRealWnd->RedrawRowCol(CGXRange().SetRows(nRow));
					bUpdated = TRUE;
				}
			}
			if (!bUpdated)
			{
				pThis->m_bEndReached = FALSE;
				pThis->FetchNewRows();
			}

		}

		return S_OK;
		break;

		default:
			return S_OK;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP CKCOMRichComboCtrl::XRowsetNotify::OnRowsetChange(IRowset * prowset,
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_ROWSET_RELEASE:
		case DBREASON_ROWSET_CHANGED:
			pThis->m_bDataSrcChanged = TRUE;

		return S_OK;
		break;
			
		default:
			return DB_S_UNWANTEDREASON;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP_(ULONG) CKCOMRichComboCtrl::XDataSourceListener::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, DataSourceListener)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMRichComboCtrl::XDataSourceListener::Release(void)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, DataSourceListener)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMRichComboCtrl::XDataSourceListener::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, DataSourceListener)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMRichComboCtrl::XDataSourceListener::dataMemberChanged(BSTR bstrDM)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, RowsetNotify)

	USES_CONVERSION;
	if ((pThis->m_strDataMember.IsEmpty() && (bstrDM == NULL || bstrDM[0] == '\0')) || pThis->m_strDataMember == OLE2T(bstrDM))
	{
		pThis->ClearInterfaces();
		pThis->m_bDataSrcChanged = TRUE;
		pThis->m_bEndReached = FALSE;
	}
	
	return S_OK;
}

STDMETHODIMP CKCOMRichComboCtrl::XDataSourceListener::dataMemberAdded(BSTR bstrDM)
{
	return S_OK;
}

STDMETHODIMP CKCOMRichComboCtrl::XDataSourceListener::dataMemberRemoved(BSTR bstrDM)
{
	METHOD_PROLOGUE(CKCOMRichComboCtrl, RowsetNotify)

	USES_CONVERSION;
	if ((pThis->m_strDataMember.IsEmpty() && (bstrDM == NULL || bstrDM[0] == '\0')) || pThis->m_strDataMember == OLE2T(bstrDM))
	{
		pThis->ClearInterfaces();
		pThis->m_bDataSrcChanged = TRUE;
	}

	return S_OK;
}

/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CKCOMRichComboCtrl, COleControl)
	//{{AFX_MSG_MAP(CKCOMRichComboCtrl)
	ON_WM_CREATE()
	ON_WM_KILLFOCUS()
	ON_WM_SETFOCUS()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_CTLCOLOR()
	ON_WM_MOUSEACTIVATE()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CKCOMRichComboCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CKCOMRichComboCtrl)
	DISP_PROPERTY_NOTIFY(CKCOMRichComboCtrl, "HeadBackColor", m_clrHeadBack, OnHeadBackColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CKCOMRichComboCtrl, "HeadForeColor", m_clrHeadFore, OnHeadForeColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CKCOMRichComboCtrl, "ListAutoPosition", m_bListAutoPosition, OnListAutoPositionChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CKCOMRichComboCtrl, "ListWidthAutoSize", m_bListWidthAutoSize, OnListWidthAutoSizeChanged, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Format", GetFormat, SetFormat, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "MaxLength", GetMaxLength, SetMaxLength, VT_I2)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "ReadOnly", GetReadOnly, SetReadOnly, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "DataSource", GetDataSource, SetDataSource, VT_UNKNOWN)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "DataMember", GetDataMember, SetDataMember, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "DefColWidth", GetDefColWidth, SetDefColWidth, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "RowHeight", GetRowHeight, SetRowHeight, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "DividerType", GetDividerType, SetDividerType, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "DividerStyle", GetDividerStyle, SetDividerStyle, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "DataMode", GetDataMode, SetDataMode, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "LeftCol", GetLeftCol, SetLeftCol, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "FirstRow", GetFirstRow, SetFirstRow, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Row", GetRow, SetRow, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Col", GetCol, SetCol, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Rows", GetRows, SetRows, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Cols", GetCols, SetCols, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "DroppedDown", GetDroppedDown, SetDroppedDown, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "DataField", GetDataField, SetDataField, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Columns", GetColumns, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "VisibleRows", GetVisibleRows, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "VisibleCols", GetVisibleCols, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "ColumnHeaders", GetColumnHeaders, SetColumnHeaders, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Font", GetFont, SetFont, VT_FONT)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "HeadFont", GetHeadFont, SetHeadFont, VT_FONT)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "FieldSeparator", GetFieldSeparator, SetFieldSeparator, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "MaxDropDownItems", GetMaxDropDownItems, SetMaxDropDownItems, VT_I2)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "MinDropDownItems", GetMinDropDownItems, SetMinDropDownItems, VT_I2)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "ListWidth", GetListWidth, SetListWidth, VT_I2)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "HeaderHeight", GetHeaderHeight, SetHeaderHeight, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Text", GetText, SetText, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Bookmark", GetBookmark, SetBookmark, VT_VARIANT)
	DISP_PROPERTY_EX(CKCOMRichComboCtrl, "Style", GetStyle, SetStyle, VT_I4)
	DISP_FUNCTION(CKCOMRichComboCtrl, "AddItem", AddItem, VT_EMPTY, VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CKCOMRichComboCtrl, "RemoveItem", RemoveItem, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CKCOMRichComboCtrl, "Scroll", Scroll, VT_EMPTY, VTS_I4 VTS_I4)
	DISP_FUNCTION(CKCOMRichComboCtrl, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichComboCtrl, "MoveFirst", MoveFirst, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichComboCtrl, "MovePrevious", MovePrevious, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichComboCtrl, "MoveNext", MoveNext, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichComboCtrl, "MoveLast", MoveLast, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichComboCtrl, "MoveRecords", MoveRecords, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CKCOMRichComboCtrl, "ReBind", ReBind, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichComboCtrl, "IsItemInList", IsItemInList, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CKCOMRichComboCtrl, "RowBookmark", RowBookmark, VT_VARIANT, VTS_I4)
	DISP_DEFVALUE(CKCOMRichComboCtrl, "Text")
	DISP_STOCKPROP_FORECOLOR()
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CKCOMRichComboCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_STOCK(CKCOMRichComboCtrl, "CtrlType", 255, CKCOMRichComboCtrl::GetCtrlType, COleControl::SetNotSupported, VT_I2)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CKCOMRichComboCtrl, COleControl)
	//{{AFX_EVENT_MAP(CKCOMRichComboCtrl)
	EVENT_CUSTOM("Scroll", FireScroll, VTS_PI2)
	EVENT_CUSTOM("ScrollAfter", FireScrollAfter, VTS_NONE)
	EVENT_CUSTOM("CloseUp", FireCloseUp, VTS_NONE)
	EVENT_CUSTOM("DropDown", FireDropDown, VTS_NONE)
	EVENT_CUSTOM("PositionList", FirePositionList, VTS_BSTR)
	EVENT_CUSTOM("Change", FireChange, VTS_NONE)
	EVENT_STOCK_CLICK()
	EVENT_STOCK_KEYDOWN()
	EVENT_STOCK_KEYPRESS()
	EVENT_STOCK_KEYUP()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CKCOMRichComboCtrl, 4)
	PROPPAGEID(CKCOMRichComboPropPage::guid)
	PROPPAGEID(CKCOMRichGridPropColumnPage::guid)
	PROPPAGEID(CLSID_StockColorPage)
	PROPPAGEID(CLSID_StockFontPage)
END_PROPPAGEIDS(CKCOMRichComboCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMRichComboCtrl, "KCOMDATA.KCOMRichComboCtrl.1",
	0x29288c9, 0x2fff, 0x11d3, 0xb4, 0x46, 0, 0x80, 0xc8, 0xf1, 0x85, 0x22)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CKCOMRichComboCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DKCOMRichCombo =
		{ 0x29288c7, 0x2fff, 0x11d3, { 0xb4, 0x46, 0, 0x80, 0xc8, 0xf1, 0x85, 0x22 } };
const IID BASED_CODE IID_DKCOMRichComboEvents =
		{ 0x29288c8, 0x2fff, 0x11d3, { 0xb4, 0x46, 0, 0x80, 0xc8, 0xf1, 0x85, 0x22 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwKCOMRichComboOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CKCOMRichComboCtrl, IDS_KCOMRICHCOMBO, _dwKCOMRichComboOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboCtrl::CKCOMRichComboCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMRichComboCtrl

BOOL CKCOMRichComboCtrl::CKCOMRichComboCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_KCOMRICHCOMBO,
			IDB_KCOMRICHCOMBO,
			afxRegApartmentThreading,
			_dwKCOMRichComboOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboCtrl::CKCOMRichComboCtrl - Constructor

CKCOMRichComboCtrl::CKCOMRichComboCtrl() : m_fntCell(NULL), m_fntHeader(NULL)
{
	InitializeIIDs(&IID_DKCOMRichCombo, &IID_DKCOMRichComboEvents);

	m_nColsPreset = 0;
	m_bEndReached = TRUE;
	m_bDataSrcChanged = FALSE;
	m_dwCookieRN = 0;
	m_bHaveBookmarks = FALSE;
	m_pColumnInfo = NULL;
	m_pStrings = NULL;
	m_nColumns = m_nColumnsBind = 0;
	m_nBookmarkSize = 0;

	m_spDataSource = NULL;
	m_spRowPosition = NULL;

	m_pDropDownRealWnd = NULL;
	m_bDataSrcChanged = FALSE;
	m_spDataSource = NULL;
	m_nCol = m_nRow = 0;
	m_nColsUsed = 0;

	m_strMask.Empty();
	m_strFormat.Empty();
	m_nMaxLength = 255;
	m_nFontHeight = BUTTONHEIGHT;
	m_hFont = NULL;

	for (int i = 0; i < 255; i ++)
		m_bIsDelimiter[i] = FALSE;

	m_pEdit = NULL;
	m_hSystemHook = NULL;
	m_bLButtonDown = FALSE;

	m_pBrhBack = NULL;
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboCtrl::~CKCOMRichComboCtrl - Destructor

CKCOMRichComboCtrl::~CKCOMRichComboCtrl()
{
	if (m_pEdit)
	{
		m_pEdit->DestroyWindow();
		delete m_pEdit;
		m_pEdit = NULL;
	}

	if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->DestroyWindow();
		delete m_pDropDownRealWnd;
		m_pDropDownRealWnd = NULL;
	}

	int i = m_apColumns.GetSize();

	for (; i > 0; i --)
		delete m_apColumns[i - 1];
	m_apColumns.RemoveAll();

	i = m_apColumnsProp.GetSize();

	for (; i > 0; i --)
		delete m_apColumnsProp[i - 1];

	m_apColumnsProp.RemoveAll();

	ClearInterfaces();

	IFont * pIfont = m_fntCell.m_pFont;
	if (m_hFont)
	{
		pIfont->ReleaseHfont(m_hFont);
		m_hFont = NULL;
	}
 
	if (m_hSystemHook)
	{
		UnhookWindowsHookEx(m_hSystemHook);
		m_hSystemHook = NULL;
	}

	if (m_pBrhBack)
	{
		delete m_pBrhBack;
		m_pBrhBack = NULL;
	}
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboCtrl::OnDraw - Drawing function

void CKCOMRichComboCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (m_hFont == NULL)
		InitCellFont();

	pdc->FillSolidRect(&rcBounds, TranslateColor(GetBackColor()));
	CRect rt = rcBounds;
	rt.DeflateRect(2, 2, 2, 2);

	m_rectButton = rt;
	m_rectButton.left = m_rectButton.right - BUTTONWIDTH;
	m_rectButton.bottom = m_rectButton.top + m_nFontHeight - 4;

	pdc->Draw3dRect(rcBounds.left, rcBounds.top, rcBounds.Width(),
		rcBounds.Height(), GetSysColor(COLOR_3DSHADOW), RGB(255, 255, 255));
	pdc->Draw3dRect(rcBounds.left + 1, rcBounds.top + 1, rcBounds.Width() - 2,
		rcBounds.Height() - 2, RGB(0, 0, 0), GetSysColor(COLOR_3DFACE));

	if (!AmbientUserMode())
	{
		CRect rectText = rt;
		rectText.DeflateRect(0, 2, 0, 2);
		CFont * pOldFont = 	m_fntCell.Select(pdc, rcBounds.Height(), m_cyExtent);
		pdc->SetBkColor(TranslateColor(GetBackColor()));
		pdc->SetTextColor(TranslateColor(GetForeColor()));
		pdc->ExtTextOut(rectText.left, rectText.top, ETO_CLIPPED | ETO_OPAQUE, &rectText, 
			AmbientDisplayName(), NULL);
		pdc->SelectObject(pOldFont);
	}
	
	CBrush brh;
	if (!m_bLButtonDown)
	{
		pdc->Draw3dRect(&m_rectButton, GetSysColor(COLOR_3DLIGHT), GetSysColor(COLOR_3DDKSHADOW));
		m_rectButton.DeflateRect(1, 1);
		pdc->Draw3dRect(&m_rectButton, GetSysColor(COLOR_3DHILIGHT), GetSysColor(COLOR_3DSHADOW));
		m_rectButton.DeflateRect(1, 1);
		brh.CreateSolidBrush(GetSysColor(COLOR_3DFACE));
		pdc->FillSolidRect(&m_rectButton, GetSysColor(COLOR_3DFACE));
		m_rectButton.InflateRect(2, 2);
	}
	else
	{
		brh.CreateSolidBrush(GetSysColor(COLOR_3DFACE));
		pdc->FillRect(&m_rectButton, &brh);
	}

	POINT pt[3];
	pt[0].x = m_rectButton.left + m_rectButton.Width() / 2 - 4;
	pt[0].y = m_rectButton.top + m_rectButton.Height() / 2 - 2;
	pt[1].x = pt[0].x + 8;
	pt[1].y = pt[0].y;
	pt[2].x = pt[0].x + 4;
	pt[2].y = pt[0].y + 4;
	
	CBrush brhNew;
	brhNew.CreateSolidBrush(RGB(0, 0, 0));
	CBrush * pBrhOld = pdc->SelectObject(&brhNew);
	pdc->Polygon(pt, 3);
	pdc->SelectObject(pBrhOld);

/*	if (AmbientUserMode())
	{
		int nStart, nEnd;
		m_pEdit->GetSel(nStart, nEnd);
		m_pEdit->SetWindowText(m_strText);
		m_pEdit->SetSel(nEnd, nEnd);
		pdc->SetWindowOrg(-2, -2);
		pdc->SetViewportOrg(rcBounds.left, rcBounds.top);
		m_pEdit->SendMessage(WM_PAINT, (WPARAM)(pdc->m_hDC), 0);
	}*/
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboCtrl::DoPropExchange - Persistence support

void CKCOMRichComboCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	PX_String(pPX, _T("Format"), m_strFormat, "");
	PX_String(pPX, _T("Text"), m_strText, "");
	PX_Bool(pPX, _T("ReadOnly"), m_bReadOnly, FALSE);
	PX_Short(pPX, _T("MaxLength"), m_nMaxLength, 255);
	PX_Long(pPX, _T("RowHeight"), m_nRowHeight, 20);
	PX_Bool(pPX, "ColumnHeaders", m_bColumnHeaders, TRUE);
	PX_Font(pPX, "Font", m_fntCell, &_KCOMFontDescDefault);
	PX_Font(pPX, "HeaderFont", m_fntHeader, &_KCOMFontDescDefault);
	PX_Long(pPX, "HeaderHeight", m_nHeaderHeight, 25);
	PX_Long(pPX, "RowHeight", m_nRowHeight, 25);
	PX_Long(pPX, "DataMode", m_nDataMode, 0);
	PX_Long(pPX, "DefColWidth", m_nDefColWidth, 100);
	PX_Long(pPX, "Rows", m_nRows, 3);
	PX_Long(pPX, "Cols", m_nCols, 3);
	PX_Long(pPX, "DividerType", m_nDividerType, 3);
	PX_Long(pPX, "DividerStyle", m_nDividerStyle, 0);
	PX_String(pPX, "FieldSeparator", m_strFieldSeparator, _T("(tab)"));
	PX_Color(pPX, "HeadForeColor", m_clrHeadFore, RGB(0, 0, 0));
	PX_Color(pPX, "HeadBackColor", m_clrHeadBack, GetSysColor(COLOR_3DFACE));
	PX_String(pPX, "DataField", m_strDataField);
	PX_String(pPX, "FieldSeparator", m_strFieldSeparator, _T("(tab)"));
	PX_Bool(pPX, "ListAutoPosition", m_bListAutoPosition, TRUE);
	PX_Bool(pPX, "ListWidthAutoSize", m_bListWidthAutoSize, TRUE);
	PX_Short(pPX, "MaxDropDownItems", m_nMaxDropDownItems, 8);
	PX_Short(pPX, "MinDropDownItems", m_nMinDropDownItems, 1);
	PX_Short(pPX, "ListWidth", m_nListWidth, 0);
	PX_Color(pPX, "BackColor", m_clrBack, GetSysColor(COLOR_WINDOW));
	PX_Long(pPX, "Style", m_nStyle, 0);

    if(!pPX->IsLoading())
    {
		//  The control's properties are being saved.
		// Create a CMemFile
		CMemFile memFile;

		// construct an archive for saving
		CArchive ar( &memFile, CArchive::store, 512);

		SerializeColumnProp(ar, FALSE);
		// iterate through the tabs and serialize
		try
		{
			ar.Close();
		}
		catch(...)
		{
			TRACE0("Exception while writing data to file in DoPropExchange\n");
		}
		
		//BLOBDATA blobData;
		
		DWORD dwLength = memFile.GetLength();
		m_hBlob = GlobalAlloc(GMEM_FIXED, sizeof(long)+dwLength);
 
		if(m_hBlob != NULL)
		{
        
			BYTE * pMem = (BYTE*)GlobalLock(m_hBlob);
		
			ASSERT(pMem);

			if(pMem != NULL)
			{
          
				  //  Fill in the first part of your structure with the
				  //  sizeof(your property data).
				  BYTE* p = memFile.Detach();  
				  
				  memcpy(pMem, (DWORD*)&dwLength, sizeof(DWORD));
				  
				  
				  DWORD cb = *(DWORD*) pMem;

				  // If you are looking at this ASSERT then an incorrect number of
				  // bytes is being indicated. PX_Blob will ASSERT.
				  ASSERT(cb == dwLength);

				  pMem+=sizeof(DWORD);


				  //  Fill in the second part of your structure with the actual
				  //  data that represents the control's property value.
          
				  memcpy(pMem, p, dwLength);
				  
				  memFile.Attach(p, dwLength);
				  ((__CMemFile*)&memFile)->m_bAutoDelete = TRUE;


				  pMem-=sizeof(DWORD);
				  cb = *(DWORD*) pMem;

				  // If you are looking at this ASSERT then an incorrect number of
				  // bytes is being indicated. PX_Blob will ASSERT.
				  
				  // do a sanity check!
				  ASSERT(cb == dwLength);

          
				  //  Call PX_Blob, unlock and free your global memory.
				  PX_Blob(pPX, _T("BlobProp"), m_hBlob);
 
				  GlobalUnlock(m_hBlob);
				  
				  //m_hBlob = NULL;
			}
			else
			{
				  // The GlobalLock call failed. Pass in a NULL HGLOBAL for the
				  // third parameter to PX_Blob. This will cause it to write a
				  // value of zero for the BLOB data.
				  HGLOBAL hTmp = NULL;
 
				  PX_Blob(pPX, _T("BlobProp"), hTmp);
			}
 
			if(m_hBlob !=NULL)
				::GlobalFree(m_hBlob);
			
			m_hBlob = NULL;
		}
		else
			// The GlobalAlloc call failed. Pass in a NULL HGLOBAL for the
			// third parameter to PX_Blob. This will cause it to write a
			// value of zero for the BLOB data.
			PX_Blob(pPX, _T("BlobProp"), m_hBlob);
    }
    else
    {
		CalcMask();
      
		//  Properties are being loaded into the control.
		//  Call PX_Blob to get the handle to the memory containing
		//  the property data.
		//  We just use an arbitrary property name. Else it will fail 
		//  with prop bags
		PX_Blob(pPX, _T("BlobProp"), m_hBlob);
 
		if(m_hBlob != NULL)
		{
			BYTE* pMem = (BYTE*)GlobalLock(m_hBlob);
			  
			  // Create a CMemFile
			CMemFile memFile;

			DWORD cb = *(DWORD*) pMem;

			pMem+=sizeof(cb);

			memFile.Attach(pMem, cb);

			// construct an archive for loading data
			CArchive ar( &memFile, CArchive::load, 512);

			// Lock the memory to get a pointer to the actual property
			// data.
			if(pMem != NULL && *(int *)pMem != 0)
			{
				//  Set the value of the property, unlock and free the global
				//  memory.
         
				SerializeColumnProp(ar, TRUE);
				try
				{
					ar.Close();
				}
				catch(...)
				{
					TRACE0("Exception while reading data from file in DoPropExchange\n");
				}
			
				// unlock				
				GlobalUnlock(m_hBlob);
			}
			else
				TRACE0("there is no columns in persistent data.\n");
		}
	}		
 
	GlobalFree(m_hBlob);
	m_hBlob = NULL;
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboCtrl::OnResetState - Reset control to default state

void CKCOMRichComboCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichComboCtrl message handlers

// the backdoor for property page :)
short CKCOMRichComboCtrl::GetCtrlType() 
{
	return 2;
}

void CKCOMRichComboCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_KCOMRICHCOMBO);
	dlgAbout.DoModal();
}

short CKCOMRichComboCtrl::GetMaxLength() 
{
	return m_nMaxLength;
}

void CKCOMRichComboCtrl::SetMaxLength(short nNewValue) 
{
	if (nNewValue <0 || nNewValue > 255)
		return;

	m_nMaxLength = nNewValue;

	SetModifiedFlag();

	BoundPropertyChanged(dispidMaxLength);
}

void CKCOMRichComboCtrl::SetRowHeight(long nNewValue)
{
	if (nNewValue < 0)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	m_nRowHeight = nNewValue;

	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetDefaultRowHeight(m_nRowHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRowHeight);
	InvalidateControl();
}

BOOL CKCOMRichComboCtrl::GetReadOnly() 
{
	return m_bReadOnly;
}

void CKCOMRichComboCtrl::SetReadOnly(BOOL bNewValue) 
{
	m_bReadOnly = bNewValue;
	
	SetModifiedFlag();
	BoundPropertyChanged(dispidReadOnly);
}

LPUNKNOWN CKCOMRichComboCtrl::GetDataSource() 
{
	if (m_spDataSource)
		m_spDataSource->AddRef();

	return m_spDataSource;
}

void CKCOMRichComboCtrl::SetDataSource(LPUNKNOWN newValue) 
{
	ClearInterfaces();
	if (m_spDataSource)
	{
		m_spDataSource->Release();
		m_spDataSource = NULL;
	}

	if (m_nDataMode != 0 && newValue)
	{
		ThrowError(CTL_E_SETNOTSUPPORTED, _T("Can not set datasource in manual mode"));
		return;
	}

	if (newValue == m_spDataSource)
		return;

	if (AmbientUserMode())
	{
		m_bDataSrcChanged = TRUE;
		m_bEndReached = FALSE;
	}

	if (!newValue)
	{
		m_spDataSource = NULL;
		return;
	}

	if (FAILED(newValue->QueryInterface(__uuidof(m_spDataSource), (void **)&m_spDataSource)))
		return;

	if (!AmbientUserMode())
		GetColumnInfo();

	SetModifiedFlag();
}

BSTR CKCOMRichComboCtrl::GetDataMember() 
{
	return m_strDataMember.AllocSysString();
}

void CKCOMRichComboCtrl::SetDataMember(LPCTSTR lpszNewValue) 
{
	if (m_nDataMode != 0)
		ThrowError(CTL_E_SETNOTSUPPORTED, _T("Can not set datasource in manual mode"));

	if (m_strDataMember == lpszNewValue)
		return;

	m_strDataMember = lpszNewValue;

	ClearInterfaces();

	SetModifiedFlag();
	
	if (!AmbientUserMode())
		GetColumnInfo();
	else
		m_bDataSrcChanged = TRUE;
}

long CKCOMRichComboCtrl::GetDefColWidth() 
{
	return m_nDefColWidth;
}

void CKCOMRichComboCtrl::SetDefColWidth(long nNewValue) 
{
	if (nNewValue < 0)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	m_nDefColWidth = nNewValue;
	
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetDefaultColWidth(m_nDefColWidth);

	SetModifiedFlag();
	BoundPropertyChanged(dispidDefColWidth);
}

long CKCOMRichComboCtrl::GetRowHeight() 
{
	if (m_pDropDownRealWnd)
		m_nRowHeight = m_pDropDownRealWnd->GetDefaultRowHeight();

	return m_nRowHeight;
}

long CKCOMRichComboCtrl::GetDividerType() 
{
	return m_nDividerType;
}

void CKCOMRichComboCtrl::SetDividerType(long nNewValue) 
{
	m_nDividerType = nNewValue;

	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerType);
}

long CKCOMRichComboCtrl::GetDividerStyle() 
{
	return m_nDividerStyle;
}

void CKCOMRichComboCtrl::SetDividerStyle(long nNewValue) 
{
	m_nDividerStyle = nNewValue;

	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerStyle);
}

long CKCOMRichComboCtrl::GetDataMode() 
{
	return m_nDataMode;
}

void CKCOMRichComboCtrl::SetDataMode(long nNewValue) 
{
	if (AmbientUserMode())
		ThrowError(CTL_E_SETNOTSUPPORTEDATRUNTIME);

	if (nNewValue == 0)
		m_nDataMode = 0;
	else
	{
		if (m_spDataSource)
			ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOCHANGETOMANUALMODE);
		else
			m_nDataMode = 1;
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidDataMode);
}

long CKCOMRichComboCtrl::GetLeftCol() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	if (m_pDropDownRealWnd)
		return m_pDropDownRealWnd->GetLeftCol();
	else
		return 0;
}

void CKCOMRichComboCtrl::SetLeftCol(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetLeftCol(nNewValue);
}

long CKCOMRichComboCtrl::GetFirstRow() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pDropDownRealWnd->GetTopRow();
}

void CKCOMRichComboCtrl::SetFirstRow(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_pDropDownRealWnd->SetTopRow(nNewValue);
}

long CKCOMRichComboCtrl::GetRow() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	ROWCOL nCol;

	m_pDropDownRealWnd->GetCurrentCell(m_nRow, nCol);

	return m_nRow;
}

void CKCOMRichComboCtrl::SetRow(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_nRow = nNewValue;

	if (m_pDropDownRealWnd->GetRowCount() >= (ROWCOL)m_nRow)
		m_pDropDownRealWnd->MoveTo(m_nRow);
	else
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRow);
}

long CKCOMRichComboCtrl::GetCol() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	ROWCOL nRow;
	
	m_pDropDownRealWnd->GetCurrentCell(nRow, m_nCol);
	return m_nCol;
}

void CKCOMRichComboCtrl::SetCol(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_nCol = nNewValue;

	if (m_pDropDownRealWnd->GetColCount() >= (ROWCOL)m_nCol)
	{
		ROWCOL nRow, nCol;
		m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
		m_pDropDownRealWnd->SetCurrentCell(nRow, m_nCol);
	}
	else
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidCol);
}

long CKCOMRichComboCtrl::GetRows() 
{
	if (!m_pDropDownRealWnd)
		return m_nRows;

	m_nRows = m_pDropDownRealWnd->GetRowCount();

	return m_nRows;
}

void CKCOMRichComboCtrl::SetRows(long nNewValue) 
{
	if (m_apColumns.GetSize())
		ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOROWSCOLSHAVINGCONTENT);

	if (nNewValue < 0)
		return;

	if (m_nDataMode == 0)
		ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOSETROWSINBINDMODE);
	if (AmbientUserMode())
		SetNotSupported();

	m_nRows = nNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidRows);
}

long CKCOMRichComboCtrl::GetCols() 
{
	if (!m_pDropDownRealWnd)
		return m_nCols;

	m_nCols = m_pDropDownRealWnd->GetColCount();

	return m_nCols;
}

void CKCOMRichComboCtrl::SetCols(long nNewValue) 
{
	if (m_apColumns.GetSize())
		ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOROWSCOLSHAVINGCONTENT);

	if (nNewValue < 0)
		return;
	else if (nNewValue > 255)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_MAX255COLS);

	if (m_nDataMode == 0)
		ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOSETCOLSINBINDMODE);
	if (AmbientUserMode())
		SetNotSupported();

	m_nCols = nNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidCols);
}

BOOL CKCOMRichComboCtrl::GetDroppedDown() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	if (!m_pDropDownRealWnd)
		return FALSE;

	return m_pDropDownRealWnd->IsWindowVisible();
}

void CKCOMRichComboCtrl::SetDroppedDown(BOOL bNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	if (!m_pDropDownRealWnd)
		return;

	if (m_pDropDownRealWnd->IsWindowVisible() && !bNewValue)
		HideDropDownWnd();
	else if (!m_pDropDownRealWnd->IsWindowVisible() && bNewValue)
		ShowDropDownWnd();
}

BSTR CKCOMRichComboCtrl::GetDataField() 
{
	return m_strDataField.AllocSysString();
}

void CKCOMRichComboCtrl::SetDataField(LPCTSTR lpszNewValue) 
{
	m_strDataField = lpszNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidDataField);
}

LPDISPATCH CKCOMRichComboCtrl::GetColumns() 
{
	CDropDownColumns * pColumns = new CDropDownColumns;
	pColumns->SetCtrlPtr(this);

	return pColumns->GetIDispatch(FALSE);
}

long CKCOMRichComboCtrl::GetVisibleRows() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pDropDownRealWnd->CalcBottomRowFromRect(m_pDropDownRealWnd->GetGridRect()) - 
		m_pDropDownRealWnd->GetTopRow() + 1;
}

long CKCOMRichComboCtrl::GetVisibleCols() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pDropDownRealWnd->CalcRightColFromRect(m_pDropDownRealWnd->GetGridRect()) - 
		m_pDropDownRealWnd->GetLeftCol() + 1;
}

void CKCOMRichComboCtrl::OnHeadBackColorChanged() 
{
	if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->RowHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
		m_pDropDownRealWnd->ColHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
	}
}

void CKCOMRichComboCtrl::OnHeadForeColorChanged() 
{
	if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->RowHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
		m_pDropDownRealWnd->ColHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
	}
}

BOOL CKCOMRichComboCtrl::GetColumnHeaders() 
{
	return m_bColumnHeaders;
}

void CKCOMRichComboCtrl::SetColumnHeaders(BOOL bNewValue) 
{
	m_bColumnHeaders = bNewValue;

	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->HideRows(0, 0, !m_bColumnHeaders);

	SetModifiedFlag();
	BoundPropertyChanged(dispidColumnHeaders);
}

LPFONTDISP CKCOMRichComboCtrl::GetHeadFont() 
{
	return m_fntHeader.GetFontDispatch();
}

void CKCOMRichComboCtrl::SetHeadFont(LPFONTDISP newValue) 
{
	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	m_fntHeader.InitializeFont(NULL, newValue);
	InitHeaderFont();
	if (m_pDropDownRealWnd)
		SetHeaderHeight(m_pDropDownRealWnd->GetRowHeight(0));

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadFont);
}

BSTR CKCOMRichComboCtrl::GetFieldSeparator() 
{
	return m_strFieldSeparator.AllocSysString();
}

void CKCOMRichComboCtrl::SetFieldSeparator(LPCTSTR lpszNewValue) 
{
	if (lstrlen(lpszNewValue) == 0 || (lstrlen(lpszNewValue) > 1 && lstrcmpi(
		lpszNewValue, _T("(space)")) && lstrcmpi(lpszNewValue, _T("(tab)"))))
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_FIELDSEPARATOR);

	m_strFieldSeparator = lpszNewValue;
	m_strSeparator = lpszNewValue;
	if (!m_strFieldSeparator.CompareNoCase(_T("(tab)")))
		m_strSeparator = _T("\t");
	else if (!m_strFieldSeparator.CompareNoCase(_T("(space)")))
		m_strSeparator = _T(" ");

	SetModifiedFlag();
	BoundPropertyChanged(dispidFieldSeparator);
}

void CKCOMRichComboCtrl::OnListAutoPositionChanged() 
{
	SetModifiedFlag();
	BoundPropertyChanged(dispidListAutoPosition);
}

void CKCOMRichComboCtrl::OnListWidthAutoSizeChanged() 
{
	SetModifiedFlag();
	BoundPropertyChanged(dispidListWidthAutoSize);
}

short CKCOMRichComboCtrl::GetMaxDropDownItems() 
{
	return m_nMaxDropDownItems;
}

void CKCOMRichComboCtrl::SetMaxDropDownItems(short nNewValue) 
{
	if (nNewValue < 1 || nNewValue > 100 || nNewValue < m_nMinDropDownItems)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDITEMSLIMITATION);

	m_nMaxDropDownItems = nNewValue;
	SetModifiedFlag();
	BoundPropertyChanged(dispidMaxDropDownItems);
}

short CKCOMRichComboCtrl::GetMinDropDownItems() 
{
	return m_nMinDropDownItems;
}

void CKCOMRichComboCtrl::SetMinDropDownItems(short nNewValue) 
{
	if (nNewValue < 1 || nNewValue > 100 || nNewValue > m_nMaxDropDownItems)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDITEMSLIMITATION);

	m_nMinDropDownItems = nNewValue;
	SetModifiedFlag();
	BoundPropertyChanged(dispidMinDropDownItems);
}

short CKCOMRichComboCtrl::GetListWidth() 
{
	return m_nListWidth;
}

void CKCOMRichComboCtrl::SetListWidth(short nNewValue) 
{
	if (nNewValue < 0)
		nNewValue = 0;

	m_nListWidth = nNewValue;
	SetModifiedFlag();
	BoundPropertyChanged(dispidListWidth);
}

long CKCOMRichComboCtrl::GetHeaderHeight() 
{
	return m_nHeaderHeight;
}

void CKCOMRichComboCtrl::SetHeaderHeight(long nNewValue) 
{
	if (nNewValue < 0)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	m_nHeaderHeight = nNewValue;

	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetRowHeight(0, 0, m_nHeaderHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeaderHeight);
}

void CKCOMRichComboCtrl::OnForeColorChanged() 
{
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->ChangeStandardStyle(CGXStyle().SetTextColor(TranslateColor(GetForeColor())));
	
	COleControl::OnForeColorChanged();
	BoundPropertyChanged(DISPID_FORECOLOR);
}

int CKCOMRichComboCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	CRect rect;
	GetClientRect(&rect);
	if (AmbientUserMode())
	{
		rect.DeflateRect(2, 4, BUTTONWIDTH, 4);
		m_pEdit = new CComboEdit(this);
		rect.right -= BUTTONWIDTH;
		m_pEdit->Create(ES_AUTOHSCROLL | WS_VISIBLE | WS_CHILD, rect, this, IDC_EDIT);
		m_pEdit->SetWindowText("");

		m_pDropDownRealWnd = new CDropDownRealWnd(this);
		m_pDropDownRealWnd->CreateEx(WS_EX_TOOLWINDOW,
			"GXWND", NULL, WS_CHILD | WS_BORDER, rect, GetDesktopWindow(), NULL, NULL);

		InitDropDownWnd();
		SetText(m_strText);
	}
	
	return 0;
}

void CKCOMRichComboCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	COleControl::OnKillFocus(pNewWnd);
	
	HideDropDownWnd();
}

BOOL CKCOMRichComboCtrl::OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray) 
{
	int i;

	switch (dispid)
	{
	case dispidFieldSeparator:
		pStringArray->Add(_T("(tab)"));
		pStringArray->Add(_T("(space)"));
		pCookieArray->Add(0);
		pCookieArray->Add(1);

		return TRUE;
		break;

	case dispidDataField:
		if (m_apColumnsProp.GetSize() == 0)
		{
			GetColumnInfo();

			for (i = 0; i < m_apColumnData.GetSize(); i ++)
			{
				pStringArray->Add(m_apColumnData[i]->strName);
				pCookieArray->Add(i);
			}
		}
		else
		{
			for (i = 0; i < m_apColumnsProp.GetSize(); i ++)
			{
				pStringArray->Add(m_apColumnsProp[i]->strName);
				pCookieArray->Add(i);
			}
		}

		return TRUE;
		break;
	}
	
	return COleControl::OnGetPredefinedStrings(dispid, pStringArray, pCookieArray);
}

BOOL CKCOMRichComboCtrl::OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut) 
{
	switch (dispid)
	{
	case dispidFieldSeparator:
		VariantInit(lpvarOut);
		lpvarOut->vt = VT_BSTR;
		if (dwCookie == 0)
			lpvarOut->bstrVal = CString(_T("(tab)")).AllocSysString();
		else if (dwCookie == 1)
			lpvarOut->bstrVal = CString(_T("(space)")).AllocSysString();

		return TRUE;
		break;

	case dispidDataField:
		VariantClear(lpvarOut);
		V_VT(lpvarOut) = VT_BSTR;
		if (m_apColumnsProp.GetSize() == 0)
		{
			GetColumnInfo();
			if (dwCookie >= 0 && dwCookie < (DWORD)m_apColumnData.GetSize())
				V_BSTR(lpvarOut) = m_apColumnData[dwCookie]->strName.AllocSysString();
		}
		else
		{
			if (dwCookie >= 0 && dwCookie < (DWORD)m_apColumnsProp.GetSize())
				V_BSTR(lpvarOut) = m_apColumnsProp[dwCookie]->strName.AllocSysString();
		}
		
		return TRUE;
		break;
	}
	
	return COleControl::OnGetPredefinedValue(dispid, dwCookie, lpvarOut);
}

// force the height to fit the font height
BOOL CKCOMRichComboCtrl::OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip) 
{
	if (lpRectPos->right - lpRectPos->left != m_nFontHeight)
	{
		SetControlSize(lpRectPos->right - lpRectPos->left, m_nFontHeight);
		if (AmbientUserMode() && m_pEdit)
			m_pEdit->SetWindowPos(&wndTop, 2, 4, lpRectPos->right - lpRectPos->left - 4 - BUTTONWIDTH, 
				m_nFontHeight - 8, SWP_NOZORDER);
	}
	
	return COleControl::OnSetObjectRects(lpRectPos, lpRectClip);
}

// force the height to fit the font height
BOOL CKCOMRichComboCtrl::OnSetExtent(LPSIZEL lpSizeL) 
{
	if (IsWindow(m_hWnd))
	{
		// calc and set the right size myself
		CDC * pDC = GetDC();
		pDC->HIMETRICtoLP(lpSizeL);
		lpSizeL->cy = m_nFontHeight;
		lpSizeL->cx = max(lpSizeL->cx , BUTTONWIDTH);
		pDC->LPtoHIMETRIC(lpSizeL);
		ReleaseDC(pDC);
	}

	BOOL bRet = COleControl::OnSetExtent(lpSizeL);
	InvalidateControl();
	
	return bRet;
}

void CKCOMRichComboCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	COleControl::OnSetFocus(pOldWnd);
	
	if (m_pEdit)
	{
		m_pEdit->SetFocus();
		m_pEdit->SetSel(0, 0xff);
	}
}

void CKCOMRichComboCtrl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	COleControl::OnLButtonDown(nFlags, point);

	if (m_rectButton.PtInRect(point))
	{
		m_bLButtonDown = TRUE;

		CDC * pDC = GetDC();
		CRect rect;
		GetClientRect(&rect);
		OnDraw(pDC, &rect, &rect);
		ReleaseDC(pDC);
		
		if (m_pEdit)
			OnDropDown();
	}
}

void CKCOMRichComboCtrl::OnLButtonUp(UINT nFlags, CPoint point) 
{
	m_bLButtonDown = FALSE;
	
	CDC * pDC = GetDC();
	CRect rect;
	GetClientRect(&rect);
	OnDraw(pDC, &rect, &rect);
	ReleaseDC(pDC);
	
	COleControl::OnLButtonUp(nFlags, point);
}

BOOL CKCOMRichComboCtrl::PreTranslateMessage(MSG* pMsg) 
{
	USHORT nCharShort;
	short nShiftState = _AfxShiftState();

	if (pMsg->hwnd == m_hWnd || ::IsChild(m_hWnd, pMsg->hwnd))
	switch (pMsg->message)
	{
	case WM_KEYDOWN:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyDown(&nCharShort, nShiftState);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));

		if (pMsg->wParam == VK_DOWN || pMsg->wParam == VK_UP 
			|| pMsg->wParam == VK_ESCAPE || pMsg->wParam == VK_RETURN ||
			pMsg->wParam == VK_NEXT || pMsg->wParam == VK_PRIOR)
		{
			if (m_pDropDownRealWnd && m_pDropDownRealWnd->IsWindowVisible())
			{
				m_pDropDownRealWnd->SendMessage(WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
				return TRUE;
			}
		}
		
		if (nCharShort == VK_LEFT || nCharShort == VK_RIGHT || nCharShort == VK_UP || 
			nCharShort == VK_DOWN || nCharShort == VK_DELETE || nCharShort == VK_TAB)
		{
			GetFocus()->SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);
			return TRUE;
		}
		else if (nCharShort == VK_F4)
		{
			OnDropDown();
			return TRUE;
		}

		break;

	case WM_KEYUP:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyUp(&nCharShort, nShiftState);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));

		if (nCharShort == VK_LEFT || nCharShort == VK_RIGHT || nCharShort == VK_UP || 
			nCharShort == VK_DOWN || nCharShort == VK_DELETE || nCharShort == VK_TAB)
		{
			GetFocus()->SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);
			return TRUE;
		}
		
		break;

	case WM_CHAR:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyPress(&nCharShort);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));
		
		break;

	case WM_SYSKEYDOWN:
		if (pMsg->wParam == VK_DOWN || pMsg->wParam == VK_UP)
		{
			OnDropDown();
			return TRUE;
		}	
	}
	
	return COleControl::PreTranslateMessage(pMsg);
}

BSTR CKCOMRichComboCtrl::GetFormat() 
{
	return m_strFormat.AllocSysString();
}

void CKCOMRichComboCtrl::SetFormat(LPCTSTR lpszNewValue) 
{
	if (IsValidFormat(lpszNewValue))
	{
		m_strFormat = lpszNewValue;

		if (AmbientUserMode())
		{
			CalcMask();
			return;
		}

		SetModifiedFlag();
		BoundPropertyChanged(dispidFormat);
	}
}

void CKCOMRichComboCtrl::CalcMask()
{
	BOOL bDelimiter = FALSE;
	char cCurrChr;
	int i, j;
	CString strInit;

	m_strMask.Empty();
	for (i = 0; i < 255; i++)
		m_bIsDelimiter[i] = FALSE;

	if (m_strFormat.IsEmpty())
	{
		m_strMask.Empty();
		return;
	}

	for (i = 0, j = 0; i < m_strFormat.GetLength(); i++)
	{
		cCurrChr = m_strFormat[i];
		if (!bDelimiter && cCurrChr == '/')
		{
			bDelimiter = TRUE;
			continue;
		}
		m_strMask += cCurrChr;
		if (bDelimiter)
			m_bIsDelimiter[j] = TRUE;
		strInit += m_bIsDelimiter[j] ? cCurrChr : '_';

		j ++;

		bDelimiter = FALSE;
	}

	if (!IsValidText(m_strText))
		SetText(strInit);
}

BOOL CKCOMRichComboCtrl::IsValidText(CString strNewText)
{
	if (m_strMask.IsEmpty())
		return strNewText.GetLength() <= m_nMaxLength;

	if (strNewText.GetLength() != m_strMask.GetLength())
		return FALSE;

	int i;
	for (i = 0; i < m_strMask.GetLength(); i ++)
		if (!IsValidChar(strNewText[i], i) && strNewText[i] != ' '
			&& strNewText[i] != '_' && strNewText[i] != ' ')
			return FALSE;
	
	return TRUE;
}

BOOL CKCOMRichComboCtrl::IsValidFormat(CString strNewFormat)
{
	BOOL bDelimiter = FALSE;
	char cCurrChr;

	if (strNewFormat.IsEmpty())
		return TRUE;

	if (strNewFormat.GetLength() > 255)
		return FALSE;

	for (int i = 0; i < strNewFormat.GetLength(); i++)
	{
		cCurrChr = strNewFormat[i];
		if (!bDelimiter && cCurrChr != '#' && cCurrChr != '9' &&
			cCurrChr != '?' && cCurrChr != 'C' && cCurrChr != 'A' &&
			cCurrChr != 'a' && cCurrChr != '/')
			return FALSE;

		bDelimiter = (cCurrChr == '/') ? TRUE : FALSE;
	}

	return TRUE;
}

BOOL CKCOMRichComboCtrl::IsValidChar(UINT nChar, int nPosition)
{
	if (nPosition >= m_strMask.GetLength())
		return FALSE;

	if (m_bIsDelimiter[nPosition])
		return (char)nChar == m_strMask[nPosition];

	switch (m_strMask[nPosition])
	{
	case '#':
		return nChar >= '0' && nChar <= '9';
		break;

		case '9':
		return nChar == ' ' || (nChar >= '0' && nChar <= '9');
		break;

	case '?':
		return (nChar >= 'a' && nChar <= 'z') ||
			(nChar >= 'A' && nChar <= 'Z');
		break;

	case 'C':
		return nChar == ' ' || (nChar >= 'a' && nChar <= 'z') 
			|| (nChar >= 'A' && nChar <= 'Z');
		break;

	case 'A':
		return (nChar >= 'a' && nChar <= 'z') ||
			(nChar >= 'A' && nChar <= 'Z') ||
			(nChar >= '0' && nChar <= '9');
		break;

	case 'a':
		return (nChar >= 32 && nChar <= 126) ||
			(nChar >= 128 && nChar <= 255);
		break;
	};

	return FALSE;
}

ROWCOL CKCOMRichComboCtrl::GetColOrdinalFromIndex(ROWCOL nColIndex)
{
	for (int i = 0; i < m_apColumns.GetSize(); i ++)
		if (m_apColumns[i]->prop.nColIndex == nColIndex)
			return i;

	return GX_INVALID;
}

ROWCOL CKCOMRichComboCtrl::GetColOrdinalFromCol(ROWCOL nCol)
{
	if (m_pDropDownRealWnd)
		return GetColOrdinalFromIndex(m_pDropDownRealWnd->GetColIndex(nCol));

	return GX_INVALID;
}

ROWCOL CKCOMRichComboCtrl::GetColOrdinalFromField(int nField)
{
	for (int i = 0; i < m_apColumns.GetSize(); i ++)
		if (m_apColumns[i]->prop.nDataField == nField)
			return i;

	return GX_INVALID;
}

int CKCOMRichComboCtrl::GetFieldFromStr(CString str)
{
	if (str.IsEmpty())
		return -1;

	for (int i = 0; i < m_apColumnData.GetSize(); i ++)
		if (!str.CompareNoCase(m_apColumnData[i]->strName))
			return m_apColumnData[i]->nColumn;

	return -1;
}

void CKCOMRichComboCtrl::UpdateDivider()
{
	if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->ChangeStandardStyle((CGXStyle().SetBorders(
			gxBorderAll, CGXPen().SetStyle(PS_NULL))));
		
		switch (m_nDividerType)
		{
		case 1:
			m_pDropDownRealWnd->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderRight, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1))));
			break;
			
		case 2:
			m_pDropDownRealWnd->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderBottom, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1))));
			break;
			
		case 3:
			m_pDropDownRealWnd->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderRight, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1))));
			m_pDropDownRealWnd->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderBottom, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1))));
			break;
		}

		m_pDropDownRealWnd->RowHeaderStyle().SetBorders(gxBorderAll, CGXPen().SetStyle(PS_NULL));
		m_pDropDownRealWnd->ColHeaderStyle().SetBorders(gxBorderAll, CGXPen().SetStyle(PS_NULL));
	}
}

void CKCOMRichComboCtrl::InitCellFont()
{
	HFONT hFont;

	if (m_fntCell.m_pFont->get_hFont(&hFont) == S_OK)
	{
		LOGFONT logCell;
		if (GetObject(hFont, sizeof(logCell), &logCell))
		{
			if (m_pDropDownRealWnd)
			{
				m_pDropDownRealWnd->ChangeStandardStyle(CGXStyle().SetFont(CGXFont().SetLogFont(logCell)));
				m_pDropDownRealWnd->UpdateFontMetrics();
			}
		}

		IFont * pIfont = m_fntCell.m_pFont;
		//release previous IFont interface pointer
		if (m_hFont)
			pIfont->ReleaseHfont(m_hFont);
		
		CY cy;
		SIZEL size;
		int x, y;
		pIfont->get_Size(&cy);
		m_hFont = hFont;
		// prevent the IFont interface from releasing gotten pointer
		pIfont->AddRefHfont(m_hFont);
		
		if (IsWindow(m_hWnd))
		{
			CDC * pDC = GetDC();
			GetControlSize(&x, &y);
			
			// calc the font height
			m_nFontHeight = cy.Lo * GetDeviceCaps(pDC->m_hDC, LOGPIXELSY) / 720000 + 12;
			
			size.cx = x;
			size.cy = m_nFontHeight;
			pDC->LPtoHIMETRIC(&size);
			// resize the control according to the new font
			OnSetExtent(&size);
//			SetControlSize(x, m_nFontHeight);
			
			ReleaseDC(pDC);
		}

		if (m_pEdit)
			m_pEdit->SendMessage(WM_SETFONT, (WPARAM)m_hFont, MAKELPARAM(TRUE, 0));
	}
}

LPFONTDISP CKCOMRichComboCtrl::GetFont() 
{
	return m_fntCell.GetFontDispatch();
}

void CKCOMRichComboCtrl::SetFont(LPFONTDISP newValue) 
{
	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	m_fntCell.InitializeFont(NULL, newValue);

	InitCellFont();
	
	if (m_pDropDownRealWnd)
		SetRowHeight(m_pDropDownRealWnd->GetFontHeight() + 4);
	else
		SetRowHeight(m_nFontHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidFont);
}

void CKCOMRichComboCtrl::SerializeColumnProp(CArchive &ar, BOOL bLoad)
{
	if (!bLoad)
	{
		int nCols = m_apColumnsProp.GetSize();
		ColumnProp * pCol;

		ar << nCols;

		for (int i = 0; i < nCols; i ++)
		{
			pCol = m_apColumnsProp[i];

			ar << pCol->bAllowSizing << pCol->bForceLock << pCol->bForceNoNullable;
			ar << pCol->bLocked << pCol->bNullable << pCol->bVisible;
			ar << pCol->clrBack << pCol->clrFore << pCol->clrHeadBack << pCol->clrHeadFore;
			ar << pCol->nAlignment << pCol->nCaptionAlignment << pCol->nCase;
			ar << pCol->nDataType << pCol->nFieldLen << pCol->nStyle;
			ar << pCol->nWidth << pCol->strCaption << pCol->strChoiceList;
			ar << pCol->strDataField << pCol->strMask << pCol->strName;
			ar << pCol->strPromptChar << pCol->bPromptInclude;
		}

		ar << m_arContent.GetSize();
		for (i = 0; i < m_arContent.GetSize(); i ++)
			ar << m_arContent[i];
	}
	else
	{
		int nCols;
		ColumnProp * pCol;
		
		ar >> nCols;
		m_nColsPreset = nCols;

		for (int i = 0; i < nCols; i ++)
		{
			pCol = new ColumnProp;

			ar >> pCol->bAllowSizing >> pCol->bForceLock >> pCol->bForceNoNullable;
			ar >> pCol->bLocked >> pCol->bNullable >> pCol->bVisible;
			ar >> pCol->clrBack >> pCol->clrFore >> pCol->clrHeadBack >> pCol->clrHeadFore;
			ar >> pCol->nAlignment >> pCol->nCaptionAlignment >> pCol->nCase;
			ar >> pCol->nDataType >> pCol->nFieldLen >> pCol->nStyle;
			ar >> pCol->nWidth >> pCol->strCaption >> pCol->strChoiceList;
			ar >> pCol->strDataField >> pCol->strMask >> pCol->strName;
			ar >> pCol->strPromptChar >> pCol->bPromptInclude;

			m_apColumnsProp.Add(pCol);
		}

		CString strContent;
		int nContents;
		m_arContent.RemoveAll();
		ar >> nContents;

		while (nContents > 0)
		{
			ar >> strContent;
			m_arContent.Add(strContent);
			nContents --;
		}
	}
}

void CKCOMRichComboCtrl::ClearInterfaces()
{
	FreeColumnMemory();
	FreeColumnData();

	if (m_dwCookieRN != 0)
	{
		AfxConnectionUnadvise(m_Rowset.m_spRowset, IID_IRowsetNotify, &m_xRowsetNotify, FALSE, m_dwCookieRN);
		m_dwCookieRN = 0;
	}
	
	if (m_spDataSource && AmbientUserMode())
		m_spDataSource->removeDataSourceListener((IDataSourceListener*)&m_xDataSourceListener);

	if (m_spRowPosition)
	{
		m_spRowPosition->Release();
		m_spRowPosition = NULL;
	}
}

void CKCOMRichComboCtrl::InitColumnObject(int nCol, CDropDownColumn *pCol)
{
	CString strName;

	strName.Format("Column%4d", m_nColsUsed ++);
	
	if (m_hWnd)
		COleControl::SetRedraw(FALSE);

	pCol->SetCtrlPtr(this);
	pCol->prop.nColIndex = m_pDropDownRealWnd->GetColIndex(nCol);
	pCol->SetName(strName);
	pCol->SetCaption(strName);
	pCol->SetBackColor(GetBackColor());
	pCol->SetForeColor(GetForeColor());
	pCol->SetHeadBackColor(m_clrHeadBack);
	pCol->SetHeadForeColor(m_clrHeadFore);
	pCol->SetAlignment(2);
	pCol->SetCaptionAlignment(2);
	pCol->SetCase(0);
	pCol->SetVisible(TRUE);
	pCol->SetWidth(m_nDefColWidth);

	if (m_hWnd)
		COleControl::SetRedraw();
	
	InvalidateControl();
}

void CKCOMRichComboCtrl::FreeColumnMemory()
{
	if (m_pColumnInfo)
	{
		CoTaskMemFree(m_pColumnInfo);
		m_pColumnInfo = NULL;
	}
	if (m_pStrings)
	{
		CoTaskMemFree(m_pStrings);
		m_pStrings = NULL;
	}
	FreeBookmarkMemory();
}

void CKCOMRichComboCtrl::InitGridFromProp()
{
	CDropDownColumn * pCol;
	ColumnProp * pProp;

	int nCols = m_apColumnsProp.GetSize();
	int nRows = nCols == 0 ? 0 : m_arContent.GetSize() / nCols;

	if (nCols > 0)
		m_nCols = nCols;
	if (m_nDataMode == 0)
		m_nRows = AmbientUserMode() ? 0 : 3;
	else if (nCols > 0)
		m_nRows = nRows;

	if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
		m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
		m_pDropDownRealWnd->SetColCount(0);
		m_pDropDownRealWnd->SetRowCount(0);
		
		m_pDropDownRealWnd->SetColCount(m_nCols);
		m_pDropDownRealWnd->SetRowCount(m_nRows);
	}

	int i;
	for (i = 0; i < m_apColumns.GetSize(); i ++)
		delete m_apColumns[i];

	m_apColumns.RemoveAll();

	for (i = 0; i < nCols; i++)
	{
		pCol = new CDropDownColumn;
		pProp = m_apColumnsProp[i];

		pCol->SetCtrlPtr(this);
		
		if (m_pDropDownRealWnd)
			pCol->prop.nColIndex = m_pDropDownRealWnd->GetColIndex(i + 1);

		pCol->prop.strDataField = pProp->strDataField;
		pCol->prop.nDataField = pProp->nDataField;
			
		pCol->SetAlignment(pProp->nAlignment);
		pCol->SetBackColor(pProp->clrBack);
		pCol->SetCaption(pProp->strCaption);
		pCol->SetCaptionAlignment(pProp->nCaptionAlignment);
		pCol->SetCase(pProp->nCase);
		pCol->SetForeColor(pProp->clrFore);
		pCol->SetHeadBackColor(pProp->clrHeadBack);
		pCol->SetHeadForeColor(pProp->clrHeadFore);
		pCol->SetName(pProp->strName);
		pCol->SetVisible(pProp->bVisible);
		pCol->SetWidth(pProp->nWidth);

		m_apColumns.Add(pCol);
	}

	if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
	}

	BoundPropertyChanged(dispidRows);
	BoundPropertyChanged(dispidCols);

	SetModifiedFlag();
}

// we do not support some data type, so to filter them out
BOOL CKCOMRichComboCtrl::GetColumnInfo()
{
	USES_CONVERSION;

	ClearInterfaces();

	HRESULT hr;
	int nColumn = 0;

	// We can't do anything if the datasource hasn't been set
	if (m_spDataSource == NULL)
		return FALSE;

	m_nColumns = 0; // Reset the column count
	// Get the rowset from the data control
	hr = GetRowset();
	if (FAILED(hr) || !m_Rowset.m_spRowset)
		return FALSE;

	// Get the column information so we know the column names etc.
	hr = m_Rowset.CAccessorRowset<CManualAccessor>::GetColumnInfo((ULONG *)&m_nColumns, &m_pColumnInfo, &m_pStrings);
	if (FAILED(hr) || !m_pColumnInfo)
		return FALSE;

	// Check if we have bookmarks we can store
	if (m_pColumnInfo->iOrdinal == 0)
	{
		m_bHaveBookmarks = TRUE;
		FreeBookmarkMemory();
		m_nBookmarkSize = m_pColumnInfo->ulColumnSize;
		nColumn = 1;
	}
	else
	{
		m_bHaveBookmarks = FALSE;

		AfxMessageBox(IDS_ERROR_NOBOOKMARK, MB_ICONERROR | MB_OK);
		return FALSE;
	}

	// Now we can create the accessor and bind the data
	// First of all we need to find the column ordinal from the name
	int nColumnsBound = 0;
	for (; nColumn < m_nColumns; nColumn++)
	{
		switch (m_pColumnInfo[nColumn].wType)
		{
		case DBTYPE_I2:
		case DBTYPE_I4:
		case DBTYPE_R4:
		case DBTYPE_R8:
		case DBTYPE_CY:
		case DBTYPE_DATE:
		case DBTYPE_BSTR:
		case DBTYPE_BOOL:
		case DBTYPE_DECIMAL:
		case DBTYPE_UI1:
		case DBTYPE_I1:
		case DBTYPE_UI2:
		case DBTYPE_UI4:
		case DBTYPE_I8:
		case DBTYPE_UI8:
		case DBTYPE_FILETIME:
		case DBTYPE_STR:
		case DBTYPE_WSTR:
		case DBTYPE_NUMERIC:
		case DBTYPE_UDT:
		case DBTYPE_DBDATE:
		case DBTYPE_DBTIME:
		case DBTYPE_DBTIMESTAMP:
			nColumnsBound ++;
			if (m_nColumnsBind < 254)
			{
				ColumnData * pCld = new ColumnData;
				pCld->strName = OLE2T(m_pColumnInfo[nColumn].pwszName);
				pCld->nColumn = nColumn;
				pCld->vt = m_pColumnInfo[nColumn].wType;
				m_apColumnData.Add(pCld);
			}
			break;
		}
	}

	m_nColumnsBind = min(m_nColumnsBind, m_apColumnData.GetSize());

	return m_apColumnData.GetSize() > 0;
}

HRESULT CKCOMRichComboCtrl::GetRowset()
{
	USES_CONVERSION;

	ASSERT(m_spDataSource != NULL);
	HRESULT hr;

	// Close anything we had open and ClearInterfaces if necessary
	ClearInterfaces();

	m_Rowset.CAccessorRowset<CManualAccessor>::Close();

	hr = m_spDataSource->getDataMember(m_strDataMember.AllocSysString(), IID_IRowPosition,
		(IUnknown **)&m_spRowPosition);
	if (FAILED(hr) || !m_spRowPosition)
		return hr;

	return m_spRowPosition->GetRowset(IID_IRowset, (IUnknown**)&m_Rowset.m_spRowset);
}

void CKCOMRichComboCtrl::InitDropDownWnd()
{
	m_pDropDownRealWnd->Initialize();
	m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);

	UpdateDivider();

	m_pDropDownRealWnd->GetParam()->EnableSelection(GX_SELFULL);
	m_pDropDownRealWnd->GetParam()->EnableMoveRows(FALSE);
	m_pDropDownRealWnd->GetParam()->EnableMoveCols(FALSE);
	m_pDropDownRealWnd->GetParam()->EnableTrackColWidth(FALSE);
	m_pDropDownRealWnd->GetParam()->EnableTrackRowHeight(FALSE);
	m_pDropDownRealWnd->GetParam()->SetSpecialMode(GX_MODELBOX_SS);
	m_pDropDownRealWnd->GetParam()->SetHideCurrentCell(GX_HIDE_ALLWAYS);
	m_pDropDownRealWnd->GetParam()->SetActivateCellFlags(FALSE);
	m_pDropDownRealWnd->GetParam()->GetProperties()->SetMarkColHeader(FALSE);
	m_pDropDownRealWnd->GetParam()->GetProperties()->SetMarkRowHeader(FALSE);

	m_pDropDownRealWnd->ChangeStandardStyle(CGXStyle().SetTextColor(TranslateColor(GetForeColor())));
	m_pDropDownRealWnd->ChangeStandardStyle(CGXStyle().SetInterior(TranslateColor(GetBackColor())));
	m_pDropDownRealWnd->ChangeStandardStyle(CGXStyle().SetVerticalAlignment(DT_VCENTER));
	m_pDropDownRealWnd->RowHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
	m_pDropDownRealWnd->RowHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
	m_pDropDownRealWnd->ColHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
	m_pDropDownRealWnd->ColHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));

	InitHeaderFont();
	InitCellFont();
	
	m_pDropDownRealWnd->SetRowHeight(0, 0, m_nHeaderHeight);
	m_pDropDownRealWnd->SetDefaultRowHeight(m_nRowHeight);
	m_pDropDownRealWnd->SetDefaultColWidth(m_nDefColWidth);

	if (m_nDataMode)
		m_pDropDownRealWnd->SetRowCount(m_nRows);
	
	InitGridFromProp();
	
	CGXData * pData = m_pDropDownRealWnd->GetParam()->GetData();
	if (m_nDataMode)
	{
		for (int i = 0, j = 0; i < m_nCols; i ++)
		{
			for (int k = 0; k < m_nRows && j < m_arContent.GetSize(); k ++)
			{
				pData->StoreValueRowCol(k + 1, i + 1, m_arContent[j ++], gxOverride);
			}
		}
	}
	
	if (m_nDataMode && m_apColumns.GetSize() == 0)
	{
		CDropDownColumn * pNewCol;
		for (int i = 0; i < m_nCols; i ++)
		{
			pNewCol = new CDropDownColumn;
			InitColumnObject(i + 1, pNewCol);
			m_apColumns.Add(pNewCol);
		}
	}
	
	m_pDropDownRealWnd->SetReadOnly(TRUE);
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);

	m_pDropDownRealWnd->HideRows(0, 0, !m_bColumnHeaders);
	m_pDropDownRealWnd->HideCols(0, 0, TRUE);
}

void CKCOMRichComboCtrl::InitHeaderFont()
{
	HFONT hFont;
	
	if (m_fntHeader.m_pFont->get_hFont(&hFont) == S_OK)
	{
		LOGFONT logHeader;
		if (GetObject(hFont, sizeof(logHeader), &logHeader))
		{
			if (m_pDropDownRealWnd)
			{
				m_pDropDownRealWnd->ColHeaderStyle().SetFont(CGXFont().SetLogFont(logHeader));
				m_pDropDownRealWnd->UpdateFontMetrics();
				m_pDropDownRealWnd->ResizeRowHeightsToFit(CGXRange().SetRows(0, 0));
			}
		}
	}
}

void CKCOMRichComboCtrl::GetCellValue(ROWCOL nRow, ROWCOL nCol, CGXStyle &style)
{
	if (m_nDataMode != 0 || !AmbientUserMode())
		return;
	
	if (!m_bEndReached && nRow == (ROWCOL)m_pDropDownRealWnd->OnGetRecordCount())
		FetchNewRows();

	if ((ROWCOL)m_apBookmark.GetSize() < nRow)
		return;

	CBookmark<0> bmk;
	HRESULT hr;

	GetBmkOfRow(nRow);
	if (!m_apBookmark[nRow - 1])
		return;

	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
	hr = m_Rowset.MoveToBookmark(bmk);
	
	if (FAILED(hr) || hr == DB_S_ENDOFROWSET || m_Rowset.m_hRow == NULL)
	{
		m_Rowset.FreeRecordMemory();
		return;
	}
	else if (hr != S_OK)
		m_Rowset.GetData();
	
	CDropDownColumn * pCol;
	ROWCOL nColToSet;
	for (int j = 0; j < m_apColumns.GetSize(); j ++)
	{
		pCol = m_apColumns[j];
		if (pCol->prop.nDataField == -1)
			continue;
		
		nColToSet = m_pDropDownRealWnd->GetColFromIndex(pCol->prop.nColIndex);
		if (nColToSet != nCol)
			continue;

		if (m_data.dwStatus[j] == DBSTATUS_S_OK && style.GetValue().Compare(m_data.strData[j]))
			style.SetValue(m_data.strData[j]);
		else if (m_data.dwStatus[j] == DBSTATUS_S_ISNULL && !style.GetValue().IsEmpty())
			style.SetValue(_T(""));
	}
	
	m_Rowset.FreeRecordMemory();
}

ROWCOL CKCOMRichComboCtrl::GetRowColFromVariant(VARIANT va)
{
	if (va.vt == VT_ERROR)
		return GX_INVALID;

	COleVariant vaNew = va;
	vaNew.ChangeType(VT_BSTR);
	vaNew.ChangeType(VT_I4);

	return vaNew.lVal;
}

ROWCOL CKCOMRichComboCtrl::GetRowFromBmk(BYTE *bmk)
{
	BYTE * bmkSearch = new BYTE[m_nBookmarkSize];
	memcpy(bmkSearch, bmk, sizeof(bmk));

	ROWCOL i, nRows = m_apBookmark.GetSize();
	for (i = 0; i < nRows; i++)
	{
		GetBmkOfRow(i + 1);
		if (m_apBookmark[i] && memcmp(m_apBookmark[i], bmkSearch, m_pColumnInfo->ulColumnSize) == 0)
		{
			delete [] bmkSearch;
			return i + 1;
		}
	}
	
	delete [] bmkSearch;
	return GX_INVALID;
}

ROWCOL CKCOMRichComboCtrl::GetRowFromHRow(HROW hRow)
{
	if (hRow == 0)
		return GX_INVALID;

	HROW hRowOld = m_Rowset.m_hRow;
	m_Rowset.m_hRow = hRow;
	HRESULT hr;

	hr = m_Rowset.GetData();
	if (FAILED(hr))
	{
		m_Rowset.m_hRow = hRowOld;
		return GX_INVALID;
	}

	ROWCOL nRow;
	nRow = GetRowFromBmk(m_data.bookmark);

	m_Rowset.FreeRecordMemory();
	m_Rowset.m_hRow = hRowOld;

	return nRow;
}

void CKCOMRichComboCtrl::FetchNewRows()
{
	USES_CONVERSION;

	m_pDropDownRealWnd->CancelRecord();

	int i = 0, nFetchOnce;
	ULONG nRowsAvailable = 0;

	// this function is useful when bound with RDC datasource, because
	// the nRowsAvailable is just the number of the records in the local
	// cache
	if (m_Rowset.GetApproximatePosition(NULL, NULL, &nRowsAvailable) == S_OK)
		nFetchOnce = nRowsAvailable - m_apBookmark.GetSize();
	else
		return;

	if (nFetchOnce == 0)
	{
		m_bEndReached = TRUE;
		return;
	}

	for (i = 0; i < nFetchOnce; i ++)
		m_apBookmark.Add(NULL);

	nFetchOnce = nRowsAvailable - m_pDropDownRealWnd->OnGetRecordCount();
	if (nFetchOnce <= 0)
		return;
	
	m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
	m_pDropDownRealWnd->InsertRows(m_pDropDownRealWnd->OnGetRecordCount() + 1, nFetchOnce);

	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
	m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
	m_pDropDownRealWnd->Redraw();
	m_pDropDownRealWnd->UpdateScrollbars();
}

void CKCOMRichComboCtrl::UndeleteRecordFromHRow(HROW hRow)
{
	m_Rowset.ReleaseRows();
	m_Rowset.m_hRow = hRow;
	if (FAILED(m_Rowset.GetData()))
		return;

	ULONG nRow;
	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_data.bookmark);
	if (FAILED(m_Rowset.GetApproximatePosition(&bmk, &nRow, NULL) || 
		nRow > m_pDropDownRealWnd->GetRowCount()))
	nRow = m_pDropDownRealWnd->GetRowCount();

	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
	
	BYTE * pBookmarkCopy;
	if (m_bHaveBookmarks)
	{
		pBookmarkCopy = new BYTE[m_nBookmarkSize];
		if (pBookmarkCopy != NULL)
			memcpy(pBookmarkCopy, &m_data.bookmark, m_nBookmarkSize);
		
		m_apBookmark.InsertAt(nRow - 1, pBookmarkCopy);
	}
	
	m_pDropDownRealWnd->InsertRows(nRow, 1);
	
	m_Rowset.FreeRecordMemory();

	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
	m_pDropDownRealWnd->Redraw();
}

void CKCOMRichComboCtrl::GetBmkOfRow(ROWCOL nRow)
{
	if (nRow < 1 || nRow > (ROWCOL)m_apBookmark.GetSize() || m_apBookmark[nRow - 1])
		return;

	BYTE * pBookmarkCopy;
	DBBOOKMARK bmFirst = DBBMK_FIRST;
	CBookmark<0> bmk;
	
	bmk.SetBookmark(m_nBookmarkSize, (BYTE *)&bmFirst);
	HRESULT hr = m_Rowset.MoveToBookmark(bmk, nRow - 1);
	if (FAILED(hr) || hr == DB_S_ENDOFROWSET || m_Rowset.m_hRow == NULL)
	{
		m_Rowset.FreeRecordMemory();
		return;
	}
	else if (hr != S_OK)
		m_Rowset.GetData();
	
	if (m_bHaveBookmarks)
	{
		pBookmarkCopy = new BYTE[m_nBookmarkSize];
		if (pBookmarkCopy != NULL)
			memcpy(pBookmarkCopy, &m_data.bookmark, m_nBookmarkSize);
		
		m_apBookmark[nRow - 1] = pBookmarkCopy;
	}
}

void CKCOMRichComboCtrl::HideDropDownWnd()
{
	if (!m_pDropDownRealWnd || !m_pDropDownRealWnd->IsWindowVisible())
		return;

	FireCloseUp();

	UnhookWindowsHookEx(m_hSystemHook);
	m_hSystemHook = NULL;

	m_pDropDownRealWnd->ShowWindow(SW_HIDE);
}

void CKCOMRichComboCtrl::UpdateDropDownWnd()
{
	USES_CONVERSION;

	if (!AmbientUserMode())
		return;

	m_pDropDownRealWnd->LockUpdate(TRUE);
	m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);

	m_bDataSrcChanged = FALSE;

	if (!GetColumnInfo())
	{
		m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
		m_pDropDownRealWnd->LockUpdate(FALSE);
		m_pDropDownRealWnd->Redraw();

		return;
	}

	int i;
	if (m_nColsPreset == 0)
	{
		i = m_apColumnsProp.GetSize();
		
		for (; i > 0; i --)
			delete m_apColumnsProp[i - 1];
		
		m_apColumnsProp.RemoveAll();
	}

	if (m_apColumnsProp.GetSize() == 0)
	{
		ColumnProp * pCol;
		m_nColumnsBind = m_apColumnData.GetSize();
		for (i = 0; i < m_nColumnsBind; i ++)
		{
			pCol = new ColumnProp;

			pCol->nDataField = m_apColumnData[i]->nColumn;
			pCol->strDataField = m_apColumnData[i]->strName;

			pCol->bVisible = TRUE;
			pCol->nAlignment = 2;
			pCol->nCaptionAlignment = 2;
			pCol->nCase = 0;
			pCol->nWidth = m_nDefColWidth;
			pCol->strCaption = m_apColumnData[i]->strName;
			pCol->strName = m_apColumnData[i]->strName;

			m_apColumnsProp.Add(pCol);
		}
	}
	else
	{
		ColumnProp * pCol;

		for (i = 0; i < m_apColumnsProp.GetSize(); i ++)
		{
			pCol = m_apColumnsProp[i];
			pCol->nDataField = GetFieldFromStr(pCol->strDataField);
		}
	}

	InitDropDownWnd();

	Bind();

	int nEntries = m_nColumnsBind + (m_bHaveBookmarks ? 1 : 0);
	m_Rowset.CreateAccessor(nEntries, &m_data, sizeof(m_data));
	if (m_bHaveBookmarks)
		m_Rowset.AddBindEntry(0, DBTYPE_BYTES, sizeof(m_data.bookmark), &m_data.bookmark);

	for (i = 0; i < m_apColumns.GetSize(); i ++)
	{
		if (m_apColumns[i]->prop.nDataField != -1)
			m_Rowset.AddBindEntry(m_apColumns[i]->prop.nDataField, DBTYPE_STR, 255 * sizeof(TCHAR), 
				&(m_data.strData[i]), NULL, &(m_data.dwStatus[i]));
	}

	HRESULT hr = m_Rowset.Bind();
	if (FAILED(hr))
	{
		m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
		m_pDropDownRealWnd->LockUpdate(FALSE);
		m_pDropDownRealWnd->Redraw();

		return;
	}

	m_Rowset.SetupOptionalRowsetInterfaces();
	FetchNewRows();
	
	// Set up the sink so we know when the rowset is repositioned
	if (AmbientUserMode())
	{
		AfxConnectionAdvise(m_Rowset.m_spRowset, IID_IRowsetNotify, &m_xRowsetNotify, FALSE, &m_dwCookieRN);
		m_spDataSource->addDataSourceListener((IDataSourceListener*)&m_xDataSourceListener);
	}

	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
	m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
	m_pDropDownRealWnd->LockUpdate(FALSE);
	m_pDropDownRealWnd->Redraw();
}

void CKCOMRichComboCtrl::DeleteRowFromSrc(ROWCOL nRow)
{
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
	m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);

	m_pDropDownRealWnd->DeleteRows(nRow, nRow);

	delete [] m_apBookmark[nRow - 1];
	m_apBookmark.RemoveAt(nRow - 1);

	m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
	m_pDropDownRealWnd->Redraw();
}

ROWCOL CKCOMRichComboCtrl::SearchText(CString strSearch, BOOL bAutoPosition)
{
	if (bAutoPosition && !m_bListAutoPosition)
		return GX_INVALID;

	FirePositionList(strSearch);

	for (int i = 0; i < m_apColumns.GetSize(); i ++)
	{
		if (!m_apColumns[i]->prop.strName.CompareNoCase(m_strDataField))
			break;
	}

	if (i >= m_apColumns.GetSize())
		return GX_INVALID;

	ROWCOL nCol = m_pDropDownRealWnd->GetColFromIndex(m_apColumns[i]->prop.nColIndex);
	CString strValue;

	if (m_apColumns[i]->prop.nDataField == -1)
	{
		for (i = 1; (ROWCOL)i <= m_pDropDownRealWnd->GetRowCount(); i ++)
		{
			strValue = m_pDropDownRealWnd->GetValueRowCol(i, nCol);
			if (!strValue.Left(strSearch.GetLength()).CompareNoCase(strSearch))
			{
				m_nRow = i;
				break;
			}
		}
		
		if ((ROWCOL)i > m_pDropDownRealWnd->GetRowCount())
			return GX_INVALID;
	}
	else
	{
		if (m_Rowset.m_spRowset == NULL)
			return GX_INVALID;

		CComQIPtr<IRowsetFind, &IID_IRowsetFind> pFind = m_Rowset.m_spRowset;
		if (!pFind)
			return GX_INVALID;

		DBBOOKMARK bmFirst = DBBMK_FIRST;
		CAccessorRowset<CManualAccessor> rac;
		rac.m_spRowset = m_Rowset.m_spRowset;
		RowData bindData;
	
		rac.CreateAccessor(1, &bindData, sizeof(bindData));
		rac.AddBindEntry(m_apColumns[i]->prop.nDataField, DBTYPE_STR, 255 * sizeof(TCHAR), 
			&(m_data.strData[0]), NULL, &(m_data.dwStatus[0]));

		if (strSearch.IsEmpty())
			m_data.dwStatus[0] = DBSTATUS_S_ISNULL;
		else
		{
			m_data.dwStatus[0] = DBSTATUS_S_OK;
			lstrcpy(m_data.strData[0], strSearch);
		}

		if (FAILED(rac.Bind()))
			return GX_INVALID;

		ULONG nRowsOb = 0;
		HROW* phRow = &rac.m_hRow;

		// first use partly match search, if fails, use exactly match search
		// not efficient but simple :)
		pFind->FindNextRow(NULL, rac.m_pAccessor->GetHAccessor(0), rac.
			m_pAccessor->GetBuffer(), DBCOMPAREOPS_EQ | DBCOMPAREOPS_CASEINSENSITIVE, 
			m_nBookmarkSize, (BYTE *)&bmFirst, 0, 1, &nRowsOb, &phRow);

		if (rac.m_hRow == 0)
			pFind->FindNextRow(NULL, rac.m_pAccessor->GetHAccessor(0), rac.
				m_pAccessor->GetBuffer(), DBCOMPAREOPS_BEGINSWITH | DBCOMPAREOPS_CASEINSENSITIVE, 
				m_nBookmarkSize, (BYTE *)&bmFirst, 0, 1, &nRowsOb, &phRow);

		if (rac.m_hRow == 0)
			return GX_INVALID;

		HROW hRow = rac.m_hRow;
		rac.ReleaseRows();
		m_Rowset.ReleaseRows();
		m_Rowset.m_hRow = hRow;
		if (FAILED(m_Rowset.GetData()))
			return GX_INVALID;

		ULONG nRow;
		CBookmark<0> bmk;
		bmk.SetBookmark(m_nBookmarkSize, m_data.bookmark);
		if (FAILED(m_Rowset.GetApproximatePosition(&bmk, &nRow, NULL) || 
			nRow > m_pDropDownRealWnd->GetRowCount()))
		nRow = m_pDropDownRealWnd->GetRowCount();

		m_nRow = nRow;
	}

	if (m_nRow > 0)
	{
		if (bAutoPosition && m_bListAutoPosition)
			m_pDropDownRealWnd->MoveTo(m_nRow);

		return m_nRow;
	}
	else
		return GX_INVALID;
}

void CKCOMRichComboCtrl::Bind()
{
	m_Rowset.ReleaseRows();

	if (FAILED(m_Rowset.ReleaseAccessors(m_Rowset.m_spRowset)))
		return;

	m_nColumnsBind = 0;
	
	for (int i = 0; i < m_apColumnsProp.GetSize(); i ++)
	{
		if (m_apColumnsProp[i]->nDataField != -1)
			m_nColumnsBind ++;
	}

	int nEntries = m_nColumnsBind + (m_bHaveBookmarks ? 1 : 0);
	m_Rowset.CreateAccessor(nEntries, &m_data, sizeof(m_data));
	if (m_bHaveBookmarks)
		m_Rowset.AddBindEntry(0, DBTYPE_BYTES, sizeof(m_data.bookmark), &m_data.bookmark);

	for (i = 0; i < m_apColumns.GetSize(); i ++)
	{
		if (m_apColumns[i]->prop.nDataField != -1)
			m_Rowset.AddBindEntry(m_apColumns[i]->prop.nDataField, DBTYPE_STR, 255 * sizeof(TCHAR), 
				&(m_data.strData[i]), NULL, &(m_data.dwStatus[i]));
	}

	HRESULT hr = m_Rowset.Bind();
	if (FAILED(hr))
	{
		m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
		m_pDropDownRealWnd->LockUpdate(FALSE);
		m_pDropDownRealWnd->Redraw();
	}
}

void CKCOMRichComboCtrl::ShowDropDownWnd()
{
	FireDropDown();

	if(m_bDataSrcChanged)
		UpdateDropDownWnd();
	
	m_pDropDownRealWnd->SetScrollBarMode(SB_BOTH, gxnDisabled);

	CPoint pt;
	CRect rect;
	GetClientRect(&rect);
	pt.x = rect.left;
	pt.y = rect.bottom;
	ClientToScreen(&pt);
	int cx = pt.x;
	int cy = pt.y;

	int x = 0, y = 0, yWndMin = 0, yWndMax = 0, xTotal = 2, yTotal = 0;
	
	// can not exceed the border of screen
	int nMaxWidth = GetSystemMetrics(SM_CXSCREEN) - cx;
	int nMaxHeight = GetSystemMetrics(SM_CYSCREEN) - cy;

	CRect rcBounds(cx, cy, 0, 0);

	if (m_bColumnHeaders)
	{
		yWndMin += m_nHeaderHeight;
		yWndMax += m_nHeaderHeight;
		yTotal += m_nHeaderHeight;
	}

	// 4 more pixels to avoid the strenge action of the grid
	yTotal += m_nRowHeight * m_pDropDownRealWnd->GetRowCount() + 4;
	yWndMax += m_nRowHeight * m_nMaxDropDownItems + 4;
	yWndMin += m_nRowHeight * m_nMinDropDownItems + 4;

	for (ROWCOL i = 1; i <= m_pDropDownRealWnd->GetColCount(); i ++)
		xTotal += m_pDropDownRealWnd->GetColWidth(i);

	y = min(yTotal, yWndMax);
	y = max(y, yWndMin);
	y = min(y, nMaxHeight);
	if (y < yTotal)
	{
		xTotal += GetSystemMetrics(SM_CXVSCROLL);
		m_pDropDownRealWnd->SetScrollBarMode(SB_VERT, gxnEnabled | gxnEnhanced);
	}

	x = (m_bListWidthAutoSize || m_nListWidth == 0) ? xTotal : m_nListWidth;
	x = min(x, nMaxWidth);
	if (x < xTotal)
	{
		y += GetSystemMetrics(SM_CYHSCROLL);
		m_pDropDownRealWnd->SetScrollBarMode(SB_HORZ, gxnEnabled | gxnEnhanced);
	}
	y = min(y, nMaxHeight);

	rcBounds.right = rcBounds.left + x;
	rcBounds.bottom = rcBounds.top + y;
	m_pDropDownRealWnd->MoveWindow(&rcBounds, FALSE);

	CString strText;
	m_pEdit->GetWindowText(strText);
	SearchText(strText);
	if (m_nRow < 1)
	{
		m_nRow = 1;
		m_pDropDownRealWnd->MoveTo(m_nRow);
	}

	m_pDropDownRealWnd->SetWindowPos(&wndTopMost, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
	m_pDropDownRealWnd->SetForegroundWindow();
	m_pDropDownRealWnd->ShowWindow(SW_SHOW);
	
	HookMouse(m_pDropDownRealWnd->m_hWnd);
}

void CKCOMRichComboCtrl::OnDropDown()
{
	if (!m_pDropDownRealWnd)
		return;

	if (m_pDropDownRealWnd->IsWindowVisible())
		HideDropDownWnd();
	else
		ShowDropDownWnd();
}

void CKCOMRichComboCtrl::HookMouse(HWND hWnd)
{
	SetHost(hWnd, m_pDropDownRealWnd->m_hWnd);
	// Install mouse hook
	m_hSystemHook = SetWindowsHookEx(WH_MOUSE, MouseProc, (HINSTANCE)GetCurrentThread(), 0);
	
	if (!m_hSystemHook)
	{
#ifdef _DEBUG
		AfxMessageBox("Failed to install hook");
#endif
	}
}

BSTR CKCOMRichComboCtrl::GetText() 
{
	CString strResult;
	if (m_pEdit)
		m_pEdit->GetWindowText(m_strText);

	strResult = m_strText;

	return strResult.AllocSysString();
}

void CKCOMRichComboCtrl::SetText(LPCTSTR lpszNewValue) 
{
	if (m_pEdit && IsValidText(lpszNewValue))
	{
		if (!AmbientUserMode())
		{
			m_strText = lpszNewValue;
			InvalidateControl();
		}

		int nStart, nEnd;
		m_pEdit->GetSel(nStart, nEnd);
		m_pEdit->SetWindowText(lpszNewValue);
		m_pEdit->SetSel(nEnd, nEnd);
	}
}

OLE_COLOR CKCOMRichComboCtrl::GetBackColor() 
{
	return m_clrBack;
}

void CKCOMRichComboCtrl::SetBackColor(OLE_COLOR nNewValue) 
{
	m_clrBack = nNewValue;

	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->ChangeStandardStyle(CGXStyle().SetInterior(TranslateColor(GetBackColor())));
	
	SetModifiedFlag();
	BoundPropertyChanged(dispidBackColor);
	InvalidateControl();
}

HBRUSH CKCOMRichComboCtrl::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	if (m_pEdit && pWnd->m_hWnd == m_pEdit->m_hWnd && nCtlColor == CTLCOLOR_EDIT)
	{
		if (m_pBrhBack)
		{
			delete m_pBrhBack;
			m_pBrhBack = NULL;
		}

		m_pBrhBack = new CBrush(TranslateColor(m_clrBack));
		pDC->SetTextColor(TranslateColor(GetForeColor()));
		pDC->SetBkColor(TranslateColor(m_clrBack));

		return (HBRUSH)*m_pBrhBack;
	}

	HBRUSH hbr = COleControl::OnCtlColor(pDC, pWnd, nCtlColor);

	return hbr;
}

void CKCOMRichComboCtrl::AddItem(LPCTSTR Item, const VARIANT FAR& RowIndex) 
{
	if (m_nDataMode == 0)
		ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	ROWCOL nRow = GetRowColFromVariant(RowIndex);
	if (nRow == GX_INVALID)
		nRow = m_pDropDownRealWnd->GetRowCount();

	if (nRow == 0 || nRow > (ROWCOL)m_pDropDownRealWnd->OnGetRecordCount() + 1)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	CString strText = Item, strField;
	int i = 1, j = 0;

	m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
	m_pDropDownRealWnd->InsertRows(nRow, 1);
	strField.Empty();
	CGXData * pData = m_pDropDownRealWnd->GetParam()->GetData();

	while (j < strText.GetLength() && (ROWCOL)i <= m_pDropDownRealWnd->GetColCount())
	{
		if (strText[j] != m_strSeparator)
			strField += strText[j];
		else
		{
			pData->StoreValueRowCol(nRow, i, strField, gxOverride);
			strField.Empty();
			i ++;
		}

		j ++;
	}

	if ((ROWCOL)i <= m_pDropDownRealWnd->GetColCount())
		pData->StoreValueRowCol(nRow, i, strField, gxOverride);

	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
	m_pDropDownRealWnd->RedrawRowCol(CGXRange().SetRows(nRow));
	m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
}

void CKCOMRichComboCtrl::RemoveItem(long RowIndex) 
{
	if (m_nDataMode == 0)
		ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	ROWCOL nRow = RowIndex;
	if (nRow == 0 || nRow > m_pDropDownRealWnd->GetRowCount())
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
	m_pDropDownRealWnd->DeleteRows(nRow, nRow);

	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
	m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
}

void CKCOMRichComboCtrl::Scroll(long Rows, long Cols) 
{
	m_pDropDownRealWnd->DoScroll(Rows > 0 ? GX_DOWN : GX_UP, abs(Rows));
	m_pDropDownRealWnd->DoScroll(Cols > 0 ? GX_RIGHT : GX_LEFT, abs(Cols));
}

void CKCOMRichComboCtrl::RemoveAll() 
{
	if (m_nDataMode == 0)
		ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	m_pDropDownRealWnd->GetParam()->EnableUndo(FALSE);
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(FALSE);
	m_pDropDownRealWnd->SetRowCount(0);
//	m_pDropDownRealWnd->SetColCount(0);
	m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
	m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);

//	for (int i = 0; i < m_apColumns.GetSize(); i ++)
//		delete m_apColumns[i];

//	m_apColumns.RemoveAll();
}

void CKCOMRichComboCtrl::MoveFirst() 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	m_pDropDownRealWnd->MoveTo(0);
}

void CKCOMRichComboCtrl::MovePrevious() 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	ROWCOL nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
	m_pDropDownRealWnd->MoveTo(nRow - 1);
}

void CKCOMRichComboCtrl::MoveNext() 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	ROWCOL nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
	m_pDropDownRealWnd->MoveTo(nRow + 1);
}

void CKCOMRichComboCtrl::MoveLast() 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	m_pDropDownRealWnd->MoveTo(m_pDropDownRealWnd->GetRowCount());
}

void CKCOMRichComboCtrl::MoveRecords(long Rows) 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	ROWCOL nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
	m_pDropDownRealWnd->MoveTo(nRow + Rows);
}

void CKCOMRichComboCtrl::ReBind() 
{
	m_bDataSrcChanged = TRUE;
}

BOOL CKCOMRichComboCtrl::IsItemInList() 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	return SearchText(GetText(), FALSE) != GX_INVALID;
}

VARIANT CKCOMRichComboCtrl::RowBookmark(long RowIndex) 
{
	VARIANT vaResult;
	VariantInit(&vaResult);

	GetBmkOfRow(RowIndex, &vaResult);

	return vaResult;
}

VARIANT CKCOMRichComboCtrl::GetBookmark() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	ROWCOL nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);

	VARIANT sa;
	VariantInit(&sa);
	GetBmkOfRow(nRow, &sa);

	return sa;
}

void CKCOMRichComboCtrl::SetBookmark(const VARIANT FAR& newValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	SetRow(GetRowFromBmk(&newValue));
}

ROWCOL CKCOMRichComboCtrl::GetRowFromIndex(ROWCOL nIndex)
{
	for (ROWCOL nRow = 1; nRow <= m_pDropDownRealWnd->GetRowCount(); nRow ++)
		if (m_pDropDownRealWnd->GetRowIndex(nRow) == nIndex)
			return nRow;

	return GX_INVALID;
}

void CKCOMRichComboCtrl::GetBmkOfRow(ROWCOL nRow, VARIANT *pVaRet)
{
	if (m_nDataMode)
	{
		pVaRet->vt = VT_I4;
		pVaRet->lVal = m_pDropDownRealWnd->GetRowIndex(nRow);

		return;
	}
	
	if (nRow > 0 && nRow <= (ROWCOL)m_apBookmark.GetSize())
	{
		GetBmkOfRow(nRow);
		
		pVaRet->vt = VT_ARRAY | VT_UI1;
		SAFEARRAYBOUND rgsabound;
		rgsabound.cElements = m_nBookmarkSize;
		rgsabound.lLbound = 0;
		
		SAFEARRAY * parray = SafeArrayCreate(VT_UI1, 1, &rgsabound);
		if (parray == NULL)
			AfxThrowMemoryException();
		
		void* pvDestData;
		AfxCheckError(SafeArrayAccessData(parray, &pvDestData));
		memcpy(pvDestData, m_apBookmark[nRow - 1], m_nBookmarkSize);
		AfxCheckError(SafeArrayUnaccessData(parray));
		
		pVaRet->parray = parray;
	}
}

ROWCOL CKCOMRichComboCtrl::GetRowFromBmk(const VARIANT *pBmk)
{
	if (m_nDataMode)
	{
		COleVariant va = *pBmk;
		va.ChangeType(VT_I4);

		return GetRowFromIndex(va.lVal);
	}

	if (pBmk->vt != (VT_UI1 | VT_ARRAY) || !pBmk->parray->pvData)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDBOOKMARK);

	LONG nLBound = 0, nUBound = 0;
	AfxCheckError(SafeArrayGetLBound(pBmk->parray, 1, &nLBound));
	AfxCheckError(SafeArrayGetUBound(pBmk->parray, 1, &nUBound));

	if (nUBound - nLBound + 1 != m_nBookmarkSize)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDBOOKMARK);

	BYTE * bookmark = NULL;

	AfxCheckError(SafeArrayAccessData(pBmk->parray, (void **)&bookmark));
	ROWCOL nRow = GetRowFromBmk(&bookmark[nLBound]);
	AfxCheckError(SafeArrayUnaccessData(pBmk->parray));

	return nRow;
}

long CKCOMRichComboCtrl::GetStyle() 
{
	return m_nStyle;
}

void CKCOMRichComboCtrl::SetStyle(long nNewValue) 
{
	m_nStyle = nNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidStyle);
}

int CKCOMRichComboCtrl::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message) 
{
	if (AmbientUserMode())
		OnActivateInPlace (TRUE, NULL);
	
	return COleControl::OnMouseActivate(pDesktopWnd, nHitTest, message);
}
