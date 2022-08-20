// KCOMRichGridCtl.cpp : Implementation of the CKCOMRichGridCtrl ActiveX Control class.

#include "stdafx.h"
#include "KCOMData.h"
#include "KCOMRichGridCtl.h"
#include "KCOMRichGridPpg.h"
#include "KCOMRichGridColumnPpg.h"
#include "GridInner.h"
#include <oledberr.h>
#include <msstkppg.h>
#include "GridColumn.h"
#include "GridColumns.h"
#include "KCOMEditControl.h"
#include "KCOMComboBox.h"
#include "KCOMDateControl.h"
#include "KCOMMaskControl.h"
#include "KCOMEditBtn.h"
#include "SelBookmarks.h"
//#include "KCOMDrawingTools.h"
#include "oleimpl2.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

IMPLEMENT_DYNCREATE(CKCOMRichGridCtrl, COleControl)

BEGIN_INTERFACE_MAP(CKCOMRichGridCtrl, COleControl)
	// Map the interfaces to their implementations
	INTERFACE_PART(CKCOMRichGridCtrl, IID_IRowPositionChange, RowPositionChange)
	INTERFACE_PART(CKCOMRichGridCtrl, IID_IRowsetNotify, RowsetNotify)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) CKCOMRichGridCtrl::XRowPositionChange::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowPositionChange)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMRichGridCtrl::XRowPositionChange::Release(void)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowPositionChange)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMRichGridCtrl::XRowPositionChange::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowPositionChange)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMRichGridCtrl::XRowPositionChange::OnRowPositionChange(
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowPositionChange)

	if (ePhase == DBEVENTPHASE_DIDEVENT)
	{
		DBPOSITIONFLAGS dwPositionFlags;
		HRESULT hr;
		HCHAPTER hChapter = DB_NULL_HCHAPTER;
		
		hr = pThis->m_spRowPosition->GetRowPosition(&hChapter, &pThis->m_Rowset.m_hRow, &dwPositionFlags);
		if (FAILED(hr) || pThis->m_Rowset.m_hRow == DB_NULL_HROW)
		{
			return S_OK;
		}
		
		ROWCOL i = pThis->GetRowFromHRow(pThis->m_Rowset.m_hRow);
		ROWCOL nRow, nCol;
		if (i != GX_INVALID)
		{
			pThis->m_pGridInner->GetCurrentCell(nRow, nCol);
			if (i != nRow)
			{
				pThis->m_pGridInner->SetCurrentCell(i, nCol);
			}
		}

		pThis->m_Rowset.FreeRecordMemory();
		pThis->m_Rowset.ReleaseRows();
		
		return S_OK;
	}
	
	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP_(ULONG) CKCOMRichGridCtrl::XRowsetNotify::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowsetNotify)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMRichGridCtrl::XRowsetNotify::Release(void)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowsetNotify)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMRichGridCtrl::XRowsetNotify::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowsetNotify)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMRichGridCtrl::XRowsetNotify::OnFieldChange(IRowset * prowset,
	HROW hRow, ULONG cColumns, ULONG rgColumns[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase)// && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_COLUMN_SET:
		case DBREASON_COLUMN_RECALCULATED:
		{
			ROWCOL nRow;
			nRow = pThis->GetRowFromHRow(hRow);
			if (nRow == GX_INVALID || nRow > (ROWCOL)pThis->m_pGridInner->GetRecordCount())
			{
				pThis->m_bEndReached = FALSE;
				pThis->FetchNewRows();
				return S_OK;
			}

			for (ULONG i = 0; i < cColumns; i ++)
			{
				pThis->UpdateCellFromSrc(nRow, rgColumns[i]);
			}
		}

		return S_OK;
		break;

		default:
			return DB_S_UNWANTEDREASON;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP CKCOMRichGridCtrl::XRowsetNotify::OnRowChange(IRowset * prowset,
	ULONG cRows, const HROW rghRows[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowsetNotify)
	
	if (DBEVENTPHASE_ABOUTTODO == ePhase)// && pThis->m_Rowset.m_spRowset == prowset)
	{	
		switch (eReason)
		{
		case DBREASON_ROW_DELETE:
		case DBREASON_ROW_UNDOINSERT:
		{
			ROWCOL nRow;

			for (ULONG i = 0; i < cRows; i ++)
			{
				nRow = pThis->GetRowFromHRow(rghRows[i]);
				if (nRow != GX_INVALID)
				{
					int j;
					for (j = 0; j < pThis->m_arRowToDelete.GetSize() && pThis->
						m_arRowToDelete[i] > nRow; j ++);

					pThis->m_arHRowToDelete.InsertAt(j, rghRows[i]);
					pThis->m_arRowToDelete.InsertAt(j, nRow);
				}
			}
		}

		default:
			return S_OK;
		}
	}

	if (DBEVENTPHASE_DIDEVENT == ePhase)// && pThis->m_Rowset.m_spRowset == prowset)
	{
		switch (eReason)
		{
		case DBREASON_ROW_DELETE:
		case DBREASON_ROW_UNDOINSERT:
		{
			for (ULONG i = 0; i < cRows; i ++)
			{
				for (int j = 0; j < pThis->m_arHRowToDelete.GetSize(); j ++)
				{
					if (pThis->m_arHRowToDelete[j] == rghRows[i])
						pThis->DeleteRowFromSrc(pThis->m_arRowToDelete[j]);
				}
			}
		}

		return S_OK;
		break;

		case DBREASON_ROW_ASYNCHINSERT:
		case DBREASON_ROW_INSERT:
		{
		pThis->m_bEndReached = FALSE;
			pThis->FetchNewRows();
		}

		return S_OK;
		break;

		case DBREASON_ROW_UNDODELETE:
		{
			for (int j = 0; j < pThis->m_arHRowToDelete.GetSize(); j ++)
			{
				for (ULONG i = 0; i < cRows; i ++)
				{
					if (pThis->m_arHRowToDelete[j] == rghRows[i])
						pThis->UndeleteRecordFromHRow(rghRows[i], pThis->m_arRowToDelete[j]);
				}
			}
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
					pThis->UpdateRowFromSrc(nRow);
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

STDMETHODIMP CKCOMRichGridCtrl::XRowsetNotify::OnRowsetChange(IRowset * prowset,
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase)// && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_ROWSET_RELEASE:
		case DBREASON_ROWSET_CHANGED:
		{
			pThis->m_bDataSrcChanged = TRUE;
			pThis->InvalidateControl();
		}

		return S_OK;
		break;
			
		default:
			return DB_S_UNWANTEDREASON;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP_(ULONG) CKCOMRichGridCtrl::XDataSourceListener::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, DataSourceListener)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMRichGridCtrl::XDataSourceListener::Release(void)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, DataSourceListener)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMRichGridCtrl::XDataSourceListener::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, DataSourceListener)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMRichGridCtrl::XDataSourceListener::dataMemberChanged(BSTR bstrDM)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowsetNotify)

	USES_CONVERSION;
	if ((pThis->m_strDataMember.IsEmpty() && (bstrDM == NULL || bstrDM[0] == '\0')) || pThis->m_strDataMember == OLE2T(bstrDM))
	{
		pThis->ClearInterfaces();
		pThis->m_bDataSrcChanged = TRUE;
		pThis->m_bEndReached = FALSE;
		pThis->UpdateControl();
	}
	
	return S_OK;
}

STDMETHODIMP CKCOMRichGridCtrl::XDataSourceListener::dataMemberAdded(BSTR bstrDM)
{
	return S_OK;
}

STDMETHODIMP CKCOMRichGridCtrl::XDataSourceListener::dataMemberRemoved(BSTR bstrDM)
{
	METHOD_PROLOGUE(CKCOMRichGridCtrl, RowsetNotify)

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

BEGIN_MESSAGE_MAP(CKCOMRichGridCtrl, COleControl)
	//{{AFX_MSG_MAP(CKCOMRichGridCtrl)
	ON_WM_CREATE()
	ON_WM_KILLFOCUS()
	ON_WM_MOUSEACTIVATE()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
	ON_REGISTERED_MESSAGE(CKCOMDataApp::m_nBringMsg, OnBringMsg)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CKCOMRichGridCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CKCOMRichGridCtrl)
	DISP_PROPERTY_NOTIFY(CKCOMRichGridCtrl, "HeadForeColor", m_clrHeadFore, OnHeadForeColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CKCOMRichGridCtrl, "HeadBackColor", m_clrHeadBack, OnHeadBackColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CKCOMRichGridCtrl, "DividerColor", m_clrDivider, OnDividerColorChanged, VT_COLOR)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "CaptionAlignment", GetCaptionAlignment, SetCaptionAlignment, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "DividerType", GetDividerType, SetDividerType, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "AllowAddNew", GetAllowAddNew, SetAllowAddNew, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "AllowDelete", GetAllowDelete, SetAllowDelete, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "AllowUpdate", GetAllowUpdate, SetAllowUpdate, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "AllowRowSizing", GetAllowRowSizing, SetAllowRowSizing, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "RecordSelectors", GetRecordSelectors, SetRecordSelectors, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "AllowColumnSizing", GetAllowColumnSizing, SetAllowColumnSizing, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "ColumnHeaders", GetColumnHeaders, SetColumnHeaders, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "HeadFont", GetHeadFont, SetHeadFont, VT_FONT)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "Font", GetFont, SetFont, VT_FONT)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "Col", GetCol, SetCol, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "Row", GetRow, SetRow, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "DataMode", GetDataMode, SetDataMode, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "VisibleCols", GetVisibleCols, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "VisibleRows", GetVisibleRows, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "LeftCol", GetLeftCol, SetLeftCol, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "Cols", GetCols, SetCols, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "Rows", GetRows, SetRows, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "FieldSeparator", GetFieldSeparator, SetFieldSeparator, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "DividerStyle", GetDividerStyle, SetDividerStyle, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "FirstRow", GetFirstRow, SetFirstRow, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "DefColWidth", GetDefColWidth, SetDefColWidth, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "RowHeight", GetRowHeight, SetRowHeight, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "Redraw", GetRedraw, SetRedraw, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "AllowColumnMoving", GetAllowColumnMoving, SetAllowColumnMoving, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "SelectByCell", GetSelectByCell, SetSelectByCell, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "Columns", GetColumns, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "DataSource", GetDataSource, SetDataSource, VT_UNKNOWN)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "DataMember", GetDataMember, SetDataMember, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "RowChanged", GetRowChanged, SetRowChanged, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "HeaderHeight", GetHeaderHeight, SetHeaderHeight, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "Bookmark", GetBookmark, SetBookmark, VT_VARIANT)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "SelBookmarks", GetSelBookmarks, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CKCOMRichGridCtrl, "FrozenCols", GetFrozenCols, SetFrozenCols, VT_I4)
	DISP_FUNCTION(CKCOMRichGridCtrl, "AddItem", AddItem, VT_EMPTY, VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CKCOMRichGridCtrl, "RemoveItem", RemoveItem, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CKCOMRichGridCtrl, "Scroll", Scroll, VT_EMPTY, VTS_I4 VTS_I4)
	DISP_FUNCTION(CKCOMRichGridCtrl, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichGridCtrl, "MoveFirst", MoveFirst, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichGridCtrl, "MovePrevious", MovePrevious, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichGridCtrl, "MoveNext", MoveNext, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichGridCtrl, "MoveLast", MoveLast, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichGridCtrl, "MoveRecords", MoveRecords, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CKCOMRichGridCtrl, "IsAddRow", IsAddRow, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CKCOMRichGridCtrl, "Update", Update, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichGridCtrl, "CancelUpdate", CancelUpdate, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichGridCtrl, "ReBind", ReBind, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichGridCtrl, "RowBookmark", RowBookmark, VT_VARIANT, VTS_I4)
	DISP_STOCKFUNC_REFRESH()
	DISP_STOCKPROP_FORECOLOR()
	DISP_STOCKPROP_HWND()
	DISP_STOCKPROP_CAPTION()
	DISP_STOCKPROP_BORDERSTYLE()
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CKCOMRichGridCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_STOCK(CKCOMRichGridCtrl, "CtrlType", 255, CKCOMRichGridCtrl::GetCtrlType, COleControl::SetNotSupported, VT_I2)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CKCOMRichGridCtrl, COleControl)
	//{{AFX_EVENT_MAP(CKCOMRichGridCtrl)
	EVENT_CUSTOM("BtnClick", FireBtnClick, VTS_NONE)
	EVENT_CUSTOM("AfterColUpdate", FireAfterColUpdate, VTS_I2)
	EVENT_CUSTOM("BeforeColUpdate", FireBeforeColUpdate, VTS_I2  VTS_BSTR  VTS_PI2)
	EVENT_CUSTOM("BeforeInsert", FireBeforeInsert, VTS_PI2)
	EVENT_CUSTOM("BeforeUpdate", FireBeforeUpdate, VTS_PI2)
	EVENT_CUSTOM("ColResize", FireColResize, VTS_I2  VTS_PI2)
	EVENT_CUSTOM("RowResize", FireRowResize, VTS_PI2)
	EVENT_CUSTOM("RowColChange", FireRowColChange, VTS_NONE)
	EVENT_CUSTOM("Scroll", FireScroll, VTS_PI2)
	EVENT_CUSTOM("Change", FireChange, VTS_NONE)
	EVENT_CUSTOM("AfterDelete", FireAfterDelete, VTS_PI2)
	EVENT_CUSTOM("BeforeDelete", FireBeforeDelete, VTS_PI2  VTS_PI2)
	EVENT_CUSTOM("AfterUpdate", FireAfterUpdate, VTS_PI2)
	EVENT_CUSTOM("AfterInsert", FireAfterInsert, VTS_PI2)
	EVENT_CUSTOM("ColSwap", FireColSwap, VTS_I2  VTS_I2  VTS_I2  VTS_PI2)
	EVENT_CUSTOM("UpdateError", FireUpdateError, VTS_NONE)
	EVENT_CUSTOM("BeforeRowColChange", FireBeforeRowColChange, VTS_PI2)
	EVENT_CUSTOM("ScrollAfter", FireScrollAfter, VTS_NONE)
	EVENT_CUSTOM("BeforeDrawCell", FireBeforeDrawCell, VTS_I4  VTS_I4  VTS_PVARIANT  VTS_PCOLOR  VTS_PI2)
	EVENT_STOCK_KEYDOWN()
	EVENT_STOCK_KEYPRESS()
	EVENT_STOCK_CLICK()
	EVENT_STOCK_DBLCLICK()
	EVENT_STOCK_KEYUP()
	EVENT_STOCK_MOUSEDOWN()
	EVENT_STOCK_MOUSEMOVE()
	EVENT_STOCK_MOUSEUP()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CKCOMRichGridCtrl, 4)
	PROPPAGEID(CKCOMRichGridPropPage::guid)
	PROPPAGEID(CKCOMRichGridPropColumnPage::guid)
	PROPPAGEID(CLSID_StockColorPage)
	PROPPAGEID(CLSID_StockFontPage)
END_PROPPAGEIDS(CKCOMRichGridCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMRichGridCtrl, "KCOMRICHGRID.KCOMRichGridCtrl.1",
	0xcfc2325, 0x12a9, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CKCOMRichGridCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DKCOMRichGrid =
		{ 0xcfc2323, 0x12a9, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_DKCOMRichGridEvents =
		{ 0xcfc2324, 0x12a9, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_IDataSource =
		{ 0x7c0ffab3, 0xcd84, 0x11d0, { 0x94, 0x9a, 0, 0xa0, 0xc9, 0x11, 0x10, 0xed } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwKCOMRichGridOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CKCOMRichGridCtrl, IDS_KCOMRICHGRID, _dwKCOMRichGridOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridCtrl::CKCOMRichGridCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMRichGridCtrl

BOOL CKCOMRichGridCtrl::CKCOMRichGridCtrlFactory::UpdateRegistry(BOOL bRegister)
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
			IDS_KCOMRICHGRID,
			IDB_KCOMRICHGRID,
			afxRegApartmentThreading,
			_dwKCOMRichGridOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridCtrl::CKCOMRichGridCtrl - Constructor

CKCOMRichGridCtrl::CKCOMRichGridCtrl() : m_fntHeader(NULL), m_fntCell(NULL)
{
	InitializeIIDs(&IID_DKCOMRichGrid, &IID_DKCOMRichGridEvents);

	m_nColsPreset = 0;
	m_bIsPosSync = TRUE;
	m_bEndReached = TRUE;
	m_bDataSrcChanged = FALSE;
	m_dwCookieRPC = m_dwCookieRN = 0;
	m_bHaveBookmarks = FALSE;
	m_pColumnInfo = NULL;
	m_pStrings = NULL;
	m_nColumns = m_nColumnsBind = 0;
	m_nBookmarkSize = 0;

	m_spDataSource = NULL;
	m_spRowPosition = NULL;

	m_pGridInner = NULL;
	m_bDataSrcChanged = FALSE;
	m_spDataSource = NULL;
	m_nCol = m_nRow = 0;
	m_bRedraw = TRUE;
	m_nColsUsed = 0;
	m_bDroppedDown = FALSE;
	m_hSystemHook = NULL;
	m_hExternalDropDownWnd = NULL;

//	((CKCOMDrawingTools *)GXDaFTools())->m_pCtrl = this;
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridCtrl::~CKCOMRichGridCtrl - Destructor

CKCOMRichGridCtrl::~CKCOMRichGridCtrl()
{
	if (m_hSystemHook)
	{
		UnhookWindowsHookEx(m_hSystemHook);
		m_hSystemHook = NULL;
	}

	if (m_pGridInner)
	{
		m_pGridInner->DestroyWindow();
		delete m_pGridInner;
		m_pGridInner = NULL;
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
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridCtrl::OnDraw - Drawing function

void CKCOMRichGridCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (m_bDataSrcChanged)
		UpdateControl();

	CRect rtGrid = rcBounds, rtCaption = rcBounds;
	CFont * pFontOld;
	CSize szCaption;

	int nSave = pdc->SaveDC();

	if (InternalGetText().GetLength())
	{
		pFontOld = SelectFontObject(pdc, m_fntHeader);

		rtCaption.bottom = rtCaption.top + m_nHeaderHeight;
		pdc->Draw3dRect(&rtCaption, GetSysColor(COLOR_3DLIGHT), GetSysColor(COLOR_3DDKSHADOW));
		rtCaption.DeflateRect(1, 1);
		pdc->Draw3dRect(&rtCaption, GetSysColor(COLOR_3DHILIGHT), GetSysColor(COLOR_3DSHADOW));
		rtCaption.DeflateRect(1, 1);
		pdc->FillRect(&rtCaption, &CBrush(TranslateColor(m_clrHeadBack)));

		pdc->SetBkColor(TranslateColor(m_clrHeadBack));
		pdc->SetTextColor(TranslateColor(m_clrHeadFore));
		
		switch (m_nCapitonAlignment)
		{
		case 0:
			pdc->SetTextAlign(TA_LEFT);
			pdc->ExtTextOut(3, (m_nHeaderHeight - pdc->GetTextExtent(
				InternalGetText()).cy) / 2, ETO_CLIPPED, &rtCaption, InternalGetText(), 
				NULL);
			break;

		case 1:
			pdc->SetTextAlign(TA_RIGHT);
			pdc->ExtTextOut(rtCaption.Width() - 3, (m_nHeaderHeight - pdc->GetTextExtent(
				InternalGetText()).cy) / 2, ETO_CLIPPED, &rtCaption, InternalGetText(), 
				NULL);
			break;

		case 2:
			pdc->SetTextAlign(TA_CENTER);
			pdc->ExtTextOut(rtCaption.Width() / 2, (m_nHeaderHeight - pdc->
				GetTextExtent(InternalGetText()).cy) / 2, ETO_CLIPPED, &rtCaption, 
				InternalGetText(), NULL);
		}

		pdc->SelectObject(pFontOld);

		rtCaption.InflateRect(2, 2);
		rtGrid.top += m_nHeaderHeight;
	}

	pdc->RestoreDC(nSave);

	if (!m_pGridInner) 
		CreateDrawGrid();		

	if (!AmbientUserMode())
	{
		m_pGridInner->SetGridRect(TRUE, &rtGrid);
		m_pGridInner->OnGridDraw(pdc);
	}
	else
	{
		m_pGridInner->Invalidate();
	}
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridCtrl::DoPropExchange - Persistence support

void CKCOMRichGridCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	ASSERT_POINTER(pPX, CPropExchange);

	ExchangeExtent(pPX);
	ExchangeStockProps(pPX);

	PX_Font(pPX, "Font", m_fntCell, &_KCOMFontDescDefault);
	PX_Font(pPX, "HeaderFont", m_fntHeader, &_KCOMFontDescDefault);
	PX_Long(pPX, "HeaderHeight", m_nHeaderHeight, 25);
	PX_Long(pPX, "RowHeight", m_nRowHeight, 20);
	PX_Long(pPX, "DataMode", m_nDataMode, 0);
	PX_Color(pPX, "HeadForeColor", m_clrHeadFore, RGB(0, 0, 0));
	PX_Color(pPX, "HeadBackColor", m_clrHeadBack, GetSysColor(COLOR_3DFACE));
	PX_Long(pPX, "DefColWidth", m_nDefColWidth, 100);
	PX_Bool(pPX, "AllowAddNew", m_bAllowAddNew, FALSE);
	PX_Bool(pPX, "AllowDelete", m_bAllowDelete, FALSE);
	PX_Bool(pPX, "AllowUpdate", m_bAllowUpdate, FALSE);
	PX_Bool(pPX, "RecordSelectors", m_bRecordSelectors, TRUE);
	PX_Long(pPX, "Rows", m_nRows, 3);
	PX_Long(pPX, "Cols", m_nCols, 3);
	PX_Long(pPX, "CapitonAlignment", m_nCapitonAlignment, 2);
	PX_Long(pPX, "DividerType", m_nDividerType, 3);
	PX_Long(pPX, "DividerStyle", m_nDividerStyle, 0);
	PX_Bool(pPX, "AllowRowSizing", m_bAllowRowSizing, TRUE);
	PX_Bool(pPX, "AllowColumnSizing", m_bAllowColumnSizing, TRUE);
	PX_Bool(pPX, "ColumnHeaders", m_bColumnHeaders, TRUE);
	PX_Bool(pPX, "AllowColumnMoving", m_bAllowColumnMoving, TRUE);
	PX_Bool(pPX, "SelectByCell", m_bSelectByCell, FALSE);
	PX_String(pPX, "FieldSeparator", m_strFieldSeparator, _T("(tab)"));
	PX_Color(pPX, "BackColor", m_clrBack, GetSysColor(COLOR_WINDOW));
	PX_Color(pPX, "DividerColor", m_clrDivider, GetSysColor(COLOR_3DFACE));
	PX_Long(pPX, "FrozenCols", m_nFrozenCols, 0);

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

		InvalidateControl();
	}
 
	GlobalFree(m_hBlob);
	m_hBlob = NULL;
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridCtrl::OnResetState - Reset control to default state

void CKCOMRichGridCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichGridCtrl message handlers

int CKCOMRichGridCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	CRect rt(0, 0, lpCreateStruct->cx, lpCreateStruct->cy);
	if (InternalGetText().GetLength())
		rt.top += m_nHeaderHeight;

	m_pGridInner = new CGridInner(this);
	m_pGridInner->Create(WS_VISIBLE | WS_CHILD, rt, this, 101);
	InitGridHeader();
	
	return 0;
}

long CKCOMRichGridCtrl::GetCaptionAlignment() 
{
	return m_nCapitonAlignment;
}

void CKCOMRichGridCtrl::SetCaptionAlignment(long nNewValue) 
{
	m_nCapitonAlignment = nNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidCaptionAlignment);
	InvalidateControl();
}

long CKCOMRichGridCtrl::GetDividerType() 
{
	return m_nDividerType;
}

void CKCOMRichGridCtrl::SetDividerType(long nNewValue) 
{
	m_nDividerType = nNewValue;

	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerType);
	InvalidateControl();
}

BOOL CKCOMRichGridCtrl::GetAllowAddNew() 
{
	return m_bAllowAddNew;
}

void CKCOMRichGridCtrl::SetAllowAddNew(BOOL bNewValue) 
{
	if (!m_bAllowUpdate && bNewValue)
		return;

	m_bAllowAddNew = bNewValue;
	if (m_pGridInner)
		m_pGridInner->SetCanAppend(m_bAllowAddNew);

	InvalidateControl();
	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowAddNew);
}

BOOL CKCOMRichGridCtrl::GetAllowDelete() 
{
	return m_bAllowDelete;
}

void CKCOMRichGridCtrl::SetAllowDelete(BOOL bNewValue) 
{
	if (!m_bAllowUpdate && bNewValue)
		return;

	m_bAllowDelete = bNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowDelete);
}

BOOL CKCOMRichGridCtrl::GetAllowUpdate() 
{
	return m_bAllowUpdate;
}

void CKCOMRichGridCtrl::SetAllowUpdate(BOOL bNewValue) 
{
	m_bAllowUpdate = bNewValue;
	if (!m_bAllowUpdate)
	{
		SetAllowDelete(FALSE);
		SetAllowAddNew(FALSE);
	}
	if (m_pGridInner && AmbientUserMode())
		m_pGridInner->SetReadOnly(!m_bAllowUpdate);

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowUpdate);
}

BOOL CKCOMRichGridCtrl::GetAllowRowSizing() 
{
	return m_bAllowRowSizing;
}

void CKCOMRichGridCtrl::SetAllowRowSizing(BOOL nNewValue) 
{
	m_bAllowRowSizing = nNewValue;

	if (AmbientUserMode())
		m_pGridInner->GetParam()->EnableTrackRowHeight(m_bAllowRowSizing);

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowRowSizing);
}

BOOL CKCOMRichGridCtrl::GetRecordSelectors() 
{
	return m_bRecordSelectors;
}

void CKCOMRichGridCtrl::SetRecordSelectors(BOOL bNewValue) 
{
	m_bRecordSelectors = bNewValue;
	if (m_pGridInner)
		m_pGridInner->HideCols(0, 0, !m_bRecordSelectors);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRecordSelectors);
	InvalidateControl();
}

BOOL CKCOMRichGridCtrl::GetAllowColumnSizing() 
{
	return m_bAllowColumnSizing;
}

void CKCOMRichGridCtrl::SetAllowColumnSizing(BOOL bNewValue) 
{
	m_bAllowColumnSizing = bNewValue;

	if (AmbientUserMode())
		m_pGridInner->GetParam()->EnableTrackColWidth(m_bAllowColumnSizing ? 
			GX_TRACK_INDIVIDUAL : FALSE);

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowColumnSizing);
}

BOOL CKCOMRichGridCtrl::GetColumnHeaders() 
{
	return m_bColumnHeaders;
}

void CKCOMRichGridCtrl::SetColumnHeaders(BOOL bNewValue) 
{
	m_bColumnHeaders = bNewValue;

	if (m_pGridInner)
		m_pGridInner->HideRows(0, 0, !m_bColumnHeaders);

	SetModifiedFlag();
	BoundPropertyChanged(dispidColumnHeaders);
	InvalidateControl();
}

LPFONTDISP CKCOMRichGridCtrl::GetHeadFont() 
{
	return m_fntHeader.GetFontDispatch();
}

void CKCOMRichGridCtrl::SetHeadFont(LPFONTDISP newValue) 
{
	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	m_fntHeader.InitializeFont(NULL, newValue);
	InitHeaderFont();
	if (!AmbientUserMode() && m_pGridInner)
		SetHeaderHeight(m_pGridInner->GetRowHeight(0));

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadFont);
	InvalidateControl();
}

LPFONTDISP CKCOMRichGridCtrl::GetFont() 
{
	return m_fntCell.GetFontDispatch();
}

void CKCOMRichGridCtrl::SetFont(LPFONTDISP newValue) 
{
	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	m_fntCell.InitializeFont(NULL, newValue);
	InitCellFont();
	if (!AmbientUserMode() && m_pGridInner)
		SetRowHeight(m_pGridInner->GetFontHeight() + 4);

	SetModifiedFlag();
	BoundPropertyChanged(dispidFont);
	InvalidateControl();
}

long CKCOMRichGridCtrl::GetCol() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	ROWCOL nRow;
	m_pGridInner->GetCurrentCell(nRow, m_nCol);
	return m_nCol;
}

void CKCOMRichGridCtrl::SetCol(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_nCol = nNewValue;

	if (m_pGridInner && m_pGridInner->GetColCount() > (ROWCOL)m_nCol)
	{
		ROWCOL nRow, nCol;
		m_pGridInner->GetCurrentCell(nRow, nCol);
		m_pGridInner->SetCurrentCell(nRow, m_nCol);
	}
	else
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidCol);
}

long CKCOMRichGridCtrl::GetRow() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	ROWCOL nCol;
	m_pGridInner->GetCurrentCell(m_nRow, nCol);
	return m_nRow;
}

void CKCOMRichGridCtrl::SetRow(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_nRow = nNewValue;

	if (m_pGridInner && m_pGridInner->GetRowCount() > (ROWCOL)m_nRow)
	{
		ROWCOL nRow, nCol;
		m_pGridInner->GetCurrentCell(nRow, nCol);
		m_pGridInner->SetCurrentCell(m_nRow, nCol);
	}
	else
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRow);
}

long CKCOMRichGridCtrl::GetDataMode() 
{
	return m_nDataMode;
}

void CKCOMRichGridCtrl::SetDataMode(long nNewValue) 
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

long CKCOMRichGridCtrl::GetVisibleCols() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pGridInner->CalcRightColFromRect(m_pGridInner->GetGridRect()) - 
		m_pGridInner->GetLeftCol() + 1;
}

long CKCOMRichGridCtrl::GetVisibleRows() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pGridInner->CalcBottomRowFromRect(m_pGridInner->GetGridRect()) - 
		m_pGridInner->GetTopRow() + 1;
}

long CKCOMRichGridCtrl::GetLeftCol() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pGridInner->GetLeftCol();
}

void CKCOMRichGridCtrl::SetLeftCol(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_pGridInner->SetLeftCol(nNewValue);
}

long CKCOMRichGridCtrl::GetCols() 
{
	if (!m_pGridInner)
		return m_nCols;

	m_nCols = m_pGridInner->GetColCount();

	return m_nCols;
}

void CKCOMRichGridCtrl::SetCols(long nNewValue) 
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
	
	if (m_pGridInner)
		m_pGridInner->SetColCount(m_nCols);

	SetModifiedFlag();
	BoundPropertyChanged(dispidCols);
	InvalidateControl();
}

long CKCOMRichGridCtrl::GetRows() 
{
	if (!m_pGridInner)
		return m_nRows;

	m_nRows = m_pGridInner->GetRowCount();

	return m_nRows;
}

void CKCOMRichGridCtrl::SetRows(long nNewValue) 
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

	m_pGridInner->SetRowCount(m_nRows);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRows);
	InvalidateControl();
}

BSTR CKCOMRichGridCtrl::GetFieldSeparator() 
{
	return m_strFieldSeparator.AllocSysString();
}

void CKCOMRichGridCtrl::SetFieldSeparator(LPCTSTR lpszNewValue) 
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

long CKCOMRichGridCtrl::GetDividerStyle() 
{
	return m_nDividerStyle;
}

void CKCOMRichGridCtrl::SetDividerStyle(long nNewValue) 
{
	m_nDividerStyle = nNewValue;

	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerStyle);
	InvalidateControl();
}

long CKCOMRichGridCtrl::GetFirstRow() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pGridInner->GetTopRow();
}

void CKCOMRichGridCtrl::SetFirstRow(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_pGridInner->SetTopRow(nNewValue);
}

long CKCOMRichGridCtrl::GetDefColWidth() 
{
	return m_nDefColWidth;
}

void CKCOMRichGridCtrl::SetDefColWidth(long nNewValue) 
{
	if (nNewValue < 0 || nNewValue > 15000)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	m_nDefColWidth = nNewValue;
	if (m_pGridInner)
	{
		m_pGridInner->GetParam()->EnableUndo(FALSE);
		m_pGridInner->SetDefaultColWidth(m_nDefColWidth);
		m_pGridInner->GetParam()->EnableUndo(TRUE);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidDefColWidth);
	InvalidateControl();
}

long CKCOMRichGridCtrl::GetRowHeight() 
{
	if (m_pGridInner)
		m_nRowHeight = m_pGridInner->GetDefaultRowHeight();

	return m_nRowHeight;
}

void CKCOMRichGridCtrl::SetRowHeight(long nNewValue) 
{
	if (nNewValue < 0 || nNewValue > 15000)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	m_nRowHeight = nNewValue;

	if (m_pGridInner)
	{
		m_pGridInner->GetParam()->EnableUndo(FALSE);
		m_pGridInner->SetDefaultRowHeight(m_nRowHeight);
		m_pGridInner->GetParam()->EnableUndo(TRUE);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidRowHeight);
	InvalidateControl();
}

BOOL CKCOMRichGridCtrl::GetRedraw() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_bRedraw;
}

void CKCOMRichGridCtrl::SetRedraw(BOOL bNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_bRedraw = bNewValue;
	m_pGridInner->LockUpdate(!m_bRedraw);
}

BOOL CKCOMRichGridCtrl::GetAllowColumnMoving() 
{
	return m_bAllowColumnMoving;
}

void CKCOMRichGridCtrl::SetAllowColumnMoving(BOOL bNewValue) 
{
	m_bAllowColumnMoving = bNewValue;

	if (AmbientUserMode())
		m_pGridInner->GetParam()->EnableMoveCols(m_bAllowColumnMoving);

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowColumnMoving);
}

BOOL CKCOMRichGridCtrl::GetSelectByCell() 
{
	return m_bSelectByCell;
}

void CKCOMRichGridCtrl::SetSelectByCell(BOOL bNewValue) 
{
	m_bSelectByCell = bNewValue;

	if (AmbientUserMode())
	{
		if (m_bSelectByCell)
		{
			m_pGridInner->GetParam()->SetSpecialMode(GX_MODELBOX_MS);
			m_pGridInner->GetParam()->SetActivateCellFlags(FALSE);
			m_pGridInner->GetParam()->EnableMoveCols(FALSE);
			m_pGridInner->GetParam()->SetHideCurrentCell(GX_HIDE_ALLWAYS);
		}
		else
		{
			m_pGridInner->GetParam()->SetSpecialMode(0);
			m_pGridInner->GetParam()->SetActivateCellFlags(TRUE);
			m_pGridInner->GetParam()->EnableMoveCols(TRUE);
			m_pGridInner->GetParam()->SetHideCurrentCell(GX_HIDE_INACTIVEFRAME);
		}
	}
	
	SetModifiedFlag();
	BoundPropertyChanged(dispidSelectByCell);
}

LPDISPATCH CKCOMRichGridCtrl::GetColumns() 
{
	CGridColumns * pColumns = new CGridColumns();
	pColumns->SetCtrlPtr(this);

	return pColumns->GetIDispatch(FALSE);
}

LPUNKNOWN CKCOMRichGridCtrl::GetDataSource() 
{
	if (m_spDataSource)
		m_spDataSource->AddRef();

	return m_spDataSource;
}

void CKCOMRichGridCtrl::SetDataSource(LPUNKNOWN newValue) 
{
	if (m_nDataMode != 0 && newValue)
	{
		ThrowError(CTL_E_SETNOTSUPPORTED, _T("Can not set datasource in manual mode"));
		goto exit;
	}

	if (newValue == m_spDataSource)
		return;

	ClearInterfaces();
	if (m_spDataSource)
	{
		m_spDataSource->Release();
		m_spDataSource = NULL;
	}

	if (AmbientUserMode())
	{
		m_bDataSrcChanged = TRUE;
		m_bEndReached = FALSE;
	}

	if (!newValue)
	{
		m_spDataSource = NULL;
		goto exit;
	}

	if (FAILED(newValue->QueryInterface(IID_IDataSource, (void **)&m_spDataSource)))
		goto exit;

	if (!AmbientUserMode())
		GetColumnInfo();

	SetModifiedFlag();

exit:
	UpdateControl();
}

BSTR CKCOMRichGridCtrl::GetDataMember() 
{
	return m_strDataMember.AllocSysString();
}

void CKCOMRichGridCtrl::SetDataMember(LPCTSTR lpszNewValue) 
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

	InvalidateControl();
}

void CKCOMRichGridCtrl::AddItem(LPCTSTR Item, const VARIANT FAR& RowIndex) 
{
	if (m_nDataMode == 0)
		ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	ROWCOL nRow = GetRowColFromVariant(RowIndex);
	if (nRow == GX_INVALID)
		nRow = m_pGridInner->OnGetRecordCount() + 1;

	if (nRow == 0 || nRow > (ROWCOL)m_pGridInner->GetRecordCount() + 1)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	CString strText = Item, strField;
	int i = 1, j = 0;

	m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
	m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	m_pGridInner->InsertRows(nRow, 1);
	strField.Empty();
	CGXData * pData = m_pGridInner->GetParam()->GetData();

	while (j < strText.GetLength() && (ROWCOL)i <= m_pGridInner->GetColCount())
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

	if ((ROWCOL)i <= m_pGridInner->GetColCount())
		pData->StoreValueRowCol(nRow, i, strField, gxOverride);

	m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pGridInner->RedrawRowCol(CGXRange().SetRows(nRow));
	m_pGridInner->GetParam()->EnableUndo(TRUE);
}

void CKCOMRichGridCtrl::RemoveItem(long RowIndex) 
{
	if (m_nDataMode == 0)
		ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	ROWCOL nRow = RowIndex;
	if (nRow == 0 || nRow > m_pGridInner->GetRowCount())
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
	m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	m_pGridInner->DeleteRows(nRow, nRow);

	ROWCOL nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);
	if (nRow > m_pGridInner->GetRecordCount())
		nRow = m_pGridInner->GetRecordCount();
	if (nCol > m_pGridInner->GetColCount())
		nCol = m_pGridInner->GetColCount();
	m_pGridInner->SetCurrentCell(nRow, nCol);

	m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pGridInner->GetParam()->EnableUndo(TRUE);
}

void CKCOMRichGridCtrl::Scroll(long Rows, long Cols) 
{
	m_pGridInner->DoScroll(Rows > 0 ? GX_DOWN : GX_UP, abs(Rows));
	m_pGridInner->DoScroll(Cols > 0 ? GX_RIGHT : GX_LEFT, abs(Cols));
}

void CKCOMRichGridCtrl::RemoveAll() 
{
	if (m_nDataMode == 0)
		ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
	m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	m_pGridInner->SetRowCount(0);
	m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pGridInner->GetParam()->EnableUndo(TRUE);
}

void CKCOMRichGridCtrl::MoveFirst() 
{
	m_pGridInner->Moveto(1);
}

void CKCOMRichGridCtrl::MovePrevious() 
{
	ROWCOL nRow, nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);
	m_pGridInner->Moveto(nRow - 1);
}

void CKCOMRichGridCtrl::MoveNext() 
{
	ROWCOL nRow, nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);
	m_pGridInner->Moveto(nRow + 1);
}

void CKCOMRichGridCtrl::MoveLast() 
{
	m_pGridInner->Moveto(m_pGridInner->GetRowCount());
}

void CKCOMRichGridCtrl::MoveRecords(long Rows) 
{
	ROWCOL nRow, nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);
	m_pGridInner->Moveto(nRow + Rows);
}

BOOL CKCOMRichGridCtrl::IsAddRow() 
{
	ROWCOL nRow, nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);

	return m_bAllowAddNew && m_pGridInner->GetAppendRow() == nRow;
}

void CKCOMRichGridCtrl::Update() 
{
	m_pGridInner->Update();
}

void CKCOMRichGridCtrl::CancelUpdate() 
{
	if (m_pGridInner->IsRecordDirty())
		m_pGridInner->CancelRecord();
}

void CKCOMRichGridCtrl::ReBind() 
{
	m_bDataSrcChanged = TRUE;

	InvalidateControl();
}

BOOL CKCOMRichGridCtrl::GetRowChanged() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pGridInner->IsRecordDirty();
}

void CKCOMRichGridCtrl::SetRowChanged(BOOL bNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	if (m_pGridInner->IsRecordDirty() && !bNewValue)
		m_pGridInner->CancelRecord();
}

BOOL CKCOMRichGridCtrl::OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip) 
{
	CRect rt;
	rt.CopyRect(lpRectPos);
	rt.OffsetRect(-rt.left, -rt.top);
	rt.top += InternalGetText().GetLength() ? m_nHeaderHeight : 0;

	if (m_pGridInner)
	{
		m_pGridInner->SetWindowPos(NULL, rt.left, rt.top, rt.Width(), rt.Height(), SWP_NOZORDER);
	}
	
	return COleControl::OnSetObjectRects(lpRectPos, lpRectClip);
}

void CKCOMRichGridCtrl::InitHeaderFont()
{
	HFONT hFont;
	
	if (m_fntHeader.m_pFont->get_hFont(&hFont) == S_OK)
	{
		LOGFONT logHeader;
		if (GetObject(hFont, sizeof(logHeader), &logHeader) && m_pGridInner)
		{
			m_pGridInner->ColHeaderStyle().SetFont(CGXFont().SetLogFont(logHeader));
			m_pGridInner->UpdateFontMetrics();
			m_pGridInner->ResizeRowHeightsToFit(CGXRange().SetRows(0, 0));
		}
	}
}

void CKCOMRichGridCtrl::InitCellFont()
{
	HFONT hFont;
	
	if (m_fntCell.m_pFont->get_hFont(&hFont) == S_OK)
	{
		LOGFONT logCell;
		if (GetObject(hFont, sizeof(logCell), &logCell) && m_pGridInner)
		{
			m_pGridInner->ChangeStandardStyle(CGXStyle().SetFont(CGXFont().SetLogFont(logCell)));
			m_pGridInner->UpdateFontMetrics();
		}
	}
}

void CKCOMRichGridCtrl::InitGridHeader()
{
	SetFieldSeparator(m_strFieldSeparator);

	m_pGridInner->Initialize();
	m_pGridInner->GetParam()->EnableUndo(FALSE);
	m_pGridInner->RegisterControl(GX_IDS_CTRL_EDIT, new CKCOMEditControl(
		m_pGridInner, GX_IDS_CTRL_EDIT));
	m_pGridInner->RegisterControl(GX_IDS_CTRL_COMBOBOX, new CKCOMComboBox(
		m_pGridInner, GX_IDS_CTRL_COMBOBOX, GX_IDS_CTRL_COMBOBOX, GXCOMBO_NOTEXTFIT));
	m_pGridInner->RegisterControl(GX_IDS_CTRL_COMBOLIST, new CKCOMComboBox(
		m_pGridInner, GX_IDS_CTRL_COMBOBOX, GX_IDS_CTRL_COMBOBOX, GXCOMBO_NOTEXTFIT, TRUE));
	m_pGridInner->RegisterControl(KCOM_IDS_CTRL_DATEEDIT, new CKCOMDateControl(
		m_pGridInner, KCOM_IDS_CTRL_DATEEDIT));
	m_pGridInner->RegisterControl(GX_IDS_CTRL_MASKEDIT, new CKCOMMaskControl(
		m_pGridInner, KCOM_IDS_CTRL_DATEEDIT));
	m_pGridInner->RegisterControl(KCOM_IDS_CTRL_EDITBUTTON, new CKCOMEditBtn(
		m_pGridInner, KCOM_IDS_CTRL_DATEEDIT));
	m_pGridInner->SetStyleRange(CGXRange().SetTable(), CGXStyle().SetUserAttribute(
		IDS_UA_GRIDTYPE, "Grid"));

	UpdateDivider();

	m_pGridInner->GetParam()->EnableTrackRowHeight(GX_TRACK_ALL);
	m_pGridInner->GetParam()->EnableSelection(GX_SELFULL);
	m_pGridInner->GetParam()->EnableMoveRows(FALSE);
	m_pGridInner->GetParam()->EnableTrackColWidth(m_bAllowColumnSizing ? 
		GX_TRACK_INDIVIDUAL : FALSE);
	m_pGridInner->GetParam()->EnableTrackRowHeight(m_bAllowRowSizing);
	m_pGridInner->SetFrozenCols(m_nFrozenCols, 0);
	
	if(AmbientUserMode())
	{
		if (m_bSelectByCell)
		{
			m_pGridInner->GetParam()->SetSpecialMode(GX_MODELBOX_MS);
			m_pGridInner->GetParam()->SetActivateCellFlags(FALSE);
			m_pGridInner->GetParam()->EnableMoveCols(FALSE);
			m_pGridInner->GetParam()->SetHideCurrentCell(GX_HIDE_ALLWAYS);
		}
		else
		{
			m_pGridInner->GetParam()->SetSpecialMode(0);
			m_pGridInner->GetParam()->SetActivateCellFlags(TRUE);
			m_pGridInner->GetParam()->EnableMoveCols(TRUE);
			m_pGridInner->GetParam()->SetHideCurrentCell(GX_HIDE_INACTIVEFRAME);
		}
	}
	
	m_pGridInner->ChangeStandardStyle(CGXStyle().SetTextColor(TranslateColor(GetForeColor())));
	m_pGridInner->ChangeStandardStyle(CGXStyle().SetInterior(TranslateColor(GetBackColor())));
	m_pGridInner->ChangeStandardStyle(CGXStyle().SetVerticalAlignment(DT_VCENTER));
	m_pGridInner->RowHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
	m_pGridInner->RowHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
	m_pGridInner->ColHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
	m_pGridInner->ColHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
	m_pGridInner->SetScrollBarMode(SB_BOTH, AmbientUserMode() ? gxnAutomatic | gxnEnhanced : gxnDisabled);

	InitHeaderFont();
	InitCellFont();
	if (!AmbientUserMode())
	{
		m_pGridInner->SetReadOnly(FALSE);
		m_pGridInner->ShowWindow(SW_HIDE);
	}
	
	m_pGridInner->SetRowHeight(0, 0, m_nHeaderHeight);
	m_pGridInner->SetDefaultRowHeight(m_nRowHeight);
	m_pGridInner->SetDefaultColWidth(m_nDefColWidth);
	CGXData * pData = m_pGridInner->GetParam()->GetData();

	if (!AmbientUserMode())
	{
		m_pGridInner->SetCanAppend(m_bAllowAddNew);
		m_pGridInner->SetRowCount(3);

		InitGridFromProp();
	}
	else
	{
		m_pGridInner->SetCanAppend(m_bAllowAddNew);
		if (m_nDataMode)
			m_pGridInner->SetRowCount(m_nRows);

		InitGridFromProp();
		
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
		m_pGridInner->Redraw();

		if (m_nDataMode && m_apColumns.GetSize() == 0)
		{
			CGridColumn * pNewCol;
			for (int i = 0; i < m_nCols; i ++)
			{
				pNewCol = new CGridColumn;
				InitColumnObject(i + 1, pNewCol);
				m_apColumns.Add(pNewCol);
			}
		}

		m_pGridInner->SetReadOnly(!m_bAllowUpdate);
	}

	m_pGridInner->HideRows(0, 0, !m_bColumnHeaders);
	m_pGridInner->HideCols(0, 0, !m_bRecordSelectors);

	if (AmbientUserMode())
		m_pGridInner->GetParam()->EnableUndo(TRUE);

	InvalidateControl();
}

void CKCOMRichGridCtrl::OnForeColorChanged() 
{
	if (m_pGridInner)
		m_pGridInner->ChangeStandardStyle(CGXStyle().SetTextColor(TranslateColor(GetForeColor())));
	
	COleControl::OnForeColorChanged();
	BoundPropertyChanged(DISPID_FORECOLOR);
}

void CKCOMRichGridCtrl::UpdateControl()
{
	USES_CONVERSION;

	if (!AmbientUserMode())
		return;

	m_pGridInner->LockUpdate(TRUE);
	m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
	m_pGridInner->GetParam()->SetLockReadOnly(FALSE);

	m_bDataSrcChanged = FALSE;

	if (!GetColumnInfo())
	{
		m_pGridInner->GetParam()->SetLockReadOnly(bLock);
		m_pGridInner->GetParam()->EnableUndo(TRUE);
		m_pGridInner->LockUpdate(FALSE);
		m_pGridInner->Redraw();

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
			pCol->nDataType = GetDataTypeFromStr(pCol->strDataField);
			pCol->bForceNoNullable = !(m_pColumnInfo[pCol->nDataField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE);
			pCol->bForceLock = !(m_pColumnInfo[pCol->nDataField].dwFlags & DBCOLUMNFLAGS_WRITE);

			pCol->bLocked = FALSE;
			pCol->bNullable = TRUE;
			pCol->bAllowSizing = TRUE;
			pCol->bPromptInclude = TRUE;
			pCol->bVisible = TRUE;
			pCol->nAlignment = 2;
			pCol->nCaptionAlignment = 2;
			pCol->nCase = 0;
			pCol->nFieldLen = 255;
			pCol->nStyle = 0;
			pCol->nWidth = m_nDefColWidth;
			pCol->strCaption = m_apColumnData[i]->strName;
			pCol->strName = m_apColumnData[i]->strName;
			pCol->strPromptChar = _T("_");

			m_apColumnsProp.Add(pCol);
		}
	}
	else
	{
		ColumnProp * pCol;

		for (int i = 0; i < m_apColumnsProp.GetSize(); i ++)
		{
			pCol = m_apColumnsProp[i];
			pCol->nDataField = GetFieldFromStr(pCol->strDataField);
			pCol->bForceNoNullable = !(m_pColumnInfo[pCol->nDataField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE);
			pCol->bForceLock = !(m_pColumnInfo[pCol->nDataField].dwFlags & DBCOLUMNFLAGS_WRITE);
		}
	}

	m_pGridInner->LockUpdate(TRUE);
	InitGridHeader();
	Bind();
	m_pGridInner->LockUpdate(FALSE);

	m_Rowset.SetupOptionalRowsetInterfaces();
	FetchNewRows();
	
	// Set up the sink so we know when the rowset is repositioned
	if (AmbientUserMode())
	{
		AfxConnectionAdvise(m_spRowPosition, IID_IRowPositionChange, &m_xRowPositionChange, FALSE, &m_dwCookieRPC);
		AfxConnectionAdvise(m_Rowset.m_spRowset, IID_IRowsetNotify, &m_xRowsetNotify, FALSE, &m_dwCookieRN);
		m_spDataSource->addDataSourceListener((IDataSourceListener*)&m_xDataSourceListener);
	}

	m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pGridInner->GetParam()->EnableUndo(TRUE);
	m_pGridInner->LockUpdate(FALSE);
	m_pGridInner->Redraw();
}

long CKCOMRichGridCtrl::GetHeaderHeight() 
{
	return m_nHeaderHeight;
}

void CKCOMRichGridCtrl::SetHeaderHeight(long nNewValue) 
{
	if (nNewValue < 0 || nNewValue > 15000)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	m_nHeaderHeight = nNewValue;
	CRect rt;

	if (m_pGridInner)
	{
		GetClientRect(&rt);
		rt.top += InternalGetText().GetLength() ? m_nHeaderHeight : 0;
		m_pGridInner->SetWindowPos(NULL, rt.left, rt.top, rt.Width(), rt.Height(), SWP_NOZORDER);
		m_pGridInner->SetRowHeight(0, 0, m_nHeaderHeight);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeaderHeight);
	InvalidateControl();
}

void CKCOMRichGridCtrl::OnHeadForeColorChanged() 
{
	if (m_pGridInner)
	{
		m_pGridInner->RowHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
		m_pGridInner->ColHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadForeColor);
	InvalidateControl();
}

void CKCOMRichGridCtrl::OnHeadBackColorChanged() 
{
	if (m_pGridInner)
	{
		m_pGridInner->RowHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
		m_pGridInner->ColHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadBackColor);
	InvalidateControl();
}

void CKCOMRichGridCtrl::UpdateDivider()
{
	if (m_pGridInner)
	{
		m_pGridInner->ChangeStandardStyle((CGXStyle().SetBorders(
			gxBorderAll, CGXPen().SetStyle(PS_NULL))));
		
		switch (m_nDividerType)
		{
		case 1:
			m_pGridInner->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderRight, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1).SetColor(TranslateColor(m_clrDivider)))));
			break;
			
		case 2:
			m_pGridInner->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderBottom, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1).SetColor(TranslateColor(m_clrDivider)))));
			break;
			
		case 3:
			m_pGridInner->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderRight, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1).SetColor(TranslateColor(m_clrDivider)))));
			m_pGridInner->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderBottom, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1).SetColor(TranslateColor(m_clrDivider)))));
			break;
		}

		m_pGridInner->RowHeaderStyle().SetBorders(gxBorderAll, CGXPen().SetStyle(PS_NULL));
		m_pGridInner->ColHeaderStyle().SetBorders(gxBorderAll, CGXPen().SetStyle(PS_NULL));
	}
}

ROWCOL CKCOMRichGridCtrl::GetRowColFromVariant(VARIANT va)
{
	if (va.vt == VT_ERROR)
		return GX_INVALID;

	COleVariant vaNew = va;
	vaNew.ChangeType(VT_BSTR);
	vaNew.ChangeType(VT_I4);

	return vaNew.lVal;
}

BOOL CKCOMRichGridCtrl::OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray) 
{
	switch (dispid)
	{
	case dispidFieldSeparator:
		pStringArray->Add(_T("(tab)"));
		pStringArray->Add(_T("(space)"));
		pCookieArray->Add(0);
		pCookieArray->Add(1);

		return TRUE;
		break;
	}
	
	return COleControl::OnGetPredefinedStrings(dispid, pStringArray, pCookieArray);
}

BOOL CKCOMRichGridCtrl::OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut) 
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
	}
	
	return COleControl::OnGetPredefinedValue(dispid, dwCookie, lpvarOut);
}

void CKCOMRichGridCtrl::InitColumnObject(int nCol, CGridColumn * pCol)
{
	CString strName;

	strName.Format("Column%4d", m_nColsUsed ++);

	if (IsWindow(m_hWnd))
		COleControl::SetRedraw(FALSE);

	pCol->SetCtrlPtr(this);
	pCol->prop.nColIndex = m_pGridInner->GetColIndex(nCol);
	pCol->SetName(strName);
	pCol->SetCaption(strName);
	pCol->SetAllowSizing(TRUE);
	pCol->SetBackColor(GetBackColor());
	pCol->SetForeColor(GetForeColor());
	pCol->SetHeadBackColor(m_clrHeadBack);
	pCol->SetHeadForeColor(m_clrHeadFore);
	pCol->SetAlignment(2);
	pCol->SetCaptionAlignment(2);
	pCol->SetCase(0);
	pCol->SetDataType(0);
	pCol->SetFieldLen(255);
	pCol->SetLocked(FALSE);
	pCol->SetMask("");
	pCol->SetNullable(TRUE);
	pCol->SetPromptChar("_");
	pCol->SetStyle(0);
	pCol->SetVisible(TRUE);
	pCol->SetWidth(m_nDefColWidth);

	if (IsWindow(m_hWnd))
		COleControl::SetRedraw();
	
	InvalidateControl();
}

void CKCOMRichGridCtrl::InitGridFromProp()
{
	CGridColumn * pCol;
	ColumnProp * pProp;

	int nCols = m_apColumnsProp.GetSize();
	int nRows = nCols == 0 ? 0 : m_arContent.GetSize() / nCols;

	m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
	m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	m_pGridInner->SetColCount(0);
	m_pGridInner->SetRowCount(0);
	
	if (nCols > 0)
		m_nCols = nCols;
	if (m_nDataMode == 0)
		m_nRows = AmbientUserMode() ? 0 : 3;
	else if (nCols > 0)
		m_nRows = nRows;

	m_pGridInner->SetColCount(m_nCols);
	m_pGridInner->SetRowCount(m_nRows);

	if (IsWindow(m_hWnd))
		COleControl::SetRedraw(FALSE);

	int i;
	for (i = 0; i < m_apColumns.GetSize(); i ++)
		delete m_apColumns[i];

	m_apColumns.RemoveAll();

	int nPrevPos = 0, nNextPos = 0;
	for (i = 0; i < nCols; i++)
	{
		pCol = new CGridColumn;
		pProp = m_apColumnsProp[i];

		pCol->SetCtrlPtr(this);
		pCol->prop.nColIndex = m_pGridInner->GetColIndex(i + 1);
		pCol->prop.strChoiceList = pProp->strChoiceList;
		pCol->m_arChoiceList.RemoveAll();
		if (!pCol->prop.strChoiceList.IsEmpty())
		{
			while (nNextPos != -1)
			{
				nNextPos = pCol->prop.strChoiceList.Find('\n', nPrevPos);
				
				if (nNextPos != -1)
					pCol->m_arChoiceList.Add(pCol->prop.strChoiceList.
						Mid(nPrevPos, nNextPos - nPrevPos));
				else
					pCol->m_arChoiceList.Add(pCol->prop.strChoiceList.
						Mid(nPrevPos));

				nPrevPos = nNextPos + 1;
			}
		}

		pCol->prop.strDataField = pProp->strDataField;
		pCol->prop.nDataField = pProp->nDataField;
		pCol->prop.bLocked = pProp->bLocked;
		pCol->prop.bNullable = pProp->bNullable;
		pCol->prop.bForceLock = pProp->bForceLock;
		pCol->prop.bForceNoNullable = pProp->bForceNoNullable;
			
		if (!pCol->prop.bLocked && pCol->prop.bForceLock)
			pCol->prop.bLocked = TRUE;
		if (pCol->prop.bNullable && pCol->prop.bForceNoNullable)
			pCol->prop.bNullable = FALSE;

		pCol->SetAlignment(pProp->nAlignment);
		pCol->SetAllowSizing(pProp->bAllowSizing);
		pCol->SetBackColor(pProp->clrBack);
		pCol->SetCaption(pProp->strCaption);
		pCol->SetCaptionAlignment(pProp->nCaptionAlignment);
		pCol->SetCase(pProp->nCase);
		pCol->SetDataType(pProp->nDataType);
		pCol->SetFieldLen((short)pProp->nFieldLen);
		pCol->SetForeColor(pProp->clrFore);
		pCol->SetHeadBackColor(pProp->clrHeadBack);
		pCol->SetHeadForeColor(pProp->clrHeadFore);
		pCol->SetNullable(pCol->prop.bNullable);
		pCol->SetLocked(pCol->prop.bLocked);
		pCol->SetMask(pProp->strMask);
		pCol->SetName(pProp->strName);
		pCol->SetPromptChar(pProp->strPromptChar);
		pCol->SetStyle(pProp->nStyle);
		pCol->SetVisible(pProp->bVisible);
		pCol->SetWidth(pProp->nWidth);
		pCol->SetPromptInclude(pProp->bPromptInclude);
		pCol->SetComboBoxStyle(pProp->nComboBoxStyle);

		m_apColumns.Add(pCol);
	}

	m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pGridInner->GetParam()->EnableUndo(TRUE);
	
	if (IsWindow(m_hWnd))
		COleControl::SetRedraw(TRUE);
	
	InvalidateControl();

	BoundPropertyChanged(dispidRows);
	BoundPropertyChanged(dispidCols);

	SetModifiedFlag();
}

void CKCOMRichGridCtrl::SerializeColumnProp(CArchive &ar, BOOL bLoad)
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
			ar << pCol->strPromptChar << pCol->bPromptInclude << pCol->nComboBoxStyle;
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
			ar >> pCol->strPromptChar >> pCol->bPromptInclude >> pCol->nComboBoxStyle;

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

BOOL CKCOMRichGridCtrl::PreTranslateMessage(MSG* pMsg) 
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
			if (m_hWndDropDown && m_bDroppedDown)
			{
				::SendMessage((HWND)m_hWndDropDown, WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
				return TRUE;
			}
		}

		if (nCharShort == VK_LEFT || nCharShort == VK_RIGHT || nCharShort == VK_UP || 
			nCharShort == VK_DOWN || nCharShort == VK_DELETE || nCharShort == VK_TAB)
		{
			GetFocus()->SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);
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
	}
	
	return COleControl::PreTranslateMessage(pMsg);
}

void CKCOMRichGridCtrl::ButtonDown(UINT nFlags, CPoint point)
{
	FireMouseDown(nFlags, _AfxShiftState(), point.x, point.y);
}

void CKCOMRichGridCtrl::ButtonUp(UINT nFlags, CPoint point)
{
	FireMouseUp(nFlags, _AfxShiftState(), point.x, point.y);
	
	CRect rect;
	GetClientRect(&rect);
	if (rect.PtInRect(point))
		FireClick();
}

void CKCOMRichGridCtrl::ButtonDblClick(UINT nFlags, CPoint point)
{
	FireDblClick();
}

HWND CKCOMRichGridCtrl::GetApplicationWindow()
{
	HWND  hwnd = NULL;
	HRESULT hr;
	
	
	if (m_pInPlaceSite != NULL)
	{
		m_pInPlaceSite->GetWindow(&hwnd);
		if (hwnd!=NULL)
			return hwnd;
		else
			return AfxGetMainWnd( )->GetSafeHwnd();
	}
	
	
	LPOLECLIENTSITE pOleClientSite = GetClientSite();
	
	if ( pOleClientSite )
	{
		IOleWindow* pOleWindow;
		hr = pOleClientSite->QueryInterface( IID_IOleWindow, (LPVOID*) 
			&pOleWindow );
		
		if ( pOleWindow )
		{
			pOleWindow->GetWindow( &hwnd );
			pOleWindow->Release();
			if (hwnd!=NULL)
				return hwnd;
			else
				
				return AfxGetMainWnd( )->GetSafeHwnd();
		}
	}
	
	return AfxGetMainWnd()->GetSafeHwnd();
}

void CKCOMRichGridCtrl::CreateDrawGrid()
{
	if (m_pGridInner)
		return;

	HWND hWnd = GetApplicationWindow(); 
	
	m_pGridInner = new CGridInner(this);
	m_pGridInner->Create(WS_VISIBLE | WS_CHILD, CRect(100,100, 110, 110), CWnd::FromHandle(hWnd), 101);
	InitGridHeader();
}

int CKCOMRichGridCtrl::GetFieldFromStr(CString str)
{
	if (str.IsEmpty())
		return -1;

	for (int i = 0; i < m_apColumnData.GetSize(); i ++)
		if (!str.CompareNoCase(m_apColumnData[i]->strName))
			return m_apColumnData[i]->nColumn;

	return -1;
}

ROWCOL CKCOMRichGridCtrl::GetRowFromBmk(BYTE *bmk)
{
	BYTE * bmkSearch = new BYTE[m_nBookmarkSize];
	memcpy(bmkSearch, bmk, sizeof(bmk));

	ROWCOL i, nRows = m_apBookmark.GetSize();
	for (i = 0; i < nRows; i++)
	{
		GetBmkOfRow(i + 1);
		if (m_apBookmark[i] && memcmp(m_apBookmark[i], bmkSearch, m_pColumnInfo->ulColumnSize) == 0)
		{
			m_bIsPosSync = TRUE;
			delete [] bmkSearch;
			return i + 1;
		}
	}
	
	delete [] bmkSearch;
	m_bIsPosSync = FALSE;
	return GX_INVALID;
}

HRESULT CKCOMRichGridCtrl::GetRowset()
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

void CKCOMRichGridCtrl::GetCellValue(ROWCOL nRow, ROWCOL nCol, CGXStyle & style)
{
	if (!AmbientUserMode())
		return;

	if (m_nDataMode != 0)
		return;
	
	if (!m_bEndReached && nRow == (ROWCOL)m_pGridInner->GetRecordCount())
		FetchNewRows();

	if ((ROWCOL)m_apBookmark.GetSize() < nRow)
		return;

	BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
	m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
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
	
	CGridColumn * pCol;
	ROWCOL nColToSet;
	for (int j = 0; j < m_apColumns.GetSize(); j ++)
	{
		pCol = m_apColumns[j];
		if (pCol->prop.nDataField == -1)
			continue;
		
		nColToSet = m_pGridInner->GetColFromIndex(pCol->prop.nColIndex);
		if (nColToSet != nCol)
			continue;

		if (m_data.dwStatus[j] == DBSTATUS_S_OK && style.GetValue().Compare(m_data.strData[j]))
			style.SetValue(m_data.strData[j]);
		else if (m_data.dwStatus[j] == DBSTATUS_S_ISNULL && !style.GetValue().IsEmpty())
			style.SetValue(_T(""));
	}
	
	m_Rowset.FreeRecordMemory();
	
	m_pGridInner->GetParam()->SetLockReadOnly(bLock);
}

void CKCOMRichGridCtrl::SetCurrentCell(ROWCOL nRow, ROWCOL nCol)
{
	if (m_nDataMode != 0)
		return;
	
	if ((int)nRow > m_apBookmark.GetSize())
		return;

	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
	
	HRESULT hr = m_Rowset.MoveToBookmark(bmk);
	if (FAILED(hr))
		return;
	
	HROW hRowNow = 0;
	DBPOSITIONFLAGS dwPositionFlags;
	hr = m_spRowPosition->GetRowPosition(NULL, &hRowNow, &dwPositionFlags);
	if (FAILED(hr) || hRowNow == DB_NULL_HROW || hRowNow == m_Rowset.m_hRow)
		return;

	m_spRowPosition->ClearRowPosition();
	m_spRowPosition->SetRowPosition(NULL, m_Rowset.m_hRow, DBPOSITION_OK);
	m_Rowset.FreeRecordMemory();
}

ROWCOL CKCOMRichGridCtrl::GetRowFromHRow(HROW hRow)
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

void CKCOMRichGridCtrl::UpdateCellFromSrc(ROWCOL nRow, ROWCOL nCol)
{
	m_pGridInner->RedrawRowCol(nRow, nCol);

	return;

	USES_CONVERSION;

	if ((int)nRow >= m_apBookmark.GetSize())
		return;

	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow]);
	
	HRESULT hr = m_Rowset.MoveToBookmark(bmk);
	if (FAILED(hr))
		return;

	int i = GetColOrdinalFromField(nCol);
	if (i == - 1)
	{
		m_Rowset.FreeRecordMemory();
		return;
	}

	CGXData * pData = m_pGridInner->GetParam()->GetData();
	if (m_data.dwStatus[i] == DBSTATUS_S_OK && m_pGridInner->GetValueRowCol(nRow, nCol)
			.CompareNoCase(m_data.strData[i]))
	{
		pData->StoreValueRowCol(nRow + 1, m_pGridInner->GetColFromIndex(m_apColumns[i]->prop.
			nColIndex), m_data.strData[i], gxOverride);

		m_pGridInner->RedrawRowCol(CGXRange(nRow + 1, m_pGridInner->GetColFromIndex(m_apColumns[i]
			->prop.nColIndex)));
	}

	m_Rowset.FreeRecordMemory();
}

void CKCOMRichGridCtrl::DeleteRowFromSrc(ROWCOL nRow)
{
	BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
	m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	m_pGridInner->GetParam()->EnableUndo(FALSE);

	if (m_pGridInner->IsAddNewMode())
		m_pGridInner->CancelRecord();
	else
	{
		m_pGridInner->CancelRecord();
		m_pGridInner->DeleteRows(nRow, nRow);

		delete [] m_apBookmark[nRow - 1];
		m_apBookmark.RemoveAt(nRow - 1);
	}

	m_pGridInner->GetParam()->EnableUndo(TRUE);
	m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pGridInner->Redraw();
}

void CKCOMRichGridCtrl::UpdateRowFromSrc(ROWCOL nRow)
{
	m_pGridInner->RedrawRowCol(CGXRange().SetRows(nRow));

	return;

	USES_CONVERSION;

	if ((int)nRow >= m_apBookmark.GetSize())
		return;

	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow]);
	
	HRESULT hr = m_Rowset.MoveToBookmark(bmk);
	if (FAILED(hr))
		return;
	
	ROWCOL nRowCurr, nColCurr;
	m_pGridInner->GetCurrentCell(nRowCurr, nColCurr);

	ROWCOL i, j;
	CGXData * pData = m_pGridInner->GetParam()->GetData();
	for (i = 1; i <= m_pGridInner->GetColCount(); i ++)
	{
		j = GetColOrdinalFromCol(i);
		if (m_apColumns[j]->prop.nDataField == -1)
			continue;

		if (m_data.dwStatus[j] == DBSTATUS_S_OK && m_pGridInner->GetValueRowCol(nRow + 1, i)
			.CompareNoCase(m_data.strData[j]))
			pData->StoreValueRowCol(nRow + 1, i, m_data.strData[j], gxOverride);
		else if (m_data.dwStatus[j] == DBSTATUS_S_ISNULL && !m_pGridInner->GetValueRowCol(nRow + 1, i)
			.IsEmpty())
			pData->StoreValueRowCol(nRow + 1, i, _T(""), gxOverride);
	}

	m_pGridInner->RedrawRowCol(CGXRange().SetRows(nRow + 1));
	m_pGridInner->SetCurrentCell(nRowCurr, nColCurr);

	m_Rowset.FreeRecordMemory();
}

void CKCOMRichGridCtrl::ClearInterfaces()
{
	FreeColumnMemory();
	FreeColumnData();

	if (m_pGridInner && AmbientUserMode())
	{
		m_pGridInner->GetParam()->EnableUndo(FALSE);
		BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
		m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
		m_pGridInner->SetColCount(0);
		m_pGridInner->SetRowCount(0);
		m_pGridInner->GetParam()->SetLockReadOnly(bLock);
		m_pGridInner->GetParam()->EnableUndo(TRUE);
	}

	if (m_dwCookieRPC != 0)
	{
		AfxConnectionUnadvise(m_spRowPosition, IID_IRowPositionChange, &m_xRowPositionChange, FALSE, m_dwCookieRPC);
		m_dwCookieRPC = 0;
	}

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

void CKCOMRichGridCtrl::FreeColumnMemory()
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

void CKCOMRichGridCtrl::FetchNewRows()
{
	USES_CONVERSION;

	m_pGridInner->CancelRecord();

	int i = 0, nFetchOnce;
	ULONG nRowsAvailable = 0;
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

	nFetchOnce = nRowsAvailable - m_pGridInner->GetRecordCount();
	if (nFetchOnce <= 0)
		return;
	
	m_pGridInner->GetParam()->EnableUndo(FALSE);
	BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
	m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	m_pGridInner->InsertRows(m_pGridInner->GetRecordCount() + 1, nFetchOnce);

	m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pGridInner->GetParam()->EnableUndo(TRUE);
	m_pGridInner->Redraw();
	m_pGridInner->UpdateScrollbars();
}

BOOL CKCOMRichGridCtrl::GetColumnInfo()
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

	return m_apColumnData.GetSize() > 0;
}

HRESULT CKCOMRichGridCtrl::UpdateRowData(ROWCOL nRow)
{
	if (m_Rowset.m_spRowset == NULL)
		return E_FAIL;

	BOOL bInsert = m_pGridInner->IsAddNewMode();
	short bCancel = 0;
	short bDispMsg = 0;

	if (!bInsert)
		FireBeforeUpdate(&bCancel);
	if (bCancel)
		return E_FAIL;

	if (nRow > (ROWCOL)m_apBookmark.GetSize() + bInsert ? 1: 0)
		return E_FAIL;

	CAccessorRowset<CManualAccessor> rac;
	RowData bindData;
	ROWCOL i;

	for (i = 0; i < 255; i ++)
		bindData.dwStatus[i] = DBSTATUS_S_OK;

	HRESULT hr;
	CString str;

	rac.m_spRowset = m_Rowset.m_spRowset;

	int nColumnsWritable = 0;
	ROWCOL nColOrdinal;
	CGridColumn * pCol;
	for (i = 1; i <= m_pGridInner->GetColCount(); i ++)
	{
		if (!m_pGridInner->m_bColDirty[i - 1] && !bInsert)
			continue;
		
		pCol = m_apColumns[GetColOrdinalFromCol(i)];
		if (pCol->prop.nDataField == -1)
			continue;

		if (m_pColumnInfo[pCol->prop.nDataField].dwFlags & DBCOLUMNFLAGS_WRITE)
			nColumnsWritable ++;
	}

	if (nColumnsWritable == 0)
		return S_OK;
	
	rac.CreateAccessor(nColumnsWritable, &bindData, sizeof(bindData));
	for (i = 1; i <= m_pGridInner->GetColCount(); i ++)
	{
		if (!m_pGridInner->m_bColDirty[i - 1] && !bInsert)
			continue;
		
		nColOrdinal = GetColOrdinalFromCol(i);
		pCol = m_apColumns[nColOrdinal];
		if (pCol->prop.nDataField == -1)
			continue;

		if (m_pColumnInfo[pCol->prop.nDataField].dwFlags & DBCOLUMNFLAGS_WRITE)
			rac.AddBindEntry(pCol->prop.nDataField, DBTYPE_STR, 255 * sizeof(TCHAR), 
				&(bindData.strData[nColOrdinal]), NULL, &(bindData.dwStatus[nColOrdinal]));
	}

	hr = rac.Bind();
	if (FAILED(hr))
	{
		ReportError();
		return E_FAIL;
	}

	if (!bInsert)
	{
		CBookmark<0> bmk;
		bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
		hr = rac.MoveToBookmark(bmk);
		if (FAILED(hr))
		{
			FireUpdateError();
			
			return E_FAIL;
		}
		rac.FreeRecordMemory();
	}

	rac.SetupOptionalRowsetInterfaces();
	
	for (i = 1; i <= m_pGridInner->GetColCount(); i ++)
	{
		if (!m_pGridInner->m_bColDirty[i - 1] && !bInsert)
			continue;
		
		nColOrdinal = GetColOrdinalFromCol(i);
		pCol = m_apColumns[nColOrdinal];
		if (pCol->prop.nDataField == -1)
			continue;
		
		if (m_pColumnInfo[pCol->prop.nDataField].dwFlags & DBCOLUMNFLAGS_WRITE)
		{
			str = m_pGridInner->m_arContent[i - 1];
			if (str.IsEmpty() && pCol->prop.bNullable)
			{
				bindData.dwStatus[nColOrdinal] = DBSTATUS_S_ISNULL;
			}
			else
			{
				lstrcpy(bindData.strData[nColOrdinal], str);
				bindData.dwStatus[nColOrdinal] = DBSTATUS_S_OK;
			}
		}
	}
	
	if (!bInsert)
	{
		try
		{
			hr = rac.SetData();
		}
		catch (CException * e)
		{
			e->ReportError();
			e->Delete();
			hr = E_FAIL;
		}
	}
	else
	{
		try
		{
			hr = rac.Insert(0, true);
		}
		catch (CException * e)
		{
			e->ReportError();
			e->Delete();
			hr = E_FAIL;
		}
	}

	if (FAILED(hr))
	{
		ReportError();
		
		for (i = 1; i <= m_pGridInner->GetColCount(); i ++)
		{
			nColOrdinal = GetColOrdinalFromCol(i);
			pCol = m_apColumns[nColOrdinal];
			
			if (pCol->prop.nDataField == -1)
				continue;
			
			if (bindData.dwStatus[nColOrdinal] != DBSTATUS_S_OK && bindData.dwStatus[nColOrdinal]
				!= DBSTATUS_S_ISNULL && bindData.dwStatus[nColOrdinal] != DBSTATUS_S_DEFAULT)
				break;
		}
		
		if (i <= m_pGridInner->GetColCount())
		{
			m_pGridInner->SetCurrentCell(nRow, i);
		}
		
		FireUpdateError();
		
		return E_FAIL;
	}
	
	try
	{
		hr = rac.Update();
	}
	catch (CException * e)
	{
		e->ReportError();
		e->Delete();
		hr = E_FAIL;
	}
	if (FAILED(hr) && hr != E_NOINTERFACE)
	{
		ReportError();
		rac.Undo();
		
		FireUpdateError();
		
		return E_FAIL;
	}
	
	if (bInsert)
	{
		FireAfterInsert(&bDispMsg);
		if (bDispMsg)
			AfxMessageBox(IDS_AFTERINSERT);
	}
	else
	{
		FireAfterUpdate(&bDispMsg);
		if (bDispMsg)
			AfxMessageBox(IDS_AFTERUPDATE);
	}

	return S_OK;
}


ROWCOL CKCOMRichGridCtrl::GetColOrdinalFromIndex(ROWCOL nColIndex)
{
	for (int i = 0; i < m_apColumns.GetSize(); i ++)
		if (m_apColumns[i]->prop.nColIndex == nColIndex)
			return i;

	return GX_INVALID;
}

ROWCOL CKCOMRichGridCtrl::GetColOrdinalFromCol(ROWCOL nCol)
{
	return GetColOrdinalFromIndex(m_pGridInner->GetColIndex(nCol));
}

ROWCOL CKCOMRichGridCtrl::GetColOrdinalFromField(int nField)
{
	for (int i = 0; i < m_apColumns.GetSize(); i ++)
		if (m_apColumns[i]->prop.nDataField == nField)
			return i;

	return GX_INVALID;
}

void CKCOMRichGridCtrl::DeleteRowData(ROWCOL nRow)
{
	short bCancel = 0, bDispMsg = 0;
	nRow = nRow >= 0 ? nRow : 0;

	if (m_apBookmark.GetSize() < (int)nRow)
		return;

	FireBeforeDelete(&bCancel, &bDispMsg);
	if (bCancel)
		return;
	if (bDispMsg && AfxMessageBox(IDS_BEFOREDELETE, MB_YESNO) != IDYES)
		return;

	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
	HRESULT hr = m_Rowset.MoveToBookmark(bmk);
	if (FAILED(hr))
		return;

	m_Rowset.FreeRecordMemory();
	HROW hRow = m_Rowset.m_hRow;
	
	try
	{
		hr = m_Rowset.Delete();
	}
	catch (CException * e)
	{
		e->ReportError();
		e->Delete();
		hr = E_FAIL;
	}
	if (FAILED(hr))
	{
		ReportError();
		return;
	}

	m_Rowset.m_hRow = hRow;
	try
	{
		hr = m_Rowset.Update();
	}
	catch (CException * e)
	{
		e->ReportError();
		e->Delete();
		hr = E_FAIL;
	}
	if (FAILED(hr) && hr != E_NOINTERFACE)
	{
		ReportError();
		m_Rowset.m_hRow = hRow;
		m_Rowset.Undo();

		return;
	}
	else
	{
		bDispMsg = 0;
		FireAfterDelete(&bDispMsg);
		if (bDispMsg)
			AfxMessageBox(IDS_AFTERDELETE);

		return;
	}
}

int CKCOMRichGridCtrl::GetDataTypeFromStr(CString strField)
{
	if (strField.IsEmpty())
		return 0;

	for (int i = 0; i < m_apColumnData.GetSize(); i ++)
		if (!strField.CompareNoCase(m_apColumnData[i]->strName))
		{
			switch (m_apColumnData[i]->vt)
			{
			case DBTYPE_UI1:
			case DBTYPE_I1:
				return 2;
				break;

			case DBTYPE_I2:
			case DBTYPE_UI2:
				return 3;
				break;

			case DBTYPE_I4:
			case DBTYPE_UI4:
				return 4;
				break;

			case DBTYPE_R4:
			case DBTYPE_NUMERIC:
			case DBTYPE_DECIMAL:
				return 5;
				break;

			case DBTYPE_CY:
				return 6;
				break;

			case DBTYPE_DATE:
			case DBTYPE_FILETIME:
			case DBTYPE_DBDATE:
			case DBTYPE_DBTIME:
			case DBTYPE_DBTIMESTAMP:
				return 7;
				break;

			case DBTYPE_BOOL:
				return 1;
				break;

			default:
				return 0;
				break;
			}
		}

	return 0;
}

void CKCOMRichGridCtrl::UndeleteRecordFromHRow(HROW hRow, ROWCOL nRow)
{
	m_Rowset.ReleaseRows();
	m_Rowset.m_hRow = hRow;
	if (FAILED(m_Rowset.GetData()))
		return;

	if (nRow > m_pGridInner->GetRowCount())
		nRow = m_pGridInner->GetRowCount();
	
	BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
	m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
	
	BYTE * pBookmarkCopy;
	if (m_bHaveBookmarks)
	{
		pBookmarkCopy = new BYTE[m_nBookmarkSize];
		if (pBookmarkCopy != NULL)
			memcpy(pBookmarkCopy, &m_data.bookmark, m_nBookmarkSize);
		
		m_apBookmark.InsertAt(nRow - 1, pBookmarkCopy);
	}
	
	m_pGridInner->InsertRows(nRow, 1);
	
	m_Rowset.FreeRecordMemory();

	m_pGridInner->GetParam()->SetLockReadOnly(bLock);
	m_pGridInner->Redraw();
}

void CKCOMRichGridCtrl::GetBmkOfRow(ROWCOL nRow)
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

short CKCOMRichGridCtrl::GetCtrlType() 
{
	return 0;
}

void CKCOMRichGridCtrl::OnDropDown(ROWCOL nRow, ROWCOL nCol)
{
	if (nRow <= 0 || nCol <= 0)
		return;

	CGridColumn * pCol = m_apColumns[GetColOrdinalFromCol(nCol)];
	
	m_hWndDropDown = (HWND)pCol->m_dropDownWnd;
	if (pCol->prop.nStyle != 1 || !IsWindow(m_hWndDropDown))
		return;

	if (m_bDroppedDown)
	{
		HideDropDownWnd();
		return;
	}

	CPoint pt;
	CRect rect = m_pGridInner->CalcRectFromRowCol(nRow, nCol);
	pt.x = rect.left;
	pt.y = rect.bottom;
	m_pGridInner->ClientToScreen(&pt);

	long lParam = pt.x + pt.y;
	lParam ^= 235;
	pt.x ^= 235;
	pt.y ^= 235;
	long wParam = MAKEWPARAM(pt.x, pt.y);

	BringInfo * pBi = new BringInfo;
	pBi->wParam = wParam;
	pBi->lParam = lParam;
	pBi->hWnd = m_hWnd;
	
	CString strText;
	CGXStyle style;

	if (m_pGridInner->IsCurrentCell(nRow, nCol) && m_pGridInner->IsActiveCurrentCell())
		m_pGridInner->GetControl(nRow, nCol)->GetValue(strText);
	else
		strText = m_pGridInner->GetValueRowCol(nRow, nCol);

	lstrcpy(pBi->strText, strText);

	::SendMessage(m_hWndDropDown, CKCOMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
}

void CKCOMRichGridCtrl::HideDropDownWnd()
{
	if (IsWindow(m_hWndDropDown) && m_bDroppedDown)
	{
		// hide the auto dropped down window
		BringInfo * pBi = new BringInfo;
		pBi->hWnd = NULL;
		
		::SendMessage(m_hWndDropDown, CKCOMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
		m_bDroppedDown = FALSE;
	}
}

void CKCOMRichGridCtrl::OnBringMsg(long wParam, long lParam)
{
	BringInfo * pBi = (BringInfo *)wParam;
	if (!pBi)
		return;
	
	if ((pBi->hWnd))
		// this message tell me the dropdown status and the dropdown
		// window's handle
	{
		m_bDroppedDown = pBi->wParam;
		TRACE1("DroppedDown is %d\t", m_bDroppedDown);
		TRACE1("hookhandle is %x\n", m_hSystemHook);
		if (m_bDroppedDown)
		{
			m_hExternalDropDownWnd = pBi->hWnd;
			HookMouse(m_hExternalDropDownWnd);
		}
		else if (m_hSystemHook)
		{
			UnhookWindowsHookEx(m_hSystemHook);
			m_hSystemHook = NULL;
		}
	}
	else
	{
		ROWCOL nRow, nCol;
		m_pGridInner->GetCurrentCell(nRow, nCol);

		BOOL bLock = m_pGridInner->GetParam()->IsLockReadOnly();
		m_pGridInner->GetParam()->SetLockReadOnly(FALSE);
		m_pGridInner->SetValueRange(CGXRange(nRow, nCol), pBi->strText);
		m_pGridInner->GetParam()->SetLockReadOnly(bLock);

		m_bDroppedDown = FALSE;
		if (m_hSystemHook)
		{
			UnhookWindowsHookEx(m_hSystemHook);
			m_hSystemHook = NULL;
		}
	}

	// free the memory previously allocated by message sender
	delete pBi;
}

void CKCOMRichGridCtrl::HookMouse(HWND hWnd)
{
	SetHost(m_hWndDropDown, hWnd);
	// Install keyboard hook to trap all keyboard messages
	m_hSystemHook = SetWindowsHookEx(WH_MOUSE, MouseProc, (HINSTANCE)GetCurrentThread(), 0);
	
	if (!m_hSystemHook)
	{
#ifdef _DEBUG
		AfxMessageBox("Failed to install hook");
#endif
	}
}

void CKCOMRichGridCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	COleControl::OnKillFocus(pNewWnd);

	HideDropDownWnd();
}

OLE_COLOR CKCOMRichGridCtrl::GetBackColor() 
{
	return m_clrBack;
}

void CKCOMRichGridCtrl::SetBackColor(OLE_COLOR nNewValue) 
{
	m_clrBack = nNewValue;

	if (m_pGridInner)
		m_pGridInner->ChangeStandardStyle(CGXStyle().SetInterior(TranslateColor(GetBackColor())));

	SetModifiedFlag();
	BoundPropertyChanged(dispidBackColor);
	InvalidateControl();
}

void CKCOMRichGridCtrl::OnTextChanged() 
{
	CRect rt;

	if (m_pGridInner)
	{
		GetClientRect(&rt);
		rt.top += InternalGetText().GetLength() ? m_nHeaderHeight : 0;
		m_pGridInner->SetWindowPos(NULL, rt.left, rt.top, rt.Width(), rt.Height(), 
			SWP_NOZORDER);
	}

	SetModifiedFlag();
	COleControl::OnTextChanged();
}

void CKCOMRichGridCtrl::OnDividerColorChanged() 
{
	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerColor);
	InvalidateControl();
}

void CKCOMRichGridCtrl::Bind()
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
		m_pGridInner->GetParam()->SetLockReadOnly(TRUE);
		m_pGridInner->GetParam()->EnableUndo(TRUE);
		m_pGridInner->LockUpdate(FALSE);
		m_pGridInner->Redraw();
	}
}

void CKCOMRichGridCtrl::ExchangeStockProps(CPropExchange *pPX)
{
	BOOL bLoading = pPX->IsLoading();
	DWORD dwStockPropMask = GetStockPropMask();
	DWORD dwPersistMask = dwStockPropMask;

	PX_ULong(pPX, _T("_StockProps"), dwPersistMask);

	if (dwStockPropMask & (STOCKPROP_CAPTION | STOCKPROP_TEXT))
	{
		CString strText;

		if (dwPersistMask & (STOCKPROP_CAPTION | STOCKPROP_TEXT))
		{
			if (!bLoading)
				strText = InternalGetText();
			if (dwStockPropMask & STOCKPROP_CAPTION)
				PX_String(pPX, _T("Caption"), strText, _T(""));
			if (dwStockPropMask & STOCKPROP_TEXT)
				PX_String(pPX, _T("Text"), strText, _T(""));
		}
		if (bLoading)
		{
			TRY
				SetText(strText);
			END_TRY
		}
	}

	if (dwStockPropMask & STOCKPROP_FORECOLOR)
	{
		if (dwPersistMask & STOCKPROP_FORECOLOR)
			PX_Color(pPX, _T("ForeColor"), m_clrForeColor, AmbientForeColor());
		else if (bLoading)
			m_clrForeColor = AmbientForeColor();
	}

	if (dwStockPropMask & STOCKPROP_BACKCOLOR)
	{
		if (dwPersistMask & STOCKPROP_BACKCOLOR)
			PX_Color(pPX, _T("BackColor"), m_clrBackColor, AmbientBackColor());
		else if (bLoading)
			m_clrBackColor = AmbientBackColor();
	}

	if (dwStockPropMask & STOCKPROP_FONT)
	{
		LPFONTDISP pFontDispAmbient = AmbientFont();
		BOOL bChanged = TRUE;

		if (dwPersistMask & STOCKPROP_FONT)
			bChanged = PX_Font(pPX, _T("Font"), m_font, NULL, pFontDispAmbient);
		else if (bLoading)
			m_font.InitializeFont(NULL, pFontDispAmbient);

		if (bLoading && bChanged)
			OnFontChanged();

		RELEASE(pFontDispAmbient);
	}

	if (dwStockPropMask & STOCKPROP_BORDERSTYLE)
	{
		short sBorderStyle = m_sBorderStyle;

		if (dwPersistMask & STOCKPROP_BORDERSTYLE)
			PX_Short(pPX, _T("BorderStyle"), m_sBorderStyle, 1);
		else if (bLoading)
			m_sBorderStyle = 0;

		if (sBorderStyle != m_sBorderStyle)
			_AfxToggleBorderStyle(this);
	}

	if (dwStockPropMask & STOCKPROP_ENABLED)
	{
		BOOL bEnabled = m_bEnabled;

		if (dwPersistMask & STOCKPROP_ENABLED)
			PX_Bool(pPX, _T("Enabled"), m_bEnabled, TRUE);
		else if (bLoading)
			m_bEnabled = TRUE;

		if ((bEnabled != m_bEnabled) && (m_hWnd != NULL))
			::EnableWindow(m_hWnd, m_bEnabled);
	}

	if (dwStockPropMask & STOCKPROP_APPEARANCE)
	{
		short sAppearance = m_sAppearance;

		if (dwPersistMask & STOCKPROP_APPEARANCE)
			PX_Short(pPX, _T("Appearance"), m_sAppearance, 0);
		else if (bLoading)
			m_sAppearance = AmbientAppearance();

		if (sAppearance != m_sAppearance)
			_AfxToggleAppearance(this);
	}
}

VARIANT CKCOMRichGridCtrl::GetBookmark() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	ROWCOL nRow, nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);

	VARIANT sa;
	VariantInit(&sa);
	GetBmkOfRow(nRow, &sa);

	return sa;
}

void CKCOMRichGridCtrl::SetBookmark(const VARIANT FAR& newValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	SetRow(GetRowFromBmk(&newValue));
}

ROWCOL CKCOMRichGridCtrl::GetRowFromIndex(ROWCOL nIndex)
{
	for (ROWCOL nRow = 1; nRow <= m_pGridInner->GetRowCount(); nRow ++)
		if (m_pGridInner->GetRowIndex(nRow) == nIndex)
			return nRow;

	return GX_INVALID;
}

void CKCOMRichGridCtrl::GetBmkOfRow(ROWCOL nRow, VARIANT *pVaRet)
{
	if (m_nDataMode)
	{
		pVaRet->vt = VT_I4;
		pVaRet->lVal = m_pGridInner->GetRowIndex(nRow);

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

ROWCOL CKCOMRichGridCtrl::GetRowFromBmk(const VARIANT *pBmk)
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

LPDISPATCH CKCOMRichGridCtrl::GetSelBookmarks()
{
	CSelBookmarks * pBmks = new CSelBookmarks;
	pBmks->SetCtrlPtr(this);

	return pBmks->GetIDispatch(FALSE);
}

VARIANT CKCOMRichGridCtrl::RowBookmark(long RowIndex) 
{
	VARIANT vaResult;
	VariantInit(&vaResult);

	GetBmkOfRow(RowIndex, &vaResult);

	return vaResult;
}

void CKCOMRichGridCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_KCOMRICHGRID);
	dlgAbout.DoModal();
}

long CKCOMRichGridCtrl::GetFrozenCols() 
{
	return m_nFrozenCols;
}

void CKCOMRichGridCtrl::SetFrozenCols(long nNewValue) 
{
	if (nNewValue < 0 || nNewValue > m_nCols)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE);

	m_nFrozenCols = nNewValue;
	if (m_pGridInner)
	{
		m_pGridInner->GetParam()->EnableUndo(FALSE);
		m_pGridInner->SetFrozenCols(m_nFrozenCols, 0);
		m_pGridInner->GetParam()->EnableUndo(TRUE);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidFrozenCols);
}

int CKCOMRichGridCtrl::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message) 
{
	if (AmbientUserMode())
		OnActivateInPlace (TRUE, NULL);

	return COleControl::OnMouseActivate(pDesktopWnd, nHitTest, message);
}
