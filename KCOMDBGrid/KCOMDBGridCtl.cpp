// KCOMDBGridCtl.cpp : Implementation of the CKCOMDBGridCtrl ActiveX Control class.

#include "stdafx.h"
#include "KCOMDBGrid.h"
#include "KCOMDBGridCtl.h"
#include "KCOMDBGridPpg.h"
#include "KCOMDBGridColumnsPpg.h"
#include "GridInner.h"
#include <oledberr.h>
#include <msstkppg.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define FETCHONCE 300

short AFXAPI _ShiftState()
{
	BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
	BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
	BOOL bAlt   = (GetKeyState(VK_MENU) < 0);

	return (short)(bShift + (bCtrl << 1) + (bAlt << 2));
}

void ShowError(HRESULT hrInput)
{
	IErrorInfo *pErrorInfo = NULL;
	CString strError;
	BSTR bstrDescription = NULL;
	HINSTANCE hInstResources = AfxGetResourceHandle();

	// go and get the error information
	//
	HRESULT hr = GetErrorInfo(0, &pErrorInfo);
	if (FAILED(hr) || !pErrorInfo)
		strError.LoadString(IDS_ERROR_GENERAL);
	else
	{
		// get the source and the description
		//
		hr = pErrorInfo->GetDescription(&bstrDescription);
		pErrorInfo->Release();

		USES_CONVERSION;
		if (SUCCEEDED(hr))
		{
			strError = W2T(bstrDescription);
			SysFreeString(bstrDescription);
		}
	}

	// show the error!
	//
	AfxMessageBox(strError, MB_ICONEXCLAMATION|MB_OK|MB_TASKMODAL);
}

AFX_STATIC_DATA const FONTDESC _KCOMFontDescDefault =
	{ sizeof(FONTDESC), OLESTR("MS Sans Serif"), FONTSIZE(8), FW_NORMAL,
	  DEFAULT_CHARSET, FALSE, FALSE, FALSE };

class __CMemFile : public CMemFile
{
	friend class CKCOMDBGridCtrl;
};

IMPLEMENT_DYNCREATE(CKCOMDBGridCtrl, COleControl)

BEGIN_INTERFACE_MAP(CKCOMDBGridCtrl, COleControl)
	// Map the interfaces to their implementations
	INTERFACE_PART(CKCOMDBGridCtrl, IID_IRowPositionChange, RowPositionChange)
	INTERFACE_PART(CKCOMDBGridCtrl, IID_IRowsetNotify, RowsetNotify)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) CKCOMDBGridCtrl::XRowPositionChange::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowPositionChange)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMDBGridCtrl::XRowPositionChange::Release(void)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowPositionChange)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMDBGridCtrl::XRowPositionChange::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowPositionChange)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMDBGridCtrl::XRowPositionChange::OnRowPositionChange(
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowPositionChange)

	if (ePhase == DBEVENTPHASE_DIDEVENT)
	{
		DBPOSITIONFLAGS dwPositionFlags;
		HRESULT hr;
		HCHAPTER hChapter = DB_NULL_HCHAPTER;
		
		if (pThis->m_pGridInner->GetColCount() == (pThis->m_bShowRecordSelector ? 1 : 0))
			return S_OK;

		hr = pThis->m_spRowPosition->GetRowPosition(&hChapter, &pThis->m_Rowset.m_hRow, &dwPositionFlags);
		if (FAILED(hr) || pThis->m_Rowset.m_hRow == DB_NULL_HROW)
		{
			return S_OK;
		}
		
		hr = pThis->m_Rowset.CRowset::GetData();
		if (FAILED(hr))
		{
			return S_OK;
		}
		
		// Now we need to find a match based upon the bookmark;
		USES_CONVERSION;

		BYTE * bmkSearch = new BYTE[pThis->m_nBookmarkSize];
		memcpy(bmkSearch, pThis->m_data.bookmark, pThis->m_nBookmarkSize);
		
		int i = pThis->GetRowFromBmk(bmkSearch);
		while (i == -1 && !pThis->m_bEndReached)
		{
			pThis->FetchNewRows();
			i = pThis->GetRowFromBmk(bmkSearch);
		}

		delete [] bmkSearch;

		CCellID cell;
		if (i != -1)
		{
			cell = pThis->m_pGridInner->GetFocusCell();
			if (cell.col < 0)
				cell.col = pThis->m_bShowRecordSelector ? 1 : 0;
			if (i != cell.row - 1)
			{
				cell.row = i + 1;
				pThis->m_pGridInner->SetFocusCell(cell);
				pThis->ScrollToRow(i + 1);
			}
		}

		pThis->m_Rowset.FreeRecordMemory();
		pThis->m_Rowset.ReleaseRows();
		
		return S_OK;
	}
	
	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP_(ULONG) CKCOMDBGridCtrl::XRowsetNotify::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowsetNotify)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMDBGridCtrl::XRowsetNotify::Release(void)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowsetNotify)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMDBGridCtrl::XRowsetNotify::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowsetNotify)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMDBGridCtrl::XRowsetNotify::OnFieldChange(IRowset * prowset,
	HROW hRow, ULONG cColumns, ULONG rgColumns[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase)// && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_COLUMN_SET:
		case DBREASON_COLUMN_RECALCULATED:
		{
			int nRow;
			nRow = pThis->GetRowFromHRow(hRow);
			if (nRow == -1 || nRow >= pThis->m_pGridInner->GetRowCount() - 1)
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

STDMETHODIMP CKCOMDBGridCtrl::XRowsetNotify::OnRowChange(IRowset * prowset,
	ULONG cRows, const HROW rghRows[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowsetNotify)
	
	if (DBEVENTPHASE_ABOUTTODO == ePhase)// && pThis->m_Rowset.m_spRowset == prowset)
	{	
		switch (eReason)
		{
		case DBREASON_ROW_DELETE:
		case DBREASON_ROW_UNDOINSERT:
		{
			int nRow;
			for (ULONG i = 0; i < cRows; i ++)
			{
				nRow = pThis->GetRowFromHRow(rghRows[i]);
				if (nRow != -1)
					pThis->DeleteRowFromSrc(nRow);
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
			for (ULONG i = 0; i < cRows; i ++)
				pThis->UndeleteRecordFromHRow(rghRows[i]);
		}

		return S_OK;
		break;

		case DBREASON_ROW_UNDOCHANGE:
		case DBREASON_ROW_UPDATE:
		{
			int nRow;
			for (ULONG i = 0; i < cRows; i ++)
			{
				nRow = pThis->GetRowFromHRow(rghRows[i]);
				if (nRow != -1)
					pThis->UpdateRowFromSrc(nRow);
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

STDMETHODIMP CKCOMDBGridCtrl::XRowsetNotify::OnRowsetChange(IRowset * prowset,
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, RowsetNotify)

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
			return S_OK;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP_(ULONG) CKCOMDBGridCtrl::XDataSourceListener::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, DataSourceListener)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMDBGridCtrl::XDataSourceListener::Release(void)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, DataSourceListener)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMDBGridCtrl::XDataSourceListener::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, DataSourceListener)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMDBGridCtrl::XDataSourceListener::dataMemberChanged(BSTR bstrDM)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, DataSourceListener)

	USES_CONVERSION;
	if ((pThis->m_strDataMember.IsEmpty() && (bstrDM == NULL || bstrDM[0] == '\0')) || pThis->m_strDataMember == OLE2T(bstrDM))
	{
		pThis->ClearInterfaces();
		pThis->m_bDataSrcChanged = TRUE;
		pThis->m_bEndReached = FALSE;
		pThis->m_nColumnsBind = 0;
		pThis->UpdateControl();
	}
	
	return S_OK;
}

STDMETHODIMP CKCOMDBGridCtrl::XDataSourceListener::dataMemberAdded(BSTR bstrDM)
{
	return S_OK;
}

STDMETHODIMP CKCOMDBGridCtrl::XDataSourceListener::dataMemberRemoved(BSTR bstrDM)
{
	METHOD_PROLOGUE(CKCOMDBGridCtrl, DataSourceListener)

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

BEGIN_MESSAGE_MAP(CKCOMDBGridCtrl, COleControl)
	//{{AFX_MSG_MAP(CKCOMDBGridCtrl)
	ON_WM_CREATE()
	ON_WM_MOUSEACTIVATE()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CKCOMDBGridCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CKCOMDBGridCtrl)
	DISP_PROPERTY_NOTIFY(CKCOMDBGridCtrl, "HeaderBackColor", m_clrHeader, OnHeaderBackColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CKCOMDBGridCtrl, "BackColor", m_clrBack, OnBackColorChanged, VT_COLOR)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "DataSource", GetDataSource, SetDataSource, VT_UNKNOWN)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "DataMember", GetDataMember, SetDataMember, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "ShowRecordSelector", GetShowRecordSelector, SetShowRecordSelector, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "AllowAddNew", GetAllowAddNew, SetAllowAddNew, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "AllowDelete", GetAllowDelete, SetAllowDelete, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "ReadOnly", GetReadOnly, SetReadOnly, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "HeaderHeight", GetHeaderHeight, SetHeaderHeight, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "RowHeight", GetRowHeight, SetRowHeight, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "AllowArrowKeys", GetAllowArrowKeys, SetAllowArrowKeys, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "AllowRowSizing", GetAllowRowSizing, SetAllowRowSizing, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "Col", GetCol, SetCol, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "DefColWidth", GetDefColWidth, SetDefColWidth, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "HeaderFont", GetHeaderFont, SetHeaderFont, VT_FONT)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "Font", GetFont, SetFont, VT_FONT)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "Row", GetRow, SetRow, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "VisibleCols", GetVisibleCols, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "VisibleRows", GetVisibleRows, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "DataMode", GetDataMode, SetDataMode, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "Cols", GetCols, SetCols, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "Rows", GetRows, SetRows, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "GridLineStyle", GetGridLineStyle, SetGridLineStyle, VT_I4)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "FieldSeparator", GetFieldSeparator, SetFieldSeparator, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMDBGridCtrl, "Caption", GetCaption, SetCaption, VT_BSTR)
	DISP_FUNCTION(CKCOMDBGridCtrl, "ColContaining", ColContaining, VT_I4, VTS_I4)
	DISP_FUNCTION(CKCOMDBGridCtrl, "RowContaining", RowContaining, VT_I4, VTS_I4)
	DISP_FUNCTION(CKCOMDBGridCtrl, "Scroll", Scroll, VT_EMPTY, VTS_I4 VTS_I4)
	DISP_FUNCTION(CKCOMDBGridCtrl, "InsertRow", InsertRow, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CKCOMDBGridCtrl, "InsertCol", InsertCol, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CKCOMDBGridCtrl, "DeleteRow", DeleteRow, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CKCOMDBGridCtrl, "DeleteCol", DeleteCol, VT_EMPTY, VTS_VARIANT)
	DISP_PROPERTY_PARAM(CKCOMDBGridCtrl, "ColWidth", GetColWidth, SetColWidth, VT_I4, VTS_I2)
	DISP_PROPERTY_PARAM(CKCOMDBGridCtrl, "CellText", GetCellText, SetCellText, VT_BSTR, VTS_I2 VTS_I2)
	DISP_PROPERTY_PARAM(CKCOMDBGridCtrl, "RowText", GetRowText, SetRowText, VT_BSTR, VTS_I2)
	DISP_STOCKFUNC_REFRESH()
	DISP_STOCKPROP_FORECOLOR()
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CKCOMDBGridCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CKCOMDBGridCtrl, COleControl)
	//{{AFX_EVENT_MAP(CKCOMDBGridCtrl)
	EVENT_CUSTOM("AfterColEdit", FireAfterColEdit, VTS_I2)
	EVENT_CUSTOM("AfterColUpdate", FireAfterColUpdate, VTS_I2)
	EVENT_CUSTOM("AfterDelete", FireAfterDelete, VTS_NONE)
	EVENT_CUSTOM("AfterInsert", FireAfterInsert, VTS_NONE)
	EVENT_CUSTOM("AfterUpdate", FireAfterUpdate, VTS_NONE)
	EVENT_CUSTOM("BeforeColEdit", FireBeforeColEdit, VTS_I2  VTS_I2  VTS_PI2)
	EVENT_CUSTOM("BeforeDelete", FireBeforeDelete, VTS_PI2)
	EVENT_CUSTOM("BeforeUpdate", FireBeforeUpdate, VTS_PI2)
	EVENT_CUSTOM("Change", FireChange, VTS_NONE)
	EVENT_CUSTOM("ColResize", FireColResize, VTS_I2  VTS_PI2)
	EVENT_CUSTOM("BeforeInsert", FireBeforeInsert, VTS_PI2)
	EVENT_CUSTOM("HeadClick", FireHeadClick, VTS_I2)
	EVENT_CUSTOM("RolColChange", FireRolColChange, VTS_NONE)
	EVENT_CUSTOM("RowResize", FireRowResize, VTS_PI2)
	EVENT_CUSTOM("Scroll", FireScroll, VTS_PI2)
	EVENT_CUSTOM("BeforeColUpdate", FireBeforeColUpdate, VTS_I2  VTS_PVARIANT  VTS_PI2)
	EVENT_STOCK_KEYDOWN()
	EVENT_STOCK_KEYPRESS()
	EVENT_STOCK_CLICK()
	EVENT_STOCK_DBLCLICK()
	EVENT_STOCK_MOUSEMOVE()
	EVENT_STOCK_MOUSEUP()
	EVENT_STOCK_MOUSEDOWN()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CKCOMDBGridCtrl, 4)
	PROPPAGEID(CKCOMDBGridPropPage::guid)
	PROPPAGEID(CKCOMDBGridColumnsPpg::guid);
	PROPPAGEID(CLSID_StockColorPage)
	PROPPAGEID(CLSID_StockFontPage)
END_PROPPAGEIDS(CKCOMDBGridCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMDBGridCtrl, "KCOMDBGRID.KCOMDBGridCtrl.1",
	0xac212646, 0xfbe2, 0x11d2, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CKCOMDBGridCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DKCOMDBGrid =
		{ 0xac212644, 0xfbe2, 0x11d2, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_DKCOMDBGridEvents =
		{ 0xac212645, 0xfbe2, 0x11d2, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_IDataSource =
		{ 0x7c0ffab3, 0xcd84, 0x11d0, { 0x94, 0x9a, 0, 0xa0, 0xc9, 0x11, 0x10, 0xed } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwKCOMDBGridOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CKCOMDBGridCtrl, IDS_KCOMDBGRID, _dwKCOMDBGridOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl::CKCOMDBGridCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMDBGridCtrl

BOOL CKCOMDBGridCtrl::CKCOMDBGridCtrlFactory::UpdateRegistry(BOOL bRegister)
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
			IDS_KCOMDBGRID,
			IDB_KCOMDBGRID,
			afxRegApartmentThreading,
			_dwKCOMDBGridOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// Licensing strings

static const TCHAR BASED_CODE _szLicFileName[] = _T("KCOMDBGrid.lic");

static const WCHAR BASED_CODE _szLicString[] =
	L"Copyright (c) 1999 BHM";


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl::CKCOMDBGridCtrlFactory::VerifyUserLicense -
// Checks for existence of a user license

BOOL CKCOMDBGridCtrl::CKCOMDBGridCtrlFactory::VerifyUserLicense()
{
	return AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl::CKCOMDBGridCtrlFactory::GetLicenseKey -
// Returns a runtime licensing key

BOOL CKCOMDBGridCtrl::CKCOMDBGridCtrlFactory::GetLicenseKey(DWORD dwReserved,
	BSTR FAR* pbstrKey)
{
	if (pbstrKey == NULL)
		return FALSE;

	*pbstrKey = SysAllocString(_szLicString);
	return (*pbstrKey != NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl::CKCOMDBGridCtrl - Constructor

CKCOMDBGridCtrl::CKCOMDBGridCtrl() : m_fntHeader(NULL), m_fntCell(NULL)
{
	InitializeIIDs(&IID_DKCOMDBGrid, &IID_DKCOMDBGridEvents);

	m_pGridInner = NULL;
	m_hResPrev = m_hResHandle = NULL;
	
	m_nCol = m_nRow = -1;
	m_bIsPosSync = TRUE;
	m_bEndReached = TRUE;
	m_bReadOnly = FALSE;
	m_bDataSrcChanged = FALSE;
	m_dwCookieRPC = m_dwCookieRN = 0;
	m_bHaveBookmarks = FALSE;
	m_pColumnInfo = NULL;
	m_pStrings = NULL;
	m_nColumns = m_nColumnsBind = 0;
	m_nBookmarkSize = 0;
	m_nColumnsBind = 0;
	for (int i = 0; i < 255; i ++)
	{
		m_nColumnBindNum[255] = -1;
		lstrcpy(m_strColumnHeader[i], _T(""));
	}

	m_spDataSource = NULL;
	m_spRowPosition = NULL;
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl::~CKCOMDBGridCtrl - Destructor

CKCOMDBGridCtrl::~CKCOMDBGridCtrl()
{
	ClearInterfaces();

	if (m_pGridInner)
	{
		m_pGridInner->DestroyWindow();
		delete m_pGridInner;
		m_pGridInner = NULL;
	}
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl::OnDraw - Drawing function

void CKCOMDBGridCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (!m_pGridInner)
	{
		pdc->SetBkColor(RGB(255, 255, 255));
		pdc->SetTextColor(RGB(0, 0, 0));
		pdc->TextOut(0, 0, "No grid");
		return;
	}
	
	if (m_bDataSrcChanged)
		UpdateControl();

	CRect rtGrid = rcBounds, rtCaption = rcBounds;
	CFont * pFontOld;
	CSize szCaption;

	if (m_strCaption.GetLength())
	{
		pFontOld = SelectFontObject(pdc, m_fntHeader);

		rtCaption.bottom = rtCaption.top + m_nHeaderHeight;
		pdc->Draw3dRect(&rtCaption, GetSysColor(COLOR_3DLIGHT), GetSysColor(COLOR_3DDKSHADOW));
		rtCaption.DeflateRect(1, 1);
		pdc->Draw3dRect(&rtCaption, GetSysColor(COLOR_3DHILIGHT), GetSysColor(COLOR_3DSHADOW));
		rtCaption.DeflateRect(1, 1);
		pdc->FillRect(&rtCaption, &CBrush(TranslateColor(m_clrHeader)));
		rtCaption.InflateRect(2, 2);

		pdc->SetTextAlign(TA_CENTER);
		pdc->SetBkColor(TranslateColor(m_clrHeader));
		pdc->ExtTextOut(rtCaption.Width() / 2, (m_nHeaderHeight - pdc->
			GetTextExtent(m_strCaption).cy) / 2, ETO_CLIPPED, &rtCaption, 
			m_strCaption, NULL);

		pdc->SelectObject(pFontOld);
		rtGrid.top += m_nHeaderHeight;
	}

	if (!AmbientUserMode())
	{
		m_pGridInner->OnDraw(pdc, rtGrid, rcInvalid);
		m_pGridInner->SendMessage(WM_NCPAINT, (WPARAM)pdc->m_hDC, NULL);
	}
	else
		m_pGridInner->Invalidate();
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl::DoPropExchange - Persistence support

void CKCOMDBGridCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	PX_String(pPX, "DataMember", m_strDataMember);
	PX_Bool(pPX, "ShowRecordSelector", m_bShowRecordSelector, TRUE);
	PX_Bool(pPX, "AllowAddNew", m_bAllowAddNew, FALSE);
	PX_Bool(pPX, "AllowDelete", m_bAllowDelete, FALSE);
	PX_Long(pPX, "HeaderHeight", m_nHeaderHeight, 25);
	PX_Long(pPX, "RowHeight", m_nRowHeight, 25);
	PX_Bool(pPX, "ReadOnly", m_bReadOnly, FALSE);
	PX_Bool(pPX, "AllowArrowKeys", m_bAllowArrowKeys, TRUE);
	PX_Bool(pPX, "AllowRowSizing", m_bAllowRowSizing, TRUE);
	PX_Color(pPX, "HeaderBackColor", m_clrHeader, GetSysColor(COLOR_3DFACE));
	PX_Color(pPX, "BackColor", m_clrBack, GetSysColor(COLOR_WINDOW));
	PX_Long(pPX, "DefColWidth", m_nDefColWidth, 100);
	PX_Font(pPX, "HeaderFont", m_fntHeader, &_KCOMFontDescDefault);
	PX_Font(pPX, "Font", m_fntCell, &_KCOMFontDescDefault);
	PX_Long(pPX, "DataMode", m_nDataMode, 0);
	PX_Long(pPX, "Rows", m_nRows, 3);
	PX_Long(pPX, "Cols", m_nCols, 3);
	PX_Long(pPX, "GridLineStyle", m_nGridLineStyle, GVL_BOTH);
	PX_String(pPX, "FieldSeparator", m_strFieldSeparator, _T("\t"));
	PX_String(pPX, "Caption", m_strCaption, _T("KCOMDBGrid1"));

    if(!pPX->IsLoading())
    {
		//  The control's properties are being saved.
		// Create a CMemFile
		CMemFile memFile;

		// construct an archive for saving
		CArchive ar( &memFile, CArchive::store, 512);

		SerializeBindInfo(ar, FALSE);
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
         
				SerializeBindInfo(ar, TRUE);
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
// CKCOMDBGridCtrl::OnResetState - Reset control to default state

void CKCOMDBGridCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl::AboutBox - Display an "About" box to the user

void CKCOMDBGridCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_KCOMDBGRID);
	dlgAbout.DoModal();
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl message handlers

LPUNKNOWN CKCOMDBGridCtrl::GetDataSource() 
{
	if (m_spDataSource)
		m_spDataSource->AddRef();

	return m_spDataSource;
}

void CKCOMDBGridCtrl::SetDataSource(LPUNKNOWN newValue) 
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

	if (FAILED(newValue->QueryInterface(IID_IDataSource, (void **)&m_spDataSource)))
		return;

	if (!AmbientUserMode())
		GetColumnInfo();

	SetModifiedFlag();
	InvalidateControl();
}

BSTR CKCOMDBGridCtrl::GetDataMember() 
{
	return m_strDataMember.AllocSysString();
}

void CKCOMDBGridCtrl::SetDataMember(LPCTSTR lpszNewValue) 
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

void CKCOMDBGridCtrl::UpdateControl()
{
	USES_CONVERSION;

	if (!AmbientUserMode())
		return;

	m_pGridInner->DeleteAllItems();
	m_bDataSrcChanged = FALSE;

	if (!GetColumnInfo())
		return;

	if (m_nColumnsBind == 0)
	{
		m_nColumnsBind = m_apColumnData.GetSize();
		for (int i = 0; i < m_nColumnsBind; i ++)
		{
			m_nColumnBindNum[i] = i;
			lstrcpy(m_strColumnHeader[i], m_apColumnData[i]->strName);
		}
	}

	if (m_pGridInner)
	{
		m_pGridInner->InsertRow();
		InitGridHeader();
	}

	if (m_bAllowAddNew)
	{
		m_pGridInner->InsertRow();
		m_pGridInner->ShowInsertImage();
	}

	for (int i = 0; i < m_nColumnsBind; i ++)
		m_pGridInner->InsertCol(m_strColumnHeader[i]);

	if (!m_bShowRecordSelector)
		m_pGridInner->DeleteCol(0);

	int nEntries = m_nColumnsBind + (m_bHaveBookmarks ? 1 : 0);
	m_Rowset.CreateAccessor(nEntries, &m_data, sizeof(m_data));
	if (m_bHaveBookmarks)
		m_Rowset.AddBindEntry(0, DBTYPE_BYTES, sizeof(m_data.bookmark), &m_data.bookmark);

	for (i = 0; i < m_nColumnsBind; i ++)
	{
		m_Rowset.AddBindEntry(m_apColumnData[m_nColumnBindNum[i]]->nColumn, DBTYPE_STR, 255 * sizeof(TCHAR), 
			&(m_data.bstrData[i]), NULL, &(m_data.dwStatus[i]));
	}

	HRESULT hr = m_Rowset.Bind();
	if (FAILED(hr))
		return;

	m_Rowset.SetupOptionalRowsetInterfaces();

	FetchNewRows();
	CCellID cell;
	cell.row = 1;
	cell.col = m_bShowRecordSelector ? 1 : 0;
	if (m_pGridInner->GetRowCount() > 1 && m_pGridInner->GetColCount() > 1)
		m_pGridInner->SetFocusCell(cell);

	m_pGridInner->SetEditable(!m_bReadOnly);
	
	// Set up the sink so we know when the rowset is repositioned
	if (AmbientUserMode())
	{
//		AfxConnectionAdvise(m_spRowPosition, IID_IRowPositionChange, &m_xRowPositionChange, FALSE, &m_dwCookieRPC);
		AfxConnectionAdvise(m_Rowset.m_spRowset, IID_IRowsetNotify, &m_xRowsetNotify, FALSE, &m_dwCookieRN);
		m_spDataSource->addDataSourceListener((IDataSourceListener*)&m_xDataSourceListener);
	}
}

int CKCOMDBGridCtrl::GetRowFromBmk(BYTE *bmk)
{
	BYTE * bmkSearch = new BYTE[m_nBookmarkSize];
	memcpy(bmkSearch, bmk, sizeof(bmk));

	int i, nRows = m_apBookmark.GetSize();
	for (i = 0; i < nRows; i++)
	{
		if (m_apBookmark[i] && memcmp(m_apBookmark[i], bmkSearch, m_pColumnInfo->ulColumnSize) == 0)
		{
			m_bIsPosSync = TRUE;
			delete [] bmkSearch;
			return i;
		}
	}
	
	delete [] bmkSearch;
	m_bIsPosSync = FALSE;
	return -1;
}

HRESULT CKCOMDBGridCtrl::GetRowset()
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

int CKCOMDBGridCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	CRect rt(0, 0, lpCreateStruct->cx, lpCreateStruct->cy);
	if (m_strCaption.GetLength())
		rt.top += m_nHeaderHeight;

	m_pGridInner = new CGridInner(this);
	m_pGridInner->Create(rt, this, 2, WS_VISIBLE | WS_CHILD);
	m_pGridInner->SetTextBkColor(m_clrBack);
	m_pGridInner->SetTextColor(TranslateColor(GetForeColor()));
	m_pGridInner->SetFixedBkColor(m_clrHeader);
	m_pGridInner->m_nDefCellWidth = m_nDefColWidth;
	InitHeaderFont();
	InitCellFont();
	if (!AmbientUserMode())
	{
		InitGridHeader();
		m_pGridInner->SetEditable(FALSE);
		if (m_bShowRecordSelector)
		{
			m_pGridInner->SetItemImage(1, 0, 0);
			m_pGridInner->RedrawCell(1, 0);
		}
		m_pGridInner->ShowWindow(SW_HIDE);
	}
	else if (m_nDataMode != 0)
	{
		InitGridHeader();
	}
	
	return 0;
}

void CKCOMDBGridCtrl::GetItemText(int nRow, int nCol)
{
	if (m_nDataMode != 0 || (nRow == m_pGridInner->GetRowCount() - 1 && m_bAllowAddNew))
		return;
	
	GV_ITEM item;
	item.row = nRow;
	item.col = nCol;
	if (!m_pGridInner->GetItem(&item) || (!m_bEndReached && nRow == 
		m_pGridInner->GetRowCount() - (m_bAllowAddNew ? 2 : 1)))
		FetchNewRows();
}

void CKCOMDBGridCtrl::InitGridHeader()
{
	m_pGridInner->m_nHeaderHeight = m_nHeaderHeight;
	m_pGridInner->m_nCellHeight = m_nRowHeight;
	m_pGridInner->SetEditable(!m_bReadOnly);
	m_pGridInner->m_nDefCellWidth = m_nDefColWidth;

	if (!AmbientUserMode() && m_nDataMode == 0)
	{
		m_pGridInner->SetRowCount(3);
		m_pGridInner->SetColCount(3);
		m_pGridInner->ShowInsertImage();
	}
	else if (m_nDataMode)
	{
		m_pGridInner->SetColCount(m_nCols + 1);
		m_nCols = m_pGridInner->GetColCount() - 1;
		m_pGridInner->SetRowCount(m_nRows + 1);
		m_nRows = m_pGridInner->GetRowCount() - 1;
		if (m_bAllowAddNew)
		{
			m_pGridInner->HideInsertImage();
			m_pGridInner->InsertRow();
			m_pGridInner->ShowInsertImage();
		}
	}

	m_pGridInner->SetFixedRowCount(1);
	m_pGridInner->SetFixedColCount(m_bShowRecordSelector ? 1 : 0);
	m_pGridInner->SetColWidth(0, 30);
	InvalidateControl();
}

void CKCOMDBGridCtrl::FetchNewRows()
{
	USES_CONVERSION;

	if (m_nColumnsBind == 0)
		return;

	int i = 0, nFetchOnce;
	HRESULT hr;
	BYTE * pBookmarkCopy;
	DBBOOKMARK bmFirst = DBBMK_FIRST;
	CBookmark<0> bmk;
	ULONG nRowsAvailable = 0;
	if (m_Rowset.GetApproximatePosition(NULL, NULL, &nRowsAvailable) == S_OK)
		nFetchOnce = nRowsAvailable - m_apBookmark.GetSize();
	else
		nFetchOnce = FETCHONCE;
	
	if (m_apBookmark.GetSize() == 0)
	{
		bmk.SetBookmark(m_nBookmarkSize, (BYTE *)&bmFirst);
		hr = m_Rowset.MoveToBookmark(bmk);
	}
	else
	{
		bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[m_apBookmark.GetSize() - 1]);
		hr = m_Rowset.MoveToBookmark(bmk, 1);
	}

	if (FAILED(hr) || hr == DB_S_ENDOFROWSET || m_Rowset.m_hRow == NULL)
	{
		m_Rowset.FreeRecordMemory();
		m_bEndReached = TRUE;

		return;
	}
	else if (hr != S_OK)
		m_Rowset.GetData();

	BOOL bIsNullRow = TRUE;
	m_pGridInner->HideInsertImage();
	
	if (nFetchOnce > FETCHONCE)
		nFetchOnce = FETCHONCE;

	if (nFetchOnce > 10)
		BeginWaitCursor();

	while (i < nFetchOnce && !m_bEndReached)
	{
		if (m_bHaveBookmarks)
		{
			pBookmarkCopy = new BYTE[m_nBookmarkSize];
			if (pBookmarkCopy != NULL)
				memcpy(pBookmarkCopy, &m_data.bookmark, m_nBookmarkSize);
			
			m_apBookmark.Add(pBookmarkCopy);
		}
		
		int nRows = m_pGridInner->GetRowCount();
		m_pGridInner->InsertRow();
		
		nRows = m_pGridInner->GetRowCount();
		for (int j = 0; j < m_nColumnsBind; j ++)
			if (m_data.dwStatus[j] == DBSTATUS_S_OK)
				m_pGridInner->SetItemText(nRows - (m_bAllowAddNew ? 2 : 1), j + (m_bShowRecordSelector ? 1 : 0), m_data.bstrData[j]);
			else
				m_pGridInner->SetItemText(nRows - (m_bAllowAddNew ? 2 : 1), j + (m_bShowRecordSelector ? 1 : 0), _T(""));

		m_Rowset.FreeRecordMemory();
		i ++;
		if (i < nFetchOnce)
		{
			bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[m_apBookmark.GetSize() - 1]);
			hr = m_Rowset.MoveToBookmark(bmk, 1);
			if (FAILED(hr) || hr == DB_S_ENDOFROWSET || m_Rowset.m_hRow == NULL)
				m_bEndReached = TRUE;
			else if (hr != S_OK)
				m_Rowset.GetData();
		}
	}

	m_pGridInner->ShowInsertImage();
	CCellID cellNow = m_pGridInner->GetFocusCell();
	m_pGridInner->RedrawCell(cellNow);

	if (nFetchOnce > 10)
		EndWaitCursor();

	m_pGridInner->Invalidate();

	FireChange();
}

void CKCOMDBGridCtrl::SetFocusCell(int nRow, int nCol)
{
	if (m_nDataMode != 0)
		return;
	
	if (nRow > m_apBookmark.GetSize())
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

void CKCOMDBGridCtrl::ScrollToRow(int nRow)
{
	CCellRange cellRange = m_pGridInner->GetVisibleNonFixedCellRange();
	
	if (nRow >= cellRange.GetMinRow() && nRow <= cellRange.GetMaxRow())
		return;

	int nScrollPos = m_pGridInner->GetScrollPos32(SB_VERT);
	if (nRow > cellRange.GetMaxRow())
		nScrollPos += (nRow - cellRange.GetMaxRow()) * m_pGridInner->m_nCellHeight;
	else
		nScrollPos -= (cellRange.GetMinRow() - nRow) * m_pGridInner->m_nCellHeight;

	m_pGridInner->SetScrollPos32(SB_VERT, nScrollPos);
	m_pGridInner->Invalidate();
}
/*
HRESULT CKCOMDBGridCtrl::ModifyCell(int nRow, int nCol, CString str)
{
	BOOL bInsert = m_bAllowAddNew && nRow == m_pGridInner->GetRowCount() - 1;
	short bCancel = 0;

	if (m_nDataMode != 0)
	{
		if (!bInsert && (nRow >= m_pGridInner->GetRowCount() || nCol
			>= m_pGridInner->GetColCount()))
			return E_FAIL;
		else 
		{
			if (bInsert)
			{
				FireBeforeInsert(&bCancel);
				if (bCancel)
					return E_FAIL;
				FireAfterInsert();
			}
			else
				FireAfterUpdate();

			FireChange();
			return S_OK;
		}
	}
	else
	{
		if (!bInsert && (nRow > m_apBookmark.GetSize() || nCol > m_nColumnsBind))
		return E_FAIL;
	}

	bCancel = 0;
	FireBeforeUpdate(&bCancel);
	if (bCancel)
		return E_FAIL;

	CAccessorRowset<CManualAccessor> rac;
	BSTR bstrData = NULL;
	HRESULT hr;
	rac.m_spRowset = m_Rowset.m_spRowset;
	rac.CreateAccessor(1, &bstrData, sizeof(bstrData));
	rac.AddBindEntry(m_apColumnData[m_nColumnBindNum[nCol - (m_bShowRecordSelector ? 1 : 0)]]->nColumn, DBTYPE_BSTR, sizeof(BSTR), &bstrData);
	
	hr = rac.Bind();
	if (FAILED(hr))
	{
		rac.m_spRowsetChange.Release();
		rac.m_spRowset = NULL;
		return hr;
	}
	
	rac.SetupOptionalRowsetInterfaces();
	
	if (!bInsert)
	{
		CBookmark<0> bmk;
		bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
		hr = rac.MoveToBookmark(bmk);
		if (FAILED(hr))
		{
			rac.m_spRowsetChange.Release();
			rac.m_spRowset = NULL;
			return E_FAIL;
		}
		
		rac.FreeRecordMemory();
		bstrData = str.AllocSysString();
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
		if (FAILED(hr))
		{
			rac.m_spRowsetChange.Release();
			rac.m_spRowset = NULL;
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
			rac.m_spRowsetChange.Release();
			rac.m_spRowset = NULL;
			return E_FAIL;
		}
		
		FireAfterUpdate();
	}
	else
	{
		short bCancel = 0;
		FireBeforeInsert(&bCancel);
		if (bCancel)
		{
			rac.m_spRowsetChange.Release();
			rac.m_spRowset = NULL;
			return E_FAIL;
		}
		
		bstrData = str.AllocSysString();
		try
		{
			hr = rac.Insert();
		}
		catch (CException * e)
		{
			e->ReportError();
			e->Delete();
			hr = E_FAIL;
		}
		if (FAILED(hr))
		{
			rac.m_spRowsetChange.Release();
			rac.m_spRowset = NULL;
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
			rac.m_spRowsetChange.Release();
			rac.m_spRowset = NULL;
			return E_FAIL;
		}
		
		FireAfterInsert();
	}
	
	rac.m_spRowsetChange.Release();
	rac.m_spRowset = NULL;
	
	FireChange();
	
	return S_OK;
}
*/

BOOL CKCOMDBGridCtrl::OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip) 
{
	if (m_pGridInner)
		m_pGridInner->SetWindowPos(NULL, 0, m_strCaption.GetLength() ? 
		m_nHeaderHeight : 0, lpRectPos->right - lpRectPos->left,
		lpRectPos->bottom - lpRectPos->top - (m_strCaption.GetLength() ?
		m_nHeaderHeight : 0), SWP_NOZORDER);

	return COleControl::OnSetObjectRects(lpRectPos, lpRectClip);
}

int CKCOMDBGridCtrl::GetRowFromHRow(HROW hRow)
{
	if (hRow == 0)
		return -1;

	HROW hRowOld = m_Rowset.m_hRow;
	m_Rowset.m_hRow = hRow;
	HRESULT hr;

	hr = m_Rowset.GetData();
	if (FAILED(hr))
	{
		m_Rowset.m_hRow = hRowOld;
		return -1;
	}

	int nRow = GetRowFromBmk(m_data.bookmark);
	m_Rowset.m_hRow = hRowOld;
	m_Rowset.FreeRecordMemory();

	return nRow;
}

void CKCOMDBGridCtrl::UpdateCellFromSrc(int nRow, int nCol)
{
	USES_CONVERSION;

	if (nRow > m_apBookmark.GetSize())
		return;

	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow]);
	HRESULT hr = m_Rowset.MoveToBookmark(bmk);
	if (FAILED(hr))
		return;
	for (int i = 0; i < m_nColumnsBind; i ++)
		if (m_apColumnData[m_nColumnBindNum[i]]->nColumn == nCol)
			break;

	if (i == m_nColumnsBind)
	{
		m_Rowset.FreeRecordMemory();
		return;
	}

	if (m_data.dwStatus[i] == DBSTATUS_S_OK)
		m_pGridInner->SetItemText(nRow + 1, i + (m_bShowRecordSelector ? 1 : 0), m_data.bstrData[i]);

	m_Rowset.FreeRecordMemory();
	m_pGridInner->RedrawCell(nRow + 1, i + (m_bShowRecordSelector ? 1 : 0));
	FireChange();
}

void CKCOMDBGridCtrl::DeleteRowFromSrc(int nRow)
{
	if (m_pGridInner->m_nRowDirty == nRow + 1)
	{
		m_pGridInner->m_bDirty = FALSE;
		m_pGridInner->m_nRowDirty = -1;
		for (int i = 0; i < 255; i++)
			m_pGridInner->m_bColDirty[i] = FALSE;
	}

	m_pGridInner->DeleteRow(nRow + 1);

	delete [] m_apBookmark[nRow];
	m_apBookmark.RemoveAt(nRow);

	m_pGridInner->Invalidate();
	FireChange();
}

void CKCOMDBGridCtrl::UpdateRowFromSrc(int nRow)
{
	USES_CONVERSION;

	if (nRow > m_apBookmark.GetSize())
		return;

	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow]);
	HRESULT hr = m_Rowset.MoveToBookmark(bmk);
	if (FAILED(hr))
		return;
	for (int i = 0; i < m_nColumnsBind; i ++)
		if (m_data.dwStatus[i] == DBSTATUS_S_OK)
			m_pGridInner->SetItemText(nRow + 1, i + (m_bShowRecordSelector ? 1 : 0), m_data.bstrData[i]);
		else if (m_data.dwStatus[i] == DBSTATUS_S_ISNULL)
			m_pGridInner->SetItemText(nRow + 1, i + (m_bShowRecordSelector ? 1 : 0), _T(""));

	m_Rowset.FreeRecordMemory();
	if (m_pGridInner->m_nRowDirty == nRow + 1)
	{
		m_pGridInner->m_bDirty = FALSE;
		m_pGridInner->m_nRowDirty = -1;
		for (int i = 0; i < 255; i++)
			m_pGridInner->m_bColDirty[i] = FALSE;
	}

	m_pGridInner->RedrawRow(nRow + 1);
	FireChange();
}

void CKCOMDBGridCtrl::ClearInterfaces()
{
	FreeColumnMemory();
	FreeColumnData();

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

void CKCOMDBGridCtrl::FreeColumnMemory()
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

BOOL CKCOMDBGridCtrl::GetShowRecordSelector() 
{
	return m_bShowRecordSelector;
}

void CKCOMDBGridCtrl::SetShowRecordSelector(BOOL bNewValue) 
{
	if (bNewValue && !m_bShowRecordSelector && m_pGridInner)
		m_pGridInner->InsertCol(_T(""), DT_CENTER | DT_VCENTER | DT_SINGLELINE, 0);
	else if (!bNewValue && m_bShowRecordSelector && m_pGridInner)
		m_pGridInner->DeleteCol(0);

	m_bShowRecordSelector = bNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidShowRecordSelector);
}

HRESULT CKCOMDBGridCtrl::DeleteRowData(int nRow)
{
	short bCancel = 0;
	nRow = nRow >= 0 ? nRow : 0;

	if (m_nDataMode != 0)
	{
		if (nRow >= m_pGridInner->GetRowCount())
			return E_FAIL;

		FireBeforeDelete(&bCancel);
		if (bCancel)
			return E_FAIL;
		
		if (m_pGridInner->m_nRowDirty == nRow)
		{
			m_pGridInner->m_bDirty = FALSE;
			m_pGridInner->m_nRowDirty = -1;
			for (int i = 0; i < 255; i++)
				m_pGridInner->m_bColDirty[i] = FALSE;
		}

		m_pGridInner->DeleteRow(nRow);
		m_pGridInner->Invalidate();
		FireAfterDelete();

		return S_OK;
	}

	if (m_apBookmark.GetSize() < nRow)
		return E_FAIL;

	FireBeforeDelete(&bCancel);
	if (bCancel)
		return E_FAIL;

	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
	HRESULT hr = m_Rowset.MoveToBookmark(bmk);
	if (FAILED(hr))
		return E_FAIL;

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
		ShowError(hr);

		return E_FAIL;
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
		ShowError(hr);
		m_Rowset.m_hRow = hRow;
		m_Rowset.Undo();
		return E_FAIL;
	}
	else
	{
		FireAfterDelete();
		return S_OK;
	}
}

BOOL CKCOMDBGridCtrl::GetAllowAddNew() 
{
	return m_bAllowAddNew;
}

void CKCOMDBGridCtrl::SetAllowAddNew(BOOL bNewValue) 
{
	if (m_bReadOnly && bNewValue)
		return;

	if (!m_bAllowAddNew && bNewValue && m_pGridInner)
	{
		m_pGridInner->InsertRow();
		m_bAllowAddNew = bNewValue;
		m_pGridInner->ShowInsertImage();
	}
	else if (m_bAllowAddNew && !bNewValue && m_pGridInner)
	{
		CCellID cellNow = m_pGridInner->GetFocusCell();
		int nRows = m_pGridInner->GetRowCount();

		m_pGridInner->DeleteRow(nRows - 1);
		nRows = m_pGridInner->GetRowCount();
		if (cellNow.row >= nRows)
			cellNow.row = nRows - 1;
		
		m_pGridInner->SetFocusCell(cellNow);
		
		if (m_pGridInner->m_bInsertMode)
		{
			m_pGridInner->m_bDirty = FALSE;
			m_pGridInner->m_nRowDirty = -1;
			for (int i = 0; i < 255; i++)
				m_pGridInner->m_bColDirty[i] = FALSE;
		}
		
		m_pGridInner->m_bInsertMode = FALSE;
		m_bAllowAddNew = bNewValue;
	}

	InvalidateControl();

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowAddNew);
}

BOOL CKCOMDBGridCtrl::GetAllowDelete() 
{
	return m_bAllowDelete;
}

void CKCOMDBGridCtrl::SetAllowDelete(BOOL bNewValue) 
{
	if (m_bReadOnly && bNewValue)
		return;

	m_bAllowDelete = bNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowDelete);
}

BOOL CKCOMDBGridCtrl::GetReadOnly() 
{
	return m_bReadOnly;
}

void CKCOMDBGridCtrl::SetReadOnly(BOOL bNewValue) 
{
	m_bReadOnly = bNewValue;
	if (m_bReadOnly)
	{
		SetAllowDelete(FALSE);
		SetAllowAddNew(FALSE);
	}
	if (m_pGridInner)
		m_pGridInner->SetEditable(!m_bReadOnly);

	SetModifiedFlag();
	BoundPropertyChanged(dispidReadOnly);
}

HRESULT CKCOMDBGridCtrl::UpdateRowData(int nRow)
{
	BOOL bInsert = m_pGridInner->m_bInsertMode;
	short bCancel = 0;

	if (bInsert)
		FireBeforeInsert(&bCancel);
	else
		FireBeforeUpdate(&bCancel);
	if (bCancel)
		return E_FAIL;

	if (m_nDataMode != 0)
	{
		if (bInsert)
		{
			m_pGridInner->HideInsertImage();
			m_pGridInner->InsertRow();
			FireAfterInsert();
			m_pGridInner->ShowInsertImage();
			ScrollToRow(m_pGridInner->GetRowCount() - 1);
			m_pGridInner->RedrawRow(m_pGridInner->GetRowCount() - 1);
		}
		else
			FireAfterUpdate();

		return S_OK;
	}

	if (!bInsert && (nRow > m_apBookmark.GetSize()))
		return E_FAIL;

	CAccessorRowset<CManualAccessor> rac;
	RowData bindData;
	int i;

	for (i = 0; i < 255; i ++)
		bindData.dwStatus[i] = DBSTATUS_S_OK;

	HRESULT hr;
	CString str;

	rac.m_spRowset = m_Rowset.m_spRowset;

	int nColumnsWritable = 0;
	for (i = 0; i < m_nColumnsBind; i ++)
		if (m_pGridInner->m_bColDirty[i] && m_pColumnInfo[m_apColumnData[i]
			->nColumn].dwFlags & DBCOLUMNFLAGS_WRITE)
			nColumnsWritable ++;
	
	rac.CreateAccessor(nColumnsWritable, &bindData, sizeof(bindData));
	for (i = 0; i < m_nColumnsBind; i ++)
	{
		if (m_pGridInner->m_bColDirty[i] && m_pColumnInfo[m_apColumnData[i]
			->nColumn].dwFlags & DBCOLUMNFLAGS_WRITE)
			rac.AddBindEntry(m_apColumnData[m_nColumnBindNum[i]]->nColumn, DBTYPE_STR, 255 * sizeof(TCHAR), 
				&(bindData.bstrData[i]), NULL, &(bindData.dwStatus[i]));
	}

	hr = rac.Bind();
	if (FAILED(hr))
	{
		ShowError(hr);
		return E_FAIL;
	}

	rac.SetupOptionalRowsetInterfaces();
	
	if (!bInsert)
	{
		CBookmark<0> bmk;
		bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
		hr = rac.MoveToBookmark(bmk);
		if (FAILED(hr))
		{
			return E_FAIL;
		}
		rac.FreeRecordMemory();

		for (i = 0; i < m_nColumnsBind; i ++)
		{
			if (m_pGridInner->m_bColDirty[i] && m_pColumnInfo[m_apColumnData[i]
				->nColumn].dwFlags & DBCOLUMNFLAGS_WRITE)
			{
				str = m_pGridInner->GetItemText(nRow, i + (m_bShowRecordSelector ? 
					1 : 0));
				if (str.IsEmpty() && m_pColumnInfo[m_apColumnData[i]->nColumn].dwFlags & DBCOLUMNFLAGS_ISNULLABLE)
				{
					bindData.dwStatus[i] = DBSTATUS_S_ISNULL;
				}
				else
				{
					lstrcpy(bindData.bstrData[i], str);
					bindData.dwStatus[i] = DBSTATUS_S_OK;
				}
			}
		}

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
		if (FAILED(hr))
		{
			ShowError(hr);

			for (i = 0; i < m_nColumnsBind; i ++)
				if (bindData.dwStatus[i] != DBSTATUS_S_OK && bindData.dwStatus[i]
					!= DBSTATUS_S_ISNULL && bindData.dwStatus[i] != DBSTATUS_S_DEFAULT)
					break;

			if (i <= m_nColumnsBind)
			{
				CCellID cell;
				cell.row = nRow;
				cell.col = i + (m_bShowRecordSelector ? 1 : 0);
				m_pGridInner->SetFocusCell(cell);
			}

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
			ShowError(hr);

			rac.Undo();
			return E_FAIL;
		}

		FireAfterUpdate();
	}
	else
	{
		for (i = 0; i < m_nColumnsBind; i ++)
		{
			if (m_pGridInner->m_bColDirty[i] && m_pColumnInfo[m_apColumnData[i]
				->nColumn].dwFlags & DBCOLUMNFLAGS_WRITE)
			{
				str = m_pGridInner->GetItemText(nRow, i + (m_bShowRecordSelector ? 
					1 : 0));
				if (str.IsEmpty() && m_pColumnInfo[m_apColumnData[i]->nColumn].dwFlags & DBCOLUMNFLAGS_ISNULLABLE)
				{
					bindData.dwStatus[i] = DBSTATUS_S_ISNULL;
				}
				else
				{
					lstrcpy(bindData.bstrData[i], str);
					bindData.dwStatus[i] = DBSTATUS_S_OK;
				}
			}
		}

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
		if (FAILED(hr))
		{
			ShowError(hr);
			
			for (i = 0; i < m_nColumnsBind; i ++)
				if (bindData.dwStatus[i] != DBSTATUS_S_OK && bindData.dwStatus[i]
					!= DBSTATUS_S_ISNULL && bindData.dwStatus[i] != DBSTATUS_S_DEFAULT)
					break;

			if (i <= m_nColumnsBind)
			{
				CCellID cell;
				cell.row = nRow;
				cell.col = i + (m_bShowRecordSelector ? 1 : 0);
				m_pGridInner->SetFocusCell(cell);
			}

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
			ShowError(hr);

			rac.Undo();
			
			return E_FAIL;
		}

		FireAfterInsert();
	}


	FireChange();

	return S_OK;
}

long CKCOMDBGridCtrl::GetHeaderHeight() 
{
	return m_nHeaderHeight;
}

void CKCOMDBGridCtrl::SetHeaderHeight(long nNewValue) 
{
	if (nNewValue < 0)
		return;

	m_nHeaderHeight = nNewValue;
	if (m_pGridInner)
		m_pGridInner->SetRowHeight(0, m_nHeaderHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeaderHeight);
	InvalidateControl();
}

long CKCOMDBGridCtrl::GetRowHeight() 
{
	return m_nRowHeight;
}

void CKCOMDBGridCtrl::SetRowHeight(long nNewValue) 
{
	if (nNewValue < 0)
		return;

	m_nRowHeight = nNewValue;
	if (m_pGridInner)
		m_pGridInner->SetRowHeight(1, m_nRowHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRowHeight);
	InvalidateControl();
}

BOOL CKCOMDBGridCtrl::GetAllowArrowKeys() 
{
	return m_bAllowArrowKeys;
}

void CKCOMDBGridCtrl::SetAllowArrowKeys(BOOL bNewValue) 
{
	m_bAllowArrowKeys = bNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowArrowKeys);
}

BOOL CKCOMDBGridCtrl::GetAllowRowSizing() 
{
	return m_bAllowRowSizing;
}

void CKCOMDBGridCtrl::SetAllowRowSizing(BOOL bNewValue) 
{
	m_bAllowRowSizing = bNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowRowSizing);
}

void CKCOMDBGridCtrl::OnBackColorChanged() 
{
	if (m_pGridInner)
		m_pGridInner->SetTextBkColor(TranslateColor(m_clrBack));

	SetModifiedFlag();
	BoundPropertyChanged(dispidBackColor);
	InvalidateControl();
}

void CKCOMDBGridCtrl::OnHeaderBackColorChanged() 
{
	if (m_pGridInner)
		m_pGridInner->SetFixedBkColor(TranslateColor(m_clrHeader));

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeaderBackColor);
	InvalidateControl();
}

void CKCOMDBGridCtrl::OnForeColorChanged() 
{
	if (m_pGridInner)	
		m_pGridInner->SetTextColor(TranslateColor(GetForeColor()));

	COleControl::OnForeColorChanged();
	BoundPropertyChanged(DISPID_FORECOLOR);
}

long CKCOMDBGridCtrl::GetCol() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_nCol;
}

void CKCOMDBGridCtrl::SetCol(long nNewValue) 
{
	if (!AmbientUserMode())
	{
		CString str;
		str.LoadString(IDS_ERROR_NOSETCOLATDESIGNTIME);
		ThrowError(CTL_E_SETNOTSUPPORTED, str);
	}

	m_nCol = nNewValue;

	if (m_pGridInner && m_pGridInner->GetColCount() >= m_nCol)
	{
		CCellID cell = m_pGridInner->GetFocusCell();
		cell.col = m_nCol;
		m_pGridInner->SetFocusCell(cell);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidCol);
}

long CKCOMDBGridCtrl::GetDefColWidth() 
{
	return m_nDefColWidth;
}

void CKCOMDBGridCtrl::SetDefColWidth(long nNewValue) 
{
	if (nNewValue < 0)
		return;

	m_nDefColWidth = nNewValue;
	if (m_pGridInner)
		m_pGridInner->m_nDefCellWidth = m_nDefColWidth;

	SetModifiedFlag();
	BoundPropertyChanged(dispidDefColWidth);
}

LPFONTDISP CKCOMDBGridCtrl::GetHeaderFont() 
{
	return m_fntHeader.GetFontDispatch();
}

void CKCOMDBGridCtrl::SetHeaderFont(LPFONTDISP newValue) 
{
	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	m_fntHeader.InitializeFont(NULL, newValue);
	InitHeaderFont();
	if (!AmbientUserMode())
		SetHeaderHeight(m_pGridInner->m_nHeaderHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeaderFont);
	InvalidateControl();
}

void CKCOMDBGridCtrl::InitHeaderFont()
{
	HFONT hFont;
	if (m_fntHeader.m_pFont->get_hFont(&hFont) == S_OK)
	{
		LOGFONT logHeader;
		if (GetObject(hFont, sizeof(logHeader), &logHeader) && m_pGridInner)
		{
			memcpy(&m_pGridInner->m_LogHeaderFont, &logHeader, sizeof(logHeader));
			m_pGridInner->InitHeaderFont();
		}
	}
}

void CKCOMDBGridCtrl::InitCellFont()
{
	HFONT hFont;
	if (m_fntCell.m_pFont->get_hFont(&hFont) == S_OK)
	{
		LOGFONT logCell;
		if (GetObject(hFont, sizeof(logCell), &logCell) && m_pGridInner)
		{
			memcpy(&m_pGridInner->m_LogCellFont, &logCell, sizeof(logCell));
			m_pGridInner->InitCellFont();
		}
	}
}

LPFONTDISP CKCOMDBGridCtrl::GetFont() 
{
	return m_fntCell.GetFontDispatch();
}

void CKCOMDBGridCtrl::SetFont(LPFONTDISP newValue) 
{
	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	m_fntCell.InitializeFont(NULL, newValue);
	InitCellFont();
	if (!AmbientUserMode())
		SetRowHeight(m_pGridInner->m_nCellHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidFont);
	InvalidateControl();
}

long CKCOMDBGridCtrl::GetRow() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	CCellID cell = m_pGridInner->GetFocusCell();
	m_nRow = cell.row;

	return m_nRow;
}

void CKCOMDBGridCtrl::SetRow(long nNewValue) 
{
	if (!AmbientUserMode())
	{
		CString str;
		str.LoadString(IDS_ERROR_NOSETROWATDESIGNTIME);
		ThrowError(CTL_E_SETNOTSUPPORTED, str);
	}

	m_nRow = nNewValue;

	if (m_pGridInner && m_pGridInner->GetRowCount() >= m_nRow)
	{
		CCellID cell = m_pGridInner->GetFocusCell();
		cell.row = m_nRow;
		m_pGridInner->SetFocusCell(cell);
		ScrollToRow(m_nRow);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidRow);
}

long CKCOMDBGridCtrl::GetVisibleCols() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	if (!m_pGridInner)
		return NULL;

	CCellRange range = m_pGridInner->GetVisibleNonFixedCellRange();
	return range.GetColSpan();
}

long CKCOMDBGridCtrl::GetVisibleRows() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	if (!m_pGridInner)
		return NULL;

	CCellRange range = m_pGridInner->GetVisibleNonFixedCellRange();
	return range.GetRowSpan();
}

long CKCOMDBGridCtrl::ColContaining(long x) 
{
	CCellID cell = m_pGridInner->GetCellFromPt(CPoint(x, 1));

	return cell.col;
}

long CKCOMDBGridCtrl::RowContaining(long y) 
{
	CCellID cell = m_pGridInner->GetCellFromPt(CPoint(1, y));

	return cell.row;
}

void CKCOMDBGridCtrl::Scroll(long Cols, long Rows) 
{
	if (Cols == Rows ==0)
		return;

	int i;
	if (Cols > 0)
		for (i = 0; i < Cols; i ++)
			m_pGridInner->OnHScroll(SB_LINERIGHT, 0, NULL);
	else
		for (i = 0; i < Cols; i ++)
			m_pGridInner->OnHScroll(SB_LINELEFT, 0, NULL);

	if (Rows > 0)
		for (i = 0; i < Cols; i ++)
			m_pGridInner->OnVScroll(SB_LINEDOWN, 0, NULL);
	else
		for (i = 0; i < Cols; i ++)
			m_pGridInner->OnVScroll(SB_LINEUP, 0, NULL);
}

long CKCOMDBGridCtrl::GetDataMode() 
{
	return m_nDataMode;
}

void CKCOMDBGridCtrl::SetDataMode(long nNewValue) 
{
	if (AmbientUserMode())
		ThrowError(CTL_E_SETNOTSUPPORTEDATRUNTIME);

	if (nNewValue == 0)
		m_nDataMode = dmBind;
	else
	{
		if (m_spDataSource)
		{
			CString str;
			str.LoadString(IDS_ERROR_NOCHANGETOMANUALMODE);
			ThrowError(CTL_E_SETNOTSUPPORTED, str);
		}
		else
			m_nDataMode = dmManual;
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidDataMode);
}

BOOL CKCOMDBGridCtrl::OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray) 
{
	CString str;
	long i;

	switch (dispid)
	{
	case dispidDataMode:
		str.LoadString(IDS_DATAMODE_BIND);
		pStringArray->Add(str);
		str.LoadString(IDS_DATAMODE_MANUAL);
		pStringArray->Add(str);
		pCookieArray->Add(0);
		pCookieArray->Add(1);
		
		return TRUE;
		break;

	case dispidGridLineStyle:
		for (i = 0; i < 4; i ++)
		{
			str.LoadString(IDS_GLSTYLE_NONE + i);
			pStringArray->Add(str);
			pCookieArray->Add(i);
		}

		return TRUE;
		break;

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

BOOL CKCOMDBGridCtrl::OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut)
{
	switch (dispid)
	{
	case dispidDataMode:
		VariantInit(lpvarOut);
		lpvarOut->vt = VT_I4;
		lpvarOut->lVal = dwCookie;
		
		return TRUE;
		break;

	case dispidGridLineStyle:
		VariantInit(lpvarOut);
		lpvarOut->vt = VT_I4;
		lpvarOut->lVal = dwCookie;

		return TRUE;
		break;

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

BOOL CKCOMDBGridCtrl::OnGetDisplayString(DISPID dispid, CString& strValue)
{
	switch (dispid)
	{
	case dispidDataMode:
		if (m_nDataMode == 0)
			strValue.LoadString(IDS_DATAMODE_BIND);
		else
			strValue.LoadString(IDS_DATAMODE_MANUAL);
		
		return TRUE;
		break;

	case dispidGridLineStyle:
		strValue.LoadString(IDS_GLSTYLE_NONE + m_nGridLineStyle);
		
		return TRUE;
		break;

	case dispidFieldSeparator:
		if (m_strFieldSeparator == _T("\t"))
			strValue = _T("(tab)");
		else if (m_strFieldSeparator == _T(" "))
			strValue = _T("(space)");
		else
			strValue = m_strFieldSeparator;

		return TRUE;
		break;
	}
	
	return COleControl::OnGetDisplayString(dispid, strValue);
}

long CKCOMDBGridCtrl::GetCols() 
{
	if (!m_pGridInner)
		return m_nCols;

	if (AmbientUserMode())
		m_nCols = m_pGridInner->GetColCount() - 1;

	return m_nCols;
}

void CKCOMDBGridCtrl::SetCols(long nNewValue) 
{
	if (!m_pGridInner)
	{
		m_nCols = nNewValue;
		return;
	}

	if (m_nDataMode == 0)
	{
		CString str;
		str.LoadString(IDS_ERROR_NOSETCOLSINBINDMODE);
		ThrowError(CTL_E_SETNOTSUPPORTED, str);
	}

	if (nNewValue <= 0)
		nNewValue = 1;

	m_nCols = nNewValue;
	if (m_pGridInner)
	{
		m_pGridInner->HideInsertImage();
		m_pGridInner->SetColCount(nNewValue + 1);
		m_pGridInner->ShowInsertImage();
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidCols);
	InvalidateControl();
}

long CKCOMDBGridCtrl::GetRows() 
{
	if (!m_pGridInner)
		return m_nRows;

	if (AmbientUserMode())
		m_nRows = m_pGridInner->GetRowCount() - (m_bAllowAddNew ? 2 : 1);

	return m_nRows;
}

void CKCOMDBGridCtrl::SetRows(long nNewValue) 
{
	if (!m_pGridInner)
	{
		m_nRows = nNewValue;
		return;
	}

	if (m_nDataMode == 0)
	{
		CString str;
		str.LoadString(IDS_ERROR_NOSETROWSINBINDMODE);
		ThrowError(CTL_E_SETNOTSUPPORTED, str);
	}

	m_nRows = nNewValue;
	if (m_pGridInner)
	{
		m_pGridInner->HideInsertImage();
		m_pGridInner->SetRowCount(nNewValue + (m_bAllowAddNew ? 2 : 1));
		m_pGridInner->ShowInsertImage();
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidRows);
	InvalidateControl();
}

void CKCOMDBGridCtrl::InsertRow(const VARIANT FAR& RowBefore) 
{
	if (m_nDataMode == 0)
		return;

	if (RowBefore.vt == VT_ERROR)
	{
		m_pGridInner->HideInsertImage();
		m_pGridInner->InsertRow();
		m_pGridInner->ShowInsertImage();
		return;
	}

	COleVariant va = RowBefore;
	va.ChangeType(VT_I4);
	m_pGridInner->HideInsertImage();
	m_pGridInner->InsertRow(va.lVal);
	m_pGridInner->RedrawRow(m_pGridInner->GetRowCount() - 1);
	m_pGridInner->ShowInsertImage();
}

void CKCOMDBGridCtrl::InsertCol(const VARIANT FAR& ColBefore) 
{
	if (m_nDataMode == 0)
		return;

	if (ColBefore.vt == VT_ERROR)
	{
		m_pGridInner->InsertCol(NULL);
		return;
	}

	COleVariant va = ColBefore;
	va.ChangeType(VT_I4);
	m_pGridInner->HideInsertImage();
	m_pGridInner->InsertCol(NULL, DT_CENTER | DT_VCENTER | DT_SINGLELINE, va.lVal);
	m_pGridInner->ShowInsertImage();
}

void CKCOMDBGridCtrl::DeleteRow(const VARIANT FAR& RowIndex) 
{
	if (m_nDataMode == 0)
		return;

	if (RowIndex.vt == VT_ERROR)
	{
		m_pGridInner->DeleteRow(m_pGridInner->GetRowCount() - 1);
		return;
	}

	COleVariant va = RowIndex;
	va.ChangeType(VT_I4);
	if (!(m_bAllowAddNew && va.lVal	== (m_pGridInner->GetRowCount() - 1)))
	{
		m_pGridInner->DeleteRow(va.lVal);
	}
}

void CKCOMDBGridCtrl::DeleteCol(const VARIANT FAR& ColIndex) 
{
	if (m_nDataMode == 0)
		return;

	if (ColIndex.vt == VT_ERROR)
	{
		m_pGridInner->DeleteCol(m_pGridInner->GetColCount() - 1);
		return;
	}

	COleVariant va = ColIndex;
	va.ChangeType(VT_I4);
	if (va.lVal >= 255)
		return;

	m_pGridInner->DeleteCol(va.lVal);
}

long CKCOMDBGridCtrl::GetGridLineStyle() 
{
	return m_nGridLineStyle;
}

void CKCOMDBGridCtrl::SetGridLineStyle(long nNewValue) 
{
	if (nNewValue < 0)
		nNewValue = 0;
	else if (nNewValue > 3)
		nNewValue = 3;

	m_nGridLineStyle = nNewValue;
	if (m_pGridInner)
		m_pGridInner->SetGridLines(m_nGridLineStyle);

	SetModifiedFlag();
	BoundPropertyChanged(dispidGridLineStyle);
	InvalidateControl();
}

BOOL CKCOMDBGridCtrl::PreTranslateMessage(MSG* pMsg) 
{
	USHORT nCharShort;
	short nShiftState = _ShiftState();

	if (pMsg->hwnd == m_hWnd || ::IsChild(m_hWnd, pMsg->hwnd))
	switch (pMsg->message)
	{
	case WM_KEYDOWN:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyDown(&nCharShort, nShiftState);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));

		if ((m_bAllowArrowKeys && (nCharShort == VK_LEFT || nCharShort == VK_RIGHT || nCharShort == VK_UP || 
			nCharShort == VK_DOWN)) || nCharShort == VK_DELETE || nCharShort == VK_TAB)
		{
			GetFocus()->SendMessage(WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
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

BOOL CKCOMDBGridCtrl::GetColumnInfo()
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
		m_bHaveBookmarks = FALSE;

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

void CKCOMDBGridCtrl::SerializeBindInfo(CArchive & ar, BOOL bLoad)
{
	CString strHeader;

	if (bLoad)
	{
		ar >> m_nColumnsBind;
		for (int i = 0; i < m_nColumnsBind; i ++)
		{
			ar >> m_nColumnBindNum[i] >> strHeader;
			lstrcpy(m_strColumnHeader[i], strHeader);
		}
	}
	else
	{
		ar << m_nColumnsBind;
		for (int i = 0; i < m_nColumnsBind; i ++)
		{
			strHeader = m_strColumnHeader[i];
			ar << m_nColumnBindNum[i] << strHeader;
		}
	}
}

long CKCOMDBGridCtrl::GetColWidth(short ColIndex) 
{
	if (ColIndex >= 255)
		return 0;

	return m_pGridInner->GetColWidth(ColIndex);
}

void CKCOMDBGridCtrl::SetColWidth(short ColIndex, long nNewValue) 
{
	if (ColIndex >= 255)
		return;

	m_pGridInner->SetColWidth(ColIndex, nNewValue);

	m_pGridInner->Invalidate();
}

/////////////////////////////////////////////////////////////////////////////
// CKCOMDBGridCtrl::GetResourceHandle - Load module which contains
// locale-specific resources

HINSTANCE CKCOMDBGridCtrl::GetResourceHandle(LCID lcid)
{
	HINSTANCE hResHandle = NULL;

	// Use lcid to determine module that contains resources
	LANGID langID = LANGIDFROMLCID(lcid);
	if (PRIMARYLANGID(langID) == LANG_CHINESE && (SUBLANGID(langID) == SUBLANG_CHINESE_SIMPLIFIED
		|| SUBLANGID(langID) == SUBLANG_NEUTRAL))
		// Load resource module
		hResHandle = LoadLibrary(_T("kcomdbrp.dll"));

	return hResHandle;
}

void CKCOMDBGridCtrl::OnSetClientSite() 
{
	// Load resource module
	m_hResPrev = AfxGetResourceHandle();
	m_hResHandle = GetResourceHandle(AmbientLocaleID());
	if (m_hResHandle)
		AfxSetResourceHandle(m_hResHandle);

	COleControl::OnSetClientSite();
}

void CKCOMDBGridCtrl::OnFinalRelease() 
{
	if (m_hResHandle)
	{
		AfxSetResourceHandle(m_hResPrev);
		FreeLibrary(m_hResHandle);
	}

	COleControl::OnFinalRelease();
}

BSTR CKCOMDBGridCtrl::GetCellText(short RowIndex, short ColIndex) 
{
	if (ColIndex >= 255)
		return NULL;

	return m_pGridInner->GetItemText(RowIndex, ColIndex).AllocSysString();
}

void CKCOMDBGridCtrl::SetCellText(short RowIndex, short ColIndex, LPCTSTR lpszNewValue) 
{
	if (ColIndex >= 255 || m_nDataMode == 0)
		return;

	if (RowIndex >= (m_pGridInner->GetRowCount() - (m_bAllowAddNew ? 1 : 0))
		|| ColIndex >= (m_pGridInner->GetColCount()))
		return;

	m_pGridInner->SetItemText(RowIndex, ColIndex, lpszNewValue);
	m_pGridInner->RedrawCell(RowIndex, ColIndex);
}

BOOL CKCOMDBGridCtrl::IsColEditable(int nCol)
{
	if (m_nDataMode != 0)
		return TRUE;

	nCol -= m_bShowRecordSelector ? 1 : 0;
	if (nCol >= m_nColumnsBind || nCol < 0)
		return FALSE;

	return (m_pColumnInfo[m_apColumnData[nCol]->nColumn].dwFlags & DBCOLUMNFLAGS_WRITE);
}

BSTR CKCOMDBGridCtrl::GetFieldSeparator() 
{
	return m_strFieldSeparator.AllocSysString();
}

void CKCOMDBGridCtrl::SetFieldSeparator(LPCTSTR lpszNewValue) 
{
	if (lstrlen(lpszNewValue) == 0 || (lstrlen(lpszNewValue) > 1 && !lstrcmpi(
		lpszNewValue, _T("(space")) && !lstrcmpi(lpszNewValue, _T("(tab)"))))
	{
		CString str;
		str.LoadString(IDS_ERROR_ERRORFIELDSEPARATOR);
		AfxMessageBox(str);

		return;
	}

	m_strFieldSeparator = lpszNewValue;
	if (!m_strFieldSeparator.CompareNoCase(_T("(tab)")))
		m_strFieldSeparator = _T("\t");
	else if (!m_strFieldSeparator.CompareNoCase(_T("(space)")))
		m_strFieldSeparator = _T(" ");

	SetModifiedFlag();
	BoundPropertyChanged(dispidFieldSeparator);
}

BSTR CKCOMDBGridCtrl::GetRowText(short RowIndex) 
{
	if (!AmbientUserMode())
		GetNotSupported();

	CString strResult;

	if (RowIndex > 0 && RowIndex < m_pGridInner->GetRowCount())
	{
		int i = m_bShowRecordSelector ? 1 : 0;
		strResult += m_pGridInner->GetItemText(RowIndex, i);
		i ++;
		for (; i < m_pGridInner->GetColCount(); i ++)
			strResult += m_strFieldSeparator + m_pGridInner->GetItemText(RowIndex, i);
	}

	return strResult.AllocSysString();
}

void CKCOMDBGridCtrl::SetRowText(short RowIndex, LPCTSTR lpszNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	if (m_nDataMode == 0)
		return;

	CString strText = lpszNewValue, strField;
	int i, j = 0;

	if (RowIndex > 0 && RowIndex < m_pGridInner->GetRowCount() - (m_bAllowAddNew ? 1 : 0))
	{
		i = m_bShowRecordSelector ? 1 : 0;
		strField.Empty();
		while (j < strText.GetLength())
		{
			if (strText[j] != m_strFieldSeparator)
				strField += strText[j];
			else
			{
				m_pGridInner->SetItemText(RowIndex, i, strField);
				strField.Empty();
				i ++;
			}

			j ++;
		}

		if (i < m_pGridInner->GetColCount())
			m_pGridInner->SetItemText(RowIndex, i, strField);

		m_pGridInner->RedrawRow(RowIndex);
	}
}

BSTR CKCOMDBGridCtrl::GetCaption() 
{
	return m_strCaption.AllocSysString();
}

void CKCOMDBGridCtrl::SetCaption(LPCTSTR lpszNewValue) 
{
	CString strNewValue = lpszNewValue;
	CRect rt;

	if (m_pGridInner)
	{
		GetClientRect(&rt);
		if (m_strCaption.GetLength() && !strNewValue.GetLength())
		{
			m_pGridInner->SetWindowPos(NULL, 0, m_nHeaderHeight, rt.Width(), 
				rt.Height() - m_nHeaderHeight, SWP_NOZORDER);
		}
		else if (!m_strCaption.GetLength() && strNewValue.GetLength())
		{
			m_pGridInner->SetWindowPos(NULL, 0, 0, rt.Width(), rt.Height(), 
				SWP_NOZORDER);
		}
	}
	m_strCaption = strNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidCaption);
	InvalidateControl();
}

void CKCOMDBGridCtrl::UndeleteRecordFromHRow(HROW hRow)
{
	m_Rowset.ReleaseRows();
	m_Rowset.m_hRow = hRow;
	if (FAILED(m_Rowset.GetData()))
		return;

	BYTE * pBookmarkCopy;
	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_data.bookmark);

	ULONG nRowPos = 0;
	if (FAILED(m_Rowset.GetApproximatePosition(&bmk, &nRowPos, NULL)))
	{
		m_Rowset.FreeRecordMemory();
		return;
	}

	ULONG nRows = m_pGridInner->GetRowCount() - (m_bAllowAddNew ? 1 : 0);
	if (nRowPos > nRows)
		nRowPos = nRows;
	
	m_pGridInner->SetRedraw(FALSE);
	BOOL bIsNullRow = TRUE;
	m_pGridInner->HideInsertImage();

	if (m_bHaveBookmarks)
	{
		pBookmarkCopy = new BYTE[m_nBookmarkSize];
		if (pBookmarkCopy != NULL)
			memcpy(pBookmarkCopy, &m_data.bookmark, m_nBookmarkSize);
		
		m_apBookmark.InsertAt(nRowPos - 1, pBookmarkCopy);
	}
	
	m_pGridInner->InsertRow(nRowPos);
	for (int j = 0; j < m_nColumnsBind; j ++)
		if (m_data.dwStatus[j] == DBSTATUS_S_OK)
			m_pGridInner->SetItemText(nRowPos, j + (m_bShowRecordSelector ? 1 : 0), m_data.bstrData[j]);
		else
			m_pGridInner->SetItemText(nRowPos, j + (m_bShowRecordSelector ? 1 : 0), _T(""));
		
		m_Rowset.FreeRecordMemory();

	m_pGridInner->ShowInsertImage();
	CCellID cellNow = m_pGridInner->GetFocusCell();
	m_pGridInner->RedrawCell(cellNow);
	m_pGridInner->SetRedraw(TRUE);
	m_pGridInner->Invalidate();
	FireChange();

	m_pGridInner->RedrawRow(nRowPos);
}

int CKCOMDBGridCtrl::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message) 
{
	if (AmbientUserMode())
		OnActivateInPlace (TRUE, NULL);
	
	return COleControl::OnMouseActivate(pDesktopWnd, nHitTest, message);
}
