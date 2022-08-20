// KCOMRichDropDownCtl.cpp : Implementation of the CKCOMRichDropDownCtrl ActiveX Control class.

#include "stdafx.h"
#include "KCOMData.h"
#include "KCOMRichDropDownCtl.h"
#include "KCOMRichDropDownPpg.h"
#include <oledberr.h>
#include <msstkppg.h>
#include "DropDownColumn.h"
#include "DropDownColumns.h"
#include "DropDownRealWnd.h"
#include "GridInner.h"
#include "KCOMRichGridColumnPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CKCOMRichDropDownCtrl, COleControl)

BEGIN_INTERFACE_MAP(CKCOMRichDropDownCtrl, COleControl)
	// Map the interfaces to their implementations
	INTERFACE_PART(CKCOMRichDropDownCtrl, IID_IRowsetNotify, RowsetNotify)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) CKCOMRichDropDownCtrl::XRowsetNotify::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, RowsetNotify)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMRichDropDownCtrl::XRowsetNotify::Release(void)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, RowsetNotify)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMRichDropDownCtrl::XRowsetNotify::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, RowsetNotify)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMRichDropDownCtrl::XRowsetNotify::OnFieldChange(IRowset * prowset,
	HROW hRow, ULONG cColumns, ULONG rgColumns[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_COLUMN_SET:
		case DBREASON_COLUMN_RECALCULATED:
		{
			ROWCOL nRow;
			nRow = pThis->GetRowFromHRow(hRow);
			if (nRow == GX_INVALID || nRow > (ROWCOL)pThis->m_pDropDownRealWnd->OnGetRecordCount())
			{
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

STDMETHODIMP CKCOMRichDropDownCtrl::XRowsetNotify::OnRowChange(IRowset * prowset,
	ULONG cRows, const HROW rghRows[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, RowsetNotify)

	
	if (DBEVENTPHASE_ABOUTTODO == ePhase && prowset == pThis->m_Rowset.m_spRowset)
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

STDMETHODIMP CKCOMRichDropDownCtrl::XRowsetNotify::OnRowsetChange(IRowset * prowset,
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, RowsetNotify)

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

STDMETHODIMP_(ULONG) CKCOMRichDropDownCtrl::XDataSourceListener::AddRef(void)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, DataSourceListener)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CKCOMRichDropDownCtrl::XDataSourceListener::Release(void)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, DataSourceListener)
	return pThis->ExternalRelease();
}

STDMETHODIMP CKCOMRichDropDownCtrl::XDataSourceListener::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, DataSourceListener)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CKCOMRichDropDownCtrl::XDataSourceListener::dataMemberChanged(BSTR bstrDM)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, RowsetNotify)

	USES_CONVERSION;
	if ((pThis->m_strDataMember.IsEmpty() && (bstrDM == NULL || bstrDM[0] == '\0')) || pThis->m_strDataMember == OLE2T(bstrDM))
	{
		pThis->ClearInterfaces();
		pThis->m_bDataSrcChanged = TRUE;
		pThis->m_bEndReached = FALSE;
	}
	
	return S_OK;
}

STDMETHODIMP CKCOMRichDropDownCtrl::XDataSourceListener::dataMemberAdded(BSTR bstrDM)
{
	return S_OK;
}

STDMETHODIMP CKCOMRichDropDownCtrl::XDataSourceListener::dataMemberRemoved(BSTR bstrDM)
{
	METHOD_PROLOGUE(CKCOMRichDropDownCtrl, RowsetNotify)

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

BEGIN_MESSAGE_MAP(CKCOMRichDropDownCtrl, COleControl)
	//{{AFX_MSG_MAP(CKCOMRichDropDownCtrl)
	ON_WM_CREATE()
	ON_WM_KEYDOWN()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
	ON_REGISTERED_MESSAGE(CKCOMDataApp::m_nBringMsg, OnBringMsg)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CKCOMRichDropDownCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CKCOMRichDropDownCtrl)
	DISP_PROPERTY_NOTIFY(CKCOMRichDropDownCtrl, "HeadBackColor", m_clrHeadBack, OnHeadBackColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CKCOMRichDropDownCtrl, "HeadForeColor", m_clrHeadFore, OnHeadForeColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CKCOMRichDropDownCtrl, "ListAutoPosition", m_bListAutoPosition, OnListAutoPositionChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CKCOMRichDropDownCtrl, "ListWidthAutoSize", m_bListWidthAutoSize, OnListWidthAutoSizeChanged, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "Font", GetFont, SetFont, VT_FONT)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "DefColWidth", GetDefColWidth, SetDefColWidth, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "RowHeight", GetRowHeight, SetRowHeight, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "DividerType", GetDividerType, SetDividerType, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "DividerStyle", GetDividerStyle, SetDividerStyle, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "DataMode", GetDataMode, SetDataMode, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "LeftCol", GetLeftCol, SetLeftCol, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "FirstRow", GetFirstRow, SetFirstRow, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "Row", GetRow, SetRow, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "Col", GetCol, SetCol, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "Rows", GetRows, SetRows, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "Cols", GetCols, SetCols, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "DroppedDown", GetDroppedDown, SetNotSupported, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "DataField", GetDataField, SetDataField, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "Columns", GetColumns, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "VisibleRows", GetVisibleRows, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "VisibleCols", GetVisibleCols, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "ColumnHeaders", GetColumnHeaders, SetColumnHeaders, VT_BOOL)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "HeadFont", GetHeadFont, SetHeadFont, VT_FONT)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "HeaderHeight", GetHeaderHeight, SetHeaderHeight, VT_I4)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "DataSource", GetDataSource, SetDataSource, VT_UNKNOWN)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "DataMember", GetDataMember, SetDataMember, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "FieldSeparator", GetFieldSeparator, SetFieldSeparator, VT_BSTR)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "MaxDropDownItems", GetMaxDropDownItems, SetMaxDropDownItems, VT_I2)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "MinDropDownItems", GetMinDropDownItems, SetMinDropDownItems, VT_I2)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "ListWidth", GetListWidth, SetListWidth, VT_I2)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CKCOMRichDropDownCtrl, "Bookmark", GetBookmark, SetBookmark, VT_VARIANT)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "AddItem", AddItem, VT_EMPTY, VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "RemoveItem", RemoveItem, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "Scroll", Scroll, VT_EMPTY, VTS_I4 VTS_I4)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "MoveFirst", MoveFirst, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "MovePrevious", MovePrevious, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "MoveNext", MoveNext, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "MoveLast", MoveLast, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "MoveRecords", MoveRecords, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "ReBind", ReBind, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CKCOMRichDropDownCtrl, "RowBookmark", RowBookmark, VT_VARIANT, VTS_I4)
	DISP_STOCKPROP_FORECOLOR()
	DISP_STOCKPROP_BORDERSTYLE()
	DISP_STOCKPROP_HWND()
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CKCOMRichDropDownCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_STOCK(CKCOMRichDropDownCtrl, "CtrlType", 255, CKCOMRichDropDownCtrl::GetCtrlType, COleControl::SetNotSupported, VT_I2)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CKCOMRichDropDownCtrl, COleControl)
	//{{AFX_EVENT_MAP(CKCOMRichDropDownCtrl)
	EVENT_CUSTOM("Scroll", FireScroll, VTS_PI2)
	EVENT_CUSTOM("CloseUp", FireCloseUp, VTS_NONE)
	EVENT_CUSTOM("DropDown", FireDropDown, VTS_NONE)
	EVENT_CUSTOM("PositionList", FirePositionList, VTS_BSTR)
	EVENT_CUSTOM("ScrollAfter", FireScrollAfter, VTS_NONE)
	EVENT_STOCK_CLICK()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CKCOMRichDropDownCtrl, 4)
	PROPPAGEID(CKCOMRichDropDownPropPage::guid)
	PROPPAGEID(CKCOMRichGridPropColumnPage::guid)
	PROPPAGEID(CLSID_StockColorPage)
	PROPPAGEID(CLSID_StockFontPage)
END_PROPPAGEIDS(CKCOMRichDropDownCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CKCOMRichDropDownCtrl, "KCOMDATA.KCOMRichDropDownCtrl.1",
	0x29288cd, 0x2fff, 0x11d3, 0xb4, 0x46, 0, 0x80, 0xc8, 0xf1, 0x85, 0x22)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CKCOMRichDropDownCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DKCOMRichDropDown =
		{ 0x29288cb, 0x2fff, 0x11d3, { 0xb4, 0x46, 0, 0x80, 0xc8, 0xf1, 0x85, 0x22 } };
const IID BASED_CODE IID_DKCOMRichDropDownEvents =
		{ 0x29288cc, 0x2fff, 0x11d3, { 0xb4, 0x46, 0, 0x80, 0xc8, 0xf1, 0x85, 0x22 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwKCOMRichDropDownOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
//	OLEMISC_INVISIBLEATRUNTIME |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CKCOMRichDropDownCtrl, IDS_KCOMRICHDROPDOWN, _dwKCOMRichDropDownOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichDropDownCtrl::CKCOMRichDropDownCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CKCOMRichDropDownCtrl

BOOL CKCOMRichDropDownCtrl::CKCOMRichDropDownCtrlFactory::UpdateRegistry(BOOL bRegister)
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
			IDS_KCOMRICHDROPDOWN,
			IDB_KCOMRICHDROPDOWN,
			afxRegApartmentThreading,
			_dwKCOMRichDropDownOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichDropDownCtrl::CKCOMRichDropDownCtrl - Constructor

CKCOMRichDropDownCtrl::CKCOMRichDropDownCtrl() : m_fntCell(NULL), m_fntHeader(NULL)
{
	InitializeIIDs(&IID_DKCOMRichDropDown, &IID_DKCOMRichDropDownEvents);

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

	m_pDrawGrid = NULL;
	m_pDropDownRealWnd = NULL;
	m_bDataSrcChanged = FALSE;
	m_spDataSource = NULL;
	m_nCol = m_nRow = 0;
	m_nColsUsed = 0;
	m_bDroppedDown = FALSE;
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichDropDownCtrl::~CKCOMRichDropDownCtrl - Destructor

CKCOMRichDropDownCtrl::~CKCOMRichDropDownCtrl()
{
	if (m_pDrawGrid)
	{
		m_pDrawGrid->DestroyWindow();
		delete m_pDrawGrid;
		m_pDrawGrid = NULL;
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
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichDropDownCtrl::OnDraw - Drawing function

void CKCOMRichDropDownCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (!m_pDrawGrid && !AmbientUserMode())
		CreateDrawGrid();

	if (AmbientUserMode())
	{
		SetControlSize(1, 1);

		if (IsWindowVisible())
			ShowWindow(SW_HIDE);

		return;
	}
	else if (m_pDrawGrid)
	{
		CRect rt = rcBounds;
		m_pDrawGrid->SetGridRect(TRUE, &rt);
		m_pDrawGrid->OnGridDraw(pdc);
	}
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichDropDownCtrl::DoPropExchange - Persistence support

void CKCOMRichDropDownCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

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
	}		
 
	GlobalFree(m_hBlob);
	m_hBlob = NULL;
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichDropDownCtrl::OnResetState - Reset control to default state

void CKCOMRichDropDownCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CKCOMRichDropDownCtrl message handlers
short CKCOMRichDropDownCtrl::GetCtrlType() 
{
	return 1;
}

ROWCOL CKCOMRichDropDownCtrl::GetColOrdinalFromIndex(ROWCOL nColIndex)
{
	for (int i = 0; i < m_apColumns.GetSize(); i ++)
		if (m_apColumns[i]->prop.nColIndex == nColIndex)
			return i;

	return GX_INVALID;
}

ROWCOL CKCOMRichDropDownCtrl::GetColOrdinalFromCol(ROWCOL nCol)
{
	if (m_pDrawGrid)
		return GetColOrdinalFromIndex(m_pDrawGrid->GetColIndex(nCol));
	else if (m_pDropDownRealWnd)
		return GetColOrdinalFromIndex(m_pDropDownRealWnd->GetColIndex(nCol));

	return GX_INVALID;
}

ROWCOL CKCOMRichDropDownCtrl::GetColOrdinalFromField(int nField)
{
	for (int i = 0; i < m_apColumns.GetSize(); i ++)
		if (m_apColumns[i]->prop.nDataField == nField)
			return i;

	return GX_INVALID;
}

int CKCOMRichDropDownCtrl::GetFieldFromStr(CString str)
{
	if (str.IsEmpty())
		return -1;

	for (int i = 0; i < m_apColumnData.GetSize(); i ++)
		if (!str.CompareNoCase(m_apColumnData[i]->strName))
			return m_apColumnData[i]->nColumn;

	return -1;
}

void CKCOMRichDropDownCtrl::CreateDrawGrid()
{
	if (m_pDrawGrid)
		return;

	HWND hWnd = GetApplicationWindow(); 
	
	m_pDrawGrid = new CGridInner(NULL);
	m_pDrawGrid->Create(WS_VISIBLE | WS_CHILD, CRect(100,100, 110, 110), CWnd::FromHandle(hWnd), 101);
	InitDrawGrid();
}

HWND CKCOMRichDropDownCtrl::GetApplicationWindow()
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


int CKCOMRichDropDownCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	CRect rt(0, 0, lpCreateStruct->cx, lpCreateStruct->cy);
	if (!AmbientUserMode())
	{
		m_pDrawGrid = new CGridInner(NULL);
		m_pDrawGrid->Create(WS_VISIBLE | WS_CHILD, rt, this, 101);
		InitDrawGrid();
	}
	else
	{
		m_pDropDownRealWnd = new CDropDownRealWnd(this);
		m_pDropDownRealWnd->CreateEx(WS_EX_TOOLWINDOW,
			"GXWND", NULL, WS_CHILD | WS_BORDER, rt, GetDesktopWindow(), NULL, NULL);

		InitDropDownWnd();
	}
	
	return 0;
}

void CKCOMRichDropDownCtrl::InitDrawGrid()
{
	m_pDrawGrid->Initialize();
	m_pDrawGrid->GetParam()->EnableUndo(FALSE);

	UpdateDivider();

	m_pDrawGrid->ChangeStandardStyle(CGXStyle().SetTextColor(TranslateColor(GetForeColor())));
	m_pDrawGrid->ChangeStandardStyle(CGXStyle().SetInterior(TranslateColor(GetBackColor())));
	m_pDrawGrid->ChangeStandardStyle(CGXStyle().SetVerticalAlignment(DT_VCENTER));
	m_pDrawGrid->RowHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
	m_pDrawGrid->RowHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
	m_pDrawGrid->ColHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
	m_pDrawGrid->ColHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
	m_pDrawGrid->SetScrollBarMode(SB_BOTH, gxnDisabled);

	InitHeaderFont();
	InitCellFont();

	m_pDrawGrid->SetReadOnly(FALSE);
	
	m_pDrawGrid->SetRowHeight(0, 0, m_nHeaderHeight);
	m_pDrawGrid->SetDefaultRowHeight(m_nRowHeight);
	m_pDrawGrid->SetDefaultColWidth(m_nDefColWidth);

	m_pDrawGrid->SetCanAppend(FALSE);
	m_pDrawGrid->SetRowCount(2);

	InitGridFromProp();

	m_pDrawGrid->HideRows(0, 0, !m_bColumnHeaders);
	m_pDrawGrid->HideCols(0, 0, TRUE);
	InvalidateControl();
}

void CKCOMRichDropDownCtrl::UpdateDivider()
{
	if (m_pDrawGrid)
	{
		m_pDrawGrid->ChangeStandardStyle((CGXStyle().SetBorders(
			gxBorderAll, CGXPen().SetStyle(PS_NULL))));
		
		switch (m_nDividerType)
		{
		case 1:
			m_pDrawGrid->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderRight, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1))));
			break;
			
		case 2:
			m_pDrawGrid->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderBottom, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1))));
			break;
			
		case 3:
			m_pDrawGrid->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderRight, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1))));
			m_pDrawGrid->ChangeStandardStyle((CGXStyle().SetBorders(
				gxBorderBottom, CGXPen().SetStyle(PS_SOLID + m_nDividerStyle)
				.SetWidth(1))));
			break;
		}

		m_pDrawGrid->RowHeaderStyle().SetBorders(gxBorderAll, CGXPen().SetStyle(PS_NULL));
		m_pDrawGrid->ColHeaderStyle().SetBorders(gxBorderAll, CGXPen().SetStyle(PS_NULL));
	}
	else if (m_pDropDownRealWnd)
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

void CKCOMRichDropDownCtrl::InitCellFont()
{
	HFONT hFont;
	
	if (m_fntCell.m_pFont->get_hFont(&hFont) == S_OK)
	{
		LOGFONT logCell;
		if (GetObject(hFont, sizeof(logCell), &logCell))
		{
			if (m_pDrawGrid)
			{
				m_pDrawGrid->ChangeStandardStyle(CGXStyle().SetFont(CGXFont().SetLogFont(logCell)));
				m_pDrawGrid->UpdateFontMetrics();
			}
			else if (m_pDropDownRealWnd)
			{
				m_pDropDownRealWnd->ChangeStandardStyle(CGXStyle().SetFont(CGXFont().SetLogFont(logCell)));
				m_pDropDownRealWnd->UpdateFontMetrics();
			}
		}
	}
}

LPFONTDISP CKCOMRichDropDownCtrl::GetFont() 
{
	return m_fntCell.GetFontDispatch();
}

void CKCOMRichDropDownCtrl::SetFont(LPFONTDISP newValue) 
{
	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	m_fntCell.InitializeFont(NULL, newValue);

	InitCellFont();
	
	if (m_pDrawGrid)
		SetRowHeight(m_pDrawGrid->GetFontHeight() + 4);
	else if (m_pDropDownRealWnd)
		SetRowHeight(m_pDropDownRealWnd->GetFontHeight() + 4);

	SetModifiedFlag();
	BoundPropertyChanged(dispidFont);
	InvalidateControl();
}

long CKCOMRichDropDownCtrl::GetDefColWidth() 
{
	return m_nDefColWidth;
}

void CKCOMRichDropDownCtrl::SetDefColWidth(long nNewValue) 
{
	if (nNewValue < 0)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	m_nDefColWidth = nNewValue;
	if (m_pDrawGrid)
		m_pDrawGrid->SetDefaultColWidth(m_nDefColWidth);
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetDefaultColWidth(m_nDefColWidth);

	SetModifiedFlag();
	BoundPropertyChanged(dispidDefColWidth);
}

long CKCOMRichDropDownCtrl::GetRowHeight() 
{
	if (m_pDropDownRealWnd)
		m_nRowHeight = m_pDropDownRealWnd->GetDefaultRowHeight();

	return m_nRowHeight;
}

void CKCOMRichDropDownCtrl::SetRowHeight(long nNewValue) 
{
	if (nNewValue < 0)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	m_nRowHeight = nNewValue;

	if (m_pDrawGrid)
		m_pDrawGrid->SetDefaultRowHeight(m_nRowHeight);
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetDefaultRowHeight(m_nRowHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRowHeight);
	InvalidateControl();
}

long CKCOMRichDropDownCtrl::GetDividerType() 
{
	return m_nDividerType;
}

void CKCOMRichDropDownCtrl::SetDividerType(long nNewValue) 
{
	m_nDividerType = nNewValue;

	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerType);
	InvalidateControl();
}

long CKCOMRichDropDownCtrl::GetDividerStyle() 
{
	return m_nDividerStyle;
}

void CKCOMRichDropDownCtrl::SetDividerStyle(long nNewValue) 
{
	m_nDividerStyle = nNewValue;

	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerStyle);
	InvalidateControl();
}

long CKCOMRichDropDownCtrl::GetDataMode() 
{
	return m_nDataMode;
}

void CKCOMRichDropDownCtrl::SetDataMode(long nNewValue) 
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

long CKCOMRichDropDownCtrl::GetLeftCol() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	if (m_pDropDownRealWnd)
		return m_pDropDownRealWnd->GetLeftCol();
	else
		return 0;
}

void CKCOMRichDropDownCtrl::SetLeftCol(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetLeftCol(nNewValue);
}

long CKCOMRichDropDownCtrl::GetFirstRow() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pDropDownRealWnd->GetTopRow();
}

void CKCOMRichDropDownCtrl::SetFirstRow(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_pDropDownRealWnd->SetTopRow(nNewValue);
}

long CKCOMRichDropDownCtrl::GetRow() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	ROWCOL nCol;

	m_pDropDownRealWnd->GetCurrentCell(m_nRow, nCol);

	return m_nRow;
}

void CKCOMRichDropDownCtrl::SetRow(long nNewValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	m_nRow = nNewValue;

	if (m_pDropDownRealWnd->GetRowCount() >= (ROWCOL)m_nRow)
	{
		ROWCOL nRow, nCol;
		m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
		m_pDropDownRealWnd->SetCurrentCell(m_nRow, nCol);
	}
	else
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDROWCOLINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRow);
}

long CKCOMRichDropDownCtrl::GetCol() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	ROWCOL nRow;
	
	m_pDropDownRealWnd->GetCurrentCell(nRow, m_nCol);
	return m_nCol;
}

void CKCOMRichDropDownCtrl::SetCol(long nNewValue) 
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

long CKCOMRichDropDownCtrl::GetRows() 
{
	if (!m_pDropDownRealWnd && !m_pDrawGrid)
		return m_nRows;

	if (m_pDrawGrid)
		m_nRows = m_pDrawGrid->GetRowCount();
	else
		m_nRows = m_pDropDownRealWnd->GetRowCount();

	return m_nRows;
}

void CKCOMRichDropDownCtrl::SetRows(long nNewValue) 
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

	m_pDrawGrid->SetRowCount(m_nRows);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRows);
	InvalidateControl();
}

long CKCOMRichDropDownCtrl::GetCols() 
{
	if (!m_pDropDownRealWnd && !m_pDrawGrid)
		return m_nCols;

	if (m_pDrawGrid)
		m_nCols = m_pDrawGrid->GetColCount();
	else
		m_nCols = m_pDropDownRealWnd->GetColCount();

	return m_nCols;
}

void CKCOMRichDropDownCtrl::SetCols(long nNewValue) 
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

	if (m_pDrawGrid)
		m_pDrawGrid->SetColCount(m_nCols);

	SetModifiedFlag();
	BoundPropertyChanged(dispidCols);
	InvalidateControl();
}

BOOL CKCOMRichDropDownCtrl::GetDroppedDown() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_bDroppedDown;
}

BSTR CKCOMRichDropDownCtrl::GetDataField() 
{
	return m_strDataField.AllocSysString();
}

void CKCOMRichDropDownCtrl::SetDataField(LPCTSTR lpszNewValue) 
{
	m_strDataField = lpszNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidDataField);
}

LPDISPATCH CKCOMRichDropDownCtrl::GetColumns() 
{
	CDropDownColumns * pColumns = new CDropDownColumns;
	pColumns->SetCtrlPtr(this);

	return pColumns->GetIDispatch(FALSE);
}

long CKCOMRichDropDownCtrl::GetVisibleRows() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pDropDownRealWnd->CalcBottomRowFromRect(m_pDropDownRealWnd->GetGridRect()) - 
		m_pDropDownRealWnd->GetTopRow() + 1;
}

long CKCOMRichDropDownCtrl::GetVisibleCols() 
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pDropDownRealWnd->CalcRightColFromRect(m_pDropDownRealWnd->GetGridRect()) - 
		m_pDropDownRealWnd->GetLeftCol() + 1;
}

void CKCOMRichDropDownCtrl::SerializeColumnProp(CArchive &ar, BOOL bLoad)
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

BOOL CKCOMRichDropDownCtrl::OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip)
{
	CRect rt;
	rt.CopyRect(lpRectPos);
	rt.OffsetRect(-rt.left, -rt.top);

	if (m_pDrawGrid)
		m_pDrawGrid->SetWindowPos(NULL, rt.left, rt.top, rt.Width(), rt.Height(), SWP_NOZORDER);
	
	return COleControl::OnSetObjectRects(lpRectPos, lpRectClip);
}

void CKCOMRichDropDownCtrl::OnForeColorChanged() 
{
	if (m_pDrawGrid)
		m_pDrawGrid->ChangeStandardStyle(CGXStyle().SetTextColor(TranslateColor(GetForeColor())));
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->ChangeStandardStyle(CGXStyle().SetTextColor(TranslateColor(GetForeColor())));
	
	COleControl::OnForeColorChanged();
	BoundPropertyChanged(DISPID_FORECOLOR);
}

void CKCOMRichDropDownCtrl::OnHeadBackColorChanged() 
{
	if (m_pDrawGrid)
	{
		m_pDrawGrid->RowHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
		m_pDrawGrid->ColHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));

		SetModifiedFlag();
		BoundPropertyChanged(dispidHeadForeColor);
		InvalidateControl();
	}
	else if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->RowHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
		m_pDropDownRealWnd->ColHeaderStyle().SetTextColor(TranslateColor(m_clrHeadFore));
	}
}

void CKCOMRichDropDownCtrl::OnHeadForeColorChanged() 
{
	if (m_pDrawGrid)
	{
		m_pDrawGrid->RowHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
		m_pDrawGrid->ColHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));

		SetModifiedFlag();
		BoundPropertyChanged(dispidHeadBackColor);
		InvalidateControl();
	}
	else if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->RowHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
		m_pDropDownRealWnd->ColHeaderStyle().SetInterior(TranslateColor(m_clrHeadBack));
	}
}

void CKCOMRichDropDownCtrl::ClearInterfaces()
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

void CKCOMRichDropDownCtrl::InitColumnObject(int nCol, CDropDownColumn *pCol)
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

void CKCOMRichDropDownCtrl::FreeColumnMemory()
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

void CKCOMRichDropDownCtrl::InitGridFromProp()
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

	if (m_pDrawGrid)
	{
		m_pDrawGrid->GetParam()->EnableUndo(FALSE);
		m_pDrawGrid->GetParam()->SetLockReadOnly(FALSE);
		m_pDrawGrid->SetColCount(0);
		m_pDrawGrid->SetRowCount(0);
		
		m_pDrawGrid->SetColCount(m_nCols);
		m_pDrawGrid->SetRowCount(m_nRows);

		if (IsWindow(m_hWnd))
			COleControl::SetRedraw(FALSE);
	}
	else if (m_pDropDownRealWnd)
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
		
		if (m_pDrawGrid)
			pCol->prop.nColIndex = m_pDrawGrid->GetColIndex(i + 1);
		else if (m_pDropDownRealWnd)
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

	if (m_pDrawGrid)
	{
		m_pDrawGrid->GetParam()->SetLockReadOnly(TRUE);
		m_pDrawGrid->GetParam()->EnableUndo(TRUE);

		if (IsWindow(m_hWnd))
			COleControl::SetRedraw(TRUE);

		InvalidateControl();
	}
	else if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->GetParam()->SetLockReadOnly(TRUE);
		m_pDropDownRealWnd->GetParam()->EnableUndo(TRUE);
	}

	BoundPropertyChanged(dispidRows);
	BoundPropertyChanged(dispidCols);

	SetModifiedFlag();
}

BOOL CKCOMRichDropDownCtrl::GetColumnInfo()
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

HRESULT CKCOMRichDropDownCtrl::GetRowset()
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

void CKCOMRichDropDownCtrl::InitDropDownWnd()
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

BOOL CKCOMRichDropDownCtrl::GetColumnHeaders() 
{
	return m_bColumnHeaders;
}

void CKCOMRichDropDownCtrl::SetColumnHeaders(BOOL bNewValue) 
{
	m_bColumnHeaders = bNewValue;

	if (m_pDrawGrid)
		m_pDrawGrid->HideRows(0, 0, !m_bColumnHeaders);
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->HideRows(0, 0, !m_bColumnHeaders);

	SetModifiedFlag();
	BoundPropertyChanged(dispidColumnHeaders);
	InvalidateControl();
}

LPFONTDISP CKCOMRichDropDownCtrl::GetHeadFont() 
{
	return m_fntHeader.GetFontDispatch();
}

void CKCOMRichDropDownCtrl::SetHeadFont(LPFONTDISP newValue) 
{
	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	m_fntHeader.InitializeFont(NULL, newValue);
	InitHeaderFont();
	if (m_pDrawGrid)
		SetHeaderHeight(m_pDrawGrid->GetRowHeight(0));
	else if (m_pDropDownRealWnd)
		SetHeaderHeight(m_pDropDownRealWnd->GetRowHeight(0));

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadFont);
	InvalidateControl();
}

long CKCOMRichDropDownCtrl::GetHeaderHeight() 
{
	return m_nHeaderHeight;
}

void CKCOMRichDropDownCtrl::SetHeaderHeight(long nNewValue) 
{
	if (nNewValue < 0)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	m_nHeaderHeight = nNewValue;

	if (m_pDrawGrid)
		m_pDrawGrid->SetRowHeight(0, 0, m_nHeaderHeight);
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetRowHeight(0, 0, m_nHeaderHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeaderHeight);
	InvalidateControl();
}

void CKCOMRichDropDownCtrl::InitHeaderFont()
{
	HFONT hFont;
	
	if (m_fntHeader.m_pFont->get_hFont(&hFont) == S_OK)
	{
		LOGFONT logHeader;
		if (GetObject(hFont, sizeof(logHeader), &logHeader))
		{
			if (m_pDrawGrid)
			{
				m_pDrawGrid->ColHeaderStyle().SetFont(CGXFont().SetLogFont(logHeader));
				m_pDrawGrid->UpdateFontMetrics();
				m_pDrawGrid->ResizeRowHeightsToFit(CGXRange().SetRows(0, 0));
			}
			else if (m_pDropDownRealWnd)
			{
				m_pDropDownRealWnd->ColHeaderStyle().SetFont(CGXFont().SetLogFont(logHeader));
				m_pDropDownRealWnd->UpdateFontMetrics();
				m_pDropDownRealWnd->ResizeRowHeightsToFit(CGXRange().SetRows(0, 0));
			}
		}
	}
}

void CKCOMRichDropDownCtrl::GetCellValue(ROWCOL nRow, ROWCOL nCol, CGXStyle &style)
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

LPUNKNOWN CKCOMRichDropDownCtrl::GetDataSource() 
{
	if (m_spDataSource)
		m_spDataSource->AddRef();

	return m_spDataSource;
}

void CKCOMRichDropDownCtrl::SetDataSource(LPUNKNOWN newValue) 
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

BSTR CKCOMRichDropDownCtrl::GetDataMember() 
{
	return m_strDataMember.AllocSysString();
}

void CKCOMRichDropDownCtrl::SetDataMember(LPCTSTR lpszNewValue) 
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

BSTR CKCOMRichDropDownCtrl::GetFieldSeparator() 
{
	return m_strFieldSeparator.AllocSysString();
}

void CKCOMRichDropDownCtrl::SetFieldSeparator(LPCTSTR lpszNewValue) 
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

void CKCOMRichDropDownCtrl::AddItem(LPCTSTR Item, const VARIANT FAR& RowIndex) 
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

void CKCOMRichDropDownCtrl::RemoveItem(long RowIndex) 
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

void CKCOMRichDropDownCtrl::Scroll(long Rows, long Cols) 
{
	m_pDropDownRealWnd->DoScroll(Rows > 0 ? GX_DOWN : GX_UP, abs(Rows));
	m_pDropDownRealWnd->DoScroll(Cols > 0 ? GX_RIGHT : GX_LEFT, abs(Cols));
}

void CKCOMRichDropDownCtrl::RemoveAll() 
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

void CKCOMRichDropDownCtrl::MoveFirst() 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	m_pDropDownRealWnd->MoveTo(0);
}

void CKCOMRichDropDownCtrl::MovePrevious() 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	ROWCOL nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
	m_pDropDownRealWnd->MoveTo(nRow - 1);
}

void CKCOMRichDropDownCtrl::MoveNext() 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	ROWCOL nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
	m_pDropDownRealWnd->MoveTo(nRow + 1);
}

void CKCOMRichDropDownCtrl::MoveLast() 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	m_pDropDownRealWnd->MoveTo(m_pDropDownRealWnd->GetRowCount());
}

void CKCOMRichDropDownCtrl::MoveRecords(long Rows) 
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	ROWCOL nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
	m_pDropDownRealWnd->MoveTo(nRow + Rows);
}

void CKCOMRichDropDownCtrl::ReBind() 
{
	m_bDataSrcChanged = TRUE;
}

ROWCOL CKCOMRichDropDownCtrl::GetRowColFromVariant(VARIANT va)
{
	if (va.vt == VT_ERROR)
		return GX_INVALID;

	COleVariant vaNew = va;
	vaNew.ChangeType(VT_BSTR);
	vaNew.ChangeType(VT_I4);

	return vaNew.lVal;
}

BOOL CKCOMRichDropDownCtrl::OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray) 
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

BOOL CKCOMRichDropDownCtrl::OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut) 
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

ROWCOL CKCOMRichDropDownCtrl::GetRowFromBmk(BYTE *bmk)
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

ROWCOL CKCOMRichDropDownCtrl::GetRowFromHRow(HROW hRow)
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

void CKCOMRichDropDownCtrl::FetchNewRows()
{
	USES_CONVERSION;

	m_pDropDownRealWnd->CancelRecord();

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

void CKCOMRichDropDownCtrl::UndeleteRecordFromHRow(HROW hRow)
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

void CKCOMRichDropDownCtrl::GetBmkOfRow(ROWCOL nRow)
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

void CKCOMRichDropDownCtrl::OnBringMsg(long wParam, long lParam)
{
	BringInfo * pBi = (BringInfo *)wParam;
	if (!pBi)
		return;

	if (!pBi->hWnd)
	{
		// the comsumer tell me that I should close the dropdown window
		delete pBi;

		if (m_pDropDownRealWnd)
			HideDropDownWnd();

		m_bDroppedDown = FALSE;
		return;
	}

	WORD x = LOWORD(pBi->wParam);
	WORD y = HIWORD(pBi->wParam);
	x ^= 235;
	y ^= 235;

	m_strText = pBi->strText;
	m_hConsumer = pBi->hWnd;
	if (x + y == (pBi->lParam ^ 235) && !m_bDroppedDown)
		Bring(x, y);
		
	// free the memory allocated by message sender
	delete pBi;
}


void CKCOMRichDropDownCtrl::HideDropDownWnd(LPCTSTR pCharValue)
{
	if (::IsWindow(m_hConsumer))
	{
		BringInfo * pBi = new BringInfo;

		pBi->wParam = FALSE;
		pBi->hWnd = (HWND)0xff;
		if (pCharValue)
		{
			CString strValue = pCharValue;
			lstrcpy(pBi->strText, strValue);
			pBi->hWnd = NULL;
		}
		::SendMessage(m_hConsumer, CKCOMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
	}

	m_pDropDownRealWnd->Showwindow(FALSE);
}

void CKCOMRichDropDownCtrl::Bring(int cx, int cy)
{
	if(m_bDataSrcChanged)
		UpdateDropDownWnd();
	
	m_pDropDownRealWnd->SetScrollBarMode(SB_BOTH, gxnDisabled);

	int x = 0, y = 0, yWndMin = 0, yWndMax = 0, xTotal = 2, yTotal = 0;
	int nMaxWidth = GetSystemMetrics(SM_CXSCREEN) - cx;
	int nMaxHeight = GetSystemMetrics(SM_CYSCREEN) - cy;

	CRect rcBounds(cx, cy, 0, 0);

	if (m_bColumnHeaders)
	{
		yWndMin += m_nHeaderHeight;
		yWndMax += m_nHeaderHeight;
		yTotal += m_nHeaderHeight;
	}

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

	if (::IsWindow(m_hConsumer))
	{
		BringInfo * pBi = new BringInfo;
		pBi->hWnd = m_pDropDownRealWnd->m_hWnd;
		pBi->wParam = TRUE;
		::SendMessage(m_hConsumer, CKCOMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
	}

	SearchText(m_strText);

	if (m_pDropDownRealWnd->GetRowCount() > 0 && m_pDropDownRealWnd-> GetColCount() > 0)
	{
		m_pDropDownRealWnd->SetCurrentCell(m_nRow, 1);
		m_pDropDownRealWnd->SelectRange(CGXRange().SetTable(), FALSE);

		if (m_nRow > 0)
			m_pDropDownRealWnd->SelectRange(CGXRange().SetRows(m_nRow), TRUE);
		else
		{
			m_nRow = 1;
			m_pDropDownRealWnd->SelectRange(CGXRange().SetRows(m_nRow), TRUE);
		}
	}

	m_pDropDownRealWnd->SetWindowPos(&wndTopMost, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
	m_pDropDownRealWnd->SetForegroundWindow();
	m_pDropDownRealWnd->Showwindow(TRUE);
}

void CKCOMRichDropDownCtrl::UpdateDropDownWnd()
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

void CKCOMRichDropDownCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (m_pDropDownRealWnd && m_pDropDownRealWnd->IsWindowVisible())
		m_pDropDownRealWnd->ProcessKeys(this, WM_KEYDOWN, nChar, nRepCnt, nFlags);
	
	COleControl::OnKeyDown(nChar, nRepCnt, nFlags);
}

void CKCOMRichDropDownCtrl::DeleteRowFromSrc(ROWCOL nRow)
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

void CKCOMRichDropDownCtrl::SearchText(CString strSearch)
{
	if (!m_bListAutoPosition)
		return;

	FirePositionList(strSearch);

	for (int i = 0; i < m_apColumns.GetSize(); i ++)
	{
		if (!m_apColumns[i]->prop.strName.CompareNoCase(m_strDataField))
			break;
	}

	if (i >= m_apColumns.GetSize())
		return;

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
		
		return;
	}
	else
	{
		if (m_Rowset.m_spRowset == NULL)
			return;

		CComQIPtr<IRowsetFind, &IID_IRowsetFind> pFind = m_Rowset.m_spRowset;
		if (!pFind)
			return;

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
			return;

		ULONG nRowsOb = 0;
		HROW* phRow = &rac.m_hRow;
		pFind->FindNextRow(NULL, rac.m_pAccessor->GetHAccessor(0), rac.
			m_pAccessor->GetBuffer(), DBCOMPAREOPS_EQ | DBCOMPAREOPS_CASEINSENSITIVE, 
			m_nBookmarkSize, (BYTE *)&bmFirst, 0, 1, &nRowsOb, &phRow);

		if (rac.m_hRow == 0)
			pFind->FindNextRow(NULL, rac.m_pAccessor->GetHAccessor(0), rac.
				m_pAccessor->GetBuffer(), DBCOMPAREOPS_BEGINSWITH | DBCOMPAREOPS_CASEINSENSITIVE, 
				m_nBookmarkSize, (BYTE *)&bmFirst, 0, 1, &nRowsOb, &phRow);

		if (rac.m_hRow == 0)
			return;

		HROW hRow = rac.m_hRow;
		rac.ReleaseRows();
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

		m_nRow = nRow;

		return;
	}
}

void CKCOMRichDropDownCtrl::OnListAutoPositionChanged() 
{
	SetModifiedFlag();
	BoundPropertyChanged(dispidListAutoPosition);
}

void CKCOMRichDropDownCtrl::OnListWidthAutoSizeChanged() 
{
	SetModifiedFlag();
	BoundPropertyChanged(dispidListWidthAutoSize);
}

short CKCOMRichDropDownCtrl::GetMaxDropDownItems() 
{
	return m_nMaxDropDownItems;
}

void CKCOMRichDropDownCtrl::SetMaxDropDownItems(short nNewValue) 
{
	if (nNewValue < 1 || nNewValue > 100 || nNewValue < m_nMinDropDownItems)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDITEMSLIMITATION);

	m_nMaxDropDownItems = nNewValue;
	SetModifiedFlag();
	BoundPropertyChanged(dispidMaxDropDownItems);
}

short CKCOMRichDropDownCtrl::GetMinDropDownItems() 
{
	return m_nMinDropDownItems;
}

void CKCOMRichDropDownCtrl::SetMinDropDownItems(short nNewValue) 
{
	if (nNewValue < 1 || nNewValue > 100 || nNewValue > m_nMaxDropDownItems)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDITEMSLIMITATION);

	m_nMinDropDownItems = nNewValue;
	SetModifiedFlag();
	BoundPropertyChanged(dispidMinDropDownItems);
}

short CKCOMRichDropDownCtrl::GetListWidth() 
{
	return m_nListWidth;
}

void CKCOMRichDropDownCtrl::SetListWidth(short nNewValue) 
{
	if (nNewValue < 0)
		nNewValue = 0;

	m_nListWidth = nNewValue;
	SetModifiedFlag();
	BoundPropertyChanged(dispidListWidth);
}

OLE_COLOR CKCOMRichDropDownCtrl::GetBackColor() 
{
	return m_clrBack;
}

void CKCOMRichDropDownCtrl::SetBackColor(OLE_COLOR nNewValue) 
{
	m_clrBack = nNewValue;

	if (m_pDrawGrid)
		m_pDrawGrid->ChangeStandardStyle(CGXStyle().SetInterior(TranslateColor(GetBackColor())));
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->ChangeStandardStyle(CGXStyle().SetInterior(TranslateColor(GetBackColor())));
	
	SetModifiedFlag();
	BoundPropertyChanged(dispidBackColor);
	InvalidateControl();
}

void CKCOMRichDropDownCtrl::Bind()
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

void CKCOMRichDropDownCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_KCOMRICHDROPDOWN);
	dlgAbout.DoModal();
}

VARIANT CKCOMRichDropDownCtrl::GetBookmark() 
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

void CKCOMRichDropDownCtrl::SetBookmark(const VARIANT FAR& newValue) 
{
	if (!AmbientUserMode())
		SetNotSupported();

	SetRow(GetRowFromBmk(&newValue));
}

VARIANT CKCOMRichDropDownCtrl::RowBookmark(long RowIndex) 
{
	VARIANT vaResult;
	VariantInit(&vaResult);

	GetBmkOfRow(RowIndex, &vaResult);

	return vaResult;
}

void CKCOMRichDropDownCtrl::GetBmkOfRow(ROWCOL nRow, VARIANT *pVaRet)
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

ROWCOL CKCOMRichDropDownCtrl::GetRowFromIndex(ROWCOL nIndex)
{
	for (ROWCOL nRow = 1; nRow <= m_pDropDownRealWnd->GetRowCount(); nRow ++)
		if (m_pDropDownRealWnd->GetRowIndex(nRow) == nIndex)
			return nRow;

	return GX_INVALID;
}

ROWCOL CKCOMRichDropDownCtrl::GetRowFromBmk(const VARIANT *pBmk)
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
