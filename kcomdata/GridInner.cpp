// GridInner.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "GridInner.h"
#include "KCOMRichGridCtl.h"
#include "GridColumns.h"
#include "GridColumn.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGridInner

CGridInner::CGridInner(CKCOMRichGridCtrl * pCtrl)
{
	m_pCtrl = pCtrl;
}

CGridInner::~CGridInner()
{
}


BEGIN_MESSAGE_MAP(CGridInner, CGXBrowserWnd)
	//{{AFX_MSG_MAP(CGridInner)
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_MOUSEMOVE()
	ON_WM_MBUTTONDBLCLK()
	ON_WM_MBUTTONDOWN()
	ON_WM_RBUTTONDBLCLK()
	ON_WM_RBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_MBUTTONUP()
	ON_WM_RBUTTONUP()
	ON_WM_KILLFOCUS()
	ON_WM_SETFOCUS()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CGridInner message handlers

ROWCOL CGridInner::GetColFromIndex(ROWCOL nColIndex)
{
	for (ROWCOL i = 1; i <= GetColCount(); i ++)
		if (GetColIndex(i) == nColIndex)
			return i;

	return GX_INVALID;
}

void CGridInner::OnGridInitialUpdate()
{
	CGXBrowserWnd::OnGridInitialUpdate();
	
	if (!m_pCtrl)
		return;

	// let the 3 lovely little icons appear in the row header
	InitBrowserSettings();

	GetParam()->SetNewGridLineMode(TRUE);
}

BOOL CGridInner::OnFlushRecord(ROWCOL nRow, CString *ps)
{
	if (!m_pCtrl)
		return CGXBrowserWnd::OnFlushRecord(nRow, ps);

	short bCancel = 0;
	
	if (m_pCtrl->m_nDataMode != 0)
	{
		m_pCtrl->FireBeforeUpdate(&bCancel);
		if (bCancel)
			return FALSE;

		if (!CGXBrowserWnd::OnFlushRecord(nRow, ps))
			return FALSE;
	}

	short bDispMsg = 0;
	if (m_pCtrl->m_nDataMode != 0)
	{
		if (GetBrowseParam()->m_nEditMode == addnew)
		{
			m_pCtrl->FireAfterInsert(&bDispMsg);
			if (bDispMsg)
				AfxMessageBox(IDS_AFTERINSERT);
		}
		else
		{
			m_pCtrl->FireAfterUpdate(&bDispMsg);
			if (bDispMsg)
				AfxMessageBox(IDS_AFTERUPDATE);
		}
		
		return TRUE;
	}
	else
	{
		CGXBrowseParam* pBrowseData = GetBrowseParam();
		
		short nField;
		ROWCOL nCols = GetColCount();
		m_arContent.SetSize(nCols);

		// fill the modification information to update the datasource
		for (ROWCOL i = 0; i < GetColCount(); i ++)
		{
			m_bColDirty[i] = pBrowseData->m_paDirtyFieldStruct[GetFieldFromCol(i + 1)].bDirty;

			nField = GetFieldFromCol(i + 1);
			if (nField >= 0 && pBrowseData->m_paDirtyFieldStruct[nField].bDirty)
				m_arContent[i] = pBrowseData->m_paDirtyFieldStruct[nField].sValue;
			else
				m_arContent[i].Empty();
		}

		if (FAILED(m_pCtrl->UpdateRowData(nRow)))
			return FALSE;
		else
			return CGXBrowserWnd::OnFlushRecord(nRow, ps);
	}

	GetParam()->EmptyUndoList();
	return TRUE;
}

void CGridInner::DeleteRecords()
{
	if (!m_pCtrl)
	{
		CGXBrowserWnd::DeleteRecords();
		return;
	}

	CGXRangeList* pSelList = GetParam()->GetRangeList();

	if (!m_pCtrl->m_bAllowDelete || IsReadOnly() || !CanUpdate())
		return;

	if (m_pCtrl->m_nDataMode)
	{
		short bCancel = 0, bDispMsg = 0;

		m_pCtrl->FireBeforeDelete(&bCancel, &bDispMsg);
		if (bCancel)
			return;
		if (bDispMsg && AfxMessageBox(IDS_BEFOREDELETE, MB_YESNO) != IDYES)
			return;

		CGXBrowserWnd::DeleteRecords();

		bDispMsg = 0;
		m_pCtrl->FireAfterDelete(&bDispMsg);
		if (bDispMsg)
			AfxMessageBox(IDS_AFTERDELETE);

		return;
	}

	// check that whole rows are selected and that "new row" is not selected
	if (!pSelList->IsAnyCellFromCol(0)
		|| pSelList->IsAnyCellFromRow( GetAppendRow() )
		)
		return;

	CRowColArray awRow;
	GetSelectedRows(awRow, TRUE, FALSE);

	m_pCtrl->DeleteRowData(awRow[0]);
}

// customize the hittest to avoid resizing this column if necessary
int CGridInner::HitTest(CPoint &pt, ROWCOL *pnRow, ROWCOL *pnCol, CRect *pRectHit)
{
	if (!m_pCtrl)
		return CGXBrowserWnd::HitTest(pt, pnRow, pnCol, pRectHit);

	ROWCOL nRow = GX_INVALID, nCol = GX_INVALID;

	int nRet = CGXBrowserWnd::HitTest(pt, &nRow, &nCol, pRectHit);
	
	if (pnRow)
		*pnRow = nRow;
	if (pnCol)
		*pnCol = nCol;

	if (nCol < 1 || nCol > (ROWCOL)m_pCtrl->m_apColumns.GetSize())
		return nRet;

	if (nRet == GX_VERTLINE && nCol > 0 && !m_pCtrl->m_apColumns[nCol - 1]
		->prop.bAllowSizing)
		nRet = GX_NOHIT;

	return nRet;
}

BOOL CGridInner::IsColDirty(ROWCOL nCol)
{
	CGXBrowseParam* pBrowseData = GetBrowseParam();

	if (pBrowseData->m_paDirtyFieldStruct == NULL || nCol >= (ROWCOL)pBrowseData
		->m_nFields)
		return FALSE;

	return pBrowseData->m_paDirtyFieldStruct[nCol - 1].bDirty;
}

// catch these messages only for firing events :(
void CGridInner::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnLButtonDblClk(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDblClick(LEFT_BUTTON, point);
}

void CGridInner::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnLButtonDown(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDown(LEFT_BUTTON, point);
}

void CGridInner::OnMouseMove(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnMouseMove(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->OnMouseMove(nFlags, point);
}

void CGridInner::OnMButtonDblClk(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnMButtonDblClk(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDblClick(MIDDLE_BUTTON, point);
}

void CGridInner::OnMButtonDown(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnMButtonDown(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDown(MIDDLE_BUTTON, point);
}

void CGridInner::OnRButtonDblClk(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnRButtonDblClk(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDblClick(RIGHT_BUTTON, point);
}

void CGridInner::OnRButtonDown(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnRButtonDown(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDown(RIGHT_BUTTON, point);
}

void CGridInner::OnLButtonUp(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnLButtonUp(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonUp(LEFT_BUTTON, point);
}

void CGridInner::OnMButtonUp(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnMButtonUp(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonUp(MIDDLE_BUTTON, point);
}

void CGridInner::OnRButtonUp(UINT nFlags, CPoint point) 
{
	CGXBrowserWnd::OnRButtonUp(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonUp(RIGHT_BUTTON, point);
}

// sync the current row postion of the datasource
BOOL CGridInner::SetCurrentCell(ROWCOL nRow, ROWCOL nCol, UINT flags)
{
	ROWCOL nRowCurr, nColCurr;
	GetCurrentCell(nRowCurr, nColCurr);

	m_pCtrl->HideDropDownWnd();

	if (!CGXBrowserWnd::SetCurrentCell(nRow, nCol, flags))
		return FALSE;

	if (nRow != nRowCurr && m_pCtrl)
		m_pCtrl->SetCurrentCell(nRow, nCol);

	return TRUE;
}

// give the grid a chance to pop up the dropdown window if this column
// is bounded to outer dropdown control to act like a combo box
void CGridInner::OnClickedButtonRowCol(ROWCOL nRow, ROWCOL nCol)
{
	CGXBrowserWnd::OnClickedButtonRowCol(nRow, nCol);

	m_pCtrl->OnDropDown(nRow, nCol);

	if (m_pCtrl)
		m_pCtrl->FireBtnClick();
}

// override this function to fire events
BOOL CGridInner::StoreStyleRowCol(ROWCOL nRow, ROWCOL nCol, const CGXStyle* pStyle, GXModifyType mt, int nType)
{
	if (!m_pCtrl)
		return CGXBrowserWnd::StoreStyleRowCol(nRow, nCol, pStyle, mt, nType);

	short bCancel = 0;
	
	if (nRow > 0 && nCol > 0 && pStyle->GetIncludeValue())
	{
		CString strValue = GetValueRowCol(nRow, nCol);

		m_pCtrl->FireBeforeColUpdate((short)nCol, strValue, &bCancel);
		if (bCancel)
		{
			EmptyEditBuffer();
			TransferCurrentCell(FALSE, GX_UPDATENOW, FALSE);
			return FALSE;
		}
	}

	if (!CGXBrowserWnd::StoreStyleRowCol(nRow, nCol, pStyle, mt, nType))
		return FALSE;

	if (pStyle->GetIncludeValue())
		m_pCtrl->FireAfterColUpdate((short)nCol);

	return TRUE;
}

// just for firing events
ROWCOL CGridInner::AddNew()
{
	if (!m_pCtrl)
		return CGXBrowserWnd::AddNew();

	short bCancel = 0;
	short bDispMsg = 0;

	m_pCtrl->FireBeforeInsert(&bCancel);
	if (bCancel)
	{
		ROWCOL nEditRow = GetBrowseParam()->m_nCurrentRow;

		GetBrowseParam()->m_nEditMode = noMode;
		EmptyEditBuffer();
		TransferCurrentCell(FALSE, GX_UPDATENOW, FALSE);
		RedrawRowCol(CGXRange().SetRows(nEditRow), GX_INVALIDATE);
		return GX_INVALID;
	}

	ROWCOL nRet;
	
	nRet = CGXBrowserWnd::AddNew();

	return nRet;
}

// just for firing events
BOOL CGridInner::OnStartTracking(ROWCOL nRow, ROWCOL nCol, int nTrackingMode)
{
	if (!m_pCtrl)
		return CGXBrowserWnd::OnStartTracking(nRow, nCol, nTrackingMode);

	short bCancel = 0;

	if (nTrackingMode == GX_TRACKWIDTH)
		m_pCtrl->FireColResize((short)nCol, &bCancel);
	else
		m_pCtrl->FireRowResize(&bCancel);

	if (bCancel)
		return FALSE;

	return CGXBrowserWnd::OnStartTracking(nRow, nCol, nTrackingMode);
}

// events! events! events!
BOOL CGridInner::MoveCurrentCell(int direction, UINT nCell, BOOL bCanSelectRange)
{
	if (!m_pCtrl)
		return CGXBrowserWnd::MoveCurrentCell(direction, nCell, bCanSelectRange);

	short bCancel = 0;

	m_pCtrl->FireBeforeRowColChange(&bCancel);
	if (bCancel)
		return FALSE;

	BOOL bRet = CGXBrowserWnd::MoveCurrentCell(direction, nCell, bCanSelectRange);
	if (bRet)
		m_pCtrl->FireRowColChange();

	return bRet;
}

BOOL CGridInner::OnLButtonHitRowCol(ROWCOL nHitRow, ROWCOL nHitCol, ROWCOL nDragRow, ROWCOL nDragCol, CPoint point, UINT flags, WORD nHitState)
{
	if (!m_pCtrl)
		return CGXBrowserWnd::OnLButtonHitRowCol(nHitRow, nHitCol, nDragRow, nDragCol, point, flags, nHitState);

	short bCancel = 0;

	if (nHitState & GX_HITINCELL)
	{
		m_pCtrl->FireBeforeRowColChange(&bCancel);
		if (bCancel)
			return FALSE;
	}

	BOOL bRet = CGXBrowserWnd::OnLButtonHitRowCol(nHitRow, nHitCol, nDragRow, nDragCol, point, flags, nHitState);
	if (bRet && (nHitState & GX_HITINCELL))
		m_pCtrl->FireRowColChange();

	return bRet;
}

BOOL CGridInner::DoScroll(int direction, ROWCOL nCell)
{
	if (!m_pCtrl)
		return CGXBrowserWnd::DoScroll(direction, nCell);

	short bCancel = 0;

	m_pCtrl->FireScroll(&bCancel);
	if (bCancel)
		return FALSE;

	if (CGXBrowserWnd::DoScroll(direction, nCell))
	{
		m_pCtrl->FireScrollAfter();
		
		return TRUE;
	}

	return FALSE;
}

void CGridInner::OnModifyCell(ROWCOL nRow, ROWCOL nCol)
{
	CGXBrowserWnd::OnModifyCell(nRow, nCol);

	if (m_pCtrl)
		m_pCtrl->FireChange();
}

BOOL CGridInner::StoreMoveCols(ROWCOL nFromCol, ROWCOL nToCol, ROWCOL nDestCol, BOOL bProcessed)
{
	short bCancel = 0;

	if (m_pCtrl)
	{
		m_pCtrl->FireColSwap((short)nFromCol, (short)nToCol, (short)nDestCol, &bCancel);
		if (bCancel)
			return FALSE;
	}

	return CGXBrowserWnd::StoreMoveCols(nFromCol, nToCol, nDestCol, bProcessed);
}

// update the cell content in bound mode
BOOL CGridInner::OnLoadCellStyle(ROWCOL nRow, ROWCOL nCol, CGXStyle &style, LPCTSTR pszExistingValue)
{
	if (!m_pCtrl)
		return FALSE;

	if (nRow > 0 && nCol > 0 && nRow <= (ROWCOL)OnGetRecordCount() && !pszExistingValue)
		m_pCtrl->GetCellValue(nRow, nCol, style);

	return FALSE;
}

// cancel the modification, if this row is a "addnew" row, delete this row
// and update the bookmark array
void CGridInner::CancelRecord()
{
	if (!m_pCtrl)
	{
		CGXBrowserWnd::CancelRecord();
		return;
	}

	if (m_pCtrl->m_nDataMode)
	{
		CGXBrowserWnd::CancelRecord();
		return;
	}

	CGXBrowseParam* pBrowseData = GetBrowseParam();
	ROWCOL nRow = pBrowseData->m_nCurrentRow;

	if (pBrowseData->m_nEditMode == addnew && (ROWCOL)m_pCtrl->m_apBookmark.GetSize() >= nRow)
	{
		delete [] m_pCtrl->m_apBookmark[nRow - 1];
		m_pCtrl->m_apBookmark.RemoveAt(nRow - 1);
	}

	CGXBrowserWnd::CancelRecord();
}

BOOL CGridInner::IsAddNewMode()
{
	return GetBrowseParam()->m_nEditMode == addnew;
}

ROWCOL CGridInner::GetRecordCount()
{
	return OnGetRecordCount() + (GetBrowseParam()->m_nEditMode == addnew ? 1 : 0);
}

void CGridInner::OnKillFocus(CWnd* pNewWnd) 
{
	if (m_pCtrl && !IsChild(pNewWnd))
		m_pCtrl->OnKillFocus(pNewWnd);

	CGXBrowserWnd::OnKillFocus(pNewWnd);
}

void CGridInner::OnSetFocus(CWnd* pOldWnd) 
{
	if (m_pCtrl)
		m_pCtrl->OnSetFocus(pOldWnd);

	CGXBrowserWnd::OnSetFocus(pOldWnd);
}

// two functions used for dropdown style cell
BOOL CGridInner::OnGridKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == VK_F4)
	{
		ROWCOL nRow, nCol;
		GetCurrentCell(nRow, nCol);
		m_pCtrl->OnDropDown(nRow, nCol);
		
		return TRUE;
	}

	return CGXBrowserWnd::OnGridKeyDown(nChar, nRepCnt, nFlags);
}

BOOL CGridInner::OnGridSysKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == VK_DOWN || nChar == VK_UP)
	{
		ROWCOL nRow, nCol;
		GetCurrentCell(nRow, nCol);
		m_pCtrl->OnDropDown(nRow, nCol);
		
		return TRUE;
	}

	return CGXBrowserWnd::OnGridSysKeyDown(nChar, nRepCnt, nFlags);
}

BOOL CGridInner::Moveto(ROWCOL nRow)
{
	if (GetColCount() == 0 || nRow > GetRowCount())
		return FALSE;

	ROWCOL nCurrRow, nCurrCol;
	GetCurrentCell(nCurrRow, nCurrCol);
	if (nCurrCol > GetColCount())
		nCurrCol = 1;

	return SetCurrentCell(nRow, nCurrCol);
}
