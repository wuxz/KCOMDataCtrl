// GridInner.cpp : implementation file
//

#include "stdafx.h"
#include "KCOMDBGrid.h"
#include "GridInner.h"
#include "kcomdbgridctl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGridInner

CGridInner::CGridInner(CKCOMDBGridCtrl *pCtrl)
{
	m_pCtrl = pCtrl;
	m_bInEdit = FALSE;
	m_bDirty = FALSE;
	m_bInsertMode = FALSE;
	m_nRowDirty = -1;

	for (int i = 0; i < 255; i++)
		m_bColDirty[i] = FALSE;

	m_ImageList.Create(MAKEINTRESOURCE(IDB_BITMAP_HEADER), 12, 1, RGB(255, 255, 255));
	if (m_ImageList.m_hImageList)
		SetImageList(&m_ImageList);
}

CGridInner::~CGridInner()
{
	m_pCtrl = NULL;
}


BEGIN_MESSAGE_MAP(CGridInner, CGridCtrl)
	//{{AFX_MSG_MAP(CGridInner)
	ON_WM_KEYDOWN()
	ON_WM_CHAR()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_MBUTTONDOWN()
	ON_WM_MBUTTONUP()
	ON_WM_RBUTTONDOWN()
	ON_WM_RBUTTONUP()
	ON_WM_HSCROLL()
	ON_WM_VSCROLL()
	ON_WM_KILLFOCUS()
	ON_WM_SETFOCUS()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CGridInner message handlers

CString CGridInner::GetItemText(int nRow, int nCol)
{
	if (nRow > 0 && nCol > 0 && !(m_pCtrl->m_bAllowAddNew && nRow == GetRowCount() - 1))
		m_pCtrl->GetItemText(nRow, nCol);

	return CGridCtrl::GetItemText(nRow, nCol);
}

CCellID CGridInner::SetFocusCell(CCellID cell)
{
	CCellID cellCur = GetFocusCell();
	
	if (cell.row == 0 || (m_pCtrl->m_bShowRecordSelector && cell.col == 0) ||
		cell.row >= GetRowCount() || cell.col > GetColCount())
		return cellCur;

	SetSelectedRange(-1, -1, -1, -1);
	cellCur = GetFocusCell();
	CCellID cellRet;
	if (m_bDirty && m_nRowDirty != -1 && m_nRowDirty != cell.row && cell.row > 0)
	{
		if (FAILED(m_pCtrl->UpdateRowData(m_nRowDirty)))
			return cellCur;
		else
		{
			m_bDirty = m_bInsertMode = FALSE;
			for (int i = 0; i < 255; i++)
				m_bColDirty[i] = FALSE;
		}
	}

	if (cellCur.row > 0 && cellCur.col >= 0 && GetFixedColCount() && 
		!(m_pCtrl->m_bAllowAddNew && cellCur.row == GetRowCount() - 1))
	{
		SetItemImage(cellCur.row, 0, -1);
		RedrawCell(cellCur.row, 0);
	}

	cellRet = CGridCtrl::SetFocusCell(cell);
	CCellID cellNow = GetFocusCell();
	if (cellNow.row >= 0 && cellNow.col	>= 0)
	{
		BOOL bInsert = m_pCtrl->m_bAllowAddNew && cellNow.row == GetRowCount() - 1;

		SetItemImage(cellNow.row, 0, bInsert ? 2 : 0);
		RedrawCell(cellNow.row, 0);
		m_pCtrl->FireRolColChange();
//		m_pCtrl->SetFocusCell(cellNow.row, cellNow.col);
	}

	return cellRet;
}

void CGridInner::OnEndEditCell(int nRow, int nCol, CString str)
{
	m_bInEdit = FALSE;
	CString strCurrent = GetItemText(nRow,nCol);
	if (strCurrent != str)
	{
		m_pCtrl->FireAfterColEdit((short)nCol);

		VARIANT va;
		short bCancel = 0;
		VariantInit(&va);
		va.vt = VT_BSTR;
		va.bstrVal = str.AllocSysString();
		m_pCtrl->FireBeforeColUpdate((short)nCol, &va, &bCancel);
		if (bCancel)
			return;

		COleVariant vaResult;
		vaResult.Clear();
		vaResult.ChangeType(VT_BSTR, &va);
		USES_CONVERSION;
		str = OLE2T(vaResult.bstrVal);
		if (va.bstrVal)
			SysFreeString(va.bstrVal);

//		if (SUCCEEDED(m_pCtrl->ModifyCell(nRow, nCol, str)))
//		{
			SetItemText(nRow, nCol, str);
			m_pCtrl->FireAfterColUpdate((short)nCol);
			m_bDirty = TRUE;
			m_nRowDirty = nRow;
			m_bInsertMode = m_pCtrl->m_bAllowAddNew && nRow == GetRowCount() - 1;
			m_bColDirty[nCol - (m_pCtrl->m_bShowRecordSelector ? 1 : 0)] = TRUE;
//		}
	}

	int nRows = GetRowCount();
	SetItemImage(nRows - 1, 0, 2);
	RedrawCell(nRows - 1, 0);
	Invalidate();
				
	CCellID cell;
	cell.row = nRow;
	cell.col = nCol;
	SetFocusCell(cell);
}

void CGridInner::OnEditCell(int nRow, int nCol, UINT nChar)
{
	if (!m_pCtrl->IsColEditable(nCol))
		return;

	CCellID cellCur = GetFocusCell();
	if (cellCur.row >= 0 && cellCur.col >= 0 && GetFixedColCount() && !(m_pCtrl->
		m_bAllowAddNew && cellCur.row == GetRowCount() - 1))
	{
		SetItemImage(cellCur.row, 0, -1);
		RedrawCell(cellCur.row, 0);
	}
	if (!m_bInEdit && GetFixedColCount())
	{
		m_bInEdit = TRUE;
		SetItemImage(nRow, 0, 1);
		RedrawCell(nRow, 0);
	}
	CGridCtrl::OnEditCell(nRow, nCol, nChar);
}

void CGridInner::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (!IsValid(m_idCurrentCell)) 
	{
		CWnd::OnKeyDown(nChar, nRepCnt, nFlags);
		return;
	}
	
	if (nChar == VK_DELETE && !(m_pCtrl->m_bAllowAddNew && m_idCurrentCell.row == GetRowCount() - 1))
	{
		if (IsEditable() && m_idCurrentCell.row > 0 && GetSelectedCount() 
			== GetColCount() - GetFixedColCount() && m_pCtrl->m_bAllowDelete)
			m_pCtrl->DeleteRowData(m_idCurrentCell.row);

		return;
	}
	if (nChar == VK_DOWN && m_pCtrl->m_bAllowAddNew && m_bInsertMode &&
		m_bDirty && m_nRowDirty != -1 && m_nRowDirty == m_idCurrentCell.row 
		&& m_idCurrentCell.row > 0)
	{
		CCellID cell;
		cell.col = m_idCurrentCell.col;

		if (FAILED(m_pCtrl->UpdateRowData(m_nRowDirty)))
			return;
		else
		{
			m_bDirty = m_bInsertMode = FALSE;
			for (int i = 0; i < 255; i++)
				m_bColDirty[i] = FALSE;
		}

		cell.row = GetRowCount() - 1;
		SetFocusCell(cell);

		return;
	}
	if (nChar == VK_ESCAPE && m_bDirty && m_nRowDirty == m_idCurrentCell.row 
		&& m_idCurrentCell.row > 0 && GetFocus() == this)
	{
		if (m_pCtrl->m_apBookmark.GetSize() >= m_nRowDirty)
			m_pCtrl->UpdateRowFromSrc(m_nRowDirty - 1);
		else
		{
			HideInsertImage();
			InsertRow(m_nRowDirty);
			DeleteRow(m_nRowDirty + 1);
			m_nRowDirty = -1;
			m_bInsertMode = FALSE;
			ShowInsertImage();
			Invalidate();
		}
		return;
	}

	if (nChar == VK_RETURN)
		nChar = VK_RIGHT;

	CGridCtrl::OnKeyDown(nChar, nRepCnt, nFlags);
}

void CGridInner::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	CGridCtrl::OnChar(nChar, nRepCnt, nFlags);
}

void CGridInner::OnFixedColClick(CCellID &cell)
{
	if (cell.row > 0 && !(m_pCtrl->m_bAllowAddNew && cell.row == GetRowCount() - 1))
		SelectRow(cell.row);
}

void CGridInner::CreateInPlaceEditControl(CRect &rect, DWORD dwStyle, UINT nID, int nRow, int nCol, LPCTSTR szText, int nChar)
{
	short bCancel = 0;
	m_pCtrl->FireBeforeColEdit((short)nCol, (short)nChar, &bCancel);
	if (bCancel)
		return;

	CGridCtrl::CreateInPlaceEditControl(rect, dwStyle, nID, nRow, nCol, szText, nChar);
}

void CGridInner::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	if (m_pCtrl->AmbientUserMode())
		m_pCtrl->OnLButtonDblClk(nFlags, point);
	
	CGridCtrl::OnLButtonDblClk(nFlags, point);
}

void CGridInner::OnMouseMove(UINT nFlags, CPoint point) 
{
	if (m_pCtrl->AmbientUserMode())
		m_pCtrl->OnMouseMove(nFlags, point);
	
	CGridCtrl::OnMouseMove(nFlags, point);
}

void CGridInner::OnLButtonDown(UINT nFlags, CPoint point) 
{
	if (m_pCtrl->AmbientUserMode())
		m_pCtrl->OnLButtonDown(nFlags, point);
	
	CGridCtrl::OnLButtonDown(nFlags, point);
}

void CGridInner::OnLButtonUp(UINT nFlags, CPoint point) 
{
	CWnd::OnLButtonUp(nFlags, point);
	
	if (m_pCtrl->AmbientUserMode())
		m_pCtrl->OnLButtonUp(nFlags, point);

	if (m_MouseMode != MOUSE_SIZING_COL && m_MouseMode != MOUSE_SIZING_ROW)
	{
		CGridCtrl::OnLButtonUp(nFlags, point);
		return;
	}

	ClipCursor(NULL);
	
	if (GetCapture()->GetSafeHwnd() == GetSafeHwnd())
	{
		ReleaseCapture();
	}
	
	short bCancel = 0;
	CRect invertedRect;
	CPoint start;
	CRect rect;
	GetClientRect(rect);
	CDC* pDC = GetDC();

	if (m_MouseMode == MOUSE_SIZING_COL)
	{
		invertedRect.SetRect(m_LastMousePoint.x, rect.top, m_LastMousePoint.x + 2, rect.bottom);
		if (pDC)
		{
			pDC->InvertRect(&invertedRect);
			ReleaseDC(pDC);
		}
		
		if (m_LeftClickDownPoint != point) 
		{
			CPoint start;
			if (!GetCellOrigin(m_LeftClickDownCell, &start)) return;
			m_pCtrl->FireColResize(m_LeftClickDownCell.col, &bCancel);
			if (!bCancel)
			{
				SetColWidth(m_LeftClickDownCell.col, point.x - start.x);
				ResetScrollBars();
				Invalidate();
			}
		}		
	}
	else if (m_MouseMode == MOUSE_SIZING_ROW)
	{
		bCancel = m_pCtrl->GetAllowRowSizing() ? 0 : 1;
		invertedRect.SetRect(rect.left, m_LastMousePoint.y, rect.right, m_LastMousePoint.y + 2);
		if (pDC)
		{
			pDC->InvertRect(&invertedRect);
			ReleaseDC(pDC);
		}
		
		if (m_LeftClickDownPoint != point) 
		{
			if (!GetCellOrigin(m_LeftClickDownCell, &start)) return;
			m_pCtrl->FireRowResize(&bCancel);
			if (!bCancel)
			{
				SetRowHeight(m_LeftClickDownCell.row, point.y - start.y);
				ResetScrollBars();
				Invalidate();
			}
		}
	} 

	m_MouseMode = MOUSE_NOTHING;
	
	SetCursor(AfxGetApp()->LoadStandardCursor(IDC_ARROW));
	
	if (!IsValid(m_LeftClickDownCell)) return;
	
	CWnd *pOwner = GetOwner();
	if (pOwner && IsWindow(pOwner->m_hWnd))
		pOwner->PostMessage(WM_COMMAND, MAKELONG(GetDlgCtrlID(), BN_CLICKED), 
		(LPARAM) GetSafeHwnd());
}

void CGridInner::OnMButtonDown(UINT nFlags, CPoint point) 
{
	if (m_pCtrl->AmbientUserMode())
		m_pCtrl->OnMButtonDown(nFlags, point);
	
	CGridCtrl::OnMButtonDown(nFlags, point);
}

void CGridInner::OnMButtonUp(UINT nFlags, CPoint point) 
{
	if (m_pCtrl->AmbientUserMode())
		m_pCtrl->OnMButtonUp(nFlags, point);
	
	CGridCtrl::OnMButtonUp(nFlags, point);
}

void CGridInner::OnRButtonDown(UINT nFlags, CPoint point) 
{
	if (m_pCtrl->AmbientUserMode())
		m_pCtrl->OnRButtonDown(nFlags, point);
	
	CGridCtrl::OnRButtonDown(nFlags, point);
}

void CGridInner::OnRButtonUp(UINT nFlags, CPoint point) 
{
	if (m_pCtrl->AmbientUserMode())
		m_pCtrl->OnRButtonUp(nFlags, point);
	
	CGridCtrl::OnRButtonUp(nFlags, point);
}

void CGridInner::OnFixedRowClick(CCellID &cell)
{
	m_pCtrl->FireHeadClick((short)cell.col);

	CGridCtrl::OnFixedRowClick(cell);
}

void CGridInner::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	short bCancel = 0;

	m_pCtrl->FireScroll(&bCancel);
	if (bCancel)
		return;
	
	CGridCtrl::OnHScroll(nSBCode, nPos, pScrollBar);
}

void CGridInner::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	short bCancel = 0;

	m_pCtrl->FireScroll(&bCancel);
	if (bCancel)
		return;
	
	CGridCtrl::OnVScroll(nSBCode, nPos, pScrollBar);
}


void CGridInner::ShowInsertImage()
{
	int nRows = GetRowCount();

	if (m_pCtrl->m_bAllowAddNew)
	{
		SetItemImage(nRows - 1, 0, 2);
		RedrawCell(nRows - 1, 0);
	}
}

void CGridInner::HideInsertImage()
{
	int nRows = GetRowCount();

	if (m_pCtrl->m_bAllowAddNew)
	{
		SetItemImage(nRows - 1, 0, -1);
		RedrawCell(nRows - 1, 0);
	}
}

void CGridInner::OnKillFocus(CWnd* pNewWnd) 
{
	m_pCtrl->OnKillFocus(pNewWnd);

	CGridCtrl::OnKillFocus(pNewWnd);
}

void CGridInner::OnSetFocus(CWnd* pOldWnd) 
{
	m_pCtrl->OnSetFocus(pOldWnd);

	CGridCtrl::OnSetFocus(pOldWnd);
}
