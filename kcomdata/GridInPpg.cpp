// GridInPpg.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "GridInPpg.h"
#include "KCOMRichGridColumnPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGridInPpg

CGridInPpg::CGridInPpg()
{
	m_bLButtonDown = FALSE;
}

CGridInPpg::~CGridInPpg()
{
}

CGridInPpg::SetPagePtr(CKCOMRichGridPropColumnPage * pPage)
{
	m_pPage = pPage;
}

BEGIN_MESSAGE_MAP(CGridInPpg, CGXBrowserWnd)
	//{{AFX_MSG_MAP(CGridInPpg)
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CGridInPpg message handlers

void CGridInPpg::SetAlignment(ROWCOL nCol, long nNewValue)
{
	DWORD dwAlign;

	switch (nNewValue)
	{
	case 0:
		dwAlign = DT_LEFT;
		break;

	case 1:
		dwAlign = DT_RIGHT;
		break;

	case 2:
		dwAlign = DT_CENTER;
		break;
	}

	SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetHorizontalAlignment(
		dwAlign));

	RedrawRowCol(CGXRange().SetCols(nCol));
}

void CGridInPpg::SetForeColor(ROWCOL nCol, OLE_COLOR nNewValue)
{
	if (nNewValue == DEFAULTCOLOR)
		return;

	COLORREF clrFore;

	OleTranslateColor(nNewValue, NULL, &clrFore);
	SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetTextColor(
		clrFore));

	RedrawRowCol(CGXRange().SetCols(nCol));
}

void CGridInPpg::SetBackColor(ROWCOL nCol, OLE_COLOR nNewValue)
{
	COLORREF clrBack;

	if (nNewValue == DEFAULTCOLOR)
		return;

	OleTranslateColor(nNewValue, NULL, &clrBack);
	SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetInterior(
		clrBack));

	RedrawRowCol(CGXRange().SetCols(nCol));
}

void CGridInPpg::SetHeadForeColor(ROWCOL nCol, OLE_COLOR nNewValue)
{
	if (nNewValue == DEFAULTCOLOR)
		return;

	COLORREF clrHeadFore;

	OleTranslateColor(nNewValue, NULL, &clrHeadFore);
	SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetTextColor(clrHeadFore));

	RedrawRowCol(CGXRange().SetCols(nCol));
}

void CGridInPpg::SetHeadBackColor(ROWCOL nCol, OLE_COLOR nNewValue)
{
	if (nNewValue == DEFAULTCOLOR)
		return;

	COLORREF clrHeadBack;

	OleTranslateColor(nNewValue, NULL, &clrHeadBack);
	SetStyleRange(CGXRange(0, nCol), CGXStyle().SetInterior(clrHeadBack));

	RedrawRowCol(CGXRange().SetCols(nCol));
}

void CGridInPpg::SetLocked(ROWCOL nCol, BOOL bNewValue)
{
	SetStyleRange(CGXRange(0, nCol), CGXStyle().SetReadOnly(bNewValue));
}

void CGridInPpg::SetVisible(ROWCOL nCol, BOOL bNewValue)
{
	HideCols(nCol, nCol, !bNewValue);
}

void CGridInPpg::SetWidth(ROWCOL nCol, long nNewValue)
{
	SetColWidth(nCol, nCol, nNewValue);

	RedrawRowCol(CGXRange().SetCols(nCol));
}

// if the datatype is date, we should use the hidden date editbox style cell
// else if the mask is set, we should use the hidden mask editbox style cell
// so the given style takes no effect
void CGridInPpg::SetStyle(ROWCOL nCol, long nNewValue)
{
	switch (nNewValue)
	{
	case 0:
		SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
			GX_IDS_CTRL_EDIT));
		break;

	case 1:
		SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
			KCOM_IDS_CTRL_EDITBUTTON));
		break;

	case 2:
		SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
			GX_IDS_CTRL_CHECKBOXEX));
		break;

	case 3:
		SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
			GX_IDS_CTRL_COMBOBOX));
		break;

	case 4:
		SetStyleRange(CGXRange().SetCols(nCol), CGXStyle().SetControl(
			GX_IDS_CTRL_PUSHBTN));
		break;
	}

	RedrawRowCol(CGXRange().SetCols(nCol));
}

void CGridInPpg::SetCaptionAlignment(ROWCOL nCol, long nNewValue)
{
	DWORD dwAlign;

	switch (nNewValue)
	{
	case 0:
		dwAlign = DT_LEFT;
		break;

	case 1:
		dwAlign = DT_RIGHT;
		break;

	case 2:
		dwAlign = DT_CENTER;
		break;
	}

	ColHeaderStyle().SetHorizontalAlignment(dwAlign);

	RedrawRowCol(CGXRange().SetCols(nCol));
}

void CGridInPpg::SetCaption(ROWCOL nCol, LPCTSTR lpszNewValue)
{
	SetValueRange(CGXRange(0, nCol), lpszNewValue);

	RedrawRowCol(CGXRange().SetCols(nCol));
}

ROWCOL CGridInPpg::GetColFromIndex(ROWCOL nColIndex)
{
	for (ROWCOL i = 1; i <= GetColCount(); i ++)
		if (GetColIndex(i) == nColIndex)
			return i;

	return GX_INVALID;
}

BOOL CGridInPpg::OnLButtonClickedRowCol(ROWCOL nRow, ROWCOL nCol, UINT nFlags, CPoint pt)
{
	BOOL bRet = CGXBrowserWnd::OnLButtonClickedRowCol(nRow, nCol, nFlags, pt);

	int nField = GetFieldFromIndex(GetColIndex(nCol));
	if (m_pPage->m_nCurrentCol == nField)
		return bRet;

	((CComboBox *)m_pPage->GetDlgItem(IDC_COMBO_COLUMN))->SetCurSel(nField);
	m_pPage->OnSelendokComboColumn();
	m_pPage->OnSelchangeListProperty();

	return bRet;
}

// calc the ordinal in the array
ROWCOL CGridInPpg::GetFieldFromIndex(ROWCOL nColIndex)
{
	for (int i = 0; i < m_pPage->m_apColumnProp.GetSize(); i ++)
		if (m_pPage->m_apColumnProp[i]->nColIndex == nColIndex)
			return i;

	return GX_INVALID;
}

void CGridInPpg::OnModifyCell(ROWCOL nRow, ROWCOL nCol)
{
	if (m_bCanDirty && !m_pPage->IsModified())
		m_pPage->SetModifiedFlag(TRUE);

	CGXBrowserWnd::OnModifyCell(nRow, nCol);
}

BOOL CGridInPpg::StoreColWidth(ROWCOL nCol, int nWidth)
{
	if (!CGXBrowserWnd::StoreColWidth(nCol, nWidth))
		return FALSE;

	if (nCol == 0)
		return TRUE;

	ColumnProp * pCol = m_pPage->m_apColumnProp[GetFieldFromIndex(
		GetColIndex(nCol))];
	
	if (m_bLButtonDown)
		return TRUE;

	pCol->nWidth = nWidth;
	m_pPage->OnSelchangeListProperty();
	m_pPage->SetModifiedFlag(TRUE);

	return TRUE;
}

BOOL CGridInPpg::StoreRemoveRows(ROWCOL nFromRow, ROWCOL nToRow, BOOL bProcessed)
{
	if (!CGXBrowserWnd::StoreRemoveRows(nFromRow, nToRow, bProcessed))
		return FALSE;

	m_pPage->SetModifiedFlag(TRUE);

	return TRUE;
}

void CGridInPpg::OnLButtonDown(UINT nFlags, CPoint point) 
{
	m_bLButtonDown = TRUE;

	CGXBrowserWnd::OnLButtonDown(nFlags, point);
}

void CGridInPpg::OnLButtonUp(UINT nFlags, CPoint point) 
{
	m_bLButtonDown = FALSE;

	CGXBrowserWnd::OnLButtonUp(nFlags, point);
}

void CGridInPpg::OnGridInitialUpdate()
{
	CGXBrowserWnd::OnGridInitialUpdate();
	InitBrowserSettings();
}
