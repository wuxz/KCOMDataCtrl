// DropDownRealWnd.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "DropDownRealWnd.h"
#include "KCOMRichDropDownCtl.h"
#include "KCOMRichComboCtl.h"
#include "DropDownColumn.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDropDownRealWnd

CDropDownRealWnd::CDropDownRealWnd(CKCOMRichDropDownCtrl * pCtrl)
{
	m_pDropDownCtrl = pCtrl;
	m_pComboCtrl = NULL;
	m_bRefreshOnSetCurrentCell = FALSE;
}

CDropDownRealWnd::CDropDownRealWnd(CKCOMRichComboCtrl * pCtrl)
{
	m_pDropDownCtrl = NULL;
	m_pComboCtrl = pCtrl;
	m_bRefreshOnSetCurrentCell = FALSE;
}

CDropDownRealWnd::~CDropDownRealWnd()
{
}


BEGIN_MESSAGE_MAP(CDropDownRealWnd, CGXBrowserWnd)
	//{{AFX_MSG_MAP(CDropDownRealWnd)
	ON_WM_KILLFOCUS()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CDropDownRealWnd message handlers

ROWCOL CDropDownRealWnd::GetColFromIndex(ROWCOL nColIndex)
{
	for (ROWCOL i = 1; i <= GetColCount(); i ++)
		if (GetColIndex(i) == nColIndex)
			return i;

	return GX_INVALID;
}

BOOL CDropDownRealWnd::OnLoadCellStyle(ROWCOL nRow, ROWCOL nCol, CGXStyle &style, LPCTSTR pszExistingValue)
{
	if (!m_pDropDownCtrl && !m_pComboCtrl)
		return FALSE;

	if (nRow > 0 && nCol > 0 && nRow <= (ROWCOL)OnGetRecordCount() && !pszExistingValue)
	{
		if (m_pDropDownCtrl)
			m_pDropDownCtrl->GetCellValue(nRow, nCol, style);
		else if (m_pComboCtrl)
			m_pComboCtrl->GetCellValue(nRow, nCol, style);
	}

	return FALSE;
}

// this funcition is used in dropdown control only
void CDropDownRealWnd::Showwindow(BOOL bShow)
{
	if (m_pDropDownCtrl)
	{
		if (bShow)
		{
			m_pDropDownCtrl->FireDropDown();
			ShowWindow(SW_SHOW);
			m_pDropDownCtrl->m_bDroppedDown = TRUE;
		}
		else
		{
			ShowWindow(SW_HIDE);
			m_pDropDownCtrl->m_bDroppedDown = FALSE;
			m_pDropDownCtrl->FireCloseUp();
		}
	}
}

void CDropDownRealWnd::OnKillFocus(CWnd* pNewWnd) 
{
	CGXBrowserWnd::OnKillFocus(pNewWnd);
	
	if (pNewWnd == this)
		return;
	
	if (::IsWindowVisible(m_hWnd))
	{
		if (m_pDropDownCtrl)
		{
			m_pDropDownCtrl->HideDropDownWnd();
			
			if (::IsWindow(m_pDropDownCtrl->m_hConsumer))
			{
				BringInfo * pBi = new BringInfo;
				pBi->hWnd = (HWND)0xff;
				pBi->wParam = FALSE;
				::PostMessage(m_pDropDownCtrl->m_hConsumer, CKCOMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
			}
		}
		else if (m_pComboCtrl)
		{
			m_pComboCtrl->HideDropDownWnd();
		}
	}
}

BOOL CDropDownRealWnd::OnLButtonClickedRowCol(ROWCOL nRow, ROWCOL nCol, UINT nFlags, CPoint pt)
{
	if (m_pDropDownCtrl)
		m_pDropDownCtrl->FireClick();
	else if (m_pComboCtrl)
		m_pComboCtrl->FireClick();

	if (nRow == 0 || nCol == 0)
		return	FALSE;

	ReturnText(nRow);
	if (m_pComboCtrl)
		m_pComboCtrl->HideDropDownWnd();

	return TRUE;
}

BOOL CDropDownRealWnd::ProcessKeys(CWnd *pSender, UINT nMessage, UINT nChar, UINT nRepCnt, UINT flags)
{
	if (nMessage == WM_KEYDOWN)
	{
		switch (nChar)
		{
		case VK_RETURN:
		{
			CRowColArray awLeft, awRight;

			if (GetSelectedRows(awLeft, awRight, TRUE, FALSE))
				ReturnText(awLeft[0]);
		}

		case VK_ESCAPE:
		{
			if (m_pDropDownCtrl)
				m_pDropDownCtrl->HideDropDownWnd();
			else if (m_pComboCtrl)
				m_pComboCtrl->HideDropDownWnd();

			return TRUE;
		}

		break;
		}
	}

	return CGXBrowserWnd::ProcessKeys(pSender, nMessage, nChar, nRepCnt, flags);
}

// return the choosed text back to host control
void CDropDownRealWnd::ReturnText(ROWCOL nRow)
{
	if (m_pDropDownCtrl)
	{
		m_pDropDownCtrl->m_nRow = nRow;
		CString strValue;
		ROWCOL col = 0;
		
		for (int i = 0; i < m_pDropDownCtrl->m_apColumns.GetSize(); i ++)
		{
			if (!m_pDropDownCtrl->m_apColumns[i]->prop.strName.CompareNoCase(
				m_pDropDownCtrl->m_strDataField))
				break;
		}
		
		if (i < m_pDropDownCtrl->m_apColumns.GetSize())
		{
			col = GetColFromIndex(m_pDropDownCtrl->m_apColumns[i]->prop.nColIndex);
			strValue = GetValueRowCol(nRow, col);
			m_pDropDownCtrl->HideDropDownWnd(strValue);
		}
	}
	else if (m_pComboCtrl)
	{
		m_pComboCtrl->m_nRow = nRow;
		CString strValue;
		ROWCOL col = 0;
		
		for (int i = 0; i < m_pComboCtrl->m_apColumns.GetSize(); i ++)
		{
			if (!m_pComboCtrl->m_apColumns[i]->prop.strName.CompareNoCase(
				m_pComboCtrl->m_strDataField))
				break;
		}
		
		if (i < m_pComboCtrl->m_apColumns.GetSize())
		{
			col = GetColFromIndex(m_pComboCtrl->m_apColumns[i]->prop.nColIndex);
			strValue = GetValueRowCol(nRow, col);
			m_pComboCtrl->SetText(strValue);
		}
	}
}

BOOL CDropDownRealWnd::DoScroll(int direction, ROWCOL nCell)
{
	short bCancel = 0;

	if (m_pDropDownCtrl)
		m_pDropDownCtrl->FireScroll(&bCancel);
	else
		m_pComboCtrl->FireScroll(&bCancel);

	if (bCancel)
		return FALSE;

	if (CGXBrowserWnd::DoScroll(direction, nCell))
	{
		if (m_pDropDownCtrl)
			m_pDropDownCtrl->FireScrollAfter();
		else
			m_pComboCtrl->FireScrollAfter();

		return TRUE;
	}

	return FALSE;
}

BOOL CDropDownRealWnd::MoveTo(ROWCOL nRow)
{
	if (!CGXBrowserWnd::MoveTo(nRow))
		return FALSE;

	ROWCOL row, col;
	GetCurrentCell(row, col);
	if ((col < 1 || col > GetColCount()) && GetColCount() > 0)
		col = 1;

	return SetCurrentCell(nRow, col);
}

// update the selection range to simulate the combobox
BOOL CDropDownRealWnd::SetCurrentCell(ROWCOL nRow, ROWCOL nCol, UINT flags)
{
	if (!CGXBrowserWnd::SetCurrentCell(nRow, nCol, flags))
		return FALSE;

	CRowColArray awLeft, awRight;

	if (nRow > 0 && m_pComboCtrl)
	{
		SelectRange(CGXRange().SetTable(), FALSE);
		SelectRange(CGXRange().SetRows(nRow), TRUE);
		ReturnText(nRow);
	}

	return TRUE;
}
