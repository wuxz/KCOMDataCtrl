// ComboEdit.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "ComboEdit.h"
#include "KCOMRichComboCtl.h"
#include "DropDownRealWnd.h"
#include <afxpriv.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CComboEdit

CComboEdit::CComboEdit(CKCOMRichComboCtrl * pCtrl)
{
	m_pCtrl = pCtrl;
}

CComboEdit::~CComboEdit()
{
}


BEGIN_MESSAGE_MAP(CComboEdit, CEdit)
	//{{AFX_MSG_MAP(CComboEdit)
	ON_WM_CHAR()
	ON_WM_KEYDOWN()
	ON_WM_SETFOCUS()
	ON_WM_LBUTTONDOWN()
	ON_MESSAGE_VOID(WM_PASTE, OnPaste)
	ON_MESSAGE_VOID(WM_CUT, OnCut)
	ON_MESSAGE_VOID(WM_CLEAR, OnClear)
	ON_WM_KILLFOCUS()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CComboEdit message handlers
void CComboEdit::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (m_pCtrl->m_nStyle || m_pCtrl->GetReadOnly())
		return;

	int nStart, nEnd;
	GetSel(nStart, nEnd);
	CString strText;
	GetWindowText(strText);
	if (!(m_pCtrl->m_strMask.IsEmpty()))
	{
		if (nEnd == strText.GetLength() && nChar != VK_BACK)
			return;
		
		else if (nChar == VK_BACK)
		{
			MoveLeft();
			GetSel(nStart, nEnd);
			strText.SetAt(nEnd, '_');
			m_pCtrl->SetText(strText);
			m_pCtrl->FireChange();
		}
		else if (m_pCtrl->IsValidChar(nChar, nEnd))
		{
			MoveRight();
			strText.SetAt(nEnd, nChar);
			m_pCtrl->SetText(strText);
			m_pCtrl->FireChange();
		}
		return;
	}
	
	// let the default window procedure to update the text and caret :)
	//		DefWindowProc(WM_CHAR, nCharRet, MAKELONG(nRepCnt, nFlags));
	CEdit::OnChar(nChar, nRepCnt, nFlags);
}

void CComboEdit::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	int nStart, nEnd;
	GetSel(nStart, nEnd);
	CString strText;
	GetWindowText(strText);
	if (!(m_pCtrl->m_strMask.IsEmpty()) && nChar == VK_DELETE)
	{
		if (m_pCtrl->m_nStyle || m_pCtrl->GetReadOnly())
			return;
		
		if (nEnd == strText.GetLength())
			return;
		strText.SetAt(nEnd, '_');
		m_pCtrl->SetText(strText);
		m_pCtrl->FireChange();
		return;
	}
	
	ROWCOL nRow;

	if (m_pCtrl->GetReadOnly())
		return;

	switch (nChar)
	{
	case VK_LEFT:
		MoveLeft();
		break;

	case VK_RIGHT:
		MoveRight();
		break;

	// just simulate the auto navigation function of the normal combobox
	case VK_UP:
		nRow = m_pCtrl->SearchText(strText, FALSE);
		if (nRow == GX_INVALID)
			nRow = 2;
		nRow --;
		m_pCtrl->m_pDropDownRealWnd->MoveTo(nRow);

		break;

	case VK_DOWN:
		nRow = m_pCtrl->SearchText(strText, FALSE);
		if (nRow == GX_INVALID)
			nRow = 0;
		nRow ++;
		m_pCtrl->m_pDropDownRealWnd->MoveTo(nRow);

		break;

	default:
		//	DefWindowProc(WM_KEYDOWN, nCharRet, MAKELONG(nRepCnt, nFlags));
		CEdit::OnKeyDown(nChar, nRepCnt, nFlags);
	}
}

void CComboEdit::OnPaste()
{
	// Do nothing for we are mask :)
}

void CComboEdit::OnClear()
{
	// Do nothing for we are mask :)
}

void CComboEdit::OnCut()
{
	// Do nothing for we are mask
}

void CComboEdit::MoveLeft()
{
	int nStart, nEnd;
	GetSel(nStart, nEnd);

	if (m_pCtrl->m_strMask.IsEmpty())
	{
		SetSel(nEnd - 1, nEnd - 1);
		return;
	}

	for (int i = nEnd - 1; i >= 0 && m_pCtrl->m_bIsDelimiter[i]; i--);

	if (i >= 0 && !m_pCtrl->m_bIsDelimiter[i])
		nEnd = i;

	SetSel(nEnd, nEnd);
}

void CComboEdit::MoveRight()
{
	int nStart, nEnd;
	CString strText;

	GetWindowText(strText);
	GetSel(nStart, nEnd);
	if (nEnd >= 0 && nEnd < strText.GetLength() && strText[nEnd] < 0)
		nEnd ++;

	if (m_pCtrl->m_strMask.IsEmpty())
	{
		SetSel(nEnd + 1, nEnd + 1);
		return;
	}

	for (int i = nEnd + 1; i < m_pCtrl->m_strMask.GetLength() && 
		m_pCtrl->m_bIsDelimiter[i]; i++);

	if (!m_pCtrl->m_bIsDelimiter[i])
		nEnd = i;

	SetSel(nEnd, nEnd);
}

// can not put the caret at thte delimiter
void CComboEdit::OnSetFocus(CWnd* pOldWnd) 
{
	CEdit::OnSetFocus(pOldWnd);
	
	int nStart, nEnd;
	GetSel(nStart, nEnd);

	if (m_pCtrl->m_bIsDelimiter[nEnd])
		MoveRight();
	GetSel(nStart, nEnd);
	if (m_pCtrl->m_bIsDelimiter[nEnd])
		MoveLeft();
}

void CComboEdit::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CEdit::OnLButtonDown(nFlags, point);

	int nStart, nEnd;
	GetSel(nStart, nEnd);
	if (m_pCtrl->m_bIsDelimiter[nEnd])
		MoveRight();
	GetSel(nStart, nEnd);
	if (m_pCtrl->m_bIsDelimiter[nEnd])
		MoveLeft();
}

void CComboEdit::OnKillFocus(CWnd* pNewWnd) 
{
	CEdit::OnKillFocus(pNewWnd);
	
	m_pCtrl->HideDropDownWnd();
}
