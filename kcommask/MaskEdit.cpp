// MaskEdit.cpp : implementation file
//

#include "stdafx.h"
#include "KCOMMask.h"
#include "MaskEdit.h"
#include "KCOMMaskCtl.h"
#include <afxpriv.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMaskEdit

CMaskEdit::CMaskEdit(CKCOMMaskCtrl * pCtrl)
{
	m_pCtrl = pCtrl;
}

CMaskEdit::~CMaskEdit()
{
}


BEGIN_MESSAGE_MAP(CMaskEdit, CEdit)
	//{{AFX_MSG_MAP(CMaskEdit)
	ON_WM_CHAR()
	ON_WM_KEYDOWN()
	ON_WM_LBUTTONDOWN()
	ON_MESSAGE_VOID(WM_PASTE, OnPaste)
	ON_MESSAGE_VOID(WM_CUT, OnCut)
	ON_MESSAGE_VOID(WM_CLEAR, OnClear)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMaskEdit message handlers

void CMaskEdit::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	CString strText;
	GetWindowText(strText);
	int nStart, nEnd;
	
	if (m_pCtrl->m_strInputMask.GetLength())
	{
		GetSel(nStart, nEnd);

		if (nChar == VK_BACK)
		{
			MoveLeft();
			GetSel(nStart, nEnd);
			strText.SetAt(nStart, m_pCtrl->m_strPromptChar[0]);
		}
		else if (m_pCtrl->IsValidChar(nChar, nStart))
		{
			GetSel(nStart, nEnd);
			strText.SetAt(nStart, nChar);
			MoveRight();
			m_pCtrl->MoveRight();
		}

		m_pCtrl->SetDisplayText(strText);

		return;
	}
	else
	{
		// let the default window procedure to update the text and caret :)
		CEdit::OnChar(nChar, nRepCnt, nFlags);

		GetWindowText(strText);
		m_pCtrl->SetDisplayText(strText);
	}
}

void CMaskEdit::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	CString strText;
	int nStart, nEnd;

	if (m_pCtrl->m_strInputMask.GetLength() == 0)
	{
		CEdit::OnKeyDown(nChar, nRepCnt, nFlags);
		GetWindowText(strText);

		if (nChar == VK_DELETE)
			m_pCtrl->SetDisplayText(strText);

		return;
	}
	
	BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
	if (bCtrl && nChar == 'C')
	{
		Copy();
		return;
	}
	else if (bCtrl && nChar == 'V')
	{
		Paste();
		return;
	}
	else if (bCtrl && nChar == 'X')
	{
		Cut();
		return;
	}

	GetWindowText(strText);
	GetSel(nStart, nEnd);
	
	if (nChar == VK_DELETE)
	{
		if(nStart < m_pCtrl->m_nMaxLength)
		{
			GetSel(nStart, nEnd);
			strText.SetAt(nStart, m_pCtrl->m_strPromptChar[0]);
			m_pCtrl->SetDisplayText(strText);
		}

		return;
	}

	switch (nChar)
	{
	case VK_HOME:
		SetSel(0, 0);
		m_pCtrl->MoveRight();
		break;

	case VK_LEFT:
	case VK_UP:
		MoveLeft();
		break;

	case VK_END:
		SetSel(strText.GetLength(), strText.GetLength());
		m_pCtrl->MoveRight();
		break;

	case VK_RIGHT:
	case VK_DOWN:
		MoveRight();
		m_pCtrl->MoveRight();
		break;

	default:
		CEdit::OnKeyDown(nChar, nRepCnt, nFlags);
	}
}

void CMaskEdit::MoveLeft()
{
	int nStart, nEnd;

	BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
	if (bShift)
		GetSel(nEnd, nStart);
	else
		GetSel(nStart, nEnd);

	if (m_pCtrl->m_strInputMask.IsEmpty())
	{
		if (bShift)
			SetSel(nStart, nEnd - 1);
		else
			SetSel(nEnd - 1, nEnd - 1);

		return;
	}

	for (int i = nEnd - 1; i >= 0 && m_pCtrl->m_bIsDelimiter[i]; i--);

	if (i >= 0 && !m_pCtrl->m_bIsDelimiter[i])
		nEnd = i;

	if (bShift)
		SetSel(nStart, nEnd);
	else
		SetSel(nEnd, nEnd);
}

void CMaskEdit::MoveRight()
{
	int nStart, nEnd;

	GetSel(nStart, nEnd);
	BOOL bShift = (GetKeyState(VK_SHIFT) < 0);

	if (m_pCtrl->m_strInputMask.IsEmpty())
	{
		if (bShift)
			SetSel(nStart, nEnd + 1);
		else
			SetSel(nEnd, nEnd + 1);
		return;
	}

	for (int i = nEnd + 1; i <= m_pCtrl->m_strMask.GetLength() && 
		m_pCtrl->m_bIsDelimiter[i]; i++);

	if (!m_pCtrl->m_bIsDelimiter[i])
		nEnd = i;

	if (bShift)
		SetSel(nStart, nEnd);
	else
		SetSel(nEnd, nEnd);
}

void CMaskEdit::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CEdit::OnLButtonDown(nFlags, point);

	m_pCtrl->MoveRight();
}

void CMaskEdit::OnPaste()
{
	DefWindowProc(WM_PASTE, 0, 0);

	CString strText;
	GetWindowText(strText);
	m_pCtrl->SetDisplayText(strText);
	m_pCtrl->MoveRight();
}

void CMaskEdit::OnCut()
{
	DefWindowProc(WM_CUT, 0, 0);

	CString strText;
	GetWindowText(strText);
	m_pCtrl->SetDisplayText(strText);
	m_pCtrl->MoveRight();
}


void CMaskEdit::OnClear()
{
	DefWindowProc(WM_CLEAR, 0, 0);

	CString strText;
	GetWindowText(strText);
	m_pCtrl->SetDisplayText(strText);
	m_pCtrl->MoveRight();
}
