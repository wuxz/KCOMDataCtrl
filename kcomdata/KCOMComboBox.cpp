// KCOMComboBox.cpp : implementation file
//

#include "stdafx.h"
#include "kcomdata.h"
#include "KCOMComboBox.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CKCOMComboBox

CKCOMComboBox::CKCOMComboBox(CGXGridCore* pGrid, UINT nEditID, UINT nListBoxID, UINT nFlags, BOOL bDropDownList) : 
	CGXComboBox(pGrid, nEditID, nListBoxID, nFlags)
{
	m_bDropDownList = bDropDownList;
}

CKCOMComboBox::~CKCOMComboBox()
{
}


BEGIN_MESSAGE_MAP(CKCOMComboBox, CGXComboBox)
	//{{AFX_MSG_MAP(CKCOMComboBox)
	ON_WM_KEYDOWN()
	ON_WM_CHAR()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CKCOMComboBox message handlers

BOOL CKCOMComboBox::GetValue(CString &strResult)
{
	if (!CGXComboBox::GetValue(strResult))
		return FALSE;

	NeedStyle();

	CString strUserAttrib;
	COleCurrency cur;
	if (m_pStyle->GetIncludeUserAttribute(GX_IDS_UA_DATATYPE))
	{
		m_pStyle->GetUserAttribute(GX_IDS_UA_DATATYPE, strUserAttrib);

		switch (atoi(strUserAttrib))
		{
		case 1:
			strResult = (atoi(strResult) == 0) ? _T("0") : _T("-1");
			break;

		case 2:
			strResult.Format("%d", (char)atoi(strResult));
			break;

		case 3:
			strResult.Format("%d", (short)atoi(strResult));
			break;

		case 4:
			strResult.Format("%d", atoi(strResult));
			break;

		case 5:
			strResult.Format("%f", atof(strResult));
			break;

		case 6:
			cur.ParseCurrency(strResult);
			if (cur.GetStatus() == COleCurrency::valid)
				strResult = cur.Format();
			else
				strResult.Empty();

			break;
		}
	}
	
	return TRUE;
}

void CKCOMComboBox::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == VK_DELETE && m_bDropDownList)
		return;
	
	CGXComboBox::OnKeyDown(nChar, nRepCnt, nFlags);
}

void CKCOMComboBox::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (!m_bDropDownList)
		CGXComboBox::OnChar(nChar, nRepCnt, nFlags);
}

BOOL CKCOMComboBox::OnGridChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (m_bDropDownList)
		return TRUE;

	return CGXComboBox::OnGridChar(nChar, nRepCnt, nFlags);
}

BOOL CKCOMComboBox::OnGridKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (nChar == VK_DELETE && m_bDropDownList)
		return TRUE;

	return CGXComboBox::OnGridKeyDown(nChar, nRepCnt, nFlags);
}
